import streamlit as st
import pandas as pd
import re
from datetime import datetime, date
from io import BytesIO
from html import escape as html_escape

st.set_page_config(page_title="Extract Excel Data", layout="wide")
st.title("Extract Excel Data")

# ── Helpers ───────────────────────────────────

DATE_PATTERNS = [
    r"^\d{1,2}/\d{1,2}/\d{2,4}$",
    r"^\d{1,2}-\d{1,2}-\d{2,4}$",
    r"^\d{4}[/-]\d{1,2}[/-]\d{1,2}$",
    r"^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,-]+\d",
]


def _is_date_cell(val) -> bool:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    if isinstance(val, datetime):
        try:
            if pd.isna(val):
                return False
        except (TypeError, ValueError):
            pass
        return True
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", ""):
        return False
    for p in DATE_PATTERNS:
        if re.search(p, s, re.IGNORECASE):
            return True
    return False


def _cell_str(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(val, datetime):
        try:
            return f"{val.month}/{val.day}/{val.year}"
        except Exception:
            return str(val)
    s = str(val).strip()
    return "" if s.lower() in ("nan", "nat") else s


def _format_number(val: str) -> str:
    """If value is numeric, format to 2 decimal places."""
    if not val:
        return val
    cleaned = val.replace(",", "")
    try:
        num = float(cleaned)
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return val


def _parse_date_from_header(header_str: str):
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(header_str, fmt).date()
        except (ValueError, TypeError):
            continue
    return None


def _normalise(text: str) -> str:
    s = " ".join(str(text).lower().split())
    s = s.lstrip("*").strip()
    return s


def _find_date_header_row(df: pd.DataFrame):
    for row_idx in range(min(len(df), 30)):
        row = df.iloc[row_idx]
        date_cols = {}
        for col_idx in range(len(row)):
            val = row.iloc[col_idx]
            if _is_date_cell(val):
                date_cols[col_idx] = _cell_str(val)
        if len(date_cols) >= 2:
            return row_idx, date_cols
    return None, {}


def _find_extra_col_indices(df_row, extra_cols: list[str]):
    indices = {}
    for col_idx in range(len(df_row)):
        cell_norm = _normalise(_cell_str(df_row.iloc[col_idx]))
        for ec in extra_cols:
            if _normalise(ec) == cell_norm and ec not in indices:
                indices[ec] = col_idx
    return indices


def _match_label(cell_text: str, label: str) -> bool:
    return _normalise(cell_text) == _normalise(label)


# ──────────────────────────────────────────────
# 1) Upload Excel file
# ──────────────────────────────────────────────
uploaded_file = st.file_uploader("Upload Excel Sheet", type=["xlsx", "xls"])

# ──────────────────────────────────────────────
# 2) Row labels (comma-separated)
# ──────────────────────────────────────────────
row_labels_input = st.text_input(
    "Row labels to extract (comma-separated)",
    placeholder='e.g. Wine, MOP Cash (Dollar), EBT Cash, MOP Credit',
)

# ──────────────────────────────────────────────
# 3) Optional extra columns
# ──────────────────────────────────────────────
extra_cols_input = st.text_input(
    "Additional columns to extract — optional (comma-separated). Date columns are included automatically.",
    placeholder="e.g. Total, Field",
)

# ──────────────────────────────────────────────
# Extract button
# ──────────────────────────────────────────────
extract_clicked = st.button("Extract", type="primary")

# ──────────────────────────────────────────────
# Extraction logic — store results in session_state
# ──────────────────────────────────────────────
if extract_clicked:
    if not uploaded_file:
        st.warning("Please upload an Excel file first.")
        st.stop()
    if not row_labels_input or not row_labels_input.strip():
        st.warning("Please enter at least one row label.")
        st.stop()

    row_labels = [lbl.strip() for lbl in row_labels_input.split(",") if lbl.strip()]
    extra_cols = [c.strip() for c in extra_cols_input.split(",") if c.strip()] if extra_cols_input else []

    try:
        all_sheets: dict[str, pd.DataFrame] = pd.read_excel(
            uploaded_file, sheet_name=None, header=None
        )
    except Exception as e:
        st.error(f"Failed to read the Excel file: {e}")
        st.stop()

    errors: list[str] = []
    results: list[dict] = []

    for sheet_name, df in all_sheets.items():
        if df.empty:
            errors.append(f"Sheet '{sheet_name}': sheet is empty — skipped.")
            continue

        header_row_idx, date_col_map = _find_date_header_row(df)
        if header_row_idx is None:
            errors.append(f"Sheet '{sheet_name}': no row with date columns found — skipped.")
            continue

        date_headers = {ci: raw for ci, raw in date_col_map.items()}

        extra_col_map = {}
        if extra_cols:
            extra_col_map = _find_extra_col_indices(df.iloc[header_row_idx], extra_cols)
            for ec in extra_cols:
                if ec not in extra_col_map:
                    errors.append(f"Sheet '{sheet_name}': extra column '{ec}' not found in header row.")

        all_col_indices = list(dict.fromkeys(list(date_headers.keys()) + list(extra_col_map.values())))

        col_display = {}
        for ci in all_col_indices:
            col_display[ci] = date_headers[ci] if ci in date_headers else _cell_str(df.iloc[header_row_idx].iloc[ci])

        for label in row_labels:
            matched = False
            for row_idx in range(header_row_idx + 1, len(df)):
                row = df.iloc[row_idx]
                for check_col in range(min(3, len(row))):
                    cell_val = _cell_str(row.iloc[check_col])
                    if cell_val and _match_label(cell_val, label):
                        matched = True
                        row_dict = {"Row Label": label}
                        for ci in all_col_indices:
                            header_name = col_display[ci]
                            val = _cell_str(row.iloc[ci]) if ci < len(row) else ""
                            row_dict[header_name] = _format_number(val)
                        results.append(row_dict)
                        break
                if matched:
                    break
            if not matched:
                errors.append(f"Sheet '{sheet_name}': row label '{label}' not found.")

    # Merge rows with the same label across sheets
    merged: dict[str, dict] = {}
    for row_dict in results:
        label = row_dict["Row Label"]
        if label not in merged:
            merged[label] = dict(row_dict)
        else:
            for k, v in row_dict.items():
                if k == "Row Label":
                    continue
                if v and (k not in merged[label] or not merged[label][k]):
                    merged[label][k] = v
    final_results = list(merged.values())

    # Save to session state
    st.session_state["extracted_df"] = pd.DataFrame(final_results) if final_results else None
    st.session_state["extract_errors"] = errors

# ──────────────────────────────────────────────
# Display results (persists across reruns via session_state)
# ──────────────────────────────────────────────
if "extracted_df" in st.session_state and st.session_state["extracted_df"] is not None:
    result_df: pd.DataFrame = st.session_state["extracted_df"].copy()

    # Drop date columns where every row is empty
    cols_to_drop = [
        col for col in result_df.columns
        if col != "Row Label" and result_df[col].fillna("").astype(str).str.strip().eq("").all()
    ]
    if cols_to_drop:
        result_df = result_df.drop(columns=cols_to_drop)

    # Detect date range from actual data columns
    date_cols_parsed: list[date] = []
    for col in result_df.columns:
        parsed = _parse_date_from_header(str(col))
        if parsed:
            date_cols_parsed.append(parsed)
    date_cols_parsed.sort()

    # ── Date range filter — defaults come from the data ──
    if date_cols_parsed:
        data_min = date_cols_parsed[0]
        data_max = date_cols_parsed[-1]

        st.write("**Date range filter**")
        dcol1, dcol2 = st.columns(2)
        with dcol1:
            start_date = st.date_input(
                "Start date",
                value=data_min,
                min_value=data_min,
                max_value=data_max,
                format="MM/DD/YYYY",
            )
        with dcol2:
            end_date = st.date_input(
                "End date",
                value=data_max,
                min_value=data_min,
                max_value=data_max,
                format="MM/DD/YYYY",
            )

        # Filter columns by date range
        cols_to_keep = []
        for col in result_df.columns:
            if col == "Row Label":
                cols_to_keep.append(col)
                continue
            parsed = _parse_date_from_header(str(col))
            if parsed is None:
                cols_to_keep.append(col)  # non-date column — always keep
            elif start_date <= parsed <= end_date:
                cols_to_keep.append(col)
        result_df = result_df[[c for c in cols_to_keep if c in result_df.columns]]

    # Ensure all column names are plain strings
    result_df.columns = [str(c) for c in result_df.columns]

    # Sort date columns chronologically; keep Row Label first, non-date cols last
    date_col_order = sorted(
        [c for c in result_df.columns if _parse_date_from_header(c)],
        key=lambda c: _parse_date_from_header(c),
    )
    non_date_cols = [c for c in result_df.columns if c != "Row Label" and _parse_date_from_header(c) is None]
    result_df = result_df[["Row Label"] + date_col_order + non_date_cols]

    st.subheader("Results")

    # Calculate height to show all rows
    table_height = min(38 + len(result_df) * 35 + 20, 2000)

    st.dataframe(
        result_df,
        use_container_width=True,
        hide_index=True,
        height=table_height,
    )

    # ── Action buttons: Copy + Download Excel ──
    copy_df = result_df.copy()
    copy_df.columns = [
        f"'{c}" if _parse_date_from_header(c) else c
        for c in copy_df.columns
    ]
    tsv = copy_df.to_csv(sep="\t", index=False)
    tsv_escaped = html_escape(tsv)
    copy_icon = "\U0001F4CB"

    btn_col1, btn_col2 = st.columns([1, 1])
    with btn_col1:
        st.components.v1.html(
            f"""
            <textarea id="tsv-data" style="position:absolute;left:-9999px">{tsv_escaped}</textarea>
            <button id="copy-btn" onclick="
                var t=document.getElementById('tsv-data');
                t.style.position='static';
                t.select();
                document.execCommand('copy');
                t.style.position='absolute';
                this.textContent='Copied!';
                setTimeout(function(){{ document.getElementById('copy-btn').textContent='{copy_icon} Copy Data'; }},1500);
            " style="padding:0.4em 1.2em;border-radius:6px;border:1px solid #ccc;
                     background:#f0f2f6;cursor:pointer;font-size:14px">
                {copy_icon} Copy Data
            </button>
            """,
            height=50,
        )
    with btn_col2:
        output = BytesIO()
        result_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)
        st.download_button(
            label="Download as Excel",
            data=output,
            file_name="extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif "extracted_df" in st.session_state and st.session_state["extracted_df"] is None:
    st.info("No matching rows found across any sheet.")

# ── Errors / Warnings ──
if st.session_state.get("extract_errors"):
    st.subheader("Errors / Warnings")
    for err in st.session_state["extract_errors"]:
        st.warning(err)
