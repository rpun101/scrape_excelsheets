import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Extract Excel Data", layout="wide")
st.title("Extract Excel Data")

# ──────────────────────────────────────────────
# 1) Upload Excel file
# ──────────────────────────────────────────────
uploaded_file = st.file_uploader("Upload Excel Sheet", type=["xlsx", "xls"])

# ──────────────────────────────────────────────
# 2) Row labels to search for (comma-separated)
# ──────────────────────────────────────────────
row_labels_input = st.text_input(
    "Row labels to extract (comma-separated)",
    placeholder='e.g. Wine, MOP Cash (Dollar), EBT Cash, MOP Credit',
)

# ──────────────────────────────────────────────
# 3) Optional extra columns (dates auto-detected)
# ──────────────────────────────────────────────
extra_cols_input = st.text_input(
    "Additional columns to extract — optional (comma-separated). Date columns are included automatically.",
    placeholder="e.g. Total, Field",
)

# ──────────────────────────────────────────────
# Extract button
# ──────────────────────────────────────────────
extract_clicked = st.button("Extract", type="primary")


# ── Helpers ───────────────────────────────────

DATE_PATTERNS = [
    r"^\d{1,2}/\d{1,2}/\d{2,4}$",          # 1/5/2025, 01/02/2024
    r"^\d{1,2}-\d{1,2}-\d{2,4}$",           # 1-5-2025
    r"^\d{4}[/-]\d{1,2}[/-]\d{1,2}$",       # 2024-01-02
    r"^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,-]+\d",
]


def _is_date_cell(val) -> bool:
    """Check if a cell value looks like a date (string or datetime object)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    if isinstance(val, datetime):
        # pd.NaT is also a datetime subclass — reject it
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
    """Convert a cell value to a clean string."""
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


def _normalise(text: str) -> str:
    """Lower-case, collapse whitespace, strip leading symbols like * for comparison."""
    s = " ".join(str(text).lower().split())
    s = s.lstrip("*").strip()
    return s


def _find_date_header_row(df: pd.DataFrame):
    """Scan rows to find the one that contains dates in multiple columns.
    Returns (row_index, {col_index: date_string}) or (None, {}).
    """
    for row_idx in range(min(len(df), 30)):  # look in first 30 rows
        row = df.iloc[row_idx]
        date_cols = {}
        for col_idx in range(len(row)):
            val = row.iloc[col_idx]
            if _is_date_cell(val):
                date_cols[col_idx] = _cell_str(val)
        if len(date_cols) >= 2:  # at least 2 date columns → header row
            return row_idx, date_cols
    return None, {}


def _find_extra_col_indices(df_row, extra_cols: list[str]):
    """Find column indices matching user-requested extra column names."""
    indices = {}
    for col_idx in range(len(df_row)):
        cell_norm = _normalise(_cell_str(df_row.iloc[col_idx]))
        for ec in extra_cols:
            if _normalise(ec) == cell_norm and ec not in indices:
                indices[ec] = col_idx
    return indices


def _match_label(cell_text: str, label: str) -> bool:
    """Case-insensitive match: exact match after normalisation."""
    return _normalise(cell_text) == _normalise(label)


# ──────────────────────────────────────────────
# Main extraction logic — only runs on button click
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

    # Read all sheets — keep raw values (no header inference)
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

        # ── Step 1: find the header row that contains dates ──
        header_row_idx, date_col_map = _find_date_header_row(df)

        if header_row_idx is None:
            errors.append(f"Sheet '{sheet_name}': no row with date columns found — skipped.")
            continue

        # Format date column headers nicely
        date_headers = {}  # col_idx → display string
        for ci, raw in date_col_map.items():
            date_headers[ci] = raw

        # ── Step 2: resolve extra columns from the same header row ──
        extra_col_map = {}  # name → col_idx
        if extra_cols:
            extra_col_map = _find_extra_col_indices(df.iloc[header_row_idx], extra_cols)
            for ec in extra_cols:
                if ec not in extra_col_map:
                    errors.append(f"Sheet '{sheet_name}': extra column '{ec}' not found in header row.")

        # Combine column indices (dates + extras, no duplicates)
        all_col_indices = list(dict.fromkeys(list(date_headers.keys()) + list(extra_col_map.values())))

        # Build column display names
        col_display = {}
        for ci in all_col_indices:
            if ci in date_headers:
                col_display[ci] = date_headers[ci]
            else:
                col_display[ci] = _cell_str(df.iloc[header_row_idx].iloc[ci])

        # ── Step 3: search for each label in rows BELOW the header ──
        for label in row_labels:
            matched = False
            for row_idx in range(header_row_idx + 1, len(df)):
                row = df.iloc[row_idx]
                # Check first few columns for the label (handles merged cells)
                for check_col in range(min(3, len(row))):
                    cell_val = _cell_str(row.iloc[check_col])
                    if cell_val and _match_label(cell_val, label):
                        matched = True
                        row_dict = {
                            "Row Label": label,
                        }
                        for ci in all_col_indices:
                            header_name = col_display[ci]
                            val = _cell_str(row.iloc[ci]) if ci < len(row) else ""
                            row_dict[header_name] = val
                        results.append(row_dict)
                        break
                if matched:
                    break

            if not matched:
                errors.append(f"Sheet '{sheet_name}': row label '{label}' not found.")

    # ──────────────────────────────────────────
    # 4) Merge rows with the same label across sheets
    # ──────────────────────────────────────────
    # Each sheet may contribute different date columns for the same label.
    # Merge them into one row per label so all values appear together.
    merged: dict[str, dict] = {}  # label → combined dict
    for row_dict in results:
        label = row_dict["Row Label"]
        if label not in merged:
            merged[label] = dict(row_dict)
        else:
            # Fill in any columns that are missing or empty
            for k, v in row_dict.items():
                if k == "Row Label":
                    continue
                if v and (k not in merged[label] or not merged[label][k]):
                    merged[label][k] = v
    final_results = list(merged.values())

    # ──────────────────────────────────────────
    # 5) Display results
    # ──────────────────────────────────────────
    st.subheader("Results")

    if final_results:
        result_df = pd.DataFrame(final_results)
        st.dataframe(result_df, use_container_width=True, hide_index=True)

        # Copy to clipboard (tab-separated for pasting into Excel/Sheets)
        tsv = result_df.to_csv(sep="\t", index=False)
        st.code(tsv, language=None)
        st.caption("Select the text above and copy, or use the buttons below.")

        col1, col2 = st.columns(2)
        with col1:
            # Download as Excel
            output = BytesIO()
            result_df.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)
            st.download_button(
                label="Download as Excel",
                data=output,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col2:
            # Download as CSV
            csv_data = result_df.to_csv(index=False)
            st.download_button(
                label="Download as CSV",
                data=csv_data,
                file_name="extracted_data.csv",
                mime="text/csv",
            )
    else:
        st.info("No matching rows found across any sheet.")

    # ── Errors / Warnings ──
    if errors:
        st.subheader("Errors / Warnings")
        for err in errors:
            st.warning(err)
