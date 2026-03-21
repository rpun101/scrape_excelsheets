# Extract Excel Data

A Streamlit web app that extracts specific row data from Excel files across multiple sheets.

## Features

- **Upload** `.xlsx` / `.xls` files
- **Search by row label** — enter comma-separated labels (case-insensitive)
- **Auto-detects date columns** in the header row (e.g. `1/5/2025`, `2024-01-02`, `Mar 2024`)
- **Optional extra columns** — specify additional column names to include
- **Multi-sheet support** — scans every sheet; the same label can be at different row numbers in each sheet
- **Merged-cell aware** — checks the first few columns for labels (handles merged/shifted cells)
- **Date range filter** — filter extracted columns by a start/end date
- **Error reporting** — shows which labels or columns were not found, per sheet (terminal-style output)
- **Copy to clipboard** — copy results as tab-separated values for pasting into Excel
- **Download** — export results as an Excel file

## TODO

- [ ] **Multiple file upload** — upload and process more than one workbook at a time; results merged into a single table
- [ ] **Column display order** — a text area with comma-separated column/label names to control which columns are shown first in the results (e.g. enter `Cash, Cash (EBT)` to pin those to the left)
- [ ] **Download as Pdf** - button to download result as pdf

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

### Libraries

| Library | Purpose |
|---|---|
| [streamlit](https://streamlit.io/) | Web UI framework — renders the page, widgets, tables, and download buttons |
| [pandas](https://pandas.pydata.org/) | Reads Excel workbooks, builds and filters the results `DataFrame` |
| [openpyxl](https://openpyxl.readthedocs.io/) | Excel engine used by pandas to open `.xlsx` files and write the downloadable output |

## Setup

```bash
cd extract_excel_data
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

## How to Use

1. **Upload** an Excel file using the file uploader.
2. **Enter row labels** in the text box, separated by commas.
   - Example: `Wine, MOP Cash (Dollar), EBT Cash, MOP Credit`
   - Matching is **case-insensitive** and ignores leading `*` characters.
   - Whitespace around labels is trimmed automatically.
3. *(Optional)* **Enter extra column names** — any non-date columns you also want extracted (e.g. `Total`).
4. Click the **Extract** button.
5. **View results** in the on-page table.
6. Click **Download results as Excel** to save the output.

## How It Works

1. The app reads every sheet in the uploaded workbook.
2. It scans rows (up to the first 30) to find the **date header row** — the row containing at least two date-formatted cells.
3. For each requested label, it searches rows below the header, checking the first few columns for a match (to handle merged cells).
4. When a match is found, it extracts the values from the date columns (and any extra columns) for that row.
5. Results from all sheets are combined into a single table.

## Example

Given an Excel file with a sheet like:

| Field              |   | 1/13/2025 | 1/14/2025 | 1/15/2025 | Total    |
|--------------------|---|-----------|-----------|-----------|----------|
| ...                |   |           |           |           |          |
| Wine               |   | 132.26    | 181.51    | 96.58     | 1,075.00 |
| MOP Cash (Dollar)  |   | 1,764.26  | 2,006.26  | 2,385.45  | 10,164.57|

Searching for `Wine, MOP Cash (Dollar)` with extra column `Total` will produce:

| Sheet  | Row Label          | Excel Row # | 1/13/2025 | 1/14/2025 | 1/15/2025 | Total    |
|--------|--------------------|-------------|-----------|-----------|-----------|----------|
| Sheet1 | Wine               | 87          | 132.26    | 181.51    | 96.58     | 1,075.00 |
| Sheet1 | MOP Cash (Dollar)  | 102         | 1,764.26  | 2,006.26  | 2,385.45  | 10,164.57|
