# Excel Rule Merger

A **.NET 8 Blazor Web App** that lets you upload an Excel workbook (`.xlsx`), apply configurable filter rules across multiple sheets, and download a single merged output file as `.xlsx` or `.csv`.

## Features

| Feature | Details |
|---|---|
| **Three filter modes** | `Between Markers` — extract rows between start/end text values in a column<br>`Filter Equals` — keep rows where a column value matches an exact string<br>`Date Range` — keep rows where a date column falls within a start/end range |
| **Multi-sheet merge** | Process all sheets or specific named sheets; results are combined into one table |
| **Date auto-detection** | Configurable regex patterns automatically identify the date column when not specified |
| **Source sheet column** | Optionally append a column containing the originating sheet name |
| **Flexible output** | Download as Excel (`.xlsx`) or CSV (`.csv`) |
| **Per-sheet logs** | Individual sheet errors are logged and displayed; processing continues for remaining sheets |

## Tech Stack

- .NET 8 / ASP.NET Core Blazor Web App (static SSR)
- [EPPlus 7](https://www.epplus.net/) for Excel read/write (non-commercial license)
- Bootstrap 5 (bundled)

## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) or later

## Setup & Run

```bash
# Clone the repository
git clone https://github.com/rpun101/scrape_excelsheets.git
cd scrape_excelsheets

# Restore dependencies
dotnet restore ExcelRuleMerger/ExcelRuleMerger.csproj

# Run the application
dotnet run --project ExcelRuleMerger/ExcelRuleMerger.csproj
```

Then open your browser to `https://localhost:5001` (or the URL shown in the terminal) and navigate to **Excel Merger** in the sidebar.

## Usage

1. Open **http://localhost:5XXX/excel-merger**
2. Upload an `.xlsx` file
3. Select sheets and processing mode
4. Fill in mode-specific fields (marker text, filter value, or date range)
5. Click **Process & Download**

### Mode examples

**Between Markers** — extracts rows in the `Section` column between "START" and "END":
- Marker Column: `Section`
- Start Marker: `START`
- End Marker: `END`

**Filter Equals** — keeps rows where `Status == "Approved"`:
- Filter Column: `Status`
- Filter Value: `Approved`

**Date Range** — keeps rows where the auto-detected date column is within 2025:
- Start Date: `2025-01-01`
- End Date: `2025-12-31`

## EPPlus License Note

This application uses EPPlus under the **non-commercial personal license**.  
For commercial deployments, obtain a commercial license from [epplus.net](https://www.epplus.net/).

## Project Structure

```
ExcelRuleMerger/
├── Components/
│   ├── Layout/          # NavMenu, MainLayout
│   ├── Pages/
│   │   ├── Home.razor          # Landing page
│   │   └── ExcelMerger.razor   # Main tool UI (/excel-merger)
│   ├── App.razor
│   └── Routes.razor
├── Services/
│   └── ExcelMergerService.cs   # Core processing logic
├── wwwroot/
│   └── js/fileDownload.js      # JS interop for file download
└── Program.cs
```
