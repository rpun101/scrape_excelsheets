using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ExcelRuleMerger.Services;

public enum MergeMode
{
    BetweenMarkers,
    FilterEquals,
    DateRange
}

public enum OutputFormat
{
    Xlsx,
    Csv
}

public class MergeOptions
{
    public bool IncludeAllSheets { get; set; } = true;
    public string SpecificSheets { get; set; } = string.Empty;
    public MergeMode Mode { get; set; } = MergeMode.BetweenMarkers;

    // Between markers
    public string MarkerColumn { get; set; } = string.Empty;
    public string StartMarker { get; set; } = string.Empty;
    public string EndMarker { get; set; } = string.Empty;
    public bool IncludeStartRow { get; set; } = true;
    public bool IncludeEndRow { get; set; } = true;

    // Filter equals
    public string FilterColumn { get; set; } = string.Empty;
    public string FilterValue { get; set; } = string.Empty;

    // Date range
    public string DateColumn { get; set; } = string.Empty;
    public string StartDate { get; set; } = string.Empty;
    public string EndDate { get; set; } = string.Empty;
    public bool DayFirst { get; set; } = false;

    // Date column patterns (newline-separated regexes)
    public string DateColumnPatterns { get; set; } =
        "(?i)^date$\n(?i)^dt$\n(?i)date\n(?i)^transaction[ _-]?date$\n(?i)^posted[ _-]?date$\n(?i)^created[ _-]?at$\n(?i)^updated[ _-]?at$\n(?i)^timestamp$";

    // Output
    public OutputFormat OutputFormat { get; set; } = OutputFormat.Xlsx;
    public bool AddSourceSheetColumn { get; set; } = true;
    public string SourceSheetColumnName { get; set; } = "_source_sheet";
}

public class SheetProcessingLog
{
    public string SheetName { get; set; } = string.Empty;
    public bool Success { get; set; }
    public int RowsExtracted { get; set; }
    public string? ErrorMessage { get; set; }
}

public class MergeResult
{
    public byte[] FileBytes { get; set; } = Array.Empty<byte>();
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public List<SheetProcessingLog> Logs { get; set; } = new();
    public int TotalRows { get; set; }
}

public class ExcelMergerService
{
    private static readonly string[] DefaultDatePatterns =
    {
        @"(?i)^date$",
        @"(?i)^dt$",
        @"(?i)date",
        @"(?i)^transaction[ _-]?date$",
        @"(?i)^posted[ _-]?date$",
        @"(?i)^created[ _-]?at$",
        @"(?i)^updated[ _-]?at$",
        @"(?i)^timestamp$"
    };

    public Task<MergeResult> ProcessAsync(Stream excelStream, MergeOptions options)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage(excelStream);

        var targetSheets = ResolveTargetSheets(package, options);
        var datePatterns = ParseDatePatterns(options.DateColumnPatterns);

        // First pass: collect headers from all sheets to unify columns
        var allHeaders = new List<string>();
        var sheetDataMap = new Dictionary<string, (List<string> Headers, List<List<object?>> Rows)>();

        var logs = new List<SheetProcessingLog>();

        foreach (var sheetName in targetSheets)
        {
            var ws = package.Workbook.Worksheets[sheetName];
            if (ws == null)
            {
                logs.Add(new SheetProcessingLog
                {
                    SheetName = sheetName,
                    Success = false,
                    ErrorMessage = "Sheet not found in workbook."
                });
                continue;
            }

            try
            {
                var (headers, allRows) = ReadSheet(ws);
                var filteredRows = ApplyFilter(headers, allRows, options, datePatterns, sheetName);

                foreach (var h in headers)
                    if (!allHeaders.Contains(h))
                        allHeaders.Add(h);

                sheetDataMap[sheetName] = (headers, filteredRows);
                logs.Add(new SheetProcessingLog
                {
                    SheetName = sheetName,
                    Success = true,
                    RowsExtracted = filteredRows.Count
                });
            }
            catch (Exception ex)
            {
                logs.Add(new SheetProcessingLog
                {
                    SheetName = sheetName,
                    Success = false,
                    ErrorMessage = ex.Message
                });
            }
        }

        if (options.AddSourceSheetColumn && !allHeaders.Contains(options.SourceSheetColumnName))
            allHeaders.Add(options.SourceSheetColumnName);

        // Build merged table
        var mergedRows = new List<(string SheetName, List<object?> Cells)>();
        foreach (var sheetName in targetSheets)
        {
            if (!sheetDataMap.TryGetValue(sheetName, out var sheetData)) continue;
            var (sheetHeaders, rows) = sheetData;

            foreach (var row in rows)
            {
                var unified = new List<object?>();
                foreach (var h in allHeaders)
                {
                    if (h == options.SourceSheetColumnName)
                    {
                        unified.Add(options.AddSourceSheetColumn ? sheetName : null);
                    }
                    else
                    {
                        var idx = sheetHeaders.IndexOf(h);
                        unified.Add(idx >= 0 && idx < row.Count ? row[idx] : null);
                    }
                }
                mergedRows.Add((sheetName, unified));
            }
        }

        var result = BuildOutput(allHeaders, mergedRows, options);
        result.Logs = logs;
        result.TotalRows = mergedRows.Count;
        return Task.FromResult(result);
    }

    private List<string> ResolveTargetSheets(ExcelPackage package, MergeOptions options)
    {
        var all = package.Workbook.Worksheets.Select(w => w.Name).ToList();
        if (options.IncludeAllSheets) return all;

        var specified = options.SpecificSheets
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList();

        return specified.Count > 0 ? specified : all;
    }

    private List<Regex> ParseDatePatterns(string patternsText)
    {
        var patterns = new List<Regex>();
        var lines = (patternsText ?? string.Empty)
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try { patterns.Add(new Regex(line, RegexOptions.Compiled)); }
            catch { /* skip invalid patterns */ }
        }

        if (patterns.Count == 0)
            foreach (var p in DefaultDatePatterns)
                patterns.Add(new Regex(p, RegexOptions.Compiled));

        return patterns;
    }

    private (List<string> Headers, List<List<object?>> Rows) ReadSheet(ExcelWorksheet ws)
    {
        var dim = ws.Dimension;
        if (dim == null) return (new List<string>(), new List<List<object?>>());

        var headers = new List<string>();
        for (int col = dim.Start.Column; col <= dim.End.Column; col++)
        {
            var h = ws.Cells[dim.Start.Row, col].Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(h)) h = $"Column{col}";
            headers.Add(h);
        }

        var rows = new List<List<object?>>();
        for (int row = dim.Start.Row + 1; row <= dim.End.Row; row++)
        {
            var cells = new List<object?>();
            for (int col = dim.Start.Column; col <= dim.End.Column; col++)
                cells.Add(ws.Cells[row, col].Value);
            rows.Add(cells);
        }

        return (headers, rows);
    }

    private List<List<object?>> ApplyFilter(
        List<string> headers,
        List<List<object?>> rows,
        MergeOptions options,
        List<Regex> datePatterns,
        string sheetName)
    {
        return options.Mode switch
        {
            MergeMode.BetweenMarkers => FilterBetweenMarkers(headers, rows, options),
            MergeMode.FilterEquals => FilterEquals(headers, rows, options),
            MergeMode.DateRange => FilterDateRange(headers, rows, options, datePatterns, sheetName),
            _ => rows
        };
    }

    private List<List<object?>> FilterBetweenMarkers(List<string> headers, List<List<object?>> rows, MergeOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.MarkerColumn))
            throw new InvalidOperationException("Marker column is required for 'between_markers' mode.");

        var colIdx = headers.IndexOf(options.MarkerColumn);
        if (colIdx < 0)
            throw new InvalidOperationException($"Marker column '{options.MarkerColumn}' not found. Available: {string.Join(", ", headers)}");

        var result = new List<List<object?>>();
        bool inSegment = false;

        for (int i = 0; i < rows.Count; i++)
        {
            var cellText = rows[i][colIdx]?.ToString() ?? string.Empty;
            bool isStart = cellText.Contains(options.StartMarker, StringComparison.OrdinalIgnoreCase);
            bool isEnd = cellText.Contains(options.EndMarker, StringComparison.OrdinalIgnoreCase);

            if (!inSegment && isStart)
            {
                inSegment = true;
                if (options.IncludeStartRow) result.Add(rows[i]);
                continue;
            }

            if (inSegment && isEnd)
            {
                if (options.IncludeEndRow) result.Add(rows[i]);
                inSegment = false;
                continue;
            }

            if (inSegment) result.Add(rows[i]);
        }

        return result;
    }

    private List<List<object?>> FilterEquals(List<string> headers, List<List<object?>> rows, MergeOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.FilterColumn))
            throw new InvalidOperationException("Filter column is required for 'filter_equals' mode.");

        var colIdx = headers.IndexOf(options.FilterColumn);
        if (colIdx < 0)
            throw new InvalidOperationException($"Filter column '{options.FilterColumn}' not found. Available: {string.Join(", ", headers)}");

        return rows.Where(r => (r[colIdx]?.ToString() ?? string.Empty) == options.FilterValue).ToList();
    }

    private List<List<object?>> FilterDateRange(
        List<string> headers,
        List<List<object?>> rows,
        MergeOptions options,
        List<Regex> datePatterns,
        string sheetName)
    {
        string? dateColName = string.IsNullOrWhiteSpace(options.DateColumn) ? null : options.DateColumn;

        if (dateColName == null)
        {
            dateColName = headers.FirstOrDefault(h => datePatterns.Any(p => p.IsMatch(h)));
            if (dateColName == null)
                throw new InvalidOperationException(
                    $"No date column found in sheet '{sheetName}'. Provide a date column name or update date column patterns. Available: {string.Join(", ", headers)}");
        }

        var colIdx = headers.IndexOf(dateColName);
        if (colIdx < 0)
            throw new InvalidOperationException($"Date column '{dateColName}' not found. Available: {string.Join(", ", headers)}");

        DateTime? start = TryParseDate(options.StartDate, options.DayFirst);
        DateTime? end = TryParseDate(options.EndDate, options.DayFirst);

        return rows.Where(r =>
        {
            var dt = TryParseCellDate(r[colIdx], options.DayFirst);
            if (dt == null) return false;
            if (start.HasValue && dt < start.Value) return false;
            if (end.HasValue && dt > end.Value) return false;
            return true;
        }).ToList();
    }

    private DateTime? TryParseDate(string? value, bool dayFirst)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var formats = dayFirst
            ? new[] { "dd/MM/yyyy", "dd-MM-yyyy", "yyyy-MM-dd", "MM/dd/yyyy" }
            : new[] { "yyyy-MM-dd", "MM/dd/yyyy", "dd/MM/yyyy", "dd-MM-yyyy" };
        if (DateTime.TryParseExact(value, formats, null, System.Globalization.DateTimeStyles.None, out var dt))
            return dt;
        if (DateTime.TryParse(value, out dt)) return dt;
        return null;
    }

    private DateTime? TryParseCellDate(object? value, bool dayFirst)
    {
        if (value == null) return null;
        if (value is DateTime dt) return dt;
        if (value is double d) return DateTime.FromOADate(d);
        return TryParseDate(value.ToString(), dayFirst);
    }

    private MergeResult BuildOutput(
        List<string> headers,
        List<(string SheetName, List<object?> Cells)> rows,
        MergeOptions options)
    {
        if (options.OutputFormat == OutputFormat.Csv)
        {
            var csv = BuildCsv(headers, rows);
            return new MergeResult
            {
                FileBytes = Encoding.UTF8.GetBytes(csv),
                FileName = "merged_output.csv",
                ContentType = "text/csv"
            };
        }
        else
        {
            var bytes = BuildXlsx(headers, rows);
            return new MergeResult
            {
                FileBytes = bytes,
                FileName = "merged_output.xlsx",
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            };
        }
    }

    private string BuildCsv(List<string> headers, List<(string SheetName, List<object?> Cells)> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(",", headers.Select(EscapeCsv)));
        foreach (var (_, cells) in rows)
            sb.AppendLine(string.Join(",", cells.Select(c => EscapeCsv(FormatCell(c)))));
        return sb.ToString();
    }

    private string EscapeCsv(string? value)
    {
        if (value == null) return string.Empty;
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }

    private string FormatCell(object? value)
    {
        if (value == null) return string.Empty;
        if (value is DateTime dt) return dt.ToString("yyyy-MM-dd HH:mm:ss");
        if (value is double d && d > 40000 && d < 100000)
        {
            try { return DateTime.FromOADate(d).ToString("yyyy-MM-dd"); }
            catch { return d.ToString(); }
        }
        return value.ToString() ?? string.Empty;
    }

    private byte[] BuildXlsx(List<string> headers, List<(string SheetName, List<object?> Cells)> rows)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Merged Output");

        for (int col = 0; col < headers.Count; col++)
        {
            ws.Cells[1, col + 1].Value = headers[col];
            ws.Cells[1, col + 1].Style.Font.Bold = true;
        }

        for (int row = 0; row < rows.Count; row++)
        {
            var cells = rows[row].Cells;
            for (int col = 0; col < cells.Count; col++)
                ws.Cells[row + 2, col + 1].Value = cells[col];
        }

        ws.Cells[ws.Dimension?.Address ?? "A1"].AutoFitColumns();
        return pkg.GetAsByteArray();
    }
}
