namespace QuickExcelCleaner.Models;

public sealed record CleanerResult(string Category, string Sheet, string Target, string Detail);

public sealed record WorkbookScanSummary(
    string FileName,
    int WorksheetCount,
    int CellFormatCount,
    int UsedStyleCount,
    int UnusedStyleCount,
    int DuplicateStyleCount,
    int DrawingCount,
    int TinyObjectCount);
