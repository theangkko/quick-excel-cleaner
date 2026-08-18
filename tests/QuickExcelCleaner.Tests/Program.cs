using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickExcelCleaner.Services;

static void Assert(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException($"TEST FAILED: {message}");
}

var tempRoot = Path.Combine(Path.GetTempPath(), "QuickExcelCleanerTests", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(tempRoot);
var workbookPath = Path.Combine(tempRoot, "scanner-test.xlsx");

try
{
    CreateWorkbook(workbookPath);

    var scanner = new WorkbookScanner();
    var results = scanner.Scan(workbookPath, 2);

    Assert(results.Any(x => x.Category == "미사용 Style" && x.Target == "cellXfs[2]"),
        "unused style cellXfs[2] was not detected");
    Assert(results.Any(x => x.Category == "중복 Style" && x.Target == "cellXfs[3]" && x.Detail.Contains("cellXfs[1]")),
        "duplicate style cellXfs[3] was not detected");

    Console.WriteLine("SCANNER TESTS PASSED");
}
finally
{
    try { Directory.Delete(tempRoot, recursive: true); } catch { }
}

QuickExcelCleaner.Tests.CleanupTests.Run();
Console.WriteLine("ALL TESTS PASSED");

static void CreateWorkbook(string path)
{
    using var document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.Packaging.SpreadsheetDocumentType.Workbook);
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = new Stylesheet(
        new Fonts(new Font()) { Count = 1U },
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })) { Count = 1U },
        new Borders(new Border()) { Count = 1U },
        new CellStyleFormats(new CellFormat()) { Count = 1U },
        new CellFormats(
            new CellFormat(),
            new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U },
            new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 1U },
            new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U }
        ) { Count = 4U },
        new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U },
        new DifferentialFormats() { Count = 0U },
        new TableStyles { Count = 0U }
    );
    stylesPart.Stylesheet.Save();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(new SheetData(
        new Row(
            new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Hello"), StyleIndex = 1U }
        )
    ));
    worksheetPart.Worksheet.Save();

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
    workbookPart.Workbook.Save();
}
