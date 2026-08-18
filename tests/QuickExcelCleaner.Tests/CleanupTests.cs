using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using QuickExcelCleaner.Services;

namespace QuickExcelCleaner.Tests;

internal static class CleanupTests
{
    public static void Run()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), "QuickExcelCleanerCleanupTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        var source = Path.Combine(tempRoot, "source.xlsx");
        var output = Path.Combine(tempRoot, "clean.xlsx");

        try
        {
            CreateWorkbook(source);

            var service = new ExcelCleanupService();
            var report = service.Clean(
                source,
                output,
                new CleanupOptions(RemoveUnusedStyles: true, MergeDuplicateStyles: true, RemoveTinyObjects: true, TinyObjectThresholdPixels: 2));

            Assert(File.Exists(report.BackupPath), "backup was not created");
            Assert(File.Exists(output), "clean output was not created");
            Assert(report.OriginalStyleCount == 4, $"expected 4 styles, got {report.OriginalStyleCount}");
            Assert(report.FinalStyleCount == 2, $"expected 2 final styles, got {report.FinalStyleCount}");
            Assert(report.RemovedStyleCount == 2, $"expected 2 removed styles, got {report.RemovedStyleCount}");
            Assert(report.RemappedCellCount == 1, $"expected 1 remapped reference, got {report.RemappedCellCount}");
            Assert(report.RemovedObjectCount == 1, $"expected 1 tiny object removed, got {report.RemovedObjectCount}");

            ExcelCleanupService.ValidateWorkbook(output);

            using var document = SpreadsheetDocument.Open(output, false);
            var workbook = document.WorkbookPart!;
            var styles = workbook.WorkbookStylesPart!.Stylesheet!;
            Assert(styles.CellFormats!.Elements<CellFormat>().Count() == 2, "output cellXfs count mismatch");
            var sheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single();
            var worksheetPart = (WorksheetPart)workbook.GetPartById(sheet.Id!);
            var cell = worksheetPart.Worksheet.Descendants<Cell>().Single();
            Assert(cell.StyleIndex?.Value == 1U, $"expected cell style 1, got {cell.StyleIndex?.Value}");
            Assert(worksheetPart.DrawingsPart?.WorksheetDrawing?.ChildElements.Any() != true, "tiny drawing object remains");

            Console.WriteLine("CLEANUP TESTS PASSED");
        }
        finally
        {
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }

    private static void CreateWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
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
            new Row(new Cell
            {
                CellReference = "A1",
                DataType = CellValues.String,
                CellValue = new CellValue("Hello"),
                StyleIndex = 3U
            })
        ));

        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        drawingsPart.WorksheetDrawing.AppendChild(
            new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 9525L, Cy = 9525L },
                new Xdr.ClientData()));
        drawingsPart.WorksheetDrawing.Save();

        worksheetPart.Worksheet.Append(new Drawing
        {
            Id = worksheetPart.GetIdOfPart(drawingsPart)
        });
        worksheetPart.Worksheet.Save();

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = "Sheet1"
        });
        workbookPart.Workbook.Save();
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException($"TEST FAILED: {message}");
    }
}
