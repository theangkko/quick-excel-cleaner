using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using QuickExcelCleaner.Services;

namespace QuickExcelCleaner.Tests;

internal static class ComplexWorkbookTests
{
    public static void Run()
    {
        var root = Path.Combine(Path.GetTempPath(), "QuickExcelCleanerComplexTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var source = Path.Combine(root, "complex.xlsx");
        var output = Path.Combine(root, "complex_clean.xlsx");

        try
        {
            CreateWorkbook(source);

            var scanner = new WorkbookScanner();
            var findings = scanner.Scan(source, 2);

            Assert(findings.Any(x => x.Category == "미사용 Style" && x.Target == "cellXfs[4]"),
                "unused style cellXfs[4] should be detected");
            Assert(findings.Any(x => x.Category == "중복 Style" && x.Target == "cellXfs[3]"),
                "duplicate style cellXfs[3] should be detected");
            Assert(findings.Count(x => x.Category == "작은 객체") == 2,
                "two tiny objects should be detected across two sheets");

            var cleaner = new ExcelCleanupService();
            var report = cleaner.Clean(source, output, new CleanupOptions(
                RemoveUnusedStyles: true,
                MergeDuplicateStyles: true,
                RemoveTinyObjects: true,
                TinyObjectThresholdPixels: 2));

            Assert(report.OriginalStyleCount == 5, $"expected 5 original styles, got {report.OriginalStyleCount}");
            Assert(report.FinalStyleCount == 3, $"expected 3 final styles, got {report.FinalStyleCount}");
            Assert(report.RemovedStyleCount == 2, $"expected 2 removed styles, got {report.RemovedStyleCount}");
            Assert(report.RemappedCellCount == 6, $"expected 6 remapped style references, got {report.RemappedCellCount}");
            Assert(report.RemovedObjectCount == 2, $"expected 2 removed tiny objects, got {report.RemovedObjectCount}");

            ExcelCleanupService.ValidateWorkbook(output);

            using var document = SpreadsheetDocument.Open(output, false);
            var workbook = document.WorkbookPart!;
            var sheets = workbook.Workbook.Sheets!.Elements<Sheet>().ToList();
            Assert(sheets.Count == 2, "worksheet count changed");

            var styles = workbook.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToList();
            var canonicalStyle = styles[1].OuterXml;

            foreach (var sheet in sheets)
            {
                var part = (WorksheetPart)workbook.GetPartById(sheet.Id!);
                var cells = part.Worksheet.Descendants<Cell>().ToList();
                Assert(cells.Any(x => x.CellValue?.Text == "KEEP"), $"KEEP value missing on {sheet.Name}");
                Assert(part.DrawingsPart?.WorksheetDrawing?.ChildElements.Count == 1,
                    $"normal drawing should remain on {sheet.Name}");
            }

            var sheet1 = (WorksheetPart)workbook.GetPartById(sheets[0].Id!);
            var row = sheet1.Worksheet.Descendants<Row>().Single(r => r.RowIndex?.Value == 2U);
            Assert(row.StyleIndex?.Value is uint rowStyleIndex && rowStyleIndex < styles.Count,
                "row style index is missing or invalid");
            Assert(styles[(int)row.StyleIndex!.Value].OuterXml == canonicalStyle,
                "row style was not normalized to the canonical style");

            var column = sheet1.Worksheet.Descendants<Column>().Single();
            Assert(column.Style?.Value is uint columnStyleIndex && columnStyleIndex < styles.Count,
                "column style index is missing or invalid");
            Assert(styles[(int)column.Style!.Value].OuterXml == canonicalStyle,
                "column style was not normalized to the canonical style");

            var sheet2 = (WorksheetPart)workbook.GetPartById(sheets[1].Id!);
            var numberCell = sheet2.Worksheet.Descendants<Cell>().Single();
            Assert(numberCell.StyleIndex?.Value == 2U, "distinct style 2 was not preserved");

            Console.WriteLine("COMPLEX WORKBOOK TESTS PASSED");
        }
        finally
        {
            try { Directory.Delete(root, true); } catch { }
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
                new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 2U },
                new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U },
                new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 3U }
            ) { Count = 5U },
            new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U },
            new DifferentialFormats() { Count = 0U },
            new TableStyles { Count = 0U }
        );
        stylesPart.Stylesheet.Save();

        var first = AddSheet(workbookPart);
        var second = AddSheet(workbookPart);

        ConfigureSheet(first, 3U);
        ConfigureSheet(second, 2U);

        workbookPart.Workbook.AppendChild(new Sheets(
            new Sheet { Id = workbookPart.GetIdOfPart(first), SheetId = 1U, Name = "Data1" },
            new Sheet { Id = workbookPart.GetIdOfPart(second), SheetId = 2U, Name = "Data2" }));
        workbookPart.Workbook.Save();
    }

    private static WorksheetPart AddSheet(WorkbookPart workbookPart)
    {
        var part = workbookPart.AddNewPart<WorksheetPart>();
        part.Worksheet = new Worksheet(new SheetData());
        part.Worksheet.Save();
        return part;
    }

    private static void ConfigureSheet(WorksheetPart part, uint cellStyle)
    {
        var sheetData = part.Worksheet.GetFirstChild<SheetData>()!;
        sheetData.Append(
            new Row { RowIndex = 1U },
            new Row(
                new Cell { CellReference = "A2", DataType = CellValues.String, CellValue = new CellValue("KEEP"), StyleIndex = cellStyle }
            ) { RowIndex = 2U, StyleIndex = 3U });

        var columns = new Columns(new Column { Min = 1U, Max = 1U, Style = 3U });
        part.Worksheet.InsertAt(columns, 0);

        var drawingsPart = part.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        drawingsPart.WorksheetDrawing.Append(TinyAnchor(), NormalAnchor());
        drawingsPart.WorksheetDrawing.Save();

        part.Worksheet.Append(new Drawing { Id = part.GetIdOfPart(drawingsPart) });
        part.Worksheet.Save();
    }

    private static Xdr.OneCellAnchor TinyAnchor() => new(
        new Xdr.FromMarker(
            new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"),
            new Xdr.RowId("0"), new Xdr.RowOffset("0")),
        new Xdr.Extent { Cx = 9525L, Cy = 9525L },
        new Xdr.ClientData());

    private static Xdr.OneCellAnchor NormalAnchor() => new(
        new Xdr.FromMarker(
            new Xdr.ColumnId("2"), new Xdr.ColumnOffset("0"),
            new Xdr.RowId("2"), new Xdr.RowOffset("0")),
        new Xdr.Extent { Cx = 952500L, Cy = 952500L },
        new Xdr.ClientData());

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException($"TEST FAILED: {message}");
    }
}
