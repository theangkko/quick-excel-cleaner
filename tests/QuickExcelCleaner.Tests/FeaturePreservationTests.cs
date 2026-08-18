using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickExcelCleaner.Services;

namespace QuickExcelCleaner.Tests;

internal static class FeaturePreservationTests
{
    public static void Run()
    {
        var root = Path.Combine(Path.GetTempPath(), "QuickExcelCleanerFeatureTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var source = Path.Combine(root, "features.xlsx");
        var output = Path.Combine(root, "features_clean.xlsx");

        try
        {
            CreateWorkbook(source);

            var cleaner = new ExcelCleanupService();
            cleaner.Clean(source, output, new CleanupOptions(
                RemoveUnusedStyles: true,
                MergeDuplicateStyles: true,
                RemoveTinyObjects: false));

            ExcelCleanupService.ValidateWorkbook(output);

            using var document = SpreadsheetDocument.Open(output, false);
            var workbook = document.WorkbookPart!;
            var sheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single();
            var part = (WorksheetPart)workbook.GetPartById(sheet.Id!);
            var worksheet = part.Worksheet;

            var merge = worksheet.Elements<MergeCells>().SingleOrDefault();
            Assert(merge?.Elements<MergeCell>().Any(x => x.Reference?.Value == "A1:B1") == true, "merged cell A1:B1 was not preserved");

            var sheetView = worksheet.Elements<SheetViews>().Single().Elements<SheetView>().Single();
            var pane = sheetView.Elements<Pane>().Single();
            Assert(pane.TopLeftCell?.Value == "B2", "freeze pane TopLeftCell was not preserved");
            Assert(pane.State?.Value == PaneStateValues.Frozen, "freeze pane state was not preserved");

            var hiddenRow = worksheet.Descendants<Row>().Single(x => x.RowIndex?.Value == 3U);
            Assert(hiddenRow.Hidden?.Value == true, "hidden row was not preserved");

            var hiddenColumn = worksheet.Descendants<Column>().Single(x => x.Min?.Value == 3U);
            Assert(hiddenColumn.Hidden?.Value == true, "hidden column was not preserved");

            var conditionalFormatting = worksheet.Elements<ConditionalFormatting>().SingleOrDefault();
            Assert(conditionalFormatting?.Elements<ConditionalFormattingRule>().Any() == true,
                "conditional formatting was not preserved");

            var styles = workbook.WorkbookStylesPart!.Stylesheet!;
            var namedStyle = styles.CellStyles?.Elements<CellStyle>().SingleOrDefault(x => x.Name?.Value == "CleanerTestStyle");
            Assert(namedStyle is not null, "named cell style was not preserved");

            var cell = worksheet.Descendants<Cell>().Single(x => x.CellReference?.Value == "A1");
            Assert(cell.CellValue?.Text == "KEEP-FEATURES", "cell value was not preserved");

            Console.WriteLine("FEATURE PRESERVATION TESTS PASSED");
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch { }
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
            new CellStyleFormats(new CellFormat(), new CellFormat()) { Count = 2U },
            new CellFormats(
                new CellFormat(),
                new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U }
            ) { Count = 2U },
            new CellStyles(
                new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U },
                new CellStyle { Name = "CleanerTestStyle", FormatId = 1U }) { Count = 2U },
            new DifferentialFormats(new DifferentialFormat(new Font(new Bold()))) { Count = 1U },
            new TableStyles { Count = 0U }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row { RowIndex = 1U,
                ChildElements =
                {
                    new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("KEEP-FEATURES"), StyleIndex = 1U },
                    new Cell { CellReference = "B1", DataType = CellValues.String, CellValue = new CellValue("MERGED") }
                } },
            new Row { RowIndex = 2U },
            new Row { RowIndex = 3U, Hidden = true }
        );

        var worksheet = new Worksheet(sheetData);

        worksheet.Append(
            new SheetViews(new SheetView { WorkbookViewId = 0U,
                Pane = new Pane { XSplit = 1D, YSplit = 1D, TopLeftCell = "B2", ActivePane = PaneValues.BottomRight, State = PaneStateValues.Frozen },
                Selection = new Selection { Pane = PaneValues.BottomRight, ActiveCell = "B2", Sqref = "B2" } }),
            new MergeCells(new MergeCell { Reference = "A1:B1" }),
            new Columns(new Column { Min = 3U, Max = 3U, Hidden = true }),
            new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A10") }),
                ChildElements = { new ConditionalFormattingRule { Type = ConditionalFormatValues.Expression, Formula = new Formula("1=1") } } }
        );

        worksheetPart.Worksheet = worksheet;
        worksheet.Save();

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Features" });
        workbookPart.Workbook.Save();
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException($"TEST FAILED: {message}");
    }
}
