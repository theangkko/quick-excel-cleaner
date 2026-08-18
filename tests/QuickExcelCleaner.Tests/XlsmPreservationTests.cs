using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickExcelCleaner.Services;

namespace QuickExcelCleaner.Tests;

internal static class XlsmPreservationTests
{
    public static void Run()
    {
        var root = Path.Combine(Path.GetTempPath(), "QuickExcelCleanerXlsmTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var sourceXlsx = Path.Combine(root, "macro-source.xlsx");
        var sourceXlsm = Path.Combine(root, "macro-source.xlsm");
        var outputXlsm = Path.Combine(root, "macro-clean.xlsm");

        try
        {
            CreateWorkbook(sourceXlsx);
            File.Copy(sourceXlsx, sourceXlsm);

            using (var zip = ZipFile.Open(sourceXlsm, ZipArchiveMode.Update))
            {
                var entry = zip.CreateEntry("xl/vbaProject.bin", CompressionLevel.NoCompression);
                using var stream = entry.Open();
                var marker = System.Text.Encoding.ASCII.GetBytes("QUICK-EXCEL-CLEANER-VBA-MARKER");
                stream.Write(marker, 0, marker.Length);
            }

            var cleaner = new ExcelCleanupService();
            var report = cleaner.Clean(sourceXlsm, outputXlsm, new CleanupOptions(
                RemoveUnusedStyles: true,
                MergeDuplicateStyles: true,
                RemoveTinyObjects: false));

            Assert(File.Exists(report.OutputPath), "xlsm output was not created");
            ExcelCleanupService.ValidateWorkbook(report.OutputPath);

            using (var zip = ZipFile.OpenRead(report.OutputPath))
            {
                var entry = zip.GetEntry("xl/vbaProject.bin");
                Assert(entry is not null, "vbaProject.bin was not preserved");
                using var stream = entry!.Open();
                using var memory = new MemoryStream();
                stream.CopyTo(memory);
                var marker = System.Text.Encoding.ASCII.GetString(memory.ToArray());
                Assert(marker == "QUICK-EXCEL-CLEANER-VBA-MARKER", "vbaProject.bin contents changed");
            }

            using var document = SpreadsheetDocument.Open(report.OutputPath, false);
            var sheet = document.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single();
            var part = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!);
            Assert(part.Worksheet.Descendants<Cell>().Single().CellValue?.Text == "MACRO-KEEP",
                "xlsm workbook content was not preserved");

            Console.WriteLine("XLSM PRESERVATION TESTS PASSED");
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

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            new Row(new Cell
            {
                CellReference = "A1",
                DataType = CellValues.String,
                CellValue = new CellValue("MACRO-KEEP")
            })));
        worksheetPart.Worksheet.Save();

        workbookPart.Workbook.AppendChild(new Sheets(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = "MacroSheet"
        }));
        workbookPart.Workbook.Save();
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException($"TEST FAILED: {message}");
    }
}
