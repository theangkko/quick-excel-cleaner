using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickExcelCleaner.Models;

namespace QuickExcelCleaner.Services;

public sealed class WorkbookScanner
{
    public IReadOnlyList<CleanerResult> Scan(string path, double tinyPixelThreshold)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var workbook = document.WorkbookPart ?? throw new InvalidDataException("WorkbookPart가 없습니다.");
        var styles = workbook.WorkbookStylesPart?.Stylesheet;
        var results = new List<CleanerResult>();

        var usedStyles = new HashSet<uint> { 0 };
        foreach (var sheet in workbook.Workbook.Sheets?.Elements<Sheet>() ?? [])
        {
            var worksheetPart = (WorksheetPart)workbook.GetPartById(sheet.Id!);
            foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
            {
                if (cell.StyleIndex?.Value is uint styleIndex)
                    usedStyles.Add(styleIndex);
            }

            var drawings = worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawings is null) continue;

            foreach (var anchor in drawings.ChildElements)
            {
                var from = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker>().FirstOrDefault();
                var to = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker>().FirstOrDefault();
                if (from is null || to is null) continue;

                var width = EstimatePixels(from, to, horizontal: true);
                var height = EstimatePixels(from, to, horizontal: false);
                if (width <= tinyPixelThreshold || height <= tinyPixelThreshold)
                {
                    results.Add(new CleanerResult("작은 객체", sheet.Name?.Value ?? "", anchor.LocalName,
                        $"약 {width:0.##} × {height:0.##} px"));
                }
            }
        }

        if (styles?.CellFormats is not null)
        {
            var formats = styles.CellFormats.Elements<CellFormat>().ToList();
            var signatures = new Dictionary<string, int>();
            for (var i = 0; i < formats.Count; i++)
            {
                if (!usedStyles.Contains((uint)i))
                    results.Add(new CleanerResult("미사용 Style", "", $"cellXfs[{i}]", "현재 셀에서 직접 참조되지 않음"));

                var signature = formats[i].OuterXml;
                if (signatures.TryGetValue(signature, out var first))
                    results.Add(new CleanerResult("중복 Style", "", $"cellXfs[{i}]", $"cellXfs[{first}]와 동일"));
                else
                    signatures[signature] = i;
            }
        }

        return results;
    }

    private static double EstimatePixels(DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker from,
        DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker to, bool horizontal)
    {
        var fromCol = from.ColumnId?.Text is { } fc && int.TryParse(fc, out var fci) ? fci : 0;
        var toCol = to.ColumnId?.Text is { } tc && int.TryParse(tc, out var tci) ? tci : 0;
        var fromRow = from.RowId?.Text is { } fr && int.TryParse(fr, out var fri) ? fri : 0;
        var toRow = to.RowId?.Text is { } tr && int.TryParse(tr, out var tri) ? tri : 0;
        var fromOffset = horizontal ? ParseEmu(from.ColumnOffset?.Text) : ParseEmu(from.RowOffset?.Text);
        var toOffset = horizontal ? ParseEmu(to.ColumnOffset?.Text) : ParseEmu(to.RowOffset?.Text);
        var cellCount = horizontal ? Math.Max(0, toCol - fromCol) : Math.Max(0, toRow - fromRow);
        const double defaultCellPx = 64;
        return Math.Max(0, cellCount * defaultCellPx + (toOffset - fromOffset) / 9525.0);
    }

    private static double ParseEmu(string? value) => long.TryParse(value, out var result) ? result : 0;
}
