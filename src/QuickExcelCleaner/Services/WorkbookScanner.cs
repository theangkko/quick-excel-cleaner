using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickExcelCleaner.Models;

namespace QuickExcelCleaner.Services;

public sealed class WorkbookScanner
{
    private const double EmuPerPixel = 9525.0;
    private const double ApproximateCellPixels = 64.0;

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

            foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
            {
                if (row.StyleIndex?.Value is uint styleIndex)
                    usedStyles.Add(styleIndex);
            }

            foreach (var column in worksheetPart.Worksheet.Descendants<Columns>().SelectMany(x => x.Elements<Column>()))
            {
                if (column.Style?.Value is uint styleIndex)
                    usedStyles.Add(styleIndex);
            }

            var drawings = worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawings is null) continue;

            foreach (var anchor in drawings.ChildElements)
            {
                if (!TryGetAnchorSize(anchor, out var width, out var height))
                    continue;

                if (width <= tinyPixelThreshold || height <= tinyPixelThreshold)
                {
                    results.Add(new CleanerResult(
                        "작은 객체",
                        sheet.Name?.Value ?? "",
                        anchor.LocalName,
                        $"약 {width:0.##} × {height:0.##} px"));
                }
            }
        }

        if (styles?.CellFormats is not null)
        {
            var formats = styles.CellFormats.Elements<CellFormat>().ToList();
            var signatures = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var i = 0; i < formats.Count; i++)
            {
                if (!usedStyles.Contains((uint)i))
                    results.Add(new CleanerResult("미사용 Style", "", $"cellXfs[{i}]", "현재 셀/행/열에서 직접 참조되지 않음"));

                var signature = formats[i].OuterXml;
                if (signatures.TryGetValue(signature, out var first))
                    results.Add(new CleanerResult("중복 Style", "", $"cellXfs[{i}]", $"cellXfs[{first}]와 동일"));
                else
                    signatures[signature] = i;
            }
        }

        return results;
    }

    private static bool TryGetAnchorSize(DocumentFormat.OpenXml.OpenXmlElement anchor, out double width, out double height)
    {
        width = 0;
        height = 0;

        if (anchor.LocalName == "oneCellAnchor")
        {
            var extent = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent>().FirstOrDefault();
            if (extent?.Cx?.Value is long cx && extent.Cy?.Value is long cy)
            {
                width = cx / EmuPerPixel;
                height = cy / EmuPerPixel;
                return true;
            }

            return false;
        }

        if (anchor.LocalName == "twoCellAnchor")
        {
            var from = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker>().FirstOrDefault();
            var to = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker>().FirstOrDefault();
            if (from is null || to is null) return false;

            var fromCol = ParseInt(from.ColumnId?.Text);
            var toCol = ParseInt(to.ColumnId?.Text);
            var fromRow = ParseInt(from.RowId?.Text);
            var toRow = ParseInt(to.RowId?.Text);
            width = Math.Max(0, toCol - fromCol) * ApproximateCellPixels +
                    Math.Max(0, ParseEmu(to.ColumnOffset?.Text) - ParseEmu(from.ColumnOffset?.Text)) / EmuPerPixel;
            height = Math.Max(0, toRow - fromRow) * ApproximateCellPixels +
                     Math.Max(0, ParseEmu(to.RowOffset?.Text) - ParseEmu(from.RowOffset?.Text)) / EmuPerPixel;
            return true;
        }

        return false;
    }

    private static int ParseInt(string? value) => int.TryParse(value, out var result) ? result : 0;
    private static long ParseEmu(string? value) => long.TryParse(value, out var result) ? result : 0;
}
