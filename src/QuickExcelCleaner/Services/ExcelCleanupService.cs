using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickExcelCleaner.Models;

namespace QuickExcelCleaner.Services;

public sealed record CleanupOptions(
    bool RemoveUnusedStyles = true,
    bool MergeDuplicateStyles = true,
    bool RemoveTinyObjects = true,
    double TinyObjectThresholdPixels = 2.0);

public sealed record CleanupReport(
    string OutputPath,
    string BackupPath,
    int OriginalStyleCount,
    int FinalStyleCount,
    int RemovedStyleCount,
    int RemappedCellCount,
    int RemovedObjectCount);

public sealed class ExcelCleanupService
{
    private const double EmuPerPixel = 9525.0;
    private const double ApproximateCellPixels = 64.0;

    public CleanupReport Clean(string sourcePath, string outputPath, CleanupOptions options)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("Excel 파일을 찾을 수 없습니다.", sourcePath);

        var sourceExtension = Path.GetExtension(sourcePath).ToLowerInvariant();
        if (sourceExtension is not ".xlsx" and not ".xlsm")
            throw new InvalidDataException("지원 형식은 .xlsx 또는 .xlsm 입니다.");

        var fullSource = Path.GetFullPath(sourcePath);
        var fullOutput = Path.GetFullPath(outputPath);
        if (string.Equals(fullSource, fullOutput, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("원본 파일과 결과 파일은 같을 수 없습니다.");

        Directory.CreateDirectory(Path.GetDirectoryName(fullOutput)!);
        var backupPath = CreateBackup(fullSource);
        File.Copy(fullSource, fullOutput, overwrite: true);

        try
        {
            using (var document = SpreadsheetDocument.Open(fullOutput, true))
            {
                var workbook = document.WorkbookPart ?? throw new InvalidDataException("WorkbookPart가 없습니다.");
                var styleReport = CleanupStyles(workbook, options);
                var removedObjects = options.RemoveTinyObjects
                    ? RemoveTinyObjects(workbook, options.TinyObjectThresholdPixels)
                    : 0;

                workbook.Workbook.Save();
                return new CleanupReport(
                    fullOutput,
                    backupPath,
                    styleReport.OriginalCount,
                    styleReport.FinalCount,
                    styleReport.RemovedCount,
                    styleReport.RemappedCount,
                    removedObjects);
            }
        }
        catch
        {
            TryDelete(fullOutput);
            throw;
        }

        static string CreateBackup(string source)
        {
            var sourceDirectory = Path.GetDirectoryName(source)!;
            var backupDirectory = Path.Combine(sourceDirectory, "ExcelCleaner_Backup");
            Directory.CreateDirectory(backupDirectory);
            var baseName = Path.GetFileNameWithoutExtension(source);
            var extension = Path.GetExtension(source);
            var candidate = Path.Combine(backupDirectory, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}_backup{extension}");
            var suffix = 1;
            while (File.Exists(candidate))
            {
                candidate = Path.Combine(backupDirectory, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}_backup_{suffix++}{extension}");
            }

            File.Copy(source, candidate, overwrite: false);
            return candidate;
        }

        static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }

    public static void ValidateWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var workbook = document.WorkbookPart ?? throw new InvalidDataException("정리된 파일에 WorkbookPart가 없습니다.");
        if (workbook.Workbook.Sheets is null || !workbook.Workbook.Sheets.Any())
            throw new InvalidDataException("정리된 파일에 Worksheet가 없습니다.");

        foreach (var sheet in workbook.Workbook.Sheets.Elements<Sheet>())
        {
            var part = (WorksheetPart)workbook.GetPartById(sheet.Id!);
            _ = part.Worksheet;
        }
    }

    private static (int OriginalCount, int FinalCount, int RemovedCount, int RemappedCount) CleanupStyles(
        WorkbookPart workbook,
        CleanupOptions options)
    {
        var stylesPart = workbook.WorkbookStylesPart;
        var cellFormats = stylesPart?.Stylesheet?.CellFormats;
        if (cellFormats is null || cellFormats.Count is null || cellFormats.Count.Value == 0)
            return (0, 0, 0, 0);

        var formats = cellFormats.Elements<CellFormat>().ToList();
        var originalCount = formats.Count;
        var used = CollectUsedStyleIndexes(workbook);

        if (!options.RemoveUnusedStyles && !options.MergeDuplicateStyles)
            return (originalCount, originalCount, 0, 0);

        var canonicalBySignature = new Dictionary<string, int>(StringComparer.Ordinal);
        var keepIndexes = new List<int>();
        var remap = new Dictionary<int, int>();

        for (var i = 0; i < formats.Count; i++)
        {
            if (i != 0 && options.RemoveUnusedStyles && !used.Contains(i))
                continue;

            var signature = formats[i].OuterXml;
            if (options.MergeDuplicateStyles && canonicalBySignature.TryGetValue(signature, out var canonical))
            {
                remap[i] = canonical;
                continue;
            }

            canonicalBySignature[signature] = i;
            keepIndexes.Add(i);
            remap[i] = i;
        }

        var newIndex = new Dictionary<int, uint>();
        var newFormats = new List<CellFormat>();
        for (var i = 0; i < keepIndexes.Count; i++)
        {
            var originalIndex = keepIndexes[i];
            newIndex[originalIndex] = (uint)i;
            newFormats.Add((CellFormat)formats[originalIndex].CloneNode(true));
        }

        var normalizedRemap = new Dictionary<uint, uint>();
        for (var i = 0; i < formats.Count; i++)
        {
            var canonicalOriginal = remap.TryGetValue(i, out var canonical) ? canonical : -1;
            if (canonicalOriginal >= 0 && newIndex.TryGetValue(canonicalOriginal, out var target))
                normalizedRemap[(uint)i] = target;
        }

        var remappedCount = ApplyStyleRemap(workbook, normalizedRemap);

        cellFormats.RemoveAllChildren<CellFormat>();
        foreach (var format in newFormats)
            cellFormats.Append(format);
        cellFormats.Count = (uint)newFormats.Count;
        stylesPart!.Stylesheet!.Save();

        return (originalCount, newFormats.Count, originalCount - newFormats.Count, remappedCount);
    }

    private static HashSet<int> CollectUsedStyleIndexes(WorkbookPart workbook)
    {
        var used = new HashSet<int> { 0 };
        foreach (var sheet in workbook.Workbook.Sheets?.Elements<Sheet>() ?? [])
        {
            var worksheetPart = (WorksheetPart)workbook.GetPartById(sheet.Id!);
            foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
            {
                if (cell.StyleIndex?.Value is uint style)
                    used.Add((int)style);
            }

            foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
            {
                if (row.StyleIndex?.Value is uint style)
                    used.Add((int)style);
            }

            foreach (var column in worksheetPart.Worksheet.Descendants<Columns>().SelectMany(x => x.Elements<Column>()))
            {
                if (column.Style?.Value is uint style)
                    used.Add((int)style);
            }
        }

        return used;
    }

    private static int ApplyStyleRemap(WorkbookPart workbook, IReadOnlyDictionary<uint, uint> remap)
    {
        var changed = 0;
        foreach (var sheet in workbook.Workbook.Sheets?.Elements<Sheet>() ?? [])
        {
            var worksheetPart = (WorksheetPart)workbook.GetPartById(sheet.Id!);

            foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
            {
                if (cell.StyleIndex?.Value is uint style && remap.TryGetValue(style, out var target) && style != target)
                {
                    cell.StyleIndex = target;
                    changed++;
                }
            }

            foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
            {
                if (row.StyleIndex?.Value is uint style && remap.TryGetValue(style, out var target) && style != target)
                {
                    row.StyleIndex = target;
                    changed++;
                }
            }

            foreach (var column in worksheetPart.Worksheet.Descendants<Columns>().SelectMany(x => x.Elements<Column>()))
            {
                if (column.Style?.Value is uint style && remap.TryGetValue(style, out var target) && style != target)
                {
                    column.Style = target;
                    changed++;
                }
            }

            worksheetPart.Worksheet.Save();
        }

        return changed;
    }

    private static int RemoveTinyObjects(WorkbookPart workbook, double thresholdPixels)
    {
        var removed = 0;
        foreach (var sheet in workbook.Workbook.Sheets?.Elements<Sheet>() ?? [])
        {
            var worksheetPart = (WorksheetPart)workbook.GetPartById(sheet.Id!);
            var drawingsPart = worksheetPart.DrawingsPart;
            var drawing = drawingsPart?.WorksheetDrawing;
            if (drawing is null) continue;

            foreach (var anchor in drawing.ChildElements.ToList())
            {
                if (IsTinyAnchor(anchor, thresholdPixels))
                {
                    anchor.Remove();
                    removed++;
                }
            }

            if (removed > 0)
                drawing.Save();
        }

        return removed;
    }

    private static bool IsTinyAnchor(OpenXmlElement anchor, double thresholdPixels)
    {
        var localName = anchor.LocalName;
        if (localName == "oneCellAnchor")
        {
            var ext = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent>().FirstOrDefault();
            if (ext?.Cx?.Value is long cx && ext.Cy?.Value is long cy)
                return cx / EmuPerPixel <= thresholdPixels || cy / EmuPerPixel <= thresholdPixels;
        }

        if (localName == "twoCellAnchor")
        {
            var from = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker>().FirstOrDefault();
            var to = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker>().FirstOrDefault();
            if (from is null || to is null) return false;

            var fromCol = ParseInt(from.ColumnId?.Text);
            var toCol = ParseInt(to.ColumnId?.Text);
            var fromRow = ParseInt(from.RowId?.Text);
            var toRow = ParseInt(to.RowId?.Text);
            var width = Math.Max(0, toCol - fromCol) * ApproximateCellPixels +
                        (ParseEmu(from.ColumnOffset?.Text) < ParseEmu(to.ColumnOffset?.Text)
                            ? (ParseEmu(to.ColumnOffset?.Text) - ParseEmu(from.ColumnOffset?.Text)) / EmuPerPixel
                            : 0);
            var height = Math.Max(0, toRow - fromRow) * ApproximateCellPixels +
                         (ParseEmu(from.RowOffset?.Text) < ParseEmu(to.RowOffset?.Text)
                             ? (ParseEmu(to.RowOffset?.Text) - ParseEmu(from.RowOffset?.Text)) / EmuPerPixel
                             : 0);
            return width <= thresholdPixels || height <= thresholdPixels;
        }

        return false;
    }

    private static int ParseInt(string? value) => int.TryParse(value, out var result) ? result : 0;
    private static long ParseEmu(string? value) => long.TryParse(value, out var result) ? result : 0;
}
