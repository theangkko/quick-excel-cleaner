using Microsoft.Win32;
using QuickExcelCleaner.Models;
using QuickExcelCleaner.Services;
using System.Collections.ObjectModel;

namespace QuickExcelCleaner;

public partial class MainWindow : System.Windows.Window
{
    private readonly WorkbookScanner _scanner = new();
    private readonly ExcelCleanupService _cleaner = new();
    private readonly ObservableCollection<CleanerResult> _results = new();
    private string? _selectedFile;

    public MainWindow()
    {
        InitializeComponent();
        ResultsGrid.ItemsSource = _results;
    }

    private void OpenButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            Multiselect = false
        };
        if (dialog.ShowDialog() != true) return;

        _selectedFile = dialog.FileName;
        FileText.Text = _selectedFile;
        ScanButton.IsEnabled = true;
        CleanButton.IsEnabled = true;
        StatusText.Text = "검사 준비 완료.";
        _results.Clear();
    }

    private async void ScanButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_selectedFile is null) return;
        if (!TryGetThreshold(out var threshold)) return;

        try
        {
            SetBusy(true, "검사 중...");

            var results = await Task.Run(() => _scanner.Scan(_selectedFile, threshold));
            _results.Clear();
            foreach (var result in results)
            {
                if (result.Category == "미사용 Style" && UnusedStylesCheck.IsChecked != true) continue;
                if (result.Category == "중복 Style" && DuplicateStylesCheck.IsChecked != true) continue;
                if (result.Category == "작은 객체" && TinyObjectsCheck.IsChecked != true) continue;
                _results.Add(result);
            }

            StatusText.Text = $"검사 완료: {_results.Count:N0}개 항목";
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(ex.Message, "Excel 검사 실패", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            StatusText.Text = "검사 실패.";
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void CleanButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_selectedFile is null) return;
        if (!TryGetThreshold(out var threshold)) return;

        var extension = System.IO.Path.GetExtension(_selectedFile);
        var dialog = new SaveFileDialog
        {
            Filter = extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
                ? "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm"
                : "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = System.IO.Path.GetFileNameWithoutExtension(_selectedFile) + "_clean" + extension,
            OverwritePrompt = true
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            SetBusy(true, "정리 중...");
            var options = new CleanupOptions(
                RemoveUnusedStyles: UnusedStylesCheck.IsChecked == true,
                MergeDuplicateStyles: DuplicateStylesCheck.IsChecked == true,
                RemoveTinyObjects: TinyObjectsCheck.IsChecked == true,
                TinyObjectThresholdPixels: threshold);

            var report = await Task.Run(() => _cleaner.Clean(_selectedFile, dialog.FileName, options));
            await Task.Run(() => ExcelCleanupService.ValidateWorkbook(report.OutputPath));

            StatusText.Text = $"정리 완료: Style {report.OriginalStyleCount:N0} → {report.FinalStyleCount:N0}, " +
                              $"객체 {report.RemovedObjectCount:N0}개 삭제";

            System.Windows.MessageBox.Show(
                $"정리된 파일이 생성되었습니다.\n\n결과: {report.OutputPath}\n백업: {report.BackupPath}\n\n" +
                $"Style: {report.OriginalStyleCount:N0} → {report.FinalStyleCount:N0}\n" +
                $"Style 삭제/병합: {report.RemovedStyleCount:N0}\n" +
                $"참조 재매핑: {report.RemappedCellCount:N0}\n" +
                $"작은 객체 삭제: {report.RemovedObjectCount:N0}",
                "Quick Excel Cleaner",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(ex.Message, "Excel 정리 실패", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            StatusText.Text = "정리 실패.";
        }
        finally
        {
            SetBusy(false);
        }
    }

    private bool TryGetThreshold(out double threshold)
    {
        if (double.TryParse(PixelThresholdText.Text, out threshold) && threshold > 0)
            return true;

        System.Windows.MessageBox.Show("픽셀 기준은 0보다 큰 숫자로 입력하세요.", "Quick Excel Cleaner");
        threshold = 2;
        return false;
    }

    private void SetBusy(bool busy, string? status = null)
    {
        OpenButton.IsEnabled = !busy;
        ScanButton.IsEnabled = !busy && _selectedFile is not null;
        CleanButton.IsEnabled = !busy && _selectedFile is not null;
        if (status is not null)
            StatusText.Text = status;
    }
}
