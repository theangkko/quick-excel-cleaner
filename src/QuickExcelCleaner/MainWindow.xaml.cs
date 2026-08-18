using Microsoft.Win32;
using QuickExcelCleaner.Models;
using QuickExcelCleaner.Services;
using System.Collections.ObjectModel;

namespace QuickExcelCleaner;

public partial class MainWindow : System.Windows.Window
{
    private readonly WorkbookScanner _scanner = new();
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
        StatusText.Text = "검사 준비 완료.";
        _results.Clear();
    }

    private async void ScanButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_selectedFile is null) return;
        if (!double.TryParse(PixelThresholdText.Text, out var threshold) || threshold <= 0)
        {
            System.Windows.MessageBox.Show("픽셀 기준은 0보다 큰 숫자로 입력하세요.", "Quick Excel Cleaner");
            return;
        }

        try
        {
            OpenButton.IsEnabled = false;
            ScanButton.IsEnabled = false;
            StatusText.Text = "검사 중...";

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
            OpenButton.IsEnabled = true;
            ScanButton.IsEnabled = _selectedFile is not null;
        }
    }
}
