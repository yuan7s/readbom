using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace readbom;

public partial class MainWindow
{
    private bool IsOfflineFileMode()
    {
        return !string.IsNullOrWhiteSpace(_offlineDocumentPath);
    }

    private void ApplyOperationMode()
    {
        var offline = IsOfflineFileMode();
        SaveToSwButton.Content = offline ? "保存本地文件" : "保存到 SW";
        RelocateFilesButton.Content = offline ? "本地复制重装" : "复制重装";
        SaveContextMenuItem.Header = offline ? "保存本地文件" : "保存到 SW";
    }

    private void OpenLocalDocument(string fullPath, string displayType)
    {
        Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
        AppendLog($"已使用系统默认程序打开{displayType}: {fullPath}");
    }

    private async void OpenOfflineFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "打开本地 SolidWorks 文件",
            Filter = "SolidWorks 文件 (*.sldasm;*.sldprt;*.slddrw)|*.sldasm;*.sldprt;*.slddrw|所有文件 (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        await ReadOfflineFileAsync(dialog.FileName, clearLog: true);
    }

    private async Task ReadOfflineFileAsync(string path, bool clearLog)
    {
        var totalWatch = Stopwatch.StartNew();
        try
        {
            var options = new ReadOptions(
                GetReadMode(),
                GetPropertySourceMode(),
                SkipVirtualCheckBox.IsChecked == true,
                GroupByConfigCheckBox.IsChecked == true);

            if (clearLog)
            {
                ClearLog();
            }

            ClearBomRows();
            _offlineDocumentPath = Path.GetFullPath(path);
            _swApp = null;
            ApplyOperationMode();
            ResetReadProgress();
            OpenOfflineFileButton.IsEnabled = false;
            ReadButton.IsEnabled = false;
            AppendLog($"离线读取本地文件: {_offlineDocumentPath}");
            SetStatus("正在离线读取本地文件");

            Action<string> log = message => Dispatcher.Invoke(() => AppendLog(message));
            Action<ReadProgress> progress = ReportReadProgress;
            var rows = await Task.Run(() => OfflineSolidWorksReader.ReadBom(
                _offlineDocumentPath,
                _propertyMapping.PropertyNames,
                options,
                progress,
                log));

            BindOfflineRows(rows);
            await ResolveBomRelatedFilesAsync();
            ValidateBomMaterials();

            SetStatus($"离线读取完成: {_rows.Count} 项");
            AppendLog($"离线读取完成，项目数: {_rows.Count}");
            var summaryWatch = Stopwatch.StartNew();
            AppendPropertyReadSummary();
            AppendLog($"离线读取计时: 属性统计 {summaryWatch.ElapsedMilliseconds}ms，总计 {totalWatch.ElapsedMilliseconds}ms");
            FinishReadProgress(_rows.Count);
        }
        catch (Exception ex)
        {
            SetStatus("离线读取失败");
            AppendLog($"离线读取失败: {ex.Message}", Brushes.Red);
            ProgressText.Text = "离线读取失败";
            ReadProgressBar.IsIndeterminate = false;
            MessageBox.Show(this, ex.Message, "离线读取失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            OpenOfflineFileButton.IsEnabled = true;
            ReadButton.IsEnabled = true;
        }
    }

    private void BindOfflineRows(IReadOnlyList<BomRow> rows)
    {
        var bindWatch = Stopwatch.StartNew();
        _rows.Clear();
        var index = 1;
        foreach (var item in rows)
        {
            item.Index = index++;
            _rows.Add(item);
        }

        AppendLog($"离线读取计时: 表格绑定 {_rows.Count} 行 {bindWatch.ElapsedMilliseconds}ms");
    }
}
