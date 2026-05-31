using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace readbom;

public partial class MainWindow : Window
{
    private dynamic? _swApp;
    private string? _offlineDocumentPath;
    private readonly ObservableCollection<BomRow> _rows = [];
    private ICollectionView? _view;
    private PropertyMappingConfig _propertyMapping;
    private readonly Dictionary<string, string> _headerFilters = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, HashSet<string>> _headerValueFilters = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<DataGridColumn, string> _dynamicPropertyColumns = new();
    private readonly Dictionary<DataGridColumn, object> _columnOriginalHeaders = new();
    private ThumbnailPreviewWindow? _thumbnailWindow;
    private BomRow? _thumbnailRow;
    private BomRow? _contextMenuRow;
    private bool _thumbnailEnabled = true;
    private bool _headerFilterEnabled;
    private int _gridSizeLevel;

    public MainWindow()
    {
        InitializeComponent();
        BomGrid.ItemsSource = _rows;
        foreach (var column in BomGrid.Columns)
        {
            column.IsReadOnly = true;
        }
        _view = CollectionViewSource.GetDefaultView(BomGrid.ItemsSource);
        try
        {
            _propertyMapping = PropertyMappingConfig.Load();
            AppendLog($"属性映射配置: {_propertyMapping.SourcePath}");
        }
        catch (Exception ex)
        {
            _propertyMapping = PropertyMappingConfig.CreateDefault(PropertyMappingConfig.DefaultPath);
            AppendLog($"属性映射配置校验失败: {ex.Message}", Brushes.Red);
            AppendLog("已使用内置默认属性映射继续运行", Brushes.Red);
            MessageBox.Show(this, ex.Message, "属性映射配置校验失败", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        ConfigurePropertyColumns();
        SetStatus("未连接 SolidWorks");
        ApplyOperationMode();
        Loaded += (_, _) => EnsureThumbnailWindowVisible();
    }

    private void PropertySettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new PropertySettingsWindow(_propertyMapping.PropertyNames)
        {
            Owner = this
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            _propertyMapping = PropertyMappingConfig.Save(dialog.PropertyNames);
            ConfigurePropertyColumns();
            _view?.Refresh();
            AppendLog($"属性设置已保存: {_propertyMapping.SourcePath}");
        }
        catch (Exception ex)
        {
            AppendLog($"属性设置保存失败: {ex.Message}", Brushes.Red);
            MessageBox.Show(this, ex.Message, "属性设置保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ConnectButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            ClearLog();
            ClearBomRows();
            if (!await SolidWorksAddinClient.IsAvailableAsync())
            {
                throw new InvalidOperationException("SW Add-in HTTP 通道不可用。请确认 SolidWorks 已启动并启用 ReadBom.SwAddin。");
            }

            var info = await SolidWorksAddinClient.GetActiveDocumentInfoAsync(message => AppendLog(message));
            _swApp = null;
            _offlineDocumentPath = null;
            ApplyOperationMode();
            SetStatus("已连接 SolidWorks");
            AppendLog("连接成功: SW Add-in HTTP 通道");
            AppendDocumentTitleLog(info.Title ?? string.Empty);
            AppendLog("插件连接状态: 已连接");
        }
        catch (Exception ex)
        {
            SetStatus("连接失败");
            AppendLog($"连接失败: {ex.Message}");
            MessageBox.Show(this, ex.Message, "连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ReadButton.IsEnabled = true;
        }
    }

    private async void ReadButton_OnClick(object sender, RoutedEventArgs e)
    {
        var totalWatch = Stopwatch.StartNew();
        try
        {
            var options = new ReadOptions(
                GetReadMode(),
                GetPropertySourceMode(),
                SkipVirtualCheckBox.IsChecked == true,
                GroupByConfigCheckBox.IsChecked == true
            );

            ResetReadProgress();
            ReadButton.IsEnabled = false;
            AppendLog("开始读取属性/BOM");
            Action<ReadProgress> progress = ReportReadProgress;
            if (IsOfflineFileMode())
            {
                await ReadOfflineFileAsync(_offlineDocumentPath!, clearLog: false);
                return;
            }

            if (!await SolidWorksAddinClient.IsAvailableAsync())
            {
                throw new InvalidOperationException("SW Add-in HTTP 通道不可用。已停用主程序 COM 回退，请在 SolidWorks 中启用 ReadBom.SwAddin 后重试。");
            }

            AppendLog("使用 SW Add-in HTTP 通道读取 BOM");
            var addinWatch = Stopwatch.StartNew();
            var result = await SolidWorksAddinClient.ReadBomAsync(_propertyMapping.PropertyNames, options, progress, message => AppendLog(message));
            AppendLog($"读取计时: Add-in返回并转换完成 {addinWatch.ElapsedMilliseconds}ms");

            var bindWatch = Stopwatch.StartNew();
            _rows.Clear();
            var index = 1;
            foreach (var item in result)
            {
                item.Index = index++;
                _rows.Add(item);
            }
            AppendLog($"读取计时: 表格绑定 {_rows.Count} 行 {bindWatch.ElapsedMilliseconds}ms");
            await ResolveBomRelatedFilesAsync();
            ValidateBomMaterials();

            SetStatus($"读取完成: {_rows.Count} 项");
            AppendLog($"读取完成，项目数: {_rows.Count}");
            var summaryWatch = Stopwatch.StartNew();
            AppendPropertyReadSummary();
            AppendLog($"读取计时: 属性统计 {summaryWatch.ElapsedMilliseconds}ms，总计 {totalWatch.ElapsedMilliseconds}ms");
            FinishReadProgress(_rows.Count);
        }
        catch (Exception ex)
        {
            SetStatus("读取失败");
            AppendLog($"读取失败: {ex.Message}");
            ProgressText.Text = "读取失败";
            ReadProgressBar.IsIndeterminate = false;
            MessageBox.Show(this, ex.Message, "读取失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ReadButton.IsEnabled = true;
        }
    }

    private bool BlockPendingAddinMigration(string commandName)
    {
        var message = $"{commandName} 需要迁移到 SW Add-in 后再启用。主程序 COM 通道已停用。";
        SetStatus("命令待迁移");
        AppendLog(message, Brushes.Red);
        MessageBox.Show(this, message, "命令待迁移", MessageBoxButton.OK, MessageBoxImage.Information);
        return true;
    }

    private async void SaveToSwButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_rows.Count == 0)
            {
                MessageBox.Show(this, "没有可保存的数据，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            BomGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            BomGrid.CommitEdit(DataGridEditingUnit.Row, true);

            var rows = _rows.ToList();
            var propertyNames = _propertyMapping.PropertyNames.ToList();
            var sourceMode = GetPropertySourceMode();

            SaveToSwButton.IsEnabled = false;
            ReadButton.IsEnabled = false;
            ResetSaveProgress();
            AppendLog(IsOfflineFileMode() ? "开始离线保存 TXT 属性到本地文件" : "开始保存 TXT 属性到 SW");
            Action<ReadProgress> progress = ReportReadProgress;

            var changedRows = rows
                .Select(row => new { Row = row, Changes = row.GetChangedProperties(propertyNames) })
                .Where(x => x.Changes.Count > 0)
                .ToList();

            if (changedRows.Count == 0)
            {
                progress(new ReadProgress("没有需要保存的修改项", 1, 1));
                SetStatus("没有需要保存的修改项");
                AppendLog("保存跳过: 没有需要保存的修改项");
                FinishSaveProgress(0, 0);
                return;
            }

            var saveRows = changedRows.Select(x => new AddinSavePropertyRow
            {
                Path = x.Row.FullPath,
                Configuration = x.Row.Configuration,
                DisplayName = string.IsNullOrWhiteSpace(x.Row.FileName) ? x.Row.FullPath : x.Row.FileName,
                Changes = x.Changes.Select(change => new AddinSavePropertyChange
                {
                    Name = change.Name,
                    Value = change.Value
                }).ToList()
            }).ToList();

            SaveResult result;
            if (IsOfflineFileMode())
            {
                Action<string> log = message => Dispatcher.Invoke(() => AppendLog(message));
                result = await Task.Run(() => OfflineSolidWorksReader.SavePropertiesBatch(saveRows, sourceMode, progress, log));
            }
            else
            {
                result = await SolidWorksAddinClient.SavePropertiesBatchAsync(
                    saveRows,
                    sourceMode,
                    progress,
                    message => AppendLog(message));
            }

            if (result.FailedRows == 0)
            {
                foreach (var item in changedRows)
                {
                    item.Row.AcceptSavedChanges(item.Changes);
                }
            }
            else
            {
                AppendLog("保存存在失败项: 为避免误清除修改标记，失败时保留所有修改单元格高亮", Brushes.Red);
            }

            SetStatus($"保存完成: {result.SavedRows}/{result.TotalRows} 个修改行");
            AppendLog($"保存完成: 修改行成功 {result.SavedRows}/{result.TotalRows}，属性 {result.SavedProperties} 个，失败 {result.FailedRows} 项");
            FinishSaveProgress(result.SavedRows, result.TotalRows);
        }
        catch (Exception ex)
        {
            SetStatus("保存失败");
            AppendLog($"保存失败: {ex.Message}", Brushes.Red);
            ProgressText.Text = "保存失败";
            ReadProgressBar.IsIndeterminate = false;
            MessageBox.Show(this, ex.Message, "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SaveToSwButton.IsEnabled = true;
            ReadButton.IsEnabled = true;
        }
    }

    private async void RelocateFilesButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var offlineMode = IsOfflineFileMode();
            if (!offlineMode && BlockPendingAddinMigration("复制重装"))
            {
                return;
            }

            var targetRows = _rows.Skip(1).Where(x => x.IsOutsideMainAssemblyDirectory).ToList();
            if (targetRows.Count == 0)
            {
                MessageBox.Show(this, "没有检测到不在主装配体目录下的文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var confirm = MessageBox.Show(
                this,
                $"将复制 {targetRows.Count} 个文件到主装配体目录，并重装对应组件引用。是否继续？",
                "复制重装",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            var rows = targetRows.ToList();
            Action<ReadProgress> progress = ReportReadProgress;

            RelocateFilesButton.IsEnabled = false;
            ReadButton.IsEnabled = false;
            ResetSaveProgress();
            ProgressText.Text = offlineMode ? "正在本地复制重装..." : "正在复制重装...";
            AppendLog(offlineMode ? "开始本地复制不在主装配体目录下的文件并重装引用" : "开始复制不在主装配体目录下的文件并重装引用");

            RelocateResult result;
            if (offlineMode)
            {
                Action<string> log = message => Dispatcher.Invoke(() => AppendLog(message));
                result = await Task.Run(() => OfflineSolidWorksReader.CopyExternalFilesToMainAssemblyDirectory(
                    _offlineDocumentPath!,
                    _rows[0],
                    rows,
                    progress,
                    log));
            }
            else
            {
                _swApp ??= SolidWorksReader.Connect();
                var swApp = _swApp;
                result = await Task.Run(() => SolidWorksReader.CopyExternalFilesToMainAssemblyDirectory(swApp, _rows[0], rows, progress));
            }

            ValidateBomPaths();
            SetStatus(offlineMode ? $"本地复制重装完成: {result.ReloadedRows}/{result.TotalRows} 项" : $"复制重装完成: {result.ReloadedRows}/{result.TotalRows} 项");
            AppendLog($"{(offlineMode ? "本地复制重装" : "复制重装")}完成: 复制 {result.CopiedFiles} 个，重装 {result.ReloadedRows}/{result.TotalRows} 项，失败 {result.FailedRows} 项");
            FinishSaveProgress(result.ReloadedRows, result.TotalRows);
            ProgressText.Text = $"{(offlineMode ? "本地复制重装" : "复制重装")}完成，成功 {result.ReloadedRows}/{result.TotalRows} 项";
        }
        catch (Exception ex)
        {
            SetStatus("复制重装失败");
            AppendLog($"复制重装失败: {ex.Message}", Brushes.Red);
            ProgressText.Text = "复制重装失败";
            ReadProgressBar.IsIndeterminate = false;
            MessageBox.Show(this, ex.Message, "复制重装失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            RelocateFilesButton.IsEnabled = true;
            ReadButton.IsEnabled = true;
        }
    }

    private void ResetReadProgress()
    {
        ProgressText.Text = "正在读取属性...";
        ReadProgressBar.IsIndeterminate = true;
        ReadProgressBar.Value = 0;
    }

    private void ResetSaveProgress()
    {
        ProgressText.Text = IsOfflineFileMode() ? "正在保存本地文件..." : "正在保存到 SW...";
        ReadProgressBar.IsIndeterminate = true;
        ReadProgressBar.Value = 0;
    }

    private void UpdateReadProgress(ReadProgress progress)
    {
        Dispatcher.Invoke(() =>
        {
            ReadProgressBar.IsIndeterminate = false;
            ReadProgressBar.Maximum = Math.Max(progress.Total, 1);
            ReadProgressBar.Value = Math.Min(progress.Completed, ReadProgressBar.Maximum);
            ProgressText.Text = $"{progress.Message} {progress.Completed}/{progress.Total}";
        });
    }

    private void ReportReadProgress(ReadProgress progress)
    {
        UpdateReadProgress(progress);
        Dispatcher.Invoke(() =>
        {
            var text = $"{progress.Message}: {progress.Completed}/{progress.Total}";
            if (progress.Message.StartsWith("压缩零件: ", StringComparison.OrdinalIgnoreCase))
            {
                AppendLogWithHighlight(text, GetSuppressedPartHighlight(progress.Message), Brushes.Red);
                return;
            }

            AppendLog(text, progress.Message.StartsWith("压缩零件汇总", StringComparison.OrdinalIgnoreCase) ? Brushes.Red : null);
        });
    }

    private void FinishReadProgress(int count)
    {
        ReadProgressBar.IsIndeterminate = false;
        ReadProgressBar.Maximum = 100;
        ReadProgressBar.Value = 100;
        ProgressText.Text = $"完成，共 {count} 项";
    }

    private void FinishSaveProgress(int saved, int total)
    {
        ReadProgressBar.IsIndeterminate = false;
        ReadProgressBar.Maximum = 100;
        ReadProgressBar.Value = 100;
        ProgressText.Text = $"保存完成，成功 {saved}/{total} 项";
    }

    private void ExportCsvButton_OnClick(object sender, RoutedEventArgs e)
    {
        var rows = GetVisibleRows();
        if (rows.Count == 0)
        {
            MessageBox.Show(this, "没有可导出的数据，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "CSV 文件 (*.csv)|*.csv",
            FileName = BuildExportFileName(rows, "csv")
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var headers = BuildExportHeaders(_propertyMapping.PropertyNames);
        var lines = new List<string> { string.Join(",", headers.Select(Escape)) };

        foreach (var row in rows)
        {
            lines.Add(string.Join(",", BuildExportValues(row, _propertyMapping.PropertyNames).Select(Escape)));
        }

        File.WriteAllLines(dialog.FileName, lines, new UTF8Encoding(true));
        AppendLog($"已导出: {dialog.FileName}");
    }

    private void ExportExcelButton_OnClick(object sender, RoutedEventArgs e)
    {
        var rows = GetVisibleRows();
        if (rows.Count == 0)
        {
            MessageBox.Show(this, "没有可导出的数据，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel 文件 (*.xlsx)|*.xlsx",
            FileName = BuildExportFileName(rows, "xlsx")
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        ExportToExcel(rows, _propertyMapping.PropertyNames, dialog.FileName);
        AppendLog($"已导出: {dialog.FileName}");
    }

    private void IncreaseRowHeightButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_gridSizeLevel < 3)
        {
            _gridSizeLevel++;
        }

        ApplyGridRowSize();
    }

    private void DecreaseRowHeightButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_gridSizeLevel > 0)
        {
            _gridSizeLevel--;
        }

        ApplyGridRowSize();
    }

    private void ThumbnailButton_OnClick(object sender, RoutedEventArgs e)
    {
        _thumbnailEnabled = !_thumbnailEnabled;
        ThumbnailButton.Content = _thumbnailEnabled ? "缩略图关" : "缩略图开";

        if (!_thumbnailEnabled)
        {
            _thumbnailWindow?.Hide();
            return;
        }

        EnsureThumbnailWindowVisible();
    }

    private void EnsureThumbnailWindowVisible()
    {
        if (!_thumbnailEnabled)
        {
            return;
        }

        if (_thumbnailWindow is null)
        {
            _thumbnailWindow = new ThumbnailPreviewWindow();
            _thumbnailWindow.Closed += (_, _) =>
            {
                _thumbnailWindow = null;
                _thumbnailEnabled = false;
                ThumbnailButton.Content = "缩略图开";
            };
        }

        PositionThumbnailWindow();
        _thumbnailWindow.Show();
        _thumbnailWindow.WindowState = WindowState.Normal;
        _thumbnailWindow.ShowInTaskbar = true;
        _thumbnailWindow.Activate();
        _thumbnailWindow.Focus();
        UpdateThumbnailWindow();
    }

    private void PositionThumbnailWindow()
    {
        if (_thumbnailWindow is null)
        {
            return;
        }

        var targetLeft = Left + ActualWidth - _thumbnailWindow.Width - 24;
        var targetTop = Top + 96;
        if (WindowState == WindowState.Maximized)
        {
            var workArea = SystemParameters.WorkArea;
            targetLeft = workArea.Right - _thumbnailWindow.Width - 24;
            targetTop = workArea.Top + 96;
        }

        _thumbnailWindow.Left = Math.Max(SystemParameters.WorkArea.Left, targetLeft);
        _thumbnailWindow.Top = Math.Max(SystemParameters.WorkArea.Top, targetTop);
    }

    private void ApplyGridRowSize()
    {
        BomGrid.RowHeight = 28 + _gridSizeLevel * 6;
        BomGrid.FontSize = 12 + _gridSizeLevel;
        BomGrid.Tag = 20 + _gridSizeLevel * 4;
        BomGrid.ColumnHeaderHeight = 36 + _gridSizeLevel * 5;
        IncreaseRowHeightButton.IsEnabled = _gridSizeLevel < 3;
        DecreaseRowHeightButton.IsEnabled = _gridSizeLevel > 0;
    }

    private void HeaderFilterButton_OnClick(object sender, RoutedEventArgs e)
    {
        _headerFilterEnabled = !_headerFilterEnabled;
        HeaderFilterButton.Content = _headerFilterEnabled ? "关闭筛选" : "表头筛选";
        if (!_headerFilterEnabled)
        {
            _headerFilters.Clear();
            _headerValueFilters.Clear();
            ApplyHeaderFilters();
        }

        UpdateHeaderFilterButtons();
        SetStatus(_headerFilterEnabled ? "已开启表头筛选，点击表头筛选按钮设置条件" : $"读取完成: {_rows.Count} 项");
    }

    private void CopyColumnButton_OnClick(object sender, RoutedEventArgs e)
    {
        var option = ShowColumnSelectionDialog("复制列", "复制源列", includeAllOption: false);
        if (option is null || string.IsNullOrWhiteSpace(option.PropertyName))
        {
            return;
        }

        var rows = GetVisibleRows();
        var values = rows.Select(row => row.GetProperty(option.PropertyName)).ToList();
        Clipboard.SetText(string.Join(Environment.NewLine, values));
        AppendLog($"复制列: {option.DisplayName}，{values.Count} 行");
        SetStatus($"已复制列: {option.DisplayName}");
    }

    private void FillColumnButton_OnClick(object sender, RoutedEventArgs e)
    {
        var result = ShowFillColumnDialog();
        if (result is null || string.IsNullOrWhiteSpace(result.Target.PropertyName))
        {
            return;
        }

        var rows = GetVisibleRows();
        foreach (var row in rows)
        {
            row.SetProperty(result.Target.PropertyName, result.Value);
        }

        RefreshBomGridSafely();
        AppendLog($"填充列: {result.Target.DisplayName}，{rows.Count} 行");
        SetStatus($"已填充列: {result.Target.DisplayName}");
    }

    private async void CalculateBlankSizeButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_rows.Count == 0)
            {
                MessageBox.Show(this, "没有可计算的数据，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (IsOfflineFileMode())
            {
                AppendLog("下料尺寸计算跳过: 离线本地文件模式不连接 SolidWorks，无法获取几何包围盒。", Brushes.Red);
                MessageBox.Show(this, "离线本地文件模式不连接 SolidWorks，无法计算下料尺寸。", "下料尺寸计算", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var target = ShowBlankSizeColumnDialog();
            if (target is null || string.IsNullOrWhiteSpace(target.PropertyName))
            {
                return;
            }

            var targetRows = GetVisibleRows()
                .Where(row => string.IsNullOrWhiteSpace(row.GetProperty(target.PropertyName)))
                .ToList();
            if (targetRows.Count == 0)
            {
                MessageBox.Show(this, $"当前筛选范围内没有 [{target.DisplayName}] 为空的项目。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Action<ReadProgress> progress = ReportReadProgress;

            CalculateBlankSizeButton.IsEnabled = false;
            ReadButton.IsEnabled = false;
            ResetSaveProgress();
            AppendLog($"开始计算下料尺寸: 目标列 [{target.DisplayName}]，空值 {targetRows.Count} 项");

            var response = await SolidWorksAddinClient.GetBoxBatchAsync(
                targetRows.Select(row => new AddinBlankSizeRow
                {
                    Path = row.FullPath,
                    Configuration = row.Configuration,
                    DisplayName = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName
                }).ToList(),
                progress,
                message => AppendLog(message));

            foreach (var item in response.Results.Where(item => item.Success))
            {
                var row = targetRows.FirstOrDefault(x => string.Equals(Path.GetFullPath(x.FullPath), Path.GetFullPath(item.Path ?? string.Empty), StringComparison.OrdinalIgnoreCase));
                if (row is null)
                {
                    continue;
                }

                var blankSize = FormatBlankSizeFromBox(item.Box);
                if (!string.IsNullOrWhiteSpace(blankSize))
                {
                    row.SetProperty(target.PropertyName, blankSize);
                }
            }

            RefreshBomGridSafely();
            SetStatus($"下料尺寸计算完成: {response.UpdatedRows}/{response.TotalRows} 项");
            AppendLog($"下料尺寸计算完成: 成功 {response.UpdatedRows}/{response.TotalRows}，失败 {response.FailedRows} 项");
            FinishSaveProgress(response.UpdatedRows, response.TotalRows);
        }
        catch (Exception ex)
        {
            SetStatus("下料尺寸计算失败");
            AppendLog($"下料尺寸计算失败: {ex.Message}", Brushes.Red);
            ProgressText.Text = "下料尺寸计算失败";
            ReadProgressBar.IsIndeterminate = false;
            MessageBox.Show(this, ex.Message, "下料尺寸计算失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            CalculateBlankSizeButton.IsEnabled = true;
            ReadButton.IsEnabled = true;
        }
    }

    private void FindReplaceButton_OnClick(object sender, RoutedEventArgs e)
    {
        var result = ShowFindReplaceDialog();
        if (result is null || string.IsNullOrEmpty(result.FindText))
        {
            return;
        }

        var targetProperties = string.IsNullOrWhiteSpace(result.Target.PropertyName)
            ? GetDynamicColumnOptions(includeAllOption: false).Select(x => x.PropertyName).ToList()
            : [result.Target.PropertyName];

        var changed = 0;
        foreach (var row in GetVisibleRows())
        {
            foreach (var propertyName in targetProperties)
            {
                var oldValue = row.GetProperty(propertyName);
                if (string.IsNullOrEmpty(oldValue) || !oldValue.Contains(result.FindText, StringComparison.Ordinal))
                {
                    continue;
                }

                row.SetProperty(propertyName, oldValue.Replace(result.FindText, result.ReplaceText, StringComparison.Ordinal));
                changed++;
            }
        }

        RefreshBomGridSafely();
        AppendLog($"查找替换: {changed} 个单元格，查找 [{result.FindText}]");
        SetStatus($"查找替换完成: {changed} 个单元格");
    }

    private static string FormatBlankSizeFromBox(IReadOnlyList<double> values)
    {
        if (values.Count < 6)
        {
            return string.Empty;
        }

        var sizes = new[]
        {
            Math.Round(Math.Abs(values[3] - values[0]) * 1000, 1),
            Math.Round(Math.Abs(values[4] - values[1]) * 1000, 1),
            Math.Round(Math.Abs(values[5] - values[2]) * 1000, 1)
        };
        Array.Sort(sizes);
        Array.Reverse(sizes);
        return string.Join("x", sizes.Select(x => x.ToString("0.#", CultureInfo.InvariantCulture)));
    }

    private void BomGrid_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        var cell = FindParent<DataGridCell>(e.OriginalSource as DependencyObject);
        if (cell?.DataContext is BomRow row)
        {
            _contextMenuRow = row;
        }

        if (!_headerFilterEnabled)
        {
            return;
        }

        var header = FindParent<DataGridColumnHeader>(e.OriginalSource as DependencyObject);
        if (header?.Column is null)
        {
            return;
        }

        var propertyName = GetColumnPropertyName(header.Column);
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return;
        }

        OpenHeaderFilterMenu(header, propertyName);
        e.Handled = true;
    }

    private void HeaderFilterGlyph_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: DataGridColumn column })
        {
            return;
        }

        var header = FindColumnHeader(column);
        if (header is null)
        {
            return;
        }

        var propertyName = GetColumnPropertyName(column);
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return;
        }

        OpenHeaderFilterMenu(header, propertyName);
        e.Handled = true;
    }

    private void BomGrid_OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.F2 || e.Key == Key.Enter)
        {
            if (BeginCurrentDynamicCellEdit(e))
            {
                e.Handled = true;
            }

            return;
        }

        if (e.Key == Key.Delete)
        {
            if (ClearSelectedDynamicCells())
            {
                e.Handled = true;
            }

            return;
        }

        if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
        {
            return;
        }

        if (e.Key == Key.A)
        {
            BomGrid.SelectAllCells();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.C)
        {
            if (CopySelectedCellsToClipboard())
            {
                e.Handled = true;
            }

            return;
        }

        if (e.Key == Key.V)
        {
            if (PasteClipboardToDynamicCells())
            {
                e.Handled = true;
            }
        }
    }

    private void BomGrid_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (string.IsNullOrEmpty(e.Text) || Keyboard.Modifiers != ModifierKeys.None)
        {
            return;
        }

        if (BeginCurrentDynamicCellEdit(e))
        {
            e.Handled = true;
        }
    }

    private void BomGrid_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        var cell = FindParent<DataGridCell>(e.OriginalSource as DependencyObject);
        if (cell is null)
        {
            return;
        }

        if (!cell.IsSelected)
        {
            BomGrid.SelectedCells.Clear();
            cell.IsSelected = true;
            cell.Focus();
            BomGrid.CurrentCell = new DataGridCellInfo(cell);
        }

        _contextMenuRow = cell.DataContext as BomRow;
    }

    private async void OpenModelMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        var row = GetCurrentBomRow();
        if (row is null)
        {
            AppendLog("打开模型失败: 未选中 BOM 行", Brushes.Red);
            return;
        }

        await OpenSolidWorksDocumentAsync(row.FullPath, "模型");
    }

    private async void OpenDrawingMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        var row = GetCurrentBomRow();
        if (row is null)
        {
            AppendLog("打开工程图失败: 未选中 BOM 行", Brushes.Red);
            return;
        }

        var drawingPath = ResolveDrawingPath(row);
        if (string.IsNullOrWhiteSpace(drawingPath))
        {
            var displayName = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName;
            AppendLog($"打开工程图失败: 未找到同名工程图 [{displayName}]", Brushes.Red);
            return;
        }

        await OpenSolidWorksDocumentAsync(drawingPath, "工程图");
    }

    private BomRow? GetCurrentBomRow()
    {
        if (_contextMenuRow is not null)
        {
            return _contextMenuRow;
        }

        if (BomGrid.CurrentCell.Item is BomRow currentCellRow)
        {
            return currentCellRow;
        }

        if (BomGrid.CurrentItem is BomRow currentRow)
        {
            return currentRow;
        }

        return BomGrid.SelectedCells
            .Select(cell => cell.Item)
            .OfType<BomRow>()
            .FirstOrDefault();
    }

}
