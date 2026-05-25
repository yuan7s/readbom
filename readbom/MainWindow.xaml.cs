using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using System.Collections;
using WpfParagraph = System.Windows.Documents.Paragraph;
using WpfRun = System.Windows.Documents.Run;

namespace readbom;

public partial class MainWindow : Window
{
    private dynamic? _swApp;
    private int _connectedSwPid;
    private readonly ObservableCollection<BomRow> _rows = [];
    private ICollectionView? _view;
    private readonly PropertyMappingConfig _propertyMapping;
    private readonly Dictionary<string, string> _headerFilters = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<DataGridColumn, string> _dynamicPropertyColumns = new();
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
    }

    private void ConnectButton_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            ClearLog();
            ClearBomRows();
            _swApp = SolidWorksReader.Connect();
            var info = SolidWorksReader.CheckConnection(_swApp);
            _connectedSwPid = info.ProcessId;
            SetStatus("已连接 SolidWorks");
            AppendLog($"连接成功。PID: {info.ProcessId}，SW版本: {info.DisplayVersion}，内部版本: {info.Version}，已打开文档: {info.OpenDocumentCount}");
            AppendLog($"SW窗口标题: {(string.IsNullOrWhiteSpace(info.MainWindowTitle) ? "(无)" : info.MainWindowTitle)}");
            AppendDocumentTitleLog(info.ActiveDocumentTitle);
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
        try
        {
            _swApp ??= SolidWorksReader.Connect();
            var swApp = _swApp;
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
            var result = await Task.Run(() => SolidWorksReader.ReadBom(swApp, options, _propertyMapping, progress));
            _rows.Clear();
            var index = 1;
            foreach (var item in result)
            {
                item.Index = index++;
                _rows.Add(item);
            }
            ValidateBomPaths();

            SetStatus($"读取完成: {_rows.Count} 项");
            AppendLog($"读取完成，项目数: {_rows.Count}");
            AppendPropertyReadSummary();
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

            _swApp ??= SolidWorksReader.Connect();
            var swApp = _swApp;
            var rows = _rows.ToList();
            var propertyNames = _propertyMapping.PropertyNames.ToList();
            var sourceMode = GetPropertySourceMode();

            SaveToSwButton.IsEnabled = false;
            ReadButton.IsEnabled = false;
            ResetSaveProgress();
            AppendLog("开始保存 TXT 属性到 SW");
            Action<ReadProgress> progress = ReportReadProgress;

            var result = await Task.Run(() => SolidWorksReader.SaveBomProperties(
                swApp,
                rows,
                propertyNames,
                sourceMode,
                progress));

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

    private void ResetReadProgress()
    {
        ProgressText.Text = "正在读取属性...";
        ReadProgressBar.IsIndeterminate = true;
        ReadProgressBar.Value = 0;
    }

    private void ResetSaveProgress()
    {
        ProgressText.Text = "正在保存到 SW...";
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
            ApplyHeaderFilters();
        }

        SetStatus(_headerFilterEnabled ? "已开启表头筛选，点击列标题设置条件" : $"读取完成: {_rows.Count} 项");
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

        BomGrid.Items.Refresh();
        AppendLog($"填充列: {result.Target.DisplayName}，{rows.Count} 行");
        SetStatus($"已填充列: {result.Target.DisplayName}");
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

        BomGrid.Items.Refresh();
        AppendLog($"查找替换: {changed} 个单元格，查找 [{result.FindText}]");
        SetStatus($"查找替换完成: {changed} 个单元格");
    }

    private void BomGrid_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
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

    private void BomGrid_OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
        {
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

    private ColumnOption? ShowColumnSelectionDialog(string title, string label, bool includeAllOption)
    {
        var options = GetDynamicColumnOptions(includeAllOption);
        if (options.Count == 0)
        {
            MessageBox.Show(this, "没有可操作的 TXT 属性列，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return null;
        }

        var comboBox = CreateColumnComboBox(options);
        SelectDefaultColumn(comboBox);

        var window = CreateToolDialog(title, 340, 150);
        var panel = CreateDialogPanel();
        panel.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 0, 0, 6) });
        panel.Children.Add(comboBox);
        AddDialogButtons(panel, window);
        window.Content = panel;

        return window.ShowDialog() == true ? comboBox.SelectedItem as ColumnOption : null;
    }

    private FillColumnResult? ShowFillColumnDialog()
    {
        var options = GetDynamicColumnOptions(includeAllOption: false);
        if (options.Count == 0)
        {
            MessageBox.Show(this, "没有可填充的 TXT 属性列，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return null;
        }

        var comboBox = CreateColumnComboBox(options);
        SelectDefaultColumn(comboBox);
        var valueBox = new TextBox { MinWidth = 260, Margin = new Thickness(0, 0, 0, 10) };
        if (BomGrid.CurrentCell.Item is BomRow row && comboBox.SelectedItem is ColumnOption option)
        {
            valueBox.Text = row.GetProperty(option.PropertyName);
        }

        comboBox.SelectionChanged += (_, _) =>
        {
            if (BomGrid.CurrentCell.Item is BomRow currentRow && comboBox.SelectedItem is ColumnOption currentOption)
            {
                valueBox.Text = currentRow.GetProperty(currentOption.PropertyName);
            }
        };

        var window = CreateToolDialog("填充列", 380, 210);
        var panel = CreateDialogPanel();
        panel.Children.Add(new TextBlock { Text = "目标列", Margin = new Thickness(0, 0, 0, 6) });
        panel.Children.Add(comboBox);
        panel.Children.Add(new TextBlock { Text = "填充内容", Margin = new Thickness(0, 10, 0, 6) });
        panel.Children.Add(valueBox);
        AddDialogButtons(panel, window);
        window.Content = panel;

        return window.ShowDialog() == true && comboBox.SelectedItem is ColumnOption target
            ? new FillColumnResult(target, valueBox.Text)
            : null;
    }

    private FindReplaceResult? ShowFindReplaceDialog()
    {
        var options = GetDynamicColumnOptions(includeAllOption: true);
        if (options.Count == 0)
        {
            MessageBox.Show(this, "没有可查找替换的 TXT 属性列，请先读取 BOM。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return null;
        }

        var comboBox = CreateColumnComboBox(options);
        SelectDefaultColumn(comboBox);
        var findBox = new TextBox { MinWidth = 260, Margin = new Thickness(0, 0, 0, 10) };
        var replaceBox = new TextBox { MinWidth = 260, Margin = new Thickness(0, 0, 0, 10) };

        var window = CreateToolDialog("查找替换", 380, 280);
        var panel = CreateDialogPanel();
        panel.Children.Add(new TextBlock { Text = "目标列", Margin = new Thickness(0, 0, 0, 6) });
        panel.Children.Add(comboBox);
        panel.Children.Add(new TextBlock { Text = "查找内容", Margin = new Thickness(0, 10, 0, 6) });
        panel.Children.Add(findBox);
        panel.Children.Add(new TextBlock { Text = "替换为", Margin = new Thickness(0, 0, 0, 6) });
        panel.Children.Add(replaceBox);
        AddDialogButtons(panel, window);
        window.Content = panel;

        return window.ShowDialog() == true && comboBox.SelectedItem is ColumnOption target
            ? new FindReplaceResult(target, findBox.Text, replaceBox.Text)
            : null;
    }

    private List<ColumnOption> GetDynamicColumnOptions(bool includeAllOption)
    {
        var options = _dynamicPropertyColumns
            .OrderBy(x => x.Key.DisplayIndex)
            .Select(x => new ColumnOption(x.Value, GetPropertyDisplayName(x.Value)))
            .ToList();

        if (includeAllOption && options.Count > 0)
        {
            options.Insert(0, new ColumnOption(string.Empty, "全部TXT属性列"));
        }

        return options;
    }

    private ComboBox CreateColumnComboBox(IReadOnlyList<ColumnOption> options)
    {
        return new ComboBox
        {
            ItemsSource = options,
            DisplayMemberPath = nameof(ColumnOption.DisplayName),
            SelectedIndex = 0,
            MinWidth = 260,
            Height = 28,
            Margin = new Thickness(0, 0, 0, 10)
        };
    }

    private void SelectDefaultColumn(ComboBox comboBox)
    {
        if (BomGrid.CurrentCell.Column is null)
        {
            return;
        }

        var propertyName = GetDynamicColumnPropertyName(BomGrid.CurrentCell.Column);
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return;
        }

        foreach (var item in comboBox.Items)
        {
            if (item is ColumnOption option && string.Equals(option.PropertyName, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = option;
                return;
            }
        }
    }

    private Window CreateToolDialog(string title, double width, double height)
    {
        return new Window
        {
            Title = title,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.Manual,
            ResizeMode = ResizeMode.NoResize,
            Width = width,
            Height = height
        };
    }

    private static StackPanel CreateDialogPanel()
    {
        return new StackPanel { Margin = new Thickness(14) };
    }

    private static void AddDialogButtons(Panel panel, Window window)
    {
        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 8, 0, 0)
        };
        var okButton = new Button { Content = "确定", Width = 72, Height = 28, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
        var cancelButton = new Button { Content = "取消", Width = 72, Height = 28, IsCancel = true };
        okButton.Click += (_, _) => window.DialogResult = true;
        buttons.Children.Add(okButton);
        buttons.Children.Add(cancelButton);
        panel.Children.Add(buttons);
    }

    private void OpenHeaderFilterMenu(DataGridColumnHeader header, string propertyName)
    {
        var current = _headerFilters.TryGetValue(propertyName, out var value) ? value : string.Empty;
        var textBox = new TextBox
        {
            Width = 180,
            Text = current,
            Margin = new Thickness(0, 6, 0, 8)
        };

        var menu = new ContextMenu
        {
            PlacementTarget = header,
            Placement = PlacementMode.Bottom,
            StaysOpen = true
        };

        var panel = new StackPanel { Width = 220, Margin = new Thickness(10) };
        panel.Children.Add(new TextBlock { Text = $"筛选: {header.Column.Header}", FontWeight = FontWeights.SemiBold });
        panel.Children.Add(textBox);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var applyButton = new Button { Content = "应用", Width = 64, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        var clearButton = new Button { Content = "清除", Width = 64, Height = 28 };
        buttons.Children.Add(applyButton);
        buttons.Children.Add(clearButton);
        panel.Children.Add(buttons);

        applyButton.Click += (_, _) =>
        {
            var filter = textBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(filter))
            {
                _headerFilters.Remove(propertyName);
            }
            else
            {
                _headerFilters[propertyName] = filter;
            }

            ApplyHeaderFilters();
            menu.IsOpen = false;
        };

        clearButton.Click += (_, _) =>
        {
            _headerFilters.Remove(propertyName);
            ApplyHeaderFilters();
            menu.IsOpen = false;
        };

        menu.Items.Add(new MenuItem { Header = panel, StaysOpenOnClick = true });
        menu.IsOpen = true;
        textBox.Focus();
        textBox.SelectAll();
    }

    private void ApplyHeaderFilters()
    {
        if (_view is null)
        {
            return;
        }

        if (_headerFilters.Count == 0)
        {
            _view.Filter = null;
        }
        else
        {
            _view.Filter = item =>
            {
                if (item is not BomRow row)
                {
                    return false;
                }

                foreach (var filter in _headerFilters)
                {
                    if (!ContainsIgnoreCase(GetCellValue(row, filter.Key), filter.Value))
                    {
                        return false;
                    }
                }

                return true;
            };
        }

        _view.Refresh();
        SetStatus(_headerFilters.Count == 0 ? $"读取完成: {_rows.Count} 项" : $"过滤后: {GetVisibleRows().Count} 项");
    }

    private string GetColumnPropertyName(DataGridColumn column)
    {
        var dynamicName = GetDynamicColumnPropertyName(column);
        if (!string.IsNullOrWhiteSpace(dynamicName))
        {
            return $"Properties[{dynamicName}]";
        }

        if (column is DataGridBoundColumn boundColumn && boundColumn.Binding is Binding binding)
        {
            return binding.Path.Path;
        }

        return column.Header?.ToString() switch
        {
            "类型" => nameof(BomRow.DocumentType),
            "工程图" => nameof(BomRow.DrawingStatus),
            _ => string.Empty
        };
    }

    private string GetDynamicColumnPropertyName(DataGridColumn column)
    {
        return _dynamicPropertyColumns.TryGetValue(column, out var propertyName) ? propertyName : string.Empty;
    }

    private static string GetCellValue(BomRow row, string propertyName)
    {
        return propertyName switch
        {
            nameof(BomRow.Index) => row.Index.ToString(CultureInfo.InvariantCulture),
            nameof(BomRow.DocumentType) => row.DocumentType,
            nameof(BomRow.DrawingStatus) => row.DrawingStatus,
            nameof(BomRow.FileName) => row.FileName,
            nameof(BomRow.Configuration) => row.Configuration,
            nameof(BomRow.Quantity) => row.Quantity.ToString(CultureInfo.InvariantCulture),
            nameof(BomRow.Material) => row.Material,
            nameof(BomRow.FullPath) => row.FullPath,
            _ when propertyName.StartsWith("Properties[", StringComparison.Ordinal) && propertyName.EndsWith("]", StringComparison.Ordinal)
                => row.GetProperty(propertyName["Properties[".Length..^1]),
            _ => row.GetProperty(propertyName)
        };
    }

    private static T? FindParent<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current is not null)
        {
            if (current is T typed)
            {
                return typed;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private bool CopySelectedCellsToClipboard()
    {
        var cells = GetOrderedSelectedCells();
        if (cells.Count == 0 && BomGrid.CurrentCell.Item is BomRow currentRow && BomGrid.CurrentCell.Column is not null)
        {
            cells.Add(new DataGridCellInfo(currentRow, BomGrid.CurrentCell.Column));
        }

        if (cells.Count == 0)
        {
            return false;
        }

        var visibleRows = GetVisibleRows();
        var rowIndex = visibleRows
            .Select((row, index) => new { row, index })
            .ToDictionary(x => x.row, x => x.index);

        var lines = cells
            .Where(x => x.Item is BomRow && x.Column is not null)
            .GroupBy(x => rowIndex.TryGetValue((BomRow)x.Item, out var index) ? index : int.MaxValue)
            .OrderBy(x => x.Key)
            .Select(group => string.Join("\t", group
                .OrderBy(x => x.Column.DisplayIndex)
                .Select(x => GetCellValue((BomRow)x.Item, GetColumnPropertyName(x.Column)))))
            .ToList();

        if (lines.Count == 0)
        {
            return false;
        }

        Clipboard.SetText(string.Join(Environment.NewLine, lines));
        return true;
    }

    private bool PasteClipboardToDynamicCells()
    {
        if (!Clipboard.ContainsText())
        {
            return false;
        }

        var text = Clipboard.GetText();
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var cells = GetOrderedSelectedCells();
        var startCell = cells.FirstOrDefault();
        if (!startCell.IsValid && BomGrid.CurrentCell.Item is BomRow)
        {
            startCell = BomGrid.CurrentCell;
        }

        if (!startCell.IsValid || startCell.Item is not BomRow startRow || startCell.Column is null)
        {
            return false;
        }

        var startProperty = GetDynamicColumnPropertyName(startCell.Column);
        if (string.IsNullOrWhiteSpace(startProperty))
        {
            return false;
        }

        var visibleRows = GetVisibleRows();
        var startRowIndex = visibleRows.IndexOf(startRow);
        if (startRowIndex < 0)
        {
            return false;
        }

        var dynamicColumns = BomGrid.Columns
            .Select(column => new { column, propertyName = GetDynamicColumnPropertyName(column) })
            .Where(x => !string.IsNullOrWhiteSpace(x.propertyName))
            .OrderBy(x => x.column.DisplayIndex)
            .ToList();
        var startColumnIndex = dynamicColumns.FindIndex(x => x.column == startCell.Column);
        if (startColumnIndex < 0)
        {
            return false;
        }

        var matrix = ParseClipboardMatrix(text);
        if (matrix.Count == 0)
        {
            return false;
        }

        if (IsSingleClipboardValue(matrix))
        {
            var fillChanged = FillSelectedDynamicCells(cells, matrix[0][0]);
            if (fillChanged > 0)
            {
                BomGrid.Items.Refresh();
                AppendLog($"粘贴TXT属性: {fillChanged} 个单元格");
                return true;
            }
        }

        var changed = 0;
        for (var r = 0; r < matrix.Count; r++)
        {
            var rowIndex = startRowIndex + r;
            if (rowIndex >= visibleRows.Count)
            {
                break;
            }

            var row = visibleRows[rowIndex];
            for (var c = 0; c < matrix[r].Count; c++)
            {
                var columnIndex = startColumnIndex + c;
                if (columnIndex >= dynamicColumns.Count)
                {
                    break;
                }

                row.SetProperty(dynamicColumns[columnIndex].propertyName, matrix[r][c]);
                changed++;
            }
        }

        if (changed == 0)
        {
            return false;
        }

        BomGrid.Items.Refresh();
        AppendLog($"粘贴TXT属性: {changed} 个单元格");
        return true;
    }

    private int FillSelectedDynamicCells(IReadOnlyList<DataGridCellInfo> cells, string value)
    {
        var changed = 0;
        foreach (var cell in cells)
        {
            if (cell.Item is not BomRow row || cell.Column is null)
            {
                continue;
            }

            var propertyName = GetDynamicColumnPropertyName(cell.Column);
            if (string.IsNullOrWhiteSpace(propertyName))
            {
                continue;
            }

            row.SetProperty(propertyName, value);
            changed++;
        }

        return changed;
    }

    private List<DataGridCellInfo> GetOrderedSelectedCells()
    {
        var visibleRows = GetVisibleRows();
        var rowIndex = visibleRows
            .Select((row, index) => new { row, index })
            .ToDictionary(x => x.row, x => x.index);

        return BomGrid.SelectedCells
            .Where(x => x.Item is BomRow && x.Column is not null)
            .OrderBy(x => rowIndex.TryGetValue((BomRow)x.Item, out var index) ? index : int.MaxValue)
            .ThenBy(x => x.Column.DisplayIndex)
            .ToList();
    }

    private static List<List<string>> ParseClipboardMatrix(string text)
    {
        return text
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .TrimEnd('\n')
            .Split('\n')
            .Where(line => line.Length > 0)
            .Select(line => line.Split('\t').ToList())
            .ToList();
    }

    private static bool IsSingleClipboardValue(IReadOnlyList<List<string>> matrix)
    {
        return matrix.Count == 1 && matrix[0].Count == 1;
    }

    private static string Escape(string? value)
    {
        var v = value ?? string.Empty;
        if (v.Contains('"') || v.Contains(',') || v.Contains('\n') || v.Contains('\r'))
        {
            return "\"" + v.Replace("\"", "\"\"") + "\"";
        }

        return v;
    }

    private static string BuildExportFileName(IReadOnlyList<BomRow> rows, string extension)
    {
        var prefix = rows.Count > 0 ? SanitizeFileName(rows[0].FileName) : string.Empty;
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
        var name = string.IsNullOrWhiteSpace(prefix)
            ? $"bom_{timestamp}"
            : $"{prefix}_bom_{timestamp}";
        return $"{name}.{extension.TrimStart('.')}";
    }

    private static string SanitizeFileName(string value)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var cleaned = new string(value
            .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
            .ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "bom" : cleaned;
    }

    private void SetStatus(string text) => StatusText.Text = text;

    private void AppendLog(string text, Brush? foreground = null)
    {
        var paragraph = new WpfParagraph(new WpfRun($"{DateTime.Now:HH:mm:ss} {text}"))
        {
            Margin = new Thickness(0),
            Foreground = foreground ?? Brushes.Black
        };
        LogTextBox.Document.Blocks.Add(paragraph);
        LogTextBox.ScrollToEnd();
    }

    private void ClearLog()
    {
        LogTextBox.Document.Blocks.Clear();
    }

    private void ClearBomRows()
    {
        _rows.Clear();
        _headerFilters.Clear();
        if (_view is not null)
        {
            _view.Filter = null;
            _view.Refresh();
        }

        HeaderFilterButton.Content = _headerFilterEnabled ? "关闭筛选" : "表头筛选";
        ResetReadProgress();
        ProgressText.Text = "就绪";
        ReadProgressBar.IsIndeterminate = false;
        ReadProgressBar.Value = 0;
    }

    private void ValidateBomPaths()
    {
        if (_rows.Count == 0)
        {
            return;
        }

        var mainDirectory = GetNormalizedDirectory(_rows[0].FullPath);
        if (string.IsNullOrWhiteSpace(mainDirectory))
        {
            AppendLog("完整路径校验跳过: 无法获取主装配体目录", Brushes.Red);
            return;
        }

        var invalidPathRows = 0;
        var unsetMaterialRows = 0;
        foreach (var row in _rows.Skip(1))
        {
            var rowDirectory = GetNormalizedDirectory(row.FullPath);
            row.IsOutsideMainAssemblyDirectory = !string.Equals(rowDirectory, mainDirectory, StringComparison.OrdinalIgnoreCase);
            if (row.IsOutsideMainAssemblyDirectory)
            {
                invalidPathRows++;
            }

            row.HasUnsetMaterial = IsUnsetMaterial(row);
            if (row.HasUnsetMaterial)
            {
                unsetMaterialRows++;
            }
        }

        if (invalidPathRows > 0)
        {
            AppendLog($"完整路径校验: {invalidPathRows} 项不在主装配体目录，已标红", Brushes.Red);
        }
        else
        {
            AppendLog("完整路径校验: 全部项目都在主装配体目录");
        }

        if (unsetMaterialRows > 0)
        {
            AppendLog($"材料校验: {unsetMaterialRows} 项 SW材料未设置，已标红", Brushes.Red);
        }
        else
        {
            AppendLog("材料校验: 全部零件 SW材料已设置");
        }

        _view?.Refresh();
    }

    private static bool IsUnsetMaterial(BomRow row)
    {
        if (!row.DocumentType.Equals("零件", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return string.IsNullOrWhiteSpace(row.Material)
               || row.Material.Equals("未设置", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetNormalizedDirectory(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            var fullPath = Path.GetFullPath(path);
            return Path.GetDirectoryName(fullPath)?.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void AppendLogWithHighlight(string text, string highlight, Brush highlightBrush)
    {
        var paragraph = new WpfParagraph { Margin = new Thickness(0) };
        paragraph.Inlines.Add(new WpfRun($"{DateTime.Now:HH:mm:ss} "));

        var index = string.IsNullOrWhiteSpace(highlight)
            ? -1
            : text.LastIndexOf(highlight, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
        {
            paragraph.Inlines.Add(new WpfRun(text));
        }
        else
        {
            paragraph.Inlines.Add(new WpfRun(text[..index]));
            paragraph.Inlines.Add(new WpfRun(text.Substring(index, highlight.Length))
            {
                Foreground = highlightBrush,
                FontWeight = FontWeights.Bold
            });
            paragraph.Inlines.Add(new WpfRun(text[(index + highlight.Length)..]));
        }

        LogTextBox.Document.Blocks.Add(paragraph);
        LogTextBox.ScrollToEnd();
    }

    private void AppendDocumentTitleLog(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            AppendLog("当前文档标题: (无)");
            return;
        }

        if (!IsUnsavedDocumentTitle(title))
        {
            AppendLog($"当前文档标题: {title}");
            return;
        }

        var cleanTitle = title.TrimEnd().TrimEnd('*').TrimEnd();
        var paragraph = new WpfParagraph { Margin = new Thickness(0) };
        paragraph.Inlines.Add(new WpfRun($"{DateTime.Now:HH:mm:ss} 当前文档标题: {cleanTitle} "));
        paragraph.Inlines.Add(new WpfRun("未保存")
        {
            Foreground = Brushes.Red,
            FontWeight = FontWeights.Bold
        });
        LogTextBox.Document.Blocks.Add(paragraph);
        LogTextBox.ScrollToEnd();
    }

    private static bool IsUnsavedDocumentTitle(string title)
    {
        return title.TrimEnd().EndsWith("*", StringComparison.Ordinal);
    }

    private static string GetSuppressedPartHighlight(string message)
    {
        const string prefix = "压缩零件: ";
        if (!message.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return string.Empty;
        }

        var name = message[prefix.Length..].Trim();
        var slash = name.LastIndexOf('/');
        if (slash >= 0 && slash < name.Length - 1)
        {
            name = name[(slash + 1)..];
        }

        var dash = name.LastIndexOf('-');
        if (dash > 0 && dash < name.Length - 1 && int.TryParse(name[(dash + 1)..], out _))
        {
            name = name[..dash];
        }

        return name;
    }

    private static bool ContainsIgnoreCase(string source, string filter)
    {
        if (string.IsNullOrWhiteSpace(filter)) return true;
        return source.Contains(filter.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private List<BomRow> GetVisibleRows()
    {
        if (_view is null)
        {
            return _rows.ToList();
        }

        return _view.Cast<object>().OfType<BomRow>().ToList();
    }

    private static void ExportToExcel(List<BomRow> rows, IReadOnlyList<string> propertyNames, string filePath)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
                        ?? throw new InvalidOperationException("未检测到 Excel，请安装 Microsoft Excel 后再导出。");

        dynamic? app = null;
        dynamic? books = null;
        dynamic? book = null;
        dynamic? sheet = null;
        try
        {
            app = Activator.CreateInstance(excelType) ?? throw new InvalidOperationException("无法启动 Excel。");
            app.DisplayAlerts = false;
            books = app.Workbooks;
            book = books.Add();
            sheet = book.Worksheets(1);

            var headers = BuildExportHeaders(propertyNames);
            for (var c = 0; c < headers.Length; c++)
            {
                sheet.Cells(1, c + 1).Value2 = headers[c];
            }

            for (var i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                var row = i + 2;
                var values = BuildExportValues(r, propertyNames);
                for (var c = 0; c < values.Length; c++)
                {
                    sheet.Cells(row, c + 1).Value2 = values[c];
                }
            }

            sheet.Columns.AutoFit();
            book.SaveAs(filePath);
        }
        finally
        {
            try { if (book != null) book.Close(false); } catch { }
            try { if (app != null) app.Quit(); } catch { }
            ReleaseCom(sheet);
            ReleaseCom(book);
            ReleaseCom(books);
            ReleaseCom(app);
        }
    }

    private static string[] BuildExportHeaders(IReadOnlyList<string> propertyNames)
    {
        return ["类型", "序号", "文件名", "配置", "数量", "SW材料", .. propertyNames.Select(GetPropertyDisplayName), "完整路径"];
    }

    private static string[] BuildExportValues(BomRow row, IReadOnlyList<string> propertyNames)
    {
        var values = new List<string>
        {
            row.DocumentType,
            row.Index.ToString(CultureInfo.InvariantCulture),
            row.FileName,
            row.Configuration,
            row.Quantity.ToString(CultureInfo.InvariantCulture),
            row.Material
        };
        values.AddRange(propertyNames.Select(name => row.GetProperty(name)));
        values.Add(row.FullPath);
        return values.ToArray();
    }

    private static void ReleaseCom(object? com)
    {
        if (com != null && Marshal.IsComObject(com))
        {
            Marshal.ReleaseComObject(com);
        }
    }

    private ReadMode GetReadMode()
    {
        return ReadModeComboBox.SelectedIndex switch
        {
            1 => ReadMode.AllComponents,
            _ => ReadMode.BomOnly
        };
    }

    private PropertySourceMode GetPropertySourceMode()
    {
        return PropertySourceComboBox.SelectedIndex == 1
            ? PropertySourceMode.CurrentConfiguration
            : PropertySourceMode.Custom;
    }

    private void ConfigurePropertyColumns()
    {
        var insertIndex = BomGrid.Columns
            .Select((column, index) => new { column, index })
            .FirstOrDefault(x => string.Equals(x.column.Header?.ToString(), "完整路径", StringComparison.OrdinalIgnoreCase))
            ?.index ?? BomGrid.Columns.Count;

        var oldDynamicColumns = BomGrid.Columns
            .Where(IsDynamicPropertyColumn)
            .ToList();
        foreach (var column in oldDynamicColumns)
        {
            BomGrid.Columns.Remove(column);
            _dynamicPropertyColumns.Remove(column);
        }

        foreach (var propertyName in _propertyMapping.PropertyNames)
        {
            var column = new DataGridTemplateColumn
            {
                Header = GetPropertyDisplayName(propertyName),
                CellTemplate = CreateEditablePropertyTemplate(propertyName),
                CellStyle = CreateChangedPropertyCellStyle(propertyName),
                Width = DataGridLength.Auto,
                MinWidth = GetPropertyColumnMinWidth(propertyName),
                IsReadOnly = false
            };
            _dynamicPropertyColumns[column] = propertyName;
            BomGrid.Columns.Insert(insertIndex++, column);
        }
    }

    private bool IsDynamicPropertyColumn(DataGridColumn column)
    {
        if (!string.IsNullOrWhiteSpace(GetDynamicColumnPropertyName(column)))
        {
            return true;
        }

        return column is DataGridBoundColumn boundColumn
               && boundColumn.Binding is Binding binding
               && IsDynamicPropertyBinding(binding);
    }

    private static bool IsDynamicPropertyBinding(Binding binding)
    {
        if (binding.Converter is BomPropertyValueConverter)
        {
            return true;
        }

        return binding.Path?.Path.StartsWith("Properties[", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static DataTemplate CreateEditablePropertyTemplate(string propertyName)
    {
        var textBox = new FrameworkElementFactory(typeof(TextBox));
        textBox.SetBinding(TextBox.TextProperty, new Binding($"Properties[{propertyName}]")
        {
            Mode = BindingMode.TwoWay,
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        });
        textBox.SetValue(Control.BorderThicknessProperty, new Thickness(0));
        textBox.SetValue(Control.PaddingProperty, new Thickness(2, 0, 2, 0));
        textBox.SetValue(Control.BackgroundProperty, Brushes.Transparent);
        textBox.SetValue(Control.VerticalContentAlignmentProperty, VerticalAlignment.Center);
        textBox.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Stretch);
        textBox.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
        textBox.SetValue(TextBox.AcceptsReturnProperty, false);
        textBox.SetValue(TextBox.AcceptsTabProperty, false);
        textBox.SetValue(FrameworkElement.TagProperty, propertyName);
        textBox.SetValue(UIElement.FocusableProperty, false);
        textBox.SetValue(TextBox.IsReadOnlyProperty, true);
        textBox.AddHandler(Control.MouseDoubleClickEvent, new MouseButtonEventHandler(EditablePropertyTextBox_OnMouseDoubleClick));
        textBox.AddHandler(UIElement.PreviewKeyDownEvent, new KeyEventHandler(EditablePropertyTextBox_OnPreviewKeyDown));
        textBox.AddHandler(UIElement.LostKeyboardFocusEvent, new KeyboardFocusChangedEventHandler(EditablePropertyTextBox_OnLostKeyboardFocus));
        textBox.AddHandler(TextBox.TextChangedEvent, new TextChangedEventHandler(EditablePropertyTextBox_OnTextChanged));

        return new DataTemplate { VisualTree = textBox };
    }

    private static void EditablePropertyTextBox_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            return;
        }

        BeginTextBoxEdit(textBox, selectAll: true);
        e.Handled = true;
    }

    private static void EditablePropertyTextBox_OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not TextBox textBox || !textBox.IsReadOnly)
        {
            return;
        }

        if (e.Key is Key.F2 or Key.Enter)
        {
            BeginTextBoxEdit(textBox, selectAll: false);
            e.Handled = true;
        }
    }

    private static void BeginTextBoxEdit(TextBox textBox, bool selectAll)
    {
        textBox.Focusable = true;
        textBox.IsReadOnly = false;
        textBox.Focus();
        if (selectAll)
        {
            textBox.SelectAll();
        }
        else
        {
            textBox.CaretIndex = textBox.Text.Length;
        }
    }

    private static void EditablePropertyTextBox_OnLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            return;
        }

        textBox.IsReadOnly = true;
        textBox.Focusable = false;
    }

    private static void EditablePropertyTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            return;
        }

        if (textBox.Tag is not string propertyName || textBox.DataContext is not BomRow row)
        {
            return;
        }

        var cell = FindParent<DataGridCell>(textBox);
        if (cell is null)
        {
            return;
        }

        cell.Background = row.IsPropertyChanged(propertyName)
            ? new SolidColorBrush(Color.FromRgb(255, 222, 228))
            : Brushes.Transparent;
    }

    private static Style CreateChangedPropertyCellStyle(string propertyName)
    {
        var style = new Style(typeof(DataGridCell));
        style.Setters.Add(new Setter(VerticalContentAlignmentProperty, VerticalAlignment.Center));
        style.Setters.Add(new Setter(PaddingProperty, new Thickness(4, 0, 4, 0)));

        var trigger = new DataTrigger
        {
            Binding = new Binding(".")
            {
                Converter = new BomPropertyChangedConverter(),
                ConverterParameter = propertyName
            },
            Value = true
        };
        trigger.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Color.FromRgb(255, 222, 228))));
        style.Triggers.Add(trigger);
        return style;
    }

    private static string GetPropertyDisplayName(string propertyName)
    {
        if (propertyName.Equals("材料", StringComparison.OrdinalIgnoreCase))
        {
            return "材料(属性)";
        }

        if (propertyName.Equals("数量", StringComparison.OrdinalIgnoreCase))
        {
            return "数量(属性)";
        }

        return propertyName;
    }

    private void AppendPropertyReadSummary()
    {
        var propertyNames = _propertyMapping.PropertyNames;
        if (propertyNames.Count == 0 || _rows.Count == 0)
        {
            return;
        }

        var matched = 0;
        foreach (var row in _rows)
        {
            matched += propertyNames.Count(name => !string.IsNullOrWhiteSpace(row.GetProperty(name)));
        }

        AppendLog($"TXT属性读取: {matched}/{_rows.Count * propertyNames.Count} 个单元格有值");
        if (matched == 0)
        {
            foreach (var row in _rows.Where(x => x.AvailablePropertyNames.Count > 0).Take(3))
            {
                AppendLog($"已读取属性名: {row.FileName}: {string.Join(", ", row.AvailablePropertyNames.Take(20))}");
            }
        }
    }

    private static double GetPropertyColumnMinWidth(string propertyName)
    {
        return propertyName switch
        {
            "零件名称" => 160,
            "物料编码" or "零件图号" or "材料" => 120,
            "表面处理" or "零件类型" => 110,
            "设计" or "版本" or "备注" => 80,
            _ => 100
        };
    }

    private sealed record ColumnOption(string PropertyName, string DisplayName);

    private sealed record FillColumnResult(ColumnOption Target, string Value);

    private sealed record FindReplaceResult(ColumnOption Target, string FindText, string ReplaceText);
}

public enum ReadMode
{
    BomOnly,
    AllComponents
}

public enum PropertySourceMode
{
    Custom,
    CurrentConfiguration
}

public record ReadOptions(ReadMode ReadMode, PropertySourceMode PropertySourceMode, bool SkipVirtual, bool GroupByConfig);

public sealed record ReadProgress(string Message, int Completed, int Total);

public sealed record SaveResult(int TotalRows, int SavedRows, int FailedRows, int SavedProperties);

public sealed record PropertyChange(string Name, string Value);

public sealed class PropertyMappingConfig
{
    public static string DefaultPath => Path.Combine(AppContext.BaseDirectory, "property-mapping.txt");
    private const string MaterialKey = "材料";
    private static readonly string[] DefaultPropertyNames = ["物料编码", "零件图号", "零件名称", "设计", MaterialKey, "零件类型", "表面处理", "版本", "备注"];

    public string SourcePath { get; private init; } = string.Empty;
    public List<string> PropertyNames { get; init; } = DefaultPropertyNames.ToList();
    public List<string> Material { get; init; } = [MaterialKey];

    public static PropertyMappingConfig Load()
    {
        var path = DefaultPath;
        if (!File.Exists(path))
        {
            var defaults = CreateDefault(path);
            File.WriteAllText(path, defaults.ToText(), new UTF8Encoding(true));
            return defaults;
        }

        var propertyNames = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var errors = new List<string>();
        var lines = File.ReadAllLines(path, Encoding.UTF8);
        for (var i = 0; i < lines.Length; i++)
        {
            ParseMappingLine(lines[i], i + 1, propertyNames, seen, errors);
        }

        if (propertyNames.Count == 0)
        {
            errors.Add("配置文件至少需要保留 1 个属性名");
        }

        if (errors.Count > 0)
        {
            throw new InvalidDataException($"配置文件 {path} 格式错误:{Environment.NewLine}{string.Join(Environment.NewLine, errors)}");
        }

        return Normalize(new PropertyMappingConfig
        {
            SourcePath = path,
            PropertyNames = propertyNames,
            Material = [MaterialKey]
        }, path);
    }

    public static PropertyMappingConfig CreateDefault(string path)
    {
        return new PropertyMappingConfig { SourcePath = path };
    }

    private static void ParseMappingLine(string rawLine, int lineNumber, List<string> propertyNames, HashSet<string> seen, List<string> errors)
    {
        var line = rawLine.Trim();
        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
        {
            return;
        }

        if (line.Contains('='))
        {
            errors.Add($"第 {lineNumber} 行不再支持 '='，请只保留属性名: {rawLine}");
            return;
        }

        if (!seen.Add(line))
        {
            errors.Add($"第 {lineNumber} 行重复属性名: {line}");
            return;
        }

        propertyNames.Add(line);
    }

    private string ToText()
    {
        return string.Join(Environment.NewLine, PropertyNames) + Environment.NewLine;
    }

    private static PropertyMappingConfig Normalize(PropertyMappingConfig config, string path)
    {
        return new PropertyMappingConfig
        {
            SourcePath = path,
            PropertyNames = Ensure(config.PropertyNames, DefaultPropertyNames),
            Material = Ensure(config.Material, MaterialKey)
        };
    }

    private static List<string> Ensure(List<string>? values, params string[] fallback)
    {
        var cleaned = values?
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return cleaned is { Count: > 0 } ? cleaned : fallback.ToList();
    }
}

public sealed class BomRow
{
    public int Index { get; set; }
    public string DocumentType { get; init; } = string.Empty;
    public string DocumentIconPath { get; init; } = string.Empty;
    public string DrawingStatus { get; init; } = string.Empty;
    public string DrawingIconPath { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string Configuration { get; init; } = string.Empty;
    public int Quantity { get; set; }
    public string Material { get; init; } = string.Empty;
    public Dictionary<string, string> Properties { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> OriginalProperties { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> AvailablePropertyNames { get; init; } = [];
    public string FullPath { get; init; } = string.Empty;
    public bool IsOutsideMainAssemblyDirectory { get; set; }
    public bool HasUnsetMaterial { get; set; }
    public bool HasValidationWarning => IsOutsideMainAssemblyDirectory || HasUnsetMaterial;

    public string GetProperty(string name)
    {
        return Properties.TryGetValue(name, out var value) ? value : string.Empty;
    }

    public void SetProperty(string name, string value)
    {
        Properties[name] = value;
    }

    public List<PropertyChange> GetChangedProperties(IReadOnlyList<string> propertyNames)
    {
        var changes = new List<PropertyChange>();
        foreach (var propertyName in propertyNames)
        {
            if (IsPropertyChanged(propertyName))
            {
                changes.Add(new PropertyChange(propertyName, NormalizeCompareValue(GetProperty(propertyName))));
            }
        }

        return changes;
    }

    public bool IsPropertyChanged(string propertyName)
    {
        var current = NormalizeCompareValue(GetProperty(propertyName));
        var original = OriginalProperties.TryGetValue(propertyName, out var value)
            ? NormalizeCompareValue(value)
            : string.Empty;
        return !string.Equals(current, original, StringComparison.Ordinal);
    }

    public void AcceptSavedChanges(IEnumerable<PropertyChange> changes)
    {
        foreach (var change in changes)
        {
            OriginalProperties[change.Name] = change.Value;
        }
    }

    private static string NormalizeCompareValue(string? value)
    {
        return value ?? string.Empty;
    }
}

public sealed class BomPropertyValueConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not BomRow row || parameter is not string propertyName)
        {
            return string.Empty;
        }

        return row.GetProperty(propertyName);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}

public sealed class BomPropertyChangedConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is BomRow row
               && parameter is string propertyName
               && row.IsPropertyChanged(propertyName);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }
}

internal static class SolidWorksReader
{
    public sealed record ConnectionInfo(int ProcessId, string Version, string DisplayVersion, int OpenDocumentCount, string ActiveDocumentTitle, string MainWindowTitle);
    private static int _lastConnectedProcessId;

    private sealed class ProgressTracker
    {
        private readonly Action<ReadProgress>? _progress;
        private readonly string _message;
        private readonly int _total;
        private int _completed;

        public ProgressTracker(Action<ReadProgress>? progress, string message, int total)
        {
            _progress = progress;
            _message = message;
            _total = Math.Max(total, 1);
        }

        public void Report()
        {
            Report(_message);
        }

        public void Report(string message)
        {
            _completed = Math.Min(_completed + 1, _total);
            _progress?.Invoke(new ReadProgress(message, _completed, _total));
        }

        public void ReportMessage(string message)
        {
            _progress?.Invoke(new ReadProgress(message, _completed, _total));
        }
    }

    public static dynamic Connect()
    {
        List<string> rotNames = [];
        var swProcessCount = Process.GetProcessesByName("SLDWORKS").Length;
        var attempts = swProcessCount > 0 ? 10 : 1;
        for (var attempt = 1; attempt <= attempts; attempt++)
        {
            // Prefer ROT enumeration so we can pick the instance that actually has opened documents.
            var fromRot = FindBestRunningSolidWorksFromRot(out rotNames);
            if (fromRot != null)
            {
                _lastConnectedProcessId = fromRot.ProcessId;
                return fromRot.App;
            }

            if (TryGetActiveSolidWorksObject(out var runningObject) && runningObject is not null)
            {
                dynamic app = runningObject;
                if (IsUsableSolidWorksApp(app))
                {
                    return runningObject;
                }
            }

            if (attempt < attempts)
            {
                Thread.Sleep(500);
            }
        }

        var rotSummary = rotNames.Count == 0
            ? "ROT中没有发现SolidWorks注册项"
            : "ROT候选: " + string.Join(" | ", rotNames.Take(5));
        throw new InvalidOperationException(
            swProcessCount > 0
                ? $"检测到 {swProcessCount} 个 SLDWORKS.exe，但没有找到可用的 SolidWorks COM 实例。{rotSummary}。请确认本工具和 SolidWorks 使用相同权限运行。"
                : "未检测到已运行的 SolidWorks 实例。请先启动并打开模型后再连接。");
    }

    private static bool TryGetActiveSolidWorksObject(out object? runningObject)
    {
        runningObject = null;
        foreach (var progId in new[] { "SldWorks.Application", "SolidWorks.Application" })
        {
            try
            {
                if (CLSIDFromProgID(progId, out var clsid) != 0)
                {
                    continue;
                }

                GetActiveObject(ref clsid, IntPtr.Zero, out runningObject);
                if (runningObject != null)
                {
                    return true;
                }
            }
            catch
            {
                runningObject = null;
            }
        }

        try
        {
            var clsid = new Guid("72B5B460-38D4-11D0-BD8B-00A0C911CE86");
            GetActiveObject(ref clsid, IntPtr.Zero, out runningObject);
            return runningObject != null;
        }
        catch
        {
            runningObject = null;
            return false;
        }
    }

    private sealed record RotCandidate(object App, int ProcessId, string DisplayName);

    private static RotCandidate? FindBestRunningSolidWorksFromRot(out List<string> rotNames)
    {
        IRunningObjectTable? rot = null;
        IEnumMoniker? enumMoniker = null;
        var candidates = new List<RotCandidate>();
        rotNames = new List<string>();
        try
        {
            GetRunningObjectTable(0, out rot);
            rot.EnumRunning(out enumMoniker);
            var monikers = new IMoniker[1];
            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                IBindCtx? ctx = null;
                try
                {
                    CreateBindCtx(0, out ctx);
                    monikers[0].GetDisplayName(ctx, null, out var displayName);
                    if (string.IsNullOrWhiteSpace(displayName) || !LooksLikeSolidWorksRotName(displayName))
                    {
                        continue;
                    }

                    rotNames.Add(displayName);
                    rot.GetObject(monikers[0], out var obj);
                    if (obj != null)
                    {
                        candidates.Add(new RotCandidate(obj, ExtractPidFromRotName(displayName), displayName));
                    }
                }
                catch
                {
                    // ignore broken ROT entry
                }
                finally
                {
                    if (ctx != null) Marshal.ReleaseComObject(ctx);
                }
            }
        }
        catch
        {
            // ignore and fall back
        }
        finally
        {
            if (enumMoniker != null) Marshal.ReleaseComObject(enumMoniker);
            if (rot != null) Marshal.ReleaseComObject(rot);
        }

        object? best = null;
        var bestScore = -1;
        foreach (var candidate in candidates)
        {
            try
            {
                dynamic sw = candidate.App;
                var info = CheckConnection(sw);
                var effectivePid = info.ProcessId > 0 ? info.ProcessId : candidate.ProcessId;
                if (!IsUsableConnection(info) && effectivePid <= 0)
                {
                    continue;
                }
                var score = (info.OpenDocumentCount > 0 ? 10 : 0)
                            + (!string.IsNullOrWhiteSpace(info.ActiveDocumentTitle) ? 1 : 0)
                            + (effectivePid > 0 ? 1 : 0);
                if (score > bestScore)
                {
                    bestScore = score;
                    best = candidate;
                }
            }
            catch
            {
                // ignore bad candidate
            }
        }

        return best is RotCandidate rotCandidate
               && (rotCandidate.ProcessId > 0 || IsUsableSolidWorksApp(rotCandidate.App))
            ? rotCandidate
            : null;
    }

    private static int ExtractPidFromRotName(string displayName)
    {
        var marker = "PID_";
        var idx = displayName.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return 0;
        idx += marker.Length;
        var end = idx;
        while (end < displayName.Length && char.IsDigit(displayName[end]))
        {
            end++;
        }

        return int.TryParse(displayName[idx..end], out var pid) ? pid : 0;
    }

    private static bool IsUsableSolidWorksApp(dynamic swApp)
    {
        try
        {
            var info = CheckConnection(swApp);
            return IsUsableConnection(info);
        }
        catch
        {
            return false;
        }
    }

    private static bool LooksLikeSolidWorksRotName(string displayName)
    {
        return displayName.Contains("SolidWorks", StringComparison.OrdinalIgnoreCase)
               || displayName.Contains("SldWorks", StringComparison.OrdinalIgnoreCase)
               || displayName.Contains("solidworks_pid_", StringComparison.OrdinalIgnoreCase)
               || displayName.Contains("SldWorks.Application", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUsableConnection(ConnectionInfo info)
    {
        return info.ProcessId > 0
               || info.OpenDocumentCount > 0
               || !string.IsNullOrWhiteSpace(info.ActiveDocumentTitle)
               || !string.IsNullOrWhiteSpace(info.MainWindowTitle)
               || !string.Equals(info.Version, "Unknown", StringComparison.OrdinalIgnoreCase);
    }

    public static ConnectionInfo CheckConnection(dynamic swApp)
    {
        var processId = GetSwProcessId(swApp);
        if (processId <= 0)
        {
            processId = _lastConnectedProcessId;
        }
        var mainWindowTitle = GetSwMainWindowTitle(swApp);
        var version = GetRevisionNumberSafe(swApp);

        var openCount = 0;
        var activeTitle = string.Empty;
        try
        {
            dynamic? doc = GetActiveDocSafe(swApp) ?? GetFirstOpenDocumentSafe(swApp);
            while (doc != null)
            {
                openCount++;
                if (string.IsNullOrWhiteSpace(activeTitle))
                {
                    try { activeTitle = (string?)doc.GetTitle() ?? string.Empty; } catch { }
                }

                try { doc = doc.GetNext(); } catch { break; }
            }
        }
        catch
        {
            // ignore
        }

        if (string.IsNullOrWhiteSpace(activeTitle))
        {
            activeTitle = ExtractDocumentTitleFromWindowTitle(mainWindowTitle);
        }

        return new ConnectionInfo(processId, version, FormatSolidWorksVersion(version, mainWindowTitle), openCount, activeTitle, mainWindowTitle);
    }

    private static string FormatSolidWorksVersion(string revisionNumber, string mainWindowTitle)
    {
        if (!string.IsNullOrWhiteSpace(mainWindowTitle))
        {
            var dash = mainWindowTitle.IndexOf(" - ", StringComparison.Ordinal);
            var titlePrefix = dash > 0 ? mainWindowTitle[..dash].Trim() : mainWindowTitle.Trim();
            if (titlePrefix.Contains("SOLIDWORKS", StringComparison.OrdinalIgnoreCase))
            {
                return titlePrefix;
            }
        }

        var parts = revisionNumber.Split('.', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 1 && int.TryParse(parts[0], out var major))
        {
            var year = major + 1992;
            var sp = parts.Length >= 2 && int.TryParse(parts[1], out var servicePack)
                ? $" SP{servicePack}.0"
                : string.Empty;
            return $"SOLIDWORKS {year}{sp}";
        }

        return "Unknown";
    }

    private static string GetSwMainWindowTitle(dynamic swApp)
    {
        try
        {
            SldWorks.SldWorks typedApp;
            dynamic frame = TryAsStrongSwApp(swApp, out typedApp) ? typedApp.Frame() : swApp.Frame();
            var hwnd = (int)frame.GetHWnd();
            if (hwnd == 0) return string.Empty;
            var len = GetWindowTextLength(new IntPtr(hwnd));
            if (len <= 0) return string.Empty;
            var sb = new StringBuilder(len + 1);
            GetWindowText(new IntPtr(hwnd), sb, sb.Capacity);
            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int GetSwProcessId(dynamic swApp)
    {
        SldWorks.SldWorks typedApp;
        if (TryAsStrongSwApp(swApp, out typedApp))
        {
            try
            {
                var pid = typedApp.GetProcessID();
                if (pid > 0) return pid;
            }
            catch
            {
                // ignore and try dynamic/frame fallback
            }
        }

        try
        {
            var pid = (int)swApp.GetProcessID();
            if (pid > 0) return pid;
        }
        catch
        {
            // ignore and try frame hwnd
        }

        try
        {
            dynamic frame = swApp.Frame();
            var hwnd = (int)frame.GetHWnd();
            if (hwnd != 0)
            {
                GetWindowThreadProcessId(new IntPtr(hwnd), out var pid);
                if (pid > 0) return pid;
            }
        }
        catch
        {
            // ignore
        }

        return 0;
    }

    private static string GetRevisionNumberSafe(dynamic swApp)
    {
        SldWorks.SldWorks typedApp;
        if (TryAsStrongSwApp(swApp, out typedApp))
        {
            try
            {
                return typedApp.RevisionNumber();
            }
            catch
            {
                // ignore and try dynamic fallback
            }
        }

        try
        {
            return (string?)swApp.RevisionNumber() ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    private static bool TryAsStrongSwApp(dynamic swApp, out SldWorks.SldWorks typedApp)
    {
        try
        {
            typedApp = (SldWorks.SldWorks)swApp;
            return typedApp != null;
        }
        catch
        {
            typedApp = null!;
            return false;
        }
    }

    public static List<BomRow> ReadBom(dynamic swApp, ReadOptions options, PropertyMappingConfig mapping, Action<ReadProgress>? progress = null)
    {
        dynamic? activeDoc = GetActiveDocSafe(swApp) ?? GetFirstOpenDocumentSafe(swApp);
        if (activeDoc is null)
        {
            var title = ExtractDocumentTitleFromWindowTitle(GetSwMainWindowTitle(swApp));
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(title)
                ? "SolidWorks 当前没有打开文档。"
                : $"已连接到 SolidWorks 窗口 [{title}]，但 COM 未返回活动文档对象。请在 SW 中切换一下活动窗口后重试。");
        }

        var path = GetModelPathSafe(activeDoc);
        if (string.IsNullOrWhiteSpace(path))
        {
            path = GetModelTitleSafe(activeDoc);
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            path = ExtractDocumentTitleFromWindowTitle(GetSwMainWindowTitle(swApp));
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("无法获取当前模型路径或标题，请先保存当前模型后再读取。");
        }

        var parentRows = new List<BomRow>();
        var map = new Dictionary<string, BomRow>(StringComparer.OrdinalIgnoreCase);
        var type = GetDocumentTypeSafe(path);
        if (type == 2)
        {
            progress?.Invoke(new ReadProgress("读取当前装配体属性", 0, 1));
            var configName = GetActiveConfigurationNameSafe(activeDoc);
            var assemblyRow = ReadFromModel(activeDoc, configName, path, options.PropertySourceMode, mapping);
            assemblyRow.Quantity = 1;
            parentRows.Add(assemblyRow);
            progress?.Invoke(new ReadProgress("读取当前装配体属性", 1, 1));
            ReadAssembly(activeDoc, map, options, swApp, mapping, progress);
            map.Remove(BuildKey(path, assemblyRow.Configuration, options.GroupByConfig));
        }
        else
        {
            progress?.Invoke(new ReadProgress("读取单个模型属性", 0, 1));
            var row = ReadFromModel(activeDoc, "Default", path, options.PropertySourceMode, mapping);
            progress?.Invoke(new ReadProgress("读取单个模型属性", 1, 1));
            row.Quantity = 1;
            parentRows.Add(row);
        }

        parentRows.AddRange(map.Values
            .OrderBy(x => x.FileName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.Configuration, StringComparer.OrdinalIgnoreCase));
        return parentRows;
    }

    public static SaveResult SaveBomProperties(
        dynamic swApp,
        IReadOnlyList<BomRow> rows,
        IReadOnlyList<string> propertyNames,
        PropertySourceMode sourceMode,
        Action<ReadProgress>? progress = null)
    {
        var savedRows = 0;
        var failedRows = 0;
        var savedProperties = 0;
        var changedRows = rows
            .Select(row => new { Row = row, Changes = row.GetChangedProperties(propertyNames) })
            .Where(x => x.Changes.Count > 0)
            .ToList();
        var total = Math.Max(changedRows.Count, 1);

        if (changedRows.Count == 0)
        {
            progress?.Invoke(new ReadProgress("没有需要保存的修改项", 1, 1));
            return new SaveResult(0, 0, 0, 0);
        }

        for (var i = 0; i < changedRows.Count; i++)
        {
            var row = changedRows[i].Row;
            var changes = changedRows[i].Changes;
            var displayName = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName;
            progress?.Invoke(new ReadProgress($"保存TXT属性到SW: {displayName} ({changes.Count}项)", i, total));

            try
            {
                if (string.IsNullOrWhiteSpace(row.FullPath))
                {
                    throw new InvalidOperationException("缺少完整路径");
                }

                var model = TryOpenModel(swApp, row.FullPath);
                if (model is null)
                {
                    throw new InvalidOperationException($"无法打开模型: {row.FullPath}");
                }

                var configName = sourceMode == PropertySourceMode.CurrentConfiguration ? NormalizeConfigName(row.Configuration) : string.Empty;
                var manager = GetCustomPropertyManagerSafe(model, configName);
                if (manager is null)
                {
                    throw new InvalidOperationException(sourceMode == PropertySourceMode.CurrentConfiguration
                        ? $"无法获取配置属性管理器: {configName}"
                        : "无法获取自定义属性管理器");
                }

                foreach (var change in changes)
                {
                    WriteProperty(manager, change.Name, change.Value);
                    savedProperties++;
                }

                SaveModel(model);
                row.AcceptSavedChanges(changes);
                savedRows++;
            }
            catch (Exception ex)
            {
                failedRows++;
                progress?.Invoke(new ReadProgress($"保存失败: {displayName} - {ex.Message}", i + 1, total));
                continue;
            }

            progress?.Invoke(new ReadProgress($"保存TXT属性到SW: {displayName} ({changes.Count}项)", i + 1, total));
        }

        return new SaveResult(changedRows.Count, savedRows, failedRows, savedProperties);
    }

    private static string NormalizeConfigName(string configName)
    {
        return configName.Equals("Default", StringComparison.OrdinalIgnoreCase) ? string.Empty : configName;
    }

    private static void ReadAssembly(dynamic assemblyDoc, Dictionary<string, BomRow> map, ReadOptions options, dynamic swApp, PropertyMappingConfig mapping, Action<ReadProgress>? progress)
    {
        if (options.ReadMode == ReadMode.AllComponents)
        {
            ReadAllComponentsViaOpenDoc(assemblyDoc, map, options, swApp, mapping, progress);
            return;
        }

        ReadBomFromAssemblyComponents(assemblyDoc, map, options, swApp, mapping, progress);
    }

    private sealed class BomSeed
    {
        public string Path { get; init; } = string.Empty;
        public string Config { get; init; } = string.Empty;
        public int Quantity { get; set; }
        public dynamic? Component { get; init; }
        public dynamic? Model { get; set; }
    }

    private static void ReadBomFromAssemblyComponents(
        dynamic assemblyDoc,
        Dictionary<string, BomRow> map,
        ReadOptions options,
        dynamic swApp,
        PropertyMappingConfig mapping,
        Action<ReadProgress>? progress)
    {
        progress?.Invoke(new ReadProgress("获取BOM组件", 0, 1));
        var rawComponents = GetAssemblyComponentsSafe(assemblyDoc, false);
        if (rawComponents is not Array components || components.Length == 0)
        {
            throw new InvalidOperationException("未能从装配体获取BOM组件列表。");
        }

        var seeds = new Dictionary<string, BomSeed>(StringComparer.OrdinalIgnoreCase);
        var suppressedComponents = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in components)
        {
            try
            {
                dynamic component = item;
                if (IsSuppressed(component))
                {
                    suppressedComponents.Add(GetComponentDisplayNameSafe(component));
                    continue;
                }

                if (options.SkipVirtual && IsVirtual(component)) continue;

                var path = GetPathSafe(component);
                if (string.IsNullOrWhiteSpace(path)) continue;
                var config = GetReferencedConfigSafe(component);
                var key = BuildKey(path, config, options.GroupByConfig);

                BomSeed seed;
                if (seeds.TryGetValue(key, out BomSeed? existingSeed) && existingSeed is not null)
                {
                    seed = existingSeed!;
                }
                else
                {
                    seed = new BomSeed
                    {
                        Path = path,
                        Config = config,
                        Quantity = 0,
                        Component = component,
                        Model = GetModelDocSafe(component)
                    };
                    seeds[key] = seed;
                }

                seed!.Quantity++;
            }
            catch
            {
                // skip one broken component
            }
        }

        var total = seeds.Count;
        progress?.Invoke(new ReadProgress("统计BOM项目", total, Math.Max(total, 1)));
        var done = 0;
        foreach (var seed in seeds.Values)
        {
            var model = seed.Model ?? TryOpenModel(swApp, seed.Path);
            if (model == null)
            {
                progress?.Invoke(new ReadProgress($"无法读取BOM属性: {Path.GetFileName(seed.Path)}", done, Math.Max(total, 1)));
                continue;
            }

            var row = ReadFromModel(model, seed.Config, seed.Path, options.PropertySourceMode, mapping);
            row.Quantity = seed.Quantity;
            map[BuildKey(seed.Path, seed.Config, options.GroupByConfig)] = row;
            done++;
            progress?.Invoke(new ReadProgress($"读取BOM属性: {Path.GetFileName(seed.Path)}", done, Math.Max(total, 1)));
        }

        ReportSuppressedComponents(suppressedComponents, progress);
    }

    private static void Traverse(dynamic component, Dictionary<string, BomRow> map, ReadOptions options, PropertyMappingConfig mapping, ProgressTracker tracker)
    {
        var children = GetChildrenSafe(component);
        if (children == null)
        {
            return;
        }

        foreach (var child in children)
        {
            try
            {
            dynamic c = child;
            if (IsSuppressed(c))
            {
                tracker.ReportMessage($"发现压缩零件: {GetComponentDisplayNameSafe(c)}");
                continue;
            }

            if (options.SkipVirtual && IsVirtual(c))
            {
                continue;
            }

            dynamic? model = GetModelDocSafe(c);
            var path = GetPathSafe(c);
            var configName = GetReferencedConfigSafe(c);

            if (model != null && !string.IsNullOrWhiteSpace(path))
            {
                var key = BuildKey(path, configName, options.GroupByConfig);
                if (!map.TryGetValue(key, out BomRow? row))
                {
                    row = ReadFromModel(model, configName, path, options.PropertySourceMode, mapping);
                    row.Quantity = 0;
                    map[key] = row;
                }

                if (row != null)
                {
                    row.Quantity += 1;
                }
                tracker.Report();
            }

            Traverse(c, map, options, mapping, tracker);
            }
            catch
            {
                // skip broken component and continue
            }
        }
    }

    private static int CountComponents(dynamic component, ReadOptions options)
    {
        var children = GetChildrenSafe(component);
        if (children == null)
        {
            return 0;
        }

        var count = 0;
        foreach (var child in children)
        {
            try
            {
                dynamic c = child;
                if (IsSuppressed(c)) continue;
                if (options.SkipVirtual && IsVirtual(c)) continue;
                if (!string.IsNullOrWhiteSpace(GetPathSafe(c)))
                {
                    count++;
                }

                count += CountComponents(c, options);
            }
            catch
            {
                // ignore broken branch when estimating progress total
            }
        }

        return count;
    }

    private static bool IsVirtual(dynamic component)
    {
        try
        {
            return (bool)component.IsVirtual;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsSuppressed(dynamic component)
    {
        try
        {
            return (bool)component.IsSuppressed();
        }
        catch { }

        try
        {
            return (bool)component.IsSuppressed;
        }
        catch { }

        try
        {
            return Convert.ToInt32(component.GetSuppression(), CultureInfo.InvariantCulture) == 0;
        }
        catch
        {
            return false;
        }
    }

    private static string GetComponentDisplayNameSafe(dynamic component)
    {
        try
        {
            var name = Convert.ToString(component.Name2, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        catch { }

        try
        {
            var name = Convert.ToString(component.Name, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        catch { }

        var path = GetPathSafe(component);
        if (!string.IsNullOrWhiteSpace(path))
        {
            return Path.GetFileName(path);
        }

        return "(未知压缩零件)";
    }

    private static void ReportSuppressedComponents(IEnumerable<string> names, Action<ReadProgress>? progress)
    {
        var list = names
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (list.Count == 0)
        {
            return;
        }

        progress?.Invoke(new ReadProgress($"压缩零件汇总: 共 {list.Count} 项", list.Count, list.Count));
        for (var i = 0; i < list.Count; i++)
        {
            progress?.Invoke(new ReadProgress($"压缩零件: {list[i]}", i + 1, list.Count));
        }
    }

    private static string BuildKey(string path, string config, bool groupByConfig)
    {
        return groupByConfig ? $"{path}|{config}" : path;
    }

    private static void ReadAllComponentsViaOpenDoc(dynamic assemblyDoc, Dictionary<string, BomRow> map, ReadOptions options, dynamic swApp, PropertyMappingConfig mapping, Action<ReadProgress>? progress)
    {
        dynamic config = GetActiveConfigurationSafe(assemblyDoc);
        if (config is null)
        {
            return;
        }
        dynamic root = GetRootComponentSafe(config);
        if (root is null)
        {
            return;
        }
        var children = GetChildrenSafe(root);
        if (children == null)
        {
            return;
        }

        progress?.Invoke(new ReadProgress("统计组件数量", 0, 1));
        var total = CountComponents(root, options);
        progress?.Invoke(new ReadProgress("统计组件数量", total, Math.Max(total, 1)));
        var tracker = new ProgressTracker(progress, "读取全部组件属性", total);

        foreach (var child in children)
        {
            try
            {
                AddWithOpenRecursively((dynamic)child, map, options, swApp, mapping, tracker);
            }
            catch
            {
                // skip one broken branch
            }
        }
    }

    private static void AddWithOpenRecursively(dynamic component, Dictionary<string, BomRow> map, ReadOptions options, dynamic swApp, PropertyMappingConfig mapping, ProgressTracker tracker)
    {
        if (IsSuppressed(component))
        {
            tracker.ReportMessage($"发现压缩零件: {GetComponentDisplayNameSafe(component)}");
            return;
        }

        AddComponentToMap(component, map, options, mapping, swApp, tracker);
        var children = GetChildrenSafe(component);
        if (children == null)
        {
            return;
        }

        foreach (var child in children)
        {
            try
            {
                AddWithOpenRecursively((dynamic)child, map, options, swApp, mapping, tracker);
            }
            catch
            {
                // skip one broken child
            }
        }
    }

    private static void AddComponentToMap(dynamic component, Dictionary<string, BomRow> map, ReadOptions options, PropertyMappingConfig mapping, dynamic? swApp = null, ProgressTracker? tracker = null)
    {
        if (IsSuppressed(component))
        {
            tracker?.ReportMessage($"发现压缩零件: {GetComponentDisplayNameSafe(component)}");
            return;
        }

        if (options.SkipVirtual && IsVirtual(component))
        {
            return;
        }

        var path = GetPathSafe(component);
        var configName = GetReferencedConfigSafe(component);
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        dynamic? model = null;
        try
        {
            model = GetModelDocSafe(component);
        }
        catch
        {
            // ignore
        }

        if (model is null && swApp is not null)
        {
            try
            {
                int err = 0, warn = 0;
                var ext = Path.GetExtension(path);
                int docType = ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase) ? 2 : 1;
                model = swApp.OpenDoc6(path, docType, 1, "", ref err, ref warn);
            }
            catch
            {
                model = null;
            }
        }

        if (model is null)
        {
            return;
        }

        var key = BuildKey(path, configName, options.GroupByConfig);
        if (!map.TryGetValue(key, out BomRow? row))
        {
            row = ReadFromModel(model, configName, path, options.PropertySourceMode, mapping);
            row.Quantity = 0;
            map[key] = row;
        }

        if (row != null)
        {
            row.Quantity += 1;
        }
        tracker?.Report($"读取全部组件属性: {Path.GetFileName(path)}");
    }

    private static BomRow ReadFromModel(dynamic model, string configName, string path, PropertySourceMode sourceMode, PropertyMappingConfig mapping)
    {
        var cfgManager = GetCustomPropertyManagerSafe(model, configName);
        var customManager = GetCustomPropertyManagerSafe(model, "");
        var preferConfig = sourceMode == PropertySourceMode.CurrentConfiguration;
        var fileName = Path.GetFileNameWithoutExtension(path);
        var hasDrawing = HasSiblingDrawing(path);
        var documentType = GetDocumentTypeSafe(path);
        var cfgProps = cfgManager is null ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : ReadAllProperties(cfgManager);
        var customProps = customManager is null ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : ReadAllProperties(customManager);
        var material = documentType == 1
            ? GetPartMaterial(model, configName, cfgProps, customProps, cfgManager, customManager, preferConfig, mapping)
            : "无需设置";
        if (documentType == 1 && string.IsNullOrWhiteSpace(material))
        {
            material = "未设置";
        }
        var properties = ReadConfiguredProperties(
            mapping.PropertyNames,
            preferConfig,
            cfgProps,
            customProps,
            cfgManager,
            customManager,
            material);

        return new BomRow
        {
            DocumentType = GetDocumentTypeLabel(path),
            DocumentIconPath = GetDocumentIconPath(path),
            DrawingStatus = hasDrawing ? "有工程图" : "无工程图",
            DrawingIconPath = hasDrawing ? "pack://application:,,,/Assets/drawing.png" : string.Empty,
            FileName = fileName,
            Configuration = string.IsNullOrWhiteSpace(configName) ? "Default" : configName,
            Material = material,
            Properties = properties,
            OriginalProperties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
            AvailablePropertyNames = BuildAvailablePropertyNames(cfgProps, customProps),
            FullPath = path
        };
    }

    private static List<string> BuildAvailablePropertyNames(
        IReadOnlyDictionary<string, string> cfgProps,
        IReadOnlyDictionary<string, string> customProps)
    {
        return cfgProps.Keys
            .Concat(customProps.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static Dictionary<string, string> ReadConfiguredProperties(
        IReadOnlyList<string> propertyNames,
        bool preferConfig,
        IReadOnlyDictionary<string, string> cfgProps,
        IReadOnlyDictionary<string, string> customProps,
        dynamic? cfgManager,
        dynamic? customManager,
        string material)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var propertyName in propertyNames)
        {
            result[propertyName] = PickMappedValue(preferConfig, cfgProps, customProps, cfgManager, customManager, new[] { propertyName });
        }

        return result;
    }

    private static dynamic? GetCustomPropertyManagerSafe(dynamic model, string configName)
    {
        if (TryAsStrongModelDoc(model, out SldWorks.ModelDoc2 typedModel))
        {
            try { return typedModel.Extension.get_CustomPropertyManager(configName); } catch { }
        }

        try { return model.Extension.CustomPropertyManager(configName); } catch { }
        try { return model.Extension.get_CustomPropertyManager(configName); } catch { }
        return null;
    }

    private static string GetPartMaterial(
        dynamic model,
        string configName,
        IReadOnlyDictionary<string, string> cfgProps,
        IReadOnlyDictionary<string, string> customProps,
        dynamic? cfgManager,
        dynamic? customManager,
        bool preferConfig,
        PropertyMappingConfig mapping)
    {
        var material = ReadSolidWorksMaterial(model, configName);
        if (!string.IsNullOrWhiteSpace(material))
        {
            return material;
        }

        return string.Empty;
    }

    private static string ReadSolidWorksMaterial(dynamic model, string configName)
    {
        if (TryAsStrongPartDoc(model, out SldWorks.IPartDoc partDoc))
        {
            foreach (var config in GetMaterialConfigCandidates(configName))
            {
                try
                {
                    string databaseName;
                    var value = partDoc.GetMaterialPropertyName2(config, out databaseName);
                    if (!string.IsNullOrWhiteSpace(value)) return value;
                }
                catch { }
            }

            try
            {
                string databaseName;
                var value = partDoc.GetMaterialPropertyName(out databaseName);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }

            try
            {
                var value = partDoc.MaterialUserName;
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }

            try
            {
                var value = partDoc.MaterialIdName;
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }
        }

        foreach (var config in GetMaterialConfigCandidates(configName))
        {
            try
            {
                var value = Convert.ToString(model.MaterialIdName2[config], CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }

            try
            {
                var value = Convert.ToString(model.GetMaterialIdName2(config), CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }

            try
            {
                string databaseName;
                var value = Convert.ToString(model.GetMaterialPropertyName2(config, out databaseName), CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
            catch { }
        }

        return string.Empty;
    }

    private static bool TryAsStrongPartDoc(dynamic model, out SldWorks.IPartDoc partDoc)
    {
        try
        {
            partDoc = (SldWorks.IPartDoc)model;
            return partDoc != null;
        }
        catch
        {
            partDoc = null!;
            return false;
        }
    }

    private static IEnumerable<string> GetMaterialConfigCandidates(string configName)
    {
        if (!string.IsNullOrWhiteSpace(configName) && !configName.Equals("Default", StringComparison.OrdinalIgnoreCase))
        {
            yield return configName;
        }

        yield return string.Empty;
        yield return "Default";
    }

    private static bool IsSolidWorksMaterialExpression(string value)
    {
        return value.TrimStart().StartsWith("\"SW-Material@", StringComparison.OrdinalIgnoreCase)
               || value.TrimStart().StartsWith("SW-Material@", StringComparison.OrdinalIgnoreCase);
    }

    private static string PickMappedValue(
        bool preferConfig,
        IReadOnlyDictionary<string, string> cfgProps,
        IReadOnlyDictionary<string, string> customProps,
        dynamic? cfgManager,
        dynamic? customManager,
        IReadOnlyList<string> candidates)
    {
        if (preferConfig)
        {
            var v = GetFirstMatch(cfgProps, candidates);
            if (!string.IsNullOrWhiteSpace(v)) return v;
            v = GetFirstMatch(customProps, candidates);
            if (!string.IsNullOrWhiteSpace(v)) return v;
            v = ReadProperty(cfgManager, candidates);
            return string.IsNullOrWhiteSpace(v) ? ReadProperty(customManager, candidates) : v;
        }

        var value = GetFirstMatch(customProps, candidates);
        if (!string.IsNullOrWhiteSpace(value)) return value;
        value = GetFirstMatch(cfgProps, candidates);
        if (!string.IsNullOrWhiteSpace(value)) return value;
        value = ReadProperty(customManager, candidates);
        return string.IsNullOrWhiteSpace(value) ? ReadProperty(cfgManager, candidates) : value;
    }

    private static string GetFirstMatch(IReadOnlyDictionary<string, string> props, IReadOnlyList<string> candidates)
    {
        foreach (var name in candidates)
        {
            if (props.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private static string ReadProperty(dynamic? manager, IReadOnlyList<string> candidates)
    {
        if (manager is null) return string.Empty;
        object managerObj = manager;
        foreach (var name in candidates)
        {
            var value = ReadPropertyByComApi(managerObj, name);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private static Dictionary<string, string> ReadAllProperties(dynamic manager)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        object managerObj = manager;

        if (TryAsStrongCustomPropertyManager(managerObj, out SldWorks.ICustomPropertyManager typedManager))
        {
            foreach (var name in GetPropertyNamesStrong(typedManager))
            {
                var value = ReadPropertyStrong(typedManager, name);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    result[name] = value;
                }
            }

            if (result.Count > 0)
            {
                return result;
            }

            if (TryReadAllStrongByGetAll3(typedManager, result)) return result;
            if (TryReadAllStrongByGetAll2(typedManager, result)) return result;
        }

        foreach (var name in GetPropertyNames(managerObj))
        {
            var value = ReadPropertyByComApi(managerObj, name);
            if (!string.IsNullOrWhiteSpace(value))
            {
                result[name] = value;
            }
        }
        if (result.Count > 0)
        {
            return result;
        }

        if (TryReadAllByGetAll3(managerObj, result)) return result;
        if (TryReadAllByGetAll2(managerObj, result)) return result;

        return result;
    }

    private static bool TryReadAllStrongByGetAll3(SldWorks.ICustomPropertyManager manager, Dictionary<string, string> result)
    {
        try
        {
            object namesObj = null!;
            object typesObj = null!;
            object valuesObj = null!;
            object resolvedObj = null!;
            object linkedObj = null!;
            manager.GetAll3(ref namesObj, ref typesObj, ref valuesObj, ref resolvedObj, ref linkedObj);
            AddPropertyList(result, namesObj, valuesObj, resolvedObj);
            return result.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryReadAllStrongByGetAll2(SldWorks.ICustomPropertyManager manager, Dictionary<string, string> result)
    {
        try
        {
            object namesObj = null!;
            object typesObj = null!;
            object valuesObj = null!;
            object resolvedObj = null!;
            manager.GetAll2(ref namesObj, ref typesObj, ref valuesObj, ref resolvedObj);
            AddPropertyList(result, namesObj, valuesObj, resolvedObj);
            return result.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    private static void AddPropertyList(Dictionary<string, string> result, object? namesObj, object? valuesObj, object? resolvedObj)
    {
        var names = ToStringList(namesObj);
        var values = ToStringList(valuesObj);
        var resolved = ToStringList(resolvedObj);
        for (var i = 0; i < names.Count; i++)
        {
            var key = names[i];
            if (string.IsNullOrWhiteSpace(key)) continue;
            var value = i < resolved.Count && !string.IsNullOrWhiteSpace(resolved[i]) ? resolved[i] :
                i < values.Count ? values[i] : string.Empty;
            value = NormalizeValue(value);
            if (!string.IsNullOrWhiteSpace(value))
            {
                result[key] = value;
            }
        }
    }

    private static List<string> GetPropertyNamesStrong(SldWorks.ICustomPropertyManager manager)
    {
        try
        {
            object namesObj = manager.GetNames();
            return ToStringList(namesObj)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static bool TryReadAllByGetAll3(object manager, Dictionary<string, string> result)
    {
        var args = new object?[] { null, null, null, null, null };
        if (!TryInvokeCom(manager, "GetAll3", args))
        {
            return false;
        }

        // GetAll3(out names, out types, out values, out resolvedValues, out linkToProperty)
        var names = ToStringList(args[0]);
        var values = ToStringList(args[2]);
        var resolved = ToStringList(args[3]);
        for (var i = 0; i < names.Count; i++)
        {
            var key = names[i];
            if (string.IsNullOrWhiteSpace(key)) continue;
            var value = i < resolved.Count && !string.IsNullOrWhiteSpace(resolved[i]) ? resolved[i] :
                        i < values.Count ? values[i] : string.Empty;
            value = NormalizeValue(value);
            if (!string.IsNullOrWhiteSpace(value))
            {
                result[key] = value;
            }
        }

        return result.Count > 0;
    }

    private static bool TryReadAllByGetAll2(object manager, Dictionary<string, string> result)
    {
        var args = new object?[] { null, null, null, null };
        if (!TryInvokeCom(manager, "GetAll2", args))
        {
            return false;
        }

        // GetAll2(out names, out types, out values, out resolvedValues)
        var names = ToStringList(args[0]);
        var values = ToStringList(args[2]);
        var resolved = ToStringList(args[3]);
        for (var i = 0; i < names.Count; i++)
        {
            var key = names[i];
            if (string.IsNullOrWhiteSpace(key)) continue;
            var value = i < resolved.Count && !string.IsNullOrWhiteSpace(resolved[i]) ? resolved[i] :
                i < values.Count ? values[i] : string.Empty;
            value = NormalizeValue(value);
            if (!string.IsNullOrWhiteSpace(value))
            {
                result[key] = value;
            }
        }

        return result.Count > 0;
    }

    private static List<string> GetPropertyNames(object manager)
    {
        var names = new List<string>();
        if (TryInvokeCom(manager, "GetNames", Array.Empty<object?>(), out var ret))
        {
            names.AddRange(ToStringList(ret));
        }
        else if (TryInvokeCom(manager, "GetNames2", Array.Empty<object?>(), out ret))
        {
            names.AddRange(ToStringList(ret));
        }

        return names
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> ToStringList(object? value)
    {
        var list = new List<string>();
        if (value is null) return list;
        if (value is string s)
        {
            if (!string.IsNullOrWhiteSpace(s)) list.Add(s);
            return list;
        }

        if (value is IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                var str = item?.ToString();
                if (!string.IsNullOrWhiteSpace(str)) list.Add(str);
            }
        }

        return list;
    }

    private static string ReadPropertyByComApi(object manager, string name)
    {
        if (TryAsStrongCustomPropertyManager(manager, out SldWorks.ICustomPropertyManager typedManager))
        {
            var typedValue = ReadPropertyStrong(typedManager, name);
            if (!string.IsNullOrWhiteSpace(typedValue))
            {
                return typedValue;
            }
        }

        if (TryGetViaGet6(manager, name, out var val6)) return val6;
        if (TryGetViaGet5(manager, name, out var val5)) return val5;
        if (TryGetViaGet4(manager, name, out var val4)) return val4;
        if (TryGetViaGet2(manager, name, out var val2)) return val2;
        if (TryGetViaGet(manager, name, out var val)) return val;
        return string.Empty;
    }

    private static string ReadPropertyStrong(SldWorks.ICustomPropertyManager manager, string name)
    {
        try
        {
            string rawValue;
            string resolvedValue;
            bool wasResolved;
            bool linkToProperty;
            manager.Get6(name, false, out rawValue, out resolvedValue, out wasResolved, out linkToProperty);
            var value = PickValue(resolvedValue, rawValue);
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }
        catch { }

        try
        {
            string rawValue;
            string resolvedValue;
            bool wasResolved;
            manager.Get5(name, false, out rawValue, out resolvedValue, out wasResolved);
            var value = PickValue(resolvedValue, rawValue);
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }
        catch { }

        try
        {
            string rawValue;
            string resolvedValue;
            if (manager.Get4(name, false, out rawValue, out resolvedValue))
            {
                var value = PickValue(resolvedValue, rawValue);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
        }
        catch { }

        try
        {
            string rawValue;
            string resolvedValue;
            manager.Get2(name, out rawValue, out resolvedValue);
            var value = PickValue(resolvedValue, rawValue);
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }
        catch { }

        try
        {
            var value = NormalizeValue(manager.Get(name));
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }
        catch { }

        return string.Empty;
    }

    private static void WriteProperty(dynamic manager, string name, string? value)
    {
        var text = value ?? string.Empty;
        object managerObj = manager;
        if (TryAsStrongCustomPropertyManager(managerObj, out SldWorks.ICustomPropertyManager typedManager))
        {
            WritePropertyStrong(typedManager, name, text);
            return;
        }

        if (TryInvokeCom(managerObj, "Add3", new object?[] { name, 30, text, 2 }))
        {
            return;
        }

        if (TryInvokeCom(managerObj, "Set2", new object?[] { name, text }))
        {
            return;
        }

        if (TryInvokeCom(managerObj, "Set", new object?[] { name, text }))
        {
            return;
        }

        if (TryInvokeCom(managerObj, "Add2", new object?[] { name, 30, text }))
        {
            return;
        }

        throw new InvalidOperationException($"属性写入失败: {name}");
    }

    private static void WritePropertyStrong(SldWorks.ICustomPropertyManager manager, string name, string value)
    {
        try
        {
            var ret = manager.Add3(name, 30, value, 2);
            if (ret >= 0)
            {
                return;
            }
        }
        catch { }

        try
        {
            var ret = manager.Set2(name, value);
            if (ret >= 0)
            {
                return;
            }
        }
        catch { }

        try
        {
            var ret = manager.Set(name, value);
            if (ret >= 0)
            {
                return;
            }
        }
        catch { }

        try
        {
            var ret = manager.Add2(name, 30, value);
            if (ret >= 0)
            {
                return;
            }
        }
        catch { }

        throw new InvalidOperationException($"属性写入失败: {name}");
    }

    private static void SaveModel(dynamic model)
    {
        if (TryAsStrongModelDoc(model, out SldWorks.ModelDoc2 typedModel))
        {
            try
            {
                int errors = 0;
                int warnings = 0;
                typedModel.Save3(1, ref errors, ref warnings);
                if (errors != 0)
                {
                    throw new InvalidOperationException($"保存模型失败，错误码: {errors}");
                }

                return;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch { }
        }

        try
        {
            int errors = 0;
            int warnings = 0;
            model.Save3(1, ref errors, ref warnings);
            if (errors != 0)
            {
                throw new InvalidOperationException($"保存模型失败，错误码: {errors}");
            }
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"保存模型失败: {ex.Message}", ex);
        }
    }

    private static bool TryGetViaGet6(object manager, string name, out string value)
    {
        // Get6(name, useCached, out val, out resolved, out wasResolved, out linkToProperty)
        var args = new object?[] { name, false, "", "", false, false };
        if (TryInvokeCom(manager, "Get6", args))
        {
            value = PickValue(args[3], args[2]);
            return !string.IsNullOrWhiteSpace(value);
        }

        value = string.Empty;
        return false;
    }

    private static bool TryGetViaGet5(object manager, string name, out string value)
    {
        // Get5(name, useCached, out val, out resolved, out wasResolved)
        var args = new object?[] { name, false, "", "", false };
        if (TryInvokeCom(manager, "Get5", args))
        {
            value = PickValue(args[3], args[2]);
            return !string.IsNullOrWhiteSpace(value);
        }

        value = string.Empty;
        return false;
    }

    private static bool TryGetViaGet4(object manager, string name, out string value)
    {
        // Get4(name, useCached, out val, out resolved)
        var args = new object?[] { name, false, "", "" };
        if (TryInvokeCom(manager, "Get4", args))
        {
            value = PickValue(args[3], args[2]);
            return !string.IsNullOrWhiteSpace(value);
        }

        value = string.Empty;
        return false;
    }

    private static bool TryGetViaGet2(object manager, string name, out string value)
    {
        // Get2(name, out val, out resolved)
        var args = new object?[] { name, "", "" };
        if (TryInvokeCom(manager, "Get2", args))
        {
            value = PickValue(args[2], args[1]);
            return !string.IsNullOrWhiteSpace(value);
        }

        value = string.Empty;
        return false;
    }

    private static bool TryGetViaGet(object manager, string name, out string value)
    {
        // Get(name)
        var args = new object?[] { name };
        if (TryInvokeCom(manager, "Get", args, out var result) && result is not null)
        {
            value = NormalizeValue(result.ToString());
            return !string.IsNullOrWhiteSpace(value);
        }

        value = string.Empty;
        return false;
    }

    private static bool TryInvokeCom(object comObj, string methodName, object?[] args)
    {
        return TryInvokeCom(comObj, methodName, args, out _);
    }

    private static bool TryInvokeCom(object comObj, string methodName, object?[] args, out object? result)
    {
        try
        {
            result = comObj.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod,
                null,
                comObj,
                args);
            return true;
        }
        catch
        {
            result = null;
            return false;
        }
    }

    private static string PickValue(object? resolved, object? raw)
    {
        var rv = NormalizeValue(resolved?.ToString());
        if (!string.IsNullOrWhiteSpace(rv)) return rv;
        return NormalizeValue(raw?.ToString());
    }

    private static string NormalizeValue(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }

        return input.Trim();
    }

    private static dynamic? GetRootComponentSafe(dynamic config)
    {
        try { return config.GetRootComponent3(true); } catch { }
        try { return config.GetRootComponent2(true); } catch { }
        try { return config.GetRootComponent(); } catch { }
        return null;
    }

    private static dynamic? GetRootComponentFromAssemblySafe(dynamic assemblyDoc)
    {
        try { return assemblyDoc.GetRootComponent3(true); } catch { }
        try { return assemblyDoc.GetRootComponent(); } catch { }
        return null;
    }

    private static dynamic? GetActiveConfigurationSafe(dynamic assemblyDoc)
    {
        if (TryAsStrongModelDoc(assemblyDoc, out SldWorks.ModelDoc2 modelDoc))
        {
            try { return modelDoc.ConfigurationManager.ActiveConfiguration; } catch { }
            try
            {
                var activeName = modelDoc.ConfigurationManager.ActiveConfiguration?.Name;
                if (!string.IsNullOrWhiteSpace(activeName))
                {
                    return modelDoc.GetConfigurationByName(activeName);
                }
            }
            catch { }

            try
            {
                var namesObj = modelDoc.GetConfigurationNames();
                if (namesObj is Array names && names.Length > 0)
                {
                    var firstName = names.GetValue(0)?.ToString();
                    if (!string.IsNullOrWhiteSpace(firstName))
                    {
                        return modelDoc.GetConfigurationByName(firstName);
                    }
                }
            }
            catch { }
        }

        try { return assemblyDoc.ConfigurationManager.ActiveConfiguration; } catch { }
        try { return assemblyDoc.GetActiveConfiguration(); } catch { }
        try
        {
            var namesObj = assemblyDoc.GetConfigurationNames();
            if (namesObj is Array names && names.Length > 0)
            {
                var firstName = names.GetValue(0)?.ToString();
                if (!string.IsNullOrWhiteSpace(firstName))
                {
                    return assemblyDoc.GetConfigurationByName(firstName);
                }
            }
        }
        catch { }
        return null;
    }

    private static string GetActiveConfigurationNameSafe(dynamic model)
    {
        if (TryAsStrongModelDoc(model, out SldWorks.ModelDoc2 modelDoc))
        {
            try
            {
                var name = modelDoc.ConfigurationManager.ActiveConfiguration?.Name;
                if (!string.IsNullOrWhiteSpace(name)) return name;
            }
            catch { }
        }

        try
        {
            var name = model.ConfigurationManager.ActiveConfiguration.Name;
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        catch { }

        try
        {
            var config = GetActiveConfigurationSafe(model);
            string? name = Convert.ToString(config?.Name, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        catch { }

        return "Default";
    }

    private static int GetDocumentTypeSafe(string path)
    {
        var ext = Path.GetExtension(path);
        if (ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return 2;
        if (ext.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return 3;
        return 1;
    }

    private static string GetDocumentTypeLabel(string path)
    {
        return GetDocumentTypeSafe(path) switch
        {
            2 => "装配体",
            3 => "工程图",
            _ => "零件"
        };
    }

    private static string GetDocumentIconPath(string path)
    {
        return GetDocumentTypeSafe(path) switch
        {
            2 => "pack://application:,,,/Assets/assembly.png",
            3 => "pack://application:,,,/Assets/drawing.png",
            _ => "pack://application:,,,/Assets/part.png"
        };
    }

    private static bool HasSiblingDrawing(string modelPath)
    {
        if (string.IsNullOrWhiteSpace(modelPath))
        {
            return false;
        }

        var type = GetDocumentTypeSafe(modelPath);
        if (type == 3)
        {
            return true;
        }

        try
        {
            var directory = Path.GetDirectoryName(modelPath);
            var fileName = Path.GetFileNameWithoutExtension(modelPath);
            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            var drawingPath = Path.Combine(directory, fileName + ".SLDDRW");
            if (File.Exists(drawingPath))
            {
                return true;
            }

            drawingPath = Path.Combine(directory, fileName + ".slddrw");
            return File.Exists(drawingPath);
        }
        catch
        {
            return false;
        }
    }

    private static dynamic? GetActiveDocSafe(dynamic swApp)
    {
        SldWorks.SldWorks typedApp;
        if (TryAsStrongSwApp(swApp, out typedApp))
        {
            try { return typedApp.ActiveDoc; } catch { }
        }

        try { return swApp.ActiveDoc; } catch { }
        return null;
    }

    private static dynamic? GetFirstOpenDocumentSafe(dynamic swApp)
    {
        SldWorks.SldWorks typedApp;
        if (TryAsStrongSwApp(swApp, out typedApp))
        {
            // Avoid IGetFirstDocument2 here: some interop builds expose it in a way that throws DISP_E_BADINDEX.
        }

        try
        {
            dynamic? first = swApp.GetFirstDocument();
            if (first != null) return first;
        }
        catch
        {
            // ignore
        }

        return null;
    }

    private static string ExtractDocumentTitleFromWindowTitle(string windowTitle)
    {
        if (string.IsNullOrWhiteSpace(windowTitle)) return string.Empty;
        var left = windowTitle.LastIndexOf('[');
        var right = windowTitle.LastIndexOf(']');
        if (left >= 0 && right > left)
        {
            return windowTitle.Substring(left + 1, right - left - 1).Trim();
        }

        const string marker = " - ";
        var idx = windowTitle.LastIndexOf(marker, StringComparison.Ordinal);
        return idx >= 0 ? windowTitle[(idx + marker.Length)..].Trim() : string.Empty;
    }

    private static Array? GetChildrenSafe(dynamic component)
    {
        try { return component.GetChildren() as Array; } catch { return null; }
    }

    private static Array? GetAssemblyComponentsSafe(dynamic assemblyDoc, bool topLevelOnly)
    {
        try
        {
            var typedAssembly = (SldWorks.AssemblyDoc)assemblyDoc;
            var components = typedAssembly.GetComponents(topLevelOnly) as Array;
            if (components != null) return components;
        }
        catch
        {
            // fallback to late binding
        }

        try
        {
            return assemblyDoc.GetComponents(topLevelOnly) as Array;
        }
        catch
        {
            return null;
        }
    }

    private static dynamic? TryOpenModel(dynamic swApp, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        path = path.Trim();

        var opened = TryGetOpenModelByPath(swApp, path);
        if (opened is not null)
        {
            return opened;
        }

        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            int err = 0, warn = 0;
            var docType = GetDocumentTypeFromPath(path);
            if (docType == 0)
            {
                return null;
            }

            if (TryAsStrongSwApp(swApp, out SldWorks.SldWorks typedApp))
            {
                return typedApp.OpenDoc6(path, docType, 1, "", ref err, ref warn);
            }

            return swApp.OpenDoc6(path, docType, 1, "", ref err, ref warn);
        }
        catch
        {
            return null;
        }
    }

    private static dynamic? TryGetOpenModelByPath(dynamic swApp, string path)
    {
        try
        {
            var doc = swApp.GetOpenDocumentByName(path);
            if (doc is not null)
            {
                return doc;
            }
        }
        catch { }

        var normalizedPath = NormalizePathForCompare(path);
        try
        {
            dynamic? doc = GetFirstOpenDocumentSafe(swApp);
            while (doc is not null)
            {
                var docPath = NormalizePathForCompare(GetModelPathSafe(doc));
                if (!string.IsNullOrWhiteSpace(docPath) && string.Equals(docPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }

                try { doc = doc.GetNext(); } catch { break; }
            }
        }
        catch { }

        return null;
    }

    private static string NormalizePathForCompare(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(path.Trim());
        }
        catch
        {
            return path.Trim();
        }
    }

    private static int GetDocumentTypeFromPath(string path)
    {
        var ext = Path.GetExtension(path);
        if (ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return 1;
        if (ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return 2;
        if (ext.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return 3;
        return 0;
    }

    private static string GetModelPathSafe(dynamic model)
    {
        SldWorks.ModelDoc2 typedModel;
        if (TryAsStrongModelDoc(model, out typedModel))
        {
            try { return typedModel.GetPathName() ?? string.Empty; } catch { }
        }

        try { return (string?)model.GetPathName() ?? string.Empty; } catch { }
        try { return (string?)model.GetPathName2() ?? string.Empty; } catch { }
        return string.Empty;
    }

    private static string GetModelTitleSafe(dynamic model)
    {
        SldWorks.ModelDoc2 typedModel;
        if (TryAsStrongModelDoc(model, out typedModel))
        {
            try { return typedModel.GetTitle() ?? string.Empty; } catch { }
        }

        try { return (string?)model.GetTitle() ?? string.Empty; } catch { }
        return string.Empty;
    }

    private static bool TryAsStrongModelDoc(dynamic model, out SldWorks.ModelDoc2 typedModel)
    {
        try
        {
            typedModel = (SldWorks.ModelDoc2)model;
            return typedModel != null;
        }
        catch
        {
            typedModel = null!;
            return false;
        }
    }

    private static bool TryAsStrongCustomPropertyManager(object manager, out SldWorks.ICustomPropertyManager typedManager)
    {
        try
        {
            typedManager = (SldWorks.ICustomPropertyManager)manager;
            return typedManager != null;
        }
        catch
        {
            typedManager = null!;
            return false;
        }
    }

    private static dynamic? GetSelectionManagerSafe(dynamic model)
    {
        try { return model.SelectionManager; } catch { }
        try { return model.ISelectionManager; } catch { }
        return null;
    }

    private static dynamic? GetModelDocSafe(dynamic component)
    {
        try { return component.GetModelDoc2(); } catch { }
        try { return component.GetModelDoc(); } catch { }
        return null;
    }

    private static string GetPathSafe(dynamic component)
    {
        try { return (string?)component.GetPathName() ?? string.Empty; } catch { return string.Empty; }
    }

    private static string GetReferencedConfigSafe(dynamic component)
    {
        try { return (string?)component.ReferencedConfiguration ?? string.Empty; } catch { return string.Empty; }
    }

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowTextLength(IntPtr hWnd);
}
