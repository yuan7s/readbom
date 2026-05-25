using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using System.Collections;

namespace readbom;

public partial class MainWindow : Window
{
    private dynamic? _swApp;
    private int _connectedSwPid;
    private readonly ObservableCollection<BomRow> _rows = [];
    private ICollectionView? _view;
    private readonly PropertyMappingConfig _propertyMapping;
    private readonly Dictionary<string, string> _headerFilters = new(StringComparer.OrdinalIgnoreCase);
    private bool _headerFilterEnabled;
    private int _gridSizeLevel;

    public MainWindow()
    {
        InitializeComponent();
        BomGrid.ItemsSource = _rows;
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

    private void ResetReadProgress()
    {
        ProgressText.Text = "正在读取属性...";
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
            FileName = $"bom_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
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
            FileName = $"bom_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
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

    private static string GetColumnPropertyName(DataGridColumn column)
    {
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

    private static string Escape(string? value)
    {
        var v = value ?? string.Empty;
        if (v.Contains('"') || v.Contains(',') || v.Contains('\n') || v.Contains('\r'))
        {
            return "\"" + v.Replace("\"", "\"\"") + "\"";
        }

        return v;
    }

    private void SetStatus(string text) => StatusText.Text = text;

    private void AppendLog(string text, Brush? foreground = null)
    {
        var paragraph = new Paragraph(new Run($"{DateTime.Now:HH:mm:ss} {text}"))
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
        var paragraph = new Paragraph { Margin = new Thickness(0) };
        paragraph.Inlines.Add(new Run($"{DateTime.Now:HH:mm:ss} "));

        var index = string.IsNullOrWhiteSpace(highlight)
            ? -1
            : text.LastIndexOf(highlight, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
        {
            paragraph.Inlines.Add(new Run(text));
        }
        else
        {
            paragraph.Inlines.Add(new Run(text[..index]));
            paragraph.Inlines.Add(new Run(text.Substring(index, highlight.Length))
            {
                Foreground = highlightBrush,
                FontWeight = FontWeights.Bold
            });
            paragraph.Inlines.Add(new Run(text[(index + highlight.Length)..]));
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
        var paragraph = new Paragraph { Margin = new Thickness(0) };
        paragraph.Inlines.Add(new Run($"{DateTime.Now:HH:mm:ss} 当前文档标题: {cleanTitle} "));
        paragraph.Inlines.Add(new Run("未保存")
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
            .Where(column => column is DataGridBoundColumn boundColumn
                             && boundColumn.Binding is Binding binding
                             && binding.Path.Path.StartsWith("Properties[", StringComparison.Ordinal))
            .ToList();
        foreach (var column in oldDynamicColumns)
        {
            BomGrid.Columns.Remove(column);
        }

        foreach (var propertyName in _propertyMapping.PropertyNames)
        {
            BomGrid.Columns.Insert(insertIndex++, new DataGridTextColumn
            {
                Header = GetPropertyDisplayName(propertyName),
                Binding = new Binding($"Properties[{propertyName}]"),
                Width = DataGridLength.Auto,
                MinWidth = GetPropertyColumnMinWidth(propertyName)
            });
        }
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
    public string FullPath { get; init; } = string.Empty;
    public bool IsOutsideMainAssemblyDirectory { get; set; }
    public bool HasUnsetMaterial { get; set; }
    public bool HasValidationWarning => IsOutsideMainAssemblyDirectory || HasUnsetMaterial;

    public string GetProperty(string name)
    {
        return Properties.TryGetValue(name, out var value) ? value : string.Empty;
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
        // Prefer ROT enumeration so we can pick the instance that actually has opened documents.
        var fromRot = FindBestRunningSolidWorksFromRot(out var rotNames);
        if (fromRot != null)
        {
            _lastConnectedProcessId = fromRot.ProcessId;
            return fromRot.App;
        }

        try
        {
            var clsid = new Guid("72B5B460-38D4-11D0-BD8B-00A0C911CE86");
            GetActiveObject(ref clsid, IntPtr.Zero, out var runningObject);
            if (runningObject != null)
            {
                dynamic app = runningObject;
                if (IsUsableSolidWorksApp(app))
                {
                    return app;
                }
            }
        }
        catch
        {
            // ignore and throw below
        }

        var swProcessCount = Process.GetProcessesByName("SLDWORKS").Length;
        var rotSummary = rotNames.Count == 0
            ? "ROT中没有发现SolidWorks注册项"
            : "ROT候选: " + string.Join(" | ", rotNames.Take(5));
        throw new InvalidOperationException(
            swProcessCount > 0
                ? $"检测到 {swProcessCount} 个 SLDWORKS.exe，但没有找到可用的 SolidWorks COM 实例。{rotSummary}。请确认本工具和 SolidWorks 使用相同权限运行。"
                : "未检测到已运行的 SolidWorks 实例。请先启动并打开模型后再连接。");
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
            FullPath = path
        };
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

        // Preferred: one-shot full property list APIs.
        if (TryReadAllByGetAll3(managerObj, result)) return result;
        if (TryReadAllByGetAll2(managerObj, result)) return result;

        // Fallback: enumerate names, then read each property value.
        foreach (var name in GetPropertyNames(managerObj))
        {
            var value = ReadPropertyByComApi(managerObj, name);
            if (!string.IsNullOrWhiteSpace(value))
            {
                result[name] = value;
            }
        }

        return result;
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
        if (TryGetViaGet6(manager, name, out var val6)) return val6;
        if (TryGetViaGet5(manager, name, out var val5)) return val5;
        if (TryGetViaGet4(manager, name, out var val4)) return val4;
        if (TryGetViaGet2(manager, name, out var val2)) return val2;
        if (TryGetViaGet(manager, name, out var val)) return val;
        return string.Empty;
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

        try
        {
            int err = 0, warn = 0;
            var docType = GetDocumentTypeFromPath(path);
            if (docType == 0)
            {
                return null;
            }

            return swApp.OpenDoc6(path, docType, 1, "", ref err, ref warn);
        }
        catch
        {
            return null;
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
