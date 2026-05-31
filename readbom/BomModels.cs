using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace readbom;

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

public sealed record RelocateResult(int TotalRows, int CopiedFiles, int ReloadedRows, int FailedRows);

public sealed record BlankSizeResult(int TotalRows, int UpdatedRows, int FailedRows);

public sealed record PropertyChange(string Name, string Value);

public sealed class AddinSavePropertyRow
{
    public string Path { get; init; } = string.Empty;
    public string Configuration { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public List<AddinSavePropertyChange> Changes { get; init; } = [];
}

public sealed class AddinSavePropertyChange
{
    public string Name { get; init; } = string.Empty;
    public string Value { get; init; } = string.Empty;
}

public sealed class AddinBlankSizeRow
{
    public string Path { get; init; } = string.Empty;
    public string Configuration { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
}

public sealed class PropertyMappingConfig
{
    public static string DefaultPath => Path.Combine(AppContext.BaseDirectory, "property-mapping.txt");
    public const string FolderNamePropertyName = "SW-文件夹名称(Folder Name)";
    public const string FileNamePropertyName = "SW-文件名称(File Name)";
    public const string QuantityPropertyName = "数量";
    public const string PartTypePropertyName = "零件类型";
    private const string MaterialKey = "材料";
    private static readonly string[] DefaultPropertyNames = ["物料编码", "零件图号", "零件名称", "设计", MaterialKey, PartTypePropertyName, "表面处理", "版本", "备注"];

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

    public static PropertyMappingConfig Save(IReadOnlyList<string> propertyNames)
    {
        var config = Normalize(new PropertyMappingConfig
        {
            SourcePath = DefaultPath,
            PropertyNames = ValidatePropertyNames(propertyNames, DefaultPath),
            Material = [MaterialKey]
        }, DefaultPath);
        File.WriteAllText(config.SourcePath, config.ToText(), new UTF8Encoding(true));
        return config;
    }

    private static List<string> ValidatePropertyNames(IReadOnlyList<string> propertyNames, string source)
    {
        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var errors = new List<string>();
        for (var i = 0; i < propertyNames.Count; i++)
        {
            ParseMappingLine(propertyNames[i], i + 1, result, seen, errors);
        }

        if (result.Count == 0)
        {
            errors.Add("至少需要保留 1 个属性名");
        }

        if (errors.Count > 0)
        {
            throw new InvalidDataException($"属性设置 {source} 格式错误:{Environment.NewLine}{string.Join(Environment.NewLine, errors)}");
        }

        return result;
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
    public string DocumentType { get; set; } = string.Empty;
    public string DocumentIconPath { get; set; } = string.Empty;
    public string DrawingStatus { get; set; } = string.Empty;
    public string DrawingIconPath { get; set; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string Configuration { get; init; } = string.Empty;
    public int Quantity { get; set; }
    public string Material { get; set; } = string.Empty;
    public Dictionary<string, string> Properties { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> OriginalProperties { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> AvailablePropertyNames { get; init; } = [];
    public string FullPath { get; set; } = string.Empty;
    public string DirectoryPath
    {
        get
        {
            if (string.IsNullOrWhiteSpace(FullPath))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetDirectoryName(FullPath) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
    public string FolderName => GetProperty(PropertyMappingConfig.FolderNamePropertyName);
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

public sealed class ThumbnailPreviewWindow : Window
{
    private readonly Image _image;
    private readonly TextBlock _title;
    private readonly TextBlock _path;
    private readonly TextBlock _message;
    private readonly Button _pinButton;

    public ThumbnailPreviewWindow()
    {
        Title = "零件缩略图";
        Width = 360;
        Height = 420;
        MinWidth = 280;
        MinHeight = 320;
        WindowStartupLocation = WindowStartupLocation.Manual;
        ShowInTaskbar = true;
        ShowActivated = true;
        Topmost = true;

        _title = new TextBlock
        {
            FontWeight = FontWeights.SemiBold,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(0, 0, 0, 6)
        };
        _path = new TextBlock
        {
            Foreground = Brushes.DimGray,
            TextWrapping = TextWrapping.Wrap,
            FontSize = 12,
            Margin = new Thickness(0, 0, 0, 10)
        };
        _image = new Image
        {
            Stretch = Stretch.Uniform,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch
        };
        _message = new TextBlock
        {
            Text = "未选择项目",
            Foreground = Brushes.DimGray,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            TextWrapping = TextWrapping.Wrap
        };
        _pinButton = new Button
        {
            Content = "📌",
            Width = 28,
            Height = 24,
            Padding = new Thickness(0),
            Margin = new Thickness(8, 0, 0, 6),
            ToolTip = "置顶",
            VerticalAlignment = VerticalAlignment.Top
        };
        _pinButton.Click += (_, _) => TogglePin();
        ApplyPinVisualState();

        var header = new DockPanel();
        DockPanel.SetDock(_pinButton, Dock.Right);
        header.Children.Add(_pinButton);
        header.Children.Add(_title);

        var previewBorder = new Border
        {
            BorderBrush = Brushes.LightGray,
            BorderThickness = new Thickness(1),
            Background = Brushes.White,
            Padding = new Thickness(8),
            Child = new Grid
            {
                Children =
                {
                    _image,
                    _message
                }
            }
        };
        Grid.SetRow(previewBorder, 2);

        Content = new Grid
        {
            Margin = new Thickness(12),
            RowDefinitions =
            {
                new RowDefinition { Height = GridLength.Auto },
                new RowDefinition { Height = GridLength.Auto },
                new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }
            },
            Children =
            {
                header,
                _path,
                previewBorder
            }
        };
        Grid.SetRow(_path, 1);
    }

    public void UpdateRow(BomRow? row)
    {
        if (row is null)
        {
            _title.Text = "未选择项目";
            _path.Text = string.Empty;
            _image.Source = null;
            _message.Text = "未选择项目";
            _message.Visibility = Visibility.Visible;
            return;
        }

        _title.Text = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName;
        _path.Text = row.FullPath;
        var image = ShellThumbnailReader.TryGetThumbnail(row.FullPath, 320);
        _image.Source = image;
        _message.Text = image is null ? "没有可用缩略图" : string.Empty;
        _message.Visibility = image is null ? Visibility.Visible : Visibility.Collapsed;
    }

    private void TogglePin()
    {
        Topmost = !Topmost;
        ApplyPinVisualState();
    }

    private void ApplyPinVisualState()
    {
        _pinButton.Background = Topmost ? Brushes.SteelBlue : SystemColors.ControlBrush;
        _pinButton.Foreground = Topmost ? Brushes.White : SystemColors.ControlTextBrush;
        _pinButton.ToolTip = Topmost ? "取消置顶" : "置顶";
    }
}

internal static class ShellThumbnailReader
{
    [Flags]
    private enum ThumbnailFlags
    {
        ResizeToFit = 0,
        BiggerSizeOk = 1,
        IconOnly = 4,
        ThumbnailOnly = 8
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeSize
    {
        public int Cx;
        public int Cy;
    }

    [ComImport]
    [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemImageFactory
    {
        void GetImage(NativeSize size, ThumbnailFlags flags, out IntPtr phbm);
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void SHCreateItemFromParsingName(
        string pszPath,
        IntPtr pbc,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IShellItemImageFactory ppv);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(
        string pszPath,
        uint dwFileAttributes,
        ref ShellFileInfo psfi,
        uint cbFileInfo,
        uint uFlags);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ShellFileInfo
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    private const uint ShgfiIcon = 0x000000100;
    private const uint ShgfiLargeIcon = 0x000000000;

    public static BitmapSource? TryGetThumbnail(string path, int size)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return null;
        }

        try
        {
            var iid = typeof(IShellItemImageFactory).GUID;
            SHCreateItemFromParsingName(path, IntPtr.Zero, iid, out var factory);
            var nativeSize = new NativeSize { Cx = size, Cy = size };
            factory.GetImage(nativeSize, ThumbnailFlags.ThumbnailOnly | ThumbnailFlags.BiggerSizeOk, out var hBitmap);
            try
            {
                var source = Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
                source.Freeze();
                return source;
            }
            finally
            {
                DeleteObject(hBitmap);
            }
        }
        catch
        {
            return TryGetFileIcon(path);
        }
    }

    private static BitmapSource? TryGetFileIcon(string path)
    {
        try
        {
            var info = new ShellFileInfo();
            var result = SHGetFileInfo(path, 0, ref info, (uint)Marshal.SizeOf<ShellFileInfo>(), ShgfiIcon | ShgfiLargeIcon);
            if (result == IntPtr.Zero || info.hIcon == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                var source = Imaging.CreateBitmapSourceFromHIcon(
                    info.hIcon,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromWidthAndHeight(96, 96));
                source.Freeze();
                return source;
            }
            finally
            {
                DestroyIcon(info.hIcon);
            }
        }
        catch
        {
            return null;
        }
    }
}
