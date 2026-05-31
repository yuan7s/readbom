using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace readbom;

public partial class PropertySettingsWindow : Window
{
    private static readonly string[] BuiltInPropertyNames =
    [
        PropertyMappingConfig.FolderNamePropertyName,
        PropertyMappingConfig.FileNamePropertyName,
        PropertyMappingConfig.QuantityPropertyName
    ];

    private static readonly string[] DefaultPropertyNames =
    [
        "物料编码",
        "零件图号",
        "零件名称",
        "设计",
        "材料",
        "零件类型",
        "表面处理",
        "版本",
        "备注"
    ];

    private readonly ObservableCollection<PropertySettingItem> _items = [];

    public List<string> PropertyNames { get; private set; } = [];

    public PropertySettingsWindow(IReadOnlyList<string> propertyNames)
    {
        InitializeComponent();
        PropertyGrid.ItemsSource = _items;
        LoadItems(propertyNames);
    }

    private void SaveButton_OnClick(object sender, RoutedEventArgs e)
    {
        PropertyGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        PropertyGrid.CommitEdit(DataGridEditingUnit.Row, true);

        var names = ReadUserPropertyNames();
        var errors = Validate(names);
        if (errors.Count > 0)
        {
            ValidationText.Text = string.Join(Environment.NewLine, errors);
            return;
        }

        PropertyNames = names;
        DialogResult = true;
    }

    private void AddButton_OnClick(object sender, RoutedEventArgs e)
    {
        var item = new PropertySettingItem { Name = string.Empty, IsBuiltIn = false };
        _items.Add(item);
        PropertyGrid.SelectedItem = item;
        PropertyGrid.ScrollIntoView(item);
        PropertyGrid.CurrentCell = new DataGridCellInfo(item, PropertyGrid.Columns[0]);
        PropertyGrid.BeginEdit();
    }

    private void DeleteButton_OnClick(object sender, RoutedEventArgs e)
    {
        var selected = PropertyGrid.SelectedItems
            .OfType<PropertySettingItem>()
            .Where(item => !item.IsBuiltIn)
            .ToList();
        foreach (var item in selected)
        {
            _items.Remove(item);
        }
    }

    private void MoveUpButton_OnClick(object sender, RoutedEventArgs e)
    {
        MoveSelectedUserItem(-1);
    }

    private void MoveDownButton_OnClick(object sender, RoutedEventArgs e)
    {
        MoveSelectedUserItem(1);
    }

    private void ResetButton_OnClick(object sender, RoutedEventArgs e)
    {
        LoadItems(DefaultPropertyNames);
        ValidationText.Text = string.Empty;
    }

    private void PropertyGrid_OnCellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.Row.Item is not PropertySettingItem { IsBuiltIn: true })
        {
            return;
        }

        e.Cancel = true;
    }

    private void MoveSelectedUserItem(int offset)
    {
        if (PropertyGrid.SelectedItem is not PropertySettingItem item || item.IsBuiltIn)
        {
            return;
        }

        var currentIndex = _items.IndexOf(item);
        var targetIndex = currentIndex + offset;
        var firstUserIndex = BuiltInPropertyNames.Length;
        if (targetIndex < firstUserIndex || targetIndex >= _items.Count)
        {
            return;
        }

        if (_items[targetIndex].IsBuiltIn)
        {
            return;
        }

        _items.Move(currentIndex, targetIndex);
        PropertyGrid.SelectedItem = item;
        PropertyGrid.ScrollIntoView(item);
    }

    private void LoadItems(IReadOnlyList<string> propertyNames)
    {
        _items.Clear();
        foreach (var name in BuiltInPropertyNames)
        {
            _items.Add(new PropertySettingItem { Name = name, IsBuiltIn = true });
        }

        foreach (var name in propertyNames.Where(name => !string.IsNullOrWhiteSpace(name)))
        {
            _items.Add(new PropertySettingItem { Name = name.Trim(), IsBuiltIn = false });
        }
    }

    private List<string> ReadUserPropertyNames()
    {
        return _items
            .Where(item => !item.IsBuiltIn)
            .Select(item => item.Name.Trim())
            .Where(name => !string.IsNullOrWhiteSpace(name) && !name.StartsWith("#", StringComparison.Ordinal))
            .ToList();
    }

    private static List<string> Validate(IReadOnlyList<string> names)
    {
        var errors = new List<string>();
        if (names.Count == 0)
        {
            errors.Add("至少需要保留 1 个用户属性名。");
            return errors;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < names.Count; i++)
        {
            var name = names[i];
            if (name.Contains('='))
            {
                errors.Add($"第 {i + 1} 个用户属性不支持 '='，请只保留属性名: {name}");
            }

            if (!seen.Add(name))
            {
                errors.Add($"第 {i + 1} 个用户属性重复: {name}");
            }
        }

        return errors;
    }
}

public sealed class PropertySettingItem : INotifyPropertyChanged
{
    private string _name = string.Empty;

    public string Name
    {
        get => _name;
        set
        {
            if (string.Equals(_name, value, StringComparison.Ordinal))
            {
                return;
            }

            _name = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
        }
    }

    public bool IsBuiltIn { get; init; }

    public string Kind => IsBuiltIn ? "内置" : "用户";

    public event PropertyChangedEventHandler? PropertyChanged;
}
