using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace readbom;

public partial class MainWindow
{
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
            var column = new DataGridTextColumn
            {
                Header = GetPropertyDisplayName(propertyName),
                Binding = new Binding($"Properties[{propertyName}]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                CellStyle = CreateChangedPropertyCellStyle(propertyName),
                ElementStyle = CreatePropertyTextElementStyle(),
                EditingElementStyle = CreatePropertyTextEditingStyle(),
                Width = new DataGridLength(GetPropertyColumnMinWidth(propertyName)),
                IsReadOnly = false
            };
            _dynamicPropertyColumns[column] = propertyName;
            BomGrid.Columns.Insert(insertIndex++, column);
        }

        UpdateHeaderFilterButtons();
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

    private static Style CreatePropertyTextElementStyle()
    {
        var style = new Style(typeof(TextBlock));
        style.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center));
        style.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(2, 0, 2, 0)));
        style.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
        return style;
    }

    private static Style CreatePropertyTextEditingStyle()
    {
        var style = new Style(typeof(TextBox));
        style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0)));
        style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(2, 0, 2, 0)));
        style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.White));
        style.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
        return style;
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
