using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WpfParagraph = System.Windows.Documents.Paragraph;
using WpfRun = System.Windows.Documents.Run;

namespace readbom;

public partial class MainWindow
{
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

    private void RefreshBomGridSafely()
    {
        BomGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        BomGrid.CommitEdit(DataGridEditingUnit.Row, true);
        Dispatcher.BeginInvoke(new Action(() =>
        {
            try
            {
                BomGrid.Items.Refresh();
            }
            catch (InvalidOperationException)
            {
                Dispatcher.BeginInvoke(new Action(RefreshBomGridSafely), System.Windows.Threading.DispatcherPriority.Background);
            }
        }), System.Windows.Threading.DispatcherPriority.Background);
    }

    private void ClearBomRows()
    {
        _rows.Clear();
        _headerFilters.Clear();
        _headerValueFilters.Clear();
        if (_view is not null)
        {
            _view.Filter = null;
            _view.Refresh();
        }

        HeaderFilterButton.Content = _headerFilterEnabled ? "关闭筛选" : "表头筛选";
        UpdateHeaderFilterButtons();
        UpdateQuickFilterButtons();
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

    private void ValidateBomMaterials()
    {
        var unsetMaterialRows = 0;
        foreach (var row in _rows.Skip(1))
        {
            row.Material = GetMaterialDisplayForValidation(row.DocumentType, row.Material);
            row.HasUnsetMaterial = IsUnsetMaterial(row);
            if (row.HasUnsetMaterial)
            {
                unsetMaterialRows++;
            }
        }

        AppendLog(unsetMaterialRows > 0
            ? $"材料校验: {unsetMaterialRows} 项 SW材料未设置，已标红"
            : "材料校验: 全部零件 SW材料已设置",
            unsetMaterialRows > 0 ? Brushes.Red : Brushes.Black);
        _view?.Refresh();
    }

    private async Task ResolveBomRelatedFilesAsync()
    {
        if (_rows.Count == 0 || string.IsNullOrWhiteSpace(_rows[0].FullPath))
        {
            return;
        }

        var watch = Stopwatch.StartNew();
        try
        {
            List<string> relatedFiles;
            if (IsOfflineFileMode())
            {
                Action<string> log = message => Dispatcher.Invoke(() => AppendLog(message));
                relatedFiles = await Task.Run(() => OfflineSolidWorksReader.GetRelatedFiles(_rows[0].FullPath, log));
            }
            else
            {
                relatedFiles = await SolidWorksAddinClient.GetRelatedFilesAsync(_rows[0].FullPath, message => AppendLog(message));
            }

            if (relatedFiles.Count == 0)
            {
                AppendLog("相关文件解析: 未获取到主装配体相关文件", Brushes.Red);
                return;
            }

            var modelPathLookup = BuildRelatedModelPathLookup(relatedFiles);
            var resolvedPathCount = 0;
            var drawingCount = 0;
            var typeResolvedCount = 0;
            foreach (var row in _rows.Skip(1))
            {
                var lookupKey = NormalizeFileLookupKey(row.FileName);
                string? resolvedPath = null;
                modelPathLookup.TryGetValue(lookupKey, out resolvedPath);

                if (!string.IsNullOrWhiteSpace(resolvedPath))
                {
                    if (string.IsNullOrWhiteSpace(row.FullPath)
                        || !string.Equals(row.FullPath, resolvedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        row.FullPath = resolvedPath;
                        resolvedPathCount++;
                    }
                }

                var resolvedType = GetDocumentTypeLabelFromPath(row.FullPath);
                if (!string.IsNullOrWhiteSpace(resolvedType)
                    && !string.Equals(row.DocumentType, resolvedType, StringComparison.OrdinalIgnoreCase))
                {
                    row.DocumentType = resolvedType;
                    row.DocumentIconPath = GetDocumentIconPathFromLabelLocal(resolvedType);
                    typeResolvedCount++;
                }

                var hasDrawing = HasDrawingForResolvedPath(row.FullPath, relatedFiles);
                row.DrawingStatus = hasDrawing ? "有工程图" : "无工程图";
                row.DrawingIconPath = hasDrawing ? "pack://application:,,,/Assets/drawing.png" : string.Empty;
                if (hasDrawing)
                {
                    drawingCount++;
                }
            }

            _view?.Refresh();
            AppendLog($"相关文件解析: 获取 {relatedFiles.Count} 个相关文件，回填完整路径 {resolvedPathCount} 项，修正类型 {typeResolvedCount} 项，工程图 {drawingCount} 项，用时 {watch.ElapsedMilliseconds}ms");
        }
        catch (Exception ex)
        {
            AppendLog($"相关文件解析失败: {ex.Message}", Brushes.Red);
        }
    }

    private static Dictionary<string, string> BuildRelatedModelPathLookup(IEnumerable<string> relatedFiles)
    {
        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in relatedFiles)
        {
            var extension = Path.GetExtension(path);
            if (!extension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)
                && !extension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var fileName = NormalizeFileLookupKey(Path.GetFileNameWithoutExtension(path));
            if (!string.IsNullOrWhiteSpace(fileName) && !lookup.ContainsKey(fileName))
            {
                lookup[fileName] = path;
            }
        }

        return lookup;
    }

    private static bool HasDrawingForResolvedPath(string modelPath, IReadOnlyCollection<string> relatedFiles)
    {
        if (string.IsNullOrWhiteSpace(modelPath))
        {
            return false;
        }

        var fileName = NormalizeFileLookupKey(Path.GetFileNameWithoutExtension(modelPath));
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        if (relatedFiles.Any(path =>
                Path.GetExtension(path).Equals(".slddrw", StringComparison.OrdinalIgnoreCase)
                && string.Equals(NormalizeFileLookupKey(Path.GetFileNameWithoutExtension(path)), fileName, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        try
        {
            var directory = Path.GetDirectoryName(modelPath);
            return !string.IsNullOrWhiteSpace(directory)
                   && Directory.Exists(directory)
                   && Directory.EnumerateFiles(directory, "*.slddrw")
                       .Any(path => string.Equals(NormalizeFileLookupKey(Path.GetFileNameWithoutExtension(path)), fileName, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizeFileLookupKey(string? value)
    {
        var key = Path.GetFileNameWithoutExtension(value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(key))
        {
            return string.Empty;
        }

        var atIndex = key.IndexOf('@');
        if (atIndex > 0)
        {
            key = key[..atIndex];
        }

        return key.Trim();
    }

    private static string GetDocumentTypeLabelFromPath(string path)
    {
        var extension = Path.GetExtension(path ?? string.Empty);
        if (extension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return "装配体";
        if (extension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return "零件";
        if (extension.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return "工程图";
        return string.Empty;
    }

    private static string GetDocumentIconPathFromLabelLocal(string label)
    {
        return label switch
        {
            "装配体" => "pack://application:,,,/Assets/assembly.png",
            "零件" => "pack://application:,,,/Assets/part.png",
            "工程图" => "pack://application:,,,/Assets/drawing.png",
            _ => string.Empty
        };
    }

    private static bool IsUnsetMaterial(BomRow row)
    {
        if (!row.DocumentType.Equals("零件", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var material = NormalizeMaterialValueForValidation(row.Material);
        return string.IsNullOrWhiteSpace(material)
               || material.Equals("未设置", StringComparison.OrdinalIgnoreCase)
               || material.Equals("<未指定>", StringComparison.OrdinalIgnoreCase)
               || material.Equals("未指定", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetMaterialDisplayForValidation(string? documentType, string? material)
    {
        var normalized = NormalizeMaterialValueForValidation(material);
        if (!string.IsNullOrWhiteSpace(normalized)
            && !normalized.Equals("<未指定>", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("未指定", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("未设置", StringComparison.OrdinalIgnoreCase))
        {
            return normalized;
        }

        return string.Equals(documentType, "装配体", StringComparison.OrdinalIgnoreCase) ? "无需设置" : "未设置";
    }

    private static string NormalizeMaterialValueForValidation(string? material)
    {
        var value = (material ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        if (value.StartsWith("$PRPSHEET:", StringComparison.OrdinalIgnoreCase))
        {
            value = value[10..].Trim();
            value = RemoveMaterialFormulaPropertyName(value);
        }
        else if (value.StartsWith("$PRP:", StringComparison.OrdinalIgnoreCase))
        {
            value = value[5..].Trim();
            value = RemoveMaterialFormulaPropertyName(value);
        }

        return value.Trim('"', '\'', ' ', '\t');
    }

    private static string RemoveMaterialFormulaPropertyName(string value)
    {
        value = value.Trim('"', '\'', ' ', '\t');
        var separatorIndex = value.IndexOfAny([' ', '\t']);
        if (separatorIndex <= 0)
        {
            return value;
        }

        var propertyName = value[..separatorIndex].Trim('"', '\'', ' ', '\t');
        var resolvedValue = value[(separatorIndex + 1)..].Trim('"', '\'', ' ', '\t');
        return propertyName.Equals("SW-Material", StringComparison.OrdinalIgnoreCase)
               || propertyName.Equals("\u6750\u8d28", StringComparison.OrdinalIgnoreCase)
               || propertyName.Equals("\u6750\u6599", StringComparison.OrdinalIgnoreCase)
            ? resolvedValue
            : value;
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

    private async Task AppendAddinConnectionStatusAsync()
    {
        if (IsOfflineFileMode())
        {
            AppendLog("本地文件模式: 未使用 SW Add-in");
            return;
        }

        if (await SolidWorksAddinClient.IsAvailableAsync())
        {
            AppendLog("插件连接状态: 已连接");
            return;
        }

        AppendLog("插件连接状态: 未连接", Brushes.Red);
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

}
