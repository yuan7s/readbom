using System.Globalization;
using System.IO;
using System.Collections;

namespace readbom;

internal static partial class SolidWorksReader
{
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

    public static BlankSizeResult CalculateBlankSizes(
        dynamic swApp,
        IReadOnlyList<BomRow> rows,
        string propertyName,
        Action<ReadProgress>? progress = null)
    {
        var updatedRows = 0;
        var failedRows = 0;
        var total = Math.Max(rows.Count, 1);
        if (rows.Count == 0)
        {
            progress?.Invoke(new ReadProgress("没有需要计算下料尺寸的项目", 1, 1));
            return new BlankSizeResult(0, 0, 0);
        }

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var displayName = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName;
            progress?.Invoke(new ReadProgress($"计算下料尺寸: {displayName}", i, total));

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

                ActivateConfigurationSafe(model, row.Configuration);
                var blankSize = CalculateBlankSize(model, row.FullPath);
                if (string.IsNullOrWhiteSpace(blankSize))
                {
                    throw new InvalidOperationException("无法获取包围盒");
                }

                row.SetProperty(propertyName, blankSize);
                updatedRows++;
            }
            catch (Exception ex)
            {
                failedRows++;
                progress?.Invoke(new ReadProgress($"下料尺寸计算失败: {displayName} - {ex.Message}", i + 1, total));
                continue;
            }

            progress?.Invoke(new ReadProgress($"计算下料尺寸: {displayName}", i + 1, total));
        }

        return new BlankSizeResult(rows.Count, updatedRows, failedRows);
    }

    public static RelocateResult CopyExternalFilesToMainAssemblyDirectory(
        dynamic swApp,
        BomRow mainAssemblyRow,
        IReadOnlyList<BomRow> rows,
        Action<ReadProgress>? progress = null)
    {
        if (rows.Count == 0)
        {
            progress?.Invoke(new ReadProgress("没有需要复制重装的文件", 1, 1));
            return new RelocateResult(0, 0, 0, 0);
        }

        var mainDirectory = Path.GetDirectoryName(Path.GetFullPath(mainAssemblyRow.FullPath));
        if (string.IsNullOrWhiteSpace(mainDirectory))
        {
            throw new InvalidOperationException("无法获取主装配体目录");
        }

        dynamic? activeDoc = GetActiveDocSafe(swApp) ?? GetFirstOpenDocumentSafe(swApp);
        if (activeDoc is null)
        {
            throw new InvalidOperationException("SolidWorks 当前没有打开装配体");
        }

        var assemblyPath = NormalizePathForCompare(GetModelPathSafe(activeDoc));
        if (!string.Equals(assemblyPath, NormalizePathForCompare(mainAssemblyRow.FullPath), StringComparison.OrdinalIgnoreCase))
        {
            activeDoc = TryOpenModel(swApp, mainAssemblyRow.FullPath) ?? activeDoc;
        }

        object assemblyDoc = TryAsStrongAssemblyDoc(activeDoc, out SldWorks.IAssemblyDoc typedAssembly)
            ? typedAssembly
            : activeDoc;
        Array? components = GetAssemblyComponentsSafe(assemblyDoc, false);
        if (components is null || components.Length == 0)
        {
            throw new InvalidOperationException("无法从当前装配体获取组件列表");
        }

        Dictionary<string, List<object>> componentMap = BuildComponentPathMap(components);
        var copiedFiles = 0;
        var reloadedRows = 0;
        var failedRows = 0;
        var total = Math.Max(rows.Count, 1);

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var displayName = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName;
            progress?.Invoke(new ReadProgress($"复制重装: {displayName}", i, total));

            try
            {
                var sourcePath = Path.GetFullPath(row.FullPath);
                if (!File.Exists(sourcePath))
                {
                    throw new FileNotFoundException("源文件不存在", sourcePath);
                }

                var targetPath = ResolveValidatedCopyTarget(sourcePath, mainDirectory);
                if (!File.Exists(targetPath) || !FilesAreSame(sourcePath, targetPath))
                {
                    File.Copy(sourcePath, targetPath, overwrite: false);
                    copiedFiles++;
                }

                var normalizedSource = NormalizePathForCompare(sourcePath);
                if (!componentMap.TryGetValue(normalizedSource, out List<object>? matchingComponents) || matchingComponents.Count == 0)
                {
                    throw new InvalidOperationException($"装配体中没有找到引用: {sourcePath}");
                }

                var replacedCount = 0;
                foreach (var component in matchingComponents)
                {
                    if (ReplaceComponentReference(component, targetPath))
                    {
                        replacedCount++;
                    }
                }

                if (replacedCount == 0)
                {
                    throw new InvalidOperationException("组件引用替换失败");
                }

                row.FullPath = targetPath;
                row.IsOutsideMainAssemblyDirectory = false;
                reloadedRows++;
                progress?.Invoke(new ReadProgress($"复制重装: {displayName}", i + 1, total));
            }
            catch (Exception ex)
            {
                failedRows++;
                progress?.Invoke(new ReadProgress($"复制重装失败: {displayName} - {ex.Message}", i + 1, total));
            }
        }

        SaveModel(activeDoc);
        return new RelocateResult(rows.Count, copiedFiles, reloadedRows, failedRows);
    }

    private static Dictionary<string, List<object>> BuildComponentPathMap(Array components)
    {
        var map = new Dictionary<string, List<object>>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in components)
        {
            if (item is null)
            {
                continue;
            }

            dynamic component = item;
            var path = NormalizePathForCompare(GetPathSafe(component));
            if (string.IsNullOrWhiteSpace(path))
            {
                continue;
            }

            if (!map.TryGetValue(path, out List<object>? list))
            {
                list = [];
                map[path] = list;
            }

            list!.Add(component);
        }

        return map;
    }

    private static string ResolveValidatedCopyTarget(string sourcePath, string targetDirectory)
    {
        var fileName = Path.GetFileName(sourcePath);
        var targetPath = Path.Combine(targetDirectory, fileName);
        if (!File.Exists(targetPath))
        {
            return targetPath;
        }

        if (FilesAreSame(sourcePath, targetPath))
        {
            return targetPath;
        }

        var name = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        for (var i = 1; i < 10000; i++)
        {
            var candidate = Path.Combine(targetDirectory, $"{name}_{i}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            if (FilesAreSame(sourcePath, candidate))
            {
                return candidate;
            }
        }

        throw new IOException($"无法生成不冲突的复制文件名: {fileName}");
    }

    private static bool FilesAreSame(string leftPath, string rightPath)
    {
        try
        {
            var left = new FileInfo(leftPath);
            var right = new FileInfo(rightPath);
            if (!left.Exists || !right.Exists || left.Length != right.Length)
            {
                return false;
            }

            using var leftStream = File.OpenRead(left.FullName);
            using var rightStream = File.OpenRead(right.FullName);
            var leftBuffer = new byte[1024 * 1024];
            var rightBuffer = new byte[1024 * 1024];
            while (true)
            {
                var leftRead = leftStream.Read(leftBuffer, 0, leftBuffer.Length);
                var rightRead = rightStream.Read(rightBuffer, 0, rightBuffer.Length);
                if (leftRead != rightRead)
                {
                    return false;
                }

                if (leftRead == 0)
                {
                    return true;
                }

                for (var i = 0; i < leftRead; i++)
                {
                    if (leftBuffer[i] != rightBuffer[i])
                    {
                        return false;
                    }
                }
            }
        }
        catch
        {
            return false;
        }
    }

    private static bool ReplaceComponentReference(dynamic component, string targetPath)
    {
        try
        {
            var result = Convert.ToInt32(component.ReplaceReference(targetPath), CultureInfo.InvariantCulture);
            if (result >= 0)
            {
                return true;
            }
        }
        catch { }

        try
        {
            component.ReplaceReference(targetPath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizeConfigName(string configName)
    {
        return configName.Equals("Default", StringComparison.OrdinalIgnoreCase) ? string.Empty : configName;
    }

    private static void ActivateConfigurationSafe(dynamic model, string configName)
    {
        if (string.IsNullOrWhiteSpace(configName) || configName.Equals("Default", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        try
        {
            model.ShowConfiguration2(configName);
        }
        catch
        {
            // Keep current configuration if the referenced one cannot be activated.
        }
    }

    private static string CalculateBlankSize(dynamic model, string path)
    {
        var documentType = GetDocumentTypeSafe(path);
        object? corners = null;
        if (documentType == 1)
        {
            if (TryAsStrongPartDoc(model, out SldWorks.IPartDoc partDoc))
            {
                try { corners = partDoc.GetPartBox(true); } catch { }
            }

            if (corners is null)
            {
                try { corners = model.GetPartBox(true); } catch { }
            }
        }
        else if (documentType == 2)
        {
            if (TryAsStrongAssemblyDoc(model, out SldWorks.IAssemblyDoc assemblyDoc))
            {
                try { corners = assemblyDoc.GetBox(1); } catch { }
            }

            if (corners is null)
            {
                try { corners = model.GetBox(1); } catch { }
            }
        }

        var values = ToDoubleList(corners);
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
        return string.Join("x", sizes.Reverse().Select(x => x.ToString("0.#", CultureInfo.InvariantCulture)));
    }

    private static List<double> ToDoubleList(object? value)
    {
        if (value is null)
        {
            return [];
        }

        if (value is IEnumerable enumerable)
        {
            return enumerable
                .Cast<object?>()
                .Select(item => Convert.ToDouble(item, CultureInfo.InvariantCulture))
                .ToList();
        }

        return [];
    }

}
