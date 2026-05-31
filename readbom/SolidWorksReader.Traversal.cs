using System.Globalization;
using System.IO;

namespace readbom;

internal static partial class SolidWorksReader
{
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
        public BomRow? Row { get; set; }
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
                    var model = GetModelDocSafe(component);
                    BomRow? row = null;
                    if (model is not null)
                    {
                        row = ReadFromModel(model, config, path, options.PropertySourceMode, mapping);
                        row.Quantity = 0;
                        map[key] = row;
                    }

                    seed = new BomSeed
                    {
                        Path = path,
                        Config = config,
                        Quantity = 0,
                        Component = component,
                        Model = model,
                        Row = row
                    };
                    seeds[key] = seed;
                }

                seed!.Quantity++;
                if (seed.Row is not null)
                {
                    seed.Row.Quantity++;
                }
            }
            catch
            {
                // skip one broken component
            }
        }

        var total = seeds.Count;
        progress?.Invoke(new ReadProgress("统计BOM项目", total, Math.Max(total, 1)));
        var done = seeds.Values.Count(seed => seed.Row is not null);
        if (done > 0)
        {
            progress?.Invoke(new ReadProgress($"已从BOM组件直接读取属性: {done} 项", done, Math.Max(total, 1)));
        }

        foreach (var seed in seeds.Values)
        {
            if (seed.Row is not null)
            {
                continue;
            }

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

}
