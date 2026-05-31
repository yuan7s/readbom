using System.Collections;
using System.IO;
using SolidWorks.Interop.swdocumentmgr;

namespace readbom;

internal static partial class OfflineSolidWorksReader
{
    public static List<string> GetRelatedFiles(string documentPath, Action<string>? log = null)
    {
        var path = NormalizeExistingSolidWorksPath(documentPath);
        var application = CreateApplication();
        var related = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        AddRelatedFile(related, path);

        using var document = OpenDocument(application, path, readOnly: true);
        if (GetDocumentType(path) == SwDmDocumentType.swDmDocumentAssembly)
        {
            var seeds = new Dictionary<string, OfflineBomSeed>(StringComparer.OrdinalIgnoreCase);
            var suppressed = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            var options = new ReadOptions(ReadMode.AllComponents, PropertySourceMode.CurrentConfiguration, SkipVirtual: false, GroupByConfig: true);
            TraverseAssembly(application, document.Document, GetActiveConfigurationName(document.Document), options, seeds, suppressed);
            foreach (var seed in seeds.Values)
            {
                AddRelatedFile(related, seed.Path);
            }
        }

        foreach (var modelPath in related.ToList())
        {
            AddSiblingDrawing(related, modelPath);
        }

        log?.Invoke($"离线相关文件解析: {related.Count} 个");
        return related.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static void TraverseAssembly(
        SwDMApplication application,
        ISwDMDocument assemblyDocument,
        string configuration,
        ReadOptions options,
        Dictionary<string, OfflineBomSeed> seeds,
        SortedSet<string> suppressed)
    {
        var visitedAssemblies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        TraverseAssemblyCore(application, assemblyDocument, configuration, options, seeds, suppressed, visitedAssemblies);
    }

    private static void TraverseAssemblyCore(
        SwDMApplication application,
        ISwDMDocument assemblyDocument,
        string configuration,
        ReadOptions options,
        Dictionary<string, OfflineBomSeed> seeds,
        SortedSet<string> suppressed,
        HashSet<string> visitedAssemblies)
    {
        var visitKey = BuildKey(assemblyDocument.FullName, configuration, groupByConfig: true);
        if (!visitedAssemblies.Add(visitKey))
        {
            return;
        }

        if (ResolveConfiguration(assemblyDocument, configuration) is not ISwDMConfiguration2 configuration2)
        {
            return;
        }

        foreach (var component in GetComponents(configuration2))
        {
            if (IsSuppressed(component))
            {
                suppressed.Add(GetComponentDisplayName(component));
                continue;
            }

            if (component is not ISwDMComponent4 component4)
            {
                continue;
            }

            if (component4.DocumentType is not (SwDmDocumentType.swDmDocumentAssembly or SwDmDocumentType.swDmDocumentPart))
            {
                continue;
            }

            if (options.ReadMode == ReadMode.BomOnly && IsExcludedFromBom(component4))
            {
                continue;
            }

            if (options.SkipVirtual && IsVirtual(component4))
            {
                continue;
            }

            using var child = TryOpenComponentDocument(application, assemblyDocument, component4);
            var childPath = NormalizeComponentPath(component4, child?.Document);
            if (string.IsNullOrWhiteSpace(childPath))
            {
                continue;
            }

            var childConfiguration = NormalizeComponentConfiguration(component.ConfigurationName, child?.Document);
            AddSeed(seeds, childPath, childConfiguration, options.GroupByConfig);

            if (options.ReadMode == ReadMode.AllComponents
                && child?.Document is not null
                && GetDocumentType(childPath) == SwDmDocumentType.swDmDocumentAssembly)
            {
                TraverseAssemblyCore(application, child.Document, childConfiguration, options, seeds, suppressed, visitedAssemblies);
            }
        }
    }

    private static List<ISwDMComponent> GetComponents(ISwDMConfiguration2 configuration)
    {
        var result = new List<ISwDMComponent>();
        try
        {
            if (configuration.GetComponents() is not IEnumerable components)
            {
                return result;
            }

            foreach (var item in components)
            {
                if (item is ISwDMComponent component)
                {
                    result.Add(component);
                }
            }
        }
        catch
        {
        }

        return result;
    }

    private static SwDmDocumentScope? TryOpenComponentDocument(SwDMApplication application, ISwDMDocument parentDocument, ISwDMComponent4 component)
    {
        try
        {
            var searchOptions = application.GetSearchOptionObject();
            searchOptions.SearchFilters = 27;
            var parentDirectory = Path.GetDirectoryName(parentDocument.FullName);
            if (!string.IsNullOrWhiteSpace(parentDirectory))
            {
                searchOptions.AddSearchPath(parentDirectory);
            }

            var document = component.GetDocument2(true, searchOptions, out SwDmDocumentOpenError error) as ISwDMDocument;
            return document is not null && error == SwDmDocumentOpenError.swDmDocumentOpenErrorNone
                ? new SwDmDocumentScope(document)
                : null;
        }
        catch
        {
            return null;
        }
    }

    private static void AddSeed(Dictionary<string, OfflineBomSeed> seeds, string path, string configuration, bool groupByConfig)
    {
        var key = BuildKey(path, configuration, groupByConfig);
        if (!seeds.TryGetValue(key, out var seed))
        {
            seed = new OfflineBomSeed
            {
                Path = path,
                Configuration = configuration,
                Quantity = 0
            };
            seeds[key] = seed;
        }

        seed.Quantity++;
    }

    private static string NormalizeComponentPath(ISwDMComponent4 component, ISwDMDocument? document)
    {
        var path = document?.FullName;
        if (string.IsNullOrWhiteSpace(path) && component is ISwDMComponent7 component7)
        {
            try
            {
                path = component7.PathName;
            }
            catch
            {
            }
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return path.Trim();
        }
    }

    private static string NormalizeComponentConfiguration(string? configuration, ISwDMDocument? document)
    {
        if (!string.IsNullOrWhiteSpace(configuration))
        {
            return configuration.Trim();
        }

        return document is null ? "Default" : GetActiveConfigurationName(document);
    }

    private static bool IsSuppressed(ISwDMComponent component)
    {
        try
        {
            return component.IsSuppressed();
        }
        catch
        {
            return false;
        }
    }

    private static bool IsVirtual(ISwDMComponent4 component)
    {
        try
        {
            return component.IsVirtual;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsExcludedFromBom(ISwDMComponent4 component)
    {
        try
        {
            return component is ISwDMComponent7 component7 && component7.ExcludeFromBOM != 0;
        }
        catch
        {
            return false;
        }
    }

    private static string GetComponentDisplayName(ISwDMComponent component)
    {
        try
        {
            if (component is ISwDMComponent7 component7 && !string.IsNullOrWhiteSpace(component7.Name2))
            {
                return component7.Name2;
            }
        }
        catch
        {
        }

        try
        {
            if (!string.IsNullOrWhiteSpace(component.Name))
            {
                return component.Name;
            }
        }
        catch
        {
        }

        return "(未知压缩零件)";
    }

    private static string BuildKey(string path, string configuration, bool groupByConfig)
    {
        return groupByConfig ? $"{path}|{configuration}" : path;
    }

    private static void AddRelatedFile(ISet<string> files, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        if (GetDocumentType(path) == SwDmDocumentType.swDmDocumentUnknown)
        {
            return;
        }

        try
        {
            files.Add(Path.GetFullPath(path));
        }
        catch
        {
            files.Add(path.Trim());
        }
    }

    private static void AddSiblingDrawing(ISet<string> files, string modelPath)
    {
        if (GetDocumentType(modelPath) == SwDmDocumentType.swDmDocumentDrawing)
        {
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(modelPath);
            var fileName = Path.GetFileNameWithoutExtension(modelPath);
            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName) || !Directory.Exists(directory))
            {
                return;
            }

            var directPath = Path.Combine(directory, fileName + ".slddrw");
            if (File.Exists(directPath))
            {
                AddRelatedFile(files, directPath);
                return;
            }

            foreach (var drawing in Directory.EnumerateFiles(directory, "*.slddrw"))
            {
                if (string.Equals(Path.GetFileNameWithoutExtension(drawing), fileName, StringComparison.OrdinalIgnoreCase))
                {
                    AddRelatedFile(files, drawing);
                    return;
                }
            }
        }
        catch
        {
        }
    }

    private static void ReportSuppressedComponents(IEnumerable<string> names, Action<ReadProgress>? progress)
    {
        var list = names.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
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
}
