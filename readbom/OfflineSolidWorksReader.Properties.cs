using System.Collections;
using System.IO;
using SolidWorks.Interop.swdocumentmgr;

namespace readbom;

internal static partial class OfflineSolidWorksReader
{
    private static BomRow CreateBomRow(
        ISwDMDocument document,
        string path,
        string configuration,
        int quantity,
        IReadOnlyList<string> propertyNames,
        PropertySourceMode sourceMode)
    {
        var customProperties = ReadDocumentProperties(document);
        var configurationProperties = ReadConfigurationProperties(ResolveConfiguration(document, configuration));
        var preferConfiguration = sourceMode == PropertySourceMode.CurrentConfiguration;
        var documentType = GetDocumentTypeLabel(path);
        var material = documentType == "零件"
            ? ReadMaterial(preferConfiguration, configurationProperties, customProperties)
            : "无需设置";
        if (documentType == "零件" && string.IsNullOrWhiteSpace(material))
        {
            material = "未设置";
        }

        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var propertyName in propertyNames.Where(x => !string.IsNullOrWhiteSpace(x)))
        {
            properties[propertyName] = PickPropertyValue(preferConfiguration, configurationProperties, customProperties, propertyName);
        }

        var hasDrawing = SolidWorksReader.HasSiblingDrawing(path);
        return new BomRow
        {
            DocumentType = documentType,
            DocumentIconPath = GetDocumentIconPath(path),
            DrawingStatus = hasDrawing ? "有工程图" : "无工程图",
            DrawingIconPath = hasDrawing ? "pack://application:,,,/Assets/drawing.png" : string.Empty,
            FileName = Path.GetFileNameWithoutExtension(path),
            Configuration = string.IsNullOrWhiteSpace(configuration) ? "Default" : configuration,
            Quantity = quantity,
            Material = material,
            Properties = properties,
            OriginalProperties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
            AvailablePropertyNames = configurationProperties.Keys
                .Concat(customProperties.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList(),
            FullPath = path
        };
    }

    private static Dictionary<string, string> ReadDocumentProperties(ISwDMDocument document)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in GetPropertyNames(document.GetCustomPropertyNames))
        {
            try
            {
                var type = SwDmCustomInfoType.swDmCustomInfoUnknown;
                var linkedTo = string.Empty;
                var value = document is ISwDMDocument5 document5
                    ? document5.GetCustomPropertyValues(name, out type, out linkedTo)
                    : document.GetCustomProperty(name, out type);
                AddProperty(result, name, value);
            }
            catch
            {
            }
        }

        return result;
    }

    private static Dictionary<string, string> ReadConfigurationProperties(ISwDMConfiguration? configuration)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (configuration is null)
        {
            return result;
        }

        foreach (var name in GetPropertyNames(configuration.GetCustomPropertyNames))
        {
            try
            {
                var type = SwDmCustomInfoType.swDmCustomInfoUnknown;
                var linkedTo = string.Empty;
                var value = configuration is ISwDMConfiguration5 configuration5
                    ? configuration5.GetCustomPropertyValues(name, out type, out linkedTo)
                    : configuration.GetCustomProperty(name, out type);
                AddProperty(result, name, value);
            }
            catch
            {
            }
        }

        return result;
    }

    private static IEnumerable<string> GetPropertyNames(Func<object> getNames)
    {
        try
        {
            return ToStringList(getNames())
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

    private static void AddProperty(IDictionary<string, string> properties, string name, string? value)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        properties[name.Trim()] = NormalizeValue(value);
    }

    private static string PickPropertyValue(
        bool preferConfiguration,
        IReadOnlyDictionary<string, string> configurationProperties,
        IReadOnlyDictionary<string, string> customProperties,
        string propertyName)
    {
        if (preferConfiguration)
        {
            var value = GetNonEmptyValue(configurationProperties, propertyName);
            return string.IsNullOrWhiteSpace(value) ? GetNonEmptyValue(customProperties, propertyName) : value;
        }

        var customValue = GetNonEmptyValue(customProperties, propertyName);
        return string.IsNullOrWhiteSpace(customValue) ? GetNonEmptyValue(configurationProperties, propertyName) : customValue;
    }

    private static string GetNonEmptyValue(IReadOnlyDictionary<string, string> properties, string propertyName)
    {
        return properties.TryGetValue(propertyName, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : string.Empty;
    }

    private static string ReadMaterial(
        bool preferConfiguration,
        IReadOnlyDictionary<string, string> configurationProperties,
        IReadOnlyDictionary<string, string> customProperties)
    {
        var candidates = new[] { "SW-Material", "SW材料", "Material", "材料" };
        foreach (var name in candidates)
        {
            var value = PickPropertyValue(preferConfiguration, configurationProperties, customProperties, name);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private static ISwDMConfiguration? ResolveConfiguration(ISwDMDocument document, string? configuration)
    {
        foreach (var candidate in GetConfigurationCandidates(document, configuration))
        {
            try
            {
                var resolved = document.ConfigurationManager.GetConfigurationByName(candidate) as ISwDMConfiguration;
                if (resolved is not null)
                {
                    return resolved;
                }
            }
            catch
            {
            }
        }

        return null;
    }

    private static IEnumerable<string> GetConfigurationCandidates(ISwDMDocument document, string? configuration)
    {
        if (!string.IsNullOrWhiteSpace(configuration) && !configuration.Equals("custom", StringComparison.OrdinalIgnoreCase))
        {
            yield return configuration.Trim();
        }

        var active = GetActiveConfigurationName(document);
        if (!string.IsNullOrWhiteSpace(active))
        {
            yield return active;
        }

        foreach (var name in GetConfigurationNames(document))
        {
            yield return name;
        }
    }

    private static string NormalizeValue(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
    }

    private static List<string> ToStringList(object? value)
    {
        if (value is null)
        {
            return [];
        }

        if (value is string text)
        {
            return string.IsNullOrWhiteSpace(text) ? [] : [text];
        }

        var result = new List<string>();
        if (value is IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                var textValue = item?.ToString();
                if (!string.IsNullOrWhiteSpace(textValue))
                {
                    result.Add(textValue);
                }
            }
        }

        return result;
    }
}
