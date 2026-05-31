using System.Globalization;
using System.IO;
using System.Reflection;
using System.Collections;

namespace readbom;

internal static partial class SolidWorksReader
{
    private static BomRow ReadFromModel(dynamic model, string configName, string path, PropertySourceMode sourceMode, PropertyMappingConfig mapping)
    {
        var cfgManager = GetCustomPropertyManagerSafe(model, configName);
        var customManager = GetCustomPropertyManagerSafe(model, "");
        var preferConfig = sourceMode == PropertySourceMode.CurrentConfiguration;
        var fileName = Path.GetFileNameWithoutExtension(path);
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
            DrawingStatus = string.Empty,
            DrawingIconPath = string.Empty,
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
            foreach (var config in GetMaterialConfigCandidates(model, configName))
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

        foreach (var config in GetMaterialConfigCandidates(model, configName))
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

    private static bool TryAsStrongAssemblyDoc(dynamic model, out SldWorks.IAssemblyDoc assemblyDoc)
    {
        try
        {
            assemblyDoc = (SldWorks.IAssemblyDoc)model;
            return assemblyDoc != null;
        }
        catch
        {
            assemblyDoc = null!;
            return false;
        }
    }

    private static IEnumerable<string> GetMaterialConfigCandidates(dynamic model, string configName)
    {
        var candidates = new List<string>();
        AddMaterialConfigCandidate(candidates, configName);
        try
        {
            AddMaterialConfigCandidate(candidates, Convert.ToString(model.ConfigurationManager.ActiveConfiguration.Name, CultureInfo.InvariantCulture));
        }
        catch { }
        AddMaterialConfigCandidate(candidates, string.Empty);
        AddMaterialConfigCandidate(candidates, "Default");
        return candidates;
    }

    private static void AddMaterialConfigCandidate(ICollection<string> candidates, string? configName)
    {
        var candidate = string.IsNullOrWhiteSpace(configName) ? string.Empty : configName.Trim();
        if (candidates.Any(x => string.Equals(x, candidate, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        candidates.Add(candidate);
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
            if (TryReadAllStrongByGetAll3(typedManager, result)) return result;
            if (TryReadAllStrongByGetAll2(typedManager, result)) return result;

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
        }

        if (TryReadAllByGetAll3(managerObj, result)) return result;
        if (TryReadAllByGetAll2(managerObj, result)) return result;

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
            AddPropertyList(result, namesObj, valuesObj);
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
            AddPropertyList(result, namesObj, valuesObj);
            return result.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    private static void AddPropertyList(Dictionary<string, string> result, object? namesObj, object? valuesObj)
    {
        var names = ToIndexedStringList(namesObj);
        var values = ToIndexedStringList(valuesObj);
        for (var i = 0; i < names.Count; i++)
        {
            var key = names[i];
            if (string.IsNullOrWhiteSpace(key)) continue;
            var value = i < values.Count ? values[i] : string.Empty;
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

        // GetAll3(out names, out types, out values, out evaluationStatuses, out linkToProperty)
        AddPropertyList(result, args[0], args[2]);

        return result.Count > 0;
    }

    private static bool TryReadAllByGetAll2(object manager, Dictionary<string, string> result)
    {
        var args = new object?[] { null, null, null, null };
        if (!TryInvokeCom(manager, "GetAll2", args))
        {
            return false;
        }

        // GetAll2(out names, out types, out values, out evaluationStatuses)
        AddPropertyList(result, args[0], args[2]);

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

    private static List<string> ToIndexedStringList(object? value)
    {
        var list = new List<string>();
        if (value is null) return list;
        if (value is string s)
        {
            list.Add(s);
            return list;
        }

        if (value is IEnumerable enumerable)
        {
            foreach (var item in enumerable)
            {
                list.Add(item?.ToString() ?? string.Empty);
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

}
