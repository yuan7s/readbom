using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace readbom;

internal static partial class SolidWorksReader
{
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

    internal static bool HasSiblingDrawing(string modelPath)
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

            if (!Directory.Exists(directory))
            {
                return false;
            }

            return Directory.EnumerateFiles(directory, "*.slddrw")
                .Any(file => string.Equals(Path.GetFileNameWithoutExtension(file), fileName, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return false;
        }
    }

    internal static string OpenDocument(dynamic swApp, string path)
    {
        var model = OpenModelVisible(swApp, path, out string openError);
        if (model is null)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(openError)
                ? $"无法打开文件: {path}"
                : $"无法打开文件: {path}。{openError}");
        }

        var title = GetModelTitleSafe(model);
        ActivateOpenedDocument(swApp, title, path);
        ActivateSolidWorksWindow(swApp);
        return string.IsNullOrWhiteSpace(title) ? Path.GetFileName(path) : title;
    }

    private static dynamic? OpenModelVisible(dynamic swApp, string path, out string error)
    {
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(path))
        {
            error = "路径为空";
            return null;
        }

        path = Path.GetFullPath(path.Trim());
        if (!File.Exists(path))
        {
            error = "文件不存在";
            return null;
        }

        var docType = GetDocumentTypeFromPath(path);
        if (docType == 0)
        {
            error = "不支持的 SolidWorks 文件类型";
            return null;
        }

        var opened = TryGetOpenModelByPath(swApp, path);
        if (opened is not null)
        {
            return opened;
        }

        try
        {
            try
            {
                swApp.DocumentVisible(true, docType);
            }
            catch
            {
                // Some SW versions or interop paths do not expose DocumentVisible through late binding.
            }

            int err = 0, warn = 0;
            dynamic? model;
            if (TryAsStrongSwApp(swApp, out SldWorks.SldWorks typedApp))
            {
                typedApp.DocumentVisible(true, docType);
                model = typedApp.OpenDoc6(path, docType, 0, "", ref err, ref warn);
            }
            else
            {
                model = swApp.OpenDoc6(path, docType, 0, "", ref err, ref warn);
            }

            if (model is null)
            {
                error = $"OpenDoc6 返回空，错误码={err}，警告码={warn}";
                return null;
            }

            if (err != 0)
            {
                error = $"OpenDoc6 错误码={err}，警告码={warn}";
            }

            return model;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return null;
        }
    }

    private static void ActivateOpenedDocument(dynamic swApp, string title, string path)
    {
        foreach (var candidate in GetActivationTitleCandidates(title, path))
        {
            try
            {
                int errors = 0;
                swApp.ActivateDoc3(candidate, false, 0, ref errors);
                if (errors == 0)
                {
                    return;
                }
            }
            catch
            {
                // Try the next title candidate.
            }
        }
    }

    private static IEnumerable<string> GetActivationTitleCandidates(string title, string path)
    {
        if (!string.IsNullOrWhiteSpace(title))
        {
            yield return title;
        }

        var fileName = Path.GetFileName(path);
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            yield return fileName;
        }

        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
        if (!string.IsNullOrWhiteSpace(fileNameWithoutExtension))
        {
            yield return fileNameWithoutExtension;
        }
    }

    internal static void ActivateSolidWorksWindow(dynamic swApp)
    {
        var hwnd = GetSolidWorksFrameHwnd(swApp);
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        ShowWindow(hwnd, SwRestore);
        SetForegroundWindow(hwnd);
        BringWindowToTop(hwnd);
    }

    private static IntPtr GetSolidWorksFrameHwnd(dynamic swApp)
    {
        try
        {
            dynamic frame = TryAsStrongSwApp(swApp, out SldWorks.SldWorks typedApp) ? typedApp.Frame() : swApp.Frame();
            var hwnd = Convert.ToInt64(frame.GetHWnd(), CultureInfo.InvariantCulture);
            return hwnd == 0 ? IntPtr.Zero : new IntPtr(hwnd);
        }
        catch
        {
            return IntPtr.Zero;
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

        path = path.Trim();

        var opened = TryGetOpenModelByPath(swApp, path);
        if (opened is not null)
        {
            return opened;
        }

        if (!File.Exists(path))
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

            if (TryAsStrongSwApp(swApp, out SldWorks.SldWorks typedApp))
            {
                return typedApp.OpenDoc6(path, docType, 1, "", ref err, ref warn);
            }

            return swApp.OpenDoc6(path, docType, 1, "", ref err, ref warn);
        }
        catch
        {
            return null;
        }
    }

    private static dynamic? TryGetOpenModelByPath(dynamic swApp, string path)
    {
        try
        {
            var doc = swApp.GetOpenDocumentByName(path);
            if (doc is not null)
            {
                return doc;
            }
        }
        catch { }

        var normalizedPath = NormalizePathForCompare(path);
        try
        {
            dynamic? doc = GetFirstOpenDocumentSafe(swApp);
            while (doc is not null)
            {
                var docPath = NormalizePathForCompare(GetModelPathSafe(doc));
                if (!string.IsNullOrWhiteSpace(docPath) && string.Equals(docPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }

                try { doc = doc.GetNext(); } catch { break; }
            }
        }
        catch { }

        return null;
    }

    private static string NormalizePathForCompare(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(path.Trim());
        }
        catch
        {
            return path.Trim();
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

    private static bool TryAsStrongCustomPropertyManager(object manager, out SldWorks.ICustomPropertyManager typedManager)
    {
        try
        {
            typedManager = (SldWorks.ICustomPropertyManager)manager;
            return typedManager != null;
        }
        catch
        {
            typedManager = null!;
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

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    private const int SwRestore = 9;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowTextLength(IntPtr hWnd);
}
