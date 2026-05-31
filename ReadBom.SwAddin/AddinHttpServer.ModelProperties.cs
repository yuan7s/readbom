using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using SldWorks;
using SwConst;

namespace ReadBom.SwAddin;

internal sealed partial class AddinHttpServer
{
    private static string GetBomTableTemplatePath()
    {
        var basePath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles), "SOLIDWORKS", "lang");
        var chinese = Path.Combine(basePath, "chinese-simplified", "bom-standard.sldbomtbt");
        if (File.Exists(chinese))
        {
            return chinese;
        }

        var english = Path.Combine(basePath, "english", "bom-standard.sldbomtbt");
        return File.Exists(english) ? english : string.Empty;
    }

    private static void DeleteBomTableFeature(ModelDoc2 model, BomTableAnnotation bomTable)
    {
        if (model == null || bomTable == null)
        {
            return;
        }

        try
        {
            var feature = bomTable.BomFeature?.GetFeature();
            if (feature == null)
            {
                return;
            }

            model.ClearSelection2(true);
            feature.Select2(false, 0);
            model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
            model.ClearSelection2(true);
            AddinLog.Write("ReadBom: hidden BOM table deleted");
        }
        catch (Exception ex)
        {
            AddinLog.Write($"ReadBom: failed to delete hidden BOM table: {ex.Message}");
        }
    }

    private object CreateBomRow(ModelDoc2 model, string path, string configName, int quantity, string[] propertyNames, DrawingLookup drawingLookup)
    {
        var rowWatch = Stopwatch.StartNew();
        var documentType = GetDocumentTypeLabel(path);
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var available = new List<string>();
        long propertyElapsed = 0;
        if (model != null)
        {
            var propertyWatch = Stopwatch.StartNew();
            var cfgManager = model.Extension.get_CustomPropertyManager(configName ?? string.Empty);
            var customManager = model.Extension.get_CustomPropertyManager(string.Empty);
            var cfgProps = ReadAllProperties(cfgManager);
            var customProps = ReadAllProperties(customManager);
            available.AddRange(cfgProps.Keys);
            available.AddRange(customProps.Keys);

            foreach (var name in propertyNames ?? Array.Empty<string>())
            {
                properties[name] = cfgProps.TryGetValue(name, out var cfgValue) && !string.IsNullOrWhiteSpace(cfgValue)
                    ? cfgValue
                    : customProps.TryGetValue(name, out var customValue) ? customValue : string.Empty;
            }
            propertyElapsed = propertyWatch.ElapsedMilliseconds;
        }

        var materialWatch = Stopwatch.StartNew();
        var material = documentType == "零件" ? ReadSolidWorksMaterial(model, configName) : "无需设置";
        var materialElapsed = materialWatch.ElapsedMilliseconds;
        if (documentType == "零件" && string.IsNullOrWhiteSpace(material))
        {
            material = "未设置";
        }

        var drawingWatch = Stopwatch.StartNew();
        var hasDrawing = drawingLookup.HasSiblingDrawing(path);
        var drawingElapsed = drawingWatch.ElapsedMilliseconds;
        if (rowWatch.ElapsedMilliseconds > 300)
        {
            AddinLog.Write($"ReadBom slow row: {Path.GetFileName(path)} total={rowWatch.ElapsedMilliseconds}ms, props={propertyElapsed}ms, material={materialElapsed}ms, drawing={drawingElapsed}ms");
        }

        return new
        {
            documentType,
            drawingStatus = hasDrawing ? "有工程图" : "无工程图",
            fileName = Path.GetFileNameWithoutExtension(path ?? string.Empty),
            configuration = string.IsNullOrWhiteSpace(configName) ? "Default" : configName,
            quantity,
            material,
            fullPath = path ?? string.Empty,
            properties,
            availablePropertyNames = available.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList()
        };
    }

    private sealed class DrawingLookup
    {
        private readonly Dictionary<string, HashSet<string>> _directoryDrawings = new(StringComparer.OrdinalIgnoreCase);

        public bool HasSiblingDrawing(string modelPath)
        {
            if (string.IsNullOrWhiteSpace(modelPath))
            {
                return false;
            }

            var type = GetDocumentTypeLabel(modelPath);
            if (type == "工程图")
            {
                return true;
            }

            try
            {
                var directory = Path.GetDirectoryName(modelPath);
                var fileName = Path.GetFileNameWithoutExtension(modelPath);
                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName) || !Directory.Exists(directory))
                {
                    return false;
                }

                if (!_directoryDrawings.TryGetValue(directory, out var drawings))
                {
                    var watch = Stopwatch.StartNew();
                    drawings = Directory.EnumerateFiles(directory, "*.slddrw")
                        .Select(Path.GetFileNameWithoutExtension)
                        .Where(name => !string.IsNullOrWhiteSpace(name))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    _directoryDrawings[directory] = drawings;
                    if (watch.ElapsedMilliseconds > 200)
                    {
                        AddinLog.Write($"ReadBom timing: drawing directory scan {directory}, drawings={drawings.Count}, elapsed={watch.ElapsedMilliseconds}ms");
                    }
                }

                return drawings.Contains(fileName);
            }
            catch
            {
                return false;
            }
        }
    }

    private static bool HasSiblingDrawing(string modelPath)
    {
        if (string.IsNullOrWhiteSpace(modelPath))
        {
            return false;
        }

        var type = GetDocumentTypeLabel(modelPath);
        if (type == "工程图")
        {
            return true;
        }

        try
        {
            var directory = Path.GetDirectoryName(modelPath);
            var fileName = Path.GetFileNameWithoutExtension(modelPath);
            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName) || !Directory.Exists(directory))
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

    private static string ReadSolidWorksMaterial(ModelDoc2 model, string configName)
    {
        if (model == null)
        {
            return string.Empty;
        }

        var partDoc = model as PartDoc;
        if (partDoc == null)
        {
            return string.Empty;
        }

        foreach (var config in GetMaterialConfigCandidates(model, configName))
        {
            try
            {
                string databaseName;
                var value = partDoc.GetMaterialPropertyName2(config, out databaseName);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            catch
            {
                // Try the next material API/configuration candidate.
            }
        }

        try
        {
            string databaseName;
            var value = partDoc.GetMaterialPropertyName(out databaseName);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }
        catch
        {
        }

        try
        {
            var value = partDoc.MaterialUserName;
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }
        catch
        {
        }

        try
        {
            var value = partDoc.MaterialIdName;
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }
        catch
        {
        }

        return string.Empty;
    }

    private static IEnumerable<string> GetMaterialConfigCandidates(ModelDoc2 model, string configName)
    {
        var names = new List<string>();
        AddMaterialConfigCandidate(names, configName);
        AddMaterialConfigCandidate(names, model.ConfigurationManager?.ActiveConfiguration?.Name);
        AddMaterialConfigCandidate(names, string.Empty);
        AddMaterialConfigCandidate(names, "Default");
        return names;
    }

    private static void AddMaterialConfigCandidate(ICollection<string> names, string name)
    {
        var candidate = string.IsNullOrWhiteSpace(name) ? string.Empty : name.Trim();
        if (names.Any(x => string.Equals(x, candidate, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        names.Add(candidate);
    }

    private object Rebuild()
    {
        var model = GetActiveModel();
        AddinLog.Write($"Rebuild: {Safe(() => model.GetTitle())}");
        model.ForceRebuild3(false);
        return new { rebuilt = true };
    }

    private object Save()
    {
        var model = GetActiveModel();
        int errors = 0;
        int warnings = 0;
        AddinLog.Write($"Save: {Safe(() => model.GetTitle())}");
        var saved = model.Save3(1, ref errors, ref warnings);
        AddinLog.Write($"Save result: saved={saved}, errors={errors}, warnings={warnings}");
        return new { saved, errors, warnings };
    }

    private ModelDoc2 TryGetOpenModelByPath(string path)
    {
        try
        {
            var model = _swApp.GetOpenDocumentByName(path) as ModelDoc2;
            if (model != null)
            {
                return model;
            }
        }
        catch
        {
        }

        var normalizedPath = NormalizePathForCompare(path);
        try
        {
            var model = _swApp.GetFirstDocument() as ModelDoc2;
            while (model != null)
            {
                var modelPath = NormalizePathForCompare(Safe(() => model.GetPathName()));
                if (!string.IsNullOrWhiteSpace(modelPath)
                    && string.Equals(modelPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
                {
                    return model;
                }

                model = model.GetNext() as ModelDoc2;
            }
        }
        catch
        {
        }

        return null;
    }

    private void ActivateDocument(string title, string path)
    {
        foreach (var candidate in GetActivationTitleCandidates(title, path))
        {
            try
            {
                var errors = 0;
                _swApp.ActivateDoc3(candidate, false, 0, ref errors);
                AddinLog.Write($"ActivateDoc3 candidate={candidate}, errors={errors}");
                if (errors == 0)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                AddinLog.Write($"ActivateDoc3 failed candidate={candidate}: {ex.Message}");
            }
        }
    }

    private static IEnumerable<string> GetActivationTitleCandidates(string title, string path)
    {
        if (!string.IsNullOrWhiteSpace(title))
        {
            yield return title;
        }

        var fileName = System.IO.Path.GetFileName(path);
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            yield return fileName;
        }

        var fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(path);
        if (!string.IsNullOrWhiteSpace(fileNameWithoutExtension))
        {
            yield return fileNameWithoutExtension;
        }
    }

    private void ActivateSolidWorksWindow()
    {
        TryActivateSolidWorksFrame();
        var hwnd = GetSolidWorksFrameHwnd();
        if (hwnd == IntPtr.Zero)
        {
            AddinLog.Write("ActivateSolidWorksWindow skipped: hwnd=0");
            return;
        }

        var foregroundBefore = GetForegroundWindow();
        var show = ShowWindow(hwnd, SwRestore);
        var top = BringWindowToTop(hwnd);
        var setPosTop = SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpShowWindow);
        var setPosNormal = SetWindowPos(hwnd, HwndNoTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpShowWindow);
        var foreground = SetForegroundWindow(hwnd);
        SwitchToThisWindow(hwnd, true);

        var foregroundAfter = GetForegroundWindow();
        AddinLog.Write($"ActivateSolidWorksWindow hwnd={hwnd}, before={foregroundBefore}, after={foregroundAfter}, show={show}, top={top}, topMost={setPosTop}, normal={setPosNormal}, foreground={foreground}");
    }

    private void TryActivateSolidWorksFrame()
    {
        try
        {
            _swApp.Visible = true;
        }
        catch (Exception ex)
        {
            AddinLog.Write("ActivateSolidWorksWindow Visible ignored: " + ex.Message);
        }

        try
        {
            var frame = _swApp.Frame();
            frame.GetType().InvokeMember(
                "Activate",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                frame,
                null);
        }
        catch (Exception ex)
        {
            AddinLog.Write("ActivateSolidWorksWindow Frame.Activate ignored: " + ex.Message);
        }
    }

    private IntPtr GetSolidWorksFrameHwnd()
    {
        try
        {
            var frame = _swApp.Frame();
            var hwndValue = frame.GetType().InvokeMember(
                "GetHWnd",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                frame,
                null);
            var hwnd = Convert.ToInt64(hwndValue);
            return hwnd == 0 ? IntPtr.Zero : new IntPtr(hwnd);
        }
        catch (Exception ex)
        {
            AddinLog.Write("GetSolidWorksFrameHwnd failed: " + ex.Message);
            return IntPtr.Zero;
        }
    }

    private static int GetDocumentTypeFromPath(string path)
    {
        var extension = System.IO.Path.GetExtension(path);
        if (extension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return (int)swDocumentTypes_e.swDocPART;
        if (extension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return (int)swDocumentTypes_e.swDocASSEMBLY;
        if (extension.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return (int)swDocumentTypes_e.swDocDRAWING;
        return 0;
    }

}
