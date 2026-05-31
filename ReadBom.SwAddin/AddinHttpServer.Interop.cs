using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using SldWorks;
using SwConst;

namespace ReadBom.SwAddin;

internal sealed partial class AddinHttpServer
{
    private static string NormalizePathForCompare(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return System.IO.Path.GetFullPath(path.Trim());
        }
        catch
        {
            return path.Trim();
        }
    }

    private ModelDoc2 GetActiveModel()
    {
        var model = _swApp.ActiveDoc as ModelDoc2;
        if (model == null)
        {
            throw new InvalidOperationException("SolidWorks 当前没有活动文档");
        }

        return model;
    }

    private static string GetConfigName(ModelDoc2 model, string requested)
    {
        if (!string.IsNullOrWhiteSpace(requested))
        {
            return requested.Equals("custom", StringComparison.OrdinalIgnoreCase) ? string.Empty : requested;
        }

        return model.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty;
    }

    private Dictionary<string, string> ReadAllProperties(CustomPropertyManager manager)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        object namesObj = null;
        object typesObj = null;
        object valuesObj = null;
        object statusObj = null;
        object linkObj = null;

        try
        {
            manager.GetAll3(ref namesObj, ref typesObj, ref valuesObj, ref statusObj, ref linkObj);
        }
        catch
        {
            manager.GetAll2(ref namesObj, ref typesObj, ref valuesObj, ref statusObj);
        }

        var names = ToIndexedStringList(namesObj);
        var values = ToIndexedStringList(valuesObj);
        for (var i = 0; i < names.Count; i++)
        {
            var name = names[i];
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            result[name] = i < values.Count ? values[i] : string.Empty;
        }

        return result;
    }

    private static void WriteProperty(CustomPropertyManager manager, string name, string value)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new InvalidOperationException("属性名不能为空");
        }

        try
        {
            var result = manager.Add3(name, (int)swCustomInfoType_e.swCustomInfoText, value ?? string.Empty, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            if (result >= 0)
            {
                return;
            }
        }
        catch
        {
        }

        try
        {
            var result = manager.Set2(name, value ?? string.Empty);
            if (result >= 0)
            {
                return;
            }
        }
        catch
        {
        }

        try
        {
            var result = manager.Set(name, value ?? string.Empty);
            if (result >= 0)
            {
                return;
            }
        }
        catch
        {
        }

        try
        {
            var result = manager.Add2(name, (int)swCustomInfoType_e.swCustomInfoText, value ?? string.Empty);
            if (result >= 0)
            {
                return;
            }
        }
        catch
        {
        }

        throw new InvalidOperationException("属性写入失败: " + name);
    }

    private static void SaveModel(ModelDoc2 model)
    {
        var errors = 0;
        var warnings = 0;
        var saved = model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
        AddinLog.Write($"SaveModel: title={Safe(() => model.GetTitle())}, saved={saved}, errors={errors}, warnings={warnings}");
        if (!saved || errors != 0)
        {
            throw new InvalidOperationException($"保存模型失败，错误码={errors}，警告码={warnings}");
        }
    }

    private static string NormalizeConfigurationName(string name)
    {
        return string.IsNullOrWhiteSpace(name) || name.Equals("Default", StringComparison.OrdinalIgnoreCase)
            ? string.Empty
            : name.Trim();
    }

    private static void ActivateConfiguration(ModelDoc2 model, string configName)
    {
        if (string.IsNullOrWhiteSpace(configName) || configName.Equals("Default", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        try
        {
            model.ShowConfiguration2(configName);
        }
        catch (Exception ex)
        {
            AddinLog.Write($"ActivateConfiguration ignored: {configName}: {ex.Message}");
        }
    }

    private static List<double> GetModelBoxValues(ModelDoc2 model, string path)
    {
        var docType = GetDocumentTypeFromPath(path);
        object corners = null;
        if (docType == (int)swDocumentTypes_e.swDocPART)
        {
            try { corners = ((PartDoc)model).GetPartBox(true); } catch { }
        }
        else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
        {
            try { corners = ((AssemblyDoc)model).GetBox(1); } catch { }
        }

        return ToDoubleList(corners);
    }

    private static List<double> ToDoubleList(object value)
    {
        var result = new List<double>();
        if (value is Array array)
        {
            foreach (var item in array)
            {
                result.Add(Convert.ToDouble(item));
            }
        }

        return result;
    }

    private static List<string> ToIndexedStringList(object value)
    {
        if (value == null)
        {
            return new List<string>();
        }

        if (value is string text)
        {
            return new List<string> { text };
        }

        if (value is System.Collections.IEnumerable enumerable)
        {
            return enumerable.Cast<object>().Select(x => x?.ToString() ?? string.Empty).ToList();
        }

        return new List<string>();
    }

    private static bool IsSuppressed(Component2 component)
    {
        try { return component.IsSuppressed(); } catch { }
        try { return Convert.ToInt32(component.GetSuppression()) == 0; } catch { }
        return false;
    }

    private static bool IsVirtual(Component2 component)
    {
        try { return component.IsVirtual; } catch { return false; }
    }

    private static string GetDocumentTypeLabel(string path)
    {
        var ext = Path.GetExtension(path ?? string.Empty);
        if (ext.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return "装配体";
        if (ext.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return "零件";
        if (ext.Equals(".slddrw", StringComparison.OrdinalIgnoreCase)) return "工程图";
        return "未知";
    }

    private static string GetAnonymousValue(object row, string name)
    {
        return row.GetType().GetProperty(name)?.GetValue(row)?.ToString() ?? string.Empty;
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex)
        {
            AddinLog.Write($"ReadBom: failed to delete temporary CSV {path}: {ex.Message}");
        }
    }

    private sealed class BomSeed
    {
        public string Path { get; set; }
        public string Config { get; set; }
        public int Quantity { get; set; }
        public ModelDoc2 Model { get; set; }
    }

    private sealed class BomTableCsvTransfer
    {
        public string CsvBase64 { get; set; }
        public int CsvByteCount { get; set; }
        public string Separator { get; set; }
        public List<string> PropertyNames { get; set; } = new List<string>();
        public string MainPath { get; set; }
        public string MainConfiguration { get; set; }
        public int RowCount { get; set; }
    }

    private void WriteJson(HttpListenerContext context, object value)
    {
        var serializeWatch = Stopwatch.StartNew();
        var text = _json.Serialize(value);
        var serializeElapsed = serializeWatch.ElapsedMilliseconds;
        var bytes = Encoding.UTF8.GetBytes(text);
        context.Response.ContentType = "application/json; charset=utf-8";
        context.Response.ContentLength64 = bytes.Length;
        var writeWatch = Stopwatch.StartNew();
        context.Response.OutputStream.Write(bytes, 0, bytes.Length);
        context.Response.OutputStream.Close();
        AddinLog.Write($"HTTP WriteJson bytes={bytes.Length}, serialize={serializeElapsed}ms, write={writeWatch.ElapsedMilliseconds}ms");
    }

    private static T Safe<T>(Func<T> func)
    {
        try { return func(); }
        catch { return default; }
    }

    private const int SwRestore = 9;
    private static readonly IntPtr HwndTopMost = new IntPtr(-1);
    private static readonly IntPtr HwndNoTopMost = new IntPtr(-2);
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpShowWindow = 0x0040;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);
}
