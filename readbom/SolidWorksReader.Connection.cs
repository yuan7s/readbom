using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace readbom;

internal static partial class SolidWorksReader
{
    public sealed record ConnectionInfo(int ProcessId, string Version, string DisplayVersion, int OpenDocumentCount, string ActiveDocumentTitle, string MainWindowTitle);
    private static int _lastConnectedProcessId;

    private sealed class ProgressTracker
    {
        private readonly Action<ReadProgress>? _progress;
        private readonly string _message;
        private readonly int _total;
        private int _completed;

        public ProgressTracker(Action<ReadProgress>? progress, string message, int total)
        {
            _progress = progress;
            _message = message;
            _total = Math.Max(total, 1);
        }

        public void Report()
        {
            Report(_message);
        }

        public void Report(string message)
        {
            _completed = Math.Min(_completed + 1, _total);
            _progress?.Invoke(new ReadProgress(message, _completed, _total));
        }

        public void ReportMessage(string message)
        {
            _progress?.Invoke(new ReadProgress(message, _completed, _total));
        }
    }

    public static dynamic Connect()
    {
        List<string> rotNames = [];
        var swProcessCount = Process.GetProcessesByName("SLDWORKS").Length;
        var attempts = swProcessCount > 0 ? 10 : 1;
        for (var attempt = 1; attempt <= attempts; attempt++)
        {
            // Prefer ROT enumeration so we can pick the instance that actually has opened documents.
            var fromRot = FindBestRunningSolidWorksFromRot(out rotNames);
            if (fromRot != null)
            {
                _lastConnectedProcessId = fromRot.ProcessId;
                return fromRot.App;
            }

            if (TryGetActiveSolidWorksObject(out var runningObject) && runningObject is not null)
            {
                dynamic app = runningObject;
                if (IsUsableSolidWorksApp(app))
                {
                    return runningObject;
                }
            }

            if (attempt < attempts)
            {
                Thread.Sleep(500);
            }
        }

        var rotSummary = rotNames.Count == 0
            ? "ROT中没有发现SolidWorks注册项"
            : "ROT候选: " + string.Join(" | ", rotNames.Take(5));
        throw new InvalidOperationException(
            swProcessCount > 0
                ? $"检测到 {swProcessCount} 个 SLDWORKS.exe，但没有找到可用的 SolidWorks COM 实例。{rotSummary}。请确认本工具和 SolidWorks 使用相同权限运行。"
                : "未检测到已运行的 SolidWorks 实例。请先启动并打开模型后再连接。");
    }

    internal static List<string> GetRelatedFiles(dynamic swApp, string mainAssemblyPath)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(mainAssemblyPath))
        {
            return [];
        }

        mainAssemblyPath = NormalizePathForCompare(mainAssemblyPath);
        AddRelatedFile(result, mainAssemblyPath);

        try
        {
            object? dependencies = TryAsStrongSwApp(swApp, out SldWorks.SldWorks typedApp)
                ? typedApp.GetDocumentDependencies2(mainAssemblyPath, true, true, true)
                : swApp.GetDocumentDependencies2(mainAssemblyPath, true, true, true);

            foreach (var value in ToStringList(dependencies))
            {
                AddRelatedFile(result, value);
            }
        }
        catch
        {
        }

        return result.ToList();
    }

    private static void AddRelatedFile(ISet<string> files, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var candidate = value.Trim();
        var extension = Path.GetExtension(candidate);
        if (!extension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".slddrw", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        try
        {
            files.Add(Path.GetFullPath(candidate));
        }
        catch
        {
            files.Add(candidate);
        }
    }

    private static bool TryGetActiveSolidWorksObject(out object? runningObject)
    {
        runningObject = null;
        foreach (var progId in new[] { "SldWorks.Application", "SolidWorks.Application" })
        {
            try
            {
                if (CLSIDFromProgID(progId, out var clsid) != 0)
                {
                    continue;
                }

                GetActiveObject(ref clsid, IntPtr.Zero, out runningObject);
                if (runningObject != null)
                {
                    return true;
                }
            }
            catch
            {
                runningObject = null;
            }
        }

        try
        {
            var clsid = new Guid("72B5B460-38D4-11D0-BD8B-00A0C911CE86");
            GetActiveObject(ref clsid, IntPtr.Zero, out runningObject);
            return runningObject != null;
        }
        catch
        {
            runningObject = null;
            return false;
        }
    }

    private sealed record RotCandidate(object App, int ProcessId, string DisplayName);

    private static RotCandidate? FindBestRunningSolidWorksFromRot(out List<string> rotNames)
    {
        IRunningObjectTable? rot = null;
        IEnumMoniker? enumMoniker = null;
        var candidates = new List<RotCandidate>();
        rotNames = new List<string>();
        try
        {
            GetRunningObjectTable(0, out rot);
            rot.EnumRunning(out enumMoniker);
            var monikers = new IMoniker[1];
            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                IBindCtx? ctx = null;
                try
                {
                    CreateBindCtx(0, out ctx);
                    monikers[0].GetDisplayName(ctx, null, out var displayName);
                    if (string.IsNullOrWhiteSpace(displayName) || !LooksLikeSolidWorksRotName(displayName))
                    {
                        continue;
                    }

                    rotNames.Add(displayName);
                    rot.GetObject(monikers[0], out var obj);
                    if (obj != null)
                    {
                        candidates.Add(new RotCandidate(obj, ExtractPidFromRotName(displayName), displayName));
                    }
                }
                catch
                {
                    // ignore broken ROT entry
                }
                finally
                {
                    if (ctx != null) Marshal.ReleaseComObject(ctx);
                }
            }
        }
        catch
        {
            // ignore and fall back
        }
        finally
        {
            if (enumMoniker != null) Marshal.ReleaseComObject(enumMoniker);
            if (rot != null) Marshal.ReleaseComObject(rot);
        }

        object? best = null;
        var bestScore = -1;
        foreach (var candidate in candidates)
        {
            try
            {
                dynamic sw = candidate.App;
                var info = CheckConnection(sw);
                var effectivePid = info.ProcessId > 0 ? info.ProcessId : candidate.ProcessId;
                if (!IsUsableConnection(info) && effectivePid <= 0)
                {
                    continue;
                }
                var score = (info.OpenDocumentCount > 0 ? 10 : 0)
                            + (!string.IsNullOrWhiteSpace(info.ActiveDocumentTitle) ? 1 : 0)
                            + (effectivePid > 0 ? 1 : 0);
                if (score > bestScore)
                {
                    bestScore = score;
                    best = candidate;
                }
            }
            catch
            {
                // ignore bad candidate
            }
        }

        return best is RotCandidate rotCandidate
               && (rotCandidate.ProcessId > 0 || IsUsableSolidWorksApp(rotCandidate.App))
            ? rotCandidate
            : null;
    }

    private static int ExtractPidFromRotName(string displayName)
    {
        var marker = "PID_";
        var idx = displayName.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return 0;
        idx += marker.Length;
        var end = idx;
        while (end < displayName.Length && char.IsDigit(displayName[end]))
        {
            end++;
        }

        return int.TryParse(displayName[idx..end], out var pid) ? pid : 0;
    }

    private static bool IsUsableSolidWorksApp(dynamic swApp)
    {
        try
        {
            var info = CheckConnection(swApp);
            return IsUsableConnection(info);
        }
        catch
        {
            return false;
        }
    }

    private static bool LooksLikeSolidWorksRotName(string displayName)
    {
        return displayName.Contains("SolidWorks", StringComparison.OrdinalIgnoreCase)
               || displayName.Contains("SldWorks", StringComparison.OrdinalIgnoreCase)
               || displayName.Contains("solidworks_pid_", StringComparison.OrdinalIgnoreCase)
               || displayName.Contains("SldWorks.Application", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUsableConnection(ConnectionInfo info)
    {
        return info.ProcessId > 0
               || info.OpenDocumentCount > 0
               || !string.IsNullOrWhiteSpace(info.ActiveDocumentTitle)
               || !string.IsNullOrWhiteSpace(info.MainWindowTitle)
               || !string.Equals(info.Version, "Unknown", StringComparison.OrdinalIgnoreCase);
    }

    public static ConnectionInfo CheckConnection(dynamic swApp)
    {
        var processId = GetSwProcessId(swApp);
        if (processId <= 0)
        {
            processId = _lastConnectedProcessId;
        }
        var mainWindowTitle = GetSwMainWindowTitle(swApp);
        var version = GetRevisionNumberSafe(swApp);

        var openCount = 0;
        var activeTitle = string.Empty;
        try
        {
            dynamic? doc = GetActiveDocSafe(swApp) ?? GetFirstOpenDocumentSafe(swApp);
            while (doc != null)
            {
                openCount++;
                if (string.IsNullOrWhiteSpace(activeTitle))
                {
                    try { activeTitle = (string?)doc.GetTitle() ?? string.Empty; } catch { }
                }

                try { doc = doc.GetNext(); } catch { break; }
            }
        }
        catch
        {
            // ignore
        }

        if (string.IsNullOrWhiteSpace(activeTitle))
        {
            activeTitle = ExtractDocumentTitleFromWindowTitle(mainWindowTitle);
        }

        return new ConnectionInfo(processId, version, FormatSolidWorksVersion(version, mainWindowTitle), openCount, activeTitle, mainWindowTitle);
    }

    private static string FormatSolidWorksVersion(string revisionNumber, string mainWindowTitle)
    {
        if (!string.IsNullOrWhiteSpace(mainWindowTitle))
        {
            var dash = mainWindowTitle.IndexOf(" - ", StringComparison.Ordinal);
            var titlePrefix = dash > 0 ? mainWindowTitle[..dash].Trim() : mainWindowTitle.Trim();
            if (titlePrefix.Contains("SOLIDWORKS", StringComparison.OrdinalIgnoreCase))
            {
                return titlePrefix;
            }
        }

        var parts = revisionNumber.Split('.', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 1 && int.TryParse(parts[0], out var major))
        {
            var year = major + 1992;
            var sp = parts.Length >= 2 && int.TryParse(parts[1], out var servicePack)
                ? $" SP{servicePack}.0"
                : string.Empty;
            return $"SOLIDWORKS {year}{sp}";
        }

        return "Unknown";
    }

    private static string GetSwMainWindowTitle(dynamic swApp)
    {
        try
        {
            SldWorks.SldWorks typedApp;
            dynamic frame = TryAsStrongSwApp(swApp, out typedApp) ? typedApp.Frame() : swApp.Frame();
            var hwnd = (int)frame.GetHWnd();
            if (hwnd == 0) return string.Empty;
            var len = GetWindowTextLength(new IntPtr(hwnd));
            if (len <= 0) return string.Empty;
            var sb = new StringBuilder(len + 1);
            GetWindowText(new IntPtr(hwnd), sb, sb.Capacity);
            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static int GetSwProcessId(dynamic swApp)
    {
        SldWorks.SldWorks typedApp;
        if (TryAsStrongSwApp(swApp, out typedApp))
        {
            try
            {
                var pid = typedApp.GetProcessID();
                if (pid > 0) return pid;
            }
            catch
            {
                // ignore and try dynamic/frame fallback
            }
        }

        try
        {
            var pid = (int)swApp.GetProcessID();
            if (pid > 0) return pid;
        }
        catch
        {
            // ignore and try frame hwnd
        }

        try
        {
            dynamic frame = swApp.Frame();
            var hwnd = (int)frame.GetHWnd();
            if (hwnd != 0)
            {
                GetWindowThreadProcessId(new IntPtr(hwnd), out var pid);
                if (pid > 0) return pid;
            }
        }
        catch
        {
            // ignore
        }

        return 0;
    }

    private static string GetRevisionNumberSafe(dynamic swApp)
    {
        SldWorks.SldWorks typedApp;
        if (TryAsStrongSwApp(swApp, out typedApp))
        {
            try
            {
                return typedApp.RevisionNumber();
            }
            catch
            {
                // ignore and try dynamic fallback
            }
        }

        try
        {
            return (string?)swApp.RevisionNumber() ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    private static bool TryAsStrongSwApp(dynamic swApp, out SldWorks.SldWorks typedApp)
    {
        try
        {
            typedApp = (SldWorks.SldWorks)swApp;
            return typedApp != null;
        }
        catch
        {
            typedApp = null!;
            return false;
        }
    }

}
