using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using SolidWorksTools;

namespace ReadBom.SwAddin;

[ComVisible(true)]
[Guid("B5B7E80B-3D31-4B72-9A92-8E1A4A3E4F18")]
[ProgId("ReadBom.SwAddin")]
[SwAddin(Description = AddinDescription, Title = AddinTitle, LoadAtStartup = true)]
public sealed class SwAddin : SolidWorks.Interop.swpublished.SwAddin
{
    private const string AddinTitle = "ReadBom HTTP Addin";
    private const string AddinDescription = "ReadBom local HTTP bridge for SolidWorks commands.";
    private SldWorks.SldWorks _swApp;
    private int _cookie;
    private AddinHttpServer _server;
    private Control _mainThreadControl;

    public bool ConnectToSW(object thisSw, int cookie)
    {
        try
        {
            AddinLog.Write("ConnectToSW called");
            _swApp = (SldWorks.SldWorks)thisSw;
            _cookie = cookie;

            // Create a hidden WinForms Control on SW's main thread to serve as
            // the dispatcher for marshaling all COM calls back to this thread.
            _mainThreadControl = new Control();
            var _ = _mainThreadControl.Handle; // Force handle creation on this thread

            try
            {
                _swApp.SetAddinCallbackInfo2(0, this, _cookie);
                AddinLog.Write("SetAddinCallbackInfo2 ok");
            }
            catch (Exception ex)
            {
                AddinLog.Write("SetAddinCallbackInfo2 ignored: " + ex.Message);
            }

            _server = new AddinHttpServer(_swApp, _mainThreadControl, "http://127.0.0.1:32127/");
            _server.Start();
            AddinLog.Write("HTTP server started");
            return true;
        }
        catch (Exception ex)
        {
            AddinLog.Write("ConnectToSW failed: " + ex);
            return false;
        }
    }

    public bool DisconnectFromSW()
    {
        AddinLog.Write("DisconnectFromSW called");
        _server?.Dispose();
        _server = null;
        _swApp = null;
        _cookie = 0;
        return true;
    }

    [ComRegisterFunction]
    public static void Register(Type type)
    {
        var attribute = (SwAddinAttribute)System.Attribute.GetCustomAttribute(type, typeof(SwAddinAttribute));
        var title = attribute?.Title ?? AddinTitle;
        var description = attribute?.Description ?? AddinDescription;
        var loadAtStartup = attribute?.LoadAtStartup == true ? 1 : 0;

        using (var key = Registry.LocalMachine.CreateSubKey($@"SOFTWARE\SolidWorks\AddIns\{{{type.GUID}}}"))
        {
            key.SetValue(null, 0, RegistryValueKind.DWord);
            key.SetValue("Title", title, RegistryValueKind.String);
            key.SetValue("Description", description, RegistryValueKind.String);
        }

        using (var key = Registry.CurrentUser.CreateSubKey($@"SOFTWARE\SolidWorks\AddIns\{{{type.GUID}}}"))
        {
            key.SetValue(null, 0, RegistryValueKind.DWord);
            key.SetValue("Title", title, RegistryValueKind.String);
            key.SetValue("Description", description, RegistryValueKind.String);
        }

        using (var key = Registry.CurrentUser.CreateSubKey($@"SOFTWARE\SolidWorks\AddInsStartup\{{{type.GUID}}}"))
        {
            key.SetValue(null, loadAtStartup, RegistryValueKind.DWord);
        }
    }

    [ComUnregisterFunction]
    public static void Unregister(Type type)
    {
        Registry.LocalMachine.DeleteSubKey($@"SOFTWARE\SolidWorks\AddIns\{{{type.GUID}}}", false);
        Registry.CurrentUser.DeleteSubKey($@"SOFTWARE\SolidWorks\AddIns\{{{type.GUID}}}", false);
        Registry.CurrentUser.DeleteSubKey($@"SOFTWARE\SolidWorks\AddInsStartup\{{{type.GUID}}}", false);
    }
}

internal static class AddinLog
{
    public static readonly string DirectoryPath = GetAddinDirectory();
    private static readonly string LogPath = Path.Combine(DirectoryPath, "SwAddin.log");

    public static void Write(string message)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
            File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{System.Environment.NewLine}");
        }
        catch
        {
            // Logging must not block Add-in loading.
        }
    }

    private static string GetAddinDirectory()
    {
        try
        {
            var location = Assembly.GetExecutingAssembly().Location;
            var directory = Path.GetDirectoryName(location);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                return directory;
            }
        }
        catch
        {
        }

        return AppDomain.CurrentDomain.BaseDirectory;
    }
}
