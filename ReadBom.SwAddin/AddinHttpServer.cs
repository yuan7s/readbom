using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using SldWorks;
using SwConst;

namespace ReadBom.SwAddin;

internal sealed partial class AddinHttpServer : IDisposable
{
    private readonly SldWorks.SldWorks _swApp;
    private readonly System.Windows.Forms.Control _mainThreadControl;
    private readonly string _prefix;
    private readonly JavaScriptSerializer _json = new JavaScriptSerializer();
    private HttpListener _listener;
    private CancellationTokenSource _cts;
    private Task _listenTask;

    public AddinHttpServer(SldWorks.SldWorks swApp, string prefix)
    {
        _swApp = swApp;
        _prefix = prefix.EndsWith("/") ? prefix : prefix + "/";
        _json.MaxJsonLength = int.MaxValue;
    }

    public AddinHttpServer(SldWorks.SldWorks swApp, System.Windows.Forms.Control mainThreadControl, string prefix)
        : this(swApp, prefix)
    {
        _mainThreadControl = mainThreadControl;
    }

    public void Start()
    {
        if (_listener != null)
        {
            return;
        }

        _cts = new CancellationTokenSource();
        _listener = new HttpListener();
        _listener.Prefixes.Add(_prefix);
        AddinLog.Write("Starting HTTP listener: " + _prefix);
        _listener.Start();
        _listenTask = Task.Run(() => ListenLoop(_cts.Token));
    }

    public void Dispose()
    {
        try { _cts?.Cancel(); } catch { }
        try { _listener?.Stop(); } catch { }
        try { _listener?.Close(); } catch { }
        _listener = null;
        _cts?.Dispose();
        _cts = null;
    }

    private async Task ListenLoop(CancellationToken token)
    {
        while (!token.IsCancellationRequested && _listener != null)
        {
            try
            {
                var context = await _listener.GetContextAsync().ConfigureAwait(false);
                _ = HandleRequest(context);
            }
            catch
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }
            }
        }
    }

    private Task<object> RunOnMainThread(Func<object> work)
    {
        if (_mainThreadControl is null || !_mainThreadControl.InvokeRequired)
        {
            return Task.FromResult(work());
        }

        var tcs = new TaskCompletionSource<object>();
        _mainThreadControl.BeginInvoke((Action)(() =>
        {
            try
            {
                tcs.TrySetResult(work());
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        }));
        return tcs.Task;
    }

    private async Task HandleRequest(HttpListenerContext context)
    {
        var watch = Stopwatch.StartNew();
        var requestPath = context.Request.Url?.AbsolutePath ?? string.Empty;
        try
        {
            AddinLog.Write($"HTTP {context.Request.HttpMethod} {requestPath} from {context.Request.RemoteEndPoint}");
            if (context.Request.HttpMethod == "GET" && requestPath == "/health")
            {
                WriteJson(context, new { ok = true, addin = "ReadBom.SwAddin", prefix = _prefix });
                AddinLog.Write($"HTTP health ok in {watch.ElapsedMilliseconds}ms");
                return;
            }

            if (context.Request.HttpMethod != "POST" || requestPath != "/command")
            {
                context.Response.StatusCode = 404;
                WriteJson(context, new { ok = false, error = "not_found" });
                AddinLog.Write($"HTTP not_found {requestPath} in {watch.ElapsedMilliseconds}ms");
                return;
            }

            var request = ReadCommand(context.Request);
            var commandName = (request.Command ?? string.Empty).Trim();
            AddinLog.Write($"Command start: {commandName}");
            var executeWatch = Stopwatch.StartNew();
            var result = await RunOnMainThread(() => ExecuteCommand(request));
            AddinLog.Write($"Command execute done: {commandName} in {executeWatch.ElapsedMilliseconds}ms");
            var writeWatch = Stopwatch.StartNew();
            WriteJson(context, new { ok = true, data = result });
            AddinLog.Write($"Command response written: {commandName} in {writeWatch.ElapsedMilliseconds}ms");
            AddinLog.Write($"Command ok: {commandName} in {watch.ElapsedMilliseconds}ms");
        }
        catch (Exception ex)
        {
            context.Response.StatusCode = 500;
            WriteJson(context, new { ok = false, error = ex.Message });
            AddinLog.Write($"HTTP failed {context.Request.HttpMethod} {requestPath} in {watch.ElapsedMilliseconds}ms: {ex}");
        }
    }

    private CommandRequest ReadCommand(HttpListenerRequest request)
    {
        using var reader = new StreamReader(request.InputStream, request.ContentEncoding ?? Encoding.UTF8);
        var body = reader.ReadToEnd();
        if (string.IsNullOrWhiteSpace(body))
        {
            return new CommandRequest();
        }

        return _json.Deserialize<CommandRequest>(body) ?? new CommandRequest();
    }

    private object ExecuteCommand(CommandRequest request)
    {
        switch ((request.Command ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "ping":
                return new { message = "pong", time = DateTime.Now };
            case "active-document":
                return GetActiveDocumentInfo();
            case "list-properties":
                return ListProperties(request);
            case "read-bom":
                return ReadBom(request);
            case "open-document":
                return OpenDocument(request);
            case "related-files":
                return GetRelatedFiles(request);
            case "save-properties-batch":
                return SavePropertiesBatch(request);
            case "calculate-blank-size":
                return CalculateBlankSizeBatch(request);
            case "set-property":
                return SetProperty(request);
            case "rebuild":
                return Rebuild();
            case "save":
                return Save();
            default:
                throw new InvalidOperationException($"未知命令: {request.Command}");
        }
    }

    private object GetActiveDocumentInfo()
    {
        var model = GetActiveModel();
        return new
        {
            title = Safe(() => model.GetTitle()),
            path = Safe(() => model.GetPathName()),
            configuration = Safe(() => model.ConfigurationManager.ActiveConfiguration.Name)
        };
    }

    private object ListProperties(CommandRequest request)
    {
        var model = GetActiveModel();
        var configName = GetConfigName(model, request.Configuration);
        var manager = model.Extension.get_CustomPropertyManager(configName);
        var values = ReadAllProperties(manager);
        return new
        {
            configuration = string.IsNullOrWhiteSpace(configName) ? "custom" : configName,
            properties = values
        };
    }

    private object SetProperty(CommandRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Name))
        {
            throw new InvalidOperationException("name 不能为空");
        }

        var model = GetActiveModel();
        var configName = GetConfigName(model, request.Configuration);
        var manager = model.Extension.get_CustomPropertyManager(configName);
        manager.Add3(request.Name, 30, request.Value ?? string.Empty, 2);
        AddinLog.Write($"SetProperty: name={request.Name}, configuration={configName}, valueLength={(request.Value ?? string.Empty).Length}");
        return new { name = request.Name, value = request.Value ?? string.Empty, configuration = configName };
    }

    private object SavePropertiesBatch(CommandRequest request)
    {
        var rows = request.SaveRows ?? new SavePropertyRow[0];
        var savedRows = 0;
        var failedRows = 0;
        var savedProperties = 0;
        var failures = new List<object>();
        var useConfiguration = string.Equals(request.PropertySourceMode, "CurrentConfiguration", StringComparison.OrdinalIgnoreCase);
        AddinLog.Write($"SavePropertiesBatch: rows={rows.Length}, sourceMode={request.PropertySourceMode}");

        for (var i = 0; i < rows.Length; i++)
        {
            var row = rows[i];
            var displayName = string.IsNullOrWhiteSpace(row.DisplayName) ? row.Path : row.DisplayName;
            try
            {
                if (string.IsNullOrWhiteSpace(row.Path))
                {
                    throw new InvalidOperationException("缺少完整路径");
                }

                var path = System.IO.Path.GetFullPath(row.Path.Trim());
                var docType = GetDocumentTypeFromPath(path);
                if (docType == 0)
                {
                    throw new InvalidOperationException("不支持的 SolidWorks 文件类型");
                }

                if (!File.Exists(path))
                {
                    throw new InvalidOperationException("文件不存在");
                }

                var model = TryGetOpenModelByPath(path);
                if (model == null)
                {
                    var errors = 0;
                    var warnings = 0;
                    try { _swApp.DocumentVisible(true, docType); } catch { }
                    model = _swApp.OpenDoc6(path, docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, string.Empty, ref errors, ref warnings) as ModelDoc2;
                    AddinLog.Write($"SavePropertiesBatch OpenDoc6: {displayName}, model={(model != null)}, errors={errors}, warnings={warnings}");
                    if (model == null)
                    {
                        throw new InvalidOperationException($"无法打开模型，错误码={errors}，警告码={warnings}");
                    }
                }

                var configName = useConfiguration ? NormalizeConfigurationName(row.Configuration) : string.Empty;
                var manager = model.Extension.get_CustomPropertyManager(configName);
                if (manager == null)
                {
                    throw new InvalidOperationException(useConfiguration
                        ? "无法获取配置属性管理器: " + configName
                        : "无法获取自定义属性管理器");
                }

                var changes = row.Changes ?? new SavePropertyChange[0];
                foreach (var change in changes)
                {
                    WriteProperty(manager, change.Name, change.Value ?? string.Empty);
                    savedProperties++;
                }

                SaveModel(model);
                savedRows++;
                AddinLog.Write($"SavePropertiesBatch row ok: {displayName}, changes={changes.Length}, index={i + 1}/{rows.Length}");
            }
            catch (Exception ex)
            {
                failedRows++;
                failures.Add(new { displayName, path = row?.Path, error = ex.Message });
                AddinLog.Write($"SavePropertiesBatch row failed: {displayName}, index={i + 1}/{rows.Length}: {ex}");
            }
        }

        AddinLog.Write($"SavePropertiesBatch done: savedRows={savedRows}, failedRows={failedRows}, savedProperties={savedProperties}");
        return new
        {
            totalRows = rows.Length,
            savedRows,
            failedRows,
            savedProperties,
            failures
        };
    }

    private object CalculateBlankSizeBatch(CommandRequest request)
    {
        var rows = request.BlankRows ?? new BlankSizeRow[0];
        var updatedRows = 0;
        var failedRows = 0;
        var results = new List<object>();
        AddinLog.Write($"CalculateBlankSizeBatch: rows={rows.Length}");

        for (var i = 0; i < rows.Length; i++)
        {
            var row = rows[i];
            var displayName = string.IsNullOrWhiteSpace(row.DisplayName) ? row.Path : row.DisplayName;
            try
            {
                if (string.IsNullOrWhiteSpace(row.Path))
                {
                    throw new InvalidOperationException("缺少完整路径");
                }

                var path = System.IO.Path.GetFullPath(row.Path.Trim());
                var docType = GetDocumentTypeFromPath(path);
                if (docType == 0)
                {
                    throw new InvalidOperationException("不支持的 SolidWorks 文件类型");
                }

                if (!File.Exists(path))
                {
                    throw new InvalidOperationException("文件不存在");
                }

                var model = TryGetOpenModelByPath(path);
                if (model == null)
                {
                    var errors = 0;
                    var warnings = 0;
                    try { _swApp.DocumentVisible(true, docType); } catch { }
                    model = _swApp.OpenDoc6(path, docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, string.Empty, ref errors, ref warnings) as ModelDoc2;
                    AddinLog.Write($"CalculateBlankSizeBatch OpenDoc6: {displayName}, model={(model != null)}, errors={errors}, warnings={warnings}");
                    if (model == null)
                    {
                        throw new InvalidOperationException($"无法打开模型，错误码={errors}，警告码={warnings}");
                    }
                }

                ActivateConfiguration(model, row.Configuration);
                var box = GetModelBoxValues(model, path);
                if (box.Count < 6)
                {
                    throw new InvalidOperationException("无法获取包围盒");
                }

                updatedRows++;
                results.Add(new { path, displayName, box, success = true, error = string.Empty });
                AddinLog.Write($"CalculateBlankSizeBatch row ok: {displayName}, box={string.Join(",", box)}, index={i + 1}/{rows.Length}");
            }
            catch (Exception ex)
            {
                failedRows++;
                results.Add(new { path = row?.Path, displayName, box = new List<double>(), success = false, error = ex.Message });
                AddinLog.Write($"CalculateBlankSizeBatch row failed: {displayName}, index={i + 1}/{rows.Length}: {ex}");
            }
        }

        AddinLog.Write($"CalculateBlankSizeBatch done: updatedRows={updatedRows}, failedRows={failedRows}");
        return new
        {
            totalRows = rows.Length,
            updatedRows,
            failedRows,
            results
        };
    }

}
