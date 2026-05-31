using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using SldWorks;
using SwConst;

namespace ReadBom.SwAddin;

internal sealed class AddinHttpServer : IDisposable
{
    private readonly SldWorks.SldWorks _swApp;
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
                _ = Task.Run(() => HandleRequest(context), token);
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

    private void HandleRequest(HttpListenerContext context)
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
            var result = ExecuteCommand(request);
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

    private object ReadBom(CommandRequest request)
    {
        var totalWatch = Stopwatch.StartNew();
        var model = GetActiveModel();
        var mainPath = Safe(() => model.GetPathName());
        var mainConfig = Safe(() => model.ConfigurationManager.ActiveConfiguration.Name);
        AddinLog.Write($"ReadBom: mainPath={mainPath}, mainConfig={mainConfig}, propertyCount={request.PropertyNames?.Length ?? 0}, groupByConfig={request.GroupByConfig}, skipVirtual={request.SkipVirtual}");

        var assembly = model as AssemblyDoc;
        if (assembly != null)
        {
            var tableResult = ExportHiddenBomTableCsv(model, mainPath, mainConfig, request);
            AddinLog.Write($"ReadBom: hidden BOM CSV path used, bytes={tableResult.CsvByteCount}");
            AddinLog.Write($"ReadBom: output CSV totalElapsed={totalWatch.ElapsedMilliseconds}ms");
            return new { tableCsv = tableResult };
        }

        var drawingLookup = new DrawingLookup();
        var rows = new List<object>
        {
            CreateBomRow(model, mainPath, mainConfig, 1, request.PropertyNames, drawingLookup)
        };

        AddinLog.Write($"ReadBom: output rows={rows.Count}, totalElapsed={totalWatch.ElapsedMilliseconds}ms");
        return new { rows };
    }

    private object OpenDocument(CommandRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Path))
        {
            throw new InvalidOperationException("path 不能为空");
        }

        var path = System.IO.Path.GetFullPath(request.Path.Trim());
        if (!File.Exists(path))
        {
            throw new InvalidOperationException("文件不存在: " + path);
        }

        var docType = GetDocumentTypeFromPath(path);
        if (docType == 0)
        {
            throw new InvalidOperationException("不支持的 SolidWorks 文件类型: " + path);
        }

        AddinLog.Write($"OpenDocument queued: path={path}, docType={docType}");
        ActivateSolidWorksWindow();
        Task.Run(() => OpenDocumentInBackground(path, docType));
        return new
        {
            path,
            title = System.IO.Path.GetFileName(path),
            accepted = true
        };
    }

    private void OpenDocumentInBackground(string path, int docType)
    {
        var watch = Stopwatch.StartNew();
        AddinLog.Write($"OpenDocument background start: path={path}, docType={docType}");
        var model = TryGetOpenModelByPath(path);
        var wasOpen = model != null;
        var errors = 0;
        var warnings = 0;
        try
        {
            if (model == null)
            {
                try { _swApp.DocumentVisible(true, docType); } catch (Exception ex) { AddinLog.Write("OpenDocument DocumentVisible ignored: " + ex.Message); }
                model = _swApp.OpenDoc6(path, docType, 0, string.Empty, ref errors, ref warnings) as ModelDoc2;
                AddinLog.Write($"OpenDocument OpenDoc6 result: model={(model != null)}, errors={errors}, warnings={warnings}");
            }

            if (model == null)
            {
                AddinLog.Write($"OpenDocument failed: OpenDoc6 returned null, errors={errors}, warnings={warnings}, path={path}");
                return;
            }

            var title = Safe(() => model.GetTitle()) ?? System.IO.Path.GetFileName(path);
            ActivateDocument(title, path);
            ActivateSolidWorksWindow();
            AddinLog.Write($"OpenDocument background ok: title={title}, wasOpen={wasOpen}, errors={errors}, warnings={warnings}, elapsed={watch.ElapsedMilliseconds}ms");
        }
        catch (Exception ex)
        {
            AddinLog.Write($"OpenDocument background failed: path={path}, elapsed={watch.ElapsedMilliseconds}ms: {ex}");
        }
    }

    private object GetRelatedFiles(CommandRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Path))
        {
            throw new InvalidOperationException("path 不能为空");
        }

        var mainPath = System.IO.Path.GetFullPath(request.Path.Trim());
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        AddRelatedFile(files, mainPath);
        AddinLog.Write("RelatedFiles: mainPath=" + mainPath);

        object dependencies = null;
        try
        {
            dependencies = _swApp.GetDocumentDependencies2(mainPath, true, true, true);
        }
        catch (Exception ex)
        {
            AddinLog.Write("RelatedFiles GetDocumentDependencies2 failed: " + ex.Message);
        }

        foreach (var item in ToIndexedStringList(dependencies))
        {
            AddRelatedFile(files, item);
        }

        AddinLog.Write("RelatedFiles: count=" + files.Count);
        return new { mainPath, files = files.ToList() };
    }

    private static void AddRelatedFile(ISet<string> files, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var candidate = value.Trim();
        var extension = System.IO.Path.GetExtension(candidate);
        if (!extension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".slddrw", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        try
        {
            files.Add(System.IO.Path.GetFullPath(candidate));
        }
        catch
        {
            files.Add(candidate);
        }
    }

    private BomTableCsvTransfer ExportHiddenBomTableCsv(
        ModelDoc2 model,
        string mainPath,
        string mainConfig,
        CommandRequest request)
    {
        var watch = Stopwatch.StartNew();
        BomTableAnnotation bomTable = null;
        try
        {
            var propertyNames = GetDistinctPropertyNames(request.PropertyNames);
            var template = GetPreparedBomTableTemplatePath(model, propertyNames);
            var insertWatch = Stopwatch.StartNew();
            AddinLog.Write($"ReadBom timing: InsertBomTable3 start, template={template}");
            bomTable = InsertBomTable(model, template, mainConfig);
            AddinLog.Write($"ReadBom timing: InsertBomTable3 only elapsed={insertWatch.ElapsedMilliseconds}ms");
            if (bomTable == null)
            {
                throw new InvalidOperationException("InsertBomTable3 返回空");
            }

            var table = (TableAnnotation)bomTable;
            var rowCountWatch = Stopwatch.StartNew();
            var rowCount = table.RowCount;
            AddinLog.Write($"ReadBom timing: get BOM row count rows={rowCount}, elapsed={rowCountWatch.ElapsedMilliseconds}ms");
            AddinLog.Write($"ReadBom timing: prepare hidden BOM table for CSV rows={rowCount}, elapsed={watch.ElapsedMilliseconds}ms");

            var exportDirectory = Path.Combine(AddinLog.DirectoryPath, "BomExports");
            Directory.CreateDirectory(exportDirectory);
            var csvPath = Path.Combine(exportDirectory, $"readbom-{Guid.NewGuid():N}.csv");
            const string separator = ",";
            var exportWatch = Stopwatch.StartNew();
            var saved = table.SaveAsText2(csvPath, separator, false);
            AddinLog.Write($"ReadBom timing: SaveAsText2 CSV saved={saved}, elapsed={exportWatch.ElapsedMilliseconds}ms, path={csvPath}");
            if (!saved || !File.Exists(csvPath))
            {
                throw new InvalidOperationException("导出 BOM CSV 失败: SaveAsText2 未生成文件");
            }

            var readCsvWatch = Stopwatch.StartNew();
            var csvBytes = File.ReadAllBytes(csvPath);
            AddinLog.Write($"ReadBom timing: read CSV file bytes={csvBytes.Length}, elapsed={readCsvWatch.ElapsedMilliseconds}ms");
            TryDeleteFile(csvPath);

            return new BomTableCsvTransfer
            {
                CsvBase64 = Convert.ToBase64String(csvBytes),
                CsvByteCount = csvBytes.Length,
                Separator = separator,
                PropertyNames = propertyNames.ToList(),
                MainPath = mainPath ?? string.Empty,
                MainConfiguration = string.IsNullOrWhiteSpace(mainConfig) ? "Default" : mainConfig,
                RowCount = rowCount
            };
        }
        finally
        {
            var deleteWatch = Stopwatch.StartNew();
            DeleteBomTableFeature(model, bomTable);
            AddinLog.Write($"ReadBom timing: delete hidden BOM elapsed={deleteWatch.ElapsedMilliseconds}ms");
        }
    }

    private static string[] GetDistinctPropertyNames(string[] propertyNames)
    {
        return (propertyNames ?? Array.Empty<string>())
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name => name.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static Dictionary<string, int> AddBomPropertyColumns(TableAnnotation table, BomTableAnnotation bomTable, string[] propertyNames)
    {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var propertyName in propertyNames ?? Array.Empty<string>())
        {
            if (string.IsNullOrWhiteSpace(propertyName) || result.ContainsKey(propertyName))
            {
                continue;
            }

            var watch = Stopwatch.StartNew();
            var columnIndex = AddBomCustomPropertyColumn(table, bomTable, propertyName, propertyName);
            AddinLog.Write($"ReadBom timing: add TXT property column name={propertyName}, index={columnIndex}, elapsed={watch.ElapsedMilliseconds}ms");
            if (columnIndex >= 0)
            {
                result[propertyName] = columnIndex;
            }
        }

        return result;
    }

    private static int AddBomCustomPropertyColumn(TableAnnotation table, BomTableAnnotation bomTable, string title, string customPropertyName)
    {
        try
        {
            var totalWatch = Stopwatch.StartNew();
            var beforeCount = table.ColumnCount;
            var insertAt = Math.Max(beforeCount - 1, 0);
            var insertWatch = Stopwatch.StartNew();
            var inserted = table.InsertColumn2(
                (int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last,
                insertAt,
                title,
                (int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
            var insertElapsed = insertWatch.ElapsedMilliseconds;
            if (!inserted)
            {
                AddinLog.Write($"ReadBom table column insert failed: {title}/{customPropertyName}");
                return -1;
            }

            var columnIndex = table.ColumnCount - 1;
            var setType2Watch = Stopwatch.StartNew();
            try { table.SetColumnType2(columnIndex, (int)swTableColumnTypes_e.swBomTableColumnType_CustomProperty, false); } catch (Exception ex) { AddinLog.Write($"ReadBom table SetColumnType2 ignored: {title}/{customPropertyName}: {ex.Message}"); }
            var setType2Elapsed = setType2Watch.ElapsedMilliseconds;

            var setType3Watch = Stopwatch.StartNew();
            try { table.SetColumnType3(columnIndex, (int)swTableColumnTypes_e.swBomTableColumnType_CustomProperty, false, customPropertyName); } catch (Exception ex) { AddinLog.Write($"ReadBom table SetColumnType3 ignored: {title}/{customPropertyName}: {ex.Message}"); }
            var setType3Elapsed = setType3Watch.ElapsedMilliseconds;

            var setCustomWatch = Stopwatch.StartNew();
            try { bomTable.SetColumnCustomProperty(columnIndex, customPropertyName); } catch (Exception ex) { AddinLog.Write($"ReadBom table SetColumnCustomProperty ignored: {title}/{customPropertyName}: {ex.Message}"); }
            var setCustomElapsed = setCustomWatch.ElapsedMilliseconds;

            var setTitleWatch = Stopwatch.StartNew();
            try { table.SetColumnTitle2(columnIndex, title, false); } catch (Exception ex) { AddinLog.Write($"ReadBom table SetColumnTitle2 ignored: {title}/{customPropertyName}: {ex.Message}"); }
            var setTitleElapsed = setTitleWatch.ElapsedMilliseconds;

            AddinLog.Write($"ReadBom timing: column detail title={title}, property={customPropertyName}, index={columnIndex}, insert={insertElapsed}ms, setType2={setType2Elapsed}ms, setType3={setType3Elapsed}ms, setCustom={setCustomElapsed}ms, setTitle={setTitleElapsed}ms, total={totalWatch.ElapsedMilliseconds}ms");
            return columnIndex;
        }
        catch (Exception ex)
        {
            AddinLog.Write($"ReadBom table column add failed: {title}/{customPropertyName}: {ex.Message}");
            return -1;
        }
    }

    private string GetPreparedBomTableTemplatePath(ModelDoc2 currentModel, string[] propertyNames)
    {
        var baseTemplate = GetBomTableTemplatePath();
        var signature = BuildPreparedBomTemplateSignature(baseTemplate, propertyNames);
        var targetTemplate = GetPreparedBomTableTemplateFile(signature);
        var signaturePath = targetTemplate + ".signature.txt";
        if (File.Exists(targetTemplate)
            && File.Exists(signaturePath)
            && string.Equals(File.ReadAllText(signaturePath, Encoding.UTF8), signature, StringComparison.Ordinal))
        {
            AddinLog.Write($"ReadBom timing: prepared BOM template cache hit, template={targetTemplate}");
            return targetTemplate;
        }

        var watch = Stopwatch.StartNew();
        ModelDoc2 templateModel = null;
        BomTableAnnotation templateBom = null;
        var currentTitle = Safe(() => currentModel.GetTitle());
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(targetTemplate));
            if (File.Exists(targetTemplate))
            {
                File.Delete(targetTemplate);
            }

            AddinLog.Write($"ReadBom timing: prepared BOM template cache miss, creating={targetTemplate}");
            templateModel = CreateTemporaryTemplateAssembly();
            var templateConfig = Safe(() => templateModel.ConfigurationManager.ActiveConfiguration.Name);
            AddinLog.Write($"ReadBom timing: temporary template assembly title={Safe(() => templateModel.GetTitle())}, config={templateConfig}");

            var insertWatch = Stopwatch.StartNew();
            templateBom = InsertBomTable(templateModel, baseTemplate, templateConfig);
            AddinLog.Write($"ReadBom timing: template InsertBomTable3 on empty assembly elapsed={insertWatch.ElapsedMilliseconds}ms");
            if (templateBom == null)
            {
                throw new InvalidOperationException("创建专用 BOM 模板失败: InsertBomTable3 返回空");
            }

            var table = (TableAnnotation)templateBom;
            var materialColumnWatch = Stopwatch.StartNew();
            var materialColumnIndex = AddBomCustomPropertyColumn(table, templateBom, "SW材料", "SW-Material");
            AddinLog.Write($"ReadBom timing: template add SW材料 column index={materialColumnIndex}, elapsed={materialColumnWatch.ElapsedMilliseconds}ms");
            if (materialColumnIndex < 0)
            {
                throw new InvalidOperationException("创建专用 BOM 模板失败: 无法添加 SW材料 列");
            }

            var propertyColumnsWatch = Stopwatch.StartNew();
            var propertyColumnIndexes = AddBomPropertyColumns(table, templateBom, propertyNames);
            AddinLog.Write($"ReadBom timing: template add TXT property columns requested={propertyNames.Length}, added={propertyColumnIndexes.Count}, elapsed={propertyColumnsWatch.ElapsedMilliseconds}ms");
            if (propertyColumnIndexes.Count != propertyNames.Length)
            {
                var missing = propertyNames.Where(name => !propertyColumnIndexes.ContainsKey(name));
                throw new InvalidOperationException("创建专用 BOM 模板失败: 未能添加属性列 " + string.Join(", ", missing));
            }

            var saveWatch = Stopwatch.StartNew();
            var saved = table.SaveAsTemplate(targetTemplate);
            AddinLog.Write($"ReadBom timing: SaveAsTemplate saved={saved}, elapsed={saveWatch.ElapsedMilliseconds}ms, template={targetTemplate}");
            if (!saved || !File.Exists(targetTemplate))
            {
                throw new InvalidOperationException("创建专用 BOM 模板失败: SaveAsTemplate 未生成文件");
            }

            File.WriteAllText(signaturePath, signature, Encoding.UTF8);
            AddinLog.Write($"ReadBom timing: prepared BOM template created elapsed={watch.ElapsedMilliseconds}ms");
            return targetTemplate;
        }
        finally
        {
            if (templateModel != null)
            {
                DeleteBomTableFeature(templateModel, templateBom);
                CloseTemporaryTemplateAssembly(templateModel, currentTitle);
            }
        }
    }

    private ModelDoc2 CreateTemporaryTemplateAssembly()
    {
        var watch = Stopwatch.StartNew();
        var assemblyTemplate = Safe(() => _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly));
        object newDoc = null;
        if (!string.IsNullOrWhiteSpace(assemblyTemplate) && File.Exists(assemblyTemplate))
        {
            newDoc = _swApp.NewDocument(assemblyTemplate, 0, 0, 0);
            AddinLog.Write($"ReadBom timing: NewDocument empty assembly elapsed={watch.ElapsedMilliseconds}ms, template={assemblyTemplate}");
        }

        if (newDoc == null)
        {
            newDoc = _swApp.NewAssembly();
            AddinLog.Write($"ReadBom timing: NewAssembly empty assembly elapsed={watch.ElapsedMilliseconds}ms");
        }

        var model = newDoc as ModelDoc2 ?? _swApp.ActiveDoc as ModelDoc2;
        if (model == null)
        {
            throw new InvalidOperationException("创建专用 BOM 模板失败: 无法创建临时空装配体");
        }

        return model;
    }

    private void CloseTemporaryTemplateAssembly(ModelDoc2 templateModel, string restoreTitle)
    {
        var title = Safe(() => templateModel.GetTitle());
        try
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                _swApp.CloseDoc(title);
                AddinLog.Write($"ReadBom timing: closed temporary template assembly title={title}");
            }
        }
        catch (Exception ex)
        {
            AddinLog.Write($"ReadBom: failed to close temporary template assembly title={title}: {ex.Message}");
        }

        try
        {
            if (!string.IsNullOrWhiteSpace(restoreTitle))
            {
                int errors = 0;
                _swApp.ActivateDoc3(restoreTitle, false, 0, ref errors);
                AddinLog.Write($"ReadBom timing: restored active document title={restoreTitle}, errors={errors}");
            }
        }
        catch (Exception ex)
        {
            AddinLog.Write($"ReadBom: failed to restore active document title={restoreTitle}: {ex.Message}");
        }
    }

    private static BomTableAnnotation InsertBomTable(ModelDoc2 model, string template, string mainConfig)
    {
        return model.Extension.InsertBomTable3(
            template,
            0,
            0,
            (int)swBomType_e.swBomType_Indented,
            mainConfig,
            true,
            (int)swNumberingType_e.swNumberingType_Detailed,
            false);
    }

    private static string BuildPreparedBomTemplateSignature(string baseTemplate, string[] propertyNames)
    {
        var baseTemplateStamp = File.Exists(baseTemplate)
            ? File.GetLastWriteTimeUtc(baseTemplate).Ticks.ToString()
            : string.Empty;
        return "readbom-bom-template-v2-empty-assembly"
               + "\nbase=" + baseTemplate
               + "\nbaseStamp=" + baseTemplateStamp
               + "\nSW材料=SW-Material"
               + "\n" + string.Join("\n", propertyNames ?? Array.Empty<string>());
    }

    private static string GetPreparedBomTableTemplateFile(string signature)
    {
        using (var sha = SHA256.Create())
        {
            var hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(signature));
            var hash = BitConverter.ToString(hashBytes).Replace("-", string.Empty).Substring(0, 16).ToLowerInvariant();
            return Path.Combine(AddinLog.DirectoryPath, "BomTemplates", $"readbom-{hash}.sldbomtbt");
        }
    }

    private List<object> ReadRowsFromComponents(AssemblyDoc assembly, CommandRequest request, DrawingLookup drawingLookup)
    {
        var stepWatch = Stopwatch.StartNew();
        var componentsObj = assembly.GetComponents(false) as Array;
        AddinLog.Write($"ReadBom timing: GetComponents raw={componentsObj?.Length ?? 0}, elapsed={stepWatch.ElapsedMilliseconds}ms");
        var seeds = new Dictionary<string, BomSeed>(StringComparer.OrdinalIgnoreCase);
        var suppressedCount = 0;
        var virtualCount = 0;
        var emptyPathCount = 0;
        var errorCount = 0;
        if (componentsObj != null)
        {
            stepWatch.Restart();
            foreach (var item in componentsObj)
            {
                try
                {
                    var component = item as Component2;
                    if (component == null) continue;
                    if (IsSuppressed(component))
                    {
                        suppressedCount++;
                        continue;
                    }

                    if (request.SkipVirtual && IsVirtual(component))
                    {
                        virtualCount++;
                        continue;
                    }

                    var path = component.GetPathName();
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        emptyPathCount++;
                        continue;
                    }

                    var config = Safe(() => component.ReferencedConfiguration);
                    var key = request.GroupByConfig ? path + "|" + config : path;
                    if (!seeds.TryGetValue(key, out var seed))
                    {
                        var getModelWatch = Stopwatch.StartNew();
                        var componentModel = component.GetModelDoc2() as ModelDoc2;
                        if (getModelWatch.ElapsedMilliseconds > 250)
                        {
                            AddinLog.Write($"ReadBom slow: GetModelDoc2 {Path.GetFileName(path)} {getModelWatch.ElapsedMilliseconds}ms");
                        }

                        seed = new BomSeed
                        {
                            Path = path,
                            Config = config,
                            Model = componentModel,
                            Quantity = 0
                        };
                        seeds[key] = seed;
                    }

                    seed.Quantity++;
                }
                catch (Exception ex)
                {
                    errorCount++;
                    AddinLog.Write($"ReadBom component skipped: {ex.Message}");
                }
            }

            AddinLog.Write($"ReadBom timing: group components rows={seeds.Count}, suppressed={suppressedCount}, virtual={virtualCount}, emptyPath={emptyPathCount}, errors={errorCount}, elapsed={stepWatch.ElapsedMilliseconds}ms");
        }

        stepWatch.Restart();
        var rows = seeds.Values
            .Select(seed => CreateBomRow(seed.Model, seed.Path, seed.Config, seed.Quantity, request.PropertyNames, drawingLookup))
            .OrderBy(row => GetAnonymousValue(row, "fileName"), StringComparer.OrdinalIgnoreCase)
            .ToList();
        AddinLog.Write($"ReadBom timing: create rows={seeds.Count}, elapsed={stepWatch.ElapsedMilliseconds}ms");
        return rows;
    }

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
