using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using SldWorks;
using SwConst;

namespace ReadBom.SwAddin;

internal sealed partial class AddinHttpServer
{
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

}
