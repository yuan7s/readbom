using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;

namespace readbom;

internal static partial class SolidWorksAddinClient
{
    private static readonly HttpClient Client = new() { Timeout = Timeout.InfiniteTimeSpan };
    private static readonly Uri BaseUri = new("http://127.0.0.1:32127/");
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = true
    };

    public static async Task<bool> IsAvailableAsync()
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            using var response = await Client.GetAsync(new Uri(BaseUri, "health"), cts.Token);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public static async Task<AddinActiveDocumentInfo> GetActiveDocumentInfoAsync(Action<string>? log = null)
    {
        return await PostCommandAsync<AddinActiveDocumentInfo>(
            new { Command = "active-document" },
            TimeSpan.FromSeconds(10),
            log);
    }

    public static async Task<AddinOpenDocumentResult> OpenDocumentAsync(string path, Action<string>? log = null)
    {
        return await PostCommandAsync<AddinOpenDocumentResult>(
            new { Command = "open-document", Path = path },
            TimeSpan.FromMinutes(2),
            log);
    }

    public static async Task<List<string>> GetRelatedFilesAsync(string mainPath, Action<string>? log = null)
    {
        var response = await PostCommandAsync<AddinRelatedFilesResult>(
            new { Command = "related-files", Path = mainPath },
            TimeSpan.FromMinutes(2),
            log);
        return response.Files ?? [];
    }

    public static async Task<SaveResult> SavePropertiesBatchAsync(
        IReadOnlyList<AddinSavePropertyRow> rows,
        PropertySourceMode sourceMode,
        Action<ReadProgress>? progress = null,
        Action<string>? log = null)
    {
        var total = Math.Max(rows.Count, 1);
        progress?.Invoke(new ReadProgress("保存TXT属性到SW(Add-in)", 0, total));
        var response = await PostCommandAsync<AddinSavePropertiesResult>(
            new
            {
                Command = "save-properties-batch",
                PropertySourceMode = sourceMode.ToString(),
                SaveRows = rows
            },
            TimeSpan.FromMinutes(10),
            log);
        progress?.Invoke(new ReadProgress("保存TXT属性到SW(Add-in)", rows.Count, total));
        return new SaveResult(response.TotalRows, response.SavedRows, response.FailedRows, response.SavedProperties);
    }

    public static async Task<AddinBoxBatchResult> GetBoxBatchAsync(
        IReadOnlyList<AddinBlankSizeRow> rows,
        Action<ReadProgress>? progress = null,
        Action<string>? log = null)
    {
        var total = Math.Max(rows.Count, 1);
        progress?.Invoke(new ReadProgress("获取包围盒(Add-in)", 0, total));
        var response = await PostCommandAsync<AddinBoxBatchResult>(
            new
            {
                Command = "calculate-blank-size",
                BlankRows = rows
            },
            TimeSpan.FromMinutes(10),
            log);
        progress?.Invoke(new ReadProgress("获取包围盒(Add-in)", rows.Count, total));
        response.Results ??= [];
        return response;
    }

    public static async Task<List<BomRow>> ReadBomAsync(
        IReadOnlyList<string> propertyNames,
        ReadOptions options,
        Action<ReadProgress>? progress = null,
        Action<string>? log = null)
    {
        progress?.Invoke(new ReadProgress("Add-in读取BOM", 0, 1));
        var request = new
        {
            Command = "read-bom",
            PropertyNames = propertyNames.ToArray(),
            GroupByConfig = options.GroupByConfig,
            SkipVirtual = options.SkipVirtual
        };
        var requestWatch = Stopwatch.StartNew();
        var response = await PostCommandAsync<ReadBomResponse>(request, TimeSpan.FromMinutes(10), log);
        log?.Invoke($"Add-in计时: HTTP请求到响应对象 {requestWatch.ElapsedMilliseconds}ms");

        var convertWatch = Stopwatch.StartNew();
        var rows = new List<BomRow>();
        if (!string.IsNullOrWhiteSpace(response.TableCsv?.CsvBase64))
        {
            rows.AddRange(CreateBomRowsFromCsv(response.TableCsv, propertyNames, log));
        }
        else if (response.Table?.Rows is { Count: > 0 } tableRows)
        {
            var tablePropertyNames = response.Table.PropertyNames ?? propertyNames.ToList();
            foreach (var row in tableRows)
            {
                rows.Add(CreateBomRowFromAddin(row, tablePropertyNames));
            }
        }
        else
        {
            foreach (var row in response.Rows ?? [])
            {
                rows.Add(CreateBomRowFromAddin(row, propertyNames));
            }
        }

        log?.Invoke($"Add-in计时: 主程序转换BomRow {rows.Count} 行 {convertWatch.ElapsedMilliseconds}ms");
        progress?.Invoke(new ReadProgress("Add-in读取BOM", 1, 1));
        return rows;
    }

    private static BomRow CreateBomRowFromAddin(AddinBomRow row, IReadOnlyList<string> fallbackPropertyNames)
    {
        var properties = new Dictionary<string, string>(
            row.Properties ?? new Dictionary<string, string>(),
            StringComparer.OrdinalIgnoreCase);
        var available = row.AvailablePropertyNames is { Count: > 0 }
            ? row.AvailablePropertyNames
            : fallbackPropertyNames
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        foreach (var name in available)
        {
            if (!properties.ContainsKey(name))
            {
                properties[name] = string.Empty;
            }
        }

        return new BomRow
        {
            DocumentType = row.DocumentType ?? string.Empty,
            DocumentIconPath = GetDocumentIconPathFromLabel(row.DocumentType),
            DrawingStatus = row.DrawingStatus ?? string.Empty,
            DrawingIconPath = string.Empty,
            FileName = row.FileName ?? string.Empty,
            Configuration = string.IsNullOrWhiteSpace(row.Configuration) ? "Default" : row.Configuration,
            Quantity = row.Quantity,
            Material = GetMaterialDisplay(row.DocumentType, row.Material),
            Properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
            OriginalProperties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
            AvailablePropertyNames = available,
            FullPath = row.FullPath ?? string.Empty
        };
    }

    private static List<BomRow> CreateBomRowsFromCsv(AddinBomCsvTable table, IReadOnlyList<string> fallbackPropertyNames, Action<string>? log)
    {
        var decodeWatch = Stopwatch.StartNew();
        var csvBytes = Convert.FromBase64String(table.CsvBase64 ?? string.Empty);
        var csvText = DecodeCsvBytes(csvBytes);
        log?.Invoke($"Add-in计时: CSV字节解码 {csvBytes.Length} bytes -> {csvText.Length} chars {decodeWatch.ElapsedMilliseconds}ms");
        SaveLastBomCsv(csvText, log);

        var parseWatch = Stopwatch.StartNew();
        var delimiter = DetectDelimitedTextSeparator(csvText, table.Separator);
        var records = ParseDelimitedText(csvText, delimiter);
        log?.Invoke($"Add-in计时: CSV解析 {records.Count} 行，分隔符 [{DescribeDelimiter(delimiter)}]，{parseWatch.ElapsedMilliseconds}ms");
        if (records.Count == 0)
        {
            return [];
        }

        var propertyNames = table.PropertyNames is { Count: > 0 }
            ? table.PropertyNames
            : fallbackPropertyNames
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        var headerIndex = FindCsvHeaderRow(records, propertyNames);
        var headers = records[headerIndex].Select(NormalizeCsvHeader).ToList();
        var quantityIndex = FindCsvColumn(headers, "\u6570\u91cf", "QTY", "QTY.", "Quantity");
        var fileNameIndex = FindCsvColumn(headers, "\u96f6\u4ef6\u53f7", "\u96f6\u4ef6\u7f16\u53f7", "Part Number", "PART NUMBER", "\u6587\u4ef6\u540d", "File Name", "\u540d\u79f0", "\u4ee3\u53f7", "PART NO.");
        var configurationIndex = FindCsvColumn(headers, "\u914d\u7f6e", "Configuration", "Config");
        var fullPathIndex = FindCsvColumn(headers, "\u5b8c\u6574\u8def\u5f84", "FullPath", "Path");
        var materialIndex = FindCsvColumn(headers, "SW\u6750\u6599", "SW-Material");
        var propertyIndexes = propertyNames.ToDictionary(name => name, name => FindCsvColumn(headers, name), StringComparer.OrdinalIgnoreCase);
        log?.Invoke($"CSV表头: 行 {headerIndex + 1}，列数 {headers.Count}，{string.Join(" | ", headers)}");
        log?.Invoke($"CSV列索引: 文件名={fileNameIndex}, 数量={quantityIndex}, 配置={configurationIndex}, SW材料={materialIndex}");
        log?.Invoke($"CSV行样例: {string.Join("; ", records.Skip(headerIndex + 1).Take(3).Select(row => row.Count + "列"))}");

        var rows = new List<BomRow>();
        if (!string.IsNullOrWhiteSpace(table.MainPath))
        {
            var mainProperties = CreateEmptyProperties(propertyNames);
            rows.Add(new BomRow
            {
                DocumentType = "装配体",
                DocumentIconPath = GetDocumentIconPathFromLabel("装配体"),
                DrawingStatus = string.Empty,
                DrawingIconPath = string.Empty,
                FileName = Path.GetFileNameWithoutExtension(table.MainPath),
                Configuration = string.IsNullOrWhiteSpace(table.MainConfiguration) ? "Default" : table.MainConfiguration,
                Quantity = 1,
                Material = "无需设置",
                Properties = mainProperties,
                OriginalProperties = new Dictionary<string, string>(mainProperties, StringComparer.OrdinalIgnoreCase),
                AvailablePropertyNames = propertyNames.ToList(),
                FullPath = table.MainPath
            });
        }

        for (var rowIndex = headerIndex + 1; rowIndex < records.Count; rowIndex++)
        {
            var record = records[rowIndex];
            if (record.All(string.IsNullOrWhiteSpace))
            {
                continue;
            }

            var fullPath = GetCsvCell(record, fullPathIndex);
            var rawFileName = GetCsvCell(record, fileNameIndex);
            if (string.IsNullOrWhiteSpace(rawFileName))
            {
                rawFileName = PickCsvDisplayName(record, headers, quantityIndex, materialIndex, propertyIndexes.Values);
            }

            if (ShouldSkipCsvBomRow(rawFileName, record, headers))
            {
                continue;
            }

            var documentType = InferDocumentType(fullPath, rawFileName);
            var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var propertyName in propertyNames)
            {
                properties[propertyName] = propertyIndexes.TryGetValue(propertyName, out var index)
                    ? GetCsvCell(record, index)
                    : string.Empty;
            }

            rows.Add(new BomRow
            {
                DocumentType = documentType,
                DocumentIconPath = GetDocumentIconPathFromLabel(documentType),
                DrawingStatus = string.Empty,
                DrawingIconPath = string.Empty,
                FileName = NormalizeCsvFileName(rawFileName, fullPath),
                Configuration = string.IsNullOrWhiteSpace(GetCsvCell(record, configurationIndex)) ? "Default" : GetCsvCell(record, configurationIndex),
                Quantity = ParseCsvQuantity(GetCsvCell(record, quantityIndex)),
                Material = GetMaterialDisplay(documentType, GetCsvCell(record, materialIndex)),
                Properties = properties,
                OriginalProperties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
                AvailablePropertyNames = propertyNames.ToList(),
                FullPath = fullPath
            });
        }

        log?.Invoke($"Add-in计时: CSV转BomRow {rows.Count} 行，headerRow={headerIndex + 1}");
        return rows;
    }

    private static Dictionary<string, string> CreateEmptyProperties(IReadOnlyList<string> propertyNames)
    {
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var propertyName in propertyNames)
        {
            properties[propertyName] = string.Empty;
        }

        return properties;
    }

}
