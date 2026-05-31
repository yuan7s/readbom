using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;

namespace readbom;

internal static partial class SolidWorksAddinClient
{
    private static char DetectDelimitedTextSeparator(string text, string? requestedSeparator)
    {
        var candidates = new List<char>();
        if (!string.IsNullOrEmpty(requestedSeparator))
        {
            candidates.Add(requestedSeparator[0]);
        }

        foreach (var candidate in new[] { ',', '\t', ';', '|' })
        {
            if (!candidates.Contains(candidate))
            {
                candidates.Add(candidate);
            }
        }

        var sampleLines = text
            .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Take(10)
            .ToList();
        if (sampleLines.Count == 0)
        {
            return candidates[0];
        }

        return candidates
            .Select(candidate => new
            {
                Separator = candidate,
                Score = sampleLines.Sum(line => CountUnquotedSeparator(line, candidate))
            })
            .OrderByDescending(x => x.Score)
            .First().Separator;
    }

    private static int CountUnquotedSeparator(string line, char separator)
    {
        var count = 0;
        var inQuotes = false;
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"' && inQuotes && i + 1 < line.Length && line[i + 1] == '"')
            {
                i++;
                continue;
            }

            if (ch == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (!inQuotes && ch == separator)
            {
                count++;
            }
        }

        return count;
    }

    private static string DescribeDelimiter(char delimiter)
    {
        return delimiter switch
        {
            '\t' => "Tab",
            ',' => "逗号",
            ';' => "分号",
            _ => delimiter.ToString()
        };
    }

    private static void SaveLastBomCsv(string csvText, Action<string>? log)
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, "last-bom.csv");
            File.WriteAllText(path, csvText, Encoding.UTF8);
            log?.Invoke($"CSV调试文件: {path}");
        }
        catch (Exception ex)
        {
            log?.Invoke($"CSV调试文件保存失败: {ex.Message}");
        }
    }

    private static List<List<string>> ParseDelimitedText(string text, char delimiter)
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var cell = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (inQuotes)
            {
                if (ch == '"' && i + 1 < text.Length && text[i + 1] == '"')
                {
                    cell.Append('"');
                    i++;
                }
                else if (ch == '"')
                {
                    inQuotes = false;
                }
                else
                {
                    cell.Append(ch);
                }

                continue;
            }

            if (ch == '"')
            {
                inQuotes = true;
            }
            else if (ch == delimiter)
            {
                row.Add(cell.ToString());
                cell.Clear();
            }
            else if (ch == '\r' || ch == '\n')
            {
                row.Add(cell.ToString());
                cell.Clear();
                if (row.Any(value => value.Length > 0))
                {
                    rows.Add(row);
                }

                row = new List<string>();
                if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }
            }
            else
            {
                cell.Append(ch);
            }
        }

        if (cell.Length > 0 || row.Count > 0)
        {
            row.Add(cell.ToString());
            if (row.Any(value => value.Length > 0))
            {
                rows.Add(row);
            }
        }

        return rows;
    }

    private static string DecodeCsvBytes(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        }

        TryRegisterCodePages();

        return GetCsvDecodeCandidates()
            .Select(encoding =>
            {
                var text = encoding.GetString(bytes);
                return new
                {
                    Text = text,
                    Score = ScoreDecodedCsvText(text)
                };
            })
            .OrderByDescending(candidate => candidate.Score)
            .First()
            .Text;
    }

    private static void TryRegisterCodePages()
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        catch
        {
            // The provider can be registered only once. If it is unavailable, the candidate list falls back.
        }
    }

    private static IEnumerable<Encoding> GetCsvDecodeCandidates()
    {
        foreach (var codePage in new[] { 54936, 936 })
        {
            Encoding? encoding = null;
            try
            {
                encoding = Encoding.GetEncoding(codePage);
            }
            catch
            {
                // Ignore unavailable code pages.
            }

            if (encoding is not null)
            {
                yield return encoding;
            }
        }

        yield return Encoding.UTF8;
        yield return Encoding.Default;
    }

    private static int ScoreDecodedCsvText(string text)
    {
        var score = 0;
        foreach (var token in new[]
                 {
                     "\u9879\u76ee\u53f7",
                     "\u96f6\u4ef6\u53f7",
                     "\u8bf4\u660e",
                     "\u6570\u91cf",
                     "\u6750\u6599",
                     "\u7269\u6599\u7f16\u7801",
                     "\u96f6\u4ef6\u56fe\u53f7",
                     "SW\u6750\u6599"
                 })
        {
            score += CountOccurrences(text, token) * 20;
        }

        score += CountOccurrences(text, ",") / 10;
        score -= CountOccurrences(text, "\uFFFD") * 50;
        return score;
    }

    private static int CountOccurrences(string text, string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return 0;
        }

        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static int FindCsvHeaderRow(IReadOnlyList<List<string>> records, IReadOnlyList<string> propertyNames)
    {
        for (var i = 0; i < records.Count; i++)
        {
            var headers = records[i].Select(NormalizeCsvHeader).ToList();
            if (FindCsvColumn(headers, "SW\u6750\u6599", "SW-Material") >= 0
                || FindCsvColumn(headers, "\u6570\u91cf", "QTY", "QTY.", "Quantity") >= 0
                || propertyNames.Any(name => FindCsvColumn(headers, name) >= 0))
            {
                return i;
            }
        }

        return 0;
    }

    private static int FindCsvColumn(IReadOnlyList<string> headers, params string[] candidates)
    {
        for (var i = 0; i < headers.Count; i++)
        {
            foreach (var candidate in candidates)
            {
                if (string.Equals(headers[i], NormalizeCsvHeader(candidate), StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
        }

        for (var i = 0; i < headers.Count; i++)
        {
            foreach (var candidate in candidates)
            {
                var normalized = NormalizeCsvHeader(candidate);
                if (!string.IsNullOrWhiteSpace(normalized)
                    && headers[i].IndexOf(normalized, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return i;
                }
            }
        }

        return -1;
    }

    private static string NormalizeCsvHeader(string? value)
    {
        return StripCsvHtmlTags(value ?? string.Empty)
            .Trim()
            .Trim('"')
            .Replace("\r", " ")
            .Replace("\n", " ");
    }

    private static string StripCsvHtmlTags(string value)
    {
        var builder = new StringBuilder(value.Length);
        var insideTag = false;
        foreach (var ch in value)
        {
            if (ch == '<')
            {
                insideTag = true;
                continue;
            }

            if (insideTag)
            {
                if (ch == '>')
                {
                    insideTag = false;
                }

                continue;
            }

            builder.Append(ch);
        }

        return builder.ToString();
    }

    private static string GetCsvCell(IReadOnlyList<string> row, int index)
    {
        return index >= 0 && index < row.Count ? row[index].Trim() : string.Empty;
    }

    private static string PickCsvDisplayName(
        IReadOnlyList<string> record,
        IReadOnlyList<string> headers,
        int quantityIndex,
        int materialIndex,
        IEnumerable<int> propertyIndexes)
    {
        var excluded = new HashSet<int>(propertyIndexes.Where(index => index >= 0)) { quantityIndex, materialIndex };
        for (var i = 0; i < record.Count; i++)
        {
            if (excluded.Contains(i))
            {
                continue;
            }

            var header = i < headers.Count ? headers[i] : string.Empty;
            if (header.Contains("项目", StringComparison.OrdinalIgnoreCase)
                || header.Equals("ITEM NO.", StringComparison.OrdinalIgnoreCase)
                || header.Equals("ITEM", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(record[i]))
            {
                return record[i].Trim();
            }
        }

        return record.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? string.Empty;
    }

    private static bool ShouldSkipCsvBomRow(string rawFileName, IReadOnlyList<string> record, IReadOnlyList<string> headers)
    {
        if (string.IsNullOrWhiteSpace(rawFileName))
        {
            return true;
        }

        var normalized = NormalizeCsvHeader(rawFileName);
        if (headers.Any(header => string.Equals(header, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        if (normalized.Equals("总计", StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("合计", StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return record.Count(value => !string.IsNullOrWhiteSpace(value)) <= 1
               && int.TryParse(normalized, NumberStyles.Integer, CultureInfo.CurrentCulture, out _);
    }

    private static int ParseCsvQuantity(string value)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out var quantity))
        {
            return quantity;
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out var numeric)
            ? Math.Max(0, (int)Math.Round(numeric))
            : 0;
    }

    private static string NormalizeCsvFileName(string rawFileName, string fullPath)
    {
        if (!string.IsNullOrWhiteSpace(fullPath))
        {
            return Path.GetFileNameWithoutExtension(fullPath);
        }

        return Path.GetFileNameWithoutExtension(rawFileName);
    }

    private static string InferDocumentType(string fullPath, string fileName)
    {
        var extension = Path.GetExtension(!string.IsNullOrWhiteSpace(fullPath) ? fullPath : fileName);
        if (extension.Equals(".sldasm", StringComparison.OrdinalIgnoreCase)) return "装配体";
        if (extension.Equals(".sldprt", StringComparison.OrdinalIgnoreCase)) return "零件";
        return string.Empty;
    }

    private static string GetMaterialDisplay(string? documentType, string? material)
    {
        var normalized = NormalizeMaterialValue(material);
        if (!string.IsNullOrWhiteSpace(normalized)
            && !normalized.Equals("<未指定>", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("未指定", StringComparison.OrdinalIgnoreCase)
            && !normalized.Equals("未设置", StringComparison.OrdinalIgnoreCase))
        {
            return normalized;
        }

        return string.Equals(documentType, "装配体", StringComparison.OrdinalIgnoreCase) ? "无需设置" : "未设置";
    }

    private static string NormalizeMaterialValue(string? material)
    {
        var value = (material ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        if (value.StartsWith("$PRP:", StringComparison.OrdinalIgnoreCase))
        {
            value = value[5..].Trim();
            value = RemoveMaterialFormulaPropertyName(value);
        }
        else if (value.StartsWith("$PRPSHEET:", StringComparison.OrdinalIgnoreCase))
        {
            value = value[10..].Trim();
            value = RemoveMaterialFormulaPropertyName(value);
        }

        value = value.Trim('"', '\'', ' ', '\t');
        return value;
    }

    private static string RemoveMaterialFormulaPropertyName(string value)
    {
        value = value.Trim('"', '\'', ' ', '\t');
        var separatorIndex = value.IndexOfAny([' ', '\t']);
        if (separatorIndex <= 0)
        {
            return value;
        }

        var propertyName = value[..separatorIndex].Trim('"', '\'', ' ', '\t');
        var resolvedValue = value[(separatorIndex + 1)..].Trim('"', '\'', ' ', '\t');
        return propertyName.Equals("SW-Material", StringComparison.OrdinalIgnoreCase)
               || propertyName.Equals("\u6750\u8d28", StringComparison.OrdinalIgnoreCase)
               || propertyName.Equals("\u6750\u6599", StringComparison.OrdinalIgnoreCase)
            ? resolvedValue
            : value;
    }

    private static async Task<T> PostCommandAsync<T>(object request, TimeSpan timeout, Action<string>? log = null)
    {
        var serializeWatch = Stopwatch.StartNew();
        var json = JsonSerializer.Serialize(request, JsonOptions);
        log?.Invoke($"Add-in计时: 请求序列化 {serializeWatch.ElapsedMilliseconds}ms, bytes={Encoding.UTF8.GetByteCount(json)}");
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var cts = new CancellationTokenSource(timeout);
        var postWatch = Stopwatch.StartNew();
        using var response = await Client.PostAsync(new Uri(BaseUri, "command"), content, cts.Token);
        log?.Invoke($"Add-in计时: HTTP等待响应头 {postWatch.ElapsedMilliseconds}ms");
        var readBodyWatch = Stopwatch.StartNew();
        var body = await response.Content.ReadAsStringAsync();
        log?.Invoke($"Add-in计时: 读取响应正文 {readBodyWatch.ElapsedMilliseconds}ms, chars={body.Length}");
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Add-in请求失败: {response.StatusCode} {body}");
        }

        var deserializeWatch = Stopwatch.StartNew();
        var envelope = JsonSerializer.Deserialize<AddinEnvelope<T>>(body, JsonOptions)
                       ?? throw new InvalidOperationException("Add-in返回空响应");
        log?.Invoke($"Add-in计时: JSON反序列化 {deserializeWatch.ElapsedMilliseconds}ms");
        if (!envelope.Ok)
        {
            throw new InvalidOperationException(envelope.Error ?? "Add-in返回失败");
        }

        return envelope.Data!;
    }

    private static string GetDocumentIconPathFromLabel(string? label)
    {
        return label switch
        {
            "装配体" => "pack://application:,,,/Assets/assembly.png",
            "零件" => "pack://application:,,,/Assets/part.png",
            _ => string.Empty
        };
    }

    private static string GetDrawingStatus(string? status, string? path)
    {
        if (!string.IsNullOrWhiteSpace(status))
        {
            return status;
        }

        return SolidWorksReader.HasSiblingDrawing(path ?? string.Empty) ? "有工程图" : "无工程图";
    }

}
