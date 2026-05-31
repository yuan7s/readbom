using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swdocumentmgr;

namespace readbom;

internal static partial class OfflineSolidWorksReader
{
    public static List<BomRow> ReadBom(
        string documentPath,
        IReadOnlyList<string> propertyNames,
        ReadOptions options,
        Action<ReadProgress>? progress = null,
        Action<string>? log = null)
    {
        var path = NormalizeExistingSolidWorksPath(documentPath);
        var application = CreateApplication();
        using var document = OpenDocument(application, path, readOnly: true);
        var activeConfiguration = GetActiveConfigurationName(document.Document);
        var rows = new List<BomRow>();

        progress?.Invoke(new ReadProgress("离线读取主文件属性", 0, 1));
        rows.Add(CreateBomRow(document.Document, path, activeConfiguration, 1, propertyNames, options.PropertySourceMode));
        progress?.Invoke(new ReadProgress("离线读取主文件属性", 1, 1));

        if (GetDocumentType(path) == SwDmDocumentType.swDmDocumentAssembly)
        {
            var seeds = new Dictionary<string, OfflineBomSeed>(StringComparer.OrdinalIgnoreCase);
            var suppressed = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            progress?.Invoke(new ReadProgress("离线读取装配体组件", 0, 1));
            TraverseAssembly(application, document.Document, activeConfiguration, options, seeds, suppressed);
            progress?.Invoke(new ReadProgress("离线读取装配体组件", seeds.Count, Math.Max(seeds.Count, 1)));

            var done = 0;
            foreach (var seed in seeds.Values.OrderBy(x => x.FileName, StringComparer.OrdinalIgnoreCase).ThenBy(x => x.Configuration, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    using var child = OpenDocument(application, seed.Path, readOnly: true);
                    rows.Add(CreateBomRow(child.Document, seed.Path, seed.Configuration, seed.Quantity, propertyNames, options.PropertySourceMode));
                }
                catch (Exception ex)
                {
                    log?.Invoke($"离线读取组件失败: {seed.Path} - {ex.Message}");
                }

                done++;
                progress?.Invoke(new ReadProgress($"离线读取组件属性: {seed.FileName}", done, Math.Max(seeds.Count, 1)));
            }

            ReportSuppressedComponents(suppressed, progress);
        }

        log?.Invoke($"离线读取完成: {rows.Count} 行，文件={path}");
        return rows;
    }

    private static SwDMApplication CreateApplication()
    {
        var candidates = SwDocumentManagerLicense.LoadCandidates();
        if (candidates.Count == 0)
        {
            throw new InvalidOperationException(
                "未找到 SolidWorks Document Manager 许可证。请设置环境变量 READBOM_SWDM_LICENSE，" +
                "或在程序目录放置 sw-document-manager-license.txt。");
        }

        var failures = new List<string>();
        foreach (var license in candidates)
        {
            try
            {
                var factory = new SwDMClassFactoryClass();
                var application = factory.GetApplication(license);
                if (application != null)
                {
                    return application;
                }

                failures.Add("GetApplication 返回空");
            }
            catch (Exception ex)
            {
                failures.Add($"0x{ex.HResult:X8} {ex.Message}");
            }
        }

        throw new InvalidOperationException(
            $"SolidWorks Document Manager 初始化失败，已尝试 {candidates.Count} 个许可证候选。最后错误: {failures.LastOrDefault() ?? "未知错误"}");
    }

    private static SwDmDocumentScope OpenDocument(SwDMApplication application, string path, bool readOnly)
    {
        var documentType = GetDocumentType(path);
        if (documentType == SwDmDocumentType.swDmDocumentUnknown)
        {
            throw new InvalidOperationException("不支持的 SolidWorks 文件类型: " + path);
        }

        ISwDMDocument? document;
        SwDmDocumentOpenError error;
        try
        {
            document = application.GetDocument(path, documentType, readOnly, out error) as ISwDMDocument;
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80040112)
        {
            throw new InvalidOperationException(
                "Document Manager 已初始化，但打开文件时授权不足。请确认 license 包含 swdocmgr_general，" +
                "如果使用多 feature 文件，请保持整串 license 内容不被拆散。",
                ex);
        }

        if (document == null || error != SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
        {
            throw new InvalidOperationException($"Document Manager 打开文件失败: {Path.GetFileName(path)}，错误={error}");
        }

        return new SwDmDocumentScope(document);
    }

    private static string NormalizeExistingSolidWorksPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("文件路径不能为空。");
        }

        var fullPath = Path.GetFullPath(path.Trim());
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("SolidWorks 文件不存在。", fullPath);
        }

        if (GetDocumentType(fullPath) == SwDmDocumentType.swDmDocumentUnknown)
        {
            throw new InvalidOperationException("请选择 .sldasm、.sldprt 或 .slddrw 文件。");
        }

        return fullPath;
    }

    private static SwDmDocumentType GetDocumentType(string path)
    {
        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".sldasm" or ".asm" or ".asmdot" => SwDmDocumentType.swDmDocumentAssembly,
            ".sldprt" or ".prt" or ".prtdot" => SwDmDocumentType.swDmDocumentPart,
            ".slddrw" or ".drw" or ".drwdot" => SwDmDocumentType.swDmDocumentDrawing,
            _ => SwDmDocumentType.swDmDocumentUnknown
        };
    }

    private static string GetDocumentTypeLabel(string path)
    {
        return GetDocumentType(path) switch
        {
            SwDmDocumentType.swDmDocumentAssembly => "装配体",
            SwDmDocumentType.swDmDocumentDrawing => "工程图",
            SwDmDocumentType.swDmDocumentPart => "零件",
            _ => "未知"
        };
    }

    private static string GetDocumentIconPath(string path)
    {
        return GetDocumentType(path) switch
        {
            SwDmDocumentType.swDmDocumentAssembly => "pack://application:,,,/Assets/assembly.png",
            SwDmDocumentType.swDmDocumentDrawing => "pack://application:,,,/Assets/drawing.png",
            _ => "pack://application:,,,/Assets/part.png"
        };
    }

    private static string GetActiveConfigurationName(ISwDMDocument document)
    {
        try
        {
            var name = document.ConfigurationManager.GetActiveConfigurationName();
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }
        }
        catch
        {
        }

        return GetConfigurationNames(document).FirstOrDefault() ?? "Default";
    }

    private static List<string> GetConfigurationNames(ISwDMDocument document)
    {
        try
        {
            return ToStringList(document.ConfigurationManager.GetConfigurationNames())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private sealed class SwDmDocumentScope(ISwDMDocument document) : IDisposable
    {
        public ISwDMDocument Document { get; } = document;

        public void Dispose()
        {
            try
            {
                Document.CloseDoc();
            }
            catch
            {
            }
        }
    }

    private sealed class OfflineBomSeed
    {
        public string Path { get; init; } = string.Empty;
        public string Configuration { get; init; } = string.Empty;
        public int Quantity { get; set; }
        public string FileName => System.IO.Path.GetFileNameWithoutExtension(Path);
    }
}
