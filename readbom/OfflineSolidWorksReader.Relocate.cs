using System.IO;
using SolidWorks.Interop.swdocumentmgr;

namespace readbom;

internal static partial class OfflineSolidWorksReader
{
    public static RelocateResult CopyExternalFilesToMainAssemblyDirectory(
        string mainAssemblyPath,
        BomRow mainAssemblyRow,
        IReadOnlyList<BomRow> rows,
        Action<ReadProgress>? progress = null,
        Action<string>? log = null)
    {
        if (rows.Count == 0)
        {
            progress?.Invoke(new ReadProgress("没有需要本地复制重装的文件", 1, 1));
            return new RelocateResult(0, 0, 0, 0);
        }

        var assemblyPath = NormalizeExistingSolidWorksPath(
            string.IsNullOrWhiteSpace(mainAssemblyPath) ? mainAssemblyRow.FullPath : mainAssemblyPath);
        if (GetDocumentType(assemblyPath) != SwDmDocumentType.swDmDocumentAssembly)
        {
            throw new InvalidOperationException("本地复制重装需要先打开装配体文件。");
        }

        var mainDirectory = Path.GetDirectoryName(assemblyPath);
        if (string.IsNullOrWhiteSpace(mainDirectory))
        {
            throw new InvalidOperationException("无法获取主装配体目录。");
        }

        var application = CreateApplication();
        using var assembly = OpenDocument(application, assemblyPath, readOnly: false);
        var externalReferences = GetExternalReferences(application, assembly.Document, mainDirectory);
        var copiedFiles = 0;
        var replacedRows = 0;
        var failedRows = 0;
        var updatedRows = new List<(BomRow Row, string TargetPath)>();
        var total = Math.Max(rows.Count, 1);

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var displayName = string.IsNullOrWhiteSpace(row.FileName) ? row.FullPath : row.FileName;
            progress?.Invoke(new ReadProgress($"本地复制重装: {displayName}", i, total));

            try
            {
                var sourcePath = NormalizeExistingSolidWorksPath(row.FullPath);
                var targetPath = ResolveValidatedCopyTarget(sourcePath, mainDirectory);
                if (!File.Exists(targetPath) || !FilesAreSame(sourcePath, targetPath))
                {
                    File.Copy(sourcePath, targetPath, overwrite: false);
                    copiedFiles++;
                }

                var referenceToReplace = ResolveReferenceForReplacement(externalReferences, sourcePath);
                if (string.IsNullOrWhiteSpace(referenceToReplace))
                {
                    referenceToReplace = sourcePath;
                }

                assembly.Document.ReplaceReference(referenceToReplace, targetPath);
                externalReferences.RemoveAll(reference => PathsEqual(reference.ResolvedPath, referenceToReplace) || PathsEqual(reference.ResolvedPath, sourcePath));
                externalReferences.Add(new ExternalReference(targetPath, targetPath));

                updatedRows.Add((row, targetPath));
                replacedRows++;
                progress?.Invoke(new ReadProgress($"本地复制重装: {displayName}", i + 1, total));
            }
            catch (Exception ex)
            {
                failedRows++;
                log?.Invoke($"本地复制重装失败: {displayName} - {ex.Message}");
                progress?.Invoke(new ReadProgress($"本地复制重装失败: {displayName}", i + 1, total));
            }
        }

        if (replacedRows > 0)
        {
            var saveError = assembly.Document.Save();
            if (saveError != SwDmDocumentSaveError.swDmDocumentSaveErrorNone)
            {
                throw new InvalidOperationException("Document Manager 保存装配体失败: " + saveError);
            }
        }

        foreach (var (row, targetPath) in updatedRows)
        {
            row.FullPath = targetPath;
            row.IsOutsideMainAssemblyDirectory = false;
        }

        return new RelocateResult(rows.Count, copiedFiles, replacedRows, failedRows);
    }

    private static List<ExternalReference> GetExternalReferences(SwDMApplication application, ISwDMDocument document, string mainDirectory)
    {
        try
        {
            var searchOptions = application.GetSearchOptionObject();
            searchOptions.SearchFilters = 27;
            searchOptions.AddSearchPath(mainDirectory);
            return ToStringList(document.GetAllExternalReferences(searchOptions))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(path => new ExternalReference(path, ResolveReferencePath(path, mainDirectory)))
                .GroupBy(reference => reference.ResolvedPath, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static string ResolveReferenceForReplacement(IReadOnlyList<ExternalReference> references, string sourcePath)
    {
        foreach (var reference in references)
        {
            if (PathsEqual(reference.ResolvedPath, sourcePath))
            {
                return reference.OriginalPath;
            }
        }

        var sourceFileName = Path.GetFileName(sourcePath);
        var fileNameMatches = references
            .Where(reference => string.Equals(Path.GetFileName(reference.ResolvedPath), sourceFileName, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return fileNameMatches.Count == 1 ? fileNameMatches[0].OriginalPath : string.Empty;
    }

    private static string ResolveReferencePath(string path, string baseDirectory)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return Path.IsPathRooted(path)
                ? Path.GetFullPath(path.Trim())
                : Path.GetFullPath(Path.Combine(baseDirectory, path.Trim()));
        }
        catch
        {
            return path.Trim();
        }
    }

    private static string ResolveValidatedCopyTarget(string sourcePath, string targetDirectory)
    {
        var fileName = Path.GetFileName(sourcePath);
        var targetPath = Path.Combine(targetDirectory, fileName);
        if (!File.Exists(targetPath) || FilesAreSame(sourcePath, targetPath))
        {
            return targetPath;
        }

        var name = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        for (var i = 1; i < 10000; i++)
        {
            var candidate = Path.Combine(targetDirectory, $"{name}_{i}{extension}");
            if (!File.Exists(candidate) || FilesAreSame(sourcePath, candidate))
            {
                return candidate;
            }
        }

        throw new IOException("无法生成不冲突的复制文件名: " + fileName);
    }

    private static bool FilesAreSame(string leftPath, string rightPath)
    {
        try
        {
            var left = new FileInfo(leftPath);
            var right = new FileInfo(rightPath);
            if (!left.Exists || !right.Exists || left.Length != right.Length)
            {
                return false;
            }

            using var leftStream = File.OpenRead(left.FullName);
            using var rightStream = File.OpenRead(right.FullName);
            var leftBuffer = new byte[1024 * 1024];
            var rightBuffer = new byte[1024 * 1024];
            while (true)
            {
                var leftRead = leftStream.Read(leftBuffer, 0, leftBuffer.Length);
                var rightRead = rightStream.Read(rightBuffer, 0, rightBuffer.Length);
                if (leftRead != rightRead)
                {
                    return false;
                }

                if (leftRead == 0)
                {
                    return true;
                }

                for (var i = 0; i < leftRead; i++)
                {
                    if (leftBuffer[i] != rightBuffer[i])
                    {
                        return false;
                    }
                }
            }
        }
        catch
        {
            return false;
        }
    }

    private static bool PathsEqual(string leftPath, string rightPath)
    {
        return string.Equals(NormalizePathForCompare(leftPath), NormalizePathForCompare(rightPath), StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizePathForCompare(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(path.Trim());
        }
        catch
        {
            return path.Trim();
        }
    }

    private sealed record ExternalReference(string OriginalPath, string ResolvedPath);
}
