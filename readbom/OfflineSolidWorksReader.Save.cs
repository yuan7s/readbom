using SolidWorks.Interop.swdocumentmgr;

namespace readbom;

internal static partial class OfflineSolidWorksReader
{
    public static SaveResult SavePropertiesBatch(
        IReadOnlyList<AddinSavePropertyRow> rows,
        PropertySourceMode sourceMode,
        Action<ReadProgress>? progress = null,
        Action<string>? log = null)
    {
        var application = CreateApplication();
        var total = Math.Max(rows.Count, 1);
        var savedRows = 0;
        var failedRows = 0;
        var savedProperties = 0;

        if (rows.Count == 0)
        {
            progress?.Invoke(new ReadProgress("没有需要保存的修改项", 1, 1));
            return new SaveResult(0, 0, 0, 0);
        }

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var displayName = string.IsNullOrWhiteSpace(row.DisplayName) ? row.Path : row.DisplayName;
            progress?.Invoke(new ReadProgress($"离线保存属性: {displayName}", i, total));

            try
            {
                using var document = OpenDocument(application, NormalizeExistingSolidWorksPath(row.Path), readOnly: false);
                foreach (var change in row.Changes.Where(x => !string.IsNullOrWhiteSpace(x.Name)))
                {
                    if (sourceMode == PropertySourceMode.CurrentConfiguration)
                    {
                        var configuration = ResolveConfiguration(document.Document, row.Configuration)
                                            ?? throw new InvalidOperationException("无法获取配置属性: " + row.Configuration);
                        WriteConfigurationProperty(configuration, change.Name, change.Value);
                    }
                    else
                    {
                        WriteDocumentProperty(document.Document, change.Name, change.Value);
                    }

                    savedProperties++;
                }

                var saveError = document.Document.Save();
                if (saveError != SwDmDocumentSaveError.swDmDocumentSaveErrorNone)
                {
                    throw new InvalidOperationException("Document Manager 保存失败: " + saveError);
                }

                savedRows++;
                progress?.Invoke(new ReadProgress($"离线保存属性: {displayName}", i + 1, total));
            }
            catch (Exception ex)
            {
                failedRows++;
                log?.Invoke($"离线保存失败: {displayName} - {ex.Message}");
                progress?.Invoke(new ReadProgress($"离线保存失败: {displayName}", i + 1, total));
            }
        }

        return new SaveResult(rows.Count, savedRows, failedRows, savedProperties);
    }

    private static void WriteDocumentProperty(ISwDMDocument document, string name, string value)
    {
        try
        {
            document.SetCustomProperty(name, value ?? string.Empty);
            return;
        }
        catch
        {
        }

        try
        {
            if (document.AddCustomProperty(name, SwDmCustomInfoType.swDmCustomInfoText, value ?? string.Empty))
            {
                return;
            }
        }
        catch
        {
        }

        try
        {
            document.DeleteCustomProperty(name);
            if (document.AddCustomProperty(name, SwDmCustomInfoType.swDmCustomInfoText, value ?? string.Empty))
            {
                return;
            }
        }
        catch
        {
        }

        throw new InvalidOperationException("属性写入失败: " + name);
    }

    private static void WriteConfigurationProperty(ISwDMConfiguration configuration, string name, string value)
    {
        try
        {
            configuration.SetCustomProperty(name, value ?? string.Empty);
            return;
        }
        catch
        {
        }

        try
        {
            if (configuration.AddCustomProperty(name, SwDmCustomInfoType.swDmCustomInfoText, value ?? string.Empty))
            {
                return;
            }
        }
        catch
        {
        }

        try
        {
            configuration.DeleteCustomProperty(name);
            if (configuration.AddCustomProperty(name, SwDmCustomInfoType.swDmCustomInfoText, value ?? string.Empty))
            {
                return;
            }
        }
        catch
        {
        }

        throw new InvalidOperationException("配置属性写入失败: " + name);
    }
}
