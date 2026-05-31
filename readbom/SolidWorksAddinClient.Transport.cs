namespace readbom;

internal static partial class SolidWorksAddinClient
{
    private static bool HasDrawing(string? status, string? path)
    {
        if (!string.IsNullOrWhiteSpace(status))
        {
            return status.Contains("有", StringComparison.OrdinalIgnoreCase);
        }

        return SolidWorksReader.HasSiblingDrawing(path ?? string.Empty);
    }

    private sealed class AddinEnvelope<T>
    {
        public bool Ok { get; set; }
        public T? Data { get; set; }
        public string? Error { get; set; }
    }

    public sealed class AddinActiveDocumentInfo
    {
        public string? Title { get; set; }
        public string? Path { get; set; }
        public string? Configuration { get; set; }
    }

    public sealed class AddinOpenDocumentResult
    {
        public string? Path { get; set; }
        public string? Title { get; set; }
        public bool Accepted { get; set; }
        public bool WasOpen { get; set; }
        public int Errors { get; set; }
        public int Warnings { get; set; }
    }

    private sealed class AddinRelatedFilesResult
    {
        public string? MainPath { get; set; }
        public List<string>? Files { get; set; }
    }

    private sealed class AddinSavePropertiesResult
    {
        public int TotalRows { get; set; }
        public int SavedRows { get; set; }
        public int FailedRows { get; set; }
        public int SavedProperties { get; set; }
    }

    public sealed class AddinBoxBatchResult
    {
        public int TotalRows { get; set; }
        public int UpdatedRows { get; set; }
        public int FailedRows { get; set; }
        public List<AddinBoxResult> Results { get; set; } = [];
    }

    public sealed class AddinBoxResult
    {
        public string? Path { get; set; }
        public string? DisplayName { get; set; }
        public List<double> Box { get; set; } = [];
        public bool Success { get; set; }
        public string? Error { get; set; }
    }

    private sealed class ReadBomResponse
    {
        public List<AddinBomRow>? Rows { get; set; }
        public AddinBomTable? Table { get; set; }
        public AddinBomCsvTable? TableCsv { get; set; }
    }

    private sealed class AddinBomTable
    {
        public List<string>? PropertyNames { get; set; }
        public List<AddinBomRow>? Rows { get; set; }
    }

    private sealed class AddinBomCsvTable
    {
        public string? CsvBase64 { get; set; }
        public int CsvByteCount { get; set; }
        public string? Separator { get; set; }
        public List<string>? PropertyNames { get; set; }
        public string? MainPath { get; set; }
        public string? MainConfiguration { get; set; }
        public int RowCount { get; set; }
    }

    private sealed class AddinBomRow
    {
        public string? DocumentType { get; set; }
        public string? DrawingStatus { get; set; }
        public string? FileName { get; set; }
        public string? Configuration { get; set; }
        public int Quantity { get; set; }
        public string? Material { get; set; }
        public string? FullPath { get; set; }
        public Dictionary<string, string>? Properties { get; set; }
        public List<string>? AvailablePropertyNames { get; set; }
    }
}
