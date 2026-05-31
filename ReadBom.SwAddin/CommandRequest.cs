namespace ReadBom.SwAddin;

public sealed class CommandRequest
{
    public string Command { get; set; }
    public string Configuration { get; set; }
    public string Name { get; set; }
    public string Value { get; set; }
    public string Path { get; set; }
    public string[] PropertyNames { get; set; }
    public string PropertySourceMode { get; set; }
    public SavePropertyRow[] SaveRows { get; set; }
    public BlankSizeRow[] BlankRows { get; set; }
    public bool GroupByConfig { get; set; } = true;
    public bool SkipVirtual { get; set; } = true;
}

public sealed class SavePropertyRow
{
    public string Path { get; set; }
    public string Configuration { get; set; }
    public string DisplayName { get; set; }
    public SavePropertyChange[] Changes { get; set; }
}

public sealed class SavePropertyChange
{
    public string Name { get; set; }
    public string Value { get; set; }
}

public sealed class BlankSizeRow
{
    public string Path { get; set; }
    public string Configuration { get; set; }
    public string DisplayName { get; set; }
}
