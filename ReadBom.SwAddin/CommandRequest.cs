namespace ReadBom.SwAddin;

public sealed class CommandRequest
{
    public string Command { get; set; }
    public string Configuration { get; set; }
    public string Name { get; set; }
    public string Value { get; set; }
    public string[] PropertyNames { get; set; }
    public bool GroupByConfig { get; set; } = true;
    public bool SkipVirtual { get; set; } = true;
}
