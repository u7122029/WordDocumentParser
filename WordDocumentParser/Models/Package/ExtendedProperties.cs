namespace WordDocumentParser.Models.Package;

/// <summary>
/// Extended document properties
/// </summary>
public class ExtendedProperties
{
    public string? Template { get; set; }
    public string? Application { get; set; }
    public string? AppVersion { get; set; }
    public string? Company { get; set; }
    public int? Pages { get; set; }
    public int? Words { get; set; }
    public int? Characters { get; set; }
    public int? CharactersWithSpaces { get; set; }
    public int? Lines { get; set; }
    public int? Paragraphs { get; set; }
    public string? Manager { get; set; }
    public int? TotalTime { get; set; }
}