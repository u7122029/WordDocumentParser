namespace WordDocumentParser.Models.Package;

/// <summary>
/// A custom XML part and its properties part.
/// </summary>
public class CustomXmlPartData
{
    /// <summary>The part's XML content.</summary>
    public string XmlContent { get; set; } = string.Empty;

    /// <summary>The associated <c>customXmlProps</c> XML, when the part has one.</summary>
    public string? PropertiesXml { get; set; }
}
