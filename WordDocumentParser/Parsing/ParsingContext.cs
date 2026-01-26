using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordDocumentParser.Parsing;

/// <summary>
/// Holds shared state during document parsing.
/// Passed to extractors to avoid tight coupling.
/// </summary>
internal sealed class ParsingContext
{
    public required WordprocessingDocument Document { get; init; }
    public required MainDocumentPart MainPart { get; init; }
    public Dictionary<string, Style> StyleCache { get; } = [];
    public Dictionary<string, string> HyperlinkUrls { get; } = [];

    /// <summary>
    /// Caches styles from the document for quick lookup.
    /// </summary>
    public void CacheStyles()
    {
        var stylesPart = MainPart.StyleDefinitionsPart;
        if (stylesPart?.Styles is null) return;

        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId?.Value is not null)
            {
                StyleCache[style.StyleId.Value] = style;
            }
        }
    }

    /// <summary>
    /// Caches hyperlink relationships for URL resolution.
    /// </summary>
    public void CacheHyperlinkRelationships()
    {
        foreach (var rel in MainPart.HyperlinkRelationships)
        {
            HyperlinkUrls[rel.Id] = rel.Uri.ToString();
        }
    }

    /// <summary>
    /// Gets the URL for a hyperlink relationship ID.
    /// </summary>
    public string? GetHyperlinkUrl(string? relationshipId)
        => relationshipId is not null && HyperlinkUrls.TryGetValue(relationshipId, out var url) ? url : null;

    /// <summary>
    /// Gets a style by its ID.
    /// </summary>
    public Style? GetStyle(string? styleId)
        => styleId is not null && StyleCache.TryGetValue(styleId, out var style) ? style : null;
}
