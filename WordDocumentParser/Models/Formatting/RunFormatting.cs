namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Represents text formatting properties for a run of text
/// </summary>
public class RunFormatting
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string? UnderlineStyle { get; set; } // Single, Double, Wave, etc.
    public bool Strike { get; set; }
    public bool DoubleStrike { get; set; }
    public string? FontFamily { get; set; }
    public string? FontFamilyAscii { get; set; }
    public string? FontFamilyEastAsia { get; set; }
    public string? FontFamilyComplexScript { get; set; }
    public string? FontSize { get; set; } // In half-points (e.g., "24" = 12pt)
    public string? FontSizeComplexScript { get; set; }
    public string? Color { get; set; } // Hex color without #
    public string? Highlight { get; set; } // Highlight color name
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }
    public bool SmallCaps { get; set; }
    public bool AllCaps { get; set; }
    public string? Shading { get; set; } // Background shading
    public string? StyleId { get; set; } // Character style reference

    public bool HasFormatting =>
        Bold || Italic || Underline || Strike || DoubleStrike ||
        FontFamily is not null || FontSize is not null || Color is not null ||
        Highlight is not null || Superscript || Subscript || SmallCaps || AllCaps;

    public RunFormatting Clone() => new()
    {
        Bold = Bold,
        Italic = Italic,
        Underline = Underline,
        UnderlineStyle = UnderlineStyle,
        Strike = Strike,
        DoubleStrike = DoubleStrike,
        FontFamily = FontFamily,
        FontFamilyAscii = FontFamilyAscii,
        FontFamilyEastAsia = FontFamilyEastAsia,
        FontFamilyComplexScript = FontFamilyComplexScript,
        FontSize = FontSize,
        FontSizeComplexScript = FontSizeComplexScript,
        Color = Color,
        Highlight = Highlight,
        Superscript = Superscript,
        Subscript = Subscript,
        SmallCaps = SmallCaps,
        AllCaps = AllCaps,
        Shading = Shading,
        StyleId = StyleId
    };
}