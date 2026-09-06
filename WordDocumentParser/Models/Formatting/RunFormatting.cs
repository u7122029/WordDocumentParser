using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Text formatting properties for a run of text.
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>). The writer applies only the properties
/// a caller actually assigned, leaving everything else in the original run XML untouched.
/// </remarks>
public class RunFormatting : TrackedModel
{
    private bool _bold;
    private bool _italic;
    private bool _underline;
    private string? _underlineStyle;
    private bool _strike;
    private bool _doubleStrike;
    private string? _fontFamily;
    private string? _fontFamilyAscii;
    private string? _fontFamilyEastAsia;
    private string? _fontFamilyComplexScript;
    private string? _fontSize;
    private string? _fontSizeComplexScript;
    private string? _color;
    private string? _highlight;
    private bool _superscript;
    private bool _subscript;
    private bool _smallCaps;
    private bool _allCaps;
    private string? _shading;
    private string? _styleId;

    /// <summary>Bold.</summary>
    public bool Bold { get => _bold; set => Set(ref _bold, value); }

    /// <summary>Italic.</summary>
    public bool Italic { get => _italic; set => Set(ref _italic, value); }

    /// <summary>Underlined.</summary>
    public bool Underline { get => _underline; set => Set(ref _underline, value); }

    /// <summary>Underline style as an OOXML token, for example <c>"single"</c> or <c>"wave"</c>.</summary>
    public string? UnderlineStyle { get => _underlineStyle; set => Set(ref _underlineStyle, value); }

    /// <summary>Single strikethrough.</summary>
    public bool Strike { get => _strike; set => Set(ref _strike, value); }

    /// <summary>Double strikethrough.</summary>
    public bool DoubleStrike { get => _doubleStrike; set => Set(ref _doubleStrike, value); }

    /// <summary>Font for high-ANSI characters (<c>w:rFonts/@w:hAnsi</c>).</summary>
    public string? FontFamily { get => _fontFamily; set => Set(ref _fontFamily, value); }

    /// <summary>Font for ASCII characters (<c>w:rFonts/@w:ascii</c>).</summary>
    public string? FontFamilyAscii { get => _fontFamilyAscii; set => Set(ref _fontFamilyAscii, value); }

    /// <summary>Font for East Asian characters (<c>w:rFonts/@w:eastAsia</c>).</summary>
    public string? FontFamilyEastAsia { get => _fontFamilyEastAsia; set => Set(ref _fontFamilyEastAsia, value); }

    /// <summary>Font for complex scripts (<c>w:rFonts/@w:cs</c>).</summary>
    public string? FontFamilyComplexScript { get => _fontFamilyComplexScript; set => Set(ref _fontFamilyComplexScript, value); }

    /// <summary>Font size in half-points, so <c>"24"</c> is 12pt.</summary>
    public string? FontSize { get => _fontSize; set => Set(ref _fontSize, value); }

    /// <summary>Complex-script font size in half-points.</summary>
    public string? FontSizeComplexScript { get => _fontSizeComplexScript; set => Set(ref _fontSizeComplexScript, value); }

    /// <summary>Text colour as a hex triplet without the leading <c>#</c>.</summary>
    public string? Color { get => _color; set => Set(ref _color, value); }

    /// <summary>Highlight colour as an OOXML token, for example <c>"yellow"</c>.</summary>
    public string? Highlight { get => _highlight; set => Set(ref _highlight, value); }

    /// <summary>Superscript.</summary>
    public bool Superscript { get => _superscript; set => Set(ref _superscript, value); }

    /// <summary>Subscript.</summary>
    public bool Subscript { get => _subscript; set => Set(ref _subscript, value); }

    /// <summary>Small capitals.</summary>
    public bool SmallCaps { get => _smallCaps; set => Set(ref _smallCaps, value); }

    /// <summary>All capitals.</summary>
    public bool AllCaps { get => _allCaps; set => Set(ref _allCaps, value); }

    /// <summary>Background shading fill as a hex triplet.</summary>
    public string? Shading { get => _shading; set => Set(ref _shading, value); }

    /// <summary>Character style reference (<c>w:rStyle</c>).</summary>
    public string? StyleId { get => _styleId; set => Set(ref _styleId, value); }

    /// <summary>Names of the font properties, for change queries.</summary>
    internal static readonly string[] FontProperties =
    [
        nameof(FontFamily), nameof(FontFamilyAscii),
        nameof(FontFamilyEastAsia), nameof(FontFamilyComplexScript)
    ];

    /// <summary>True when any visible formatting is applied.</summary>
    public bool HasFormatting =>
        Bold || Italic || Underline || Strike || DoubleStrike ||
        FontFamily is not null || FontSize is not null || Color is not null ||
        Highlight is not null || Superscript || Subscript || SmallCaps || AllCaps;

    /// <summary>True when a caller has assigned any of the font-family properties.</summary>
    public bool IsFontChanged => IsAnyChanged(FontProperties);

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public RunFormatting Clone()
    {
        var clone = new RunFormatting
        {
            _bold = _bold,
            _italic = _italic,
            _underline = _underline,
            _underlineStyle = _underlineStyle,
            _strike = _strike,
            _doubleStrike = _doubleStrike,
            _fontFamily = _fontFamily,
            _fontFamilyAscii = _fontFamilyAscii,
            _fontFamilyEastAsia = _fontFamilyEastAsia,
            _fontFamilyComplexScript = _fontFamilyComplexScript,
            _fontSize = _fontSize,
            _fontSizeComplexScript = _fontSizeComplexScript,
            _color = _color,
            _highlight = _highlight,
            _superscript = _superscript,
            _subscript = _subscript,
            _smallCaps = _smallCaps,
            _allCaps = _allCaps,
            _shading = _shading,
            _styleId = _styleId
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
