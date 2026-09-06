using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Paragraph formatting properties.
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>). The writer applies only the properties
/// a caller actually assigned, leaving the rest of the original paragraph XML untouched.
/// String-valued enumerations hold the OOXML wire token, for example <c>"center"</c>, not
/// <c>"Center"</c>; the writer accepts either spelling.
/// </remarks>
public class ParagraphFormatting : TrackedModel
{
    private string? _styleId;
    private string? _alignment;
    private string? _indentLeft;
    private string? _indentRight;
    private string? _indentFirstLine;
    private string? _indentHanging;
    private string? _spacingBefore;
    private string? _spacingAfter;
    private string? _lineSpacing;
    private string? _lineSpacingRule;
    private bool _keepNext;
    private bool _keepLines;
    private bool _pageBreakBefore;
    private bool _widowControl;
    private string? _outlineLevel;
    private string? _shadingFill;
    private string? _shadingColor;
    private BorderFormatting? _topBorder;
    private BorderFormatting? _bottomBorder;
    private BorderFormatting? _leftBorder;
    private BorderFormatting? _rightBorder;
    private int? _numberingId;
    private int? _numberingLevel;

    /// <summary>Paragraph style reference (<c>w:pStyle</c>).</summary>
    public string? StyleId { get => _styleId; set => Set(ref _styleId, value); }

    /// <summary>Alignment as an OOXML token: <c>"left"</c>, <c>"center"</c>, <c>"right"</c>, <c>"both"</c>.</summary>
    public string? Alignment { get => _alignment; set => Set(ref _alignment, value); }

    /// <summary>Left indent in twips.</summary>
    public string? IndentLeft { get => _indentLeft; set => Set(ref _indentLeft, value); }

    /// <summary>Right indent in twips.</summary>
    public string? IndentRight { get => _indentRight; set => Set(ref _indentRight, value); }

    /// <summary>First-line indent in twips.</summary>
    public string? IndentFirstLine { get => _indentFirstLine; set => Set(ref _indentFirstLine, value); }

    /// <summary>Hanging indent in twips.</summary>
    public string? IndentHanging { get => _indentHanging; set => Set(ref _indentHanging, value); }

    /// <summary>Space before the paragraph, in twips.</summary>
    public string? SpacingBefore { get => _spacingBefore; set => Set(ref _spacingBefore, value); }

    /// <summary>Space after the paragraph, in twips.</summary>
    public string? SpacingAfter { get => _spacingAfter; set => Set(ref _spacingAfter, value); }

    /// <summary>Line spacing, interpreted according to <see cref="LineSpacingRule"/>.</summary>
    public string? LineSpacing { get => _lineSpacing; set => Set(ref _lineSpacing, value); }

    /// <summary>Line spacing rule as an OOXML token: <c>"auto"</c>, <c>"exact"</c>, <c>"atLeast"</c>.</summary>
    public string? LineSpacingRule { get => _lineSpacingRule; set => Set(ref _lineSpacingRule, value); }

    /// <summary>Keep this paragraph on the same page as the next one.</summary>
    public bool KeepNext { get => _keepNext; set => Set(ref _keepNext, value); }

    /// <summary>Keep all lines of this paragraph on one page.</summary>
    public bool KeepLines { get => _keepLines; set => Set(ref _keepLines, value); }

    /// <summary>Start this paragraph on a new page.</summary>
    public bool PageBreakBefore { get => _pageBreakBefore; set => Set(ref _pageBreakBefore, value); }

    /// <summary>Suppress widow and orphan lines.</summary>
    public bool WidowControl { get => _widowControl; set => Set(ref _widowControl, value); }

    /// <summary>Outline level, where <c>"0"</c> is the top level.</summary>
    public string? OutlineLevel { get => _outlineLevel; set => Set(ref _outlineLevel, value); }

    /// <summary>Shading fill as a hex triplet.</summary>
    public string? ShadingFill { get => _shadingFill; set => Set(ref _shadingFill, value); }

    /// <summary>Shading pattern colour as a hex triplet.</summary>
    public string? ShadingColor { get => _shadingColor; set => Set(ref _shadingColor, value); }

    /// <summary>Top border.</summary>
    public BorderFormatting? TopBorder { get => _topBorder; set => Set(ref _topBorder, value); }

    /// <summary>Bottom border.</summary>
    public BorderFormatting? BottomBorder { get => _bottomBorder; set => Set(ref _bottomBorder, value); }

    /// <summary>Left border.</summary>
    public BorderFormatting? LeftBorder { get => _leftBorder; set => Set(ref _leftBorder, value); }

    /// <summary>Right border.</summary>
    public BorderFormatting? RightBorder { get => _rightBorder; set => Set(ref _rightBorder, value); }

    /// <summary>Numbering definition ID (<c>w:numId</c>).</summary>
    public int? NumberingId { get => _numberingId; set => Set(ref _numberingId, value); }

    /// <summary>Numbering level (<c>w:ilvl</c>).</summary>
    public int? NumberingLevel { get => _numberingLevel; set => Set(ref _numberingLevel, value); }

    /// <summary>True when any formatting worth writing is present.</summary>
    public bool HasFormatting =>
        StyleId is not null || Alignment is not null ||
        IndentLeft is not null || IndentRight is not null ||
        IndentFirstLine is not null || IndentHanging is not null ||
        SpacingBefore is not null || SpacingAfter is not null ||
        LineSpacing is not null || KeepNext || KeepLines ||
        PageBreakBefore || ShadingFill is not null;

    /// <summary>
    /// True when this paragraph's formatting, or any of its border objects, has pending changes.
    /// </summary>
    public bool HasFormattingChanges =>
        HasChanges ||
        _topBorder?.HasChanges is true || _bottomBorder?.HasChanges is true ||
        _leftBorder?.HasChanges is true || _rightBorder?.HasChanges is true;

    /// <summary>Clears the change record on this model and on its border objects.</summary>
    public void AcceptAllChanges()
    {
        AcceptChanges();
        _topBorder?.AcceptChanges();
        _bottomBorder?.AcceptChanges();
        _leftBorder?.AcceptChanges();
        _rightBorder?.AcceptChanges();
    }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public ParagraphFormatting Clone()
    {
        var clone = new ParagraphFormatting
        {
            _styleId = _styleId,
            _alignment = _alignment,
            _indentLeft = _indentLeft,
            _indentRight = _indentRight,
            _indentFirstLine = _indentFirstLine,
            _indentHanging = _indentHanging,
            _spacingBefore = _spacingBefore,
            _spacingAfter = _spacingAfter,
            _lineSpacing = _lineSpacing,
            _lineSpacingRule = _lineSpacingRule,
            _keepNext = _keepNext,
            _keepLines = _keepLines,
            _pageBreakBefore = _pageBreakBefore,
            _widowControl = _widowControl,
            _outlineLevel = _outlineLevel,
            _shadingFill = _shadingFill,
            _shadingColor = _shadingColor,
            _topBorder = _topBorder?.Clone(),
            _bottomBorder = _bottomBorder?.Clone(),
            _leftBorder = _leftBorder?.Clone(),
            _rightBorder = _rightBorder?.Clone(),
            _numberingId = _numberingId,
            _numberingLevel = _numberingLevel
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
