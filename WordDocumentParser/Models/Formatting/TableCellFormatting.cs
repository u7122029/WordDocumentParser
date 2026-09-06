using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Table cell formatting.
/// </summary>
/// <remarks>
/// Assignments are tracked (see <see cref="TrackedModel"/>) so the writer can rewrite only the cell
/// properties a caller actually set, preserving the original element ordering for the rest.
/// </remarks>
public class TableCellFormatting : TrackedModel
{
    private string? _width;
    private string? _widthType;
    private int _gridSpan = 1;
    private string? _verticalMerge;
    private string? _verticalAlignment;
    private string? _shadingFill;
    private string? _shadingColor;
    private string? _shadingPattern;
    private BorderFormatting? _topBorder;
    private BorderFormatting? _bottomBorder;
    private BorderFormatting? _leftBorder;
    private BorderFormatting? _rightBorder;
    private string? _textDirection;
    private bool _noWrap;

    /// <summary>Cell width value (<c>w:tcW/@w:w</c>).</summary>
    public string? Width { get => _width; set => Set(ref _width, value); }

    /// <summary>Cell width unit as an OOXML token: <c>"dxa"</c>, <c>"pct"</c>, <c>"auto"</c>.</summary>
    public string? WidthType { get => _widthType; set => Set(ref _widthType, value); }

    /// <summary>Number of grid columns this cell spans. 1 means no horizontal merge.</summary>
    public int GridSpan { get => _gridSpan; set => Set(ref _gridSpan, value); }

    /// <summary>Vertical merge state: <c>"Restart"</c>, <c>"Continue"</c>, or null.</summary>
    public string? VerticalMerge { get => _verticalMerge; set => Set(ref _verticalMerge, value); }

    /// <summary>Vertical alignment as an OOXML token: <c>"top"</c>, <c>"center"</c>, <c>"bottom"</c>.</summary>
    public string? VerticalAlignment { get => _verticalAlignment; set => Set(ref _verticalAlignment, value); }

    /// <summary>Shading fill as a hex triplet, or <c>"auto"</c>.</summary>
    public string? ShadingFill { get => _shadingFill; set => Set(ref _shadingFill, value); }

    /// <summary>Shading pattern colour as a hex triplet.</summary>
    public string? ShadingColor { get => _shadingColor; set => Set(ref _shadingColor, value); }

    /// <summary>Shading pattern as an OOXML token, for example <c>"clear"</c>.</summary>
    public string? ShadingPattern { get => _shadingPattern; set => Set(ref _shadingPattern, value); }

    /// <summary>Top border.</summary>
    public BorderFormatting? TopBorder { get => _topBorder; set => Set(ref _topBorder, value); }

    /// <summary>Bottom border.</summary>
    public BorderFormatting? BottomBorder { get => _bottomBorder; set => Set(ref _bottomBorder, value); }

    /// <summary>Left border.</summary>
    public BorderFormatting? LeftBorder { get => _leftBorder; set => Set(ref _leftBorder, value); }

    /// <summary>Right border.</summary>
    public BorderFormatting? RightBorder { get => _rightBorder; set => Set(ref _rightBorder, value); }

    /// <summary>Text direction as an OOXML token, for example <c>"tbRl"</c>.</summary>
    public string? TextDirection { get => _textDirection; set => Set(ref _textDirection, value); }

    /// <summary>Suppress text wrapping in this cell.</summary>
    public bool NoWrap { get => _noWrap; set => Set(ref _noWrap, value); }

    /// <summary>True when any border object on this cell has pending changes.</summary>
    public bool HasBorderChanges =>
        IsAnyChanged(nameof(TopBorder), nameof(BottomBorder), nameof(LeftBorder), nameof(RightBorder)) ||
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
    public TableCellFormatting Clone()
    {
        var clone = new TableCellFormatting
        {
            _width = _width,
            _widthType = _widthType,
            _gridSpan = _gridSpan,
            _verticalMerge = _verticalMerge,
            _verticalAlignment = _verticalAlignment,
            _shadingFill = _shadingFill,
            _shadingColor = _shadingColor,
            _shadingPattern = _shadingPattern,
            _topBorder = _topBorder?.Clone(),
            _bottomBorder = _bottomBorder?.Clone(),
            _leftBorder = _leftBorder?.Clone(),
            _rightBorder = _rightBorder?.Clone(),
            _textDirection = _textDirection,
            _noWrap = _noWrap
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
