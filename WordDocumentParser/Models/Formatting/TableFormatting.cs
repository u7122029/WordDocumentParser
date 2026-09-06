using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Table-level formatting.
/// </summary>
public class TableFormatting : TrackedModel
{
    private string? _width;
    private string? _widthType;
    private string? _alignment;
    private string? _indentFromLeft;
    private BorderFormatting? _topBorder;
    private BorderFormatting? _bottomBorder;
    private BorderFormatting? _leftBorder;
    private BorderFormatting? _rightBorder;
    private BorderFormatting? _insideHorizontalBorder;
    private BorderFormatting? _insideVerticalBorder;
    private string? _cellMarginTop;
    private string? _cellMarginBottom;
    private string? _cellMarginLeft;
    private string? _cellMarginRight;
    private List<string>? _gridColumnWidths;

    /// <summary>Table width value (<c>w:tblW/@w:w</c>).</summary>
    public string? Width { get => _width; set => Set(ref _width, value); }

    /// <summary>Table width unit as an OOXML token: <c>"dxa"</c>, <c>"pct"</c>, <c>"auto"</c>.</summary>
    public string? WidthType { get => _widthType; set => Set(ref _widthType, value); }

    /// <summary>Table alignment as an OOXML token: <c>"left"</c>, <c>"center"</c>, <c>"right"</c>.</summary>
    public string? Alignment { get => _alignment; set => Set(ref _alignment, value); }

    /// <summary>Table indent from the left margin, in twips.</summary>
    public string? IndentFromLeft { get => _indentFromLeft; set => Set(ref _indentFromLeft, value); }

    /// <summary>Outer top border.</summary>
    public BorderFormatting? TopBorder { get => _topBorder; set => Set(ref _topBorder, value); }

    /// <summary>Outer bottom border.</summary>
    public BorderFormatting? BottomBorder { get => _bottomBorder; set => Set(ref _bottomBorder, value); }

    /// <summary>Outer left border.</summary>
    public BorderFormatting? LeftBorder { get => _leftBorder; set => Set(ref _leftBorder, value); }

    /// <summary>Outer right border.</summary>
    public BorderFormatting? RightBorder { get => _rightBorder; set => Set(ref _rightBorder, value); }

    /// <summary>Border between rows.</summary>
    public BorderFormatting? InsideHorizontalBorder { get => _insideHorizontalBorder; set => Set(ref _insideHorizontalBorder, value); }

    /// <summary>Border between columns.</summary>
    public BorderFormatting? InsideVerticalBorder { get => _insideVerticalBorder; set => Set(ref _insideVerticalBorder, value); }

    /// <summary>Default top cell margin, in twips.</summary>
    public string? CellMarginTop { get => _cellMarginTop; set => Set(ref _cellMarginTop, value); }

    /// <summary>Default bottom cell margin, in twips.</summary>
    public string? CellMarginBottom { get => _cellMarginBottom; set => Set(ref _cellMarginBottom, value); }

    /// <summary>Default left cell margin, in twips.</summary>
    public string? CellMarginLeft { get => _cellMarginLeft; set => Set(ref _cellMarginLeft, value); }

    /// <summary>Default right cell margin, in twips.</summary>
    public string? CellMarginRight { get => _cellMarginRight; set => Set(ref _cellMarginRight, value); }

    /// <summary>Widths of the table grid columns, in twips, in column order.</summary>
    public List<string>? GridColumnWidths { get => _gridColumnWidths; set => Set(ref _gridColumnWidths, value); }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public TableFormatting Clone()
    {
        var clone = new TableFormatting
        {
            _width = _width,
            _widthType = _widthType,
            _alignment = _alignment,
            _indentFromLeft = _indentFromLeft,
            _topBorder = _topBorder?.Clone(),
            _bottomBorder = _bottomBorder?.Clone(),
            _leftBorder = _leftBorder?.Clone(),
            _rightBorder = _rightBorder?.Clone(),
            _insideHorizontalBorder = _insideHorizontalBorder?.Clone(),
            _insideVerticalBorder = _insideVerticalBorder?.Clone(),
            _cellMarginTop = _cellMarginTop,
            _cellMarginBottom = _cellMarginBottom,
            _cellMarginLeft = _cellMarginLeft,
            _cellMarginRight = _cellMarginRight,
            _gridColumnWidths = _gridColumnWidths is null ? null : [.. _gridColumnWidths]
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
