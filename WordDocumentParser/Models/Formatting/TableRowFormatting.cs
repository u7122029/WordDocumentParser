using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Table row formatting.
/// </summary>
public class TableRowFormatting : TrackedModel
{
    private string? _height;
    private string? _heightRule;
    private bool _isHeader;
    private bool _cantSplit;

    /// <summary>Row height in twips.</summary>
    public string? Height { get => _height; set => Set(ref _height, value); }

    /// <summary>Height rule as an OOXML token: <c>"auto"</c>, <c>"exact"</c>, <c>"atLeast"</c>.</summary>
    public string? HeightRule { get => _heightRule; set => Set(ref _heightRule, value); }

    /// <summary>Repeat this row as a header at the top of each page.</summary>
    public bool IsHeader { get => _isHeader; set => Set(ref _isHeader, value); }

    /// <summary>Prevent this row from splitting across pages.</summary>
    public bool CantSplit { get => _cantSplit; set => Set(ref _cantSplit, value); }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public TableRowFormatting Clone()
    {
        var clone = new TableRowFormatting
        {
            _height = _height,
            _heightRule = _heightRule,
            _isHeader = _isHeader,
            _cantSplit = _cantSplit
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
