using WordDocumentParser.Core;

namespace WordDocumentParser.Models.Formatting;

/// <summary>
/// Border formatting for a paragraph, table, or cell edge.
/// </summary>
public class BorderFormatting : TrackedModel
{
    private string? _style;
    private string? _size;
    private string? _color;
    private string? _space;

    /// <summary>Border style as an OOXML token, for example <c>"single"</c> or <c>"dashDotStroked"</c>.</summary>
    public string? Style { get => _style; set => Set(ref _style, value); }

    /// <summary>Border width in eighths of a point, so <c>"4"</c> is 0.5pt.</summary>
    public string? Size { get => _size; set => Set(ref _size, value); }

    /// <summary>Border colour as a hex triplet, or <c>"auto"</c>.</summary>
    public string? Color { get => _color; set => Set(ref _color, value); }

    /// <summary>Space between the border and the content, in points.</summary>
    public string? Space { get => _space; set => Set(ref _space, value); }

    /// <summary>Creates a copy that carries the same pending changes as this instance.</summary>
    /// <returns>The copy.</returns>
    public BorderFormatting Clone()
    {
        var clone = new BorderFormatting
        {
            _style = _style,
            _size = _size,
            _color = _color,
            _space = _space
        };
        clone.CopyChangesFrom(this);
        return clone;
    }
}
