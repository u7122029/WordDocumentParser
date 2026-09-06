using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Tables;

/// <summary>
/// A table cell with its content and formatting.
/// </summary>
public class TableCell : TrackedModel
{
    private int _rowIndex;
    private int _columnIndex;
    private int _rowSpan = 1;
    private int _colSpan = 1;
    private List<DocumentNode> _content = [];
    private TableCellFormatting? _formatting;

    /// <summary>Zero-based row index of this cell.</summary>
    public int RowIndex { get => _rowIndex; set => Set(ref _rowIndex, value); }

    /// <summary>
    /// Zero-based logical column index of this cell, counting merged columns.
    /// A cell following a <see cref="ColSpan"/> of 2 starts two columns further right, so this can
    /// differ from the cell's position within <see cref="TableRow.Cells"/>.
    /// </summary>
    public int ColumnIndex { get => _columnIndex; set => Set(ref _columnIndex, value); }

    /// <summary>Number of rows this cell spans. 1 means no vertical merge.</summary>
    public int RowSpan { get => _rowSpan; set => Set(ref _rowSpan, value); }

    /// <summary>Number of columns this cell spans. 1 means no horizontal merge.</summary>
    public int ColSpan { get => _colSpan; set => Set(ref _colSpan, value); }

    /// <summary>
    /// Document nodes contained within this cell.
    /// Mutating this list directly does not mark the cell as changed; call
    /// <see cref="MarkContentChanged"/> afterwards, or use the extension methods, which do so.
    /// </summary>
    public List<DocumentNode> Content { get => _content; set { Set(ref _content, value ?? []); MarkContentChanged(); } }

    /// <summary>Combined text of all nodes in this cell.</summary>
    public string TextContent => string.Join(" ", Content.ConvertAll(c => c.Text));

    /// <summary>Cell formatting properties.</summary>
    public TableCellFormatting? Formatting { get => _formatting; set => Set(ref _formatting, value); }

    /// <summary>
    /// Records that this cell's content list was added to, removed from, or cleared.
    /// </summary>
    /// <remarks>
    /// The writer needs this to distinguish a cell whose content the caller emptied from a cell that
    /// simply parsed to no content nodes. Without it, clearing a cell was a silent no-op and the old
    /// text was written back out — which matters when the caller was redacting.
    /// </remarks>
    public void MarkContentChanged() => MarkChanged(nameof(Content));

    /// <summary>True when this cell's content list has been added to, removed from, or cleared.</summary>
    public bool IsContentChanged => IsChanged(nameof(Content));

    /// <summary>True when the cell, its content nodes, or its formatting have pending changes.</summary>
    public bool HasCellChanges =>
        HasChanges ||
        _formatting?.HasChanges is true ||
        _content.Exists(node => node.HasChanges);

    /// <summary>Clears the change record on this cell, its formatting, and its content nodes.</summary>
    public void AcceptAllChanges()
    {
        AcceptChanges();
        _formatting?.AcceptAllChanges();
        foreach (var node in _content)
        {
            node.AcceptAllChanges();
        }
    }
}
