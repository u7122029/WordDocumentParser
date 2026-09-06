using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Tables;

/// <summary>
/// A table row with its cells and formatting.
/// </summary>
public class TableRow : TrackedModel
{
    private int _rowIndex;
    private List<TableCell> _cells = [];
    private bool _isHeader;
    private TableRowFormatting? _formatting;

    /// <summary>Zero-based index of this row in the table.</summary>
    public int RowIndex { get => _rowIndex; set => Set(ref _rowIndex, value); }

    /// <summary>Cells contained in this row, in document order.</summary>
    public List<TableCell> Cells { get => _cells; set { Set(ref _cells, value ?? []); MarkCellsChanged(); } }

    /// <summary>Whether this row repeats as a header at the top of each page.</summary>
    public bool IsHeader { get => _isHeader; set => Set(ref _isHeader, value); }

    /// <summary>Row formatting properties.</summary>
    public TableRowFormatting? Formatting { get => _formatting; set => Set(ref _formatting, value); }

    /// <summary>Records that cells were added to or removed from this row.</summary>
    public void MarkCellsChanged() => MarkChanged(nameof(Cells));

    /// <summary>True when cells have been added to or removed from this row.</summary>
    public bool IsCellsChanged => IsChanged(nameof(Cells));

    /// <summary>True when the row, its formatting, or any of its cells have pending changes.</summary>
    public bool HasRowChanges =>
        HasChanges ||
        _formatting?.HasChanges is true ||
        _cells.Exists(cell => cell.HasCellChanges);

    /// <summary>Clears the change record on this row, its formatting, and its cells.</summary>
    public void AcceptAllChanges()
    {
        AcceptChanges();
        _formatting?.AcceptChanges();
        foreach (var cell in _cells)
        {
            cell.AcceptAllChanges();
        }
    }
}
