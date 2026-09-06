using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Tables;

/// <summary>
/// A complete table structure with rows, cells, and formatting.
/// </summary>
public class TableData : TrackedModel
{
    private List<TableRow> _rows = [];
    private int _columnCount;
    private TableFormatting? _formatting;

    /// <summary>All rows in the table, in document order.</summary>
    public List<TableRow> Rows { get => _rows; set { Set(ref _rows, value ?? []); MarkRowsChanged(); } }

    /// <summary>Number of rows in the table.</summary>
    public int RowCount => Rows.Count;

    /// <summary>Number of logical columns in the table, counting merged columns once each.</summary>
    public int ColumnCount { get => _columnCount; set => Set(ref _columnCount, value); }

    /// <summary>Table-level formatting properties.</summary>
    public TableFormatting? Formatting { get => _formatting; set => Set(ref _formatting, value); }

    /// <summary>Records that rows were added to or removed from this table.</summary>
    public void MarkRowsChanged() => MarkChanged(nameof(Rows));

    /// <summary>True when rows have been added to or removed from this table.</summary>
    public bool IsRowsChanged => IsChanged(nameof(Rows));

    /// <summary>True when the table, its formatting, or anything inside it has pending changes.</summary>
    public bool HasTableChanges =>
        HasChanges ||
        _formatting?.HasChanges is true ||
        _rows.Exists(row => row.HasRowChanges);

    /// <summary>Clears the change record on the table and everything inside it.</summary>
    public void AcceptAllChanges()
    {
        AcceptChanges();
        _formatting?.AcceptChanges();
        foreach (var row in _rows)
        {
            row.AcceptAllChanges();
        }
    }

    /// <summary>
    /// Gets the cell occupying the given logical position.
    /// </summary>
    /// <param name="row">Zero-based row index.</param>
    /// <param name="column">Zero-based logical column index.</param>
    /// <returns>The cell covering that position, or null when there is none.</returns>
    /// <remarks>
    /// A horizontally merged cell covers every column in its span, so asking for any column within
    /// the span returns the merged cell.
    /// </remarks>
    public TableCell? GetCell(int row, int column)
    {
        if (row < 0 || row >= Rows.Count || column < 0)
            return null;

        foreach (var cell in Rows[row].Cells)
        {
            if (column >= cell.ColumnIndex && column < cell.ColumnIndex + Math.Max(1, cell.ColSpan))
                return cell;
        }

        return null;
    }

    /// <summary>
    /// Gets all cell text as a two-dimensional array indexed by row and logical column.
    /// </summary>
    /// <returns>The cell text, with empty strings for positions no cell covers.</returns>
    /// <remarks>
    /// Filled by walking each row once rather than searching the row for every column, so the cost
    /// is linear in the number of cells instead of quadratic in the table's dimensions.
    /// </remarks>
    public string[,] ToTextArray()
    {
        var result = new string[RowCount, ColumnCount];

        for (var rowIndex = 0; rowIndex < RowCount; rowIndex++)
        {
            for (var column = 0; column < ColumnCount; column++)
            {
                result[rowIndex, column] = string.Empty;
            }

            foreach (var cell in Rows[rowIndex].Cells)
            {
                var text = cell.TextContent;
                var span = Math.Max(1, cell.ColSpan);

                for (var offset = 0; offset < span; offset++)
                {
                    var column = cell.ColumnIndex + offset;
                    if (column >= 0 && column < ColumnCount)
                    {
                        result[rowIndex, column] = text;
                    }
                }
            }
        }

        return result;
    }
}
