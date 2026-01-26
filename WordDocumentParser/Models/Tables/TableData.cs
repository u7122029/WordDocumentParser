using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Tables;

/// <summary>
/// Represents a complete table structure with rows, cells, and formatting.
/// </summary>
public class TableData
{
    /// <summary>All rows in the table</summary>
    public List<TableRow> Rows { get; set; } = [];

    /// <summary>Number of rows in the table</summary>
    public int RowCount => Rows.Count;

    /// <summary>Number of columns in the table</summary>
    public int ColumnCount { get; set; }

    /// <summary>Table-level formatting properties for round-trip fidelity</summary>
    public TableFormatting? Formatting { get; set; }

    /// <summary>
    /// Gets a cell at the specified position.
    /// </summary>
    /// <param name="row">Zero-based row index</param>
    /// <param name="column">Zero-based column index</param>
    /// <returns>The cell at the position, or null if not found</returns>
    public TableCell? GetCell(int row, int column)
    {
        if (row < 0 || row >= Rows.Count)
            return null;

        var tableRow = Rows[row];
        return tableRow.Cells.Find(c => c.ColumnIndex == column);
    }

    /// <summary>
    /// Gets all text content as a 2D array for easy access.
    /// </summary>
    /// <returns>2D string array with cell text content</returns>
    public string[,] ToTextArray()
    {
        var result = new string[RowCount, ColumnCount];
        for (var i = 0; i < RowCount; i++)
        {
            for (var j = 0; j < ColumnCount; j++)
            {
                var cell = GetCell(i, j);
                result[i, j] = cell?.TextContent ?? string.Empty;
            }
        }
        return result;
    }
}
