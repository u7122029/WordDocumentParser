using WordDocumentParser.FormattingModels;

namespace WordDocumentParser;

/// <summary>
/// Represents a table cell with its content
/// </summary>
public class TableCell
{
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public int RowSpan { get; set; } = 1;
    public int ColSpan { get; set; } = 1;
    public List<DocumentNode> Content { get; set; } = [];
    public string TextContent => string.Join(" ", Content.ConvertAll(c => c.Text));

    /// <summary>
    /// Cell formatting properties for round-trip fidelity
    /// </summary>
    public TableCellFormatting? Formatting { get; set; }
}

/// <summary>
/// Represents a table row
/// </summary>
public class TableRow
{
    public int RowIndex { get; set; }
    public List<TableCell> Cells { get; set; } = [];
    public bool IsHeader { get; set; }

    /// <summary>
    /// Row formatting properties for round-trip fidelity
    /// </summary>
    public TableRowFormatting? Formatting { get; set; }
}

/// <summary>
/// Represents a complete table structure
/// </summary>
public class TableData
{
    public List<TableRow> Rows { get; set; } = [];
    public int RowCount => Rows.Count;
    public int ColumnCount { get; set; }

    /// <summary>
    /// Table formatting properties for round-trip fidelity
    /// </summary>
    public TableFormatting? Formatting { get; set; }

    /// <summary>
    /// Gets a cell at the specified position
    /// </summary>
    public TableCell? GetCell(int row, int column)
    {
        if (row < 0 || row >= Rows.Count) return null;
        var tableRow = Rows[row];
        return tableRow.Cells.Find(c => c.ColumnIndex == column);
    }

    /// <summary>
    /// Gets all text content as a 2D array for easy access
    /// </summary>
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

/// <summary>
/// Represents image data extracted from the document
/// </summary>
public class ImageData
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public byte[]? Data { get; set; }
    public double WidthInches { get; set; }
    public double HeightInches { get; set; }
    public string? AltText { get; set; }
    public string? Description { get; set; }

    /// <summary>
    /// Image dimensions in EMUs for precise round-trip
    /// </summary>
    public long WidthEmu { get; set; }
    public long HeightEmu { get; set; }

    /// <summary>
    /// Image positioning and formatting for round-trip fidelity
    /// </summary>
    public ImageFormatting? Formatting { get; set; }
}
