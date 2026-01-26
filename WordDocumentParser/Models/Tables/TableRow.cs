using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Tables;

/// <summary>
/// Represents a table row with its cells and formatting.
/// </summary>
public class TableRow
{
    /// <summary>Zero-based index of this row in the table</summary>
    public int RowIndex { get; set; }

    /// <summary>Cells contained in this row</summary>
    public List<TableCell> Cells { get; set; } = [];

    /// <summary>Whether this row is a header row (repeats on page breaks)</summary>
    public bool IsHeader { get; set; }

    /// <summary>Row formatting properties for round-trip fidelity</summary>
    public TableRowFormatting? Formatting { get; set; }
}
