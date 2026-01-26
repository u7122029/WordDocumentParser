using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Models.Tables;

/// <summary>
/// Represents a table cell with its content and formatting.
/// </summary>
public class TableCell
{
    /// <summary>Zero-based row index of this cell</summary>
    public int RowIndex { get; set; }

    /// <summary>Zero-based column index of this cell</summary>
    public int ColumnIndex { get; set; }

    /// <summary>Number of rows this cell spans (1 = no span)</summary>
    public int RowSpan { get; set; } = 1;

    /// <summary>Number of columns this cell spans (1 = no span)</summary>
    public int ColSpan { get; set; } = 1;

    /// <summary>Document nodes contained within this cell</summary>
    public List<DocumentNode> Content { get; set; } = [];

    /// <summary>Combined text content of all nodes in this cell</summary>
    public string TextContent => string.Join(" ", Content.ConvertAll(c => c.Text));

    /// <summary>Cell formatting properties for round-trip fidelity</summary>
    public TableCellFormatting? Formatting { get; set; }
}
