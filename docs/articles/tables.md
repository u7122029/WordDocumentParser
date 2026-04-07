# Working with Tables

@WordDocumentParser.Extensions.TableExtensions provides full 2D access to tables, including structural modification, cell formatting, and nested table support. All structural operations preserve the original table style by modifying the XML in-place rather than rebuilding from scratch.

## Reading Tables

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

// Find tables
var tables = doc.FindAllTables(includeNested: true);
var table = tables.First();
var (rows, cols) = table.GetDimensions();

// 2D cell access — returns a TableCell with TextContent, Formatting, etc.
string? text = table.GetCellText(0, 0);
var cell = table.GetCell(1, 2);
Console.WriteLine(cell?.TextContent);

// Iterate by row, column, or all cells
foreach (var c in table.GetRowCells(0)) { /* header cells */ }
foreach (var c in table.GetColumnCells(0)) { /* first column */ }
foreach (var (row, col, c) in table.EnumerateCells()) { /* all */ }

// Convert to 2D array or text representation
string[,]? array = table.ToTextArray();
Console.WriteLine(table.ToTextRepresentation());
```

> [!NOTE]
> `FindAllTables(includeNested: false)` returns only top-level tables. Set `includeNested: true` to also return tables inside cells of other tables.

## Modifying Table Structure

These operations clone the formatting from adjacent rows/columns, so new rows and columns inherit the original table's style:

```csharp
// Add rows and columns at the end
table.AddRow("Cell 1", "Cell 2", "Cell 3");
table.AddColumn("Header", "Value 1", "Value 2");

// Insert at a specific index
table.InsertRow(1, "Inserted A", "Inserted B", "Inserted C");
table.InsertColumn(0, "New First Col Header", "Row 1", "Row 2");

// Remove rows and columns
table.RemoveRow(3);
table.RemoveColumn(2);
```

> [!WARNING]
> The number of values passed to `AddRow` / `InsertRow` must match the table's column count. Likewise, the number of values passed to `AddColumn` / `InsertColumn` must match the row count.

## Modifying Cell Content

```csharp
table.SetCellText(0, 0, "Updated header");

var cell = table.GetCell(1, 0);
cell.AppendText("Additional paragraph");
cell.RemoveText("unwanted substring");
cell.ClearContent();
```

## Cell and Row Formatting

Use @WordDocumentParser.Models.Formatting.TableCellFormatting and @WordDocumentParser.Models.Formatting.TableRowFormatting to control appearance:

```csharp
// Cell formatting
cell.SetShading("FFFF00");               // Yellow background
cell.SetVerticalAlignment("center");      // top, center, bottom
cell.SetBorders("single", 8, "000000");  // Style, size, color
cell.SetContentStyle("Heading2");         // Apply paragraph style

// Row operations
var row = table.GetRow(0);
row.SetAsHeader(true);                    // Repeat on page breaks
row.SetRowShading("D9E2F3");             // Row background color

// Table alignment
table.SetTableAlignment("Center");
```

## Nested Tables

Cells can contain nested tables. Use the @WordDocumentParser.Models.Tables.TableCell methods to access them:

```csharp
if (cell.HasNestedTable())
{
    var nested = cell.GetFirstNestedTable();
    var (r, c) = nested.GetDimensions();
    nested.SetCellText(0, 0, "Nested cell updated");
}
```

## Next Steps

- **[Document Merging](merging.md)** — combine documents with automatic resource remapping
- **[Table Parsing Demo](../demos/table-parsing.md)** — full table access and formatting example
- **[Table Modification Demo](../demos/table-modification.md)** — structural modification walkthrough
- @WordDocumentParser.Extensions.TableExtensions — API reference for all 30+ table methods
