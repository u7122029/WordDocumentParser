# Table Parsing Demo

**Source:** `WordDocumentParser.Demo/Features/Tables/TableParsing.cs`

Demonstrates full table access: finding tables (including nested), iterating cells, modifying content, applying formatting, and converting to 2D arrays.

## What It Does

1. Finds all tables (top-level and nested) with `FindAllTables()`
2. Displays each table's dimensions and text representation
3. Demonstrates cell access by row, column, and coordinates
4. Modifies cell text with `SetCellText()`
5. Applies header row formatting (yellow shading, header repeat)
6. Applies alternating row shading (zebra striping)
7. Accesses and modifies nested tables within cells
8. Sets cell borders and vertical alignment
9. Converts a table to a 2D `string[,]` array
10. Saves and re-parses to verify all changes persisted

## Key APIs Used

```csharp
// Finding tables
var allTables = doc.FindAllTables(includeNested: true);
var topLevelOnly = doc.FindAllTables(includeNested: false);

// Cell access
var (rows, cols) = table.GetDimensions();
string? text = table.GetCellText(0, 0);
var cell = table.GetCell(1, 2);

// Row and column iteration
foreach (var cell in table.GetRowCells(0)) { ... }
foreach (var cell in table.GetColumnCells(0)) { ... }

// Formatting
headerRow.SetRowShading("FFFF00");
headerRow.SetAsHeader(true);
cell.SetBorders(style: "single", size: 8, color: "000000");
cell.SetVerticalAlignment("center");

// Nested tables
if (cell.HasNestedTable())
{
    var nested = cell.GetFirstNestedTable();
    nested.SetCellText(0, 0, "[NESTED-MODIFIED] " + original);
}

// 2D array
string[,]? array = table.ToTextArray();
```

## Related

- [Working with Tables](../articles/tables.md)
- [Table Modification Demo](table-modification.md)
