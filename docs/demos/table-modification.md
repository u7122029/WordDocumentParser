# Table Modification Demo

**Source:** `WordDocumentParser.Demo/Features/Tables/TableModificationDemo.cs`

Demonstrates structural table modifications: adding and inserting rows and columns, appending and removing cell text, and verifying that all changes survive a save/reload cycle.

## What It Does

1. Finds a specific table by matching header cell text
2. Adds a new row at the end with `AddRow()`
3. Adds a new column at the end with `AddColumn()`
4. Inserts a row at index 1 with `InsertRow()`
5. Inserts a column at index 1 with `InsertColumn()`
6. Appends text to a cell with `AppendText()`
7. Removes a substring from a cell with `RemoveText()`
8. Saves the document and re-parses to verify

## Key APIs Used

```csharp
// Find a specific table
var table = doc.FindAll(node =>
{
    if (node.Type != ContentType.Table) return false;
    return node.GetCellText(0, 0)!.Trim() == "Acronym";
}).First();

// Structural modifications
table.AddRow("Cell 1", "Cell 2", "Cell 3");
table.AddColumn("Header", "Value 1", "Value 2");
table.InsertRow(1, "Inserted A", "Inserted B", "Inserted C");
table.InsertColumn(1, "Col Header", "Row 1", "Row 2");

// Cell text operations
var cell = table.GetCell(0, 0);
cell.AppendText(" (appended paragraph)");
cell.RemoveText("Cell ");

// Dimensions after modifications
var (rows, cols) = table.GetDimensions();
```

## Related

- [Working with Tables](../articles/tables.md)
- [Table Parsing Demo](table-parsing.md)
