# Document Structure Demo

**Source:** `WordDocumentParser.Demo/Features/Parsing/DocumentStructureDemo.cs`

Demonstrates parsing a Word document and displaying its full structure, including the heading tree, statistics, table of contents, table dimensions, and image metadata.

## What It Does

1. Parses a `.docx` file into the heading-based tree
2. Prints the tree structure with `ToTreeString()`
3. Shows document statistics &mdash; counts by `ContentType`
4. Generates a table of contents from the heading hierarchy
5. Lists all tables with dimensions and heading-path location
6. Lists all images with name, dimensions, and content type

## Key APIs Used

```csharp
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile(filePath);

// Display tree
Console.WriteLine(doc.ToTreeString());

// Statistics
var counts = doc.CountByType();

// Table of contents
var toc = doc.GetTableOfContents();
foreach (var (level, title, _) in toc) { ... }

// Tables
var tables = doc.GetAllTables();
foreach (var table in tables)
{
    var tableData = table.GetTableData();
    Console.WriteLine($"{tableData.RowCount} rows x {tableData.ColumnCount} columns");
    Console.WriteLine($"Location: {table.GetHeadingPath()}");
}

// Images
var images = doc.GetAllImages();
foreach (var image in images)
{
    var imageData = image.GetImageData();
    Console.WriteLine($"{imageData.Name}: {imageData.WidthInches}\" x {imageData.HeightInches}\"");
}
```

## Related

- [Document Tree Structure](../articles/document-tree.md)
- [Tree Navigation](../articles/navigation.md)
