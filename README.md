# WordDocumentParser

A .NET library for parsing Word documents (.docx) into a hierarchical tree structure and writing them back with full round-trip fidelity.

## Features

- **Tree-based parsing**: Parses Word documents into a hierarchical structure organized by heading levels
- **Round-trip fidelity**: Preserves all formatting, styles, document properties, and dynamic references when writing back
- **Rich content support**: Handles paragraphs, headings, tables, images, lists, hyperlinks, and structured document tags
- **Full formatting capture**: Extracts paragraph, run, table, and image formatting properties
- **Document properties**: Preserves core, extended, and custom document properties
- **Dynamic references**: Maintains DOCPROPERTY fields, BIBLIOGRAPHY, CITATION, and other field codes
- **Fluent API**: Extension methods for easy tree navigation and manipulation

## Installation

### Requirements

- .NET 9.0 or later
- DocumentFormat.OpenXml 3.0.2

### Add to your project

Reference the `WordDocumentParser` project or add the compiled DLL to your project:

```xml
<ItemGroup>
    <ProjectReference Include="..\WordDocumentParser\WordDocumentParser.csproj" />
</ItemGroup>
```

Or if using the compiled library:

```xml
<ItemGroup>
    <Reference Include="WordDocumentParser">
        <HintPath>path\to\WordDocumentParser.dll</HintPath>
    </Reference>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.0.2" />
</ItemGroup>
```

## Quick Start

### Parsing a Document

```csharp
using WordDocumentParser;

// Parse a Word document
using var parser = new WordDocumentTreeParser();
var documentTree = parser.ParseFromFile("document.docx");

// Display the tree structure
Console.WriteLine(documentTree.ToTreeString());
```

### Writing a Document

```csharp
// Save a parsed document to a new file (preserves all formatting)
documentTree.SaveToFile("output.docx");

// Or save to a stream
using var stream = new MemoryStream();
documentTree.SaveToStream(stream);

// Or get as byte array
var bytes = documentTree.ToDocxBytes();
```

### Creating a Document from Scratch

```csharp
// Create a new document
var root = new DocumentNode(ContentType.Document, "My Document");

// Add a heading
var heading = new DocumentNode(ContentType.Heading, 1, "Introduction");
root.AddChild(heading);

// Add paragraphs under the heading
heading.AddChild(new DocumentNode(ContentType.Paragraph, "This is the first paragraph."));
heading.AddChild(new DocumentNode(ContentType.Paragraph, "This is the second paragraph."));

// Add another section
var methods = new DocumentNode(ContentType.Heading, 1, "Methods");
root.AddChild(methods);
methods.AddChild(new DocumentNode(ContentType.Paragraph, "Description of methods..."));

// Save the document
root.SaveToFile("new_document.docx");
```

## Document Tree Structure

The parser organizes content hierarchically based on heading levels:

```
Document (root)
  +-- H1: Introduction
  |     +-- Paragraph: Some text...
  |     +-- H2: Background
  |     |     +-- Paragraph: More text...
  |     |     +-- Table: [Table: 3x4]
  |     +-- H2: Purpose
  |           +-- Paragraph: Purpose text...
  +-- H1: Methods
        +-- H2: Data Collection
        |     +-- Image: [Image: figure1.png]
        +-- H2: Analysis
              +-- Paragraph: Analysis details...
```

## API Reference

### Core Classes

#### `WordDocumentTreeParser`

Parses Word documents into a tree structure.

```csharp
using var parser = new WordDocumentTreeParser();

// Parse from file
var tree = parser.ParseFromFile("document.docx");

// Parse from stream
var tree = parser.ParseFromStream(stream, "documentName");
```

#### `DocumentNode`

Represents a node in the document tree.

| Property | Type | Description |
|----------|------|-------------|
| `Id` | `string` | Unique identifier for the node |
| `Type` | `ContentType` | Type of content (Document, Heading, Paragraph, etc.) |
| `HeadingLevel` | `int` | Heading level (1-9) for headings, 0 for other types |
| `Text` | `string` | Plain text content |
| `Children` | `List<DocumentNode>` | Child nodes |
| `Parent` | `DocumentNode?` | Parent node reference |
| `Runs` | `List<FormattedRun>` | Formatted text runs with styling |
| `ParagraphFormatting` | `ParagraphFormatting?` | Paragraph-level formatting |
| `OriginalXml` | `string?` | Original OpenXML for round-trip fidelity |
| `PackageData` | `DocumentPackageData?` | Document package data (root node only) |
| `Metadata` | `Dictionary<string, object>` | Additional metadata |

#### `ContentType` Enum

```csharp
public enum ContentType
{
    Document,      // Root document node
    Heading,       // Heading (H1-H9)
    Paragraph,     // Regular paragraph
    Table,         // Table
    Image,         // Image
    List,          // List container
    ListItem,      // List item
    HyperlinkText, // Hyperlink text
    TextRun        // Text run with formatting
}
```

### Extension Methods

The `DocumentTreeExtensions` class provides fluent methods for tree navigation:

#### Finding Nodes

```csharp
// Find all nodes matching a predicate
var matches = root.FindAll(n => n.Text.Contains("search term"));

// Find first matching node
var node = root.FindFirst(n => n.Type == ContentType.Table);

// Get section by heading text (case-insensitive partial match)
var section = root.GetSection("Methods");
```

#### Accessing Headings

```csharp
// Get all headings
var allHeadings = root.GetAllHeadings();

// Get headings at specific level
var h2Headings = root.GetHeadingsAtLevel(2);

// Get table of contents
var toc = root.GetTableOfContents();
foreach (var (level, title, node) in toc)
{
    Console.WriteLine($"{"".PadLeft(level * 2)}{title}");
}
```

#### Accessing Tables

```csharp
// Get all tables
var tables = root.GetAllTables();

foreach (var tableNode in tables)
{
    var tableData = tableNode.GetTableData();
    if (tableData != null)
    {
        // Get dimensions
        Console.WriteLine($"Table: {tableData.RowCount}x{tableData.ColumnCount}");

        // Access as 2D array
        var array = tableData.ToTextArray();
        Console.WriteLine($"Cell [0,0]: {array[0, 0]}");

        // Access specific cell
        var cell = tableData.GetCell(0, 1);
        Console.WriteLine($"Cell content: {cell?.TextContent}");
    }
}
```

#### Accessing Images

```csharp
// Get all images
var images = root.GetAllImages();

foreach (var imageNode in images)
{
    var imageData = imageNode.GetImageData();
    if (imageData != null)
    {
        Console.WriteLine($"Image: {imageData.Name}");
        Console.WriteLine($"Size: {imageData.WidthInches:F1}\" x {imageData.HeightInches:F1}\"");
        Console.WriteLine($"Type: {imageData.ContentType}");

        // Save image to file
        if (imageData.Data != null)
        {
            File.WriteAllBytes($"{imageData.Name}.png", imageData.Data);
        }
    }
}
```

#### Navigation

```csharp
// Get text content under a node
var text = section.GetAllText();

// Get breadcrumb path
var path = node.GetHeadingPath(); // e.g., "Document > Introduction > Background"

// Get path as node list
var pathNodes = node.GetPath();

// Navigate siblings
var nextSibling = node.GetNextSibling();
var prevSibling = node.GetPreviousSibling();
var allSiblings = node.GetSiblings();

// Get all descendants
var descendants = root.GetDescendants();

// Get node depth
var depth = node.GetDepth();
```

#### Statistics

```csharp
// Count nodes by type
var counts = root.CountByType();
Console.WriteLine($"Paragraphs: {counts[ContentType.Paragraph]}");
Console.WriteLine($"Tables: {counts[ContentType.Table]}");

// Flatten tree to list
var allNodes = root.Flatten();
```

### Formatting Models

#### `RunFormatting`

Text run formatting properties:

```csharp
public class RunFormatting
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string? UnderlineStyle { get; set; }  // Single, Double, Wave, etc.
    public bool Strike { get; set; }
    public bool DoubleStrike { get; set; }
    public string? FontFamily { get; set; }
    public string? FontSize { get; set; }        // In half-points (e.g., "24" = 12pt)
    public string? Color { get; set; }           // Hex color without #
    public string? Highlight { get; set; }       // Highlight color name
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }
    public bool SmallCaps { get; set; }
    public bool AllCaps { get; set; }
    public string? StyleId { get; set; }         // Character style reference
}
```

#### `ParagraphFormatting`

Paragraph formatting properties:

```csharp
public class ParagraphFormatting
{
    public string? StyleId { get; set; }
    public string? Alignment { get; set; }       // Left, Center, Right, Both
    public string? IndentLeft { get; set; }      // In twips
    public string? IndentRight { get; set; }
    public string? IndentFirstLine { get; set; }
    public string? SpacingBefore { get; set; }   // In twips
    public string? SpacingAfter { get; set; }
    public string? LineSpacing { get; set; }
    public bool KeepNext { get; set; }
    public bool KeepLines { get; set; }
    public bool PageBreakBefore { get; set; }
    // ... and more
}
```

#### `TableData`, `TableRow`, `TableCell`

Table structure with formatting:

```csharp
var tableData = tableNode.GetTableData();

// Access rows
foreach (var row in tableData.Rows)
{
    Console.WriteLine($"Row {row.RowIndex}, IsHeader: {row.IsHeader}");

    foreach (var cell in row.Cells)
    {
        Console.WriteLine($"  Cell [{cell.RowIndex},{cell.ColumnIndex}]: {cell.TextContent}");
        Console.WriteLine($"  ColSpan: {cell.ColSpan}, RowSpan: {cell.RowSpan}");
    }
}
```

#### `ImageData`

Image data with dimensions and positioning:

```csharp
var imageData = imageNode.GetImageData();

// Dimensions
double widthInches = imageData.WidthInches;
double heightInches = imageData.HeightInches;
long widthEmu = imageData.WidthEmu;   // For precise round-trip

// Image data
byte[] data = imageData.Data;
string contentType = imageData.ContentType;
string altText = imageData.AltText;

// Positioning
var formatting = imageData.Formatting;
bool isInline = formatting.IsInline;
string wrapType = formatting.WrapType;  // None, Square, Tight, etc.
```

## Project Structure

```
WordDocumentParser/
├── WordDocumentParser.sln           # Solution file
├── README.md                        # This file
│
├── WordDocumentParser/              # Core library (class library)
│   ├── WordDocumentParser.csproj
│   ├── DocumentNode.cs              # Document tree node
│   ├── DocumentPackageData.cs       # Package data for round-trip
│   ├── DocumentTreeExtensions.cs    # Extension methods
│   ├── FormattingModels.cs          # Formatting classes
│   ├── TableAndImageModels.cs       # Table and image models
│   ├── WordDocumentTreeParser.cs    # Document parser
│   └── WordDocumentTreeWriter.cs    # Document writer
│
└── WordDocumentParser.Demo/         # Demo application (console app)
    ├── WordDocumentParser.Demo.csproj
    ├── Program.cs                   # Demo entry point
    ├── DocumentComparison.cs        # Document comparison utilities
    ├── DocumentValidator.cs         # Document validation utilities
    └── TableHelper.cs               # Table creation helpers
```

## Building

```bash
# Build the entire solution
dotnet build

# Build only the library
dotnet build WordDocumentParser/WordDocumentParser.csproj

# Run the demo
dotnet run --project WordDocumentParser.Demo
```

## Round-Trip Fidelity

The library preserves the following when parsing and writing back:

- **Styles**: All paragraph and character styles
- **Formatting**: Bold, italic, underline, fonts, colors, spacing, etc.
- **Document Properties**: Title, author, subject, keywords, custom properties
- **Dynamic References**: DOCPROPERTY fields, BIBLIOGRAPHY, CITATION, TOC
- **Structure**: Headers, footers, sections, page layout
- **Media**: Images with original dimensions and positioning
- **Tables**: Cell merging, borders, shading, column widths
- **Numbering**: List definitions and formatting
- **Glossary**: Building blocks and Quick Parts

## Example: Complete Workflow

```csharp
using WordDocumentParser;

// 1. Parse an existing document
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile("input.docx");

// 2. Analyze the structure
Console.WriteLine("Document Structure:");
Console.WriteLine(doc.ToTreeString());

var stats = doc.CountByType();
Console.WriteLine($"\nStatistics:");
foreach (var (type, count) in stats)
{
    Console.WriteLine($"  {type}: {count}");
}

// 3. Find specific content
var methodsSection = doc.GetSection("Methods");
if (methodsSection != null)
{
    Console.WriteLine($"\nMethods section text:");
    Console.WriteLine(methodsSection.GetAllText());
}

// 4. Work with tables
foreach (var table in doc.GetAllTables())
{
    var data = table.GetTableData();
    Console.WriteLine($"\nTable at: {table.GetHeadingPath()}");
    Console.WriteLine($"Dimensions: {data.RowCount}x{data.ColumnCount}");
}

// 5. Save with full fidelity
doc.SaveToFile("output.docx");
```

## License

[Add your license here]

## Contributing

[Add contribution guidelines here]
