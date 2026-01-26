# WordDocumentParser

A .NET library for parsing Word documents (.docx) into a hierarchical tree structure and writing them back with full round-trip fidelity.

## Features

- **Tree-based parsing**: Parses Word documents into a hierarchical structure organized by heading levels
- **Round-trip fidelity**: Preserves all formatting, styles, document properties, and dynamic references when writing back
- **Rich content support**: Handles paragraphs, headings, tables, images, lists, hyperlinks, and content controls
- **Content Controls**: Full support for all Structured Document Tag (SDT) types including text, date, dropdown, checkbox, and document property controls
- **Document Properties**: Easy access to core, extended, and custom properties with dictionary-style syntax
- **DOCPROPERTY Fields**: Detects and preserves document property field codes with value resolution
- **Full formatting capture**: Extracts paragraph, run, table, and image formatting properties
- **Table of Contents**: Preserves TOC and generates heading-based table of contents
- **Fluent API**: Extension methods for easy tree navigation, querying, and manipulation

## Installation

### Requirements

- .NET 9.0 or later
- DocumentFormat.OpenXml 3.4.1

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
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.4.1" />
</ItemGroup>
```

## Quick Start

### Parsing a Document

```csharp
using WordDocumentParser;

// Parse a Word document
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile("document.docx");

// Display the tree structure
Console.WriteLine(doc.Root.ToTreeString());

// Access document properties
Console.WriteLine($"Title: {doc["Title"]}");
Console.WriteLine($"Author: {doc["Author"]}");
```

### Working with Document Properties

```csharp
// Dictionary-style access (case-insensitive)
doc["Title"] = "My Document";
doc["Author"] = "John Doe";
doc["Department"] = "Engineering";  // Creates custom property

// Method-based access
string? company = doc.GetProperty("Company");
doc.SetProperty("Project", "WordParser");
bool hasKeywords = doc.HasProperty("Keywords");
doc.RemoveProperty("OldProperty");

// Get all properties
var allProps = doc.GetAllProperties();
foreach (var (name, value) in allProps)
{
    Console.WriteLine($"{name}: {value}");
}
```

### Writing a Document

```csharp
using WordDocumentParser.Extensions;

// Save a parsed document to a new file (preserves all formatting)
doc.SaveToFile("output.docx");

// Or save to a stream
using var stream = new MemoryStream();
doc.SaveToStream(stream);

// Or get as byte array
var bytes = doc.ToDocxBytes();
```

### Working with Content Controls

```csharp
using WordDocumentParser.Extensions;

// Find all content controls
var controls = doc.GetAllContentControls();

// Find by tag or alias
var clientControl = doc.FindContentControlByTag("ClientName");
var dateControl = doc.FindContentControlByAlias("Document Date");

// Update content control values
doc.SetContentControlValueByTag("ClientName", "ABC Corporation");
doc.SetContentControlValueByAlias("ProjectCode", "PRJ-2024-001");

// Get all tags in use
var tags = doc.GetContentControlTags();

// Remove content controls (keeps text content)
doc.RemoveContentControlByTag("TemporaryField");
doc.RemoveAllContentControls();  // Remove all, keep content
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

// Create WordDocument and save
var doc = new WordDocument(root);
doc["Title"] = "My New Document";
doc["Author"] = "Jane Smith";
doc.SaveToFile("new_document.docx");
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

#### `WordDocument`

The primary wrapper class representing a complete Word document.

```csharp
// Properties
WordDocument doc = ...;
DocumentNode root = doc.Root;                    // Root of content tree
string fileName = doc.FileName;                  // Original file name
DocumentPackageData package = doc.PackageData;  // Full package for round-trip

// Document Properties (dictionary-style access)
doc["Title"] = "New Title";
string? author = doc["Author"];

// Property methods
doc.SetProperty("Company", "ACME Corp");
string? value = doc.GetProperty("Keywords");
bool exists = doc.HasProperty("Subject");
doc.RemoveProperty("OldProp");
var all = doc.GetAllProperties();
```

#### `WordDocumentTreeParser`

Parses Word documents into a tree structure.

```csharp
using var parser = new WordDocumentTreeParser();

// Parse from file
var doc = parser.ParseFromFile("document.docx");

// Parse from stream
var doc = parser.ParseFromStream(stream, "documentName");
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
| `ContentControlProperties` | `ContentControlProperties?` | Content control metadata |
| `OriginalXml` | `string?` | Original OpenXML for round-trip fidelity |
| `Metadata` | `Dictionary<string, object>` | Additional metadata |

#### `ContentType` Enum

```csharp
public enum ContentType
{
    Document,       // Root document node
    Heading,        // Heading (H1-H9)
    Paragraph,      // Regular paragraph
    Table,          // Table
    Image,          // Image
    List,           // List container
    ListItem,       // List item
    HyperlinkText,  // Hyperlink text
    TextRun,        // Text run with formatting
    ContentControl  // Structured Document Tag (SDT)
}
```

### Content Controls

The library provides full support for Word content controls (Structured Document Tags).

#### `ContentControlType` Enum

```csharp
public enum ContentControlType
{
    Unknown,
    RichText,              // Rich text content
    PlainText,             // Plain text only
    Picture,               // Image placeholder
    Date,                  // Date picker
    DropDownList,          // Dropdown selection
    ComboBox,              // Editable dropdown
    Checkbox,              // Checkbox control
    RepeatingSection,      // Repeating content
    RepeatingSectionItem,
    BuildingBlockGallery,  // Quick Parts gallery
    Group,                 // Group container
    Bibliography,          // Bibliography field
    Citation,              // Citation field
    Equation,              // Equation placeholder
    DocumentProperty       // Linked to document property
}
```

#### `ContentControlProperties`

```csharp
public class ContentControlProperties
{
    public int? Id { get; set; }           // Unique identifier
    public string? Tag { get; set; }       // Developer tag
    public string? Alias { get; set; }     // Display name/title
    public ContentControlType Type { get; set; }
    public string? Value { get; set; }     // Current content

    // Lock settings
    public bool LockContentControl { get; set; }  // Can't delete
    public bool LockContents { get; set; }        // Can't edit

    // Data binding (for document property controls)
    public string? DataBindingXPath { get; set; }
    public string? DataBindingStoreItemId { get; set; }

    // Type-specific
    public string? DateFormat { get; set; }
    public List<ContentControlListItem> ListItems { get; set; }
    public bool? IsChecked { get; set; }
}
```

#### Content Control Extension Methods

```csharp
// Finding controls
var all = doc.GetAllContentControls();
var byType = doc.GetContentControlsByType(ContentControlType.Date);
var byTag = doc.FindContentControlByTag("CustomerName");
var byAlias = doc.FindContentControlByAlias("Invoice Date");
var byId = doc.FindContentControlById(12345);

// Getting metadata
var tags = doc.GetContentControlTags();
var props = doc.GetContentControlPropertiesByTag("FieldTag");

// Modifying values
doc.SetContentControlValueByTag("ClientName", "New Client");
doc.SetContentControlValueByAlias("Date", "2024-01-15");

// Removing controls (text content is preserved)
doc.RemoveContentControlByTag("TempField");
doc.RemoveContentControlByAlias("Draft Notice");
doc.RemoveContentControl(12345);
doc.RemoveAllContentControls();
```

### Extension Methods

#### Tree Navigation

```csharp
// Finding nodes
var matches = root.FindAll(n => n.Text.Contains("search term"));
var node = root.FindFirst(n => n.Type == ContentType.Table);
var section = root.GetSection("Methods");  // Case-insensitive

// Navigation
var path = node.GetPath();                 // Nodes from root
var breadcrumb = node.GetHeadingPath();    // "Doc > H1 > H2"
var siblings = node.GetSiblings();
var next = node.GetNextSibling();
var prev = node.GetPreviousSibling();
var depth = node.GetDepth();
var flat = root.Flatten();                 // All nodes as list
```

#### Tree Queries

```csharp
// Headings
var allHeadings = doc.GetAllHeadings();
var h2Headings = doc.GetHeadingsAtLevel(2);
var toc = doc.GetTableOfContents();  // (Level, Title, Node) tuples

// Content types
var tables = doc.GetAllTables();
var images = doc.GetAllImages();

// Text extraction
var text = section.GetAllText();

// Statistics
var counts = root.CountByType();
Console.WriteLine($"Paragraphs: {counts[ContentType.Paragraph]}");
```

#### Working with Tables

```csharp
var tableNode = doc.GetAllTables().First();
var table = tableNode.GetTableData();

// Dimensions
Console.WriteLine($"Size: {table.RowCount}x{table.ColumnCount}");

// Access as 2D array
var array = table.ToTextArray();
Console.WriteLine($"Cell [0,0]: {array[0, 0]}");

// Access specific cell
var cell = table.GetCell(1, 2);
Console.WriteLine($"Content: {cell?.TextContent}");
Console.WriteLine($"ColSpan: {cell?.ColSpan}");

// Iterate rows and cells
foreach (var row in table.Rows)
{
    foreach (var c in row.Cells)
    {
        Console.WriteLine($"[{c.RowIndex},{c.ColumnIndex}]: {c.TextContent}");
    }
}
```

#### Working with Images

```csharp
foreach (var imageNode in doc.GetAllImages())
{
    var img = imageNode.GetImageData();
    if (img != null)
    {
        Console.WriteLine($"Name: {img.Name}");
        Console.WriteLine($"Size: {img.WidthInches:F1}\" x {img.HeightInches:F1}\"");
        Console.WriteLine($"Type: {img.ContentType}");
        Console.WriteLine($"Alt: {img.AltText}");

        // Save to file
        if (img.Data != null)
        {
            File.WriteAllBytes($"extracted_{img.Name}", img.Data);
        }
    }
}
```

### Formatting Models

#### `RunFormatting` (Text-level)

```csharp
public class RunFormatting
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string? UnderlineStyle { get; set; }  // Single, Double, Wave
    public bool Strike { get; set; }
    public bool DoubleStrike { get; set; }
    public string? FontFamily { get; set; }
    public string? FontSize { get; set; }        // Half-points ("24" = 12pt)
    public string? Color { get; set; }           // Hex without #
    public string? Highlight { get; set; }
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }
    public bool SmallCaps { get; set; }
    public bool AllCaps { get; set; }
    public string? StyleId { get; set; }
}
```

#### `ParagraphFormatting`

```csharp
public class ParagraphFormatting
{
    public string? StyleId { get; set; }
    public string? Alignment { get; set; }       // Left, Center, Right, Both
    public string? IndentLeft { get; set; }      // Twips
    public string? IndentRight { get; set; }
    public string? IndentFirstLine { get; set; }
    public string? SpacingBefore { get; set; }
    public string? SpacingAfter { get; set; }
    public string? LineSpacing { get; set; }
    public bool KeepNext { get; set; }
    public bool KeepLines { get; set; }
    public bool PageBreakBefore { get; set; }
    public int? NumberingId { get; set; }
    public int? NumberingLevel { get; set; }
}
```

#### `TableFormatting`

```csharp
public class TableFormatting
{
    public string? Width { get; set; }
    public string? WidthType { get; set; }       // Pct, Dxa, Auto
    public string? Alignment { get; set; }
    public BorderFormatting? TopBorder { get; set; }
    public BorderFormatting? BottomBorder { get; set; }
    public BorderFormatting? LeftBorder { get; set; }
    public BorderFormatting? RightBorder { get; set; }
    public BorderFormatting? InsideHorizontalBorder { get; set; }
    public BorderFormatting? InsideVerticalBorder { get; set; }
    public List<string>? GridColumnWidths { get; set; }
}
```

## Document Properties

The library provides comprehensive access to all three types of Word document properties:

### Core Properties

Standard document metadata:
- Title, Subject, Creator/Author, Keywords, Description
- LastModifiedBy, Revision, Category, ContentStatus
- Created, Modified (dates)

### Extended Properties

Application-specific metadata:
- Template, Application, AppVersion
- Company, Manager
- Pages, Words, Characters, Lines, Paragraphs

### Custom Properties

User-defined key-value pairs that can store any additional metadata.

```csharp
// All property types accessible uniformly
doc["Title"] = "Annual Report";           // Core
doc["Company"] = "ACME Corporation";      // Extended
doc["ProjectCode"] = "PRJ-2024-001";      // Custom (auto-created)
doc["Confidential"] = "Yes";              // Custom

// Check property existence
if (doc.HasProperty("Department"))
{
    Console.WriteLine(doc["Department"]);
}

// Remove property
doc.RemoveProperty("OldField");

// List all
foreach (var (name, value) in doc.GetAllProperties())
{
    Console.WriteLine($"{name}: {value}");
}
```

## Project Structure

```
WordDocumentParser/
├── WordDocumentParser.sln
├── README.md
│
├── WordDocumentParser/                    # Core library
│   ├── Core/
│   │   ├── IDocumentParser.cs
│   │   ├── IDocumentWriter.cs
│   │   └── ContentType.cs
│   ├── Models/
│   │   ├── Formatting/                   # RunFormatting, ParagraphFormatting, etc.
│   │   ├── ContentControls/              # ContentControlProperties, types
│   │   ├── Tables/                       # TableData, TableRow, TableCell
│   │   ├── Images/                       # ImageData, ImageFormatting
│   │   └── Package/                      # CoreProperties, ExtendedProperties, etc.
│   ├── Extensions/
│   │   ├── TreeNavigationExtensions.cs
│   │   ├── TreeQueryExtensions.cs
│   │   ├── ContentControlExtensions.cs
│   │   ├── DocumentPropertyExtensions.cs
│   │   └── SerializationExtensions.cs
│   ├── Parsing/
│   │   ├── ParsingContext.cs
│   │   └── Extractors/
│   ├── WordDocument.cs                   # Main document wrapper
│   ├── DocumentNode.cs                   # Tree node
│   ├── WordDocumentTreeParser.cs         # Parser
│   └── WordDocumentTreeWriter.cs         # Writer
│
└── WordDocumentParser.Demo/              # Demo application
    ├── Program.cs
    └── Features/
        ├── Parsing/
        ├── ContentControls/
        ├── DocumentProperties/
        ├── DocumentCreation/
        └── RoundTrip/
```

## Round-Trip Fidelity

The library preserves the following when parsing and writing back:

- **Styles**: All paragraph and character styles
- **Formatting**: Bold, italic, underline, fonts, colors, spacing, borders, shading
- **Document Properties**: Core, extended, and custom properties
- **Content Controls**: All SDT types with properties and data binding
- **Dynamic References**: DOCPROPERTY fields, TOC, BIBLIOGRAPHY, CITATION
- **Structure**: Headers, footers, sections, page layout
- **Media**: Images with dimensions, alt text, and positioning
- **Tables**: Cell merging, borders, shading, column widths, nested content
- **Numbering**: List definitions and formatting
- **Hyperlinks**: External URLs and internal anchors
- **Glossary**: Building blocks, Quick Parts, document property fields

## Example: Complete Workflow

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

// 1. Parse an existing document
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile("template.docx");

// 2. Display structure
Console.WriteLine(doc.Root.ToTreeString());

// 3. Update document properties
doc["Title"] = "Quarterly Report Q4 2024";
doc["Author"] = "Finance Department";
doc["Department"] = "Finance";
doc["ReportDate"] = DateTime.Now.ToShortDateString();

// 4. Update content controls
doc.SetContentControlValueByTag("CompanyName", "ACME Corporation");
doc.SetContentControlValueByTag("ReportPeriod", "Q4 2024");
doc.SetContentControlValueByTag("PreparedBy", "John Smith");

// 5. Analyze content
var toc = doc.GetTableOfContents();
Console.WriteLine("\nTable of Contents:");
foreach (var (level, title, _) in toc)
{
    Console.WriteLine($"{"".PadLeft(level * 2)}{title}");
}

// 6. Work with tables
foreach (var table in doc.GetAllTables())
{
    var data = table.GetTableData();
    Console.WriteLine($"\nTable: {data.RowCount}x{data.ColumnCount}");
    Console.WriteLine($"Location: {table.GetHeadingPath()}");
}

// 7. Extract images
foreach (var img in doc.GetAllImages())
{
    var data = img.GetImageData();
    Console.WriteLine($"\nImage: {data?.Name} ({data?.ContentType})");
}

// 8. Get statistics
var counts = doc.Root.CountByType();
Console.WriteLine($"\nDocument Statistics:");
Console.WriteLine($"  Headings: {counts.GetValueOrDefault(ContentType.Heading)}");
Console.WriteLine($"  Paragraphs: {counts.GetValueOrDefault(ContentType.Paragraph)}");
Console.WriteLine($"  Tables: {counts.GetValueOrDefault(ContentType.Table)}");
Console.WriteLine($"  Images: {counts.GetValueOrDefault(ContentType.Image)}");

// 9. Save with full fidelity
doc.SaveToFile("output.docx");
Console.WriteLine("\nDocument saved successfully!");
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

## License

[Add your license here]

## Contributing

[Add contribution guidelines here]
