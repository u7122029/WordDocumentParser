# WordDocumentParser

A .NET library for parsing Word documents (.docx) into a hierarchical tree structure and writing them back with full round-trip fidelity. Unlike raw OpenXML, this library provides a heading-based document tree, 2D table access, document merging with resource remapping, and surgical modification that preserves formatting you didn't touch.

## Features

- **Tree-based parsing**: Organizes flat OpenXML elements into a heading-level hierarchy with parent-child navigation
- **Round-trip fidelity**: Stores original XML per node and only regenerates what you modify — formatting, styles, and structure pass through untouched
- **Table manipulation**: 2D cell access, add/insert/remove rows and columns, cell formatting, nested table support — all with style preservation
- **Document merging**: Append, concatenate, extract sections, and insert nodes across documents with automatic image/hyperlink relationship remapping
- **Content controls**: Find, update, and remove Structured Document Tags by tag, alias, ID, or type
- **Style and font management**: Bulk style changes, font replacement, per-run or per-paragraph font control with character-set awareness
- **Document properties**: Uniform dictionary-style access to core, extended, and custom properties
- **DOCPROPERTY fields**: Detects and preserves document property field codes with value resolution

## Installation

### Requirements

- .NET 9.0 or later
- DocumentFormat.OpenXml 3.4.1

### Add to your project

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
using WordDocumentParser.Extensions;

// Parse a Word document into a tree
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile("document.docx");

// Display the heading-based tree structure
Console.WriteLine(doc.Root.ToTreeString());

// Access document properties
Console.WriteLine($"Title: {doc["Title"]}");
Console.WriteLine($"Author: {doc["Author"]}");
```

### Writing a Document

```csharp
// Save a parsed (and optionally modified) document — preserves all formatting
doc.SaveToFile("output.docx");

// Or save to a stream / byte array
using var stream = new MemoryStream();
doc.SaveToStream(stream);
byte[] bytes = doc.ToDocxBytes();
```

### Creating a Document from Scratch

```csharp
var root = new DocumentNode(ContentType.Document, "My Document");

var heading = new DocumentNode(ContentType.Heading, 1, "Introduction");
root.AddChild(heading);
heading.AddChild(new DocumentNode(ContentType.Paragraph, "First paragraph."));
heading.AddChild(new DocumentNode(ContentType.Paragraph, "Second paragraph."));

var doc = new WordDocument(root);
doc["Title"] = "My New Document";
doc.SaveToFile("new_document.docx");
```

## Document Tree Structure

The parser organizes content hierarchically based on heading levels. OpenXML stores body elements as a flat list — this library infers the hierarchy:

```
Document (root)
  +-- H1: Introduction
  |     +-- Paragraph: Some text...
  |     +-- H2: Background
  |     |     +-- Paragraph: More text...
  |     |     +-- Table: [3x4]
  |     +-- H2: Purpose
  |           +-- Paragraph: Purpose text...
  +-- H1: Methods
        +-- H2: Data Collection
        |     +-- Image: [figure1.png]
        +-- H2: Analysis
              +-- Paragraph: Analysis details...
```

This enables operations impossible with flat OpenXML: `GetSection("Methods")`, `GetHeadingPath()` breadcrumbs, `GetTableOfContents()`, upward navigation via `Parent`.

## Working with Tables

### Reading Tables

```csharp
// Find tables
var tables = doc.FindAllTables(includeNested: true);
var table = tables.First();
var (rows, cols) = table.GetDimensions();

// 2D cell access
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

### Modifying Table Structure

All structural operations preserve the original table style by modifying the XML in-place rather than rebuilding from scratch.

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

### Modifying Cell Content

```csharp
// Set, append, or remove text
table.SetCellText(0, 0, "Updated header");
var cell = table.GetCell(1, 0);
cell.AppendText("Additional paragraph");
cell.RemoveText("unwanted substring");
cell.ClearContent();
```

### Cell and Row Formatting

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

### Nested Tables

```csharp
if (cell.HasNestedTable())
{
    var nested = cell.GetFirstNestedTable();
    var (r, c) = nested.GetDimensions();
    nested.SetCellText(0, 0, "Nested cell updated");
}
```

## Document Merging

```csharp
using var parser1 = new WordDocumentTreeParser();
using var parser2 = new WordDocumentTreeParser();
var doc1 = parser1.ParseFromFile("first.docx");
var doc2 = parser2.ParseFromFile("second.docx");

// Append one document to another (handles image/hyperlink remapping)
doc1.AppendDocument(doc2, addPageBreak: true);

// Append multiple documents
doc1.AppendDocuments(new[] { doc2, doc3 }, addPageBreaks: true);

// Create a new combined document
var combined = DocumentMergeExtensions.ConcatenateDocuments(
    new[] { doc1, doc2, doc3 });
```

### Section Extraction and Insertion

```csharp
// Extract a section by heading text
var section = doc.ExtractSection("Chapter 3", includeNestedHeadings: true);

// Extract all tables or headings at a level
var tables = doc.ExtractTables();
var h2s = doc.ExtractHeadingsAtLevel(2);

// Insert nodes from one document into another
target.InsertSectionAfterHeading("Chapter 2", source, "New Section");
target.ReplaceSection("Old Section", source, "Replacement Section");
```

## Content Controls

```csharp
// Find controls
var all = doc.GetAllContentControls();
var byTag = doc.FindContentControlByTag("ClientName");
var byAlias = doc.FindContentControlByAlias("Document Date");
var byId = doc.FindContentControlById(12345);
var byType = doc.GetContentControlsByType(ContentControlType.Date);

// Update values
doc.SetContentControlValueByTag("ClientName", "ABC Corporation");
doc.SetContentControlValueByAlias("ProjectCode", "PRJ-2024-001");

// Remove controls (text content is preserved)
doc.RemoveContentControlByTag("TemporaryField");
doc.RemoveAllContentControls();

// Get metadata
var tags = doc.GetContentControlTags();
var props = doc.GetContentControlPropertiesByTag("FieldTag");
```

## Document Properties

All three property types (core, extended, custom) are accessible through a uniform API:

```csharp
// Dictionary-style access (case-insensitive)
doc["Title"] = "Annual Report";           // Core property
doc["Company"] = "ACME Corporation";      // Extended property
doc["ProjectCode"] = "PRJ-2024-001";      // Custom (auto-created)

// Method-based access
string? company = doc.GetProperty("Company");
doc.SetProperty("Department", "Engineering");
bool exists = doc.HasProperty("Keywords");
doc.RemoveProperty("OldField");

// List all properties
foreach (var (name, value) in doc.GetAllProperties())
    Console.WriteLine($"{name}: {value}");
```

## Style Management

```csharp
// Find nodes by style
var normalParas = doc.FindByStyle("Normal");
var headings = doc.FindByStyles("Heading1", "Heading2");

// Change styles
node.ChangeStyle("Heading2");  // Also updates HeadingLevel and Type
doc.ChangeStyleBulk("OldStyle", "NewStyle");
doc.ChangeStyleWhere(n => n.Text.StartsWith("Note:"), "NoteStyle");

// Query style usage
var distribution = doc.GetStyleDistribution();
bool isHeading = node.HasStyle("Heading1");
```

## Font Management

```csharp
// Set font for a paragraph, section, or entire document
node.SetParagraphFont("Calibri");
doc.SetDocumentFont("Arial");

// Target specific text within a paragraph
node.SetFontForText("important", "Arial Black");
node.SetFontForRange(0, 5, "Courier New");

// Replace fonts globally
doc.ReplaceFont("Times New Roman", "Calibri");

// Query fonts in use
var fonts = doc.GetAllFontsUsed();
```

## Tree Navigation

```csharp
// Finding nodes
var matches = root.FindAll(n => n.Text.Contains("search term"));
var first = root.FindFirst(n => n.Type == ContentType.Table);
var section = root.GetSection("Methods");

// Navigation
var path = node.GetPath();              // Ancestor chain from root
var breadcrumb = node.GetHeadingPath(); // "Doc > Chapter 1 > Section 1.1"
var siblings = node.GetSiblings();
var next = node.GetNextSibling();
var depth = node.GetDepth();
var flat = root.Flatten();

// Queries
var toc = doc.GetTableOfContents();     // (Level, Title, Node) tuples
var counts = doc.Root.CountByType();    // ContentType → count
var allText = section.GetAllText();     // Recursive text extraction
```

## API Reference

### Core Classes

| Class | Description |
|-------|-------------|
| `WordDocument` | Primary document wrapper with property access and content tree |
| `DocumentNode` | Tree node with type, text, formatting, children, and parent reference |
| `WordDocumentTreeParser` | Parses .docx files into the tree model |
| `WordDocumentTreeWriter` | Writes the tree model back to .docx with formatting preservation |

### DocumentNode Properties

| Property | Type | Description |
|----------|------|-------------|
| `Id` | `string` | Unique identifier |
| `Type` | `ContentType` | Document, Heading, Paragraph, Table, Image, List, ListItem, HyperlinkText, TextRun, ContentControl |
| `HeadingLevel` | `int` | 1-9 for headings, 0 for other types |
| `Text` | `string` | Plain text content |
| `Children` | `List<DocumentNode>` | Child nodes |
| `Parent` | `DocumentNode?` | Parent node |
| `Runs` | `List<FormattedRun>` | Formatted text runs with styling |
| `ParagraphFormatting` | `ParagraphFormatting?` | Paragraph-level formatting |
| `ContentControlProperties` | `ContentControlProperties?` | Content control metadata |
| `OriginalXml` | `string?` | Original OpenXML for round-trip fidelity |
| `Metadata` | `Dictionary<string, object>` | Additional metadata (e.g., TableData for table nodes) |

### Extension Method Categories

| Extension Class | Methods | Purpose |
|----------------|---------|---------|
| `TableExtensions` | 30+ | Cell access, structural modification, formatting, nested tables |
| `DocumentMergeExtensions` | 15+ | Append, concatenate, extract sections, insert nodes, clone |
| `ContentControlExtensions` | 20+ | Find, update, remove SDT controls |
| `FontExtensions` | 12+ | Paragraph/document/range font changes, font queries |
| `StyleExtensions` | 10+ | Find by style, change styles, style distribution |
| `TreeNavigationExtensions` | 8+ | FindAll, GetPath, GetHeadingPath, siblings, flatten |
| `TreeQueryExtensions` | 10+ | GetAllHeadings, GetAllTables, GetTableOfContents, CountByType |
| `DocumentPropertyExtensions` | 8+ | Property field queries, metadata text extraction |
| `SerializationExtensions` | 3 | SaveToFile, SaveToStream, ToDocxBytes |

## Round-Trip Fidelity

Saving edits a copy of the source package rather than reassembling one from the model, and applies only the changes you actually made. So anything you don't touch is carried over — including parts this library has no model for:

- **Styles**: All paragraph and character styles, theme, font table
- **Formatting**: Bold, italic, underline, fonts, colors, spacing, borders, shading
- **Document Properties**: Core, extended, and custom properties, including their declared types
- **Content Controls**: All SDT types with properties and data binding, including `w14`/`w15`/`w16` extension elements such as checkbox state
- **Dynamic References**: DOCPROPERTY fields, TOC, BIBLIOGRAPHY, CITATION
- **Structure**: Headers, footers and the relationships they own, sections, page layout, numbering definitions
- **Media**: Images with dimensions, alt text, and positioning
- **Tables**: Cell merging, borders, shading, column widths, nested tables — structural modifications (add/insert/remove rows and columns) preserve the original table style
- **Hyperlinks**: External URLs and internal anchors with relationship preservation
- **Glossary**: Building blocks, Quick Parts, custom XML parts
- **Everything else**: Comments, tracked changes, embedded objects, charts, ink, and any other part

Documents you build in code have no source package, so they contain only what the model represents. See [Round-Trip Fidelity](docs/articles/round-trip.md) for the full boundaries, failure behaviour, and concurrency rules.

### Untrusted Input

A `.docx` is a zip archive whose parts decompress to a size its author chooses. Parsing applies no bounds by default; pass `DocumentLimits.Untrusted` for documents that didn't come from you:

```csharp
using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
var doc = parser.ParseFromFile(uploadedPath);
```

### Failure Behaviour

A part that can't be preserved throws `DocumentPreservationException` rather than being dropped silently, so a lossy result is never reported as a faithful one. Set `RecoveryOptions.ContinueOnPreservationFailure` to salvage what you can and inspect `Diagnostics` afterwards. Saving is atomic — the package is built in memory and staged beside the destination, so a failure leaves any existing file intact.

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
│   │   ├── ContentType.cs
│   │   ├── TrackedModel.cs               # Records which properties a caller assigned
│   │   ├── OoxmlEnum.cs                  # Typed enum <-> OOXML token conversion
│   │   ├── DocumentLimits.cs             # Bounds for untrusted input
│   │   └── DocumentPreservationException.cs
│   ├── Models/
│   │   ├── Formatting/                   # RunFormatting, ParagraphFormatting, TableFormatting, etc.
│   │   ├── ContentControls/              # ContentControlProperties, ContentControlType
│   │   ├── Tables/                       # TableData, TableRow, TableCell
│   │   ├── Images/                       # ImageData
│   │   └── Package/                      # CoreProperties, ExtendedProperties, DocumentPackageData
│   ├── Extensions/
│   │   ├── TableExtensions.cs            # Table access, modification, formatting
│   │   ├── DocumentMergeExtensions.cs    # Document merging and section operations
│   │   ├── ContentControlExtensions.cs   # SDT find, update, remove
│   │   ├── FontExtensions.cs             # Font management
│   │   ├── StyleExtensions.cs            # Style queries and changes
│   │   ├── TreeNavigationExtensions.cs   # FindAll, GetPath, siblings, flatten
│   │   ├── TreeQueryExtensions.cs        # GetAllHeadings, GetAllTables, CountByType
│   │   ├── DocumentPropertyExtensions.cs # Document property field operations
│   │   └── SerializationExtensions.cs    # SaveToFile, SaveToStream, ToDocxBytes
│   ├── Parsing/
│   │   ├── ParsingContext.cs             # Style cache, hyperlink resolution, shared media buffers
│   │   └── Extractors/                   # FormattingExtractor, ImageExtractor, TableExtractor
│   ├── Writing/
│   │   ├── DocumentWriteSession.cs       # Per-write state; edits a copy of the source package
│   │   ├── ParagraphEditor.cs            # Applies node edits onto original paragraph XML
│   │   ├── RunEditor.cs                  # Maps edited runs onto existing runs by offset
│   │   ├── TableEditor.cs                # Applies table edits onto original table XML
│   │   ├── TableBuilder.cs               # Builds tables created in code
│   │   ├── DefaultParts.cs               # Styles and numbering for new documents
│   │   └── OoxmlOrder.cs                 # Schema-ordered property insertion
│   ├── WordDocument.cs                   # Main document wrapper
│   ├── DocumentNode.cs                   # Tree node
│   ├── DocumentPropertyHelpers.cs        # Property name/type utilities
│   ├── WordDocumentTreeParser.cs         # Parser with heading hierarchy inference
│   └── WordDocumentTreeWriter.cs         # Writer facade
│
├── WordDocumentParser.Tests/             # Regression tests (xUnit)
│
└── WordDocumentParser.Demo/              # Demo application
    ├── Program.cs
    └── Features/
        ├── Tables/                       # TableParsing, TableModificationDemo
        ├── Concatenation/                # DocumentConcatenationDemo
        ├── ContentControls/              # ContentControlsDemo, ContentControlRemovalDemo
        ├── DocumentCreation/             # DocumentCreationDemo, TableHelper
        ├── DocumentProperties/           # DocumentPropertyDemo
        ├── Examples/                     # ExampleUsageDemo
        ├── Fonts/                        # FontDemo
        ├── Parsing/                      # DocumentStructureDemo
        ├── RoundTrip/                    # RoundTripDemo, DocumentComparison, DocumentValidator
        └── Styles/                       # ParagraphStyleDemo
```

## Building and Testing

```bash
# Build the entire solution
dotnet build

# Build only the library
dotnet build WordDocumentParser/WordDocumentParser.csproj

# Run the regression tests
dotnet test

# Run the demo against a document (defaults to SampleDocument.docx)
dotnet run --project WordDocumentParser.Demo -- path/to/document.docx
```

CI builds with `-warnaserror`. `GenerateDocumentationFile` is on, so a public member without XML documentation fails the build.

## License

See [LICENSE](LICENSE) for details.
