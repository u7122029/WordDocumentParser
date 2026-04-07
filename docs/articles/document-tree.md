# Document Tree Structure

The parser organizes content hierarchically based on heading levels. OpenXML stores body elements as a flat list; this library infers the hierarchy from heading levels to produce a navigable tree.

## Tree Layout

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

This enables operations impossible with flat OpenXML:

- `GetSection("Methods")` — retrieve a heading and all its children
- `GetHeadingPath()` — breadcrumb trail like `"Doc > Chapter 1 > Section 1.1"`
- `GetTableOfContents()` — structured TOC entries
- `Parent` — upward navigation from any node

> [!NOTE]
> Content that appears before the first heading is attached directly to the root @WordDocumentParser.DocumentNode as children.

## Core Types

### WordDocument

@WordDocumentParser.WordDocument is the primary entry point. It holds the tree root, the @WordDocumentParser.Models.Package.DocumentPackageData (images, hyperlinks, custom XML), and provides dictionary-style property access:

```csharp
var doc = parser.ParseFromFile("document.docx");
doc["Title"] = "My Report";       // sets a core property
string? author = doc["Author"];   // reads a core property
```

### DocumentNode

@WordDocumentParser.DocumentNode represents a single element in the tree. Key properties:

| Property | Type | Description |
|----------|------|-------------|
| `Id` | `string` | Unique identifier |
| `Type` | @WordDocumentParser.Core.ContentType | Document, Heading, Paragraph, Table, Image, etc. |
| `HeadingLevel` | `int` | 1–9 for headings, 0 otherwise |
| `Text` | `string` | Plain text content |
| `Children` | `List<DocumentNode>` | Child nodes |
| `Parent` | `DocumentNode?` | Parent node |
| `Runs` | `List<FormattedRun>` | Formatted text runs (see @WordDocumentParser.Models.Formatting.FormattedRun) |
| `ParagraphFormatting` | @WordDocumentParser.Models.Formatting.ParagraphFormatting | Paragraph-level formatting |
| `ContentControlProperties` | @WordDocumentParser.Models.ContentControls.ContentControlProperties | Content control metadata (if this node is an SDT) |
| `OriginalXml` | `string?` | Original OpenXML for [round-trip fidelity](round-trip.md) |
| `Metadata` | `Dictionary<string, object>` | Additional metadata (e.g., @WordDocumentParser.Models.Tables.TableData for table nodes) |

### ContentType Enum

@WordDocumentParser.Core.ContentType defines the type of each node:

- `Document` — Root document node
- `Heading` — Heading paragraph (H1–H9)
- `Paragraph` — Body text paragraph
- `Table` — Table container
- `Image` — Embedded image
- `List` — List container
- `ListItem` — List item
- `HyperlinkText` — Hyperlink text span
- `TextRun` — Inline text run
- `ContentControl` — Structured Document Tag

> [!TIP]
> Use `node.Type` to filter or branch on node kind. For example, `root.FindAll(n => n.Type == ContentType.Table)` returns every table in the document.

## Next Steps

- **[Working with Tables](tables.md)** — 2D cell access, modification, and formatting
- **[Tree Navigation](navigation.md)** — finding nodes, breadcrumbs, flattening
- **[Document Structure Demo](../demos/document-structure.md)** — see the tree in action
