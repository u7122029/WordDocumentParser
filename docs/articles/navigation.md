# Tree Navigation

The hierarchical tree structure enables powerful navigation and querying operations that are impossible with flat OpenXML. These are provided by @WordDocumentParser.Extensions.TreeNavigationExtensions and @WordDocumentParser.Extensions.TreeQueryExtensions.

## Finding Nodes

```csharp
using WordDocumentParser;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

// Find all nodes matching a predicate (depth-first)
var matches = root.FindAll(n => n.Text.Contains("search term"));

// Find the first match
var first = root.FindFirst(n => n.Type == ContentType.Table);

// Get a section by heading text (returns the heading and all its children)
var section = root.GetSection("Methods");
```

## Navigation

```csharp
// Ancestor chain from root to this node
var path = node.GetPath();

// Breadcrumb string: "Doc > Chapter 1 > Section 1.1"
var breadcrumb = node.GetHeadingPath();

// Sibling navigation
var siblings = node.GetSiblings();
var next = node.GetNextSibling();

// Depth in the tree (root = 0)
var depth = node.GetDepth();

// Flatten to a list (depth-first traversal)
var flat = root.Flatten();
```

> [!TIP]
> `GetHeadingPath()` is useful for reporting a node's location in the document — for example, logging which section a table was found in.

## Queries

```csharp
// Structured table of contents: (Level, Title, Node) tuples
var toc = doc.GetTableOfContents();

// Count nodes by type — returns Dictionary<ContentType, int>
var counts = doc.Root.CountByType();

// Get all text under a node (recursive)
var allText = section.GetAllText();

// Get all headings or all headings at a specific level
var allHeadings = doc.Root.GetAllHeadings();
var h2s = doc.Root.GetHeadingsAtLevel(2);

// Get all tables or images
var tables = doc.Root.GetAllTables();
var images = doc.Root.GetAllImages();
```

## Next Steps

- **[Round-Trip Fidelity](round-trip.md)** — understand what is preserved during save
- **[Document Structure Demo](../demos/document-structure.md)** — parsing and displaying tree structure
- @WordDocumentParser.Extensions.TreeNavigationExtensions — navigation API reference
- @WordDocumentParser.Extensions.TreeQueryExtensions — query API reference
