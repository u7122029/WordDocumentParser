---
_layout: landing
---

# WordDocumentParser

A .NET library for parsing Word documents (.docx) into a hierarchical tree structure and writing them back with full round-trip fidelity.

Unlike raw OpenXML, this library provides a heading-based document tree, 2D table access, document merging with resource remapping, and surgical modification that preserves formatting you didn't touch.

## Get Started in 30 Seconds

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

// Parse any .docx into a heading-based tree
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile("report.docx");

// Navigate, modify, save — formatting is preserved automatically
Console.WriteLine(doc.Root.ToTreeString());
doc["Title"] = "Updated Report";
doc.SaveToFile("output.docx");
```

> [!TIP]
> New here? Start with the [Getting Started](articles/getting-started.md) guide, then explore the feature articles.

## Key Features

| Feature | What it does |
|---------|-------------|
| [Tree-based parsing](articles/document-tree.md) | Organizes flat OpenXML into a heading-level hierarchy with parent-child navigation |
| [Round-trip fidelity](articles/round-trip.md) | Stores original XML per node; only regenerates what you modify |
| [Table manipulation](articles/tables.md) | 2D cell access, add/insert/remove rows and columns, nested tables |
| [Document merging](articles/merging.md) | Append, concatenate, extract sections with automatic resource remapping |
| [Content controls](articles/content-controls.md) | Find, update, and remove SDTs by tag, alias, ID, or type |
| [Style & font management](articles/styles.md) | Bulk style changes, font replacement, per-run font control |
| [Document properties](articles/properties.md) | Uniform dictionary-style access to core, extended, and custom properties |

## Explore

- **[Articles](articles/index.md)** — Conceptual guides for each feature area
- **[Demo Scripts](demos/index.md)** — Runnable examples from the `WordDocumentParser.Demo` project
- **[API Reference](api/index.md)** — Full API docs generated from XML comments
