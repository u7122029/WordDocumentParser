# Round-Trip Fidelity

WordDocumentParser is designed for surgical modification — it preserves everything you don't explicitly change. Each @WordDocumentParser.DocumentNode stores its `OriginalXml`, and the @WordDocumentParser.WordDocumentTreeWriter only regenerates elements that have been modified.

## What Is Preserved

- **Styles** — All paragraph and character styles, theme, font table
- **Formatting** — Bold, italic, underline, fonts, colors, spacing, borders, shading
- **Document Properties** — Core, extended, and custom properties
- **Content Controls** — All SDT types with properties and data binding
- **Dynamic References** — DOCPROPERTY fields, TOC, BIBLIOGRAPHY, CITATION
- **Structure** — Headers, footers, sections, page layout, numbering definitions
- **Media** — Images with dimensions, alt text, and positioning
- **Tables** — Cell merging, borders, shading, column widths, nested tables; structural modifications preserve the original table style
- **Hyperlinks** — External URLs and internal anchors with relationship preservation
- **Glossary** — Building blocks, Quick Parts, custom XML parts

## How It Works

1. **Parsing**: @WordDocumentParser.WordDocumentTreeParser reads the `.docx` package, builds the heading-based tree, and stores the original OpenXML for each element in `DocumentNode.OriginalXml`.
2. **Modification**: You modify nodes through the API — changing text, formatting, table structure, etc.
3. **Writing**: @WordDocumentParser.WordDocumentTreeWriter walks the tree. For unmodified nodes, it writes back the original XML verbatim. For modified nodes, it regenerates the XML from the node's current state.

> [!IMPORTANT]
> This approach means that formatting, styles, and structural details you didn't touch pass through exactly as they were in the original document. There is no lossy conversion step.

## Validation

You can validate output documents using the OpenXML SDK:

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

using var validationDoc = WordprocessingDocument.Open(outputPath, false);
var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
var errors = validator.Validate(validationDoc).ToList();

if (errors.Count == 0)
    Console.WriteLine("No validation errors.");
else
    foreach (var error in errors)
        Console.WriteLine($"- {error.Description}");
```

## See Also

- **[Round-Trip Demo](../demos/round-trip.md)** — parse, write, validate, and compare
- @WordDocumentParser.WordDocumentTreeWriter — writer API reference
