# Demo Scripts

The `WordDocumentParser.Demo` project contains runnable examples demonstrating every major feature of the library. Each demo is a self-contained static class in `WordDocumentParser.Demo/Features/`.

## Running Demos

The demo project's `Program.cs` calls one demo at a time. Edit it to run the demo you want:

```csharp
// In Program.cs, call the demo you want to run:
DocumentStructureDemo.Run(inputDoc);
TableParsing.Run(inputDoc);
TableModificationDemo.Run(inputDoc);
DocumentCreationDemo.Run();
DocumentConcatenationDemo.Run(firstDoc, secondDoc);
ContentControlsDemo.Run(inputDoc);
ContentControlRemovalDemo.Run(inputDoc);
DocumentPropertyDemo.Run(inputDoc);
ParagraphStyleDemo.Run(inputDoc);
FontDemo.Run(inputDoc);
RoundTripDemo.Run(inputDoc);
```

Then run:

```bash
dotnet run --project WordDocumentParser.Demo
```

## Available Demos

| Demo | Description |
|------|-------------|
| [Document Structure](document-structure.md) | Parse a document and display its tree, statistics, TOC, tables, and images |
| [Table Parsing](table-parsing.md) | 2D cell access, iteration, modification, formatting, and nested tables |
| [Table Modification](table-modification.md) | Add/insert/remove rows and columns, append and remove cell text |
| [Document Creation](document-creation.md) | Build a document from scratch using the tree API |
| [Document Concatenation](document-concatenation.md) | Append, concatenate, and insert sections across documents |
| [Content Controls](content-controls.md) | Find, inspect, and modify SDT controls (dropdowns, dates, checkboxes) |
| [Content Control Removal](content-control-removal.md) | Strip content controls while preserving text content |
| [Document Properties](document-properties.md) | Get, set, and delete core, extended, and custom properties |
| [Paragraph Styles](paragraph-styles.md) | Query style distribution, change styles, bulk replace |
| [Fonts](fonts.md) | Per-run, per-paragraph, per-range, and global font operations |
| [Round-Trip](round-trip.md) | Parse, write, validate, and compare documents |
