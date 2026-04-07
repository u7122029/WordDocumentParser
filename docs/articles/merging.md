# Document Merging

@WordDocumentParser.Extensions.DocumentMergeExtensions provides methods for appending, concatenating, and inserting content across multiple Word documents. Image and hyperlink relationships are automatically remapped to avoid ID conflicts.

## Appending Documents

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

using var parser1 = new WordDocumentTreeParser();
using var parser2 = new WordDocumentTreeParser();
var doc1 = parser1.ParseFromFile("first.docx");
var doc2 = parser2.ParseFromFile("second.docx");

// Append one document to another
doc1.AppendDocument(doc2, addPageBreak: true);

// Append multiple documents
doc1.AppendDocuments(new[] { doc2, doc3 }, addPageBreaks: true);
```

## Concatenating Documents

Create a new @WordDocumentParser.WordDocument from multiple source documents without modifying the originals:

```csharp
var combined = DocumentMergeExtensions.ConcatenateDocuments(
    new[] { doc1, doc2, doc3 });
```

## Section Extraction and Insertion

Extract and transplant sections between documents:

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

## Merge Statistics

Before merging, you can inspect what will be combined using @WordDocumentParser.Extensions.MergeStatistics:

```csharp
var stats = doc1.GetMergeStatistics(doc2);
Console.WriteLine($"Target: {stats.TargetNodeCount} nodes, {stats.TargetImageCount} images");
Console.WriteLine($"Source: {stats.SourceNodeCount} nodes, {stats.SourceImageCount} images");
```

## Resource Remapping

> [!NOTE]
> When merging documents, images and hyperlinks from the source document are remapped to avoid ID conflicts with the target. This happens automatically — you do not need to manage relationship IDs yourself.

## Next Steps

- **[Content Controls](content-controls.md)** — find, update, and remove SDTs
- **[Document Concatenation Demo](../demos/document-concatenation.md)** — full merging walkthrough
- @WordDocumentParser.Extensions.DocumentMergeExtensions — API reference
