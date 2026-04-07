# Document Creation Demo

**Source:** `WordDocumentParser.Demo/Features/DocumentCreation/DocumentCreationDemo.cs`

Demonstrates creating a complete Word document from scratch using the tree API, including headings, paragraphs, a table, and list items.

## What It Does

1. Creates a `DocumentNode` tree with headings at multiple levels
2. Adds paragraphs under each heading
3. Creates a table using a helper and adds it under a heading
4. Adds numbered list items with metadata for list level and ID
5. Wraps the tree in a `WordDocument` and saves to `.docx`
6. Displays the resulting tree structure

## Key APIs Used

```csharp
// Build the tree
var root = new DocumentNode(ContentType.Document, "Sample Document");

var intro = new DocumentNode(ContentType.Heading, 1, "Introduction");
root.AddChild(intro);
intro.AddChild(new DocumentNode(ContentType.Paragraph, "Some text."));

var background = new DocumentNode(ContentType.Heading, 2, "Background");
intro.AddChild(background);

// Add a table
var tableNode = TableHelper.CreateSampleTable();
methods.AddChild(tableNode);

// Add list items
var item = new DocumentNode(ContentType.ListItem, "First finding");
item.Metadata["ListLevel"] = 0;
item.Metadata["ListId"] = 1;
results.AddChild(item);

// Save
var document = new WordDocument(root);
document.SaveToFile("SampleDocument.docx");
```

## Related

- [Getting Started](../articles/getting-started.md)
- [Document Tree Structure](../articles/document-tree.md)
