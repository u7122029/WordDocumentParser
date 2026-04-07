# Getting Started

## Requirements

- .NET 9.0 or later
- [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) 3.4.1

## Installation

Add the project reference to your `.csproj`:

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

## Parsing a Document

Use @WordDocumentParser.WordDocumentTreeParser to parse a `.docx` file into a @WordDocumentParser.WordDocument, which holds the heading-based tree and document metadata:

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile("document.docx");

// Display the heading-based tree structure
Console.WriteLine(doc.Root.ToTreeString());

// Access document properties with the dictionary-style indexer
Console.WriteLine($"Title: {doc["Title"]}");
Console.WriteLine($"Author: {doc["Author"]}");
```

> [!NOTE]
> `WordDocumentTreeParser` implements `IDisposable` because it holds an open handle to the `.docx` package during parsing. Always wrap it in a `using` statement.

## Writing a Document

Save a parsed (and optionally modified) document. The library only regenerates elements you changed; everything else passes through as the original XML:

```csharp
using WordDocumentParser.Extensions;

doc.SaveToFile("output.docx");

// Or save to a stream / byte array
using var stream = new MemoryStream();
doc.SaveToStream(stream);
byte[] bytes = doc.ToDocxBytes();
```

## Creating a Document from Scratch

You can build a @WordDocumentParser.DocumentNode tree programmatically and save it without parsing an existing file:

```csharp
using WordDocumentParser;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;

var root = new DocumentNode(ContentType.Document, "My Document");

var heading = new DocumentNode(ContentType.Heading, 1, "Introduction");
root.AddChild(heading);
heading.AddChild(new DocumentNode(ContentType.Paragraph, "First paragraph."));
heading.AddChild(new DocumentNode(ContentType.Paragraph, "Second paragraph."));

var doc = new WordDocument(root);
doc["Title"] = "My New Document";
doc.SaveToFile("new_document.docx");
```

> [!TIP]
> See the [Document Creation Demo](../demos/document-creation.md) for a fuller example that includes tables, list items, and nested headings.

## Building the Solution

```bash
# Build the entire solution
dotnet build

# Build only the library
dotnet build WordDocumentParser/WordDocumentParser.csproj

# Run the demo project
dotnet run --project WordDocumentParser.Demo
```

## Next Steps

- **[Document Tree Structure](document-tree.md)** — understand the heading-based hierarchy and core types
- **[Working with Tables](tables.md)** — 2D access, modification, and formatting
- **[API Reference](../api/index.md)** — browse all classes and methods
