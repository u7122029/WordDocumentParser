# Document Properties

@WordDocumentParser.WordDocument provides uniform dictionary-style access to all three types of document properties: core (title, author, etc.), extended (company, manager, etc.), and custom (user-defined key-value pairs).

## Dictionary-Style Access

All property types are accessible through a single indexer, which is case-insensitive:

```csharp
using WordDocumentParser;

// Core property
doc["Title"] = "Annual Report";

// Extended property
doc["Company"] = "ACME Corporation";

// Custom (auto-created if the name is not a known core/extended property)
doc["ProjectCode"] = "PRJ-2024-001";
```

> [!NOTE]
> The indexer automatically routes to the correct property store. Known names like `Title`, `Author`, `Subject` go to @WordDocumentParser.Models.Package.CoreProperties. Names like `Company`, `Manager` go to @WordDocumentParser.Models.Package.ExtendedProperties. Everything else becomes a custom property.

## Method-Based Access

```csharp
string? company = doc.GetProperty("Company");
doc.SetProperty("Department", "Engineering");
bool exists = doc.HasProperty("Keywords");
doc.RemoveProperty("OldField");
```

## Listing All Properties

```csharp
foreach (var (name, value) in doc.GetAllProperties())
    Console.WriteLine($"{name}: {value}");
```

## Deleting Properties

Properties can be deleted by setting them to `null` via the indexer or by calling `RemoveProperty`:

```csharp
doc["Department"] = null;           // Delete via indexer
doc.RemoveProperty("ReviewStatus"); // Delete via method
```

## Custom Properties

Custom properties are stored separately and can be accessed directly:

```csharp
foreach (var (name, value) in doc.CustomProperties)
    Console.WriteLine($"{name} = {value}");
```

## DOCPROPERTY Fields

The library detects and preserves `DOCPROPERTY` field codes in the document. These are dynamic references that Word resolves at print/update time. Use @WordDocumentParser.Extensions.DocumentPropertyExtensions to query nodes with document property fields and extract their metadata.

## Next Steps

- **[Style Management](styles.md)** — query, change, and bulk-replace paragraph styles
- **[Document Properties Demo](../demos/document-properties.md)** — full walkthrough
- @WordDocumentParser.Extensions.DocumentPropertyExtensions — API reference
