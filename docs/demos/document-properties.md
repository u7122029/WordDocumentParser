# Document Properties Demo

**Source:** `WordDocumentParser.Demo/Features/DocumentProperties/DocumentPropertyDemo.cs`

Demonstrates the full document properties API: reading, setting, updating, and deleting core, extended, and custom properties.

## What It Does

1. Displays all existing properties with `GetAllProperties()`
2. Reads properties via the `doc["Name"]` indexer
3. Sets built-in properties (Title, Author, Subject, Company)
4. Sets custom properties (any unknown name becomes custom)
5. Accesses custom properties directly via `doc.CustomProperties`
6. Updates an existing custom property
7. Deletes properties via `doc["Name"] = null` and `doc.RemoveProperty()`
8. Saves and re-parses to verify persistence

## Key APIs Used

```csharp
// List all
foreach (var (name, value) in doc.GetAllProperties())
    Console.WriteLine($"{name} = {value}");

// Read
string? title = doc["Title"];

// Set (core, extended, or custom)
doc["Title"] = "Demo Document";
doc["Company"] = "Demo Corp";
doc["ProjectCode"] = "DEMO-001";  // custom

// Check existence
bool exists = doc.HasProperty("Keywords");

// Delete
doc["Department"] = null;
doc.RemoveProperty("ReviewStatus");

// Custom properties
foreach (var (name, value) in doc.CustomProperties) { ... }
```

## Related

- [Document Properties](../articles/properties.md)
