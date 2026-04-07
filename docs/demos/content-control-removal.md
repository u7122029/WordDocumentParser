# Content Control Removal Demo

**Source:** `WordDocumentParser.Demo/Features/ContentControls/ContentControlRemovalDemo.cs`

Demonstrates removing content controls and document property fields from a document while preserving the underlying text content.

## What It Does

1. Displays all current content controls and document property fields
2. Removes a specific control by tag with `RemoveContentControlByTag()`
3. Removes a control by alias with `RemoveContentControlByAlias()`
4. Removes a control directly from a node with `RemoveContentControl()`
5. Removes a document property field with `RemoveDocumentPropertyField()`
6. Removes all remaining controls with `RemoveAllContentControls()`
7. Saves and verifies that text content is preserved without control wrappers

## Key APIs Used

```csharp
// Query
var controls = doc.GetAllContentControls();
var docProps = doc.Root.GetNodesWithDocumentPropertyFields();

// Remove by tag or alias
doc.RemoveContentControlByTag("combobox");
doc.RemoveContentControlByAlias("dropdown");

// Remove directly
node.RemoveContentControl(controlId);

// Remove document property fields
node.RemoveDocumentPropertyField();

// Remove all
int removed = doc.RemoveAllContentControls();
```

## Related

- [Content Controls](../articles/content-controls.md)
- [Content Controls Demo](content-controls.md)
