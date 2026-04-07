# Content Controls Demo

**Source:** `WordDocumentParser.Demo/Features/ContentControls/ContentControlsDemo.cs`

Demonstrates finding, inspecting, and modifying content controls (Structured Document Tags) in a Word document. Covers both block-level and inline controls.

## What It Does

1. Finds all content controls with `GetAllContentControls()`
2. Displays detailed properties for each control: type, ID, tag, alias, value, list items, date format, checkbox state, lock settings
3. Shows text with metadata annotations using `GetTextWithMetadata()`
4. Modifies control values based on type:
   - **Dropdown/ComboBox**: selects a different list item
   - **Date**: sets to the current date
   - **Checkbox**: toggles the checked state
   - **PlainText/RichText**: prepends "Modified: "
5. Saves the document and re-parses to verify

## Key APIs Used

```csharp
// Find controls
var controls = doc.GetAllContentControls();

// Inspect properties
var props = node.ContentControlProperties;
// props.Type, props.Id, props.Tag, props.Alias, props.Value
// props.ListItems, props.DateFormat, props.DateValue, props.IsChecked

// Inline controls within a paragraph
var inlineProps = node.GetInlineContentControlProperties();

// Text with metadata
string annotated = node.GetTextWithMetadata();

// Modify block-level control
node.Text = newValue;
props.Value = newValue;
node.Runs.Clear();
node.Runs.Add(new FormattedRun(newValue, formatting));

// Modify inline control
foreach (var run in node.Runs.Where(r => r.ContentControlProperties == props))
    run.Text = newValue;
props.Value = newValue;
```

## Related

- [Content Controls](../articles/content-controls.md)
- [Content Control Removal Demo](content-control-removal.md)
