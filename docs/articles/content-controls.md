# Content Controls

Content controls (Structured Document Tags / SDTs) are interactive form fields in Word documents — dropdowns, date pickers, checkboxes, text inputs, and more. @WordDocumentParser.Extensions.ContentControlExtensions provides full access to find, read, update, and remove them.

## Finding Controls

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.ContentControls;

var all = doc.GetAllContentControls();

// Find by tag, alias, ID, or type
var byTag = doc.FindContentControlByTag("ClientName");
var byAlias = doc.FindContentControlByAlias("Document Date");
var byId = doc.FindContentControlById(12345);
var byType = doc.GetContentControlsByType(ContentControlType.Date);
```

## Updating Values

```csharp
doc.SetContentControlValueByTag("ClientName", "ABC Corporation");
doc.SetContentControlValueByAlias("ProjectCode", "PRJ-2024-001");
```

## Removing Controls

When a content control is removed, its text content is preserved — only the control wrapper is stripped:

```csharp
doc.RemoveContentControlByTag("TemporaryField");
doc.RemoveAllContentControls();
```

> [!TIP]
> Removing a content control does not delete any text. The underlying paragraph text remains in the document. This is useful for "flattening" template fields after filling them in.

## Querying Metadata

Each content control has a @WordDocumentParser.Models.ContentControls.ContentControlProperties object with its full metadata:

```csharp
var tags = doc.GetContentControlTags();

var props = doc.GetContentControlPropertiesByTag("FieldTag");
// props.Type — ContentControlType enum
// props.Id — unique integer ID
// props.Tag / props.Alias — user-defined identifiers
// props.Value — current text value
// props.ListItems — dropdown/combobox options (List<ContentControlListItem>)
// props.DateFormat / props.DateValue — for Date controls
// props.IsChecked — for Checkbox controls
```

## Supported Control Types

@WordDocumentParser.Models.ContentControls.ContentControlType defines:

- `PlainText` / `RichText` — text input fields
- `DropDownList` / `ComboBox` — selection lists with @WordDocumentParser.Models.ContentControls.ContentControlListItem entries
- `Date` — date pickers with format strings
- `Checkbox` — checked/unchecked toggle
- `Picture` — image placeholders

## Next Steps

- **[Document Properties](properties.md)** — uniform access to core, extended, and custom properties
- **[Content Controls Demo](../demos/content-controls.md)** — reading and modifying controls
- **[Content Control Removal Demo](../demos/content-control-removal.md)** — stripping controls while preserving text
- @WordDocumentParser.Extensions.ContentControlExtensions — API reference
