# Style Management

Paragraph styles control the visual appearance of text in Word. @WordDocumentParser.Extensions.StyleExtensions lets you query, change, and bulk-replace styles across the document.

## Finding Nodes by Style

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

var normalParas = doc.FindByStyle("Normal");
var headings = doc.FindByStyles("Heading1", "Heading2");
```

## Changing Styles

```csharp
// Change a single node's style
node.ChangeStyle("Heading2");  // Also updates HeadingLevel and Type

// Bulk replace one style with another across the entire document
doc.ChangeStyleBulk("OldStyle", "NewStyle");

// Conditional style change
doc.ChangeStyleWhere(n => n.Text.StartsWith("Note:"), "NoteStyle");
```

> [!NOTE]
> `ChangeStyle` on a heading node updates both the `HeadingLevel` and `Type` properties to match the new style. For example, changing from `Heading1` to `Heading2` sets `HeadingLevel = 2` and may reparent the node in the tree.

## Querying Style Usage

```csharp
// Style distribution: style name -> paragraph count
var distribution = doc.GetStyleDistribution();

// Check a node's style
bool isHeading = node.HasStyle("Heading1");
string? style = node.GetStyle();
```

## Next Steps

- **[Font Management](fonts.md)** — per-run and per-paragraph font control (distinct from styles)
- **[Paragraph Style Demo](../demos/paragraph-styles.md)** — full walkthrough of style operations
- @WordDocumentParser.Extensions.StyleExtensions — API reference
