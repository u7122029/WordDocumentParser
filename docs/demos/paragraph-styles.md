# Paragraph Styles Demo

**Source:** `WordDocumentParser.Demo/Features/Styles/ParagraphStyleDemo.cs`

Demonstrates querying and modifying paragraph styles: viewing style distribution, finding nodes by style, changing individual styles, and bulk-replacing styles.

## What It Does

1. Displays style distribution with `GetStyleDistribution()`
2. Finds paragraphs by style with `FindByStyle()`
3. Changes a single node's style with `ChangeStyle()` (Heading1 to Heading2)
4. Changes individual paragraphs' styles (BodyText to Quote)
5. Bulk-replaces a style across the document with `ChangeStyleBulk()` (Caption to Subtitle)
6. Displays updated style distribution
7. Saves and verifies with `HasStyle()` and `FindByStyle()`

## Key APIs Used

```csharp
// Query
var distribution = doc.GetStyleDistribution();
var heading1s = doc.FindByStyle("Heading1");
string? style = node.GetStyle();
bool isH1 = node.HasStyle("Heading1");

// Change
node.ChangeStyle("Heading2");

// Bulk replace
int changed = doc.ChangeStyleBulk("Caption", "Subtitle");
```

## Related

- [Style Management](../articles/styles.md)
