# Font Management

Font management controls the typeface used to render text, which is distinct from paragraph styles. @WordDocumentParser.Extensions.FontExtensions supports per-run, per-paragraph, per-range, and document-wide font operations.

> [!NOTE]
> Fonts and styles are independent. Changing a paragraph's style (e.g., `Heading1`) may imply a default font, but explicitly setting a font with these methods overrides the style's font for the targeted runs.

## Setting Fonts

```csharp
using WordDocumentParser;
using WordDocumentParser.Extensions;

// Set font for an entire paragraph (all runs)
node.SetParagraphFont("Calibri");

// Set font for the entire document
doc.SetDocumentFont("Arial");

// Target specific text within a paragraph
node.SetFontForText("important", "Arial Black");

// Target a character range (start index, length)
node.SetFontForRange(0, 5, "Courier New");
```

## Replacing Fonts

Replace all occurrences of a font globally across the document:

```csharp
int replaced = doc.ReplaceFont("Times New Roman", "Calibri");
Console.WriteLine($"Replaced {replaced} run(s)");
```

## Querying Fonts

```csharp
// All fonts used across the document
var fonts = doc.GetAllFontsUsed();

// Fonts used in a specific paragraph
var paraFonts = node.GetFontsUsed();

// Font of a specific run
string? font = run.GetFont();
```

## Per-Run Font Control

Each @WordDocumentParser.Models.Formatting.FormattedRun carries its own @WordDocumentParser.Models.Formatting.RunFormatting, including font. You can read or set the font on individual runs:

```csharp
var firstRun = node.Runs[0];
Console.WriteLine(firstRun.GetFont());  // Current font
firstRun.SetFont("Cascadia Code");      // Change it
```

> [!TIP]
> When you call `SetFontForText`, the library automatically splits runs at word boundaries so only the matched text gets the new font. The surrounding text keeps its original formatting.

## Next Steps

- **[Tree Navigation](navigation.md)** — finding nodes, breadcrumbs, and queries
- **[Font Demo](../demos/fonts.md)** — full walkthrough of all font operations
- @WordDocumentParser.Extensions.FontExtensions — API reference
