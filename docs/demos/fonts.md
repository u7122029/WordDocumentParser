# Font Demo

**Source:** `WordDocumentParser.Demo/Features/Fonts/FontDemo.cs`

Demonstrates all font management operations: per-run, per-paragraph, per-text-span, per-range, global replacement, and querying.

## What It Does

1. Lists all fonts currently used with `GetAllFontsUsed()`
2. Changes the font of a single run with `SetFont()`
3. Changes the font of an entire paragraph with `SetParagraphFont()`
4. Changes the font of a specific word with `SetFontForText()`
5. Changes the font of a character range with `SetFontForRange()`
6. Replaces a font globally with `ReplaceFont()`
7. Sets all headings to a specific font
8. Displays updated font usage
9. Saves, validates with OpenXML SDK, and re-parses to verify

## Key APIs Used

```csharp
// Query
var fonts = doc.GetAllFontsUsed();
var paraFonts = node.GetFontsUsed();
string? font = run.GetFont();

// Per-run
run.SetFont("Cascadia Code");

// Per-paragraph
node.SetParagraphFont("Arial");

// Per-text-span (finds matching text within the paragraph)
int modified = node.SetFontForText("important", "Comic Sans MS");

// Per-character-range
bool success = node.SetFontForRange(0, 10, "Georgia");

// Global replacement
int replaced = doc.ReplaceFont("Times New Roman", "Consolas");

// All headings
foreach (var heading in doc.Root.FindAll(n => n.Type == ContentType.Heading))
    heading.SetParagraphFont("Trebuchet MS");
```

## Related

- [Font Management](../articles/fonts.md)
