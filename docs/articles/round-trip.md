# Round-Trip Fidelity

WordDocumentParser is designed for surgical modification — it preserves everything you don't explicitly change.

## How It Works

1. **Parsing**: @WordDocumentParser.WordDocumentTreeParser reads the `.docx` package, builds the heading-based tree, stores the original OpenXML for each element in `DocumentNode.OriginalXml`, and — unless you turn `CaptureOriginalPackage` off — keeps a copy of the source package itself.
2. **Modification**: You modify nodes through the API. Every assignment is recorded, so the library knows which values you set and which it merely read.
3. **Writing**: @WordDocumentParser.WordDocumentTreeWriter edits a copy of the source package. It replaces the body with the current tree, applies your recorded changes to each element's original XML, and overwrites only those package parts whose value differs from what was parsed.

> [!IMPORTANT]
> Preservation comes from *not taking the document apart*. The writer edits the original package rather than reassembling one from the model, so parts the library has no model for travel through untouched. Reassembling could only ever preserve what the model happened to capture.

## What Is Preserved

Because the source package is edited rather than rebuilt, anything you do not touch is carried over, including parts this library does not model:

- **Styles** — All paragraph and character styles, theme, font table
- **Formatting** — Bold, italic, underline, fonts, colors, spacing, borders, shading
- **Document Properties** — Core, extended, and custom properties, including their declared types
- **Content Controls** — All SDT types with properties and data binding, including `w14`/`w15`/`w16` extension elements such as checkbox state
- **Dynamic References** — DOCPROPERTY fields, TOC, BIBLIOGRAPHY, CITATION
- **Structure** — Headers, footers and the relationships they own, sections, page layout, numbering definitions
- **Media** — Images with dimensions, alt text, and positioning
- **Tables** — Cell merging, borders, shading, column widths, nested tables
- **Hyperlinks** — External URLs and internal anchors with relationship preservation
- **Glossary** — Building blocks, Quick Parts, custom XML parts
- **Everything else** — Comments, tracked changes, embedded objects, charts, ink, and any other part, carried through as-is

## Limits

- **The body is rebuilt from the tree.** Paragraphs, tables, and block-level content controls are written from their nodes. Body-level elements the tree does not model — bookmarks spanning blocks, revision range markers — are carried over and reinserted at their original position, which is exact when the tree's structure is unchanged and best-effort when nodes were added or removed.
- **Documents built in code have no source package.** A @WordDocumentParser.WordDocument you construct yourself is assembled from the model, so only what the model represents ends up in the file.
- **`CaptureOriginalPackage = false`** trades preservation for memory: saving then reassembles from the model alone.
- **Merging is not a package merge.** @WordDocumentParser.Extensions.DocumentMergeExtensions brings the source's body content, images, hyperlinks, and numbering definitions across. The source's styles, theme, headers, and footers are not merged — the target's are used, so content relying on a style the target lacks will render with the target's defaults.

## Failure Behaviour

By default, a part that cannot be preserved throws a @WordDocumentParser.Core.DocumentPreservationException rather than being silently dropped, so a lossy result is never reported as a faithful one. To salvage what you can instead, opt in and read the diagnostics afterwards:

```csharp
var recovery = new RecoveryOptions { ContinueOnPreservationFailure = true };

using var parser = new WordDocumentTreeParser { Recovery = recovery };
var document = parser.ParseFromFile(inputPath);

foreach (var diagnostic in recovery.Diagnostics)
    Console.WriteLine(diagnostic);   // e.g. "rId7: Image part could not be read. (IOException)"
```

Saving is atomic: the package is built completely in memory and staged beside the destination before replacing it, so a failure part way through leaves any existing file intact.

## Untrusted Input

A `.docx` is a zip archive whose parts decompress to a size the author chooses. Pass @WordDocumentParser.Core.DocumentLimits when the document did not come from you:

```csharp
using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
var document = parser.ParseFromFile(uploadedPath);
```

`DocumentLimits.Trusted` — the default — applies no bounds, matching the behaviour appropriate when you control the input files.

## Validation

You can validate output documents using the OpenXML SDK:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

using var validationDoc = WordprocessingDocument.Open(outputPath, false);
var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
var errors = validator.Validate(validationDoc).ToList();

if (errors.Count == 0)
    Console.WriteLine("No validation errors.");
else
    foreach (var error in errors)
        Console.WriteLine($"- [{error.Part?.Uri}] {error.Description}");
```

Documents this library creates from scratch validate cleanly against Office 2019, which the test suite asserts.

## Concurrency

@WordDocumentParser.WordDocumentTreeParser and @WordDocumentParser.WordDocumentTreeWriter hold no state between calls and can be reused sequentially. Neither is safe to share across threads; use one instance per thread. A parsed @WordDocumentParser.WordDocument is likewise not safe to modify from several threads at once.

## See Also

- **[Round-Trip Demo](../demos/round-trip.md)** — parse, write, validate, and compare
- @WordDocumentParser.WordDocumentTreeWriter — writer API reference
