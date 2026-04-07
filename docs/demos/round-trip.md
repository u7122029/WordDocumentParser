# Round-Trip Demo

**Source:** `WordDocumentParser.Demo/Features/RoundTrip/RoundTripDemo.cs`

Demonstrates round-trip fidelity: parsing a document, writing it back, validating the output, and comparing the original with the copy.

## What It Does

1. Parses a `.docx` file into the document tree
2. Saves it to a new file with `SaveToFile()`
3. Validates the output using `DocumentValidator.ValidateAndReport()`
4. Compares the original and copy with `DocumentComparison.CompareDocuments()`

## Supporting Classes

### DocumentValidator

Uses the OpenXML SDK's `OpenXmlValidator` to check the output file for schema errors.

### DocumentComparison

Re-parses both files and compares node counts, heading structure, table dimensions, and other structural properties.

## Key APIs Used

```csharp
// Parse
using var parser = new WordDocumentTreeParser();
var doc = parser.ParseFromFile(inputPath);

// Write
doc.SaveToFile(outputPath);

// Validate
DocumentValidator.ValidateAndReport(outputPath);

// Compare
DocumentComparison.CompareDocuments(inputPath, outputPath);
```

## Related

- [Round-Trip Fidelity](../articles/round-trip.md)
