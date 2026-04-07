# Document Concatenation Demo

**Source:** `WordDocumentParser.Demo/Features/Concatenation/DocumentConcatenationDemo.cs`

Demonstrates merging Word documents: appending one to another, concatenating multiple documents, extracting sections, and inserting content across documents.

## What It Does

### `Run(firstDocPath, secondDocPath)`

1. Parses both documents and displays node/image counts
2. Shows pre-merge statistics with `GetMergeStatistics()`
3. Appends the second document with a page break
4. Updates the combined document's title
5. Saves, validates with OpenXML SDK, and re-parses to verify

### `RunMultiple(params string[] paths)`

1. Parses all input documents
2. Creates a new combined document with `ConcatenateDocuments()`
3. Saves the result

### `RunSectionInsertion(targetDocPath, sourceDocPath)`

1. Lists headings in both documents
2. Extracts a section from the source with `ExtractSection()`
3. Inserts the section after a heading in the target with `InsertNodesAfter()`
4. Saves and displays the result

### `RunNodeExtraction(targetDocPath, sourceDocPath)`

1. Extracts all tables from the source with `ExtractTables()`
2. Clones them for the target with `CloneNodesForDocument()`
3. Inserts the cloned tables into the target document

## Key APIs Used

```csharp
// Append
doc1.AppendDocument(doc2, addPageBreak: true);

// Concatenate
var combined = DocumentMergeExtensions.ConcatenateDocuments(documents, addPageBreaks: true);

// Extract and insert
var section = sourceDoc.ExtractSection(headingText, includeNestedHeadings: true);
targetDoc.InsertNodesAfter(insertAfterHeading, section, sourceDoc);

// Clone with resource remapping
var cloned = targetDoc.CloneNodesForDocument(sourceDoc, tables);

// Merge statistics
var stats = doc1.GetMergeStatistics(doc2);
```

## Related

- [Document Merging](../articles/merging.md)
