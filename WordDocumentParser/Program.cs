using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordDocumentParser
{
    /// <summary>
    /// Demonstration program showing how to use the Word Document Tree Parser and Writer
    /// </summary>
    class Program
    {
        private const string SourceDirUrl =
            "https://hackage-content.haskell.org/package/pandoc-3.8.3/src/test/docx/";

        private const string DownloadFolder = @"C:\isolated";

        static void Main(string[] args)
        {
            Directory.CreateDirectory(DownloadFolder);

            // Download and process all .docx files from the pandoc test directory
            try
            {
                var downloadedPaths = DownloadAllDocxFromDirectory(SourceDirUrl, DownloadFolder);

                if (downloadedPaths.Count == 0)
                {
                    Console.WriteLine("No .docx files were found to download.");
                    return;
                }

                foreach (var path in downloadedPaths)
                {
                    Console.WriteLine();
                    Console.WriteLine("===================================================");
                    Console.WriteLine($"Processing: {path}");
                    Console.WriteLine("===================================================");

                    ParseAndDisplayDocument(path);

                    Console.WriteLine("\n\nWriting Document Demo:");
                    Console.WriteLine("======================");
                    DemoWriteDocument(path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to download/process documents.");
                Console.WriteLine(ex);

                // Optional fallback: preserve your original behavior
                Console.WriteLine("\n\nWord Document Tree Parser & Writer - Usage Examples");
                Console.WriteLine("===================================================\n");
                ShowExampleUsage();

                Console.WriteLine("\n\nCreating Sample Document:");
                Console.WriteLine("=========================");
                DemoCreateDocument();
            }
        }

        /// <summary>
        /// Downloads every *.docx linked on a simple directory listing page into targetFolder.
        /// Returns full local file paths.
        /// </summary>
        private static List<string> DownloadAllDocxFromDirectory(string directoryUrl, string targetFolder)
        {
            using var http = new HttpClient();

            // Some servers behave better if a User-Agent is present
            http.DefaultRequestHeaders.UserAgent.ParseAdd("WordDocumentParser/1.0 (+https://example.local)");

            Console.WriteLine($"Fetching listing: {directoryUrl}");
            var html = http.GetStringAsync(directoryUrl).GetAwaiter().GetResult();

            // Extract href="something.docx" (also handles single quotes)
            // This is intentionally lightweight to avoid extra dependencies.
            var hrefRegex = new Regex(
                @"href\s*=\s*(['""])(?<href>[^'""]+?\.docx)\1",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            var baseUri = new Uri(directoryUrl, UriKind.Absolute);

            var docxUris = hrefRegex.Matches(html)
                .Select(m => m.Groups["href"].Value)
                .Select(href => new Uri(baseUri, href)) // resolves relative links
                .Distinct()
                .ToList();

            Console.WriteLine($"Found {docxUris.Count} .docx link(s).");

            var downloaded = new List<string>(docxUris.Count);

            foreach (var docxUri in docxUris)
            {
                var fileName = Path.GetFileName(docxUri.LocalPath);

                // Very defensive: ensure a usable filename even if URL is odd
                if (string.IsNullOrWhiteSpace(fileName))
                    fileName = Guid.NewGuid().ToString("N") + ".docx";

                var localPath = Path.Combine(targetFolder, fileName);

                Console.WriteLine($"Downloading: {docxUri} -> {localPath}");

                // Download bytes and write
                var bytes = http.GetByteArrayAsync(docxUri).GetAwaiter().GetResult();
                File.WriteAllBytes(localPath, bytes);

                downloaded.Add(localPath);
            }

            return downloaded;
        }

        /// <summary>
        /// Demonstrates writing a parsed document back to a new file
        /// </summary>
        static void DemoWriteDocument(string inputPath)
        {
            using var parser = new WordDocumentTreeParser();
            var documentTree = parser.ParseFromFile(inputPath);

            var outputPath = Path.Combine(Path.GetDirectoryName(inputPath)!,
                Path.GetFileNameWithoutExtension(inputPath) + "_copy.docx");

            documentTree.SaveToFile(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");

            // Validate the saved document
            ValidateDocument(outputPath);

            // Compare specific elements
            CompareDocuments(inputPath, outputPath);
        }

        /// <summary>
        /// Compares specific elements between original and copy
        /// </summary>
        static void CompareDocuments(string originalPath, string copyPath)
        {
            Console.WriteLine("\nComparing documents...");

            using var origDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(originalPath, false);
            using var copyDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(copyPath, false);

            var origBody = origDoc.MainDocumentPart?.Document?.Body;
            var copyBody = copyDoc.MainDocumentPart?.Document?.Body;

            if (origBody == null || copyBody == null)
            {
                Console.WriteLine("Could not access document bodies");
                return;
            }

            // Count elements
            var origParas = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
            var copyParas = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count();
            Console.WriteLine($"Paragraphs: Original={origParas}, Copy={copyParas}");

            var origTables = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();
            var copyTables = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();
            Console.WriteLine($"Tables: Original={origTables}, Copy={copyTables}");

            // Check for field characters
            var origFldChar = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldChar>().Count();
            var copyFldChar = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldChar>().Count();
            Console.WriteLine($"Field characters: Original={origFldChar}, Copy={copyFldChar}");

            // Check section properties
            var origSectPr = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Count();
            var copySectPr = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Count();
            Console.WriteLine($"Section properties: Original={origSectPr}, Copy={copySectPr}");

            // Check section properties locations
            Console.WriteLine("\nOriginal section properties locations:");
            int idx = 0;
            foreach (var sectPr in origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>())
            {
                var parent = sectPr.Parent;
                var grandParent = parent?.Parent;
                Console.WriteLine($"  {idx++}: Parent={parent?.GetType().Name}, GrandParent={grandParent?.GetType().Name}");
            }

            Console.WriteLine("\nCopy section properties locations:");
            idx = 0;
            foreach (var sectPr in copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>())
            {
                var parent = sectPr.Parent;
                var grandParent = parent?.Parent;
                Console.WriteLine($"  {idx++}: Parent={parent?.GetType().Name}, GrandParent={grandParent?.GetType().Name}");
            }

            // Check if any paragraphs contain sectPr in original
            var parasWithSectPr = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .Where(p => p.ParagraphProperties?.SectionProperties != null)
                .Count();
            Console.WriteLine($"\nParagraphs with section properties in original: {parasWithSectPr}");

            // Check TOC field structure
            Console.WriteLine("\nField structure analysis:");
            var origInstrTexts = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldCode>()
                .Select(f => f.Text?.Trim())
                .Where(t => !string.IsNullOrEmpty(t) && t.Contains("TOC"))
                .ToList();
            var copyInstrTexts = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldCode>()
                .Select(f => f.Text?.Trim())
                .Where(t => !string.IsNullOrEmpty(t) && t.Contains("TOC"))
                .ToList();
            Console.WriteLine($"TOC field instructions in original: {origInstrTexts.Count}");
            Console.WriteLine($"TOC field instructions in copy: {copyInstrTexts.Count}");
            if (origInstrTexts.Count > 0)
                Console.WriteLine($"  Original TOC: {origInstrTexts[0]}");
            if (copyInstrTexts.Count > 0)
                Console.WriteLine($"  Copy TOC: {copyInstrTexts[0]}");

            // Check for field begin/end balance
            var origBegin = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldChar>()
                .Count(f => f.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Begin);
            var origEnd = origBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldChar>()
                .Count(f => f.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.End);
            var copyBegin = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldChar>()
                .Count(f => f.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Begin);
            var copyEnd = copyBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldChar>()
                .Count(f => f.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.End);
            Console.WriteLine($"\nField balance - Original: Begin={origBegin}, End={origEnd}");
            Console.WriteLine($"Field balance - Copy: Begin={copyBegin}, End={copyEnd}");
        }

        /// <summary>
        /// Validates a Word document and reports any errors
        /// </summary>
        static void ValidateDocument(string filePath)
        {
            Console.WriteLine("\nValidating document...");
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false);
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
            var errors = validator.Validate(doc);

            if (!errors.Any())
            {
                Console.WriteLine("Document is valid - no errors found.");
            }
            else
            {
                Console.WriteLine($"Found {errors.Count()} validation errors:");
                foreach (var error in errors.Take(20))
                {
                    Console.WriteLine($"  - {error.Description}");
                    if (error.Node != null)
                    {
                        var xml = error.Node.OuterXml;
                        if (xml.Length > 100) xml = xml.Substring(0, 100) + "...";
                        Console.WriteLine($"    Node: {xml}");
                    }
                }
                if (errors.Count() > 20)
                {
                    Console.WriteLine($"  ... and {errors.Count() - 20} more errors");
                }
            }
        }

        /// <summary>
        /// Demonstrates creating a document from scratch using the tree structure
        /// </summary>
        static void DemoCreateDocument()
        {
            // Create a new document tree programmatically
            var root = new DocumentNode(ContentType.Document, "Sample Document");

            // Add Introduction section
            var intro = new DocumentNode(ContentType.Heading, 1, "Introduction");
            root.AddChild(intro);
            intro.AddChild(new DocumentNode(ContentType.Paragraph,
                "This document was created programmatically using WordDocumentTreeWriter."));
            intro.AddChild(new DocumentNode(ContentType.Paragraph,
                "It demonstrates the ability to generate Word documents from a tree structure."));

            // Add a subsection
            var background = new DocumentNode(ContentType.Heading, 2, "Background");
            intro.AddChild(background);
            background.AddChild(new DocumentNode(ContentType.Paragraph,
                "The document tree structure follows the heading hierarchy of the document."));

            // Add Methods section with a table
            var methods = new DocumentNode(ContentType.Heading, 1, "Methods");
            root.AddChild(methods);
            methods.AddChild(new DocumentNode(ContentType.Paragraph,
                "The following table shows our methodology:"));

            // Create a sample table
            var tableNode = CreateSampleTable();
            methods.AddChild(tableNode);

            // Add Results section with list items
            var results = new DocumentNode(ContentType.Heading, 1, "Results");
            root.AddChild(results);
            results.AddChild(new DocumentNode(ContentType.Paragraph, "Key findings include:"));

            var item1 = new DocumentNode(ContentType.ListItem, "First finding: Improved efficiency");
            item1.Metadata["ListLevel"] = 0;
            item1.Metadata["ListId"] = 1;
            results.AddChild(item1);

            var item2 = new DocumentNode(ContentType.ListItem, "Second finding: Better accuracy");
            item2.Metadata["ListLevel"] = 0;
            item2.Metadata["ListId"] = 1;
            results.AddChild(item2);

            var item3 = new DocumentNode(ContentType.ListItem, "Third finding: Reduced costs");
            item3.Metadata["ListLevel"] = 0;
            item3.Metadata["ListId"] = 1;
            results.AddChild(item3);

            // Add Conclusion
            var conclusion = new DocumentNode(ContentType.Heading, 1, "Conclusion");
            root.AddChild(conclusion);
            conclusion.AddChild(new DocumentNode(ContentType.Paragraph,
                "This demonstrates the round-trip capability of parsing and writing Word documents."));

            // Save the document
            var outputPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.docx");
            root.SaveToFile(outputPath);
            Console.WriteLine($"Sample document created: {outputPath}");

            // Display the tree structure
            Console.WriteLine("\nDocument Tree Structure:");
            Console.WriteLine(root.ToTreeString());
        }

        /// <summary>
        /// Creates a sample table node with data
        /// </summary>
        static DocumentNode CreateSampleTable()
        {
            var tableData = new TableData { ColumnCount = 3 };

            // Header row
            var headerRow = new TableRow { RowIndex = 0, IsHeader = true };
            headerRow.Cells.Add(CreateCell(0, 0, "Step"));
            headerRow.Cells.Add(CreateCell(0, 1, "Description"));
            headerRow.Cells.Add(CreateCell(0, 2, "Duration"));
            tableData.Rows.Add(headerRow);

            // Data rows
            var row1 = new TableRow { RowIndex = 1 };
            row1.Cells.Add(CreateCell(1, 0, "1"));
            row1.Cells.Add(CreateCell(1, 1, "Data Collection"));
            row1.Cells.Add(CreateCell(1, 2, "2 weeks"));
            tableData.Rows.Add(row1);

            var row2 = new TableRow { RowIndex = 2 };
            row2.Cells.Add(CreateCell(2, 0, "2"));
            row2.Cells.Add(CreateCell(2, 1, "Analysis"));
            row2.Cells.Add(CreateCell(2, 2, "3 weeks"));
            tableData.Rows.Add(row2);

            var row3 = new TableRow { RowIndex = 3 };
            row3.Cells.Add(CreateCell(3, 0, "3"));
            row3.Cells.Add(CreateCell(3, 1, "Report Writing"));
            row3.Cells.Add(CreateCell(3, 2, "1 week"));
            tableData.Rows.Add(row3);

            var tableNode = new DocumentNode(ContentType.Table, $"[Table: {tableData.RowCount}x{tableData.ColumnCount}]");
            tableNode.Metadata["TableData"] = tableData;
            tableNode.Metadata["RowCount"] = tableData.RowCount;
            tableNode.Metadata["ColumnCount"] = tableData.ColumnCount;

            return tableNode;
        }

        static TableCell CreateCell(int row, int col, string text)
        {
            var cell = new TableCell { RowIndex = row, ColumnIndex = col };
            cell.Content.Add(new DocumentNode(ContentType.Paragraph, text));
            return cell;
        }

        static void ParseAndDisplayDocument(string filePath)
        {
            Console.WriteLine($"Parsing: {filePath}\n");

            using var parser = new WordDocumentTreeParser();
            var documentTree = parser.ParseFromFile(filePath);

            // Display the tree structure
            Console.WriteLine("Document Tree Structure:");
            Console.WriteLine("========================");
            Console.WriteLine(documentTree.ToTreeString());

            // Display statistics
            Console.WriteLine("\nDocument Statistics:");
            Console.WriteLine("====================");
            var counts = documentTree.CountByType();
            foreach (var kvp in counts.OrderBy(k => k.Key.ToString()))
            {
                Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
            }

            // Display table of contents
            var toc = documentTree.GetTableOfContents();
            if (toc.Any())
            {
                Console.WriteLine("\nTable of Contents:");
                Console.WriteLine("==================");
                foreach (var (level, title, _) in toc)
                {
                    var indent = new string(' ', (level - 1) * 2);
                    Console.WriteLine($"{indent}{level}. {title}");
                }
            }

            // Display tables info
            var tables = documentTree.GetAllTables().ToList();
            if (tables.Any())
            {
                Console.WriteLine($"\nTables Found: {tables.Count}");
                Console.WriteLine("=============");
                int tableNum = 1;
                foreach (var table in tables)
                {
                    var tableData = table.GetTableData();
                    if (tableData != null)
                    {
                        Console.WriteLine($"  Table {tableNum}: {tableData.RowCount} rows × {tableData.ColumnCount} columns");
                        Console.WriteLine($"    Location: {table.GetHeadingPath()}");
                    }
                    tableNum++;
                }
            }

            // Display images info
            var images = documentTree.GetAllImages().ToList();
            if (images.Any())
            {
                Console.WriteLine($"\nImages Found: {images.Count}");
                Console.WriteLine("=============");
                foreach (var image in images)
                {
                    var imageData = image.GetImageData();
                    if (imageData != null)
                    {
                        Console.WriteLine($"  - {imageData.Name}: {imageData.WidthInches:F1}\" × {imageData.HeightInches:F1}\" ({imageData.ContentType})");
                        Console.WriteLine($"    Location: {image.GetHeadingPath()}");
                    }
                }
            }
        }

        static void ShowExampleUsage()
        {
            Console.WriteLine(@"
// Basic Usage - Parse a Word document
using var parser = new WordDocumentTreeParser();
var documentTree = parser.ParseFromFile(""document.docx"");

// The tree structure follows heading hierarchy:
// Document (root)
//   ├── H1: Introduction
//   │     ├── Paragraph: Some text...
//   │     ├── H2: Background
//   │     │     ├── Paragraph: More text...
//   │     │     └── Table: [Table: 3x4]
//   │     └── H2: Purpose
//   │           └── Paragraph: Purpose text...
//   └── H1: Methods
//         ├── H2: Data Collection
//         │     └── Image: [Image: figure1.png]
//         └── H2: Analysis
//               └── Paragraph: Analysis details...

// Print the tree structure
Console.WriteLine(documentTree.ToTreeString());

// Find all headings
var allHeadings = documentTree.GetAllHeadings();
foreach (var heading in allHeadings)
{
    Console.WriteLine($""H{heading.HeadingLevel}: {heading.Text}"");
}

// Get specific heading level
var h2Headings = documentTree.GetHeadingsAtLevel(2);

// Find a section by name
var methodsSection = documentTree.GetSection(""Methods"");
if (methodsSection != null)
{
    // Get all text under this section
    var sectionText = methodsSection.GetAllText();
    Console.WriteLine($""Methods section content:\n{sectionText}"");
}

// Work with tables
var tables = documentTree.GetAllTables();
foreach (var tableNode in tables)
{
    var tableData = tableNode.GetTableData();
    if (tableData != null)
    {
        // Access as 2D array
        var array = tableData.ToTextArray();
        Console.WriteLine($""Cell [0,0]: {array[0, 0]}"");
        
        // Or iterate rows/cells
        foreach (var row in tableData.Rows)
        {
            foreach (var cell in row.Cells)
            {
                Console.WriteLine($""Row {cell.RowIndex}, Col {cell.ColumnIndex}: {cell.TextContent}"");
            }
        }
    }
}

// Work with images
var images = documentTree.GetAllImages();
foreach (var imageNode in images)
{
    var imageData = imageNode.GetImageData();
    if (imageData?.Data != null)
    {
        // Save image to file
        File.WriteAllBytes($""{imageData.Name}"", imageData.Data);
    }
}

// Navigation
var node = documentTree.FindFirst(n => n.Text.Contains(""specific text""));
if (node != null)
{
    // Get breadcrumb path
    var path = node.GetHeadingPath(); // ""Document > Introduction > Background""
    
    // Navigate to parent
    var parent = node.Parent;
    
    // Get siblings
    var siblings = node.GetSiblings();
    var nextSibling = node.GetNextSibling();
}

// Custom search
var paragraphsWithLinks = documentTree.FindAll(n => 
    n.Type == ContentType.Paragraph && 
    n.Metadata.ContainsKey(""HasHyperlinks""));

// Statistics
var stats = documentTree.CountByType();
Console.WriteLine($""Total paragraphs: {stats[ContentType.Paragraph]}"");
Console.WriteLine($""Total tables: {stats[ContentType.Table]}"");

// ============================================
// WRITING DOCUMENTS
// ============================================

// Save a parsed document back to a new file
documentTree.SaveToFile(""output.docx"");

// Or save to a stream
using var stream = new MemoryStream();
documentTree.SaveToStream(stream);

// Or get as byte array
var bytes = documentTree.ToDocxBytes();

// Create a new document from scratch
var newDoc = new DocumentNode(ContentType.Document, ""My Document"");

// Add a heading
var heading = new DocumentNode(ContentType.Heading, 1, ""Chapter 1"");
newDoc.AddChild(heading);

// Add paragraphs under the heading
heading.AddChild(new DocumentNode(ContentType.Paragraph, ""This is the first paragraph.""));
heading.AddChild(new DocumentNode(ContentType.Paragraph, ""This is the second paragraph.""));

// Add a sub-heading
var subHeading = new DocumentNode(ContentType.Heading, 2, ""Section 1.1"");
heading.AddChild(subHeading);
subHeading.AddChild(new DocumentNode(ContentType.Paragraph, ""Content under section 1.1""));

// Add list items
var listItem = new DocumentNode(ContentType.ListItem, ""First bullet point"");
listItem.Metadata[""ListLevel""] = 0;
listItem.Metadata[""ListId""] = 1;
heading.AddChild(listItem);

// Create a table
var tableData = new TableData { ColumnCount = 2 };
var row = new TableRow { RowIndex = 0 };
row.Cells.Add(new TableCell { RowIndex = 0, ColumnIndex = 0, Content = { new DocumentNode(ContentType.Paragraph, ""Cell 1"") } });
row.Cells.Add(new TableCell { RowIndex = 0, ColumnIndex = 1, Content = { new DocumentNode(ContentType.Paragraph, ""Cell 2"") } });
tableData.Rows.Add(row);

var table = new DocumentNode(ContentType.Table, ""[Table]"");
table.Metadata[""TableData""] = tableData;
heading.AddChild(table);

// Save the new document
newDoc.SaveToFile(""new_document.docx"");
");
        }
    }
}
