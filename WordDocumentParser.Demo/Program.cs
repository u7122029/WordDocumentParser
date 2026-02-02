using System;
using System.IO;
using System.Linq;
using WordDocumentParser.Demo.Features.Concatenation;
using WordDocumentParser.Demo.Features.ContentControls;
using WordDocumentParser.Demo.Features.DocumentCreation;
using WordDocumentParser.Demo.Features.DocumentProperties;
using WordDocumentParser.Demo.Features.Examples;
using WordDocumentParser.Demo.Features.Parsing;
using WordDocumentParser.Demo.Features.RoundTrip;
using WordDocumentParser.Demo.Features.Styles;
using WordDocumentParser.Demo.Features.Tables;
using WordDocumentParser.Demo.Features.Fonts;

namespace WordDocumentParser.Demo;

/// <summary>
/// Demonstration program showing how to use the Word Document Tree Parser and Writer library.
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        // Demo: Document Concatenation and Section Insertion
        string firstDoc = @"C:\isolated\sgp.docx";
        string secondDoc = @"C:\isolated\action_guide.docx";

        if (File.Exists(firstDoc) && File.Exists(secondDoc))
        {
            // Demo 1: Full document concatenation
            Console.WriteLine("\n\nDocument Concatenation Demo:");
            Console.WriteLine("============================");
            DocumentConcatenationDemo.Run(firstDoc, secondDoc);

            // Demo 2: Section insertion (copy-paste style)
            Console.WriteLine("\n\n");
            DocumentConcatenationDemo.RunSectionInsertion(firstDoc, secondDoc);

            // Demo 3: Extract and insert specific nodes (tables)
            Console.WriteLine("\n\n");
            DocumentConcatenationDemo.RunNodeExtraction(firstDoc, secondDoc);
        }
        else
        {
            if (!File.Exists(firstDoc))
                Console.WriteLine($"First document not found: {firstDoc}");
            if (!File.Exists(secondDoc))
                Console.WriteLine($"Second document not found: {secondDoc}");

            Console.WriteLine("\nWord Document Tree Parser & Writer - Usage Examples");
            Console.WriteLine("===================================================\n");

            // Show example code
            ExampleUsageDemo.Show();

            // Demo: Create a document from scratch
            Console.WriteLine("\n\nCreating Sample Document:");
            Console.WriteLine("=========================");
            DocumentCreationDemo.Run();
        }
    }
}
