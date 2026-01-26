using System;
using System.IO;
using WordDocumentParser.Demo.Features.ContentControls;
using WordDocumentParser.Demo.Features.DocumentCreation;
using WordDocumentParser.Demo.Features.DocumentProperties;
using WordDocumentParser.Demo.Features.Examples;
using WordDocumentParser.Demo.Features.Parsing;
using WordDocumentParser.Demo.Features.RoundTrip;

namespace WordDocumentParser.Demo;

/// <summary>
/// Demonstration program showing how to use the Word Document Tree Parser and Writer library.
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        // Example usage with a file path
        string filePath = args.Length > 0 ? args[0] : "C:\\isolated\\content_control_example.docx";

        if (File.Exists(filePath))
        {
            DocumentStructureDemo.Run(filePath);

            // // Demo: Content Controls - Read and Modify
            // Console.WriteLine("\n\nContent Control Demo:");
            // Console.WriteLine("=====================");
            // ContentControlsDemo.Run(filePath);
            //
            // // Demo: Removing Content Controls and Document Properties
            // Console.WriteLine("\n\nRemoving Content Controls Demo:");
            // Console.WriteLine("================================");
            // ContentControlRemovalDemo.Run(filePath);

            // Demo: Add, change, and remove a document property
            Console.WriteLine("\n\nDocument Property Demo:");
            Console.WriteLine("=======================");
            DocumentPropertyDemo.Run(filePath);

            // Demo: Write the parsed document back to a new file
            // Console.WriteLine("\n\nWriting Document Demo:");
            // Console.WriteLine("======================");
            // RoundTripDemo.Run(filePath);
        }
        else
        {
            Console.WriteLine("Word Document Tree Parser & Writer - Usage Examples");
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
