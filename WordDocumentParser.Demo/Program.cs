using System;
using System.IO;
using WordDocumentParser.Demo.Features.Tables;
using WordDocumentParser.Demo.Features.RoundTrip;
using WordDocumentParser.Demo.Features.ContentControls;
using WordDocumentParser.Demo.Features.DocumentProperties;

namespace WordDocumentParser.Demo;

/// <summary>
/// Demonstration program showing how to use the Word Document Tree Parser and Writer library.
/// </summary>
internal static class Program
{
    private static int Main(string[] args)
    {
        // The document to work on comes from the command line, falling back to the sample checked in
        // beside the solution. It used to be a hardcoded absolute path on one machine.
        var inputDoc = args.Length > 0 ? args[0] : FindSampleDocument();

        if (inputDoc is null)
        {
            Console.Error.WriteLine("Usage: WordDocumentParser.Demo <path-to-document.docx>");
            Console.Error.WriteLine("No path was given and SampleDocument.docx could not be found.");
            return 1;
        }

        if (!File.Exists(inputDoc))
        {
            Console.Error.WriteLine($"Document not found: {inputDoc}");
            return 1;
        }

        Console.WriteLine("Performing ContentControls Demo");
        ContentControlsDemo.Run(inputDoc);

        Console.WriteLine("Performing DocumentProperties Demo");
        DocumentPropertyDemo.Run(inputDoc);

        Console.WriteLine("Performing TableModification Demo");
        TableModificationDemo.Run(inputDoc);

        Console.WriteLine("Performing RoundTrip Demo");
        RoundTripDemo.Run(inputDoc);
        return 0;
    }

    /// <summary>
    /// Looks for the sample document by walking up from the working directory to the repository root.
    /// </summary>
    private static string? FindSampleDocument()
    {
        var directory = new DirectoryInfo(Directory.GetCurrentDirectory());

        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "SampleDocument.docx");
            Console.WriteLine(candidate);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        return null;
    }
}
