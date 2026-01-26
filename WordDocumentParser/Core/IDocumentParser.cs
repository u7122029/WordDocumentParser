namespace WordDocumentParser.Core;

/// <summary>
/// Interface for parsing Word documents into a document tree structure.
/// </summary>
public interface IDocumentParser : IDisposable
{
    /// <summary>
    /// Parses a Word document from a file path.
    /// </summary>
    /// <param name="filePath">Path to the .docx file</param>
    /// <returns>The parsed Word document with metadata and content structure</returns>
    WordDocument ParseFromFile(string filePath);

    /// <summary>
    /// Parses a Word document from a stream.
    /// </summary>
    /// <param name="stream">Stream containing the .docx data</param>
    /// <param name="documentName">Optional name for the document</param>
    /// <returns>The parsed Word document with metadata and content structure</returns>
    WordDocument ParseFromStream(Stream stream, string documentName = "Document");
}
