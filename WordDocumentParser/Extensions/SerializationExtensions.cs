namespace WordDocumentParser.Extensions;

/// <summary>
/// Extension methods for saving and exporting documents.
/// </summary>
public static class SerializationExtensions
{
    /// <summary>
    /// Saves the document tree to a Word document file (.docx).
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <param name="filePath">Path where the document will be saved</param>
    public static void SaveToFile(this DocumentNode root, string filePath)
    {
        using var writer = new WordDocumentTreeWriter();
        writer.WriteToFile(root, filePath);
    }

    /// <summary>
    /// Saves the document tree to a stream as a Word document (.docx).
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <param name="stream">Stream to write to</param>
    public static void SaveToStream(this DocumentNode root, Stream stream)
    {
        using var writer = new WordDocumentTreeWriter();
        writer.WriteToStream(root, stream);
    }

    /// <summary>
    /// Saves the document tree to a byte array as a Word document (.docx).
    /// </summary>
    /// <param name="root">The root document node</param>
    /// <returns>Document as a byte array</returns>
    public static byte[] ToDocxBytes(this DocumentNode root)
    {
        using var stream = new MemoryStream();
        root.SaveToStream(stream);
        return stream.ToArray();
    }
}
