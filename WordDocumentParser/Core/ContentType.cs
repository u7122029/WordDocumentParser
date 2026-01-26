namespace WordDocumentParser.Core;

/// <summary>
/// Represents the type of content in a document node.
/// </summary>
public enum ContentType
{
    /// <summary>Root document node</summary>
    Document,

    /// <summary>Heading paragraph (H1-H9)</summary>
    Heading,

    /// <summary>Regular paragraph</summary>
    Paragraph,

    /// <summary>Table element</summary>
    Table,

    /// <summary>Embedded image</summary>
    Image,

    /// <summary>List container</summary>
    List,

    /// <summary>Individual list item</summary>
    ListItem,

    /// <summary>Hyperlink text span</summary>
    HyperlinkText,

    /// <summary>Formatted text run</summary>
    TextRun,

    /// <summary>Structured Document Tag (content control)</summary>
    ContentControl
}
