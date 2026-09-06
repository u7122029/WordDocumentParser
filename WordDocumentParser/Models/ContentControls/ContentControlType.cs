namespace WordDocumentParser.Models.ContentControls;

/// <summary>
/// The kind of a structured document tag (content control).
/// </summary>
public enum ContentControlType
{
    /// <summary>The control declares no recognised type.</summary>
    Unknown,

    /// <summary>Rich text: formatted content, including other blocks.</summary>
    RichText,

    /// <summary>Plain text: unformatted text only.</summary>
    PlainText,

    /// <summary>Picture placeholder.</summary>
    Picture,

    /// <summary>Date picker.</summary>
    Date,

    /// <summary>Drop-down list, restricted to its listed items.</summary>
    DropDownList,

    /// <summary>Combo box: a list that also accepts typed text.</summary>
    ComboBox,

    /// <summary>Checkbox, whose state lives in a <c>w14:checkbox</c> element.</summary>
    Checkbox,

    /// <summary>Repeating section, which repeats its item control.</summary>
    RepeatingSection,

    /// <summary>One item of a repeating section.</summary>
    RepeatingSectionItem,

    /// <summary>Building block gallery, backed by the glossary document.</summary>
    BuildingBlockGallery,

    /// <summary>Group: a container that protects the content inside it.</summary>
    Group,

    /// <summary>Bibliography, generated from the document's sources.</summary>
    Bibliography,

    /// <summary>A single citation within a bibliography.</summary>
    Citation,

    /// <summary>Equation placeholder.</summary>
    Equation,

    /// <summary>Bound to a document property through a data binding.</summary>
    DocumentProperty
}
