namespace WordDocumentParser.FormattingModels;

/// <summary>
/// Represents the type of content control
/// </summary>
public enum ContentControlType
{
    Unknown,
    RichText,
    PlainText,
    Picture,
    Date,
    DropDownList,
    ComboBox,
    Checkbox,
    RepeatingSection,
    RepeatingSectionItem,
    BuildingBlockGallery,
    Group,
    Bibliography,
    Citation,
    Equation,
    DocumentProperty
}