namespace WordDocumentParser.Models.Package;

/// <summary>
/// Stores the original document package data for round-trip fidelity.
/// This preserves styles, themes, fonts, properties, and other document parts.
/// </summary>
public class DocumentPackageData
{
    /// <summary>
    /// The bytes of the source <c>.docx</c> package, captured at parse time.
    /// </summary>
    /// <remarks>
    /// <para>
    /// When this is set the writer edits a copy of the original package in place rather than
    /// assembling a new one from the properties below. That is what makes preservation total: parts
    /// this library has no model for — comments, revisions, embedded objects, ink, charts, VBA — and
    /// the relationships that headers, footers, and notes own travel through untouched, as do
    /// extension-namespace elements such as the <c>w14:checkbox</c> that carries a checkbox's state.
    /// </para>
    /// <para>
    /// Rebuilding from the model could only ever preserve what the model happened to capture, so
    /// anything outside it was silently dropped. Documents constructed in code have no source
    /// package and are still assembled from the properties below.
    /// </para>
    /// </remarks>
    public byte[]? OriginalPackageBytes { get; set; }

    /// <summary>
    /// The values captured from the source package, used to tell an edited part from an untouched
    /// one. The writer overwrites a part in the copied package only where the current value differs
    /// from its baseline.
    /// </summary>
    public PackageBaseline? Baseline { get; set; }

    /// <summary>
    /// Original styles.xml content
    /// </summary>
    public string? StylesXml { get; set; }

    /// <summary>
    /// Original theme XML content (theme/theme1.xml)
    /// </summary>
    public string? ThemeXml { get; set; }

    /// <summary>
    /// Original font table XML (fontTable.xml)
    /// </summary>
    public string? FontTableXml { get; set; }

    /// <summary>
    /// Original numbering definitions XML (numbering.xml)
    /// </summary>
    public string? NumberingXml { get; set; }

    /// <summary>
    /// Original document settings XML (settings.xml)
    /// </summary>
    public string? SettingsXml { get; set; }

    /// <summary>
    /// Original web settings XML (webSettings.xml)
    /// </summary>
    public string? WebSettingsXml { get; set; }

    /// <summary>
    /// Original footnotes XML
    /// </summary>
    public string? FootnotesXml { get; set; }

    /// <summary>
    /// Original endnotes XML
    /// </summary>
    public string? EndnotesXml { get; set; }

    /// <summary>
    /// Core document properties (author, title, etc.)
    /// </summary>
    public CoreProperties? CoreProperties { get; set; }

    /// <summary>
    /// Extended document properties (company, template, word count, etc.)
    /// </summary>
    public ExtendedProperties? ExtendedProperties { get; set; }

    /// <summary>
    /// Custom document properties
    /// </summary>
    public string? CustomPropertiesXml { get; set; }

    /// <summary>
    /// Header parts - key is the relationship ID, value is the XML content
    /// </summary>
    public Dictionary<string, string> Headers { get; set; } = [];

    /// <summary>
    /// Footer parts - key is the relationship ID, value is the XML content
    /// </summary>
    public Dictionary<string, string> Footers { get; set; } = [];

    /// <summary>
    /// Image parts from headers - key is header relationship ID, value is dictionary of image relationship ID to image data
    /// </summary>
    public Dictionary<string, Dictionary<string, ImagePartData>> HeaderImages { get; set; } = [];

    /// <summary>
    /// Image parts from footers - key is footer relationship ID, value is dictionary of image relationship ID to image data
    /// </summary>
    public Dictionary<string, Dictionary<string, ImagePartData>> FooterImages { get; set; } = [];

    /// <summary>
    /// Image parts - key is the relationship ID, value is the image data
    /// </summary>
    public Dictionary<string, ImagePartData> Images { get; set; } = [];

    /// <summary>
    /// Section properties from the original document
    /// </summary>
    public List<string> SectionPropertiesXml { get; set; } = [];

    /// <summary>
    /// The original document.xml content (for reference)
    /// </summary>
    public string? OriginalDocumentXml { get; set; }

    /// <summary>
    /// Custom XML parts - key is the part URI, value is the XML content
    /// </summary>
    public Dictionary<string, CustomXmlPartData> CustomXmlParts { get; set; } = [];

    /// <summary>
    /// Original core.xml content for exact round-trip
    /// </summary>
    public string? CorePropertiesXml { get; set; }

    /// <summary>
    /// Original app.xml content for exact round-trip
    /// </summary>
    public string? AppPropertiesXml { get; set; }

    /// <summary>
    /// Hyperlink relationships - key is the relationship ID, value is the URL
    /// </summary>
    public Dictionary<string, HyperlinkRelationshipData> HyperlinkRelationships { get; set; } = [];

    /// <summary>
    /// Glossary document XML (for Quick Parts, building blocks, document property fields)
    /// </summary>
    public string? GlossaryDocumentXml { get; set; }

    /// <summary>
    /// Glossary document styles XML
    /// </summary>
    public string? GlossaryStylesXml { get; set; }

    /// <summary>
    /// Glossary document fonts XML
    /// </summary>
    public string? GlossaryFontTableXml { get; set; }

    /// <summary>
    /// Images from glossary document part - key is relationship ID, value is image data
    /// </summary>
    public Dictionary<string, ImagePartData> GlossaryImages { get; set; } = [];

    /// <summary>
    /// Every relationship ID the main document part used in the source package.
    /// </summary>
    /// <remarks>
    /// Merging allocates new IDs against this set rather than against the images alone. Allocating
    /// from one resource kind at a time handed out an ID another kind already held, and the winner
    /// took over the loser's references — an image merge could leave an existing hyperlink pointing
    /// at a picture.
    /// </remarks>
    public HashSet<string> KnownRelationshipIds { get; set; } = new(StringComparer.Ordinal);

    /// <summary>
    /// True when the writer can edit a copy of the source package instead of assembling a new one.
    /// </summary>
    public bool CanEditInPlace => OriginalPackageBytes is { Length: > 0 };

    /// <summary>
    /// Allocates a relationship ID that no resource in this package is using, and records it.
    /// </summary>
    /// <param name="prefix">The ID prefix, normally <c>"rId"</c>.</param>
    /// <returns>The new ID.</returns>
    public string AllocateRelationshipId(string prefix = "rId")
    {
        _nextRelationshipNumber = Math.Max(_nextRelationshipNumber, 1000);

        string candidate;
        do
        {
            candidate = $"{prefix}{_nextRelationshipNumber++}";
        }
        while (KnownRelationshipIds.Contains(candidate) ||
               Images.ContainsKey(candidate) ||
               HyperlinkRelationships.ContainsKey(candidate));

        KnownRelationshipIds.Add(candidate);
        return candidate;
    }

    private int _nextRelationshipNumber;

    /// <summary>
    /// Records the current values of the modelled parts as the baseline for change detection.
    /// </summary>
    public void CaptureBaseline()
    {
        Baseline = PackageBaseline.From(this);
        CoreProperties?.AcceptChanges();
        ExtendedProperties?.AcceptChanges();
    }
}

/// <summary>
/// A snapshot of the modelled package parts as they were read from the source document.
/// </summary>
/// <remarks>
/// The writer compares the live values against this snapshot to decide which parts to overwrite in
/// the copied package. Parts that match their baseline are left exactly as the source had them,
/// which preserves attributes and extension elements the model does not represent.
/// </remarks>
public sealed class PackageBaseline
{
    /// <summary>Creates a baseline from the current values of a package data instance.</summary>
    /// <param name="data">The package data to snapshot.</param>
    /// <returns>The snapshot.</returns>
    public static PackageBaseline From(DocumentPackageData data) => new()
    {
        StylesXml = data.StylesXml,
        ThemeXml = data.ThemeXml,
        FontTableXml = data.FontTableXml,
        NumberingXml = data.NumberingXml,
        SettingsXml = data.SettingsXml,
        WebSettingsXml = data.WebSettingsXml,
        FootnotesXml = data.FootnotesXml,
        EndnotesXml = data.EndnotesXml,
        CustomPropertiesXml = data.CustomPropertiesXml,
        GlossaryDocumentXml = data.GlossaryDocumentXml,
        Headers = new Dictionary<string, string>(data.Headers),
        Footers = new Dictionary<string, string>(data.Footers),
        ImageRelationshipIds = [.. data.Images.Keys],
        HyperlinkRelationshipIds = [.. data.HyperlinkRelationships.Keys],
        CustomXmlPartUris = [.. data.CustomXmlParts.Keys]
    };

    /// <summary>styles.xml as parsed.</summary>
    public string? StylesXml { get; init; }

    /// <summary>theme1.xml as parsed.</summary>
    public string? ThemeXml { get; init; }

    /// <summary>fontTable.xml as parsed.</summary>
    public string? FontTableXml { get; init; }

    /// <summary>numbering.xml as parsed.</summary>
    public string? NumberingXml { get; init; }

    /// <summary>settings.xml as parsed.</summary>
    public string? SettingsXml { get; init; }

    /// <summary>webSettings.xml as parsed.</summary>
    public string? WebSettingsXml { get; init; }

    /// <summary>footnotes.xml as parsed.</summary>
    public string? FootnotesXml { get; init; }

    /// <summary>endnotes.xml as parsed.</summary>
    public string? EndnotesXml { get; init; }

    /// <summary>custom.xml as parsed.</summary>
    public string? CustomPropertiesXml { get; init; }

    /// <summary>The glossary document as parsed.</summary>
    public string? GlossaryDocumentXml { get; init; }

    /// <summary>Header XML by relationship ID, as parsed.</summary>
    public Dictionary<string, string> Headers { get; init; } = [];

    /// <summary>Footer XML by relationship ID, as parsed.</summary>
    public Dictionary<string, string> Footers { get; init; } = [];

    /// <summary>Relationship IDs of the images present in the source package.</summary>
    public HashSet<string> ImageRelationshipIds { get; init; } = [];

    /// <summary>Relationship IDs of the hyperlinks present in the source package.</summary>
    public HashSet<string> HyperlinkRelationshipIds { get; init; } = [];

    /// <summary>URIs of the custom XML parts present in the source package.</summary>
    public HashSet<string> CustomXmlPartUris { get; init; } = [];
}