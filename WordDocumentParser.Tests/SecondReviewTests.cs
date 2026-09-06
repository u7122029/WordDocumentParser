using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Regressions for text corruption, lost edits, ineffective removals, and unbounded parsing.
/// </summary>
public class SecondReviewTests
{
    // ---- Text ordering -----------------------------------------------------

    [Fact]
    public void StylingInsideARunWithLaterChildrenKeepsTheContentOrder()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>abc</w:t><w:tab/><w:t>def</w:t></w:r></w:p>"));

        Assert.Equal(1, document.Root.Children[0].SetFontForText("b", "Arial"));

        var saved = XDocument.Parse(SavedBodyXml(document));
        Assert.Equal("abc\tdef", VisibleText(saved));
    }

    [Fact]
    public void StylingAcrossARunBoundaryWithABreakKeepsTheContentOrder()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>one</w:t><w:br/><w:t>two</w:t></w:r><w:r><w:t>three</w:t></w:r></w:p>"));

        Assert.Equal(1, document.Root.Children[0].SetFontForText("wo", "Arial"));

        var saved = XDocument.Parse(SavedBodyXml(document));

        // The break renders as a space; nothing separates "two" from "three" in the source.
        Assert.Equal("one twothree", VisibleText(saved));

        // The break must still sit between "one" and "two", not after them.
        var order = saved.Descendants(XName.Get("p", W)).First().Descendants()
            .Where(e => e.Name.LocalName is "t" or "br")
            .Select(e => e.Name.LocalName == "br" ? "<br>" : e.Value);
        Assert.Equal(["one", "<br>", "t", "wo", "three"], order);
    }

    /// <summary>
    /// Reads the paragraph's visible text in document order, rendering tabs and breaks the way the
    /// model does so ordering regressions show up as text differences.
    /// </summary>
    private static string VisibleText(XDocument saved) =>
        string.Concat(saved.Descendants(XName.Get("p", W)).First().Descendants()
            .Where(e => e.Name.LocalName is "t" or "tab" or "br")
            .Select(e => e.Name.LocalName switch
            {
                "t" => e.Value,
                "tab" => "\t",
                _ => " "
            }));

    // ---- Saving must not consume the model ---------------------------------

    [Fact]
    public void SavingTwiceKeepsAnInlineContentControlEdit()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>Before </w:t></w:r>" +
            "<w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='name'/><w:text/></w:sdtPr>" +
            "<w:sdtContent><w:r><w:t>OLD</w:t></w:r></w:sdtContent></w:sdt>" +
            "<w:r><w:t> After</w:t></w:r></w:p>"));

        // Assigning the run's text is the whole edit: the node must notice it on its own.
        document.Root.Children[0].Runs.First(r => r.Text == "OLD").Text = "NEW";

        Assert.Equal("Before NEW After", SavedText(document));
        Assert.Equal("Before NEW After", SavedText(document));
    }

    [Fact]
    public void ADocumentWideFontChangeReachesInlineContentControlText()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:t>Before </w:t></w:r>" +
            "<w:sdt><w:sdtPr><w:id w:val='7'/><w:tag w:val='name'/><w:text/></w:sdtPr>" +
            "<w:sdtContent><w:r><w:t>Inside</w:t></w:r></w:sdtContent></w:sdt></w:p>"));

        document.SetDocumentFont("Arial");

        var controlRun = XDocument.Parse(SavedBodyXml(document))
            .Descendants(XName.Get("sdtContent", W))
            .Descendants(XName.Get("r", W))
            .First();

        Assert.Contains(
            controlRun.Descendants(XName.Get("rFonts", W)),
            e => (string?)e.Attribute(XName.Get("ascii", W)) == "Arial");
    }

    // ---- Control removal ---------------------------------------------------

    [Fact]
    public void RemovingAMultiBlockContentControlKeepsEveryBlock()
    {
        var document = Parse(Create(
            "<w:sdt><w:sdtPr><w:tag w:val='group'/></w:sdtPr><w:sdtContent>" +
            Paragraph("ONE") + Paragraph("TWO") +
            "</w:sdtContent></w:sdt>"));

        Assert.True(document.RemoveContentControlByTag("group"));

        var savedText = SavedText(document);
        Assert.Contains("ONE", savedText, StringComparison.Ordinal);
        Assert.Contains("TWO", savedText, StringComparison.Ordinal);
        Assert.DoesNotContain("<w:sdt", SavedBodyXml(document), StringComparison.Ordinal);
    }

    [Fact]
    public void RemovingASingleParagraphContentControlKeepsItsFormatting()
    {
        var document = Parse(Create(
            "<w:sdt><w:sdtPr><w:tag w:val='one'/><w:text/></w:sdtPr><w:sdtContent>" +
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>BOLD</w:t></w:r></w:p>" +
            "</w:sdtContent></w:sdt>"));

        Assert.True(document.RemoveContentControlByTag("one"));

        var saved = SavedBodyXml(document);
        Assert.Equal("BOLD", TextOf(saved));
        Assert.DoesNotContain("<w:sdt", saved, StringComparison.Ordinal);
        Assert.Contains("<w:b", saved, StringComparison.Ordinal);
    }

    // ---- Property removal --------------------------------------------------

    [Fact]
    public void RemovingACorePropertyClearsItInTheSavedDocument()
    {
        var package = Create(Paragraph("Body"), p => p.PackageProperties.Title = "Original Title");

        var document = Parse(package);
        Assert.Equal("Original Title", document["Title"]);
        Assert.True(document.RemoveProperty("Title"));

        using var saved = Open(Save(document));
        Assert.True(string.IsNullOrEmpty(saved.Document.PackageProperties.Title));
    }

    [Fact]
    public void RemovingAnExtendedPropertyClearsItInTheSavedDocument()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddExtendedFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties(
                new DocumentFormat.OpenXml.ExtendedProperties.Company("Acme"));
        });

        var document = Parse(package);
        Assert.Equal("Acme", document["Company"]);
        Assert.True(document.RemoveProperty("Company"));

        using var saved = Open(Save(document));
        Assert.Null(saved.Document.ExtendedFilePropertiesPart!.Properties.Company);
    }

    [Fact]
    public void AnUntouchedCorePropertySurvivesASave()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            p.PackageProperties.Title = "Keep Me";
            p.PackageProperties.Creator = "Author";
        });

        var document = Parse(package);
        document["Creator"] = "New Author";

        using var saved = Open(Save(document));
        Assert.Equal("Keep Me", saved.Document.PackageProperties.Title);
        Assert.Equal("New Author", saved.Document.PackageProperties.Creator);
    }

    // ---- Numbering ---------------------------------------------------------

    [Fact]
    public void MergingCollidingListIdsKeepsEachParagraphOnItsOwnDefinition()
    {
        var source = Parse(Create(
            NumberedParagraph("First list", 1) + NumberedParagraph("Second list", 2),
            p => p.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                AbstractNum(11), AbstractNum(12),
                new NumberingInstance(new AbstractNumId { Val = 11 }) { NumberID = 1 },
                new NumberingInstance(new AbstractNumId { Val = 12 }) { NumberID = 2 })));

        var target = Parse(Create(
            NumberedParagraph("Target list", 1),
            p => p.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                AbstractNum(21),
                new NumberingInstance(new AbstractNumId { Val = 21 }) { NumberID = 1 })));

        target.AppendDocument(source, addPageBreak: false);

        using var saved = Open(Save(target));
        var numbering = saved.MainPart.NumberingDefinitionsPart!.Numbering;
        var definedIds = numbering.Elements<NumberingInstance>().Select(n => n.NumberID!.Value).ToList();

        var referenced = saved.MainPart.Document.Descendants<NumberingId>()
            .Select(n => n.Val!.Value)
            .ToList();

        // Three paragraphs, three definitions, and the two source lists must stay distinct.
        Assert.Equal(3, referenced.Count);
        Assert.Equal(3, referenced.Distinct().Count());
        Assert.All(referenced, id => Assert.Contains(id, definedIds));
    }

    private static string NumberedParagraph(string text, int numberingId) =>
        $"<w:p><w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='{numberingId}'/></w:numPr></w:pPr>" +
        $"<w:r><w:t>{text}</w:t></w:r></w:p>";

    private static AbstractNum AbstractNum(int id) => new(
        new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = "%1." })
        { LevelIndex = 0})
    { AbstractNumberId = id };

    // ---- Limits ------------------------------------------------------------

    [Fact]
    public void AnOversizedCustomXmlPartIsRejected()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var customXml = p.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = customXml.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write("<root>");
            writer.Write(new string('x', 1_000_000));
            writer.Write("</root>");
        });

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxCharactersInPart = 4096 }
        };

        using var input = new MemoryStream(package);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Fact]
    public void DeeplyNestedTablesAreRejected()
    {
        var body = Paragraph("innermost");
        for (var i = 0; i < 12; i++)
        {
            body = Table(body);
        }

        var package = Create(body);

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxElementDepth = 4 }
        };

        using var input = new MemoryStream(package);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Fact]
    public void NestingWithinTheDepthLimitIsAccepted()
    {
        var package = Create(Table(Table(Paragraph("inner"))));

        // The limit counts XML elements, so two nested tables sit well under a realistic bound.
        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxElementDepth = 32 }
        };

        using var input = new MemoryStream(package);
        Assert.Equal(2, parser.ParseFromStream(input).FindAllTables().Count());
    }

    [Fact]
    public void AnOrdinaryDocumentPassesTheUntrustedDepthLimit()
    {
        var package = Create(
            Paragraph("Title", 1) +
            Table(Table(Paragraph("nested"))) +
            "<w:p><w:hyperlink r:id='rId1'><w:r><w:rPr><w:b/></w:rPr><w:t>LINK</w:t></w:r></w:hyperlink></w:p>",
            p => p.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId1"));

        using var parser = new WordDocumentTreeParser { Limits = DocumentLimits.Untrusted };
        using var input = new MemoryStream(package);

        Assert.NotNull(parser.ParseFromStream(input));
    }

    [Fact]
    public void AnOversizedPackageIsRejectedBeforeItIsBuffered()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var imagePart = p.MainDocumentPart!.AddImagePart("image/png");
            var payload = new byte[512 * 1024];
            Random.Shared.NextBytes(payload);
            using var stream = new MemoryStream(payload);
            imagePart.FeedData(stream);
        });

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxTotalBinaryBytes = 64 * 1024 }
        };

        using var input = new MemoryStream(package);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    // ---- Custom properties without package capture -------------------------

    [Fact]
    public void CustomPropertiesSurviveWithPackageCaptureDisabled()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                    new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR("Kept"))
                {
                    Name = "Marker",
                    PropertyId = 2,
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                });
        });

        using var parser = new WordDocumentTreeParser { CaptureOriginalPackage = false };
        using var input = new MemoryStream(package);
        var document = parser.ParseFromStream(input);

        Assert.Equal("Kept", document["Marker"]);

        using var saved = Open(Save(document));
        Assert.NotNull(saved.Document.CustomFilePropertiesPart);
        Assert.Contains("Kept", saved.Document.CustomFilePropertiesPart!.Properties.InnerXml, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOutOfRangeValueFallsBackToTextRatherThanBreakingItsType()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var part = p.AddCustomFilePropertiesPart();
            part.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties(
                new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty(
                    new DocumentFormat.OpenXml.VariantTypes.VTInt32("42"))
                {
                    Name = "Count",
                    PropertyId = 2,
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                });
        });

        var document = Parse(package);
        document["Count"] = "2147483648";

        var bytes = Save(document);

        using var saved = Open(bytes);
        var property = saved.Document.CustomFilePropertiesPart!.Properties.FirstChild!;
        Assert.Equal("lpwstr", property.FirstChild!.LocalName);

        Assert.Empty(Validate(bytes));
    }

    // ---- Cleared run collection --------------------------------------------

    [Fact]
    public void ClearingTheRunCollectionEmptiesTheParagraph()
    {
        var document = Parse(Create(Paragraph("SECRET")));

        document.Root.Children[0].Runs.Clear();
        document.Root.Children[0].MarkRunsChanged();

        Assert.Equal(string.Empty, SavedText(document));
    }

    // ---- Theme fonts -------------------------------------------------------

    [Fact]
    public void AnExplicitFontClearsTheThemeReferenceItReplaces()
    {
        var document = Parse(Create(
            "<w:p><w:r><w:rPr><w:rFonts w:asciiTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/></w:rPr>" +
            "<w:t>Text</w:t></w:r></w:p>"));

        document.Root.Children[0].SetParagraphFont("Arial");

        var fonts = XDocument.Parse(SavedBodyXml(document))
            .Descendants(XName.Get("rFonts", W))
            .Single();

        Assert.Equal("Arial", (string?)fonts.Attribute(XName.Get("ascii", W)));
        Assert.Null(fonts.Attribute(XName.Get("asciiTheme", W)));
        Assert.Null(fonts.Attribute(XName.Get("hAnsiTheme", W)));
    }

    // ---- Allocation growth -------------------------------------------------

    [Fact]
    public void RepeatedMatchFontChangesAllocateInProportionToTheMatchCount()
    {
        static long Measure(int matches)
        {
            var node = new DocumentNode(ContentType.Paragraph,
                string.Concat(Enumerable.Repeat("match filler ", matches)));

            var before = GC.GetAllocatedBytesForCurrentThread();
            node.SetFontForText("match", "Arial", allOccurrences: true);
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }

        Measure(50);   // warm up

        var small = Measure(500);
        var large = Measure(2000);

        // Four times the matches should cost roughly four times the allocation. Rebuilding the run
        // collection per match made it grow with the square instead.
        Assert.True(large < small * 8, $"2000 matches allocated {large} bytes against {small} for 500.");
    }
}
