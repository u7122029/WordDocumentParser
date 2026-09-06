using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using WordDocumentParser.Models.Package;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Tree construction, section extraction, merging, and failure behaviour.
/// </summary>
public class StructureAndMergeTests
{
    [Fact]
    public void HeadingsNestUnderTheirOwnSectionAndKeepDocumentOrder()
    {
        var document = Parse(Create(
            Paragraph("A", 1) + Paragraph("A2", 2) + Paragraph("B", 1) +
            Paragraph("B4", 4) + Paragraph("B3", 3)));

        var b3 = document.FindFirst(n => n.Text == "B3")!;
        Assert.Equal("B", b3.Parent!.Text);

        Assert.Equal("AA2BB4B3", SavedText(document));
    }

    [Fact]
    public void AHeadingWrappedInAContentControlKeepsTheRestOfItsSection()
    {
        var document = Parse(Create(
            "<w:sdt><w:sdtPr><w:tag w:val='heading'/><w:text/></w:sdtPr>" +
            $"<w:sdtContent>{Paragraph("Title", 1)}</w:sdtContent></w:sdt>" +
            Paragraph("Following body")));

        Assert.Equal("TitleFollowing body", SavedText(document));
    }

    [Fact]
    public void ExtractSectionFindsANestedHeading()
    {
        var document = Parse(Create(Paragraph("Parent", 1) + Paragraph("Nested", 2) + Paragraph("Body")));

        var section = document.ExtractSection("Nested");

        Assert.Single(section);
        Assert.Equal("Nested", section[0].Text);
    }

    [Fact]
    public void ExtractSectionWithoutNestedHeadingsExcludesThem()
    {
        var document = Parse(Create(Paragraph("Parent", 1) + Paragraph("Nested", 2) + Paragraph("Body")));

        var section = document.ExtractSection("Parent", includeNestedHeadings: false);

        Assert.DoesNotContain(section.SelectMany(n => n.GetAllHeadings()), n => n.HeadingLevel == 2);
    }

    [Fact]
    public void EditingAConcatenatedDocumentLeavesTheSourceAlone()
    {
        var source = Parse(Create(Table(Paragraph("OLD"))));
        var combined = DocumentMergeExtensions.ConcatenateDocuments([source]);

        combined.FindAllTables().First().SetCellText(0, 0, "NEW");

        Assert.Equal("OLD", source.FindAllTables().First().GetCellText(0, 0));
        Assert.Equal("NEW", combined.FindAllTables().First().GetCellText(0, 0));
    }

    [Fact]
    public void CloningCarriesPropertiesSetSinceParsing()
    {
        var source = new WordDocument();
        source["ProjectCode"] = "NEW";

        var clone = DocumentMergeExtensions.ConcatenateDocuments([source]);

        Assert.Equal("NEW", clone["ProjectCode"]);
    }

    [Fact]
    public void MergingImagesDoesNotStealAnExistingHyperlinksRelationshipId()
    {
        var target = Parse(Create(
            "<w:p><w:hyperlink r:id='rId1000'><w:r><w:t>LINK</w:t></w:r></w:hyperlink></w:p>",
            p => p.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rId1000")));

        var source = new WordDocument();
        source.Images["rId9"] = new ImagePartData { ContentType = "image/png", Data = TinyPng };

        target.AppendDocument(source, addPageBreak: false);

        using var saved = Open(Save(target));
        var linkId = saved.MainPart.Document.Descendants<Hyperlink>().First().Id!.Value!;
        Assert.Contains(saved.MainPart.HyperlinkRelationships, rel => rel.Id == linkId);
    }

    [Fact]
    public void MergingBringsTheSourcesNumberingDefinitionsAlong()
    {
        var source = Parse(Create(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='42'/></w:numPr></w:pPr>" +
            "<w:r><w:t>Numbered</w:t></w:r></w:p>",
            p => p.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." })
                    { LevelIndex = 0 })
                { AbstractNumberId = 41 },
                new NumberingInstance(new AbstractNumId { Val = 41 }) { NumberID = 42 })));

        var target = new WordDocument();
        target.AppendDocument(source, addPageBreak: false);

        using var saved = Open(Save(target));
        var numbering = saved.MainPart.NumberingDefinitionsPart!.Numbering;

        var instance = numbering.Elements<NumberingInstance>().Single();
        var abstractId = instance.AbstractNumId!.Val!.Value;

        Assert.Contains(numbering.Elements<AbstractNum>(), a => a.AbstractNumberId!.Value == abstractId);

        // The body's reference must point at an instance that exists.
        var referenced = saved.MainPart.Document.Descendants<NumberingId>().Single().Val!.Value;
        Assert.Equal(instance.NumberID!.Value, referenced);
    }

    [Fact]
    public void AMissingReplacementSectionLeavesTheTargetIntact()
    {
        var target = Parse(Create(Paragraph("Keep", 1) + Paragraph("Body")));

        Assert.Throws<ArgumentException>(() => target.ReplaceSection("Keep", new WordDocument(), "Missing"));

        Assert.Single(target.Root.Children);
        Assert.Equal("Keep", target.Root.Children[0].Text);
    }

    [Fact]
    public void AFailedSaveLeavesTheExistingFileInPlace()
    {
        var path = Path.Combine(Path.GetTempPath(), $"wdp-{Guid.NewGuid():N}.docx");
        File.WriteAllText(path, "ORIGINAL");

        try
        {
            var document = new WordDocument();
            document.PackageData.HyperlinkRelationships["rId99"] =
                new HyperlinkRelationshipData { Url = "http://[not a uri", IsExternal = true };

            using var writer = new WordDocumentTreeWriter();
            Assert.Throws<DocumentPreservationException>(() => writer.WriteToFile(document, path));

            Assert.Equal("ORIGINAL", File.ReadAllText(path));

            // The staging file must not be left behind either.
            Assert.DoesNotContain(
                Directory.GetFiles(Path.GetDirectoryName(path)!, $"{Path.GetFileName(path)}*"),
                file => file.EndsWith(".tmp", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AWriterCanBeReusedForASecondDocument()
    {
        var document = new WordDocument();
        document.Root.AddChild(new DocumentNode(ContentType.ListItem, "Item"));

        using var writer = new WordDocumentTreeWriter();
        writer.BuildPackage(document);

        using var second = Open(writer.BuildPackage(document));
        Assert.NotNull(second.MainPart.NumberingDefinitionsPart);
    }

    [Fact]
    public void AParserCanBeReusedForASecondDocument()
    {
        var first = Create(Paragraph("First"));
        var second = Create(Paragraph("Second"));

        using var parser = new WordDocumentTreeParser();

        using var firstStream = new MemoryStream(first);
        Assert.Equal("First", parser.ParseFromStream(firstStream).Root.Children[0].Text);

        using var secondStream = new MemoryStream(second);
        Assert.Equal("Second", parser.ParseFromStream(secondStream).Root.Children[0].Text);
    }

    [Fact]
    public void ADocumentCreatedInCodeValidatesAgainstOffice2019()
    {
        var document = new WordDocument();
        document.Root.AddChild(new DocumentNode(ContentType.Heading, 1, "Title"));
        document.Root.AddChild(new DocumentNode(ContentType.Paragraph, "Body"));

        Assert.Empty(Validate(Save(document)));
    }

    [Fact]
    public void ParsingRejectsADocumentThatExceedsItsBinaryBudget()
    {
        var package = Create(Paragraph("Body"), p =>
        {
            var imagePart = p.MainDocumentPart!.AddImagePart("image/png");
            using var stream = new MemoryStream(TinyPng);
            imagePart.FeedData(stream);
        });

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxTotalBinaryBytes = 1 }
        };

        using var input = new MemoryStream(package);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Fact]
    public void ParsingRejectsADocumentWithTooManyParts()
    {
        // Two parts: the main document, plus an image hanging off it.
        var package = Create(Paragraph("Body"), p =>
        {
            var imagePart = p.MainDocumentPart!.AddImagePart("image/png");
            using var stream = new MemoryStream(TinyPng);
            imagePart.FeedData(stream);
        });

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxPartCount = 1 }
        };

        using var input = new MemoryStream(package);
        Assert.Throws<DocumentLimitExceededException>(() => parser.ParseFromStream(input));
    }

    [Fact]
    public void ParsingAcceptsADocumentWithinItsPartBudget()
    {
        var package = Create(Paragraph("Body"));

        using var parser = new WordDocumentTreeParser
        {
            Limits = new DocumentLimits { MaxPartCount = 64 }
        };

        using var input = new MemoryStream(package);
        Assert.Equal("Body", parser.ParseFromStream(input).Root.Children[0].Text);
    }
}
