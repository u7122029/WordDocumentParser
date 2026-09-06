using DocumentFormat.OpenXml.Packaging;
using WordDocumentParser.Core;
using WordDocumentParser.Extensions;
using Xunit;
using static WordDocumentParser.Tests.DocumentFixture;

namespace WordDocumentParser.Tests;

/// <summary>
/// Guards the costs that grew faster than the work: repeated media, and tree rendering.
/// </summary>
public class ResourceUseTests
{
    [Fact]
    public void AnImageUsedManyTimesIsHeldOnce()
    {
        const int occurrences = 32;
        var payload = new byte[64 * 1024];
        Random.Shared.NextBytes(payload);

        var drawing =
            "<w:p><w:r><w:drawing><wp:inline distT='0' distB='0' distL='0' distR='0' " +
            "xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'>" +
            "<wp:extent cx='914400' cy='914400'/><wp:docPr id='1' name='Image'/>" +
            "<a:graphic xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>" +
            "<a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/picture'>" +
            "<pic:pic xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture'>" +
            "<pic:nvPicPr><pic:cNvPr id='1' name='Image'/><pic:cNvPicPr/></pic:nvPicPr>" +
            "<pic:blipFill><a:blip r:embed='rIdImg'/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
            "<pic:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='914400' cy='914400'/></a:xfrm>" +
            "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom></pic:spPr></pic:pic>" +
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

        var package = Create(
            string.Concat(Enumerable.Repeat(drawing, occurrences)),
            p =>
            {
                var imagePart = p.MainDocumentPart!.AddImagePart("image/png", "rIdImg");
                using var stream = new MemoryStream(payload);
                imagePart.FeedData(stream);
            });

        var document = Parse(package);

        var distinctBuffers = document.Root
            .FindAllContent(n => n.Type == ContentType.Image)
            .Select(n => n.GetImageData()?.Data)
            .Where(data => data is not null)
            .Distinct(ReferenceEqualityComparer.Instance)
            .Count();

        Assert.Equal(occurrences, document.Root.FindAllContent(n => n.Type == ContentType.Image).Count());
        Assert.Equal(1, distinctBuffers);
    }

    [Fact]
    public void RenderingATreeAllocatesInProportionToItsSize()
    {
        static long MeasureAllocations(int nodeCount)
        {
            var root = new DocumentNode(ContentType.Document, "root");
            for (var i = 0; i < nodeCount; i++)
            {
                root.AddChild(new DocumentNode(ContentType.Paragraph, $"Paragraph number {i}"));
            }

            // Warm up so the measurement excludes one-time costs.
            root.ToTreeString();

            var before = GC.GetAllocatedBytesForCurrentThread();
            root.ToTreeString();
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }

        var small = MeasureAllocations(200);
        var large = MeasureAllocations(800);

        // Four times the nodes should cost roughly four times the allocation. Concatenating each
        // subtree's rendered string instead cost sixteen times as much, and kept getting worse.
        Assert.True(large < small * 8, $"800 nodes allocated {large} bytes against {small} for 200.");
    }

    [Fact]
    public void TreeRenderingRespectsThePreviewLengthAtEveryDepth()
    {
        var root = new DocumentNode(ContentType.Document, "root");
        var child = new DocumentNode(ContentType.Paragraph, new string('x', 200));
        root.AddChild(child);

        var rendered = root.ToTreeString(previewLength: 10);

        Assert.DoesNotContain(new string('x', 11), rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void ShortTextRendersWithoutSlicingPastItsEnd()
    {
        var node = new DocumentNode(ContentType.Paragraph, "short");

        Assert.Contains("short", node.ToTreeString(previewLength: 100), StringComparison.Ordinal);
        Assert.Contains("short", node.ToString(100), StringComparison.Ordinal);
    }
}
