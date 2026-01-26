using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;
using WordDocumentParser.Models.Images;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace WordDocumentParser.Parsing.Extractors;

/// <summary>
/// Extracts image data and formatting from Word documents.
/// </summary>
internal sealed class ImageExtractor
{
    private readonly ParsingContext _context;

    public ImageExtractor(ParsingContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Extracts all images from a paragraph.
    /// </summary>
    public List<DocumentNode> ExtractImages(Paragraph para)
    {
        List<DocumentNode> images = [];

        var drawings = para.Descendants<Drawing>().ToList();
        foreach (var drawing in drawings)
        {
            var imageNode = ProcessDrawing(drawing);
            if (imageNode is not null)
                images.Add(imageNode);
        }

        return images;
    }

    /// <summary>
    /// Processes a drawing element to extract image information.
    /// </summary>
    public DocumentNode? ProcessDrawing(Drawing drawing)
    {
        var inline = drawing.Inline;
        var anchor = drawing.Anchor;

        var extent = inline?.Extent ?? anchor?.GetFirstChild<DW.Extent>();
        var docPr = inline?.DocProperties ?? anchor?.GetFirstChild<DW.DocProperties>();
        var graphic = inline?.Graphic ?? anchor?.GetFirstChild<A.Graphic>();

        if (graphic is null) return null;

        var blip = graphic.Descendants<A.Blip>().FirstOrDefault();
        if (blip is null) return null;

        var imageData = new ImageData();

        // Get image relationship ID and data
        var embedId = blip.Embed?.Value;
        if (!string.IsNullOrEmpty(embedId))
        {
            imageData.Id = embedId;

            try
            {
                if (_context.MainPart.GetPartById(embedId) is ImagePart imagePart)
                {
                    imageData.ContentType = imagePart.ContentType;
                    using var stream = imagePart.GetStream();
                    using var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    imageData.Data = ms.ToArray();
                }
            }
            catch
            {
                // Image extraction failed, continue without data
            }
        }

        // Get dimensions in EMUs for precise round-trip
        if (extent is not null)
        {
            imageData.WidthEmu = extent.Cx?.Value ?? 0;
            imageData.HeightEmu = extent.Cy?.Value ?? 0;
            imageData.WidthInches = imageData.WidthEmu / 914400.0;
            imageData.HeightInches = imageData.HeightEmu / 914400.0;
        }

        // Get alt text and description
        if (docPr is not null)
        {
            imageData.Name = docPr.Name?.Value ?? "";
            imageData.Description = docPr.Description?.Value;
            imageData.AltText = docPr.Title?.Value;
        }

        // Extract image formatting/positioning
        imageData.Formatting = ExtractImageFormatting(inline, anchor);

        var node = new DocumentNode(ContentType.Image, $"[Image: {imageData.Name}]");
        node.Metadata["ImageData"] = imageData;
        node.Metadata["Width"] = imageData.WidthInches;
        node.Metadata["Height"] = imageData.HeightInches;
        node.Metadata["ContentType"] = imageData.ContentType;

        return node;
    }

    /// <summary>
    /// Extracts image formatting and positioning information.
    /// </summary>
    public static ImageFormatting ExtractImageFormatting(DW.Inline? inline, DW.Anchor? anchor)
    {
        var formatting = new ImageFormatting();

        if (inline is not null)
        {
            formatting.IsInline = true;
            formatting.DistanceFromTop = inline.DistanceFromTop?.Value;
            formatting.DistanceFromBottom = inline.DistanceFromBottom?.Value;
            formatting.DistanceFromLeft = inline.DistanceFromLeft?.Value;
            formatting.DistanceFromRight = inline.DistanceFromRight?.Value;
        }
        else if (anchor is not null)
        {
            formatting.IsInline = false;
            formatting.DistanceFromTop = anchor.DistanceFromTop?.Value;
            formatting.DistanceFromBottom = anchor.DistanceFromBottom?.Value;
            formatting.DistanceFromLeft = anchor.DistanceFromLeft?.Value;
            formatting.DistanceFromRight = anchor.DistanceFromRight?.Value;
            formatting.AllowOverlap = anchor.AllowOverlap?.Value ?? false;
            formatting.BehindDocument = anchor.BehindDoc?.Value ?? false;
            formatting.LayoutInCell = anchor.LayoutInCell?.Value ?? false;
            formatting.Locked = anchor.Locked?.Value ?? false;
            formatting.RelativeHeight = anchor.RelativeHeight?.Value;

            // Horizontal position
            var hPos = anchor.HorizontalPosition;
            if (hPos is not null)
            {
                formatting.HorizontalRelativeTo = hPos.RelativeFrom?.Value.ToString();
                formatting.HorizontalPosition = hPos.PositionOffset?.Text;
            }

            // Vertical position
            var vPos = anchor.VerticalPosition;
            if (vPos is not null)
            {
                formatting.VerticalRelativeTo = vPos.RelativeFrom?.Value.ToString();
                formatting.VerticalPosition = vPos.PositionOffset?.Text;
            }

            // Wrap type
            formatting.WrapType = anchor switch
            {
                _ when anchor.GetFirstChild<DW.WrapNone>() is not null => "None",
                _ when anchor.GetFirstChild<DW.WrapSquare>() is not null => "Square",
                _ when anchor.GetFirstChild<DW.WrapTight>() is not null => "Tight",
                _ when anchor.GetFirstChild<DW.WrapThrough>() is not null => "Through",
                _ when anchor.GetFirstChild<DW.WrapTopBottom>() is not null => "TopAndBottom",
                _ => null
            };
        }

        return formatting;
    }
}
