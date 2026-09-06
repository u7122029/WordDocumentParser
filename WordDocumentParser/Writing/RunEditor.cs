using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentParser.Core;
using WordDocumentParser.Models.Formatting;

namespace WordDocumentParser.Writing;

/// <summary>
/// Applies edited <see cref="FormattedRun"/> values back onto the runs of an existing paragraph.
/// </summary>
/// <remarks>
/// <para>
/// The paragraph's own XML is kept and edited in place. Regenerating it from the model instead
/// would discard every structure the model does not represent — hyperlink wrappers, field codes,
/// bookmarks, comment anchors, and drawings — on any edit at all, however unrelated.
/// </para>
/// <para>
/// Model runs are matched to XML runs by character offset. Because both sides describe the same
/// text, a caller's edit maps to an exact range even when the model runs were split — which is what
/// <c>SetFontForText</c> does — and the XML run is split to match rather than the paragraph being
/// rebuilt.
/// </para>
/// </remarks>
internal static class RunEditor
{
    /// <summary>
    /// One text-bearing piece of the paragraph: a single <c>w:t</c>, <c>w:tab</c>, or <c>w:br</c>
    /// together with the run that owns it. Enumerated in the same order, and with the same field
    /// handling, that the parser used to produce the model runs, so the two sides line up.
    /// </summary>
    private sealed class RunPiece
    {
        /// <summary>The run that currently owns this piece. Splitting reassigns it.</summary>
        public required Run Run { get; set; }

        public required OpenXmlElement Content { get; init; }
        public required string Text { get; init; }

        /// <summary>
        /// Every run making up the field this piece stands for — delimiters, code, and result —
        /// in document order, or null when the piece is ordinary content.
        /// </summary>
        /// <remarks>
        /// A field computes its own text, so the whole construct is the unit of deletion and of
        /// formatting; emptying the result element alone would leave the field to regenerate it.
        /// </remarks>
        public List<Run>? FieldRuns { get; init; }

        /// <summary>True when this piece stands for a whole field rather than one text element.</summary>
        public bool IsField => FieldRuns is not null;
    }

    /// <summary>
    /// Applies the node's edited runs and text to a paragraph, leaving everything the caller did not
    /// change exactly as it was.
    /// </summary>
    /// <param name="paragraph">The paragraph to edit in place.</param>
    /// <param name="node">The node holding the edits.</param>
    public static void ApplyRunEdits(Paragraph paragraph, DocumentNode node)
    {
        var pieces = EnumeratePieces(paragraph);

        if (node.IsRunsChanged && node.Runs.Count > 0)
        {
            ApplyModelRuns(paragraph, pieces, node.Runs);
            return;
        }

        // The caller emptied the run collection. That is a deletion, so the paragraph's text goes
        // with it — except where the caller also assigned Text as the replacement.
        if (node.IsRunsChanged && node.Runs.Count == 0)
        {
            ReplaceParagraphText(paragraph, pieces, node.IsTextChanged ? node.Text : string.Empty);
            return;
        }

        // No run-level edits, but the caller assigned the node's plain text.
        if (node.IsTextChanged)
        {
            ReplaceParagraphText(paragraph, pieces, node.Text);
        }
    }

    /// <summary>
    /// Replaces a paragraph's visible text, keeping the first text element (and therefore its run's
    /// formatting) and emptying the rest.
    /// </summary>
    public static void ReplaceParagraphText(Paragraph paragraph, string newText)
        => ReplaceParagraphText(paragraph, EnumeratePieces(paragraph), newText);

    private static void ReplaceParagraphText(Paragraph paragraph, List<RunPiece> pieces, string newText)
    {
        // Fields contribute text the paragraph no longer has. Emptying their result element would
        // not remove them — the field would recompute its value — so the construct goes entirely.
        foreach (var piece in pieces)
        {
            if (piece.FieldRuns is null) continue;

            foreach (var fieldRun in piece.FieldRuns)
            {
                fieldRun.Remove();
            }
        }

        var textPieces = pieces.FindAll(piece => piece.Content is Text && piece.Run.Parent is not null);

        if (textPieces.Count == 0)
        {
            if (string.IsNullOrEmpty(newText)) return;

            var run = paragraph.Elements<Run>().FirstOrDefault();
            if (run is null)
            {
                run = new Run();
                paragraph.Append(run);
            }
            run.Append(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
            return;
        }

        var first = (Text)textPieces[0].Content;
        first.Text = newText;
        first.Space = SpaceProcessingModeValues.Preserve;

        for (var i = 1; i < textPieces.Count; i++)
        {
            ((Text)textPieces[i].Content).Text = string.Empty;
        }
    }

    /// <summary>
    /// Maps each edited model run onto the XML pieces covering the same character range and applies
    /// the properties the caller actually changed.
    /// </summary>
    private static void ApplyModelRuns(Paragraph paragraph, List<RunPiece> pieces, List<FormattedRun> modelRuns)
    {
        var aligned = Align(pieces, modelRuns);

        if (aligned is null)
        {
            // The model's text no longer matches the paragraph's, so offsets cannot be trusted.
            // Rewrite the text through the existing runs and apply formatting positionally, which
            // still keeps hyperlink and field containers intact.
            RewriteWithoutAlignment(paragraph, pieces, modelRuns);
            return;
        }

        for (var i = 0; i < modelRuns.Count; i++)
        {
            var modelRun = modelRuns[i];
            var piece = aligned[i];
            if (piece is null) continue;

            if (modelRun.IsTextChanged && piece.Content is Text text)
            {
                text.Text = modelRun.Text;
                text.Space = SpaceProcessingModeValues.Preserve;
            }

            if (!modelRun.Formatting.HasChanges) continue;

            if (piece.IsField)
            {
                // A field's formatting belongs to the whole construct.
                foreach (var fieldRun in piece.FieldRuns!)
                {
                    ApplyRunFormatting(fieldRun, modelRun.Formatting);
                }
                continue;
            }

            // Run properties apply to everything in the run, so the piece needs a run of its own
            // before its formatting is set. A run holding "abc", a tab and "def" would otherwise
            // hand all three the formatting meant for one.
            IsolatePiece(pieces, piece);
            ApplyRunFormatting(piece.Run, modelRun.Formatting);
        }
    }

    /// <summary>
    /// Ensures a piece's content is the only content in its run, moving whatever preceded and
    /// followed it into runs of their own.
    /// </summary>
    private static void IsolatePiece(List<RunPiece> pieces, RunPiece piece)
    {
        var run = piece.Run;
        var contents = run.ChildElements.Where(child => child is not RunProperties).ToList();
        if (contents.Count <= 1) return;

        var position = contents.IndexOf(piece.Content);
        if (position < 0) return;

        MoveIntoSiblingRun(pieces, run, contents.Take(position).ToList(), before: true);
        MoveIntoSiblingRun(pieces, run, contents.Skip(position + 1).ToList(), before: false);
    }

    /// <summary>
    /// Moves a run's leading or trailing content into a new sibling run that copies its properties,
    /// and repoints the affected pieces at it.
    /// </summary>
    private static void MoveIntoSiblingRun(
        List<RunPiece> pieces, Run run, List<OpenXmlElement> contents, bool before)
    {
        if (contents.Count == 0) return;

        var sibling = new Run();
        if (run.RunProperties is not null)
        {
            sibling.RunProperties = (RunProperties)run.RunProperties.CloneNode(true);
        }

        foreach (var content in contents)
        {
            content.Remove();
            sibling.Append(content);
        }

        if (before)
        {
            run.InsertBeforeSelf(sibling);
        }
        else
        {
            run.InsertAfterSelf(sibling);
        }

        var moved = new HashSet<OpenXmlElement>(contents);
        foreach (var piece in pieces)
        {
            if (ReferenceEquals(piece.Run, run) && moved.Contains(piece.Content))
            {
                piece.Run = sibling;
            }
        }
    }

    /// <summary>
    /// Produces one XML piece per model run, splitting XML runs where a model boundary falls inside
    /// one. Returns null when the two sides describe different text, in which case offsets are
    /// meaningless and the caller falls back to positional rewriting.
    /// </summary>
    private static List<RunPiece?>? Align(List<RunPiece> pieces, List<FormattedRun> modelRuns)
    {
        var pieceText = string.Concat(pieces.Select(p => p.Text));
        var modelText = string.Concat(modelRuns.Select(ModelText));

        if (!string.Equals(pieceText, modelText, StringComparison.Ordinal))
        {
            return null;
        }

        var result = new List<RunPiece?>(modelRuns.Count);
        var pieceIndex = 0;
        var offsetInPiece = 0;

        foreach (var modelRun in modelRuns)
        {
            var needed = ModelText(modelRun).Length;

            // Skip over pieces already fully consumed, and over empty pieces.
            while (pieceIndex < pieces.Count && offsetInPiece >= pieces[pieceIndex].Text.Length &&
                   pieces[pieceIndex].Text.Length > 0)
            {
                pieceIndex++;
                offsetInPiece = 0;
            }

            if (pieceIndex >= pieces.Count)
            {
                result.Add(null);
                continue;
            }

            var piece = pieces[pieceIndex];
            var remaining = piece.Text.Length - offsetInPiece;

            if (needed == 0)
            {
                result.Add(piece);
                continue;
            }

            if (needed >= remaining)
            {
                // The model run covers the rest of this piece. When it covers exactly the rest and
                // the piece starts here, the whole piece belongs to it.
                if (offsetInPiece > 0 && piece.Content is Text partial)
                {
                    piece = SplitTextPiece(pieces, pieceIndex, partial, offsetInPiece);
                }

                result.Add(piece);
                pieceIndex = pieces.IndexOf(piece) + 1;
                offsetInPiece = 0;

                // A model run spanning several pieces keeps only its first; the remainder is left as
                // it was, which is correct because the run's own text is unchanged in that case.
                var consumed = remaining;
                while (consumed < needed && pieceIndex < pieces.Count)
                {
                    consumed += pieces[pieceIndex].Text.Length;
                    pieceIndex++;
                }
            }
            else
            {
                // The model run covers part of this piece: split so the range gets its own run.
                if (piece.Content is Text splittable)
                {
                    if (offsetInPiece > 0)
                    {
                        piece = SplitTextPiece(pieces, pieceIndex, splittable, offsetInPiece);
                        pieceIndex = pieces.IndexOf(piece);
                        offsetInPiece = 0;
                        splittable = (Text)piece.Content;
                    }

                    SplitTextPiece(pieces, pieceIndex, splittable, needed);
                    result.Add(pieces[pieceIndex]);
                    pieceIndex++;
                    offsetInPiece = 0;
                }
                else
                {
                    result.Add(piece);
                    offsetInPiece += needed;
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Splits the text piece at <paramref name="index"/> at <paramref name="offset"/> characters,
    /// giving each half its own run with cloned run properties, and returns the trailing piece.
    /// </summary>
    /// <remarks>
    /// Everything that followed the split point inside the original run moves into the trailing run.
    /// A run can hold several children — text, tabs, breaks, drawings — and leaving them behind puts
    /// them ahead of the text they follow, reordering the paragraph.
    /// </remarks>
    private static RunPiece SplitTextPiece(List<RunPiece> pieces, int index, Text text, int offset)
    {
        var sourceRun = pieces[index].Run;
        var head = text.Text[..offset];
        var tail = text.Text[offset..];

        // Capture the siblings after the split point before the tree is modified.
        var trailingChildren = new List<OpenXmlElement>();
        for (var sibling = text.NextSibling(); sibling is not null; sibling = sibling.NextSibling())
        {
            trailingChildren.Add(sibling);
        }

        text.Text = head;
        text.Space = SpaceProcessingModeValues.Preserve;

        var tailRun = new Run();
        if (sourceRun.RunProperties is not null)
        {
            tailRun.RunProperties = (RunProperties)sourceRun.RunProperties.CloneNode(true);
        }

        var tailText = new Text(tail) { Space = SpaceProcessingModeValues.Preserve };
        tailRun.Append(tailText);

        foreach (var child in trailingChildren)
        {
            child.Remove();
            tailRun.Append(child);
        }

        sourceRun.InsertAfterSelf(tailRun);

        pieces[index] = new RunPiece { Run = sourceRun, Content = text, Text = head };
        var tailPiece = new RunPiece { Run = tailRun, Content = tailText, Text = tail };
        pieces.Insert(index + 1, tailPiece);

        // The pieces for the moved children now belong to the trailing run. A run's pieces are
        // contiguous, so this stops at the first piece owned by a different run.
        for (var i = index + 2; i < pieces.Count && ReferenceEquals(pieces[i].Run, sourceRun); i++)
        {
            pieces[i].Run = tailRun;
        }

        return tailPiece;
    }

    /// <summary>
    /// Applies formatting positionally when offsets cannot be aligned, so containers survive even
    /// though an exact character mapping is unavailable.
    /// </summary>
    private static void RewriteWithoutAlignment(Paragraph paragraph, List<RunPiece> pieces, List<FormattedRun> modelRuns)
    {
        var textPieces = pieces.FindAll(piece => piece.Content is Text);
        var textRuns = modelRuns.FindAll(run => !run.IsTab && !run.IsBreak);

        for (var i = 0; i < textRuns.Count && i < textPieces.Count; i++)
        {
            var target = (Text)textPieces[i].Content;
            target.Text = textRuns[i].Text;
            target.Space = SpaceProcessingModeValues.Preserve;

            if (textRuns[i].Formatting.HasChanges)
            {
                ApplyRunFormatting(textPieces[i].Run, textRuns[i].Formatting);
            }
        }

        // Model gained runs: append them after the last existing run, inheriting its formatting.
        for (var i = textPieces.Count; i < textRuns.Count; i++)
        {
            var run = new Run();
            var template = textPieces.Count > 0 ? textPieces[^1].Run.RunProperties : null;
            if (template is not null)
            {
                run.RunProperties = (RunProperties)template.CloneNode(true);
            }

            if (textRuns[i].Formatting.HasChanges)
            {
                ApplyRunFormatting(run, textRuns[i].Formatting);
            }

            run.Append(new Text(textRuns[i].Text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }

        // Model lost runs: empty the surplus rather than removing the runs, so any container or
        // bookmark anchored to them stays put.
        for (var i = textRuns.Count; i < textPieces.Count; i++)
        {
            ((Text)textPieces[i].Content).Text = string.Empty;
        }
    }

    /// <summary>
    /// Applies only the run properties the caller assigned onto an existing run's <c>w:rPr</c>.
    /// </summary>
    public static void ApplyRunFormatting(Run run, RunFormatting formatting)
    {
        var runProps = run.RunProperties;
        if (runProps is null)
        {
            runProps = new RunProperties();
            run.InsertAt(runProps, 0);
        }

        if (formatting.IsFontChanged)
        {
            var fonts = runProps.GetFirstChild<RunFonts>();
            if (fonts is null)
            {
                fonts = new RunFonts();
                OoxmlOrder.InsertRunProperty(runProps, fonts);
            }

            // Assigning a font name must also clear that slot's theme reference. A theme reference
            // on the same element takes precedence over the explicit name, so leaving it behind
            // would let the theme font win over the one the caller asked for.
            if (formatting.IsChanged(nameof(RunFormatting.FontFamilyAscii)) && formatting.FontFamilyAscii is not null)
            {
                fonts.Ascii = formatting.FontFamilyAscii;
                fonts.AsciiTheme = null;
            }
            if (formatting.IsChanged(nameof(RunFormatting.FontFamily)) && formatting.FontFamily is not null)
            {
                fonts.HighAnsi = formatting.FontFamily;
                fonts.HighAnsiTheme = null;
            }
            if (formatting.IsChanged(nameof(RunFormatting.FontFamilyEastAsia)) && formatting.FontFamilyEastAsia is not null)
            {
                fonts.EastAsia = formatting.FontFamilyEastAsia;
                fonts.EastAsiaTheme = null;
            }
            if (formatting.IsChanged(nameof(RunFormatting.FontFamilyComplexScript)) && formatting.FontFamilyComplexScript is not null)
            {
                fonts.ComplexScript = formatting.FontFamilyComplexScript;
                fonts.ComplexScriptTheme = null;
            }
        }

        SetToggle<Bold>(runProps, formatting, nameof(RunFormatting.Bold), formatting.Bold);
        SetToggle<Italic>(runProps, formatting, nameof(RunFormatting.Italic), formatting.Italic);
        SetToggle<Strike>(runProps, formatting, nameof(RunFormatting.Strike), formatting.Strike);
        SetToggle<DoubleStrike>(runProps, formatting, nameof(RunFormatting.DoubleStrike), formatting.DoubleStrike);
        SetToggle<Caps>(runProps, formatting, nameof(RunFormatting.AllCaps), formatting.AllCaps);
        SetToggle<SmallCaps>(runProps, formatting, nameof(RunFormatting.SmallCaps), formatting.SmallCaps);

        if (formatting.IsChanged(nameof(RunFormatting.Color)))
        {
            SetOrRemove(runProps, formatting.Color, value => new Color { Val = value });
        }

        if (formatting.IsChanged(nameof(RunFormatting.FontSize)))
        {
            SetOrRemove(runProps, formatting.FontSize, value => new FontSize { Val = value });
        }

        if (formatting.IsChanged(nameof(RunFormatting.FontSizeComplexScript)))
        {
            SetOrRemove(runProps, formatting.FontSizeComplexScript, value => new FontSizeComplexScript { Val = value });
        }

        if (formatting.IsChanged(nameof(RunFormatting.Highlight)))
        {
            SetOrRemove(runProps, formatting.Highlight, value =>
                OoxmlEnum.Parse<HighlightColorValues>(value) is { } parsed ? new Highlight { Val = parsed } : null);
        }

        if (formatting.IsChanged(nameof(RunFormatting.Shading)))
        {
            SetOrRemove(runProps, formatting.Shading, value => new Shading { Fill = value });
        }

        if (formatting.IsChanged(nameof(RunFormatting.StyleId)))
        {
            SetOrRemove(runProps, formatting.StyleId, value => new RunStyle { Val = value });
        }

        if (formatting.IsAnyChanged(nameof(RunFormatting.Underline), nameof(RunFormatting.UnderlineStyle)))
        {
            runProps.GetFirstChild<Underline>()?.Remove();
            if (formatting.Underline)
            {
                var underline = new Underline
                {
                    Val = OoxmlEnum.Parse<UnderlineValues>(formatting.UnderlineStyle) ??
                          new EnumValue<UnderlineValues>(UnderlineValues.Single)
                };
                OoxmlOrder.InsertRunProperty(runProps, underline);
            }
        }

        if (formatting.IsAnyChanged(nameof(RunFormatting.Superscript), nameof(RunFormatting.Subscript)))
        {
            runProps.GetFirstChild<VerticalTextAlignment>()?.Remove();
            if (formatting.Superscript || formatting.Subscript)
            {
                OoxmlOrder.InsertRunProperty(runProps, new VerticalTextAlignment
                {
                    Val = formatting.Superscript ? VerticalPositionValues.Superscript : VerticalPositionValues.Subscript
                });
            }
        }

        if (!runProps.HasChildren)
        {
            runProps.Remove();
        }
    }

    private static void SetToggle<T>(RunProperties runProps, RunFormatting formatting, string propertyName, bool enabled)
        where T : OpenXmlLeafElement, new()
    {
        if (!formatting.IsChanged(propertyName)) return;

        runProps.GetFirstChild<T>()?.Remove();
        if (enabled)
        {
            OoxmlOrder.InsertRunProperty(runProps, new T());
        }
    }

    private static void SetOrRemove<T>(RunProperties runProps, string? value, Func<string, T?> factory)
        where T : OpenXmlElement
    {
        runProps.GetFirstChild<T>()?.Remove();
        if (string.IsNullOrEmpty(value)) return;

        var element = factory(value);
        if (element is not null)
        {
            OoxmlOrder.InsertRunProperty(runProps, element);
        }
    }

    private static string ModelText(FormattedRun run) =>
        run.IsTab ? "\t" : run.IsBreak ? " " : run.Text;

    /// <summary>
    /// Walks the paragraph the same way the parser did, so the pieces line up one-for-one with the
    /// model runs it produced: top-level runs, runs inside hyperlinks, runs inside inline content
    /// controls, with field syntax collapsed the same way.
    /// </summary>
    private static List<RunPiece> EnumeratePieces(Paragraph paragraph)
    {
        var pieces = new List<RunPiece>();

        var inField = false;
        string? fieldCode = null;
        var resultRuns = new List<Run>();
        var constructRuns = new List<Run>();
        var fieldPieceStart = 0;

        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    var fieldChar = run.GetFirstChild<FieldChar>();
                    if (fieldChar?.FieldCharType?.Value is { } charType)
                    {
                        if (charType == FieldCharValues.Begin)
                        {
                            inField = true;
                            fieldCode = null;
                            resultRuns.Clear();
                            constructRuns.Clear();
                            fieldPieceStart = pieces.Count;
                        }

                        constructRuns.Add(run);

                        if (charType == FieldCharValues.End)
                        {
                            if (fieldCode is not null && IsDocPropertyField(fieldCode) && resultRuns.Count > 0)
                            {
                                // The parser collapses a DOCPROPERTY field's result into one model
                                // run, so collapse the pieces to match. The whole construct is kept
                                // so formatting reaches every part of it and deletion removes it all.
                                var collapsedText = string.Concat(
                                    pieces.Skip(fieldPieceStart).Select(p => p.Text));
                                pieces.RemoveRange(fieldPieceStart, pieces.Count - fieldPieceStart);
                                pieces.Add(new RunPiece
                                {
                                    Run = resultRuns[0],
                                    Content = resultRuns[0],
                                    Text = collapsedText,
                                    FieldRuns = [.. constructRuns]
                                });
                            }

                            inField = false;
                            fieldCode = null;
                            resultRuns.Clear();
                            constructRuns.Clear();
                        }
                        continue;
                    }

                    if (run.GetFirstChild<FieldCode>() is { } code && inField)
                    {
                        fieldCode = (fieldCode ?? "") + code.Text;
                        constructRuns.Add(run);
                        continue;
                    }

                    if (inField && fieldCode is not null)
                    {
                        resultRuns.Add(run);
                        constructRuns.Add(run);
                    }

                    AddPieces(pieces, run);
                    break;

                case Hyperlink hyperlink:
                    foreach (var hyperlinkRun in hyperlink.Elements<Run>())
                    {
                        AddPieces(pieces, hyperlinkRun);
                    }
                    break;

                case SdtRun sdtRun:
                    foreach (var sdtContentRun in sdtRun.SdtContentRun?.Elements<Run>() ?? [])
                    {
                        AddPieces(pieces, sdtContentRun);
                    }
                    break;
            }
        }

        return pieces;
    }

    private static void AddPieces(List<RunPiece> pieces, Run run)
    {
        foreach (var content in run.ChildElements)
        {
            switch (content)
            {
                case Text text:
                    pieces.Add(new RunPiece { Run = run, Content = text, Text = text.Text });
                    break;
                case TabChar:
                    pieces.Add(new RunPiece { Run = run, Content = content, Text = "\t" });
                    break;
                case Break or CarriageReturn:
                    pieces.Add(new RunPiece { Run = run, Content = content, Text = " " });
                    break;
            }
        }
    }

    private static bool IsDocPropertyField(string fieldCode) =>
        fieldCode.TrimStart().StartsWith("DOCPROPERTY", StringComparison.OrdinalIgnoreCase);
}
