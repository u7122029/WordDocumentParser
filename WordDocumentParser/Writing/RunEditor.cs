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
/// Model runs use source identity when text or run order changes. Each edited slice retains its
/// source formatting and container. Character offsets suffice when the original sequence and text
/// still match, and XML runs are split where a formatting boundary requires it.
/// </para>
/// </remarks>
internal static class RunEditor
{
    /// <summary>
    /// One field in the paragraph — its delimiters, instruction, and result runs in document order.
    /// </summary>
    /// <remarks>
    /// A field computes its own text, so the construct is the unit of both deletion and formatting.
    /// Emptying a result element leaves the field to regenerate its value, which is why removing the
    /// text a field contributed means removing the whole thing.
    /// </remarks>
    private sealed class FieldConstruct
    {
        public required List<Run> Runs { get; init; }

        /// <summary>
        /// True when the parser collapsed this field's result into a single model run, which it does
        /// for <c>DOCPROPERTY</c>. Other fields keep one piece per result element.
        /// </summary>
        public required bool IsCollapsed { get; init; }
    }

    /// <summary>A text, tab, break, or collapsed field, in the parser's enumeration order.</summary>
    private sealed class RunPiece
    {
        /// <summary>The run that currently owns this piece. Splitting reassigns it.</summary>
        public required Run Run { get; set; }

        public required OpenXmlElement Content { get; init; }
        public required string Text { get; init; }

        /// <summary>
        /// This piece's position in the paragraph's enumeration, matching the
        /// <see cref="FormattedRun.SourceOrdinal"/> the parser stamped on the run it produced.
        /// </summary>
        public int Ordinal { get; set; }

        /// <summary>The field this piece's text comes from, or null for ordinary content.</summary>
        public FieldConstruct? Field { get; init; }

        /// <summary>True when this piece stands for a whole collapsed field.</summary>
        public bool IsCollapsedField => Field is { IsCollapsed: true };
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
        RemoveFields(pieces.Select(piece => piece.Field));

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
        // A deleted field can leave the paragraph's text unchanged, so offsets cannot detect it.
        var sourceSequenceChanged = modelRuns.Any(run => run.SourceOrdinal.HasValue) &&
            (modelRuns.Count != pieces.Count ||
             modelRuns.Where((run, index) => run.SourceOrdinal != index).Any() ||
             pieces.Any(piece => piece.Text.Length == 0));
        var aligned = sourceSequenceChanged || HasDeletedField(pieces, modelRuns)
            ? null : Align(pieces, modelRuns);

        if (aligned is null)
        {
            // The model's text no longer matches the paragraph's, so offsets cannot be trusted.
            // Reconcile source identities in the model's order, retaining their XML containers.
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

            if (piece.IsCollapsedField)
            {
                // The model has one run for the whole field, so its formatting belongs to all of it.
                foreach (var fieldRun in piece.Field!.Runs)
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
    /// meaningless and the caller falls back to source identity.
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

        var original = pieces[index];
        pieces[index] = new RunPiece { Run = sourceRun, Content = text, Text = head, Ordinal = original.Ordinal, Field = original.Field };
        var tailPiece = new RunPiece { Run = tailRun, Content = tailText, Text = tail, Ordinal = original.Ordinal, Field = original.Field };
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
    /// Reconciles edited slices by source identity and orders the resulting runs by the model.
    /// </summary>
    private static void RewriteWithoutAlignment(Paragraph paragraph, List<RunPiece> pieces, List<FormattedRun> modelRuns)
    {
        var survivors = new Dictionary<int, List<int>>();
        var output = new List<Run>?[modelRuns.Count];
        for (var i = 0; i < modelRuns.Count; i++)
        {
            if (modelRuns[i].SourceOrdinal is not { } ordinal || ordinal < 0 || ordinal >= pieces.Count) continue;
            if (!survivors.TryGetValue(ordinal, out var indices)) survivors[ordinal] = indices = [];
            indices.Add(i);
        }

        var fields = pieces.Where(piece => piece.Field is not null).GroupBy(piece => piece.Field!).ToList();
        var fieldIndices = fields.ToDictionary(group => group.Key, group => group
            .SelectMany(piece => survivors.GetValueOrDefault(piece.Ordinal) ?? []).Order().ToList());
        RemoveFields(fields.Where(group => fieldIndices[group.Key].Count == 0).Select(group => group.Key));

        foreach (var piece in pieces)
        {
            var indices = survivors.GetValueOrDefault(piece.Ordinal);
            if (piece.Field is { } field && fieldIndices[field].Count == 0) continue;
            if (piece.IsCollapsedField)
            {
                // A collapsed property field retains its computed result and is formatted as a unit.
                foreach (var index in indices ?? [])
                {
                    if (!modelRuns[index].Formatting.HasChanges) continue;
                    foreach (var run in piece.Field!.Runs) ApplyRunFormatting(run, modelRuns[index].Formatting);
                }
                continue;
            }

            if (indices is null)
            {
                if (piece.Content is Text text) text.Text = string.Empty;
                else piece.Content.Remove();
                continue;
            }

            IsolatePiece(pieces, piece);
            // Snapshot before styling the first slice; later slices inherit the original properties.
            var template = indices.Count > 1 ? (Run)piece.Run.CloneNode(true) : null;
            var previous = piece.Run;
            foreach (var index in indices)
            {
                var run = index == indices[0] ? piece.Run : (Run)template!.CloneNode(true);
                if (index != indices[0]) previous.InsertAfterSelf(run);
                SetRunContent(run, modelRuns[index]);
                output[index] = [run];
                previous = run;
            }
        }

        // Field delimiters travel with their results. Interleaving independent content into a field
        // would change which text it computes, so reject that ambiguous structural edit.
        foreach (var (field, indices) in fieldIndices)
        {
            if (indices.Count == 0) continue;
            if (indices[^1] - indices[0] + 1 != indices.Count)
            {
                throw new DocumentPreservationException("paragraph runs", "A field's result runs must remain together.");
            }
            var results = indices.SelectMany(index => output[index] ?? []).ToArray();
            var resultRuns = results.ToHashSet();
            foreach (var index in indices) output[index] = null;
            var runs = new List<Run>();
            var resultIndex = 0;
            for (OpenXmlElement? element = field.Runs[0]; element is not null; element = element.NextSibling())
            {
                // Order cached result slices by the model while leaving the instruction and
                // delimiter slots intact. Collapsed fields have no individually modeled results.
                if (element is Run run) runs.Add(resultRuns.Contains(run) ? results[resultIndex++] : run);
                if (ReferenceEquals(element, field.Runs[^1])) break;
            }
            output[indices[0]] = runs;
        }

        var nextSources = new Run?[modelRuns.Count];
        Run? next = null;
        for (var i = modelRuns.Count - 1; i >= 0; i--)
        {
            nextSources[i] = next;
            if (output[i] is { Count: > 0 } runs) next = runs[0];
        }
        Run? prior = null;
        for (var i = 0; i < modelRuns.Count; i++)
        {
            if (modelRuns[i].SourceOrdinal is not { } ordinal || !survivors.ContainsKey(ordinal))
            {
                var run = new Run();
                var template = prior?.RunProperties ?? nextSources[i]?.RunProperties;
                if (template is not null) run.RunProperties = (RunProperties)template.CloneNode(true);
                SetRunContent(run, modelRuns[i]);
                // Insertions between two runs in the same container stay in that container.
                var parent = prior?.Parent is OpenXmlCompositeElement container &&
                             ReferenceEquals(container, nextSources[i]?.Parent) ? container : paragraph;
                parent.Append(run);
                output[i] = [run];
            }
            if (output[i] is { Count: > 0 } runs) prior = runs[^1];
        }

        var ranks = new Dictionary<OpenXmlElement, int>();
        foreach (var runs in output)
        {
            foreach (var run in runs ?? []) ranks.Add(run, ranks.Count);
        }
        OrderRunContainers(paragraph, ranks);
    }

    /// <summary>Writes one model slice into an isolated run, retaining its original run properties.</summary>
    private static void SetRunContent(Run run, FormattedRun model)
    {
        foreach (var child in run.ChildElements.Where(child => child is not RunProperties).ToList()) child.Remove();
        if (model.IsTab) run.Append(new TabChar());
        else if (model.IsBreak)
        {
            if (model.BreakType == "CarriageReturn") run.Append(new CarriageReturn());
            else run.Append(new Break { Type = OoxmlEnum.Parse<BreakValues>(model.BreakType ?? "textWrapping") });
        }
        else run.Append(new Text(model.Text) { Space = SpaceProcessingModeValues.Preserve });
        if (model.Formatting.HasChanges) ApplyRunFormatting(run, model.Formatting);
    }

    /// <summary>
    /// Orders modeled children inside their existing containers. Unmodeled children retain their
    /// slots, and container properties, relationships, bookmarks, and drawings remain intact.
    /// </summary>
    private static (int First, int Last)? OrderRunContainers(
        OpenXmlElement element, Dictionary<OpenXmlElement, int> ranks)
    {
        if (ranks.TryGetValue(element, out var rank)) return (rank, rank);
        if (element is Run || element is not OpenXmlCompositeElement container) return null;

        var children = container.ChildElements.ToArray();
        var ordered = new List<(OpenXmlElement Element, int First, int Last)>();
        var slots = new List<int>();
        var alreadyOrdered = true;
        for (var i = 0; i < children.Length; i++)
        {
            if (OrderRunContainers(children[i], ranks) is not { } range) continue;
            if (ordered.Count > 0 && ordered[^1].Last >= range.First) alreadyOrdered = false;
            ordered.Add((children[i], range.First, range.Last));
            slots.Add(i);
        }
        if (ordered.Count == 0) return null;
        if (alreadyOrdered) return (ordered[0].First, ordered[^1].Last);
        ordered.Sort((left, right) => left.First.CompareTo(right.First));
        for (var i = 1; i < ordered.Count; i++)
        {
            if (ordered[i - 1].Last >= ordered[i].First)
            {
                throw new DocumentPreservationException("paragraph runs",
                    "Run order interleaves distinct XML containers. Keep each container's runs together.");
            }
        }
        var changed = false;
        for (var i = 0; i < slots.Count; i++)
        {
            changed |= !ReferenceEquals(children[slots[i]], ordered[i].Element);
            children[slots[i]] = ordered[i].Element;
        }
        if (changed)
        {
            container.RemoveAllChildren();
            container.Append(children);
        }
        return (ordered[0].First, ordered[^1].Last);
    }

    /// <summary>
    /// Returns true when a field's every piece is gone from the model, meaning the caller deleted it.
    /// </summary>
    /// <remarks>
    /// Offsets alone cannot see this: a field whose result was empty leaves the paragraph's text
    /// unchanged, so the aligned path would map cleanly and write the deleted field back out.
    /// </remarks>
    private static bool HasDeletedField(List<RunPiece> pieces, List<FormattedRun> modelRuns)
    {
        var fields = pieces.Where(piece => piece.Field is not null).ToList();
        if (fields.Count == 0) return false;

        var sourceOrdinals = modelRuns
            .Where(run => run.SourceOrdinal.HasValue)
            .Select(run => run.SourceOrdinal!.Value)
            .ToHashSet();

        return fields
            .GroupBy(piece => piece.Field!)
            .Any(field => !field.Any(piece => sourceOrdinals.Contains(piece.Ordinal)));
    }

    /// <summary>
    /// Removes each distinct field construct in the sequence, delimiters and all.
    /// </summary>
    private static void RemoveFields(IEnumerable<FieldConstruct?> fields)
    {
        var removed = new HashSet<FieldConstruct>();

        foreach (var field in fields)
        {
            if (field is null || !removed.Add(field)) continue;

            foreach (var run in field.Runs)
            {
                run.Remove();
            }
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
                            CloseField(pieces, fieldPieceStart, fieldCode, resultRuns, constructRuns);

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

                    if (inField)
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

        // Assigned only once the field collapsing above has settled, so these line up with the
        // ordinals the parser stamped on the model runs.
        for (var i = 0; i < pieces.Count; i++)
        {
            pieces[i].Ordinal = i;
        }

        return pieces;
    }

    /// <summary>
    /// Attaches the field construct to the pieces its result produced, collapsing them into one when
    /// the parser would have collapsed the model runs the same way.
    /// </summary>
    private static void CloseField(
        List<RunPiece> pieces, int fieldPieceStart, string? fieldCode,
        List<Run> resultRuns, List<Run> constructRuns)
    {
        if (constructRuns.Count == 0) return;

        var collapse = fieldCode is not null && IsDocPropertyField(fieldCode);
        var field = new FieldConstruct { Runs = [.. constructRuns], IsCollapsed = collapse };

        if (collapse)
        {
            // The parser turns a DOCPROPERTY field's result into one model run, so collapse the
            // pieces to match and keep the two sides aligned.
            var collapsedText = string.Concat(pieces.Skip(fieldPieceStart).Select(p => p.Text));
            pieces.RemoveRange(fieldPieceStart, pieces.Count - fieldPieceStart);
            pieces.Add(new RunPiece
            {
                Run = resultRuns.FirstOrDefault() ?? constructRuns[0],
                Content = resultRuns.FirstOrDefault() ?? constructRuns[0],
                Text = collapsedText,
                Field = field
            });
            return;
        }

        // Every other field keeps one piece per result element, but each still belongs to the
        // construct — so deleting its text deletes the field rather than emptying its result.
        for (var i = fieldPieceStart; i < pieces.Count; i++)
        {
            pieces[i] = new RunPiece
            {
                Run = pieces[i].Run,
                Content = pieces[i].Content,
                Text = pieces[i].Text,
                Field = field
            };
        }
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
