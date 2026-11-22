using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>Places a reviewer's ranges on the surviving text of a paragraph another reviewer
/// replaces whole. Source ordinals survive the package writer's annotation-id renumbering.</summary>
internal static class IrReplacementAnnotations
{
    internal static IEnumerable<(XElement Element, int Offset)> Markers(XElement paragraph)
    {
        // A tab or image occupies one placement unit. Its enclosed text, if any, belongs to
        // another stream, just as in the aligned token spans.
        if (paragraph.Annotation<IrSourceContent>() is not { } source)
            yield break;
        var lengths = source.PlacementWidths();
        int offset = 0;
        foreach (var element in paragraph.Descendants())
        {
            if (lengths.TryGetValue(element, out int length))
                offset += length;
            else if ((IrReader.IsRangeAnnotation(element) || element.Name == W.commentReference) &&
                IrMarkupRenderer.IsRunLevelCommentMarker(element) && element.Ancestors(W.p).FirstOrDefault() == paragraph)
                yield return (element, offset);
        }
    }

    internal static IrNodeList<IrEqualReviewerMarkers> Plan(IrCompositeOp winner, IrNodeList<IrEqualReviewerMarkers> reviewers,
        IrDocument original, IReadOnlyList<(string Author, IrDocument Ir)> documents, IrDiffSettings settings)
    {
        if (winner.Op is not { RequiresWholeParagraphReplace: true, TokenDiff: { } diff, LeftAnchor: { } la, RightAnchor: { } ra } ||
            winner.SourceReviewer < 0 || original.AnchorIndex.GetValueOrDefault(la) is not IrParagraph left ||
            documents[winner.SourceReviewer].Ir.AnchorIndex.GetValueOrDefault(ra) is not IrParagraph right)
            return reviewers;
        var lt = IrTokenText.PlacementTokens(left, settings);
        var rt = IrTokenText.PlacementTokens(right, settings);
        var spans = new List<(int Start, int End, int Shift)>();
        foreach (var op in diff.Ops)
        {
            if (!IrTokenText.KeepsTextUnchanged(op, lt, rt)) continue;
            var (start, end) = IrTokenText.LeftSpanPositions(lt, op);
            int shift = IrTokenText.RightSpanPositions(rt, op).Start - start;
            if (spans.Count > 0 && spans[^1].End == start && spans[^1].Shift == shift)
                spans[^1] = (spans[^1].Start, end, shift);
            else
                spans.Add((start, end, shift));
        }

        IrEqualReviewerMarkers Place(IrEqualReviewerMarkers reviewer)
        {
            if (documents[reviewer.Reviewer].Ir.AnchorIndex.GetValueOrDefault(reviewer.Anchor)?.Source.Element is not { } source)
                return reviewer;
            var placements = new List<IrMarkerPlacement>();
            var ranges = Markers(source).Select((marker, ordinal) => (marker.Element, marker.Offset, ordinal))
                .GroupBy(m => (Bookmark: m.Element.Name == W.bookmarkStart || m.Element.Name == W.bookmarkEnd,
                    Id: (string?)m.Element.Attribute(W.id)));
            foreach (var range in ranges)
            {
                // Consolidation imports reviewer comment definitions only for body references.
                if (!range.Key.Bookmark && !source.Ancestors(W.body).Any()) continue;
                // A range crossing into another paragraph needs a cross-block placement. Here both
                // endpoints and the comment reference must survive together in the same text span.
                var markers = range.ToList();
                if (range.Key.Id is null || !markers.Any(m => m.Element.Name == (range.Key.Bookmark ? W.bookmarkStart : W.commentRangeStart))
                    || !markers.Any(m => m.Element.Name == (range.Key.Bookmark ? W.bookmarkEnd : W.commentRangeEnd))
                    || !range.Key.Bookmark && !markers.Any(m => m.Element.Name == W.commentReference)) continue;
                foreach (var span in spans)
                    if (markers[0].Offset >= span.Start && markers[^1].Offset <= span.End)
                    {
                        placements.AddRange(markers.Select(m => new IrMarkerPlacement(m.ordinal, m.Offset + span.Shift)));
                        break;
                    }
            }
            return placements.Count == 0 ? reviewer : reviewer with
            {
                ReplacementPositions = IrNodeList.From(placements.OrderBy(p => p.Ordinal)),
            };
        }
        return IrNodeList.From(reviewers.Select(Place));
    }
}
