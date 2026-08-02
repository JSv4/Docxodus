#nullable enable

using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Mechanics for native Word comment authoring on <see cref="DocxSession"/> (issue #300):
/// part scaffold, id allocation, the definition/reference run shapes, and threading-metadata
/// pruning. The public ops (<see cref="DocxSession.AddComment"/> et al.) own guards, snapshots
/// and the <see cref="EditResult"/> envelope; everything OOXML-shaped lives here — the
/// <see cref="AnnotationOps"/> split, applied to real <c>w:comment</c> markup instead of the
/// bookmark + custom-XML overlay.
/// </summary>
internal static class CommentOps
{
    /// <summary>The paragraph style id a comment body paragraph wears.</summary>
    internal const string CommentTextStyleId = "CommentText";

    /// <summary>The character style id worn by the reference run and the annotationRef mark.</summary>
    internal const string CommentReferenceStyleId = "CommentReference";

    /// <summary>
    /// Find-or-create the <c>WordprocessingCommentsPart</c>. A part created here gets a bare
    /// <c>w:comments</c> root (there are no Word-reserved comment definitions, unlike notes).
    /// </summary>
    internal static WordprocessingCommentsPart EnsureCommentsPart(MainDocumentPart main)
    {
        var part = main.WordprocessingCommentsPart;
        if (part is not null) return part;

        part = main.AddNewPart<WordprocessingCommentsPart>();
        part.PutXDocument(new XDocument(
            new XElement(W.comments,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "r", R.r))));
        return part;
    }

    /// <summary>
    /// Allocate the next comment id: max existing id + 1, scanning the definitions <em>and</em>
    /// every body-side <c>commentReference</c>/<c>commentRangeStart</c>/<c>commentRangeEnd</c>
    /// across the parts that can host one — a dangling marker whose definition was lost must not
    /// alias a fresh comment. Comments have no reference-order invariant (renderers pair markers
    /// with definitions by id, not position — unlike footnotes), so plain max+1 is safe anywhere
    /// in the document.
    /// </summary>
    internal static int NextCommentId(MainDocumentPart main)
    {
        int max = 0;

        var commentsRoot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (commentsRoot is not null)
            foreach (var c in commentsRoot.Elements(W.comment))
                if (int.TryParse((string?)c.Attribute(W.id), NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) && id > max)
                    max = id;

        foreach (var part in ReferenceHostParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            foreach (var el in root.Descendants())
                if ((el.Name == W.commentReference || el.Name == W.commentRangeStart || el.Name == W.commentRangeEnd)
                    && int.TryParse((string?)el.Attribute(W.id), NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
                    && id > max)
                    max = id;
        }
        return max + 1;
    }

    /// <summary>The parts whose content can carry comment range markers / references.</summary>
    private static IEnumerable<OpenXmlPart> ReferenceHostParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }

    /// <summary>The body-side reference run: <c>w:r[rStyle=CommentReference]/w:commentReference</c>.</summary>
    internal static XElement BuildReferenceRun(int id) =>
        new XElement(W.r,
            new XElement(W.rPr,
                new XElement(W.rStyle, new XAttribute(W.val, CommentReferenceStyleId))),
            new XElement(W.commentReference,
                new XAttribute(W.id, id.ToString(CultureInfo.InvariantCulture))));

    /// <summary>
    /// The definition's mark run — <c>w:annotationRef</c> is the comment analogue of
    /// <c>w:footnoteRef</c>, rendering the comment's marker inside the comment pane.
    /// </summary>
    internal static XElement BuildAnnotationRefRun() =>
        new XElement(W.r,
            new XElement(W.rPr,
                new XElement(W.rStyle, new XAttribute(W.val, CommentReferenceStyleId))),
            new XElement(W.annotationRef));

    /// <summary>
    /// Stamp <c>CommentText</c> on every body paragraph that has no style of its own (a heading
    /// payload keeps its Heading style), and prepend the <c>w:annotationRef</c> mark run to the
    /// first paragraph — the shape Word writes for every comment body.
    /// </summary>
    internal static void ApplyCommentBodyStyle(IReadOnlyList<XElement> paras)
    {
        foreach (var p in paras)
        {
            var pPr = p.Element(W.pPr);
            if (pPr is null)
            {
                pPr = new XElement(W.pPr);
                p.AddFirst(pPr);
            }
            if (pPr.Element(W.pStyle) is null)
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, CommentTextStyleId)));
        }

        var first = paras[0];
        var mark = BuildAnnotationRefRun();
        var firstPPr = first.Element(W.pPr);
        if (firstPPr is not null) firstPPr.AddAfterSelf(mark);
        else first.AddFirst(mark);
    }

    /// <summary>
    /// Format a comment date the way Word writes <c>w:date</c>: UTC, second precision, trailing
    /// <c>Z</c>. An Unspecified-kind value is treated as already-UTC rather than local, so the
    /// output never depends on the machine's timezone.
    /// </summary>
    internal static string FormatDate(System.DateTime date)
    {
        var utc = date.Kind == System.DateTimeKind.Unspecified
            ? System.DateTime.SpecifyKind(date, System.DateTimeKind.Utc)
            : date.ToUniversalTime();
        return utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
    }
}
