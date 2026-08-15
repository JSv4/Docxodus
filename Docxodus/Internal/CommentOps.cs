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
    internal static readonly XNamespace W15 =
        "http://schemas.microsoft.com/office/word/2012/wordml";

    internal static readonly XNamespace W16Cid =
        "http://schemas.microsoft.com/office/word/2016/wordml/cid";

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

    /// <summary>
    /// True when <paramref name="commentId"/> has at least one live body-side
    /// <c>w:commentReference</c>. A reply can mirror a point comment (reference only) or a ranged
    /// comment, but it cannot faithfully attach to an orphaned definition with no reference at all.
    /// </summary>
    internal static bool HasCommentReference(MainDocumentPart main, string commentId) =>
        ReferenceHostParts(main)
            .Select(p => p.GetXDocument().Root)
            .Where(r => r is not null)
            .SelectMany(r => r!.Descendants(W.commentReference))
            .Any(r => (string?)r.Attribute(W.id) == commentId);

    /// <summary>
    /// Add the body-side reference for a reply immediately after its parent's reference. Word's
    /// native threaded shape keeps the range markers on the thread root and gives replies only an
    /// adjacent <c>w:commentReference</c>; <c>w15:paraIdParent</c> supplies the relationship and
    /// makes every descendant share that root range. This also makes replying to an existing reply
    /// work when (as Word writes it) the immediate parent has no range markers of its own.
    /// </summary>
    internal static IReadOnlyList<XElement> InsertReplyReference(
        MainDocumentPart main, string parentId, int replyId)
    {
        var roots = ReferenceHostParts(main)
            .Select(p => p.GetXDocument().Root)
            .Where(r => r is not null)
            .Select(r => r!)
            .ToList();

        // Materialize before inserting: a new reference must not become input to this same pass.
        var referenceRuns = roots.SelectMany(r => r.Descendants(W.commentReference))
            .Where(e => (string?)e.Attribute(W.id) == parentId)
            .Select(e => e.AncestorsAndSelf(W.r).FirstOrDefault())
            .Where(r => r is not null)
            .Select(r => r!)
            .Distinct()
            .ToList();

        if (referenceRuns.Count == 0)
            throw new InvalidOperationException($"parent comment id {parentId} has no live commentReference");

        // The document-side paragraphs are the blocks whose content changes. Return them so the
        // public mutation envelope can report/patch the same scope callers would re-render.
        var hostParagraphs = referenceRuns
            .Select(r => r.Ancestors(W.p).FirstOrDefault())
            .Where(p => p is not null)
            .Select(p => p!)
            .Distinct()
            .ToList();

        foreach (var parentRun in referenceRuns)
        {
            var replyRun = BuildReferenceRun(replyId);
            UnidHelper.AssignToSelfAndDescendants(replyRun);
            parentRun.AddAfterSelf(replyRun);
        }

        return hostParagraphs;
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
    /// Find-or-create Word's two paraId-keyed comment metadata parts and ensure
    /// <paramref name="comment"/> has entries in both. New ids are deterministic: max existing
    /// eight-hex value + 1 (with a collision-free first-free fallback only at UInt32 overflow).
    /// Existing <c>w15:done</c>/<c>w15:paraIdParent</c> values survive unless the caller explicitly
    /// supplies a replacement. Returns the comment's last-paragraph <c>w14:paraId</c>.
    /// </summary>
    internal static string EnsureThreadingMetadata(
        MainDocumentPart main, XElement comment, string? parentParaId = null, bool? resolved = null)
    {
        var commentsPart = main.WordprocessingCommentsPart
            ?? throw new InvalidOperationException("cannot create comment metadata without comments.xml");
        var commentsRoot = commentsPart.GetXDocument().Root
            ?? throw new InvalidOperationException("comments.xml has no root element");

        var lastPara = comment.Elements(W.p).LastOrDefault();
        if (lastPara is null)
        {
            // Word-authored definitions end in a paragraph because commentsExtended keys on that
            // paragraph. Be total over a malformed/table-only input by adding the smallest legal
            // carrier rather than producing an extension entry that cannot resolve.
            lastPara = new XElement(W.p);
            ApplyCommentBodyStyle(new[] { lastPara });
            comment.Add(lastPara);
            UnidHelper.AssignToSelfAndDescendants(lastPara);
        }

        // A w14 attribute is an ignorable extension to older consumers. Preserve any existing
        // compatibility tokens while making that contract explicit before adding/using paraId.
        if (commentsRoot.GetNamespaceOfPrefix("w14") is null)
            commentsRoot.Add(new XAttribute(XNamespace.Xmlns + "w14", W14.w14));
        EnsureIgnorablePrefix(commentsRoot, "w14", W14.w14);

        var paraId = (string?)lastPara.Attribute(W14.paraId);
        if (string.IsNullOrEmpty(paraId))
        {
            paraId = NextParaId(main);
            lastPara.SetAttributeValue(W14.paraId, paraId);
        }

        var exPart = main.WordprocessingCommentsExPart;
        if (exPart is null)
        {
            exPart = main.AddNewPart<WordprocessingCommentsExPart>();
            exPart.PutXDocument(new XDocument(
                new XElement(W15 + "commentsEx",
                    new XAttribute(XNamespace.Xmlns + "w15", W15))));
        }
        var exRoot = exPart.GetXDocument().Root
            ?? throw new InvalidOperationException("commentsExtended.xml has no root element");
        var commentEx = exRoot.Elements(W15 + "commentEx")
            .FirstOrDefault(e => (string?)e.Attribute(W15 + "paraId") == paraId);
        if (commentEx is null)
        {
            commentEx = new XElement(W15 + "commentEx",
                new XAttribute(W15 + "paraId", paraId),
                new XAttribute(W15 + "done", "0"));
            exRoot.Add(commentEx);
        }
        if (parentParaId is not null)
            commentEx.SetAttributeValue(W15 + "paraIdParent", parentParaId);
        if (resolved.HasValue)
            commentEx.SetAttributeValue(W15 + "done", resolved.Value ? "1" : "0");

        var idsPart = main.WordprocessingCommentsIdsPart;
        if (idsPart is null)
        {
            idsPart = main.AddNewPart<WordprocessingCommentsIdsPart>();
            idsPart.PutXDocument(new XDocument(
                new XElement(W16Cid + "commentsIds",
                    new XAttribute(XNamespace.Xmlns + "w16cid", W16Cid))));
        }
        var idsRoot = idsPart.GetXDocument().Root
            ?? throw new InvalidOperationException("commentsIds.xml has no root element");
        if (!idsRoot.Elements(W16Cid + "commentId")
                .Any(e => (string?)e.Attribute(W16Cid + "paraId") == paraId))
        {
            idsRoot.Add(new XElement(W16Cid + "commentId",
                new XAttribute(W16Cid + "paraId", paraId),
                new XAttribute(W16Cid + "durableId", NextDurableId(idsRoot))));
        }

        commentsPart.PutXDocument();
        exPart.PutXDocument();
        idsPart.PutXDocument();
        return paraId;
    }

    private static void EnsureIgnorablePrefix(XElement root, string prefix, XNamespace ns)
    {
        if (root.GetNamespaceOfPrefix("mc") != MC.mc)
            root.SetAttributeValue(XNamespace.Xmlns + "mc", MC.mc.NamespaceName);
        if (root.GetNamespaceOfPrefix(prefix) != ns)
            root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);

        var tokens = ((string?)root.Attribute(MC.Ignorable) ?? string.Empty)
            .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .ToList();
        if (!tokens.Contains(prefix, StringComparer.Ordinal))
            tokens.Add(prefix);
        root.SetAttributeValue(MC.Ignorable, string.Join(" ", tokens));
    }

    /// <summary>Interpret the OOXML on/off lexical forms accepted by <c>w15:done</c>.</summary>
    internal static bool ParseDone(string? value) =>
        value is "1" or "true" or "on";

    private static string NextParaId(MainDocumentPart main) => new ParaIdAllocator(main).Next();

    /// <summary>
    /// The single owner of package-unique <c>w14:paraId</c> minting. Word keys comment
    /// threading, coauthoring and revision identity off a paraId, so a value must be unique
    /// across every part that can carry one — including the two paraId-keyed comment metadata
    /// parts, whose entries outlive the paragraph they name.
    /// </summary>
    /// <remarks>
    /// A caller minting several ids for a subtree that is still <em>detached</em> from the
    /// package (a clone not yet inserted) must share one allocator instance: each minted value
    /// is retained here, so a later call cannot alias an earlier one. Calling
    /// <see cref="NextParaId"/> repeatedly would return the same value until each is written
    /// back into a live part.
    /// </remarks>
    internal sealed class ParaIdAllocator
    {
        private readonly List<string> _used;

        internal ParaIdAllocator(MainDocumentPart main) => _used = CollectParaIds(main);

        internal string Next()
        {
            var value = NextEightHex(_used);
            _used.Add(value);
            return value;
        }
    }

    private static List<string> CollectParaIds(MainDocumentPart main)
    {
        var values = new List<string>();
        foreach (var part in ReferenceHostParts(main).Append<OpenXmlPart?>(main.WordprocessingCommentsPart))
        {
            var root = part?.GetXDocument().Root;
            if (root is null) continue;
            values.AddRange(root.DescendantsAndSelf().Attributes(W14.paraId).Select(a => (string)a));
        }

        var exRoot = main.WordprocessingCommentsExPart?.GetXDocument().Root;
        if (exRoot is not null)
        {
            values.AddRange(exRoot.Elements(W15 + "commentEx")
                .SelectMany(e => new[]
                {
                    (string?)e.Attribute(W15 + "paraId"),
                    (string?)e.Attribute(W15 + "paraIdParent"),
                })
                .Where(v => !string.IsNullOrEmpty(v))
                .Select(v => v!));
        }
        var idsRoot = main.WordprocessingCommentsIdsPart?.GetXDocument().Root;
        if (idsRoot is not null)
            values.AddRange(idsRoot.Elements(W16Cid + "commentId")
                .Select(e => (string?)e.Attribute(W16Cid + "paraId"))
                .Where(v => !string.IsNullOrEmpty(v)).Select(v => v!));

        return values;
    }

    private static string NextDurableId(XElement idsRoot) =>
        NextEightHex(idsRoot.Elements(W16Cid + "commentId")
            .Select(e => (string?)e.Attribute(W16Cid + "durableId"))
            .Where(v => !string.IsNullOrEmpty(v)).Select(v => v!));

    private static string NextEightHex(IEnumerable<string> values)
    {
        var used = new HashSet<uint>();
        uint max = 0;
        foreach (var value in values)
        {
            if (value.Length != 8
                || !uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var parsed))
                continue;
            used.Add(parsed);
            if (parsed > max) max = parsed;
        }

        uint next;
        if (max < uint.MaxValue)
        {
            next = max + 1;
            if (next != 0 && !used.Contains(next))
                return next.ToString("X8", CultureInfo.InvariantCulture);
        }

        next = 1;
        while (used.Contains(next))
        {
            if (next == uint.MaxValue)
                throw new InvalidOperationException("all eight-hex comment metadata ids are allocated");
            next++;
        }
        return next.ToString("X8", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Re-home the comment markers inside a tracked-move SOURCE block onto freshly cloned comment
    /// definitions, so the source copy and the destination clone no longer share comment ids.
    /// </summary>
    /// <remarks>
    /// A tracked move leaves two live copies of the block. Copying the markers verbatim would give
    /// one comment two ranges and two references — a schema violation (<c>w:id</c> is
    /// uniqueness-constrained) and a visible defect: the comment shows twice in the Reviewing
    /// pane. This mirrors <c>IrMarkupRenderer.NormalizeComments</c> step (B): the DELETED copy —
    /// here the <c>w:moveFrom</c> source — takes the fresh id and a cloned definition, so
    /// accepting keeps the destination's original comment (with its thread intact) and rejecting
    /// keeps the source's clone.
    /// <para>
    /// Every comment referenced from the block is cloned, so a whole thread (a root plus the
    /// replies whose references live in the same paragraph) is cloned together and the clones'
    /// <c>w15:paraIdParent</c> links are re-pointed at the cloned parents rather than at the
    /// originals.
    /// </para>
    /// </remarks>
    internal static void CloneCommentsForMoveSource(MainDocumentPart main, XElement source)
    {
        static bool IsMarker(XElement e) =>
            e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd || e.Name == W.commentReference;

        var markers = source.DescendantsAndSelf().Where(IsMarker).ToList();
        if (markers.Count == 0) return;

        var commentsRoot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (commentsRoot is null) return;

        var oldIds = markers.Select(m => (string?)m.Attribute(W.id))
            .Where(id => id is not null).Distinct().ToList();

        // Original last-paragraph paraId → the clone's, so cloned replies re-point at cloned parents.
        var paraIdMap = new Dictionary<string, string>();
        var clonedParents = new List<(XElement Clone, string OldParentParaId)>();
        var exRoot = main.WordprocessingCommentsExPart?.GetXDocument().Root;

        foreach (var oldId in oldIds)
        {
            var definition = commentsRoot.Elements(W.comment)
                .FirstOrDefault(c => (string?)c.Attribute(W.id) == oldId);
            if (definition is null) continue; // dangling marker: nothing to clone, id stays put

            var oldParaId = (string?)definition.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
            var oldEx = oldParaId is null || exRoot is null
                ? null
                : exRoot.Elements(W15 + "commentEx")
                    .FirstOrDefault(e => (string?)e.Attribute(W15 + "paraId") == oldParaId);

            var clone = new XElement(definition);
            clone.SetAttributeValue(W.id, NextCommentId(main).ToString(CultureInfo.InvariantCulture));
            // Strip the identities the clone must not share: paraId keys the threading parts and
            // Unid keys the projection index. EnsureThreadingMetadata mints a fresh paraId below.
            foreach (var el in clone.DescendantsAndSelf())
            {
                el.Attributes(W14.paraId).Remove();
                el.Attributes(PtOpenXml.Unid).Remove();
            }
            commentsRoot.Add(clone);

            var newId = (string)clone.Attribute(W.id)!;
            foreach (var marker in markers.Where(m => (string?)m.Attribute(W.id) == oldId))
                marker.SetAttributeValue(W.id, newId);

            var newParaId = EnsureThreadingMetadata(
                main, clone, resolved: ParseDone((string?)oldEx?.Attribute(W15 + "done")));
            if (oldParaId is not null) paraIdMap[oldParaId] = newParaId;
            if ((string?)oldEx?.Attribute(W15 + "paraIdParent") is { } oldParent)
                clonedParents.Add((clone, oldParent));
        }

        // Re-point cloned replies at their cloned parent. A parent outside the moved block was not
        // cloned — the reply then becomes top-level rather than dangling, matching the teardown
        // policy in PruneThreadingMetadata.
        foreach (var (clone, oldParent) in clonedParents)
        {
            var paraId = (string?)clone.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
            var entry = paraId is null ? null : main.WordprocessingCommentsExPart?.GetXDocument().Root?
                .Elements(W15 + "commentEx")
                .FirstOrDefault(e => (string?)e.Attribute(W15 + "paraId") == paraId);
            if (entry is null) continue;
            if (paraIdMap.TryGetValue(oldParent, out var newParent))
                entry.SetAttributeValue(W15 + "paraIdParent", newParent);
            else
                entry.Attributes(W15 + "paraIdParent").Remove();
        }
        main.WordprocessingCommentsExPart?.PutXDocument();
    }

    /// <summary>
    /// Prune Word's comment-threading metadata for removed comment definitions:
    /// <c>commentsExtended.xml</c> (<c>w15:commentEx</c>) and <c>commentsIds.xml</c>
    /// (<c>w16cid:commentId</c>) entries key on a definition paragraph's <c>w14:paraId</c>, so a
    /// removed definition's entries must go too — and a surviving reply whose
    /// <c>w15:paraIdParent</c> pointed at a removed comment becomes top-level (the attribute is
    /// dropped) instead of dangling. The parts are never created here; a document without
    /// threading metadata is untouched.
    /// </summary>
    internal static void PruneThreadingMetadata(
        WordprocessingDocument doc, IReadOnlyCollection<string> removedParaIds)
    {
        if (removedParaIds.Count == 0) return;
        var main = doc.MainDocumentPart;
        if (main is null) return;

        PruneEntries(main.WordprocessingCommentsExPart,
            W15 + "commentEx", W15 + "paraId", W15 + "paraIdParent", removedParaIds);
        PruneEntries(main.WordprocessingCommentsIdsPart,
            W16Cid + "commentId", W16Cid + "paraId", parentAttr: null, removedParaIds);
    }

    private static void PruneEntries(
        OpenXmlPart? part, XName entryName, XName paraIdAttr, XName? parentAttr,
        IReadOnlyCollection<string> removedParaIds)
    {
        var root = part?.GetXDocument().Root;
        if (root is null) return;

        bool changed = false;
        foreach (var entry in root.Descendants(entryName).ToList())
        {
            var paraId = (string?)entry.Attribute(paraIdAttr);
            if (paraId is not null && removedParaIds.Contains(paraId))
            {
                entry.Remove();
                changed = true;
                continue;
            }
            if (parentAttr is not null
                && (string?)entry.Attribute(parentAttr) is { } parent
                && removedParaIds.Contains(parent))
            {
                entry.Attribute(parentAttr)!.Remove();
                changed = true;
            }
        }
        if (changed) part!.PutXDocument();
    }

    /// <summary>
    /// Flatten a <c>w:comment</c> body to plain text: per-paragraph <c>w:t</c> concatenation,
    /// paragraphs joined by a single space, runs carrying the <c>w:annotationRef</c> mark
    /// excluded — the same rule the HTML renderer and the markdown projection apply.
    /// </summary>
    internal static string FlattenBodyText(XElement comment)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var p in comment.Elements(W.p))
        {
            var text = string.Concat(
                p.Descendants(W.t)
                    .Where(t => !t.Ancestors(W.r).Any(r => r.Elements(W.annotationRef).Any()))
                    .Select(t => (string)t));
            if (text.Length == 0) continue;
            if (sb.Length > 0) sb.Append(' ');
            sb.Append(text);
        }
        return sb.ToString();
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
