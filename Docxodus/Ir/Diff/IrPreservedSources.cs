#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>Original input markup paired with the accepted comparison sources. The pairings and
/// the anchors that require a preserved copy are shared by section planning and document rendering.</summary>
internal sealed class IrPreservedSources
{
    internal Dictionary<XElement, List<XElement>>? RightOriginals { get; private init; }
    internal Dictionary<XElement, List<XElement>>? LeftOriginals { get; private init; }
    internal Dictionary<XElement, List<XElement>>? SharedRightOriginals { get; private init; }
    internal Dictionary<XElement, XElement>? LeftDeletedInsertions { get; private init; }
    internal IReadOnlySet<string>? ReplacementAnchors { get; private init; }

    internal static IrPreservedSources Read(IrDocument irLeft, IrDocument irRight, WmlDocument left, WmlDocument right)
    {
        if (HasTrackedRevisionMarkup(left))
            return new IrPreservedSources
            {
                LeftOriginals = BuildPreservedOriginalIndex(irLeft, left, carryRemovedTables: false),
                SharedRightOriginals = BuildPreservedOriginalIndex(irRight, right, carryRemovedTables: false),
                LeftDeletedInsertions = BuildLeftDeletedInsertionIndex(irLeft, irRight, left),
            };

        var originals = BuildPreservedOriginalIndex(irRight, right, carryRemovedTables: true);
        HashSet<string>? replacements = null;
        if (originals is { Count: > 0 })
            foreach (var (anchor, block) in irRight.AnchorIndex)
                if (block.Source.Element is { } element && originals.TryGetValue(element, out var group) &&
                    LowersToPreservedReplace(group))
                    (replacements ??= new(StringComparer.Ordinal)).Add(anchor);
        return new IrPreservedSources { RightOriginals = originals, ReplacementAnchors = replacements };
    }

    /// <summary>Every element name <see cref="RevisionProcessor"/> recognizes as tracked-revision markup —
    /// the "does this block carry pre-existing revisions worth preserving?" gate.</summary>
    internal static readonly HashSet<XName> TrackedRevisionNames = new(RevisionProcessor.TrackedRevisionsElements);

    /// <summary>True when any story part that <see cref="RevisionProcessor"/> considers carries tracked
    /// revision markup. The scan is namespace-aware so an alternate prefix for WordprocessingML cannot evade
    /// the one-sided input-revision-preservation gate. It deliberately covers the main document plus headers,
    /// footers, endnotes, and footnotes — a dirty header is just as asymmetric as a dirty body when only the
    /// RIGHT input's revisions would otherwise be preserved.</summary>
    private static bool HasTrackedRevisionMarkup(WmlDocument doc)
    {
        try
        {
            using var ms = new MemoryStream(doc.DocumentByteArray, writable: false);
            using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);
            foreach (var entry in zip.Entries)
            {
                // Mirrors RevisionProcessor.HasTrackedRevisions' story-part scope without opening the
                // document through the SDK before the renderer's package pass. Header/footer part names are
                // generated as word/header*.xml and word/footer*.xml by Word/Open XML SDK; scanning every
                // such part is conservative for an orphaned relationship and never risks asymmetric output.
                bool trackedStoryPart = entry.FullName is "word/document.xml" or "word/endnotes.xml" or "word/footnotes.xml" ||
                    (entry.FullName.StartsWith("word/header", StringComparison.Ordinal) &&
                     entry.FullName.EndsWith(".xml", StringComparison.Ordinal)) ||
                    (entry.FullName.StartsWith("word/footer", StringComparison.Ordinal) &&
                     entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
                if (!trackedStoryPart)
                    continue;

                using var stream = entry.Open();
                using var reader = System.Xml.XmlReader.Create(stream);
                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        TrackedRevisionNames.Contains(XName.Get(reader.LocalName, reader.NamespaceURI)))
                        return true;
                }
            }
            return false;
        }
        catch (Exception e) when (e is InvalidDataException or System.Xml.XmlException)
        {
            return false;
        }
    }

    /// <summary>
    /// Build the accepted-working-element → ORIGINAL right body element(s) map that powers
    /// <c>PreserveInputRevisions</c>. The renderer's IR read normalizes revisions by ACCEPTING the whole
    /// working copy first (see <see cref="IrReaderOptions.RevisionView"/>), so every retained
    /// <c>Source.Element</c> is revision-free; preserving Word-style requires reaching back to the ORIGINAL
    /// elements. Pairing is an in-order two-pointer walk over the two bodies' child sequences that mirrors
    /// the document-level accept's modeled body restructurings — a paragraph whose MARK is deleted merges into
    /// the NEXT paragraph (a fully-deleted paragraph vanishes the same way), while adjacent tables with the
    /// same bidi setting coalesce. A GROUP of original elements grows until its last member no longer
    /// merge-continues, then its accepted visible
    /// text to equal the working block's text. A verified group maps working → [originals] (one working
    /// block may correspond to SEVERAL originals: the mark-deleted members ride along and vanish again on
    /// accept, exactly Word's shape). Any unexplained divergence (removed content controls, trailing
    /// unmatched originals) stops the walk and returns the PARTIAL map — every entry
    /// already emitted was verified in order, and unmapped blocks degrade to accepted-view emission rather
    /// than risking a wrong pairing. Only markup-bearing groups are mapped (a clean block's accepted
    /// emission already equals its original content, so mapping it would change nothing but bytes churn).
    /// <para><b>Note scopes ride the same map.</b> Footnote/endnote definitions pair by <c>w:id</c> (stable
    /// across accept) and each paired definition's child blocks are aligned with the same walk — the note
    /// renderer dispatches through the shared <see cref="RenderBlockOp"/>, so Equal/Insert note blocks then
    /// preserve through the same two emit hooks with zero extra plumbing.</para>
    /// </summary>
    private static Dictionary<XElement, List<XElement>>? BuildPreservedOriginalIndex(
        IrDocument irRight, WmlDocument right, bool carryRemovedTables)
    {
        // The working (accepted) part trees: the IR read pins each part's parsed XDocument in Sources;
        // every body block's Source.Element lives in the w:document tree, note blocks in w:footnotes/w:endnotes.
        XElement? WorkingRoot(XName rootName) =>
            irRight.Sources.Values.Select(xd => xd.Root).FirstOrDefault(r => r?.Name == rootName);

        var workingBody = WorkingRoot(W.document)?.Element(W.body);
        if (workingBody == null)
            return null;

        var map = new Dictionary<XElement, List<XElement>>();
        using (var streamDoc = new OpenXmlMemoryStreamDocument(right))
        using (var wDoc = streamDoc.GetMeasuredWordprocessingDocument())
        {
            var main = wDoc.MainDocumentPart;
            var originalBody = main?.GetXDocument().Root?.Element(W.body);
            if (originalBody == null)
                return null;
            AlignPreservedChildren(workingBody, originalBody, map, carryRemovedTables);

            AlignPreservedNoteScope(WorkingRoot(W.footnotes),
                main?.FootnotesPart?.GetXDocument().Root, W.footnote, map, carryRemovedTables);
            AlignPreservedNoteScope(WorkingRoot(W.endnotes),
                main?.EndnotesPart?.GetXDocument().Root, W.endnote, map, carryRemovedTables);
        }
        return map.Count == 0 ? null : map;
    }

    /// <summary>
    /// Find the deliberately tiny inverse of the dirty-left insertion projection: a raw LEFT whole-block
    /// deletion that Word carries forward as its ORIGINAL insertion when the accepted RIGHT view contains that
    /// block. This is intentionally not an alignment algorithm. A target must be a direct child of the RIGHT
    /// working body's exact ordinal, and the raw LEFT child at that SAME ordinal must be a fully deleted
    /// <c>w:p</c>/<c>w:tbl</c> which has no accepted LEFT block at that ordinal. Matching the block name and
    /// projected text is the final guard. That shape occurs in Word's own reintroduced-deletion redlines; any
    /// shifted, nested, mixed, field-bearing, move/property, or partially bare shape remains on the ordinary
    /// comparer-authored insertion fallback.
    /// </summary>
    private static Dictionary<XElement, XElement>? BuildLeftDeletedInsertionIndex(
        IrDocument irLeft,
        IrDocument irRight,
        WmlDocument left)
    {
        XElement? WorkingBody(IrDocument ir) =>
            ir.Sources.Values.Select(xd => xd.Root).FirstOrDefault(r => r?.Name == W.document)?.Element(W.body);

        var acceptedLeftBody = WorkingBody(irLeft);
        var workingRightBody = WorkingBody(irRight);
        if (acceptedLeftBody == null || workingRightBody == null)
            return null;

        using var streamDoc = new OpenXmlMemoryStreamDocument(left);
        using var wDoc = streamDoc.GetMeasuredWordprocessingDocument();
        var rawLeftBody = wDoc.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (rawLeftBody == null)
            return null;

        // Ordinals intentionally count EVERY direct body element, including comment/bookmark/range leaves.
        // The raw and working bodies may have different lengths after accepting revisions; that is precisely
        // why this is a per-ordinal eligibility check rather than a two-pointer resynchronization walk.
        var rawLeft = rawLeftBody.Elements().ToList();
        var acceptedLeft = acceptedLeftBody.Elements().ToList();
        var workingRight = workingRightBody.Elements().ToList();
        var map = new Dictionary<XElement, XElement>();

        for (int ordinal = 0; ordinal < workingRight.Count && ordinal < rawLeft.Count; ordinal++)
        {
            var target = workingRight[ordinal];
            var candidate = rawLeft[ordinal];
            if ((target.Name != W.p && target.Name != W.tbl) || candidate.Name != target.Name)
                continue;

            // The raw deleted block must not still be represented by an accepted LEFT p/tbl at this position.
            // We deliberately do not search elsewhere: a shifted match could reattribute an unrelated change.
            if (ordinal < acceptedLeft.Count &&
                (acceptedLeft[ordinal].Name == W.p || acceptedLeft[ordinal].Name == W.tbl))
                continue;

            if (!IsProjectableLeftDeletionAsInsertion(candidate))
                continue;
            if (!string.Equals(DeletedProjectionVisibleText(candidate), VisibleText(target), StringComparison.Ordinal))
                continue;

            map[target] = candidate;
        }

        return map.Count == 0 ? null : map;
    }

    /// <summary>Whether one raw LEFT direct body block is safe to turn from native deletion to native insertion.
    /// The source has to be all-and-only deletion markup: paragraph candidates need a deleted paragraph mark;
    /// table candidates need every direct row marked deleted. Fields, moves, property revisions, a foreign
    /// insertion, malformed <c>w:t</c> in deleted content, and every bare paragraph child are rejected.
    /// </summary>
    private static bool IsProjectableLeftDeletionAsInsertion(XElement candidate)
    {
        if (candidate.Name == W.p)
        {
            if (candidate.Element(W.pPr)?.Element(W.rPr)?.Element(W.del) == null)
                return false;
        }
        else if (candidate.Name == W.tbl)
        {
            var rows = candidate.Elements(W.tr).ToList();
            if (rows.Count == 0 || rows.Any(row => row.Element(W.trPr)?.Element(W.del) == null))
                return false;
        }
        else
        {
            return false;
        }

        // A simple/complex field has distinct containment and text-conversion rules. This projection does not
        // expand or rebuild it, so it is safer to leave it to the proven accepted-view whole-block renderer.
        if (candidate.DescendantsAndSelf().Any(e =>
                e.Name == W.fldSimple || e.Name == W.fldChar ||
                e.Name == W.instrText || e.Name == W.delInstrText))
            return false;

        // RevisionProcessor's public tracked-element list does not include the custom-XML move range
        // endpoints even though it transforms them. They are move provenance, never simple deletion
        // provenance, so keep all four on the accepted-view fallback explicitly.
        if (candidate.DescendantsAndSelf().Any(e =>
                e.Name == W.customXmlMoveFromRangeStart || e.Name == W.customXmlMoveFromRangeEnd ||
                e.Name == W.customXmlMoveToRangeStart || e.Name == W.customXmlMoveToRangeEnd))
            return false;

        var tracked = candidate.DescendantsAndSelf()
            .Where(e => TrackedRevisionNames.Contains(e.Name))
            .ToList();
        if (!tracked.Any(e => e.Name == W.del) ||
            tracked.Any(e => e.Name != W.del && e.Name != W.delText))
            return false;

        // `w:delText` is valid only under a deletion wrapper. A raw `w:t` would need a different semantic
        // conversion and proves the candidate is not wholly deleted.
        if (candidate.DescendantsAndSelf(W.delText).Any(t => !t.Ancestors(W.del).Any()) ||
            candidate.DescendantsAndSelf(W.t).Any())
            return false;

        // At the paragraph level no run-level child may be bare: preserving an arbitrary marker/run alongside
        // the converted wrappers would make part of an ostensibly deleted block survive the projection.
        return candidate.DescendantsAndSelf(W.p)
            .All(p => p.Elements().All(child => child.Name == W.pPr || child.Name == W.del));
    }

    /// <summary>The visible text a raw deleted block will expose after its <c>w:delText</c> nodes become
    /// <c>w:t</c>. Called only after <see cref="IsProjectableLeftDeletionAsInsertion"/> verified the strict
    /// all-deleted shape.</summary>
    private static string DeletedProjectionVisibleText(XElement candidate) =>
        string.Concat(candidate.Descendants(W.delText).Select(t => t.Value));

    /// <summary>Pair each working note definition with the original of the SAME <c>w:id</c> (note ids are
    /// untouched by the accept normalization) and align their child blocks — the note-scope leg of
    /// <see cref="BuildPreservedOriginalIndex"/>. Null roots (scope absent on either side) are a no-op.</summary>
    private static void AlignPreservedNoteScope(
        XElement? workingRoot, XElement? originalRoot, XName noteName, Dictionary<XElement, List<XElement>> map,
        bool carryRemovedTables)
    {
        if (workingRoot == null || originalRoot == null)
            return;
        var originalById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var note in originalRoot.Elements(noteName))
            if ((string?)note.Attribute(W.id) is { } id && !originalById.ContainsKey(id))
                originalById[id] = note;
        foreach (var workingNote in workingRoot.Elements(noteName))
            if ((string?)workingNote.Attribute(W.id) is { } id && originalById.TryGetValue(id, out var originalNote))
                AlignPreservedChildren(workingNote, originalNote, map, carryRemovedTables);
    }

    /// <summary>The container-level two-pointer alignment walk of <see cref="BuildPreservedOriginalIndex"/>
    /// (see there for the model and the conservative-bail rules). Adds verified markup-bearing groups to
    /// <paramref name="map"/>; a divergence stops THIS container's walk only (entries already added stand).
    ///
    /// <para>With <paramref name="carryRemovedTables"/>, a table that accepting removes (every row
    /// tracked-deleted, or emptied by a move) joins the group of the block that follows it, as a leading
    /// member, so the redline keeps it in place with its rows still deleted (issue #866). Such a table
    /// contributes no accepted text, so it never takes part in the group's identity check. The LEFT map
    /// does not carry them: its delete-side projection accepts only groups of insertions.</para></summary>
    private static void AlignPreservedChildren(
        XElement workingContainer, XElement originalContainer, Dictionary<XElement, List<XElement>> map,
        bool carryRemovedTables)
    {
        var working = workingContainer.Elements().ToList();
        var original = originalContainer.Elements().ToList();

        int i = 0;
        foreach (var w in working)
        {
            var group = new List<XElement>();

            // AcceptDeletedAndMoveFromParagraphMarks rebuilds body/note containers from p/tbl blocks and
            // retains only body sectPr. Direct leaf annotations (comment/bookmark/range markers, etc.) are
            // therefore discarded. Skip only such leaves, and only when the working side does not carry the
            // same name, so a non-normalized 1:1 marker still has to match exactly.
            SkipAcceptDiscarded(w);
            if (i >= original.Count)
                return;   // originals exhausted early — alignment lost; keep what was verified.

            // AcceptRevisions merges a maximal direct run of clean, same-bidi tables into one table. This is
            // resynchronization only: the accepted working table is already the correct renderer source, and
            // mapping several raw tables back to it could later re-emit a structurally different table run.
            if (TryConsumeCleanTableMerge(w, original, ref i))
                continue;

            // Any remaining non-block child (not an accept-discarded leaf) requires an exact 1:1 name match
            // to stay aligned; it is never preserved itself.
            if (w.Name != W.p && w.Name != W.tbl)
            {
                if (original[i].Name != w.Name)
                    return;
                i++;
                continue;
            }

            string target = VisibleText(w);
            var acc = new System.Text.StringBuilder();
            while (true)
            {
                // Cover a direct leaf that sat between a mark-deleted paragraph and the paragraph it merges
                // into. The accept transform drops it just like a leaf before an ordinary working block.
                SkipAcceptDiscarded(w);
                if (i >= original.Count)
                    return;   // ran out of originals mid-group — alignment lost.
                var o = original[i];
                group.Add(o);
                acc.Append(AcceptedVisibleText(o));
                i++;
                // A paragraph whose MARK is deleted merge-continues into the next original — keep growing.
                if (HasDeletedParagraphMark(o))
                    continue;
                // Group boundary: the last member contributes the block identity. Verify.
                if (o.Name != w.Name || !string.Equals(acc.ToString(), target, StringComparison.Ordinal))
                    return;   // accept semantics we do not model (removed sdt, etc.) — stop.
                break;
            }

            if (group.Any(g => g.Descendants().Any(d => TrackedRevisionNames.Contains(d.Name))))
                map[w] = group;

            // Step over what accept discards. A removed table is carried in the group when asked for; it is
            // dropped with the group if the walk then bails or resynchronizes, exactly as before.
            void SkipAcceptDiscarded(XElement workingBlock)
            {
                while (i < original.Count && IsAcceptDiscarded(original[i], workingBlock))
                {
                    if (carryRemovedTables && original[i].Name == W.tbl)
                        group.Add(original[i]);
                    i++;
                }
            }
        }
    }

    /// <summary>A member of a preserved group that is a table accepting removes: it was carried in front of
    /// its group's block by <see cref="AlignPreservedChildren"/> and is emitted verbatim, rows still deleted,
    /// by the equal, inserted and modified emitters (issue #866). Move, split and merge renderers do not read
    /// preserved groups, so there it is still dropped.</summary>
    internal static bool IsCarriedRemovedTable(XElement member) =>
        member.Name == W.tbl && RevisionProcessor.AcceptRemovesTable(member);

    /// <summary>Whether a modified block whose revision's original is <paramref name="group"/> is drawn as the
    /// conservative whole-block pair, to keep the revision's own tracked changes (the paragraph placement plan).</summary>
    private static bool LowersToPreservedReplace(List<XElement> group) =>
        group.Any(member => !IsCarriedRemovedTable(member) &&
            member.DescendantsAndSelf().Any(e => TrackedRevisionNames.Contains(e.Name)));

    /// <summary>True when accepting revisions leaves nothing of the original child in the working container:
    /// an accept-discarded leaf, or a table accept removes outright (every row tracked-deleted, or emptied
    /// by a move). Without the table case, one wholly-deleted table stopped the walk and every block after
    /// it lost its preserved revisions (issue #837).</summary>
    private static bool IsAcceptDiscarded(XElement original, XElement working) =>
        IsAcceptDiscardedLeaf(original, working) || RevisionProcessor.AcceptRemovesTable(original);

    /// <summary>True when the original child is a direct leaf discarded by the accept transform that rebuilds
    /// body/note block containers. A wrapper with any nested paragraph/table is deliberately NOT skipped: its
    /// topology needs explicit modeling. The name guard preserves exact matching when the working container
    /// still carries the same leaf.</summary>
    private static bool IsAcceptDiscardedLeaf(XElement original, XElement working) =>
        original.Name != working.Name &&
        original.Name != W.p && original.Name != W.tbl && original.Name != W.sectPr &&
        !original.Descendants().Any(e => e.Name == W.p || e.Name == W.tbl);

    /// <summary>Consume exactly the clean table run that
    /// <see cref="RevisionProcessor.AcceptRevisions"/> coalesces into <paramref name="working"/>. The run is
    /// never entered into the preservation map: cloning several raw tables for one accepted source would change
    /// table structure. Failure leaves <paramref name="originalIndex"/> unchanged, so the ordinary strict
    /// 1:1 walk either verifies a non-merged table or bails conservatively.</summary>
    private static bool TryConsumeCleanTableMerge(
        XElement working, IReadOnlyList<XElement> original, ref int originalIndex)
    {
        if (working.Name != W.tbl || originalIndex >= original.Count || original[originalIndex].Name != W.tbl)
            return false;

        bool BidiVisual(XElement table) => table.Element(W.tblPr)?.Element(W.bidiVisual) != null;
        bool bidiVisual = BidiVisual(original[originalIndex]);
        int end = originalIndex + 1;
        while (end < original.Count && original[end].Name == W.tbl && BidiVisual(original[end]) == bidiVisual)
            end++;
        if (end - originalIndex < 2)
            return false;

        var run = original.Skip(originalIndex).Take(end - originalIndex).ToList();
        if (run.Any(table => table.DescendantsAndSelf().Any(e => TrackedRevisionNames.Contains(e.Name))))
            return false;
        if (!string.Equals(VisibleText(working), string.Concat(run.Select(AcceptedVisibleText)), StringComparison.Ordinal))
            return false;
        if (working.Elements(W.tr).Count() != run.Sum(table => table.Elements(W.tr).Count()))
            return false;

        originalIndex = end;
        return true;
    }

    /// <summary>True when a paragraph's MARK is revision-deleted (<c>w:pPr/w:rPr/w:del</c> or
    /// <c>w:moveFrom</c>) — on accept the paragraph merges into the NEXT one (vanishing entirely when its
    /// content is all delete-grade), the one body restructuring the preservation walk models.</summary>
    private static bool HasDeletedParagraphMark(XElement block) =>
        block.Name == W.p &&
        (block.Element(W.pPr)?.Element(W.rPr)?.Elements()
            .Any(e => e.Name == W.del || e.Name == W.moveFrom) ?? false);

    /// <summary>Concatenated <c>w:t</c> text of an (already accepted) working block.</summary>
    private static string VisibleText(XElement block) =>
        string.Concat(block.Descendants(W.t).Select(t => t.Value));

    /// <summary>The accepted-view visible text of an ORIGINAL (markup-bearing) block: every <c>w:t</c> not
    /// inside delete-grade markup (<c>w:del</c> holds <c>w:delText</c>, excluded implicitly; <c>w:moveFrom</c>
    /// holds <c>w:t</c> that accept REMOVES, excluded explicitly).</summary>
    private static string AcceptedVisibleText(XElement block) =>
        string.Concat(block.Descendants(W.t)
            .Where(t => !t.Ancestors(W.moveFrom).Any() && !t.Ancestors(W.del).Any())
            .Select(t => t.Value));

}
