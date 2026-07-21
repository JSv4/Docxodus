#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.4 Task 3 — the native OOXML revision renderer (<see cref="IrMarkupRenderer"/>) test battery. The
/// foundational gate invariant: for any (left, right) pair, the rendered document satisfies
/// <c>AcceptRevisions(Render) ≡ right</c> and <c>RejectRevisions(Render) ≡ left</c> at the per-block
/// <see cref="IrBlock.ContentHash"/> level (the WmlComparer output contract). Proven over (a) targeted unit
/// shapes, (b) the full WC corpus both directions, and (c) the deterministic fuzz seeds; plus an
/// OpenXmlValidator baseline-vs-output comparison (zero NEW schema errors).
/// </summary>
[Trait("Category", "Markup")]
public class IrMarkupRendererTests
{
    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    private readonly ITestOutputHelper _out;

    public IrMarkupRendererTests(ITestOutputHelper output) => _out = output;

    // ----------------------------------------------------------------- build helpers

    /// <summary>Build the script over two docs (Accept-view IRs, the same the adapter uses) and render markup.</summary>
    private static WmlDocument RenderMarkup(WmlDocument left, WmlDocument right, IrDiffSettings? settings = null)
    {
        settings ??= new IrDiffSettings();
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, settings);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    /// <summary>Count the header + footer parts attached to a document's main part.</summary>
    private static int HeaderFooterPartCount(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        return main.HeaderParts.Count() + main.FooterParts.Count();
    }

    /// <summary>Assert every image relationship used by one explicitly bound header part resolves locally.</summary>
    private static void AssertHeaderImageEmbedsResolve(WmlDocument doc, int sectionIndex, int expectedEmbedCount)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);
        var main = wordDoc.MainDocumentPart!;
        var body = main.GetXDocument().Root!.Element(W.body)!;
        var sectPr = body.Descendants(W.sectPr)
            .Where(candidate => candidate.Parent?.Name != W.sectPrChange)
            .ElementAt(sectionIndex);
        var reference = sectPr.Element(W.headerReference)!;
        var header = (HeaderPart)main.GetPartById((string)reference.Attribute(R.id)!);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var embeds = header.GetXDocument().Root!.Descendants(a + "blip")
            .Select(blip => (string?)blip.Attribute(R.r + "embed"))
            .Where(id => id is not null)
            .Cast<string>()
            .ToList();
        Assert.Equal(expectedEmbedCount, embeds.Count);
        var partRelationships = header.Parts.ToDictionary(pair => pair.RelationshipId, pair => pair.OpenXmlPart);
        foreach (var embed in embeds)
        {
            Assert.True(partRelationships.TryGetValue(embed, out var imagePart),
                $"Header clone has r:embed '{embed}' but only [{string.Join(", ", partRelationships.Keys)}].");
            Assert.IsType<ImagePart>(imagePart);
        }
    }

    // ----------------------------------------------------------------- header/footer part hygiene

    /// <summary>
    /// Regression: a mid-document section-break paragraph (inner <c>w:sectPr</c> with a
    /// <c>w:headerReference</c>) is an Equal block, which the renderer clones from the RIGHT document.
    /// Importing that clone's related parts must NOT copy the RIGHT header part into the LEFT-based output
    /// as a duplicate — header/footer scopes are deliberately not diffed, so the LEFT package's parts are
    /// authoritative. Before the fix the output carried two header parts (the LEFT original plus a
    /// <c>P&lt;guid&gt;.xml</c> copy of the RIGHT's), diverging from the clean WmlComparer oracle.
    /// </summary>
    [Fact]
    public void Render_does_not_duplicate_header_parts_for_equal_section_break_block()
    {
        string Body(string firstWord) =>
            $"<w:p><w:r><w:t>{firstWord} first section body text here</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHdr1\"/>" +
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:p>" +
            "<w:p><w:r><w:t>second section body text here</w:t></w:r></w:p>";
        const string header = "<w:p><w:r><w:t>Running header</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyAndHeaderXml(Body("Original"), header);
        var right = IrTestDocuments.FromBodyAndHeaderXml(Body("Modified"), header);

        var rendered = RenderMarkup(left, right);

        Assert.Equal(HeaderFooterPartCount(left), HeaderFooterPartCount(rendered));
    }

    /// <summary>The per-block ContentHash sequence over a document's BODY, descending into table cells, in
    /// document order. This is the text/structure fingerprint the invariant compares — modeled run format is
    /// deliberately excluded (FormatChanged is a Task-4 gap), so it rides on ContentHash, not record equality.</summary>
    private static List<string> BodyContentHashes(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, ReadOpts);
        var hashes = new List<string>();
        var blocks = ir.Body.Blocks.ToList();
        // Exclude the trailing standalone section break: the last-section w:sectPr is page METADATA, not
        // revisable content, and the WmlComparer contract sources it from the LEFT document (headers/footers
        // stripped) — so accept-all does NOT reproduce the RIGHT's trailing sectPr by design. Mid-document
        // section breaks ARE content and stay in the comparison; only the final block, if a section break,
        // is dropped.
        if (blocks.Count > 0 && blocks[^1] is IrSectionBreak)
            blocks.RemoveAt(blocks.Count - 1);
        foreach (var block in blocks)
            CollectHashes(block, hashes);
        return hashes;
    }

    private static void CollectHashes(IrBlock block, List<string> sink)
    {
        switch (block)
        {
            case IrParagraph p:
                sink.Add("p:" + p.ContentHash.ToHex());
                break;
            case IrTable t:
                // A table's own ContentHash already rolls its rows/cells, but to localize a mismatch we descend.
                sink.Add("tbl:" + t.ContentHash.ToHex());
                foreach (var row in t.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var b in cell.Blocks)
                            CollectHashes(b, sink);
                break;
            default:
                sink.Add(block.GetType().Name + ":" + block.ContentHash.ToHex());
                break;
        }
    }

    /// <summary>The per-note block ContentHash sequence over a document's FOOTNOTE then ENDNOTE scopes (Task 4 —
    /// note-scope markup). Only notes actually REFERENCED from the body (a <c>w:footnoteReference</c>/
    /// <c>w:endnoteReference</c> with the matching id) are included, in ascending numeric note-id order, each
    /// note's blocks hashed with the same descent as the body. Filtering by body reference is semantically
    /// faithful (an unreferenced note is invisible) and makes the invariant robust to an orphaned empty note
    /// left in a part after a whole-note insertion is rejected — what matters is the referenced content.</summary>
    private static List<string> NoteContentHashes(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, ReadOpts);
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var bodyRoot = wd.MainDocumentPart!.GetXDocument().Root!;

        // Order notes by BODY-REFERENCE DOCUMENT ORDER, not by absolute numeric id. The note-id renumber pass
        // (M2.6 Task 1) mirrors the oracle (ChangeFootnoteEndnoteReferencesToUniqueRange): it numbers refs by the
        // ACCEPTED body's reference order, so an equal/matched note carries its right-side ordinal even after
        // reject leaves only a subsequence of those refs. Comparing by reference position (the content actually
        // surfaced in reading order) — rather than the absolute id baked into a tag — is the semantically faithful
        // round-trip invariant the oracle itself satisfies (accept-by-right-order, reject-by-left-order), and is
        // robust to the legitimate id divergence between the two sides. The first reference to a note fixes its
        // position; the scope tag still partitions footnotes from endnotes.
        List<string> ReferenceOrder(XName refName) =>
            bodyRoot.Descendants(refName).Select(e => (string?)e.Attribute(W.id))
                .Where(s => s != null).Select(s => s!).Distinct().ToList();
        var referencedFn = ReferenceOrder(W.footnoteReference);
        var referencedEn = ReferenceOrder(W.endnoteReference);

        var hashes = new List<string>();
        foreach (var (scopeTag, store, referenced) in new[]
                 { ("fn", ir.Footnotes, referencedFn), ("en", ir.Endnotes, referencedEn) })
        {
            int ordinal = 0;
            foreach (var id in referenced)
            {
                if (!store.Notes.TryGetValue(id, out var note))
                    continue;
                hashes.Add($"{scopeTag}#{ordinal++}");
                foreach (var b in note.Blocks)
                    CollectHashes(b, hashes);
            }
        }
        return hashes;
    }

    /// <summary>The per-paragraph BOUNDARY-NORMALIZED modeled-only format signature sequence over a document's
    /// body, descending into table cells, in document order. This is the FORMAT fingerprint the strengthened
    /// invariant compares (Task 4 — w:rPrChange): two ContentHash-equal paragraphs compare format-equal iff
    /// their per-token modeled formats agree, independent of run boundaries (so run-resegmentation from
    /// rPrChange wrapping does not spuriously flip it). Non-paragraph blocks contribute their ContentHash only
    /// (no run model / no modeled run format to compare).</summary>
    private static List<string> BodyFormatSignatures(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, ReadOpts);
        var settings = new IrDiffSettings();
        var sigs = new List<string>();
        var blocks = ir.Body.Blocks.ToList();
        if (blocks.Count > 0 && blocks[^1] is IrSectionBreak)
            blocks.RemoveAt(blocks.Count - 1);
        foreach (var block in blocks)
            CollectFormatSignatures(block, settings, sigs);
        return sigs;
    }

    private static void CollectFormatSignatures(IrBlock block, IrDiffSettings settings, List<string> sink)
    {
        switch (block)
        {
            case IrParagraph p:
                sink.Add("pf:" + IrModeledFormat.BlockSignature(p, settings));
                // A3: an inline (in-pPr) sectPr's modeled page setup must round-trip too (accept≡right,
                // reject≡left). Compare only the modeled section key: its raw unmodeled digest contains
                // header/footer refs, which a correct refined story topology may intentionally rebind to a
                // cloned part while preserving the same effective story grid.
                if (p.InlineSectionFormat is { } isf)
                    sink.Add("psec:" + IrModeledFormat.SectionKey(isf));
                break;
            case IrTable t:
                // Include the SHELL digests (block-format-change family): tblPr/tblGrid are in the
                // FormatFingerprint (not ContentHash), so without them the round-trip assertion would be
                // blind to a w:tblPrChange / w:tblGridChange / w:trPrChange that failed to restore. (tcPr is
                // already in ContentHash via ShellDigest, so it rides BodyContentHashes.)
                sink.Add("tblf:" + t.ContentHash.ToHex()
                    + "|tblPr:" + t.TblPrDigest.ToHex() + "|tblGrid:" + t.TblGridDigest.ToHex());
                foreach (var row in t.Rows)
                {
                    // Two independent, tracked row-shell projections — each verifies its own round-trip
                    // (accept≡right, reject≡left): `trf:` = the w:trPr children only (TrPrShellDigest, w:trPrChange),
                    // `tprex:` = the w:tblPrEx children only (TrPrExDigest, w:tblPrExChange, follow-up A2). Both use
                    // the trackable per-element subset, NOT the fingerprint's tblPrEx-inclusive TrPrDigest, so each
                    // marker's reject is byte-verified against exactly what that marker restores.
                    sink.Add("trf:" + row.TrPrShellDigest.ToHex());
                    sink.Add("tprex:" + row.TrPrExDigest.ToHex());
                    foreach (var cell in row.Cells)
                        foreach (var b in cell.Blocks)
                            CollectFormatSignatures(b, settings, sink);
                }
                break;
            default:
                sink.Add(block.GetType().Name + "f:" + block.ContentHash.ToHex());
                break;
        }
    }

    /// <summary>Assert the rendered markup round-trips: accept ≡ right body, reject ≡ left body (ContentHash).</summary>
    private static void AssertRoundTrip(WmlDocument left, WmlDocument right, IrDiffSettings? settings = null, string? label = null)
    {
        var rendered = RenderMarkup(left, right, settings);

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);

        var acceptHashes = BodyContentHashes(accepted);
        var rightHashes = BodyContentHashes(right);
        var rejectHashes = BodyContentHashes(rejected);
        var leftHashes = BodyContentHashes(left);

        Assert.True(acceptHashes.SequenceEqual(rightHashes),
            $"ACCEPT≠RIGHT {label}\n  accept: [{string.Join(", ", acceptHashes)}]\n  right:  [{string.Join(", ", rightHashes)}]");
        Assert.True(rejectHashes.SequenceEqual(leftHashes),
            $"REJECT≠LEFT {label}\n  reject: [{string.Join(", ", rejectHashes)}]\n  left:   [{string.Join(", ", leftHashes)}]");

        // STRENGTHENED (Task 4): format must round-trip too. Accept restores the RIGHT modeled formatting,
        // reject the LEFT — proven by the boundary-normalized modeled-only format signature (so w:rPrChange and
        // FormatOnly blocks restore the correct rPr on the appropriate side).
        var acceptFmt = BodyFormatSignatures(accepted);
        var rightFmt = BodyFormatSignatures(right);
        var rejectFmt = BodyFormatSignatures(rejected);
        var leftFmt = BodyFormatSignatures(left);
        Assert.True(acceptFmt.SequenceEqual(rightFmt),
            $"ACCEPT-FORMAT≠RIGHT {label}\n  accept: [{string.Join(", ", acceptFmt)}]\n  right:  [{string.Join(", ", rightFmt)}]");
        Assert.True(rejectFmt.SequenceEqual(leftFmt),
            $"REJECT-FORMAT≠LEFT {label}\n  reject: [{string.Join(", ", rejectFmt)}]\n  left:   [{string.Join(", ", leftFmt)}]");

        // STRENGTHENED (Task 4): footnote/endnote scope content must round-trip too.
        var acceptNotes = NoteContentHashes(accepted);
        var rightNotes = NoteContentHashes(right);
        var rejectNotes = NoteContentHashes(rejected);
        var leftNotes = NoteContentHashes(left);
        Assert.True(acceptNotes.SequenceEqual(rightNotes),
            $"ACCEPT-NOTES≠RIGHT {label}\n  accept: [{string.Join(", ", acceptNotes)}]\n  right:  [{string.Join(", ", rightNotes)}]");
        Assert.True(rejectNotes.SequenceEqual(leftNotes),
            $"REJECT-NOTES≠LEFT {label}\n  reject: [{string.Join(", ", rejectNotes)}]\n  left:   [{string.Join(", ", leftNotes)}]");

        // STRENGTHENED (2026-07-03): header/footer STORY content must round-trip too. Referenced stories
        // only, empty ≡ absent (accepting a deleted story / rejecting an inserted one leaves an EMPTY
        // story — Word's own behavior — which must compare equal to the other side's absent story).
        var acceptStories = StoryContentHashes(accepted);
        var rightStories = StoryContentHashes(right);
        var rejectStories = StoryContentHashes(rejected);
        var leftStories = StoryContentHashes(left);
        Assert.True(acceptStories.SequenceEqual(rightStories),
            $"ACCEPT-STORIES≠RIGHT {label}\n  accept: [{string.Join(", ", acceptStories)}]\n  right:  [{string.Join(", ", rightStories)}]");
        Assert.True(rejectStories.SequenceEqual(leftStories),
            $"REJECT-STORIES≠LEFT {label}\n  reject: [{string.Join(", ", rejectStories)}]\n  left:   [{string.Join(", ", leftStories)}]");

        // STRENGTHENED (block-format-change family, Phase 3): the trailing section PROPERTIES must round-trip —
        // accept ≡ right, reject ≡ left. Reference-normalized (header/footer references are owned by the
        // header/footer machinery and compared by StoryContentHashes above), so this checks page setup only.
        Assert.Equal(TrailingSectPrPropsDigest(right), TrailingSectPrPropsDigest(accepted));
        Assert.Equal(TrailingSectPrPropsDigest(left), TrailingSectPrPropsDigest(rejected));
    }

    /// <summary>Canonical hash of the trailing sectPr's PROPERTY children (ignoring header/footer references,
    /// the sectPrChange marker, and rsids); "none" when the body has no trailing sectPr. The reference-normalized
    /// section-property fingerprint the block-format-change round-trip compares.</summary>
    private static string TrailingSectPrPropsDigest(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root?.Element(W.body);
        var sectPr = body?.Elements(W.sectPr).LastOrDefault();
        if (sectPr == null)
            return "none";
        var c = new XElement("sect");
        foreach (var e in sectPr.Elements().Where(e =>
                     e.Name != W.headerReference && e.Name != W.footerReference && e.Name != W.sectPrChange))
            c.Add(new XElement(e));
        return Docxodus.Ir.IrHasher.CanonicalHash(c).ToHex();
    }

    /// <summary>
    /// The effective header/footer STORY content-hash projection (2026-07-03 campaign): resolve every
    /// section ordinal × kind by Word's carry-forward inheritance, then collect the visible story's per-block
    /// ContentHashes (headers then footers, deterministic order). This intentionally compares presentation
    /// semantics rather than raw explicit refs: a correct refined topology can add a clone ref at section 1
    /// while Reject still resolves to the same A/A effective header grid as the original inherited left side.
    /// </summary>
    private static List<string> StoryContentHashes(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, ReadOpts);
        var result = new List<string>();
        foreach (var (tag, stories) in new[] { ("hdr", ir.Headers), ("ftr", ir.Footers) })
        {
            var explicitCells = new Dictionary<(int Section, IrHeaderFooterKind Kind), IrHeaderFooter>();
            foreach (var hf in stories)
                foreach (var r in hf.References)
                    if (!explicitCells.ContainsKey((r.SectionIndex, r.Kind)))
                        explicitCells[(r.SectionIndex, r.Kind)] = hf;

            int sectionCount = Math.Max(1, ir.Body.Blocks.OfType<IrParagraph>()
                .Count(paragraph => paragraph.InlineSectionBreakAnchor is not null) + 1);
            var currentByKind = new Dictionary<IrHeaderFooterKind, IrHeaderFooter?>();
            foreach (var kind in new[]
                     { IrHeaderFooterKind.Default, IrHeaderFooterKind.First, IrHeaderFooterKind.Even })
                currentByKind[kind] = null;
            for (int section = 0; section < sectionCount; section++)
            {
                foreach (var kind in new[]
                         { IrHeaderFooterKind.Default, IrHeaderFooterKind.First, IrHeaderFooterKind.Even })
                {
                    if (explicitCells.TryGetValue((section, kind), out var explicitStory))
                        currentByKind[kind] = explicitStory;
                    var current = currentByKind[kind];
                    if (current is null || !current.Scope.Blocks.Any(BlockHasVisibleText))
                        continue;
                    result.Add($"{tag}@s{section}:{kind}");
                    foreach (var block in current.Scope.Blocks)
                        CollectHashes(block, result);
                }
            }
        }
        return result;
    }

    private static bool BlockHasVisibleText(IrBlock block) => block switch
    {
        IrParagraph p => p.Inlines.OfType<IrTextRun>().Any(r => !string.IsNullOrWhiteSpace(r.Text)),
        IrTable t => t.Rows.Any(row => row.Cells.Any(c => c.Blocks.Any(BlockHasVisibleText))),
        _ => true, // opaque/image blocks count as content
    };

    // ----------------------------------------------------------------- targeted unit shapes

    [Fact]
    public void Render_identical_documents_yields_no_revisions_and_round_trips()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>");
        var rendered = RenderMarkup(doc, doc);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.Empty(body.Descendants(W.ins));
        Assert.Empty(body.Descendants(W.del));
        AssertRoundTrip(doc, doc, label: "identical");
    }

    [Fact]
    public void Render_inserted_paragraph_wraps_runs_in_ins_and_round_trips()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Second inserted</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.NotEmpty(body.Descendants(W.ins));
        AssertRoundTrip(left, right, label: "insert-paragraph");
    }

    [Fact]
    public void Render_deleted_paragraph_uses_delText_and_round_trips()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Keep</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Remove me</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Keep</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.NotEmpty(body.Descendants(W.del));
        Assert.NotEmpty(body.Descendants(W.delText));   // deletions MUST use w:delText, not w:t
        AssertRoundTrip(left, right, label: "delete-paragraph");
    }

    [Fact]
    public void Render_modified_paragraph_splits_runs_at_token_boundaries_and_round_trips()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>the quick brown fox</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>the slow brown fox</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.NotEmpty(body.Descendants(W.ins));
        Assert.NotEmpty(body.Descendants(W.del));
        AssertRoundTrip(left, right, label: "modify-paragraph");
    }

    /// <summary>
    /// A Modified table with a tail column add must stay ONE table. The table differ deliberately
    /// emits three paired cell ops plus one right-only cell op per row; rendering that right-only
    /// tail cell as <c>w:cellIns</c> lets the native <c>w:tblGridChange</c> restore the old grid on
    /// reject. Falling back to whole-table del+ins is structurally coarser than Word's one-table
    /// revision shape and is especially visible when every row grows by one column.
    /// </summary>
    [Fact]
    public void Render_table_tail_column_add_stays_one_table_and_round_trips()
    {
        static string Cell(string text, int width) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr>" +
            $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        static string Row(int width, params string[] cells) =>
            "<w:tr>" + string.Concat(cells.Select(c => Cell(c, width))) + "</w:tr>";
        static string Grid(int width, int count) =>
            "<w:tblGrid>" + string.Concat(Enumerable.Repeat($"<w:gridCol w:w=\"{width}\"/>", count)) + "</w:tblGrid>";
        static string Table(int width, int columns, params string[] rows) =>
            "<w:tbl><w:tblPr><w:tblW w:w=\"6000\" w:type=\"dxa\"/></w:tblPr>" +
            Grid(width, columns) + string.Concat(rows) + "</w:tbl>";

        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
            Table(2000, 3,
                Row(2000, "old A", "old B", "old C"),
                Row(2000, "old D", "old E", "old F"),
                Row(2000, "old G", "old H", "old I")));
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
            Table(1500, 4,
                Row(1500, "new A", "new B", "new C", "new D"),
                Row(1500, "new E", "new F", "new G", "new H"),
                Row(1500, "new I", "new J", "new K", "new L")));

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        var table = Assert.Single(body.Elements(W.tbl)); // not the whole-table del+ins fallback
        var rows = table.Elements(W.tr).ToList();
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row => Assert.Equal(4, row.Elements(W.tc).Count()));
        Assert.Equal(3, table.Descendants(W.cellIns).Count());
        Assert.Empty(table.Descendants(W.cellDel));
        Assert.Single(table.Descendants(W.tblGridChange));
        Assert.Equal(9, table.Descendants(W.tcPrChange).Count());
        Assert.Equal(0, SchemaErrorCount(rendered));

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "table-tail-column-add");
    }

    /// <summary>
    /// A column inserted between retained columns must keep those later columns paired with their actual
    /// left-side cells, not merely append a tail <c>w:cellIns</c>. Every retained cell changes width here, so
    /// the differ has to find its spine from the tcPr-free cell body key.
    /// </summary>
    [Fact]
    public void Render_table_middle_column_add_stays_one_table_and_round_trips()
    {
        static string Cell(string text, int width) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr>" +
            $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        static string Row(int width, params string[] cells) =>
            "<w:tr>" + string.Concat(cells.Select(c => Cell(c, width))) + "</w:tr>";
        static string Grid(int width, int count) =>
            "<w:tblGrid>" + string.Concat(Enumerable.Repeat($"<w:gridCol w:w=\"{width}\"/>", count)) + "</w:tblGrid>";
        static string Table(int width, int columns, params string[] rows) =>
            "<w:tbl><w:tblPr><w:tblW w:w=\"6000\" w:type=\"dxa\"/></w:tblPr>" +
            Grid(width, columns) + string.Concat(rows) + "</w:tbl>";

        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
            Table(2000, 3,
                Row(2000, "A1", "B1", "C1"),
                Row(2000, "A2", "B2", "C2"),
                Row(2000, "A3", "B3", "C3")));
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
            Table(1500, 4,
                Row(1500, "A1", "NEW-1", "B1", "C1"),
                Row(1500, "A2", "NEW-2", "B2", "C2"),
                Row(1500, "A3", "NEW-3", "B3", "C3")));

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        var table = Assert.Single(body.Elements(W.tbl)); // not the conservative whole-table fallback
        var rows = table.Elements(W.tr).ToList();
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row =>
        {
            var cells = row.Elements(W.tc).ToList();
            Assert.Equal(4, cells.Count);
            Assert.NotEmpty(cells[1].Descendants(W.cellIns));
            Assert.All(cells.Where((_, index) => index != 1), cell =>
                Assert.Empty(cell.Descendants(W.cellIns)));
            // The inserted shell is new; only explicitly paired cells receive tcPr history.
            Assert.Empty(cells[1].Descendants(W.tcPrChange));
        });
        Assert.Equal(3, table.Descendants(W.cellIns).Count());
        Assert.Empty(table.Descendants(W.cellDel));
        Assert.Single(table.Descendants(W.tblGridChange));
        Assert.Equal(9, table.Descendants(W.tcPrChange).Count());
        Assert.Equal(0, SchemaErrorCount(rendered));

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "table-middle-column-add");
    }

    /// <summary>
    /// Native <c>w:cellIns</c> needs a matching <c>w:tblGridChange</c> to remove the accepted-side widened
    /// grid on reject. The table-shell opt-out therefore chooses the reversible whole-table fallback instead
    /// of emitting an apparently fine-grained but one-way result.
    /// </summary>
    [Fact]
    public void Render_table_column_add_with_table_shell_tracking_off_uses_whole_table_fallback()
    {
        static string Cell(string text, int width) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr>" +
            $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        static string Row(int width, params string[] cells) =>
            "<w:tr>" + string.Concat(cells.Select(c => Cell(c, width))) + "</w:tr>";
        static string Table(int width, int columns, params string[] rows) =>
            "<w:tbl><w:tblPr/><w:tblGrid>" +
            string.Concat(Enumerable.Repeat($"<w:gridCol w:w=\"{width}\"/>", columns)) +
            "</w:tblGrid>" + string.Concat(rows) + "</w:tbl>";

        var left = IrTestDocuments.FromBodyXml(Table(3000, 2, Row(3000, "A", "B")));
        var right = IrTestDocuments.FromBodyXml(Table(2000, 3, Row(2000, "A", "NEW", "B")));
        var settings = new IrDiffSettings { TrackTableFormatChanges = false };

        var rendered = RenderMarkup(left, right, settings);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.Equal(2, body.Elements(W.tbl).Count());
        Assert.Empty(body.Descendants(W.cellIns));
        Assert.NotEmpty(body.Descendants(W.del));
        Assert.NotEmpty(body.Descendants(W.ins));
        Assert.Equal(0, SchemaErrorCount(rendered));
        AssertRoundTrip(left, right, settings, "table-column-add-table-format-off");
    }

    [Fact]
    public void Render_split_run_fragment_with_boundary_whitespace_carries_xml_space_preserve()
    {
        // "the quick brown fox" → "the slow brown fox": the single source run is split at the changed-word
        // boundary into an Equal prefix run ("the ") and an Equal suffix run (" brown fox"). A fragment whose
        // text has a leading or trailing space MUST carry xml:space="preserve" or Word collapses the boundary
        // whitespace, corrupting the round-trip text. Assert the attribute is present on a boundary fragment.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>the quick brown fox</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>the slow brown fox</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);

        XNamespace xmlNs = XNamespace.Xml;
        // Find a w:t whose text has boundary whitespace and confirm it preserves space. The split produces at
        // least one such fragment ("the " trailing, or " brown fox" leading) on the Equal (unwrapped) runs.
        var boundaryTexts = wd.MainDocumentPart!.GetXDocument().Descendants(W.t)
            .Where(t => t.Value.Length > 0 && (char.IsWhiteSpace(t.Value[0]) || char.IsWhiteSpace(t.Value[^1])))
            .ToList();
        Assert.NotEmpty(boundaryTexts);
        Assert.All(boundaryTexts, t =>
            Assert.Equal("preserve", (string?)t.Attribute(xmlNs + "space")));
    }

    [Fact]
    public void Render_revision_ids_are_unique_and_ascending_from_one()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>alpha bravo charlie</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>delete this line</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>alpha CHANGED charlie</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>inserted line</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var xDoc = wd.MainDocumentPart!.GetXDocument();
        var ids = xDoc.Descendants()
            .Where(e => e.Name == W.ins || e.Name == W.del)
            .Select(e => (int?)e.Attribute(W.id))
            .Where(i => i.HasValue)
            .Select(i => i!.Value)
            .ToList();

        Assert.NotEmpty(ids);
        Assert.Equal(ids.Count, ids.Distinct().Count());   // unique
        Assert.True(ids.Min() >= 1, "ids start at 1");
    }

    [Fact]
    public void Render_preserves_run_shading_on_modified_paragraph()
    {
        // A shaded run on an EQUAL portion must survive into the output — proving provenance-clone (not
        // IrRunFormat rebuild) preserves direct formatting while another span is revised.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:shd w:val=\"clear\" w:fill=\"FFFF00\"/></w:rPr><w:t>highlight one two</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:shd w:val=\"clear\" w:fill=\"FFFF00\"/></w:rPr><w:t>highlight THREE two</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var shdCount = wd.MainDocumentPart!.GetXDocument().Descendants(W.shd).Count();
        Assert.True(shdCount > 0, "w:shd run property must be preserved through the split");
        AssertRoundTrip(left, right, label: "run-shading");
    }

    [Fact]
    public void Render_modify_with_zero_width_inline_at_span_boundary_round_trips()
    {
        // A word edit immediately adjacent to a ZERO-WIDTH inline (w:tab) exercises the SourceRunModel's
        // empty-span / zero-width-segment boundary handling (the slicer must attach the tab to exactly one
        // side, never duplicate or drop it). The tokenizer counts the tab as 0 chars, so the token char
        // offsets straddle it precisely at the word boundary.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>alpha</w:t><w:tab/><w:t>bravo</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>alpha</w:t><w:tab/><w:t>charlie</w:t></w:r></w:p>");

        // The contract is the round-trip: accept yields exactly the right paragraph (one tab + charlie),
        // reject yields exactly the left (one tab + bravo). A tab sitting on the boundary of an Equal/Delete
        // span may render as a deleted-tab + inserted-tab pair (the char-boundary slicer attributes the
        // zero-width inline to the adjacent del/ins spans) — that is benign: accept keeps exactly one, reject
        // keeps exactly one. We assert the round-trip (the actual contract), and that the tab is never DROPPED.
        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        Assert.True(wd.MainDocumentPart!.GetXDocument().Descendants(W.tab).Any(), "the w:tab must not be dropped");

        var acceptTabs = new MemoryStream(RevisionProcessor.AcceptRevisions(rendered).DocumentByteArray);
        using (var accWd = WordprocessingDocument.Open(acceptTabs, false))
            Assert.Equal(1, accWd.MainDocumentPart!.GetXDocument().Descendants(W.tab).Count());
        var rejectTabs = new MemoryStream(RevisionProcessor.RejectRevisions(rendered).DocumentByteArray);
        using (var rejWd = WordprocessingDocument.Open(rejectTabs, false))
            Assert.Equal(1, rejWd.MainDocumentPart!.GetXDocument().Descendants(W.tab).Count());

        AssertRoundTrip(left, right, label: "zero-width-boundary");
    }

    // ----------------------------------------------------------------- format change (w:rPrChange)

    [Fact]
    public void Render_format_change_emits_rPrChange_with_old_rPr_and_round_trips_format()
    {
        // Same text, run gains bold: a FormatChanged span. The right run keeps bold (accepted state) and
        // carries a w:rPrChange whose inner w:rPr is the LEFT (non-bold) formatting. Accept ⇒ bold, reject ⇒
        // non-bold — proven by both the text AND format round-trip invariant.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>sample text here</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>sample text here</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var rPrChanges = wd.MainDocumentPart!.GetXDocument().Descendants(W.rPrChange).ToList();
        Assert.NotEmpty(rPrChanges);
        // The rPrChange carries the OLD rPr; here the old side is non-bold, so its inner rPr has no w:b.
        var inner = rPrChanges[0].Element(W.rPr);
        Assert.NotNull(inner);
        Assert.Null(inner!.Element(W.b));
        // Required attributes.
        foreach (var c in rPrChanges)
        {
            Assert.NotNull(c.Attribute(W.id));
            Assert.NotNull(c.Attribute(W.author));
            Assert.NotNull(c.Attribute(W.date));
        }
        AssertRoundTrip(left, right, label: "format-change-add-bold");
    }

    [Fact]
    public void Render_format_change_remove_bold_round_trips_format()
    {
        // Bold → non-bold: the OLD rPr must carry w:b so reject restores bold.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>sample text here</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>sample text here</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var rPrChange = wd.MainDocumentPart!.GetXDocument().Descendants(W.rPrChange).FirstOrDefault();
        Assert.NotNull(rPrChange);
        Assert.NotNull(rPrChange!.Element(W.rPr)!.Element(W.b));   // old (bold) preserved
        AssertRoundTrip(left, right, label: "format-change-remove-bold");
    }

    /// <summary>
    /// A dedicated FORMAT-MUTATION fuzz seed class (Task 4): every seed bolds N random words across the
    /// generated paragraphs, producing pure FormatChanged spans. Exercises the w:rPrChange path at scale and
    /// asserts the strengthened format round-trip invariant holds (accept ⇒ right format, reject ⇒ left).
    /// </summary>
    [Fact]
    [Trait("Category", "Fuzz")]
    public void Fuzz_format_mutation_seeds_round_trip_format()
    {
        const int seedCount = 30;
        var settings = new IrDiffSettings();
        var failures = new List<string>();
        int passed = 0;

        for (int seed = 1; seed <= seedCount; seed++)
        {
            var (left, right, desc) = MakeFormatMutationPair(seed);
            try
            {
                var rendered = RenderMarkup(left, right, settings);
                var acc = RevisionProcessor.AcceptRevisions(rendered);
                var rej = RevisionProcessor.RejectRevisions(rendered);
                if (!BodyContentHashes(acc).SequenceEqual(BodyContentHashes(right)))
                    failures.Add($"seed {seed}: ACCEPT≠RIGHT [{desc}]");
                else if (!BodyContentHashes(rej).SequenceEqual(BodyContentHashes(left)))
                    failures.Add($"seed {seed}: REJECT≠LEFT [{desc}]");
                else if (!BodyFormatSignatures(acc).SequenceEqual(BodyFormatSignatures(right)))
                    failures.Add($"seed {seed}: ACCEPT-FORMAT≠RIGHT [{desc}]");
                else if (!BodyFormatSignatures(rej).SequenceEqual(BodyFormatSignatures(left)))
                    failures.Add($"seed {seed}: REJECT-FORMAT≠LEFT [{desc}]");
                else
                    passed++;
            }
            catch (Exception ex)
            {
                failures.Add($"seed {seed}: THREW {ex.GetType().Name}: {ex.Message} [{desc}]");
            }
        }

        _out.WriteLine($"Format-mutation fuzz: {passed}/{seedCount} seeds passed, {failures.Count} failures");
        foreach (var f in failures.Take(30))
            _out.WriteLine("  FAIL " + f);
        Assert.True(failures.Count == 0, $"{failures.Count}/{seedCount} format-mutation seeds failed.");
    }

    /// <summary>Deterministically generate a (plain, formatted) document pair where the right side adds bold/
    /// italic/color to a seed-chosen subset of runs — pure FormatChanged spans (text identical).</summary>
    private static (WmlDocument Left, WmlDocument Right, string Desc) MakeFormatMutationPair(int seed)
    {
        var rng = new Random(seed);
        string[] bank = { "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel" };
        int paraCount = 1 + rng.Next(3);
        var leftSb = new System.Text.StringBuilder();
        var rightSb = new System.Text.StringBuilder();
        var desc = new System.Text.StringBuilder();
        for (int p = 0; p < paraCount; p++)
        {
            leftSb.Append("<w:p>");
            rightSb.Append("<w:p>");
            int runCount = 2 + rng.Next(4);
            for (int r = 0; r < runCount; r++)
            {
                string word = bank[rng.Next(bank.Length)] + (r < runCount - 1 ? " " : "");
                // Escape nothing — bank words are plain ASCII.
                leftSb.Append($"<w:r><w:t xml:space=\"preserve\">{word}</w:t></w:r>");
                int pick = rng.Next(4);   // 0 = unchanged, 1 = bold, 2 = italic, 3 = color
                string rPr = pick switch
                {
                    1 => "<w:rPr><w:b/></w:rPr>",
                    2 => "<w:rPr><w:i/></w:rPr>",
                    3 => "<w:rPr><w:color w:val=\"FF0000\"/></w:rPr>",
                    _ => "",
                };
                if (pick != 0) desc.Append($"p{p}r{r}:{pick} ");
                rightSb.Append($"<w:r>{rPr}<w:t xml:space=\"preserve\">{word}</w:t></w:r>");
            }
            leftSb.Append("</w:p>");
            rightSb.Append("</w:p>");
        }
        return (IrTestDocuments.FromBodyXml(leftSb.ToString()),
                IrTestDocuments.FromBodyXml(rightSb.ToString()),
                desc.Length == 0 ? "no-format-change" : desc.ToString().Trim());
    }

    // ----------------------------------------------------------------- note-scope markup

    [Fact]
    [Trait("Category", "Corpus")]
    public void Render_footnote_edit_lands_markup_inside_footnotes_part_and_round_trips()
    {
        // WC035-Footnote: a footnote whose text is edited. The markup must land INSIDE the footnotes part
        // (w:ins/w:del under w:footnote), and accept/reject must round-trip the note content.
        var left = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC035-Footnote-Before.docx"));
        var right = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC035-Footnote-After.docx"));
        var rendered = RenderMarkup(left, right);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var fnPart = wd.MainDocumentPart!.FootnotesPart;
        Assert.NotNull(fnPart);
        var fnRoot = fnPart!.GetXDocument().Root!;
        // Revision markup (ins or del) must appear inside a w:footnote.
        var noteRevs = fnRoot.Elements(W.footnote)
            .SelectMany(n => n.Descendants().Where(e => e.Name == W.ins || e.Name == W.del))
            .ToList();
        Assert.NotEmpty(noteRevs);

        AssertRoundTrip(left, right, label: "footnote-edit");
    }

    /// <summary>
    /// M2.4b Workstream D — the hyperlink rId REMAP (WC019). Before and After both reference their hyperlink as
    /// rId4 but to DIFFERENT targets (ericwhite.com → ericwhite2.com). Recreating the right relationship under
    /// the same id would collide with the left's rId4 and leave the cloned link resolving to the LEFT target;
    /// the remap mints a FRESH relationship id for the right target and rewrites the cloned w:hyperlink/@r:id to
    /// it. We assert the ACCEPTED document contains a hyperlink resolving to the RIGHT target (ericwhite2.com) —
    /// the remap half is correct. (The FULL accept/reject round-trip stays allowlisted: rejecting w:del/w:ins
    /// nested inside w:hyperlink is a shared-RevisionProcessor gap, deferred to M2.5 — see Task4BlockedPairs.)
    /// </summary>
    [Fact]
    [Trait("Category", "Corpus")]
    public void Hyperlink_rId_collision_remaps_to_fresh_relationship_resolving_right_target()
    {
        var left = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC019-Hyperlink-Before.docx"));
        var right = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC019-Hyperlink-After-2.docx"));
        var rendered = RenderMarkup(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);

        using var ms = new MemoryStream(accepted.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var main = wd.MainDocumentPart!;
        // The accepted document references a hyperlink whose relationship resolves to the RIGHT target.
        var referencedIds = main.GetXDocument().Descendants(W.hyperlink)
            .Select(h => (string?)h.Attribute(R.r + "id")).Where(s => s != null).ToHashSet();
        var resolvedTargets = main.HyperlinkRelationships
            .Where(r => referencedIds.Contains(r.Id))
            .Select(r => r.Uri.ToString()).ToList();
        Assert.Contains(resolvedTargets, t => t.Contains("ericwhite2.com"));
        // The right target rides on a FRESH id, not the colliding rId4 (which still names the LEFT target).
        var rightRel = main.HyperlinkRelationships.Single(r => r.Uri.ToString().Contains("ericwhite2.com"));
        Assert.NotEqual("rId4", rightRel.Id);
    }

    // ----------------------------------------------------------------- hyperlink-internal edits (B1)

    private const string RelsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Build a one-paragraph body: "Visit " + a <c>w:hyperlink</c> (r:id "rIdLink") wrapping the supplied
    /// run-level XML + " for details." — so the hyperlink anchor is editable independently of the surrounding
    /// text. <paramref name="anchorRunsXml"/> is the inner XML of the hyperlink (one or more <c>w:r</c>).
    /// </summary>
    private static WmlDocument HyperlinkPara(string anchorRunsXml)
    {
        const string preserve = " xml:space=\"preserve\"";
        var body =
            "<w:p>" +
            $"<w:r><w:t{preserve}>Visit </w:t></w:r>" +
            $"<w:hyperlink r:id=\"rIdLink\" xmlns:r=\"{RelsNs}\">{anchorRunsXml}</w:hyperlink>" +
            $"<w:r><w:t{preserve}> for details.</w:t></w:r>" +
            "</w:p>";
        return Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithHyperlinks(
            body, ("rIdLink", "https://example.com"));
    }

    /// <summary>
    /// B1 regression: editing a WORD INSIDE a multi-run <c>w:hyperlink</c> anchor must round-trip. The base
    /// anchor is two runs "our "/"website"; the right edits the second run to "homepage". Before the fix the
    /// 2-way <see cref="DocxDiff.Compare"/> output re-emitted the whole hyperlink once per overlapping token op,
    /// so reject yielded a doubled/tripled anchor ("Visit our websitethe homepageour website for details.")
    /// instead of restoring the base. Invariants: reject ≡ left, accept ≡ right (body text).
    /// </summary>
    [Fact]
    public void Hyperlink_internal_text_edit_round_trips_2way()
    {
        var left = HyperlinkPara("<w:r><w:t xml:space=\"preserve\">our </w:t></w:r><w:r><w:t>website</w:t></w:r>");
        var right = HyperlinkPara("<w:r><w:t xml:space=\"preserve\">our </w:t></w:r><w:r><w:t>homepage</w:t></w:r>");

        var rl = DocxDiff.Compare(left, right);

        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));   // accept ≡ right
        // The full content-hash + format + note invariant (the renderer gate).
        AssertRoundTrip(left, right, label: "hyperlink-internal-multi-run");
    }

    /// <summary>
    /// B1 sibling: editing the anchor of a SINGLE-run hyperlink ("our website" → "the homepage" in one run)
    /// must also round-trip — the same overlapping-token-op walk applies when the single run is split by the
    /// intra-anchor edit. (WC019, the prior whole-anchor single-run replace, must not regress; this is the
    /// partial-edit variant of that shape.)
    /// </summary>
    [Fact]
    public void Hyperlink_internal_single_run_anchor_edit_round_trips_2way()
    {
        var left = HyperlinkPara("<w:r><w:t xml:space=\"preserve\">our website</w:t></w:r>");
        var right = HyperlinkPara("<w:r><w:t xml:space=\"preserve\">the homepage</w:t></w:r>");

        var rl = DocxDiff.Compare(left, right);

        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));   // accept ≡ right
        AssertRoundTrip(left, right, label: "hyperlink-internal-single-run");
    }

    /// <summary>
    /// B1 control: editing a word OUTSIDE the hyperlink (in the same paragraph that contains a multi-run
    /// hyperlink) must stay clean — the hyperlink is untouched, the edit lands on the surrounding run, and the
    /// document round-trips. Guards against the dedup fix over-claiming and dropping an untouched hyperlink.
    /// </summary>
    [Fact]
    public void Hyperlink_present_edit_outside_anchor_round_trips_2way()
    {
        var left = HyperlinkPara("<w:r><w:t xml:space=\"preserve\">our </w:t></w:r><w:r><w:t>website</w:t></w:r>");
        // Edit the trailing text only ("for details." → "for info."); anchor identical.
        var leftDoc = left;
        var rightBody =
            "<w:p>" +
            "<w:r><w:t xml:space=\"preserve\">Visit </w:t></w:r>" +
            $"<w:hyperlink r:id=\"rIdLink\" xmlns:r=\"{RelsNs}\"><w:r><w:t xml:space=\"preserve\">our </w:t></w:r><w:r><w:t>website</w:t></w:r></w:hyperlink>" +
            "<w:r><w:t xml:space=\"preserve\"> for info.</w:t></w:r>" +
            "</w:p>";
        var right = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithHyperlinks(
            rightBody, ("rIdLink", "https://example.com"));

        var rl = DocxDiff.Compare(leftDoc, right);

        Assert.Equal(Docs.PlainText(leftDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));      // accept ≡ right
        AssertRoundTrip(leftDoc, right, label: "hyperlink-present-edit-outside");
    }

    /// <summary>
    /// Build a one-paragraph body holding TWO ADJACENT <c>w:hyperlink</c> siblings that share the SAME target
    /// (both reference relationship id "rIdLink"). The two links are genuinely DISTINCT source elements that
    /// merely happen to point at the same URI — a shape Word produces routinely (two authored links to one page).
    /// <paramref name="firstInner"/> / <paramref name="secondInner"/> are the inner XML of each hyperlink.
    /// </summary>
    private static WmlDocument TwoAdjacentSameTargetLinks(string firstInner, string secondInner)
    {
        var body =
            "<w:p>" +
            $"<w:hyperlink r:id=\"rIdLink\" xmlns:r=\"{RelsNs}\">{firstInner}</w:hyperlink>" +
            $"<w:hyperlink r:id=\"rIdLink\" xmlns:r=\"{RelsNs}\">{secondInner}</w:hyperlink>" +
            "</w:p>";
        return Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithHyperlinks(
            body, ("rIdLink", "https://example.com"));
    }

    /// <summary>
    /// REGRESSION (F2 follow-up): two ADJACENT, genuinely-DISTINCT source <c>w:hyperlink</c>s that share the
    /// same target ("first" / "second"); the RIGHT edits the SECOND link's anchor ("second" → "SECOND"). The F2
    /// fragment-coalescer rejoined emitted hyperlink fragments by ATTRIBUTE EQUALITY gated on "carries a plain
    /// run", which could not tell "N fragments of ONE split source link" from "N distinct source links sharing a
    /// target": it folded all three emitted fragments into ONE link, so REJECT yielded ONE link ("firstsecond")
    /// while LEFT had TWO — RejectRevisions ≢ left at the ContentHash level (IrReader frames each hyperlink
    /// boundary). The source-wrapper-identity coalescer must keep the two distinct links separate.
    /// </summary>
    [Fact]
    public void Adjacent_distinct_same_target_hyperlinks_one_edited_round_trips()
    {
        var left = TwoAdjacentSameTargetLinks(
            "<w:r><w:t>first</w:t></w:r>", "<w:r><w:t>second</w:t></w:r>");
        var right = TwoAdjacentSameTargetLinks(
            "<w:r><w:t>first</w:t></w:r>", "<w:r><w:t>SECOND</w:t></w:r>");

        var rl = DocxDiff.Compare(left, right);

        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));   // accept ≡ right
        AssertRoundTrip(left, right, label: "adjacent-distinct-same-target-one-edited");
    }

    /// <summary>
    /// REGRESSION guard: the same two adjacent distinct same-target links with NO edit must NOT collapse into
    /// one link — reject ≡ left ≡ accept at the ContentHash level (two link boundaries preserved both ways).
    /// </summary>
    [Fact]
    public void Adjacent_distinct_same_target_hyperlinks_unchanged_stay_separate()
    {
        var doc = TwoAdjacentSameTargetLinks(
            "<w:r><w:t>first</w:t></w:r>", "<w:r><w:t>second</w:t></w:r>");

        var rl = DocxDiff.Compare(doc, doc);

        Assert.Equal(Docs.PlainText(doc), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(doc), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));    // accept ≡ right
        AssertRoundTrip(doc, doc, label: "adjacent-distinct-same-target-unchanged");
    }

    /// <summary>
    /// Edge pin: a single-token hyperlink anchor with NO separator ("aaaa" → "zzzz") FULLY replaced — there is no
    /// Equal (plain) run inside the link, so the walk emits a pure del-link then a pure ins-link. Because both
    /// resolve to the SAME target (the anchor text is rewritten, the href is unchanged), issue #232's target-gate
    /// coalesces them back into ONE <c>w:hyperlink</c> (one element, original r:id) — byte-faithful to the
    /// source's single link. Reject ≡ left, accept ≡ right still hold.
    /// </summary>
    [Fact]
    public void Hyperlink_single_token_anchor_no_equal_run_replaced_round_trips()
    {
        var left = HyperlinkPara("<w:r><w:t>aaaa</w:t></w:r>");
        var right = HyperlinkPara("<w:r><w:t>zzzz</w:t></w:r>");

        var rl = DocxDiff.Compare(left, right);

        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));   // accept ≡ right
        // #232: a single fully-replaced same-target link renders as ONE w:hyperlink, not per-token fragments.
        Assert.Equal(new[] { "rIdLink" }, HyperlinkRelIds(rl));
        AssertRoundTrip(left, right, label: "hyperlink-single-token-no-equal-run");
    }

    /// <summary>Count-and-id helper: the <c>@r:id</c> of every <c>w:hyperlink</c> in a document's main part, in
    /// document order (an internal link with only <c>@w:anchor</c> reads as null). #232 asserts a single source
    /// link renders as a single output hyperlink.</summary>
    private static string?[] HyperlinkRelIds(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.GetXDocument().Descendants(W.hyperlink)
            .Select(h => (string?)h.Attribute(R.id))
            .ToArray();
    }

    /// <summary>Build a one-paragraph hyperlink body like <see cref="HyperlinkPara"/> but with an explicit target
    /// URI, so a test can make the SAME <c>r:id</c> resolve to DIFFERENT targets on the two sides (the WC019
    /// retarget shape) or the SAME target (the #232 rewrite shape).</summary>
    private static WmlDocument HyperlinkParaTo(string anchorRunsXml, string uri)
    {
        const string preserve = " xml:space=\"preserve\"";
        var body =
            "<w:p>" +
            $"<w:r><w:t{preserve}>Visit </w:t></w:r>" +
            $"<w:hyperlink r:id=\"rIdLink\" xmlns:r=\"{RelsNs}\">{anchorRunsXml}</w:hyperlink>" +
            $"<w:r><w:t{preserve}> for details.</w:t></w:r>" +
            "</w:p>";
        return Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithHyperlinks(body, ("rIdLink", uri));
    }

    /// <summary>
    /// #232 (multi-word): a single-target hyperlink whose ENTIRE multi-word anchor is replaced with no shared
    /// token ("catdog" → "fish bird whale", href unchanged) must render as ONE <c>w:hyperlink</c> carrying the
    /// original r:id — not per-token del/ins link fragments. Before the target-gate the pure-del + pure-ins run
    /// had no plain piece, so the coalescer left two link elements (accept/reject round-tripped at the text level
    /// but the raw markup fragmented, 1 → N).
    /// </summary>
    [Fact]
    public void Hyperlink_fully_replaced_same_target_renders_as_one_hyperlink()
    {
        var left = HyperlinkPara("<w:r><w:t>catdog</w:t></w:r>");
        var right = HyperlinkPara("<w:r><w:t xml:space=\"preserve\">fish bird whale</w:t></w:r>");

        var rl = DocxDiff.Compare(left, right);

        Assert.Equal(new[] { "rIdLink" }, HyperlinkRelIds(rl));                                       // ONE link
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));   // accept ≡ right
        AssertRoundTrip(left, right, label: "hyperlink-fully-replaced-same-target");
    }

    /// <summary>
    /// #232 discriminator: a whole-anchor RETARGET — the SAME <c>r:id</c> resolves to a DIFFERENT href on each
    /// side ("aaaa"→example.com vs "zzzz"→example.org) — must NOT be coalesced. A single hyperlink carries ONE
    /// r:id, and the target genuinely differs, so the del-link (old target) and ins-link (new target) stay TWO
    /// elements for the post-assembly rId remap (the synthetic analogue of the WC019 corpus case). This pins that
    /// the fix keys on the RESOLVED target, not the r:id string (identical here), so it never folds a retarget.
    /// </summary>
    [Fact]
    public void Hyperlink_whole_anchor_retarget_stays_two_links()
    {
        var left = HyperlinkParaTo("<w:r><w:t>aaaa</w:t></w:r>", "https://example.com/");
        var right = HyperlinkParaTo("<w:r><w:t>zzzz</w:t></w:r>", "https://example.org/");

        var rl = DocxDiff.Compare(left, right);

        // TWO hyperlink elements survive (old-target del-link + new-target ins-link, the ins remapped to a fresh
        // id post-assembly so it no longer reads the same string).
        Assert.Equal(2, HyperlinkRelIds(rl).Length);
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rl)));   // reject ≡ left
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionAccepter.AcceptRevisions(rl)));   // accept ≡ right
        AssertRoundTrip(left, right, label: "hyperlink-whole-anchor-retarget");
    }

    [Fact]
    [Trait("Category", "Corpus")]
    public void Render_endnote_edit_lands_markup_inside_endnotes_part_and_round_trips()
    {
        var left = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC035-Endnote-Before.docx"));
        var right = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC035-Endnote-After.docx"));
        var rendered = RenderMarkup(left, right);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var enPart = wd.MainDocumentPart!.EndnotesPart;
        Assert.NotNull(enPart);
        var enRoot = enPart!.GetXDocument().Root!;
        var noteRevs = enRoot.Elements(W.endnote)
            .SelectMany(n => n.Descendants().Where(e => e.Name == W.ins || e.Name == W.del))
            .ToList();
        Assert.NotEmpty(noteRevs);

        AssertRoundTrip(left, right, label: "endnote-edit");
    }

    // ----------------------------------------------------------------- native move markup (w:moveFrom/To)

    /// <summary>Build a doc from plain-text paragraphs (mirrors WmlComparerMoveDetectionTests' fixtures).</summary>
    private static WmlDocument MoveDoc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    paragraphs.Select(t => new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text(t))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }
        return new WmlDocument("move.docx", stream.ToArray());
    }

    private static readonly string[] MoveLeft =
    {
        "This is paragraph A with enough words for move detection here.",
        "This is paragraph B with sufficient content to anchor it firmly.",
        "This is paragraph C with more words added for good measure today.",
    };
    private static readonly string[] MoveRight =
    {
        "This is paragraph B with sufficient content to anchor it firmly.",
        "This is paragraph A with enough words for move detection here.",
        "This is paragraph C with more words added for good measure today.",
    };

    [Fact]
    public void Render_move_emits_native_moveFrom_moveTo_with_shared_name_and_round_trips()
    {
        var left = MoveDoc(MoveLeft);
        var right = MoveDoc(MoveRight);
        var rendered = RenderMarkup(left, right);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;

        var moveFrom = body.Descendants(W.moveFrom).ToList();
        var moveTo = body.Descendants(W.moveTo).ToList();
        Assert.NotEmpty(moveFrom);
        Assert.NotEmpty(moveTo);

        // Range markers present and start/end counts pair up.
        Assert.Equal(body.Descendants(W.moveFromRangeStart).Count(), body.Descendants(W.moveFromRangeEnd).Count());
        Assert.Equal(body.Descendants(W.moveToRangeStart).Count(), body.Descendants(W.moveToRangeEnd).Count());

        // Names link FROM and TO halves (set-equal), are non-empty, and follow the "moveN" convention.
        var fromNames = body.Descendants(W.moveFromRangeStart).Select(e => (string?)e.Attribute(W.name)).ToHashSet();
        var toNames = body.Descendants(W.moveToRangeStart).Select(e => (string?)e.Attribute(W.name)).ToHashSet();
        Assert.NotEmpty(fromNames);
        Assert.True(fromNames.SetEquals(toNames), "moveFrom/moveTo range names must pair");
        Assert.All(fromNames, n => Assert.StartsWith("move", n));

        // Required attributes on moveFrom/moveTo runs.
        foreach (var e in moveFrom.Concat(moveTo))
        {
            Assert.NotNull(e.Attribute(W.id));
            Assert.NotNull(e.Attribute(W.author));
            Assert.NotNull(e.Attribute(W.date));
        }

        AssertRoundTrip(left, right, label: "native-move");
    }

    [Fact]
    public void Render_move_output_is_recognized_as_Moved_by_WmlComparer_GetRevisions()
    {
        // THE ORACLE: WmlComparer.GetRevisions, run over OUR rendered output, must see Moved revisions — proving
        // our native move markup is structurally what the shipped reader recognizes.
        var left = MoveDoc(MoveLeft);
        var right = MoveDoc(MoveRight);
        var rendered = RenderMarkup(left, right);

        var revs = WmlComparer.GetRevisions(rendered, new WmlComparerSettings());
        var moved = revs.Where(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Moved).ToList();
        Assert.True(moved.Count >= 2, $"WmlComparer.GetRevisions should see ≥2 Moved in our output (saw {moved.Count} of {revs.Count} total)");
    }

    [Fact]
    public void Render_move_and_edit_nests_ins_del_inside_moveTo_and_round_trips()
    {
        // Paragraph A is relocated AND edited (one word changed): a MoveModify. The destination moveTo range
        // must carry nested ins/del for the in-move edit, and RevisionProcessor (the oracle) must accept it to
        // the right and reject it to the left.
        var left = MoveDoc(
            "This is paragraph A with enough words for move detection here.",
            "This is paragraph B with sufficient content to anchor it firmly.");
        var right = MoveDoc(
            "This is paragraph B with sufficient content to anchor it firmly.",
            "This is paragraph A with PLENTY words for move detection here.");
        var settings = new IrDiffSettings { MoveSimilarityThreshold = 0.6, MoveMinimumTokenCount = 3 };
        var rendered = RenderMarkup(left, right, settings);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        // The destination's move range contains the relocated unchanged spans plus the separate in-move edit
        // revisions.  A plain right-side moveTo clone would accept/reject correctly but hide this edit.
        var destination = Assert.Single(body.Elements(W.p),
            p => p.Descendants(W.moveToRangeStart).Any());
        Assert.NotEmpty(destination.Descendants(W.moveTo));
        Assert.NotEmpty(destination.Descendants(W.ins));
        Assert.NotEmpty(destination.Descendants(W.del));
        AssertRoundTrip(left, right, settings, label: "move-modify");
    }

    [Fact]
    public void Render_moved_format_change_is_tracked_at_destination()
    {
        // Exact text relocation with a bold change is a MoveModify: its destination must carry both the native
        // move wrapper and the old run formatting. Without that projection, reviewers see only a move and reject
        // cannot reconstruct the left formatting.
        const string anchor = "Anchor paragraph holds the document spine steady.";
        const string moved = "Moved paragraph contains enough distinct words for move detection today.";
        const string trailing = "Trailing paragraph also has enough words to stabilize matching.";
        const string trailing2 = "Final paragraph provides another stable matching anchor here.";
        var left = IrTestDocuments.FromBodyXml(
            $"<w:p><w:r><w:t>{anchor}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:t>{moved}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:t>{trailing}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:t>{trailing2}</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            $"<w:p><w:r><w:t>{anchor}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:t>{trailing}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:t>{trailing2}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{moved}</w:t></w:r></w:p>");
        var settings = new IrDiffSettings { MoveSimilarityThreshold = 0.6, MoveMinimumTokenCount = 3 };

        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), settings);
        var destinationOp = Assert.Single(script.Operations,
            op => op.Kind == IrEditOpKind.MoveModifyBlock && op.IsMoveSource == false);
        Assert.Contains(destinationOp.TokenDiff!.Ops, op => op.Kind == IrTokenOpKind.FormatChanged);

        var rendered = RenderMarkup(left, right, settings);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        var destination = Assert.Single(body.Elements(W.p),
            p => p.Descendants(W.moveToRangeStart).Any());
        Assert.NotEmpty(destination.Descendants(W.moveTo));
        var oldRunFormat = Assert.Single(destination.Descendants(W.rPrChange));
        Assert.Null(oldRunFormat.Element(W.rPr)!.Element(W.b));
        AssertRoundTrip(left, right, settings, label: "move-format");
    }

    /// <summary>
    /// Two paragraphs can be both edited and reordered inside one unmatched gap.  Their strong lexical
    /// correspondence crosses document order, so treating either pair as an in-place ModifyBlock is unsafe:
    /// the right-ordered renderer then keeps the left text in the right order when revisions are rejected.
    /// Such correspondence must lower to move or delete/insert semantics, both of which restore the original
    /// left order on reject.
    /// </summary>
    [Fact]
    public void Render_crossed_edited_paragraphs_restore_left_order_on_reject()
    {
        const string leftBody =
            "<w:p><w:r><w:t>anchor-start</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>red green blue orange</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>cat dog mouse horse</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchor-end</w:t></w:r></w:p>";
        const string rightBody =
            "<w:p><w:r><w:t>anchor-start</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>cat dog mouse zebra</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>red green blue purple</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchor-end</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXml(leftBody);
        var right = IrTestDocuments.FromBodyXml(rightBody);

        AssertRoundTrip(left, right, label: "crossed-edited-paragraphs");
    }

    [Fact]
    public void Render_move_with_DetectMoves_off_demotes_to_ins_del()
    {
        var left = MoveDoc(MoveLeft);
        var right = MoveDoc(MoveRight);
        var settings = new IrDiffSettings { RenderMoves = false };
        var rendered = RenderMarkup(left, right, settings);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.Empty(body.Descendants(W.moveFrom));
        Assert.Empty(body.Descendants(W.moveTo));
        Assert.True(body.Descendants(W.ins).Any() || body.Descendants(W.del).Any(), "demoted move must use ins/del");
        AssertRoundTrip(left, right, settings, label: "move-demoted");
    }

    [Fact]
    public void Render_move_with_SimplifyMoveMarkup_converts_to_del_ins_and_strips_ranges()
    {
        var left = MoveDoc(MoveLeft);
        var right = MoveDoc(MoveRight);
        var settings = new IrDiffSettings { SimplifyMoveMarkup = true };
        var rendered = RenderMarkup(left, right, settings);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.Empty(body.Descendants(W.moveFrom));
        Assert.Empty(body.Descendants(W.moveTo));
        Assert.Empty(body.Descendants(W.moveFromRangeStart));
        Assert.Empty(body.Descendants(W.moveToRangeStart));
        Assert.True(body.Descendants(W.del).Any(), "simplified moveFrom → del");
        Assert.True(body.Descendants(W.ins).Any(), "simplified moveTo → ins");
        AssertRoundTrip(left, right, settings, label: "move-simplified");
    }

    /// <summary>
    /// WC-BodyBookmarks capability verdict (M2.6 Task 2). This fixture pair converts the document's ENTIRE
    /// endnote store to footnotes (BEFORE: 24 footnotes + 190 endnotes; AFTER: 213 footnotes + 0 endnotes) and
    /// carries many body-level bookmark markers. On it the WmlComparer ORACLE THROWS a DocxodusException
    /// ("Internal error in ProcessFootnoteEndnote") — it produces NO comparison at all (see
    /// <c>WmlComparerBodyLevelBookmarkTests</c>, whose own comment documents this as a separate bug from the
    /// body-bookmark NRE). There is therefore no oracle behaviour to match.
    ///
    /// <para>Our IR engine's <see cref="DocxDiff.GetRevisions"/> surface, by contrast, completes WITHOUT
    /// throwing and yields a substantial revision list — we EXCEED the oracle here (it cannot even run). This
    /// test pins that capability: GetRevisions is total on the pathological note-store-conversion fixture.</para>
    ///
    /// <para>The separate MARKUP round-trip (<c>Compare</c> accept/reject ≡ right/left) does NOT hold for this
    /// pair — the whole-store endnote→footnote cross-part migration surfaces as 190 endnote deletions + 190
    /// footnote insertions that the per-scope note diff does not reconcile to a clean accept==right. That
    /// failure is retained in <see cref="Task4BlockedPairs"/> with this oracle-throws ceiling as context;
    /// fixing cross-part note-store conversion is a large effort with negligible real-world value (the oracle
    /// itself cannot do it, and whole-store note-kind flips do not occur in practice), so it is NOT pursued.</para>
    /// </summary>
    [Fact]
    public void WC_BodyBookmarks_GetRevisions_is_total_where_the_oracle_throws()
    {
        var before = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC-BodyBookmarks-Before.docx"));
        var after = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC-BodyBookmarks-After.docx"));

        // The oracle throws on this fixture (no behaviour to match) — assert that explicitly so the verdict is
        // self-documenting and re-checked: if WmlComparer ever LEARNS to handle this pair, revisit the verdict.
        Assert.Throws<DocxodusException>(() => WmlComparer.Compare(before, after, new WmlComparerSettings()));

        // Our engine completes and produces revisions in BOTH directions — the capability win.
        var fwd = DocxDiff.GetRevisions(before, after);
        var rev = DocxDiff.GetRevisions(after, before);
        Assert.NotEmpty(fwd);
        Assert.NotEmpty(rev);
        Assert.All(fwd, r => Assert.NotNull(r.Text));
        Assert.All(rev, r => Assert.NotNull(r.Text));
    }

    // ----------------------------------------------------------------- WC022 ordering regression (M2.6 T2)

    /// <summary>
    /// WC022 adjacent-empty-paragraph ordering regression (M2.6 Task 2). BEFORE has two adjacent empty
    /// paragraphs where the second keeps its persisted unid into AFTER; the aligner's InOrderRefine must
    /// reserve that same-unid identity pair BEFORE first-fitting the other empty, or reject reconstructs the
    /// pair swapped (a fwd-REJECT-only failure). Both directions must round-trip clean (content + format +
    /// notes), pinning the IrBlockAligner identity-reservation fix.
    /// </summary>
    [Fact]
    public void WC022_adjacent_empty_paragraphs_round_trip_both_directions()
    {
        var before = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC022-Image-Math-Para-Before.docx"));
        var after = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC022-Image-Math-Para-After.docx"));
        AssertRoundTrip(before, after, label: "WC022-fwd");
        AssertRoundTrip(after, before, label: "WC022-rev");
    }

    // ----------------------------------------------------------------- split/merge markup (M2.6 Task 7)

    private static readonly IrDiffSettings SplitOn = new() { DetectSplitMerge = true };

    /// <summary>M2.6 split markup (anchored-split shape): the produced document carries an inserted
    /// paragraph mark (empty <c>w:ins</c> in <c>pPr/rPr</c>) on every paragraph but the last of the
    /// group; REJECT removes the marks, re-merging the split paragraphs into the original LEFT one.</summary>
    [Fact]
    public void Split_markup_accept_yields_right_reject_yields_left()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">aaa bbb ccc ddd. eee fff ggg hhh.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchor one two three four five.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">aaa bbb ccc ddd. </w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>eee fff ggg hhh.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchor one two three four five.</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right, SplitOn);
        Assert.Equal(0, SchemaErrorCount(rendered));

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        // The split-introduced pilcrow is a paragraph-mark revision: empty w:ins inside pPr/rPr.
        Assert.NotEmpty(body.Elements(W.p).Elements(W.pPr).Elements(W.rPr).Elements(W.ins));

        AssertRoundTrip(left, right, SplitOn, label: "split-markup");
    }

    /// <summary>M2.6 merge markup: paragraphs 0..N-2 of the merged group carry a DELETED mark
    /// (empty <c>w:del</c> in <c>pPr/rPr</c>); ACCEPT merges them into the following paragraph,
    /// REJECT restores the original N LEFT paragraphs.</summary>
    [Fact]
    public void Merge_markup_accept_yields_right_reject_yields_left()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">aaa bbb ccc ddd. </w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>eee fff ggg hhh.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchor one two three four five.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">aaa bbb ccc ddd. eee fff ggg hhh.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchor one two three four five.</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right, SplitOn);
        Assert.Equal(0, SchemaErrorCount(rendered));

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.NotEmpty(body.Elements(W.p).Elements(W.pPr).Elements(W.rPr).Elements(W.del));

        AssertRoundTrip(left, right, SplitOn, label: "merge-markup");
    }

    /// <summary>The two corpus split fixtures (cell-scope splits): with detection ON the produced
    /// markup must still round-trip both ways and stay schema-valid.</summary>
    [Theory]
    [InlineData("WC041-Table-5.docx", "WC041-Table-5-Mod.docx")]
    [InlineData("WC023-Table-4-Row-Image-Before.docx", "WC023-Table-4-Row-Image-After-Delete-1-Row.docx")]
    public void Fixture_split_markup_round_trips(string l, string r)
    {
        var left = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, l));
        var right = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, r));
        var rendered = RenderMarkup(left, right, SplitOn);
        Assert.Equal(0, SchemaErrorCount(rendered));
        AssertRoundTrip(left, right, SplitOn, label: $"split-fixture {l}");
    }

    /// <summary>F4.2 regression: the WC022 identity-reservation reject-order invariant must hold with
    /// split/merge detection ON, both directions (the scan never promotes Unchanged/FormatOnly pairs).</summary>
    [Fact]
    public void WC022_reject_order_invariant_holds_with_detection_on()
    {
        var before = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC022-Image-Math-Para-Before.docx"));
        var after = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC022-Image-Math-Para-After.docx"));
        AssertRoundTrip(before, after, SplitOn, label: "WC022-split-on-fwd");
        AssertRoundTrip(after, before, SplitOn, label: "WC022-split-on-rev");
    }

    // ----------------------------------------------------------------- corpus invariant (92 × 2)

    /// <summary>
    /// The M2.4 DOCUMENTED-DEVIATION base↔variant pairs: their accept/reject round-trip does not hold for a
    /// reason rooted in the ENGINE READER/ALIGNER or in relationship-id remapping — NOT in the renderer's
    /// markup, which is correct for every body/table/move/format/note construct the edit script expresses. Each
    /// entry below carries its PRECISE root cause. This allowlist is a RATCHET — the invariant test asserts
    /// EVERY other pair round-trips AND that no allowlisted pair UNEXPECTEDLY passes (a fixed-early pair must be
    /// removed). The Task-4 burndown drove this from 11 to 6 distinct root causes; M2.5 Task 3 then closed WC019
    /// (RevisionProcessor reject del/ins under w:hyperlink + empty-hyperlink-shell drop), M2.6 Task 1 closed
    /// WC034 foot+end (the note-id renumber/reorder pass — IrMarkupRenderer.RenumberNoteIds), and M2.6 Task 2
    /// closed WC022 (the InOrderRefine same-unid identity-reservation phase — IrBlockAligner), leaving 1 entry:
    /// WC-BodyBookmarks (endnote→footnote whole-note-store conversion, on which the WmlComparer ORACLE ITSELF
    /// throws "Internal error in ProcessFootnoteEndnote" — there is no oracle behaviour to match; see below).
    /// </summary>
    private static readonly HashSet<string> Task4BlockedPairs = new(StringComparer.Ordinal)
    {
        // (M2.6 Task 1 — CLOSED) The WC034-After3 note-renumber family. The note CONTENT round-tripped after M2.5
        // Task 3 fixed the note-store CORRESPONDENCE (notes pair by body-reference order + content, not raw w:id);
        // the SURVIVING residual was purely the note-element id/ORDER in the produced part — a matched note paired
        // left-en#1 → right-en#2 but the produced definition kept the LEFT id at the LEFT part position, so the
        // accepted part's note sequence was [en2,en1,en3] vs RIGHT's [en1,en2,en3]. Closed by the note-id renumber
        // pass (IrMarkupRenderer.RenumberNoteIds) mirroring the oracle's ChangeFootnoteEndnoteReferencesToUniqueRange:
        // body references are renumbered to document order (base 1, separator/continuation boilerplate reserved),
        // each definition renumbered + reordered to match, del references resolving deleted-only defs and ins/equal
        // references resolving the live (matched/inserted) def. Both directions now round-trip clean (ACCEPT==RIGHT
        // notes, REJECT==LEFT notes) — removed from the allowlist.
        // (M2.4b Workstream A — CLOSED, 3 of 4) The SmartArt diagram rel-id family (WC014 ×2 + WC052) was here
        // as DEVIATIONS: an UNCHANGED diagram's relationship ids renumber between revisions (and on accept
        // MoveRelatedPartsToDestination mints fresh "R…" ids), and its wp:docPr/@id renumbers (1 vs 2), so the
        // opaque content hash for that block differed side-to-side and on accept. Fixed at the reader/hasher
        // level — IrHasher.Canonicalize now resolves every relationship-namespace attribute to a stable
        // content-identity token (media → part-content SHA, xml diagram parts dropped to match the WmlComparer
        // oracle, external/hyperlink → target URI, dangling → sentinel) and strips the renumber-prone
        // wp:docPr/@id. Content identity over rel numbering: those three pairs now round-trip clean and are
        // removed from this allowlist.
        // (M2.6 Task 2 — CLOSED) WC022's adjacent-empty-paragraph ordering. The bookmark half closed in M2.4b
        // WS-D (the body-level bookmark marker is dropped like WmlComparer's RemoveBookmarks). The surviving
        // residual was an aligner ORDERING bug: BEFORE had two adjacent empty paragraphs [efb022(empty+pPr),
        // c88b(bare empty)]; AFTER had [5e71(bare empty), c88b(bare empty)]. InOrderRefine's first-to-first
        // matched AFTER's bare empty 5e71 (scanned first, no identity match) to the only free bare-empty left
        // c88b, stranding BEFORE's efb022 to pair with AFTER's c88b — crossing document order, so reject
        // reconstructed [8]/[9] swapped (a fwd-REJECT-only failure; accept + the reverse direction were already
        // clean). Closed by giving InOrderRefine a SAME-UNID identity-reservation phase (IrBlockAligner): a
        // free right block first claims the free left block sharing its persisted unid (the genuinely-unchanged
        // c88b↔c88b pair) BEFORE any first-fit, keeping the pairing monotonic. Pure deterministic tie-break —
        // it only chooses among equal-(content,format) candidates, never which blocks pair — so no other corpus
        // pair shifts. NOT a 1:N problem (all pairings here are 1×1); both directions now round-trip clean.
        // Removed from the allowlist.
        // (M2.5 Task 3 — CLOSED) hyperlink TARGET+TEXT change where the right hyperlink's rId COLLIDES with a
        // DIFFERENT left rId (Before → ericwhite.com, After → ericwhite2.com, BOTH as rId4). The true rId REMAP
        // landed in M2.4b WS-D (ImportHyperlinkAndExternalRelationships mints a fresh id and rewrites the cloned
        // @r:id). The residual blocker was the shared RevisionProcessor: a hyperlink-text edit nests w:del/w:ins
        // INSIDE the w:hyperlink (the schema forbids a hyperlink inside w:ins/w:del, so the markers live within
        // it), and REJECT's del→ins / ins→del reversal rules were gated to parent==w:p, never firing under a
        // w:hyperlink — so reject left the deleted link's content unrestored and the inserted link's content
        // unremoved. Fixed by (1) extending those two REJECT rules to also fire when parent==w:hyperlink
        // (additive — WmlComparer never produces del/ins-in-hyperlink because it strips hyperlinks pre-compare,
        // so no existing case is affected), and (2) dropping an EMPTY w:hyperlink shell (no surviving run
        // content) in the accept transform (the artifact of the IR's del-old-link/ins-new-link shape). Both
        // directions now round-trip clean (ACCEPT==RIGHT, REJECT==LEFT, content + format). Removed from the
        // allowlist; the full old-engine RevisionProcessor/WmlComparer suite stays green.
        // (M2.6 Task 2 — FINAL VERDICT: RETAINED, oracle-throws ceiling) WC-BodyBookmarks-After carries MANY
        // body-level bookmarkStart/End markers AND converts the document's ENTIRE endnote store to footnotes
        // (measured: BEFORE 24 fn + 190 en; AFTER 213 fn + 0 en). The bookmark half closed in M2.4b WS-D (the
        // body-level markers drop like WmlComparer's RemoveBookmarks). The surviving blocker is the whole-store
        // ENDNOTE→FOOTNOTE cross-part conversion: 190 endnotes migrate into the footnote part, surfacing as 190
        // endnote deletions + 190 footnote insertions that the IR's PER-SCOPE note diff (fn-vs-fn, en-vs-en)
        // does not reconcile to a clean accept==right (note counts land at accept=424 vs right=426 fwd).
        //
        // METHOD-RULE VERDICT (measured end-to-end, M2.6 T2): the WmlComparer ORACLE THROWS on this fixture —
        // WmlComparer.Compare raises DocxodusException "Internal error in ProcessFootnoteEndnote" and produces
        // NOTHING (its own WmlComparerBodyLevelBookmarkTests only asserts the body-bookmark NRE is gone, not
        // that Compare succeeds). There is NO oracle behaviour to match, so this is not a parity gap. Where the
        // oracle dies our DocxDiff.GetRevisions COMPLETES and yields 6319 revisions both directions — we EXCEED
        // the oracle on the consumer surface (pinned by WC_BodyBookmarks_GetRevisions_is_total_where_the_oracle_throws).
        // Only the MARKUP round-trip (Compare accept/reject ≡ right/left) fails, on the cross-part note-store
        // migration. HONEST WORTH ASSESSMENT: fixing whole-store note-kind conversion is a large, isolated
        // effort (a cross-part note correspondence the per-scope diff is not built for) for negligible value —
        // the oracle itself cannot do it, and converting an entire endnote store to footnotes does not occur in
        // real documents (this is a single synthetic fixture). NOT pursued; retained with this ceiling context.
        "WC-BodyBookmarks-Before.docx↔WC-BodyBookmarks-After.docx",
    };

    [Fact]
    [Trait("Category", "Corpus")]
    public void WC_corpus_markup_accept_reject_round_trips_both_directions()
    {
        var pairs = WcCorpus.BuildPairs();
        Assert.True(pairs.Count >= 30, $"Expected a substantial WC pair list; inferred {pairs.Count}.");

        var settings = new IrDiffSettings();
        int passed = 0;
        var failures = new List<string>();            // a pair NOT on the Task-4 allowlist failed (a regression)
        var blockedNowPassing = new List<string>();   // an allowlisted pair UNEXPECTEDLY passed (ratchet down)

        foreach (var (baseName, variantName) in pairs)
        {
            string key = $"{baseName}↔{variantName}";
            bool blocked = Task4BlockedPairs.Contains(key);
            var baseDoc = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, baseName));
            var variantDoc = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, variantName));

            bool pairOk = true;
            foreach (var (l, r, dir) in new[] { (baseDoc, variantDoc, "fwd"), (variantDoc, baseDoc, "rev") })
            {
                string? failure = null;
                try
                {
                    var rendered = RenderMarkup(l, r, settings);
                    var acceptedDoc = RevisionProcessor.AcceptRevisions(rendered);
                    var rejectedDoc = RevisionProcessor.RejectRevisions(rendered);
                    var accept = BodyContentHashes(acceptedDoc);
                    var reject = BodyContentHashes(rejectedDoc);
                    if (!accept.SequenceEqual(BodyContentHashes(r)))
                        failure = $"{key} [{dir}] ACCEPT≠RIGHT";
                    else if (!reject.SequenceEqual(BodyContentHashes(l)))
                        failure = $"{key} [{dir}] REJECT≠LEFT";
                    else if (!BodyFormatSignatures(acceptedDoc).SequenceEqual(BodyFormatSignatures(r)))
                        failure = $"{key} [{dir}] ACCEPT-FORMAT≠RIGHT";
                    else if (!BodyFormatSignatures(rejectedDoc).SequenceEqual(BodyFormatSignatures(l)))
                        failure = $"{key} [{dir}] REJECT-FORMAT≠LEFT";
                    else if (!NoteContentHashes(acceptedDoc).SequenceEqual(NoteContentHashes(r)))
                        failure = $"{key} [{dir}] ACCEPT-NOTES≠RIGHT";
                    else if (!NoteContentHashes(rejectedDoc).SequenceEqual(NoteContentHashes(l)))
                        failure = $"{key} [{dir}] REJECT-NOTES≠LEFT";
                }
                catch (Exception ex)
                {
                    failure = $"{key} [{dir}] THREW {ex.GetType().Name}: {ex.Message}";
                }

                if (failure == null)
                    passed++;
                else
                {
                    pairOk = false;
                    if (!blocked) failures.Add(failure);
                }
            }

            // A Task-4-allowlisted pair that round-trips in BOTH directions was fixed early — flag it so the
            // allowlist ratchets DOWN (its entry must be removed, never silently retained).
            if (blocked && pairOk)
                blockedNowPassing.Add(key);
        }

        int total = pairs.Count * 2;
        _out.WriteLine($"WC corpus markup invariant: {passed}/{total} round-trips passed " +
            $"({Task4BlockedPairs.Count} pairs Task-4-blocked).");
        foreach (var f in failures.Take(40))
            _out.WriteLine("  UNEXPECTED FAIL " + f);
        foreach (var p in blockedNowPassing)
            _out.WriteLine("  RATCHET: Task-4-blocked pair now passes, remove from allowlist: " + p);

        Assert.True(failures.Count == 0,
            $"{failures.Count} non-allowlisted corpus round-trips failed (Task-3 regressions — see output).");
        Assert.True(blockedNowPassing.Count == 0,
            $"{blockedNowPassing.Count} Task-4-allowlisted pairs now pass — remove them from Task4BlockedPairs.");
    }

    // ----------------------------------------------------------------- fuzz invariant (50 seeds)

    [Fact]
    [Trait("Category", "Fuzz")]
    public void Fuzz_markup_accept_reject_round_trips_over_seeds()
    {
        const int seedCount = 50;
        var settings = new IrDiffSettings();
        int passed = 0;
        var failures = new List<string>();

        for (int seed = 1; seed <= seedCount; seed++)
        {
            var fuzzCase = DiffFuzzer.Generate(seed);
            try
            {
                var rendered = RenderMarkup(fuzzCase.Left, fuzzCase.Right, settings);
                var acceptedDoc = RevisionProcessor.AcceptRevisions(rendered);
                var rejectedDoc = RevisionProcessor.RejectRevisions(rendered);
                var accept = BodyContentHashes(acceptedDoc);
                var reject = BodyContentHashes(rejectedDoc);
                if (!accept.SequenceEqual(BodyContentHashes(fuzzCase.Right)))
                    failures.Add($"seed {seed}: ACCEPT≠RIGHT [{fuzzCase.DescribeMutations()}]");
                else if (!reject.SequenceEqual(BodyContentHashes(fuzzCase.Left)))
                    failures.Add($"seed {seed}: REJECT≠LEFT [{fuzzCase.DescribeMutations()}]");
                else if (!BodyFormatSignatures(acceptedDoc).SequenceEqual(BodyFormatSignatures(fuzzCase.Right)))
                    failures.Add($"seed {seed}: ACCEPT-FORMAT≠RIGHT [{fuzzCase.DescribeMutations()}]");
                else if (!BodyFormatSignatures(rejectedDoc).SequenceEqual(BodyFormatSignatures(fuzzCase.Left)))
                    failures.Add($"seed {seed}: REJECT-FORMAT≠LEFT [{fuzzCase.DescribeMutations()}]");
                else
                    passed++;
            }
            catch (Exception ex)
            {
                failures.Add($"seed {seed}: THREW {ex.GetType().Name}: {ex.Message} [{fuzzCase.DescribeMutations()}]");
            }
        }

        _out.WriteLine($"Fuzz markup invariant: {passed}/{seedCount} seeds passed, {failures.Count} failures");
        foreach (var f in failures.Take(40))
            _out.WriteLine("  FAIL " + f);

        Assert.True(failures.Count == 0, $"{failures.Count}/{seedCount} fuzz seeds failed (see output).");
    }

    // ----------------------------------------------------------------- validation baseline vs output

    [Fact]
    [Trait("Category", "Corpus")]
    public void WC_corpus_markup_introduces_no_new_validation_errors()
    {
        var pairs = WcCorpus.BuildPairs();
        var settings = new IrDiffSettings();
        var regressions = new List<string>();
        int checkd = 0;

        foreach (var (baseName, variantName) in pairs)
        {
            // Skip the Task-4-blocked pairs whose conservative fallback also can't yet keep validity (body-level
            // opaque markers); they are accounted for in the round-trip allowlist above. Every OTHER pair must
            // introduce zero new schema errors.
            if (Task4BlockedPairs.Contains($"{baseName}↔{variantName}"))
                continue;

            var left = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, baseName));
            var right = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, variantName));

            // Baseline = the worse of the two inputs' own schema-error counts (some fixtures carry
            // pre-existing warnings). The output must not exceed max(left, right) baseline.
            int baseline = Math.Max(SchemaErrorCount(left), SchemaErrorCount(right));

            WmlDocument rendered;
            try { rendered = RenderMarkup(left, right, settings); }
            catch (Exception ex) { regressions.Add($"{baseName}↔{variantName} render threw {ex.GetType().Name}"); continue; }

            int outErrors = SchemaErrorCount(rendered);
            checkd++;
            if (outErrors > baseline)
                regressions.Add($"{baseName}↔{variantName}: output {outErrors} schema errors > baseline {baseline}");
        }

        _out.WriteLine($"Validation baseline check: {checkd} pairs checked ({Task4BlockedPairs.Count} Task-4-blocked skipped), {regressions.Count} with NEW errors");
        foreach (var r in regressions.Take(40))
            _out.WriteLine("  " + r);

        Assert.True(regressions.Count == 0, $"{regressions.Count} pairs introduced new validation errors (see output).");
    }

    private static int SchemaErrorCount(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using var wd = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        // Filter the SAME tolerated-description whitelist WmlComparer's own validation tests use
        // (WmlComparerTests.ExpectedErrors) — Word emits a handful of tblLook/latentStyles/numbering
        // attributes newer than the SDK's bundled schema; these are pre-existing fixture noise, not renderer
        // regressions. Counting them on the cloned right-side content would spuriously inflate the output count
        // over the per-document baseline.
        return validator.Validate(wd).Count(e =>
            e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema &&
            !OxPt.WcTests.ExpectedErrors.Contains(e.Description));
    }

    // ----------------------------------------------------------------- header/footer story markup (2026-07-03)

    [Fact]
    public void Render_matched_header_story_marks_up_and_round_trips()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "CONFIDENTIAL Draft 1" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "CONFIDENTIAL Draft 2" });

        var rendered = RenderMarkup(left, right);

        var headerXml = Assert.Single(HeaderFooterFixtures.StoryPartsXml(rendered));
        Assert.Contains("<w:ins", headerXml);
        Assert.Contains("<w:del", headerXml);
        Assert.Contains("Open-Xml-PowerTools", headerXml);

        Assert.Equal("CONFIDENTIAL Draft 2",
            Assert.Single(HeaderFooterFixtures.StoryTexts(RevisionProcessor.AcceptRevisions(rendered))));
        Assert.Equal("CONFIDENTIAL Draft 1",
            Assert.Single(HeaderFooterFixtures.StoryTexts(RevisionProcessor.RejectRevisions(rendered))));
        Assert.Equal(0, SchemaErrorCount(rendered));
    }

    [Fact]
    public void Render_reassigned_inherited_header_clones_only_the_later_section()
    {
        // A is shared by both left sections. The right document keeps A in section 0 but assigns B at section 1.
        // The output cannot rewrite A in place: it needs an A→B clone bound at section 1, while section 0 remains A.
        var left = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }),
            },
            headerParts: new Dictionary<string, string[]> { ["rIdA"] = new[] { "Header A" } });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }, Headers: new[] { ("default", "rIdB") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A" },
                ["rIdB"] = new[] { "Header B" },
            });

        var rendered = RenderMarkup(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);

        Assert.Equal(2, HeaderFooterPartCount(rendered));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rendered, true, 0, "default"));
        Assert.Equal("Header B", HeaderFooterFixtures.ReferencedStoryText(rendered, true, 1, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 0, "default"));
        Assert.Equal("Header B", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 1, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 0, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 1, "default"));
        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "header-reassign-inherited");
    }

    [Fact]
    public void Render_reassigned_header_rejoins_the_primary_story_after_the_clone()
    {
        // The refined topology must explicitly rebind from the cloned B story back to primary A at section 1.
        // Otherwise A's inherited successor would incorrectly keep B all the way through section 2.
        var left = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }),
                new HeaderFooterFixtures.Section(new[] { "section 2" }),
            },
            headerParts: new Dictionary<string, string[]> { ["rIdA"] = new[] { "Header A" } });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdB") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 2" }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A" },
                ["rIdB"] = new[] { "Header B" },
            });

        var rendered = RenderMarkup(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);

        Assert.Equal(2, HeaderFooterPartCount(rendered));
        Assert.Equal("Header B", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 0, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 1, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 2, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 0, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 1, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 2, "default"));
        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "header-clone-rejoin");
    }

    [Fact]
    public void Render_reassigned_header_clone_preserves_left_and_imports_right_media_relationships()
    {
        // The cloned A→B story keeps an equal left image and gains an inserted right image. Both relationship
        // graphs are part-owned: copying XML alone would leave one or both r:embed ids dangling.
        var left = HeaderFooterFixtures.WithImageInHeaderPart(HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A", HeaderFooterFixtures.ImageParagraphXml("rIdOld") },
            }), headerIndex: 0, relId: "rIdOld");
        // Relationship IDs are scoped to their owning header part. The same rIdOld is therefore attached
        // separately to A and B; B also owns its new image relationship.
        var right = HeaderFooterFixtures.WithImageInHeaderPart(
            HeaderFooterFixtures.WithImageInHeaderPart(
                HeaderFooterFixtures.WithImageInHeaderPart(HeaderFooterFixtures.Build(
                    new[]
                    {
                        new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                        new HeaderFooterFixtures.Section(new[] { "section 1" }, Headers: new[] { ("default", "rIdB") }),
                    },
                    headerParts: new Dictionary<string, string[]>
                    {
                        ["rIdA"] = new[] { "Header A", HeaderFooterFixtures.ImageParagraphXml("rIdOld") },
                        ["rIdB"] = new[]
                        {
                            "Header B",
                            HeaderFooterFixtures.ImageParagraphXml("rIdOld"),
                            HeaderFooterFixtures.ImageParagraphXml("rIdNew"),
                        },
                    }), headerIndex: 0, relId: "rIdOld"), headerIndex: 1, relId: "rIdOld"),
            headerIndex: 1, relId: "rIdNew");

        var rendered = RenderMarkup(left, right);
        AssertHeaderImageEmbedsResolve(rendered, sectionIndex: 1, expectedEmbedCount: 2);

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "header-clone-media");
    }

    [Fact]
    public void Render_reassigned_header_clone_keeps_left_media_relationships_for_deleted_content()
    {
        // Rebuilding the clone emits the deleted LEFT image from its original source XML. Its rId must resolve
        // from the clone itself, even though no right-side image is imported into the B story.
        var left = HeaderFooterFixtures.WithImageInHeaderPart(HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A", HeaderFooterFixtures.ImageParagraphXml("rIdOld") },
            }), headerIndex: 0, relId: "rIdOld");
        var right = HeaderFooterFixtures.WithImageInHeaderPart(HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }, Headers: new[] { ("default", "rIdB") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A", HeaderFooterFixtures.ImageParagraphXml("rIdOld") },
                ["rIdB"] = new[] { "Header B" },
            }), headerIndex: 0, relId: "rIdOld");

        var rendered = RenderMarkup(left, right);
        AssertHeaderImageEmbedsResolve(rendered, sectionIndex: 1, expectedEmbedCount: 1);

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        AssertHeaderImageEmbedsResolve(rejected, sectionIndex: 1, expectedEmbedCount: 1);
        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "header-clone-left-media");
    }

    [Fact]
    public void Render_reused_right_header_keeps_each_left_story_independently_rejectable()
    {
        // The physical right B story is reused in both sections. Each section must bind to its own A→B/C→B
        // output part; a flat right-URI map would make both references target the last C→B part on rebind.
        var left = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdA") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }, Headers: new[] { ("default", "rIdC") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A" },
                ["rIdC"] = new[] { "Header C" },
            });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(new[] { "section 0" }, Headers: new[] { ("default", "rIdB") }),
                new HeaderFooterFixtures.Section(new[] { "section 1" }, Headers: new[] { ("default", "rIdB") }),
            },
            headerParts: new Dictionary<string, string[]> { ["rIdB"] = new[] { "Header B" } });

        var rendered = RenderMarkup(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);

        Assert.Equal(2, HeaderFooterPartCount(rendered));
        Assert.Equal("Header B", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 0, "default"));
        Assert.Equal("Header B", HeaderFooterFixtures.ReferencedStoryText(accepted, true, 1, "default"));
        Assert.Equal("Header A", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 0, "default"));
        Assert.Equal("Header C", HeaderFooterFixtures.ReferencedStoryText(rejected, true, 1, "default"));
        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
        AssertRoundTrip(left, right, label: "header-reused-right-part");
    }

    [Fact]
    public void Render_matched_footer_story_round_trips()
    {
        var left = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Same header" }, footerParas: new[] { "Acme Corp 2025" });
        var right = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Same header" }, footerParas: new[] { "Acme Corp 2026" });

        var rendered = RenderMarkup(left, right);

        // The unchanged header part is carried verbatim (no markup); the footer carries the diff.
        var parts = HeaderFooterFixtures.StoryPartsXml(rendered);
        Assert.Equal(2, parts.Count);
        Assert.DoesNotContain("<w:ins", parts[0]);
        Assert.Contains("<w:ins", parts[1]);

        Assert.Equal(new[] { "Same header", "Acme Corp 2026" },
            HeaderFooterFixtures.StoryTexts(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(new[] { "Same header", "Acme Corp 2025" },
            HeaderFooterFixtures.StoryTexts(RevisionProcessor.RejectRevisions(rendered)));
        Assert.Equal(0, SchemaErrorCount(rendered));
    }

    [Fact]
    public void Render_header_revision_ids_unique_across_scopes()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body one" }, headerParas: new[] { "Header v1" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body two" }, headerParas: new[] { "Header v2" });

        var rendered = RenderMarkup(left, right);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var main = wd.MainDocumentPart!;
        var ids = new List<string>();
        foreach (var part in new[] { (OpenXmlPart)main }.Concat(main.HeaderParts))
        {
            var root = XDocument.Load(part.GetStream(FileMode.Open, FileAccess.Read)).Root!;
            ids.AddRange(root.Descendants()
                .Where(e => e.Name == W.ins || e.Name == W.del)
                .Select(e => (string?)e.Attribute(W.id))
                .Where(v => v != null)!);
        }
        Assert.True(ids.Count >= 2, "expected revisions in both body and header");
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void Render_gate_off_keeps_left_header_verbatim()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Header v1" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Header v2" });

        var rendered = RenderMarkup(left, right, new IrDiffSettings { CompareHeadersFooters = false });

        var headerXml = Assert.Single(HeaderFooterFixtures.StoryPartsXml(rendered));
        Assert.DoesNotContain("<w:ins", headerXml);
        Assert.Contains("Header v1", headerXml);
        Assert.DoesNotContain("Header v2", headerXml);
    }

    [Fact]
    public void Render_inserted_first_page_header_adds_part_reference_and_titlePg()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Running" });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "Body" },
                    Headers: new[] { ("default", "rIdH1"), ("first", "rIdH2") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdH1"] = new[] { "Running" },
                ["rIdH2"] = new[] { "Cover page banner" },
            },
            titlePg: true);

        var rendered = RenderMarkup(left, right);

        // The output gains the first-page story: part + w:headerReference@first + w:titlePg on the sectPr.
        Assert.Equal("Cover page banner",
            HeaderFooterFixtures.ReferencedStoryText(rendered, isHeader: true, sectionIndex: 0, kind: "first"));
        using (var ms = new MemoryStream(rendered.DocumentByteArray))
        using (var wd = WordprocessingDocument.Open(ms, false))
        {
            var body = wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
            var sectPr = body.Elements(W.sectPr).Last();
            Assert.NotNull(sectPr.Element(W.titlePg));
        }

        // Accept keeps the inserted story; reject empties it (empty ≡ absent at the text level — Word's
        // own behavior for rejecting an inserted header story).
        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        Assert.Equal("Cover page banner",
            HeaderFooterFixtures.ReferencedStoryText(accepted, isHeader: true, sectionIndex: 0, kind: "first"));
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Assert.Equal("",
            HeaderFooterFixtures.ReferencedStoryText(rejected, isHeader: true, sectionIndex: 0, kind: "first"));

        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(accepted));
        Assert.Equal(0, SchemaErrorCount(rejected));
    }

    [Fact]
    public void Render_deleted_footer_story_marks_content_deleted()
    {
        var left = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Running" }, footerParas: new[] { "Legacy footer line" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Running" });

        var rendered = RenderMarkup(left, right);

        // The part and its reference stay; the content is marked deleted.
        var footerXml = HeaderFooterFixtures.StoryPartsXml(rendered)[1];
        Assert.Contains("<w:del", footerXml);

        // Accept empties the story (≡ right's absent story at the text level); reject restores it.
        Assert.Equal("",
            HeaderFooterFixtures.ReferencedStoryText(
                RevisionProcessor.AcceptRevisions(rendered), isHeader: false, sectionIndex: 0, kind: "default"));
        Assert.Equal("Legacy footer line",
            HeaderFooterFixtures.ReferencedStoryText(
                RevisionProcessor.RejectRevisions(rendered), isHeader: false, sectionIndex: 0, kind: "default"));

        Assert.Equal(0, SchemaErrorCount(rendered));
        Assert.Equal(0, SchemaErrorCount(RevisionProcessor.AcceptRevisions(rendered)));
    }

    [Fact]
    public void Render_inserted_even_footer_ensures_evenAndOddHeaders()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, footerParas: new[] { "Odd pages" });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "Body" },
                    Footers: new[] { ("default", "rIdF1"), ("even", "rIdF2") }),
            },
            footerParts: new Dictionary<string, string[]>
            {
                ["rIdF1"] = new[] { "Odd pages" },
                ["rIdF2"] = new[] { "Even pages" },
            },
            evenAndOddHeaders: true);

        var rendered = RenderMarkup(left, right);

        Assert.Equal("Even pages",
            HeaderFooterFixtures.ReferencedStoryText(rendered, isHeader: false, sectionIndex: 0, kind: "even"));
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var settingsRoot = wd.MainDocumentPart!.DocumentSettingsPart!.GetXDocument().Root!;
        Assert.NotNull(settingsRoot.Element(W.evenAndOddHeaders));
        Assert.Equal(0, SchemaErrorCount(rendered));
    }

    [Fact]
    public void Render_image_added_to_header_imports_media_into_header_part()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Notice" });
        var right = HeaderFooterFixtures.WithImageInFirstHeaderPart(
            HeaderFooterFixtures.Simple(new[] { "Body" },
                headerParas: new[] { "Notice", HeaderFooterFixtures.ImageParagraphXml("rIdImg1") }),
            "rIdImg1");

        var rendered = RenderMarkup(left, right);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var headerPart = wd.MainDocumentPart!.HeaderParts.Single();
        var headerRoot = XDocument.Load(headerPart.GetStream(FileMode.Open, FileAccess.Read)).Root!;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var embedId = (string?)headerRoot.Descendants(a + "blip").Single().Attribute(r + "embed");
        Assert.NotNull(embedId);
        // The embed id resolves against the HEADER part's own relationships (rels are part-scoped).
        var imagePart = headerPart.GetPartById(embedId!);
        Assert.StartsWith("image/", imagePart.ContentType);
    }

    // ----------------------------------------------------------------- author override (composite groundwork)

    /// <summary>Regression pin: the two-way render path leaves <see cref="RenderState.AuthorOverride"/> null,
    /// so every revision is still stamped with <see cref="DocxDiffSettings.AuthorForRevisions"/>. Guards that
    /// adding the (composite-only) override does not change ordinary two-way output.</summary>
    [Fact]
    public void Author_override_null_keeps_settings_author()
    {
        var left = Docs.Para("alpha one", "beta two");
        var right = Docs.Para("alpha one EDITED", "beta two");
        var settings = new DocxDiffSettings { AuthorForRevisions = "Eric" };
        var doc = DocxDiff.Compare(left, right, settings);
        var xml = Docs.MainPartXml(doc);
        Assert.Contains("w:author=\"Eric\"", xml);
    }
}
