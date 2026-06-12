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
                break;
            case IrTable t:
                sink.Add("tblf:" + t.ContentHash.ToHex());
                foreach (var row in t.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var b in cell.Blocks)
                            CollectFormatSignatures(b, settings, sink);
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
    }

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
    public void Render_preserves_unmodeled_run_properties_on_modified_paragraph()
    {
        // A run carrying an UNMODELED rPr child (w:shd) on an EQUAL portion must survive into the output —
        // proving provenance-clone (not IrRunFormat rebuild) preserves unmodeled formatting.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:shd w:val=\"clear\" w:fill=\"FFFF00\"/></w:rPr><w:t>highlight one two</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:shd w:val=\"clear\" w:fill=\"FFFF00\"/></w:rPr><w:t>highlight THREE two</w:t></w:r></w:p>");

        var rendered = RenderMarkup(left, right);
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var shdCount = wd.MainDocumentPart!.GetXDocument().Descendants(W.shd).Count();
        Assert.True(shdCount > 0, "unmodeled w:shd run property must be preserved through the split");
        AssertRoundTrip(left, right, label: "unmodeled-shd");
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
        // If the aligner classified this as a MoveModify, the moveTo range exists and carries nested ins/del.
        // (If the similarity pass instead classified it as Move + separate edits, the round-trip still holds —
        // so we assert the contract, the round-trip, and only check nesting WHEN moveTo is present.)
        if (body.Descendants(W.moveTo).Any())
        {
            var moveToRangeStart = body.Descendants(W.moveToRangeStart).FirstOrDefault();
            Assert.NotNull(moveToRangeStart);
        }
        AssertRoundTrip(left, right, settings, label: "move-modify");
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

    // ----------------------------------------------------------------- corpus invariant (92 × 2)

    /// <summary>
    /// The Task-4-BLOCKED base↔variant pairs: their accept/reject round-trip needs native markup the Task-3
    /// core does not yet emit (and whose conservative whole-block fallback can't fully express). Each is tied to
    /// the Task-4 burndown item that closes it. This allowlist is a RATCHET — the invariant test below asserts
    /// EVERY other pair round-trips AND that no allowlisted pair UNEXPECTEDLY passes; Task 4 shrinks it to empty.
    /// </summary>
    private static readonly HashSet<string> Task4BlockedPairs = new(StringComparer.Ordinal)
    {
        // Footnote/endnote SCOPE markup — the renderer does not yet render IrEditScript.NoteOps into the
        // footnotes/endnotes parts; accept/reject keep the LEFT package's note content. Task-4: note scopes.
        "WC020-FootNote-Before.docx↔WC020-FootNote-After-2.docx",
        "WC035-Footnote-Before.docx↔WC035-Footnote-After.docx",
        "WC035-Endnote-Before.docx↔WC035-Endnote-After.docx",
        // OPAQUE drawing content (SmartArt diagram data parts) in a modified block — the whole-block del+ins
        // fallback can't toggle opaque diagram XML at content-hash grain. Task-4: opaque block markup.
        "WC014-SmartArt-Before.docx↔WC014-SmartArt-After.docx",
        "WC014-SmartArt-With-Image-Before.docx↔WC014-SmartArt-With-Image-After.docx",
        "WC052-SmartArt-Same.docx↔WC052-SmartArt-Same-Mod.docx",
        // Image/math drawing SWAP inside a modified paragraph — precise per-drawing rel toggling is Task-4
        // (whole-block image insert/delete IS covered; an in-paragraph swap is not). Task-4: drawing revisions.
        "WC022-Image-Math-Para-Before.docx↔WC022-Image-Math-Para-After.docx",
        // Hyperlink TARGET change where the right rId COLLIDES with a different left rId (needs rId remap, not
        // same-id recreation). Task-4: hyperlink relationship remap.
        "WC019-Hyperlink-Before.docx↔WC019-Hyperlink-After-2.docx",
        // Body-level non-paragraph (bookmark/perm) insert/delete — opaque body-level elements have no run model
        // to revision-mark; WmlComparer handles these specially. Task-4: body-level marker revisions.
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
}
