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
        // TABLE-structural: a paragraph moved INTO a table, and a modify adjacent to a table — proper row/cell
        // revision markup (Task 4) is needed; the whole-table del+ins fallback leaves an empty shell on the
        // toggled side. Task-4: table row/cell revision markup.
        "WC007-Unmodified.docx↔WC007-Moved-into-Table.docx",
        "WC010-Para-Before-Table-Unmodified.docx↔WC010-Para-Before-Table-Mod.docx",
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
                    var accept = BodyContentHashes(RevisionProcessor.AcceptRevisions(rendered));
                    var reject = BodyContentHashes(RevisionProcessor.RejectRevisions(rendered));
                    if (!accept.SequenceEqual(BodyContentHashes(r)))
                        failure = $"{key} [{dir}] ACCEPT≠RIGHT";
                    else if (!reject.SequenceEqual(BodyContentHashes(l)))
                        failure = $"{key} [{dir}] REJECT≠LEFT";
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
                var accept = BodyContentHashes(RevisionProcessor.AcceptRevisions(rendered));
                var reject = BodyContentHashes(RevisionProcessor.RejectRevisions(rendered));
                if (!accept.SequenceEqual(BodyContentHashes(fuzzCase.Right)))
                    failures.Add($"seed {seed}: ACCEPT≠RIGHT [{fuzzCase.DescribeMutations()}]");
                else if (!reject.SequenceEqual(BodyContentHashes(fuzzCase.Left)))
                    failures.Add($"seed {seed}: REJECT≠LEFT [{fuzzCase.DescribeMutations()}]");
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
