#nullable enable

using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Cross-paragraph token-stream diff over word-matched runs (<see cref="IrDiffSettings.CrossParagraphTokenDiff"/>,
/// 2026-07-25) — the within-run flat word+pilcrow stream decoded from Word's compare output: within a run of
/// adjacent word-matched paragraph pairs, retained words from base paragraph k can land in output paragraph
/// k±1 (crossing the pilcrow), pilcrows are stream tokens marked ¶INS/¶DEL by side, and the output paragraph
/// count follows the token-level LCS interleave (a 2×2 word-matched region can emit 3 paragraphs). The gate:
/// with the setting ON, a run of ≥2 adjacent Modified pairs renders via ONE
/// <see cref="IrEditOpKind.CrossParagraphRunBlock"/> op yet still satisfies the sacred round-trip contract —
/// <c>accept ≡ right</c> and <c>reject ≡ left</c> (<see cref="Docs.PlainText"/> encodes both per-paragraph
/// text AND paragraph count via the newline join, so a stray/missing/fused paragraph is caught). Fixtures put
/// an unchanged paragraph AFTER the run so the fuse-into-the-following-block failure mode is observable.
/// </summary>
[Trait("Category", "Markup")]
public class DocxDiffCrossParagraphTests
{
    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    private static IrEditScript BuildScript(WmlDocument left, WmlDocument right, bool crossParagraph)
    {
        var settings = new IrDiffSettings { CrossParagraphTokenDiff = crossParagraph };
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        return IrEditScriptBuilder.Build(irLeft, irRight, settings);
    }

    private static WmlDocument RenderMarkup(WmlDocument left, WmlDocument right, bool crossParagraph)
    {
        var settings = new IrDiffSettings { CrossParagraphTokenDiff = crossParagraph };
        var script = BuildScript(left, right, crossParagraph);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    /// <summary>Assert the fused render round-trips (accept≡right, reject≡left by block text) AND that the
    /// cross-paragraph path was actually exercised (a CrossParagraphRunBlock op was produced).</summary>
    private static void AssertFusedRoundTrip(
        WmlDocument left, WmlDocument right, int expectedFusedOps, string label)
    {
        var script = BuildScript(left, right, crossParagraph: true);
        int fusedOps = script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);
        Assert.True(fusedOps == expectedFusedOps,
            $"{label}: expected {expectedFusedOps} CrossParagraphRunBlock op(s), got {fusedOps}.");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);

        Assert.Equal(Docs.PlainText(right), Docs.PlainText(accepted));   // accept ≡ right
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(rejected));    // reject ≡ left
    }

    /// <summary>
    /// THE decoded law's minimal pin: a 2×2 word-matched region (two adjacent Modified pairs) whose flat
    /// word+pilcrow stream carries a retained word ("charlie") ACROSS the pair boundary — base paragraph 1's
    /// word lands in output paragraph 2. The token-level LCS interleave emits THREE output paragraphs for the
    /// region (¶INS closing "alpha bravo", ¶DEL closing the crossing "charlie", the retained final pilcrow) —
    /// the shape no per-paragraph model reproduces — and accept/reject reconstruct the right/left sides
    /// exactly, paragraph count included.
    /// </summary>
    [Fact]
    public void CrossParagraph_2x2_boundary_crossing_word_emits_three_paragraphs_and_round_trips()
    {
        var left = Docs.Para("HEAD",
            "alpha bravo charlie",
            "delta echo foxtrot",
            "TAIL");
        var right = Docs.Para("HEAD",
            "alpha bravo",
            "charlie delta echo foxtrot",
            "TAIL");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "2x2-boundary-crossing");

        // The rendered region is THREE paragraphs (HEAD + 3 + TAIL = 5 body w:p), with the crossing word
        // "charlie" rendered ONCE as a shared (unwrapped) run — retained content, not a del/ins pair.
        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        var paras = body.Elements(ns + "p").ToList();
        Assert.Equal(5, paras.Count);

        var charlieRuns = body.Descendants(ns + "t").Where(t => t.Value.Contains("charlie")).ToList();
        var charlieDel = body.Descendants(ns + "delText").Where(t => t.Value.Contains("charlie")).ToList();
        Assert.Single(charlieRuns);
        Assert.Empty(charlieDel);
        Assert.DoesNotContain(charlieRuns[0].Ancestors(), a => a.Name == ns + "ins" || a.Name == ns + "del");
    }

    /// <summary>A run of 3 word-matched pairs where content flows across BOTH interior boundaries. The single
    /// fused op must reproduce right (accept) and left (reject) exactly — paragraph counts included.</summary>
    [Fact]
    public void CrossParagraph_3x3_flow_round_trips()
    {
        var left = Docs.Para("HEAD",
            "one two three four",
            "five six seven eight",
            "nine ten eleven twelve",
            "TAIL");
        var right = Docs.Para("HEAD",
            "one two three",
            "four five six seven",
            "eight nine ten eleven twelve",
            "TAIL");
        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "3x3-flow");
    }

    /// <summary>Two word-matched pairs whose pilcrows all match (no cross-boundary word movement): the fused
    /// path must be output-equivalent to the per-pair path — same paragraph count, same round-trip.</summary>
    [Fact]
    public void CrossParagraph_aligned_pairs_stay_two_paragraphs_and_round_trip()
    {
        // Cleanly-aligned pairs (every residue re-matches within its own pair, every boundary
        // pilcrow pairs): the STRUCTURE gate declines the fusion — a fused stream that changes no
        // paragraph structure would only redistribute within-pair anchors versus the ordinary
        // per-pair differ. The run falls back to the per-pair path, which produces the identical
        // 4-paragraph shape and round-trip.
        var left = Docs.Para("HEAD",
            "alpha bravo charlie delta",
            "echo foxtrot golf hotel",
            "TAIL");
        var right = Docs.Para("HEAD",
            "alpha bravo charlie replaced",
            "echo foxtrot golf changed",
            "TAIL");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 0, "aligned-pairs");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        Assert.Equal(4, body.Elements(ns + "p").Count()); // HEAD + 2 + TAIL — no extra paragraphs invented
    }

    /// <summary>
    /// Block-format history on the fused run's SHARED (Equal-mark) pilcrows: when the paired paragraphs'
    /// pPr differ (justify → center here), the fused output paragraph carries the RIGHT props plus a
    /// <c>w:pPrChange</c> archiving the LEFT props — so reject restores the left alignment at the
    /// property level, exactly like the per-pair path's <c>ApplyBlockFormatChanges</c>.
    /// </summary>
    [Fact]
    public void CrossParagraph_equal_cell_tracks_paragraph_format_change_and_reject_restores_it()
    {
        static string Body(string jc, string p1, string p2) =>
            $"<w:p><w:pPr><w:jc w:val=\"{jc}\"/></w:pPr><w:r><w:t>{p1}</w:t></w:r></w:p>" +
            $"<w:p><w:pPr><w:jc w:val=\"{jc}\"/></w:pPr><w:r><w:t>{p2}</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>TAIL</w:t></w:r></w:p>";

        var left = IrTestDocuments.FromBodyXml(Body("both", "alpha bravo charlie delta", "echo foxtrot golf hotel"));
        var right = IrTestDocuments.FromBodyXml(Body("center", "alpha bravo charlie word", "echo foxtrot golf word"));

        // Cleanly-aligned pairs: the STRUCTURE gate declines the fusion (no paragraph-structure
        // change) and the run falls back to the per-pair path — whose ApplyBlockFormatChanges
        // produces the identical pPrChange discipline asserted below.
        var script = BuildScript(left, right, crossParagraph: true);
        Assert.DoesNotContain(script.Operations, o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var renderedBody = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;

        // Both fused paragraphs carry right props (center) + a pPrChange archiving the left props (both).
        var changed = renderedBody.Descendants(ns + "pPrChange").ToList();
        Assert.Equal(2, changed.Count);
        Assert.All(changed, ch =>
            Assert.Equal("both", (string?)ch.Element(ns + "pPr")?.Element(ns + "jc")?.Attribute(ns + "val")));

        string JcVals(WmlDocument d) => string.Join(",",
            XDocument.Parse(Docs.MainPartXml(d)).Root!.Element(ns + "body")!.Elements(ns + "p")
                .Select(p => (string?)p.Element(ns + "pPr")?.Element(ns + "jc")?.Attribute(ns + "val") ?? "-"));

        Assert.Equal(JcVals(right), JcVals(RevisionProcessor.AcceptRevisions(rendered))); // accept keeps center
        Assert.Equal(JcVals(left), JcVals(RevisionProcessor.RejectRevisions(rendered)));  // reject restores both
    }

    /// <summary>A single word-matched pair (run of 1) is NOT fused — the flat-stream shape is decoded for
    /// runs of ≥2 pairs only; a lone pair keeps its ordinary per-pair token diff.</summary>
    [Fact]
    public void CrossParagraph_single_pair_is_not_fused()
    {
        var left = Docs.Para("HEAD", "alpha bravo charlie delta", "TAIL");
        var right = Docs.Para("HEAD", "alpha bravo charlie changed", "TAIL");

        var script = BuildScript(left, right, crossParagraph: true);
        Assert.DoesNotContain(script.Operations, o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));
    }

    // ---------------------------------------------- run + story-final tail (one-sided members, 2026-07-27)

    /// <summary>
    /// THE decoded run+gap law ("justified body → large font"): a word-matched pair whose LONG left tail
    /// crosses forward into the paragraph carrying a trailing INSERT, retaining a shared word there. The
    /// aligner pairs the 12-word body with the 3-word rewrite (≥2 shared content words waive the size-parity
    /// veto), absorbs the surplus insert as a story-tail split, and the builder streams run+tail as ONE op:
    /// [E title][I pair-head][E final fused cell] — the crossing word ("foxtrot") renders ONCE as retained
    /// content inside the final paragraph, and accept/reject reconstruct right/left exactly, paragraph
    /// counts included.
    /// </summary>
    [Fact]
    public void CrossParagraph_pair_with_trailing_insert_tail_streams_and_round_trips()
    {
        var left = Docs.Para(
            "Title Zulu Demo",
            "alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima");
        var right = Docs.Para(
            "Title Yankee Demo",
            "alpha bravo charlie",
            "whiskey november foxtrot oscar papa");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "pair+insert-tail");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        Assert.Equal(3, body.Elements(ns + "p").Count()); // title + pair cell + final fused cell

        // The crossing word is retained — rendered once, outside any ins/del wrapper.
        var foxtrotRuns = body.Descendants(ns + "t").Where(t => t.Value.Contains("foxtrot")).ToList();
        Assert.Single(foxtrotRuns);
        Assert.Empty(body.Descendants(ns + "delText").Where(t => t.Value.Contains("foxtrot")));
        Assert.DoesNotContain(foxtrotRuns[0].Ancestors(), a => a.Name == ns + "ins" || a.Name == ns + "del");
    }

    /// <summary>
    /// Run + trailing REPLACE gap (deleted AND inserted paragraphs at the story end): the pair's residue
    /// crosses into the trailing insert ("golf" retained there), the unpaired gap paragraphs ride the
    /// stream one-sided, and accept/reject reconstruct both sides exactly.
    /// </summary>
    [Fact]
    public void CrossParagraph_run_with_trailing_replace_gap_streams_and_round_trips()
    {
        var left = Docs.Para(
            "Title Zulu Demo",
            "alpha bravo charlie delta echo golf",
            "xray yankee zulu");
        var right = Docs.Para(
            "Title Zulu Demo",
            "alpha bravo charlie",
            "hotel india golf juliet");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "run+replace-gap");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        var golfRuns = body.Descendants(ns + "t").Where(t => t.Value.Contains("golf")).ToList();
        Assert.Single(golfRuns);
        Assert.Empty(body.Descendants(ns + "delText").Where(t => t.Value.Contains("golf")));
        Assert.DoesNotContain(golfRuns[0].Ancestors(), a => a.Name == ns + "ins" || a.Name == ns + "del");
    }

    /// <summary>
    /// The tail STRUCTURE gate: a run whose trailing replace gap shares NO content with the run (zero
    /// strictly-crossing matches) must NOT stream — one-sided absorption with no cross-flow falls back to
    /// the ordinary replace-gap grammar, which still round-trips.
    /// </summary>
    [Fact]
    public void CrossParagraph_tail_without_crossings_is_not_fused_and_round_trips()
    {
        var left = Docs.Para(
            "Title Zulu Demo",
            "alpha bravo charlie delta",
            "xray yankee zulu");
        var right = Docs.Para(
            "Title Zulu Demo",
            "alpha bravo charlie",
            "hotel india juliet");

        var script = BuildScript(left, right, crossParagraph: true);
        Assert.Equal(0, CountFusedOps(script));

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));
    }

    /// <summary>
    /// A GENUINE story-final split behind a word-matched pair also streams (the split's surplus member is
    /// the tail; the singular's residue crossing into it always fires the gate) — and the stream shape
    /// degenerates to the split grammar: the split text stays fully retained (no del/ins), only the new
    /// interior pilcrow is inserted, and accept/reject reconstruct both sides exactly.
    /// </summary>
    [Fact]
    public void CrossParagraph_story_final_genuine_split_behind_pair_streams_and_round_trips()
    {
        var left = Docs.Para(
            "HEAD alpha bravo",
            "one two three four five six");
        var right = Docs.Para(
            "HEAD alpha changed",
            "one two three",
            "four five six");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "genuine-split-tail");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        // The split content is retained, not rewritten: no delText anywhere in the split region.
        foreach (var word in new[] { "one", "two", "three", "four", "five", "six" })
        {
            Assert.Single(body.Descendants(ns + "t").Where(t => t.Value.Contains(word)));
            Assert.Empty(body.Descendants(ns + "delText").Where(t => t.Value.Contains(word)));
        }
    }

    /// <summary>
    /// A story-final MERGE tail without cross-flow keeps the ordinary merge grammar (the merge pair's
    /// in-slot anchors seal the windows, so no strictly-crossing match can fire) — no fused op, and the
    /// merge still round-trips.
    /// </summary>
    [Fact]
    public void CrossParagraph_story_final_merge_without_crossing_keeps_merge_grammar()
    {
        var left = Docs.Para(
            "HEAD alpha bravo",
            "one two three",
            "four five six");
        var right = Docs.Para(
            "HEAD alpha changed",
            "one two three four five six");

        var script = BuildScript(left, right, crossParagraph: true);
        Assert.Equal(0, CountFusedOps(script));
        Assert.Contains(script.Operations, o => o.Kind == IrEditOpKind.MergeBlock);

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));
    }

    // ------------------------------------------- story-final pair retention law (decoded, 2026-07-27)

    /// <summary>
    /// THE decoded story-final retention law ("centered bold" class): the story's last word-matched pair
    /// is mostly rewritten and its ISOLATED same-ordinal matches ("golfington"/"hoteliers" — no common
    /// prefix, not an adjacent bigram) do NOT retain in place, even though they outweigh the crossing
    /// chain in matched characters. Suppressing them lets the interior pair's residue chain
    /// ("charlie delta") cross forward into the final right paragraph: the run streams, the interior
    /// boundary splits (¶INS + ¶DEL), the crossing words render ONCE as retained content, and the final
    /// left paragraph is wholly struck. Before the law, the isolated recoveries sealed the window and
    /// the whole run fell back to three per-pair paragraphs.
    /// </summary>
    [Fact]
    public void CrossParagraph_story_final_isolated_matches_yield_to_the_crossing_chain()
    {
        var left = Docs.Para("HEAD",
            "alpha bravo charlie delta echo",
            "golfington mike hoteliers november india");
        var right = Docs.Para("HEAD",
            "alpha zulu bravo yankee xray",
            "whiskey golfington charlie hoteliers delta");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "story-final-isolated-yield");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        Assert.Equal(4, body.Elements(ns + "p").Count()); // HEAD + ¶INS cell + ¶DEL crossing cell + struck final

        foreach (var word in new[] { "charlie", "delta" })
        {
            var runs = body.Descendants(ns + "t").Where(t => t.Value.Contains(word)).ToList();
            Assert.Single(runs);
            Assert.Empty(body.Descendants(ns + "delText").Where(t => t.Value.Contains(word)));
            Assert.DoesNotContain(runs[0].Ancestors(), a => a.Name == ns + "ins" || a.Name == ns + "del");
        }
    }

    /// <summary>
    /// The law's VALVE ("italic underline" class): when the story-final pair's same-ordinal matches
    /// include an adjacent BIGRAM ("golfington hoteliers" contiguous on BOTH sides), the pair genuinely
    /// retains that content in place — the valve stands, no crossing is invented, the structure gate
    /// declines, and the run keeps its per-pair three-paragraph shape.
    /// </summary>
    [Fact]
    public void CrossParagraph_story_final_bigram_retains_in_place_and_declines_fusion()
    {
        var left = Docs.Para("HEAD",
            "alpha bravo charlie delta echo",
            "golfington hoteliers mike november india");
        var right = Docs.Para("HEAD",
            "alpha zulu bravo yankee xray",
            "whiskey golfington hoteliers charlie delta");

        var script = BuildScript(left, right, crossParagraph: true);
        Assert.Equal(0, CountFusedOps(script));

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        Assert.Equal(3, body.Elements(ns + "p").Count()); // no paragraph structure invented
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));
    }

    /// <summary>
    /// Suppression alone never INVENTS structure ("yellow highlight" class): the story-final pair keeps
    /// its common unit PREFIX ("sierra") plus an isolated interior match, and NO residue of an earlier
    /// pair can cross. With the isolated match suppressed there are zero crossings, every boundary
    /// pairs, the structure gate declines, and the run falls back to the ordinary per-pair path — same
    /// paragraph count, prefix retained, exact round-trip.
    /// </summary>
    [Fact]
    public void CrossParagraph_story_final_prefix_without_crossing_never_invents_structure()
    {
        var left = Docs.Para("HEAD",
            "alpha bravo charlie delta",
            "sierra tango uniform victor");
        var right = Docs.Para("HEAD",
            "alpha bravo xray yankee",
            "sierra whiskey uniform zulu");

        var script = BuildScript(left, right, crossParagraph: true);
        Assert.Equal(0, CountFusedOps(script));

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        Assert.Equal(3, body.Elements(ns + "p").Count());
        var sierraRuns = body.Descendants(ns + "t").Where(t => t.Value.Contains("sierra")).ToList();
        Assert.Single(sierraRuns);
        Assert.DoesNotContain(sierraRuns[0].Ancestors(), a => a.Name == ns + "ins" || a.Name == ns + "del");
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));
    }

    /// <summary>With the setting OFF (the default), the same fixtures produce NO CrossParagraphRunBlock op —
    /// the byte-identical baseline. (The parity scoreboards are the broad OFF-is-byte-identical canary; this
    /// is the focused local guard.)</summary>
    [Fact]
    public void Setting_off_never_emits_a_fused_op_and_still_round_trips()
    {
        var left = Docs.Para("HEAD", "alpha bravo charlie", "delta echo foxtrot", "TAIL");
        var right = Docs.Para("HEAD", "alpha bravo", "charlie delta echo foxtrot", "TAIL");

        var script = BuildScript(left, right, crossParagraph: false);
        Assert.DoesNotContain(script.Operations, o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);

        var rendered = RenderMarkup(left, right, crossParagraph: false);
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));
    }

    /// <summary>
    /// Fusion is gated to the top-level BODY projection only. A word-matched run INSIDE a table cell (a
    /// nested scope the fused renderer is not verified for) keeps the ordinary per-pair path even with the
    /// setting on: the script carries NO CrossParagraphRunBlock op anywhere (body or nested), and the cell
    /// still round-trips. This pins the conservative scoping the merger/notes/headers/textbox paths rely on.
    /// </summary>
    [Fact]
    public void CrossParagraph_run_inside_table_cell_is_not_fused_and_round_trips()
    {
        static string Body(string p1, string p2) =>
            "<w:p><w:r><w:t>HEAD</w:t></w:r></w:p>" +
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"9000\"/></w:tblGrid>" +
            "<w:tr><w:tc>" +
            $"<w:p><w:r><w:t>{p1}</w:t></w:r></w:p>" +
            $"<w:p><w:r><w:t>{p2}</w:t></w:r></w:p>" +
            "</w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:t>TAIL</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXml(Body("alpha bravo charlie", "delta echo foxtrot"));
        var right = IrTestDocuments.FromBodyXml(Body("alpha bravo changed", "delta echo changed"));

        var script = BuildScript(left, right, crossParagraph: true);
        Assert.Equal(0, CountFusedOps(script)); // gated out of the cell scope

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        Assert.Equal(Docs.StructuralBody(right), Docs.StructuralBody(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.StructuralBody(left), Docs.StructuralBody(RevisionProcessor.RejectRevisions(rendered)));
    }

    /// <summary>End-to-end via the public <see cref="DocxDiff"/> surface:
    /// <c>DocxDiffSettings.CrossParagraphTokenDiff</c> flows through <c>ToIrDiffSettings</c> to the same
    /// round-trippable markup, and the DATA surfaces (revisions / edit-script JSON) stay un-fused.</summary>
    [Fact]
    public void Public_DocxDiff_surfaces_with_cross_paragraph_setting()
    {
        var left = Docs.Para("HEAD", "alpha bravo charlie", "delta echo foxtrot", "TAIL");
        var right = Docs.Para("HEAD", "alpha bravo", "charlie delta echo foxtrot", "TAIL");
        var settings = new DocxDiffSettings { CrossParagraphTokenDiff = true };

        var rendered = DocxDiff.Compare(left, right, settings);
        Assert.Equal(Docs.PlainText(right), Docs.PlainText(RevisionProcessor.AcceptRevisions(rendered)));
        Assert.Equal(Docs.PlainText(left), Docs.PlainText(RevisionProcessor.RejectRevisions(rendered)));

        // Data paths force the fusion off: the JSON never carries the new kind, and GetRevisions works.
        Assert.DoesNotContain("CrossParagraphRunBlock", DocxDiff.GetEditScriptJson(left, right, settings));
        Assert.NotEmpty(DocxDiff.GetRevisions(left, right, settings));
    }

    /// <summary>
    /// Whitespace/punctuation attachment across the fused stream (decoded 2026-07-27 from reference
    /// compare output on the right-alignment demo pair). Three laws in one pin:
    /// (1) a separator flanking a matched pair on both sides pairs WITH it — the retained "text" keeps
    /// its trailing space in the ¶INS paragraph ("text " + ins "alignment.", not "text" + ins
    /// " alignment."), and "is"/"aligned" keep theirs;
    /// (2) the one-base-space / two-next-space shape resolves by trailing attachment: the ins run
    /// carries the SECOND next-side space ("…document " before retained "is");
    /// (3) a lone punctuation match with changed words on both sides and NO retained mark at the cell
    /// end is suppressed — the fused ¶DEL cell duplicates the period into both regions
    /// (ins "…margin." / del "and italic."), never a shared ".".
    /// </summary>
    [Fact]
    public void CrossParagraph_separator_attachment_matches_the_decoded_reference_shape()
    {
        var left = Docs.Para(
            "Right Aligned Italic Demo",
            "This text is right-aligned and italic.",
            "Right-aligned italic text creates an elegant signature effect.");
        var right = Docs.Para(
            "Right Alignment Demo",
            "This document demonstrates right text alignment.",
            "All text in this document is aligned to the right margin.");

        AssertFusedRoundTrip(left, right, expectedFusedOps: 1, "separator-attachment");

        var rendered = RenderMarkup(left, right, crossParagraph: true);
        var ns = (XNamespace)IrTestDocuments.W;
        var body = XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
        var paras = body.Elements(ns + "p").ToList();
        Assert.Equal(4, paras.Count);

        Assert.Equal("R[This ]I[document demonstrates right ]R[text ]I[alignment.]", StreamSig(paras[1], ns));
        Assert.Equal(
            "I[All text in this document ]R[is ]D[right-]R[aligned ]I[to the right margin.]D[and italic.]",
            StreamSig(paras[2], ns));
    }

    /// <summary>Per-paragraph run-stream signature: adjacent same-state text coalesced, states R (retained),
    /// I (w:ins), D (w:del) — the decode notation used against the reference outputs.</summary>
    private static string StreamSig(XElement para, XNamespace ns)
    {
        var parts = new List<(char State, string Text)>();
        void Walk(XElement el, char state)
        {
            foreach (var c in el.Elements())
            {
                if (c.Name == ns + "ins") Walk(c, 'I');
                else if (c.Name == ns + "del") Walk(c, 'D');
                else if (c.Name == ns + "pPr") { }
                else if (c.Name == ns + "t" || c.Name == ns + "delText")
                    parts.Add((state, c.Value));
                else Walk(c, state);
            }
        }
        Walk(para, 'R');

        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < parts.Count; i++)
        {
            if (i > 0 && parts[i].State == parts[i - 1].State)
            {
                sb.Length--;                      // reopen the previous bracket
                sb.Append(parts[i].Text).Append(']');
            }
            else
            {
                sb.Append(parts[i].State).Append('[').Append(parts[i].Text).Append(']');
            }
        }
        return sb.ToString();
    }

    /// <summary>Count CrossParagraphRunBlock ops ANYWHERE in a script — body, table cells, textbox interiors,
    /// and note/header-footer scopes — so a leak into a nested scope is caught.</summary>
    private static int CountFusedOps(IrEditScript script)
    {
        int n = script.Operations.Sum(CountFusedOps);
        if (script.NoteOps is { } notes)
            n += notes.SelectMany(nd => nd.Ops).Sum(CountFusedOps);
        if (script.HeaderFooterOps is { } hf)
            n += hf.SelectMany(h => h.Ops).Sum(CountFusedOps);
        return n;
    }

    private static int CountFusedOps(IrEditOp op)
    {
        int n = op.Kind == IrEditOpKind.CrossParagraphRunBlock ? 1 : 0;
        if (op.TextboxDiffs is { } tbx)
            n += tbx.SelectMany(t => t.Ops).Sum(CountFusedOps);
        if (op.TableDiff is { } td)
            n += td.RowOps
                .Where(r => r.CellOps is not null)
                .SelectMany(r => r.CellOps!)
                .Where(c => c.BlockOps is not null)
                .SelectMany(c => c.BlockOps!)
                .Sum(CountFusedOps);
        return n;
    }
}
