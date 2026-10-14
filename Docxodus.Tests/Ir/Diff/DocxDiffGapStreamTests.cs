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
/// Story-final GAP-REGION token stream (2026-07-27) — the decoded extension of the cross-paragraph
/// word+pilcrow stream to regions the pair-run substrate never covered: (a) story-final replace
/// regions with ZERO word-matched pairs whose one-sided paragraphs still share words (function words
/// included — reference compare output retains "This"/"is"/"." across such gaps, materializing
/// ¶INS/¶DEL splits from the interleave), and (b) regions where one-sided members precede the
/// word-matched pair (the leading del/ins joins the same stream, so an early one-sided paragraph's
/// word can land retained inside a later paired output paragraph). Plus the decoded MERGE-SLOT law:
/// a story-tail base surplus merges at the FIRST pair whose next-side tail residue shares an
/// unmatched content word with the following base paragraph — not blindly at the last pair — and the
/// displaced pairs re-slot forward one base paragraph each.
/// All shapes keep the sacred round-trip contract: accept ≡ right, reject ≡ left (paragraph counts
/// included, via <see cref="Docs.PlainText"/>'s newline join).
/// </summary>
[Trait("Category", "Markup")]
public class DocxDiffGapStreamTests
{
    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    private static IrEditScript BuildScript(WmlDocument left, WmlDocument right)
    {
        var settings = new IrDiffSettings { CrossParagraphTokenDiff = true };
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        return IrEditScriptBuilder.Build(irLeft, irRight, settings);
    }

    private static WmlDocument RenderMarkup(WmlDocument left, WmlDocument right)
    {
        var settings = new IrDiffSettings { CrossParagraphTokenDiff = true };
        var script = BuildScript(left, right);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    private static void AssertRoundTrip(WmlDocument left, WmlDocument right, string label)
    {
        var rendered = RenderMarkup(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Assert.True(Docs.PlainText(right) == Docs.PlainText(accepted), $"{label}: accept ≢ right");
        Assert.True(Docs.PlainText(left) == Docs.PlainText(rejected), $"{label}: reject ≢ left");
    }

    private static XElement Body(WmlDocument rendered)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        return XDocument.Parse(Docs.MainPartXml(rendered)).Root!.Element(ns + "body")!;
    }

    private static void AssertRetainedOnce(XElement body, string word, string label)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        var runs = body.Descendants(ns + "t").Where(t => t.Value.Contains(word)).ToList();
        var dels = body.Descendants(ns + "delText").Where(t => t.Value.Contains(word)).ToList();
        Assert.True(runs.Count == 1, $"{label}: expected \"{word}\" in exactly one live run, got {runs.Count}");
        Assert.True(dels.Count == 0, $"{label}: expected \"{word}\" in no delText, got {dels.Count}");
        Assert.False(runs[0].Ancestors().Any(a => a.Name == ns + "ins" || a.Name == ns + "del"),
            $"{label}: \"{word}\" must render as a shared (unwrapped) run");
    }

    private static (int InsMarks, int DelMarks) PilcrowMarks(XElement body)
    {
        var ns = (XNamespace)IrTestDocuments.W;
        int ins = 0, del = 0;
        foreach (var p in body.Elements(ns + "p"))
        {
            var rpr = p.Element(ns + "pPr")?.Element(ns + "rPr");
            if (rpr is null)
                continue;
            if (rpr.Element(ns + "ins") is not null) ins++;
            if (rpr.Element(ns + "del") is not null) del++;
        }
        return (ins, del);
    }

    /// <summary>
    /// (a) ZERO-PAIR story-final region: no cross-side paragraph pair clears the aligner's pairing
    /// floor, yet the sides share words in the decoded crossing pattern (left p2's "vexed" ↔ right
    /// p1's head, left p2's "is" ↔ right p2). The stream deletes the unmatched left paragraph's
    /// pilcrow (¶DEL), inserts the right head paragraph's (¶INS mid-left-p2), retains "vexed" inside
    /// the ¶INS-terminated paragraph, and the story-final pilcrow pairs — four body paragraphs, the
    /// shape no per-paragraph or gap-grammar model reproduces.
    /// </summary>
    [Fact]
    public void GapStream_zero_pair_story_final_region_streams_shared_words()
    {
        var left = Docs.Para("HEAD",
            "gargle mumble tumble quince",
            "vexed zonk is trimmed in cyan boxes");
        var right = Docs.Para("HEAD",
            "vexed cursor shows granite marble with cobalt fizz styles",
            "granite fizz is standard for formal quartz sessions");

        var script = BuildScript(left, right);
        int fused = script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);
        Assert.True(fused == 1, $"zero-pair region: expected 1 CrossParagraphRunBlock, got {fused}");

        AssertRoundTrip(left, right, "zero-pair region");

        var body = Body(RenderMarkup(left, right));
        var ns = (XNamespace)IrTestDocuments.W;
        Assert.Equal(4, body.Elements(ns + "p").Count()); // HEAD + 3 stream paragraphs
        AssertRetainedOnce(body, "vexed", "zero-pair region");
        var (insMarks, delMarks) = PilcrowMarks(body);
        Assert.Equal(1, insMarks);
        Assert.Equal(1, delMarks);
    }

    /// <summary>
    /// (b) LEADING one-sided members join the stream: region = [deleted left p1, inserted right p1,
    /// word-matched pair (left p2 ↔ right p2)]. The deleted paragraph's "vexed" matches the paired
    /// right paragraph's head (a one-sided → paired forward flow), so it renders RETAINED inside the
    /// ¶DEL-terminated paragraph; accept fuses it into the pair's paragraph reproducing right p2
    /// exactly, reject reproduces left p1/p2.
    /// </summary>
    [Fact]
    public void GapStream_leading_one_sided_members_join_the_pair_stream()
    {
        var left = Docs.Para("HEAD",
            "vexed poodle combines azure quilt jackets",
            "ochre nimbus text resembles hyperlinks in scrolls");
        var right = Docs.Para("HEAD",
            "demonstrating waffle and iron combined",
            "vexed text is both waffle and iron simultaneously");

        var script = BuildScript(left, right);
        int fused = script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);
        Assert.True(fused == 1, $"leading one-sided: expected 1 CrossParagraphRunBlock, got {fused}");

        AssertRoundTrip(left, right, "leading one-sided");

        var body = Body(RenderMarkup(left, right));
        var ns = (XNamespace)IrTestDocuments.W;
        Assert.Equal(4, body.Elements(ns + "p").Count()); // HEAD + ¶INS + ¶DEL + shared final
        AssertRetainedOnce(body, "vexed", "leading one-sided");
        var (insMarks, delMarks) = PilcrowMarks(body);
        Assert.Equal(1, insMarks);
        Assert.Equal(1, delMarks);
    }

    /// <summary>A 1×1 story-final del+ins gap NEVER enters the region stream — the replace-gap
    /// grammar owns that shape (its fused-tail re-diff already retains shared tokens).</summary>
    [Fact]
    public void GapStream_1x1_story_final_gap_stays_with_the_gap_grammar()
    {
        var left = Docs.Para("HEAD", "alpha bravo shared quill");
        var right = Docs.Para("HEAD", "shared charlie delta echo foxtrot golf hotel india juliet");

        var script = BuildScript(left, right);
        Assert.Equal(0, script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock));
        AssertRoundTrip(left, right, "1x1 gap");
    }

    /// <summary>A zero-pair region whose sides share exactly ONE word does NOT stream — the decoded
    /// interleave boundary is ≥2 matched word units (the same threshold as the residue-pair law);
    /// a lone shared word renders via the replace-gap grammar with no retention.</summary>
    [Fact]
    public void GapStream_single_shared_word_region_falls_back()
    {
        var left = Docs.Para("HEAD",
            "gargle mumble tumble quince",
            "vexed floop dorble strom quibble");
        var right = Docs.Para("HEAD",
            "vexed cursor granite marble cobalt fizz wental",
            "standard formal quartz dune pipkin");

        var script = BuildScript(left, right);
        Assert.Equal(0, script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock));
        AssertRoundTrip(left, right, "single-shared-word gap");
    }

    /// <summary>A zero-pair story-final region with NO shared words falls back to the replace-gap
    /// grammar (the stream finds no matches; the structure gate declines).</summary>
    [Fact]
    public void GapStream_zero_shared_words_region_falls_back()
    {
        var left = Docs.Para("HEAD",
            "gargle mumble tumble quince",
            "zonk wibble frond peppercorn");
        var right = Docs.Para("HEAD",
            "cursor granite marble cobalt",
            "standard formal quartz sessions");

        var script = BuildScript(left, right);
        Assert.Equal(0, script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock));
        AssertRoundTrip(left, right, "no-shared-words gap");
    }

    /// <summary>
    /// COUNT-EQUAL BOUNDARY CONSTRUCT, story-final: a zero-pair region whose sides share exactly ONE
    /// word can still stream when the same number of matched tokens precedes a left and a right
    /// pilcrow (here: one match — "vexed" — precedes both left p3's and right p1's pilcrows), which
    /// pairs those pilcrows into a retained-¶ construct mid-region; the following inserted paragraph,
    /// the hoisted leading ins-words inside the ¶DEL cells, and the trailing deletions all emit from
    /// the same stream. Decoded from the reference compare output (the lone-shared-word NO-construct
    /// case stays with the grammar — see the single-shared-word pin).
    /// </summary>
    [Fact]
    public void GapStream_count_equal_boundary_construct_streams_single_match_region()
    {
        var left = Docs.Para("HEAD",
            "gargle mumble policy vexon",
            "wibble dates frong",
            "scope employees vexed contractors flim glorb",
            "password rekwire ments",
            "muffa reqired for all");
        var right = Docs.Para("HEAD",
            "italic vexed underline combo demo quix",
            "demonstrating flopping and prongs",
            "sample of flopped prongdom");

        var script = BuildScript(left, right);
        int fused = script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);
        Assert.True(fused == 1, $"count-equal construct: expected 1 CrossParagraphRunBlock, got {fused}");

        AssertRoundTrip(left, right, "count-equal construct");

        var body = Body(RenderMarkup(left, right));
        AssertRetainedOnce(body, "vexed", "count-equal construct");
    }

    /// <summary>
    /// INTERIOR region (not story-final): the same zero-pair stream fires MID-document when a
    /// count-equal boundary construct forms — the leading inserted paragraph, the hoist of the next
    /// paragraph's leading ins-word into the ¶DEL cell, the retained-¶ construct, and a TRAILING
    /// pure-insert cell, all before an unchanged anchor paragraph. Accept/reject reconstruct the
    /// sides exactly (the trailing ¶INS cell carries no left content, so nothing can fuse into the
    /// following anchor).
    /// </summary>
    [Fact]
    public void GapStream_interior_region_with_construct_streams()
    {
        var left = Docs.Para("HEAD",
            "xrayed mumble vexony",
            "a comprehensive flarn demonstration wexler quibble",
            "TAIL");
        var right = Docs.Para("HEAD",
            "double flopping bold demo",
            "this wexler demonstrates prong line flopping dorble",
            "bold flopped text sample",
            "TAIL");

        var script = BuildScript(left, right);
        int fused = script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);
        Assert.True(fused == 1, $"interior construct: expected 1 CrossParagraphRunBlock, got {fused}");

        AssertRoundTrip(left, right, "interior construct");

        var body = Body(RenderMarkup(left, right));
        var ns = (XNamespace)IrTestDocuments.W;
        Assert.Equal(6, body.Elements(ns + "p").Count()); // HEAD + 4 stream paragraphs + TAIL
        AssertRetainedOnce(body, "wexler", "interior construct");
    }

    /// <summary>
    /// SECTION-BREAK TRANSPARENCY + ONE-CONSTRUCT law: when the sides' section properties differ,
    /// the aligner interleaves a Deleted and an Inserted section break with the story-final replace
    /// region — the stream looks straight through them (their ops re-emit after the fused op). And a
    /// zero-pair region forms AT MOST ONE count-equal construct: matches past the first construct's
    /// boundary ("prongs", four paragraphs down) are suppressed — the region's remainder renders as
    /// the plain replace-gap arrangement (own-¶INS inserted paragraph, the last right's runs fused
    /// into the following ¶DEL cell, plain deletions, live story-final pilcrow).
    /// </summary>
    [Fact]
    public void GapStream_section_break_transparent_and_single_construct()
    {
        static WmlDocument BodyDoc(string pgSz, params string[] texts)
        {
            var paras = string.Concat(texts.Select(t => $"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>"));
            return IrTestDocuments.FromBodyXml(paras + $"<w:sectPr><w:pgSz w:w=\"{pgSz}\" w:h=\"15840\"/></w:sectPr>");
        }

        var left = BodyDoc("12240",
            "gargle mumble policy vexon",
            "wibble dates frong",
            "scope employees vexed contractors flim glorb",
            "password rekwire ments",
            "muffa reqired and quonk zibble wamp");
        var right = BodyDoc("11906",
            "italic vexed underline combo demo quix",
            "demonstrating flopping wexcel groves",
            "sample of flopped and prongdom");

        var script = BuildScript(left, right);
        int fused = script.Operations.Count(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock);
        Assert.True(fused == 1, $"sect-transparent construct: expected 1 CrossParagraphRunBlock, got {fused} " +
            $"(ops: [{string.Join(", ", script.Operations.Select(o => o.Kind))}])");

        AssertRoundTrip(left, right, "sect-transparent construct");

        var body = Body(RenderMarkup(left, right));
        AssertRetainedOnce(body, "vexed", "sect-transparent construct");
        // The post-construct shared word ("and", four paragraphs past the construct) is never
        // retained: the live occurrence stays inserted, the base occurrence stays deleted.
        var ns = (XNamespace)IrTestDocuments.W;
        var andLive = body.Descendants(ns + "t").Where(t => t.Value.Contains(" and")).ToList();
        foreach (var t in andLive)
            Assert.True(t.Ancestors().Any(a => a.Name == ns + "ins"),
                "post-construct match must not be retained (live occurrence outside w:ins)");
        Assert.True(body.Descendants(ns + "delText").Any(t => t.Value.Contains(" and ")),
            "post-construct match must not be retained (base occurrence must stay deleted)");
    }

    /// <summary>
    /// MERGE-SLOT law: with a story-tail base surplus, the 2:1 merge forms at the FIRST pair whose
    /// next-side TAIL residue (unmatched words after the pair's last matched word) shares an
    /// unmatched content word with the FOLLOWING base paragraph — here n1's tail "sigma tau hotel"
    /// shares "hotel" with b2 — so the merge is (b1,b2 → n1), the displaced pair re-slots to
    /// (b3 ↔ n2), and the story-final weak pair flushes to a whole ins+del. "hotel" renders RETAINED
    /// (it flows from n1's tail into merge member b2).
    /// </summary>
    [Fact]
    public void MergeSlot_tail_evidence_moves_merge_to_earlier_pair()
    {
        var left = Docs.Para("HEAD",
            "alpha bravo charlie delta echo golf omega",
            "vexed hotel uses kappa lambda golf",
            "yankee zulu november oscar");
        var right = Docs.Para("HEAD",
            "alpha bravo charlie golf sigma tau hotel",
            "papa quebec romeo golf whiskey");

        var script = BuildScript(left, right);
        var kinds = script.Operations.Select(o => o.Kind).ToList();
        int mergeIdx = kinds.IndexOf(IrEditOpKind.MergeBlock);
        Assert.True(mergeIdx == 1, // Equal(HEAD) first, then the merge — NOT a leading 1:1 pair
            $"tail-evidence merge slot: expected MergeBlock at op index 1, got kinds [{string.Join(", ", kinds)}]");

        AssertRoundTrip(left, right, "tail-evidence merge slot");

        var body = Body(RenderMarkup(left, right));
        AssertRetainedOnce(body, "hotel", "tail-evidence merge slot");
    }

    /// <summary>Without tail-residue evidence the story-tail surplus keeps its decoded default: the
    /// merge forms at the LAST pair (surplus absorbed there), the leading pair stays 1:1.</summary>
    [Fact]
    public void MergeSlot_without_evidence_keeps_story_tail_surplus_at_last_pair()
    {
        var left = Docs.Para("HEAD",
            "alpha bravo charlie delta echo golf omega",
            "vexed uses kappa lambda golf",
            "yankee zulu november oscar");
        var right = Docs.Para("HEAD",
            "alpha bravo charlie golf sigma tau",
            "papa quebec romeo golf whiskey");

        var script = BuildScript(left, right);
        var kinds = script.Operations.Select(o => o.Kind).ToList();
        int mergeIdx = kinds.IndexOf(IrEditOpKind.MergeBlock);
        Assert.True(mergeIdx == 2, // Equal(HEAD), the 1:1 pair, then the surplus merge at the last pair
            $"no-evidence merge slot: expected MergeBlock at op index 2, got kinds [{string.Join(", ", kinds)}]");
        AssertRoundTrip(left, right, "no-evidence merge slot");
    }
}
