#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Pins for the aligner's same-slot positional pre-pairing, decoded from Word's compare output.
/// Word's replace-gap matcher is positional-first: it pairs the k-th unmatched base paragraph with
/// the k-th unmatched next paragraph whenever they share at least one non-function content word —
/// at word-overlap levels far below any fuzzy-similarity floor — and only then folds the genuinely
/// surplus paragraphs into adjacent output. Without the pass, the content-greedy machinery can pick
/// a one-slot-DISPLACED assignment (a tail pair or an added-text merge steals the k-th base
/// paragraph), which renders every later pair against the wrong counterpart.
/// </summary>
public class DocxDiffSameSlotPairingTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void ReplaceGap_PairsKthWithKth_OnSharedContentWord()
    {
        // Decoded from Word's compare output: four base paragraphs against three next paragraphs,
        // where each slot k shares at least one content word with its counterpart (slot 0:
        // "Showcase", slot 1: "ledger"/"krypton", slot 2: "prose") at word-Jaccard far below the
        // fuzzy-similarity floor. Word pairs slot 0/1/2 positionally and deletes the surplus base
        // paragraph. Without the same-slot pass, the added-text merge scan consumed base
        // paragraphs 1+2 for next paragraph 1 (wholesale-deleting base paragraph 1) and the residue
        // force-pair matched base 3 to next 2 — every pairing one slot displaced from Word's.
        var left = Doc(
            "Krypton Gauge Showcase",
            "This ledger outlines krypton gauge revisions.",
            "This prose employs a broader krypton gauge of 30vw.",
            "Krypton gauge sways legibility and ledger contour.");
        var right = Doc(
            "Teal Stout Prose Showcase",
            "This ledger mixes teal krypton hue with stout styling.",
            "Teal stout prose heralds upbeat results and triumph.");
        var result = DocxDiff.Compare(left, right);

        var paras = BodyParas(result);
        Assert.Equal(4, paras.Count);
        // Slot k retained text proves the k-th↔k-th assignment (a displaced assignment retains
        // nothing in the k-th output paragraph or deletes it wholesale).
        Assert.Contains("Showcase", RetainedText(paras[0]));
        Assert.Contains("ledger", RetainedText(paras[1]));
        Assert.Contains("prose", RetainedText(paras[2]));
        // The surplus base paragraph falls out as a deletion — never a displaced re-pairing.
        Assert.Contains("sways", DeletedText(paras[3]));
        Assert.DoesNotContain("sways", RetainedText(paras[3]));
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void ReplaceGap_FunctionWordOverlapOnly_DoesNotPair()
    {
        // Slot 0 shares only closed-class function words ("This", "with"); slot 1 shares nothing.
        // Stopword-grade overlap is positional scaffolding, not pairing evidence — Word keeps the
        // region as separate inserted and deleted paragraphs (mid-document arrangement).
        var left = Doc(
            "head anchor",
            "This chapter is written with care about many unrelated topics kept apart.",
            "Winter evenings stay quiet.",
            "tail anchor");
        var right = Doc(
            "head anchor",
            "This story goes with speed.",
            "Summer mornings ring loud.",
            "tail anchor");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "INS:INS", "DEL:DEL", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void ReplaceGap_SizeParityGuard_BlocksSharedWordPairAcrossSizes()
    {
        // One shared content word ("orbit") but a 1-vs-11 word-count ratio: on shared-word-only
        // evidence a paragraph does not pair with one several times its size (Word deletes the
        // short label and inserts the long sentence separately).
        var left = Doc("head anchor", "Orbit", "tail anchor");
        var right = Doc(
            "head anchor",
            "The orbit calculations for the seasonal launch window remain incomplete today.",
            "tail anchor");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void ReplaceGap_ShiftedDiagonal_CompetitorEvidenceWins()
    {
        // A leading insertion shifts the true correspondence off the slot diagonal: every list item
        // shares the boilerplate word "entry" with its slot counterpart, but the shifted counterpart
        // shares strictly more ("First"+"entry"). Word follows the stronger content diagonal —
        // the slot pass must yield (decoded from Word's compare output; the k-th↔k-th pairing on
        // "entry" alone rendered every item against the wrong counterpart).
        var left = Doc(
            "Register List Panel",
            "First entry",
            "Second entry",
            "Third entry",
            "Fourth entry");
        var right = Doc(
            "Register List Slanted Panel",
            "This ledger shows registered lists with slanted styling:",
            "First slanted registered entry",
            "Second slanted registered entry",
            "Third slanted registered entry");
        var result = DocxDiff.Compare(left, right);

        var paras = BodyParas(result);
        Assert.Equal(6, paras.Count);
        Assert.Contains("First", RetainedText(paras[2]));
        Assert.Contains("Second", RetainedText(paras[3]));
        Assert.Contains("Third", RetainedText(paras[4]));
        // Surplus "Fourth entry" joins the last pair at the story tail (deleted pilcrow on the
        // pair, spanning diff retains the shared "entry", story-final pilcrow lives).
        Assert.Equal("DEL", ParaMark(paras[4]));
        Assert.Null(ParaMark(paras[5]));
        Assert.Contains("Fourth", DeletedText(paras[5]));
        AssertRoundTrip(result, left, right);
    }

    // ---------------------------------------------------------------- story-tail surplus absorption

    [Fact]
    public void TailSurplusBase_JoinsAdjacentPair_DeletedPilcrowOnPair_SpanningDiff()
    {
        // Decoded from Word's compare output: the surplus base paragraph after the gap's LAST pair
        // at the story end does not keep its own deleted pilcrow — the PAIR paragraph's pilcrow
        // carries the deletion mark, the story-final pilcrow stays live, and the token diff spans
        // the joined text (the surplus's shared words are retained against the pair's counterpart).
        var left = Doc(
            "Krypton Gauge Showcase",
            "This ledger outlines krypton gauge revisions.",
            "This prose employs a broader krypton gauge of 30vw.",
            "Krypton gauge sways legibility and ledger contour.");
        var right = Doc(
            "Teal Stout Prose Showcase",
            "This ledger mixes teal krypton hue with stout styling.",
            "Teal stout prose heralds upbeat results and triumph.");
        var result = DocxDiff.Compare(left, right);

        var paras = BodyParas(result);
        Assert.Equal(4, paras.Count);
        Assert.Equal("DEL", ParaMark(paras[2]));   // the pair's pilcrow carries the deletion
        Assert.Null(ParaMark(paras[3]));           // the story-final pilcrow stays live
        Assert.Contains("sways", DeletedText(paras[3]));
        Assert.Contains("and", RetainedText(paras[3])); // spanning diff retains the shared word
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void TailSurplusNext_JoinsAdjacentPair_InsertedPilcrowOnPair_SpanningDiff()
    {
        // Mirror of the base-surplus law for a surplus NEXT paragraph: the pair paragraph's pilcrow
        // carries the insertion mark and the surplus's shared words are retained.
        var left = Doc(
            "Krypton Gauge 24 Showcase",
            "This ledger outlines krypton gauge 24.",
            "Broader krypton gauges lift legibility for briefings.");
        var right = Doc(
            "Krypton Gauge Showcase",
            "This ledger outlines krypton gauge revisions.",
            "This prose employs a broader krypton gauge of 30vw.",
            "Krypton gauge sways legibility and ledger contour.");
        var result = DocxDiff.Compare(left, right);

        var paras = BodyParas(result);
        Assert.Equal(4, paras.Count);
        Assert.Equal("INS", ParaMark(paras[2]));   // the pair's pilcrow carries the insertion
        Assert.Null(ParaMark(paras[3]));           // the story-final pilcrow stays live
        Assert.Contains("legibility", RetainedText(paras[3])); // spanning diff retains the shared word
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void MidStorySurplus_StaysSeparate_WithOwnDeletedPilcrow()
    {
        // The absorption is a story-tail law only (decoded: a deleted paragraph after a pair
        // mid-story keeps its own deleted pilcrow — the following retained anchor pins the shape).
        var left = Doc(
            "Krypton Gauge Showcase",
            "Krypton gauge sways legibility and ledger contour.",
            "closing anchor stays put");
        var right = Doc(
            "Teal Krypton Showcase",
            "closing anchor stays put");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "MIXED:-", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    // ---------------------------------------------------------------- similarity-pass evidence discipline

    [Fact]
    public void SimilarityPair_SeparatorSkeletonOverlap_DoesNotPair_BesideContentPair()
    {
        // Decoded from Word's compare output: two same-length prose sentences sharing ONLY a
        // closed-class function word ("This") clear the general similarity score purely on their
        // separator skeleton (spaces + final period are shared tokens). Beside a REAL content-word
        // pair (the "Demo" titles), Word flushes that scaffolding-grade correspondence and treats
        // the region as a full rewrite — separate inserted and deleted paragraphs, no interleave.
        var left = Doc(
            "head anchor",
            "Krypton Gauge Demo",
            "This text will be indented soon.",
            "Winter evenings stay quiet mostly.",
            "tail anchor");
        var right = Doc(
            "head anchor",
            "Stout Prose Demo",
            "This story reads with vigor today.",
            "Summer mornings ring loud always.",
            "tail anchor");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "MIXED:-", "INS:INS", "INS:INS", "DEL:DEL", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void SimilarityPair_FunctionGradeOverlap_StillPairs_WhenGapHasNoContentCorrespondence()
    {
        // The counterpart law: in a gap where NO candidate pairing carries content-word evidence,
        // whatever weak overlap exists is what Word's flat matcher pairs on (decoded: a lone shared
        // "and" is retained inside an otherwise zero-overlap rewrite). The similarity pass falls
        // back to its lexical-overlap grain there rather than dissolving the whole region.
        var left = Doc(
            "head anchor",
            "This text will be indented soon.",
            "Winter evenings stay quiet mostly.",
            "tail anchor");
        var right = Doc(
            "head anchor",
            "This story reads with vigor today.",
            "Summer mornings ring loud always.",
            "tail anchor");
        var result = DocxDiff.Compare(left, right);

        // The first sentence pair interleaves (shared "This" + separator skeleton clears the
        // score); the second stays below the similarity floor and falls out as delete + insert.
        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "MIXED:-", "INS:INS", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void SimilarityPair_FunctionWordPairSuppressed_TrailingGapKeepsWordArrangement()
    {
        // The corpus-decoded composite: slot 0 pairs on a real content word ("Demo"); the remaining
        // trailing region shares only function-word scaffolding, so no further pair forms and the
        // replace-gap arrangement applies — next-side surplus ¶INS first, the last next paragraph's
        // runs fused into the FIRST deleted paragraph (¶DEL), the interior deletion ¶DEL, and the
        // story-final base paragraph keeping the live pilcrow.
        var left = Doc(
            "Krypton Gauge Demo",
            "This text will be indented soon.",
            "Winter evenings stay quiet mostly.",
            "Autumn rains drift slowly northward.");
        var right = Doc(
            "Stout Prose Demo",
            "This story reads with vigor today.",
            "Summer mornings ring loud always.");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "MIXED:-", "INS:INS", "MIXED:DEL", "DEL:DEL", "DEL:-" }, shape);
        var paras = BodyParas(result);
        Assert.Contains("Demo", RetainedText(paras[0]));
        // The fused paragraph carries the LAST next paragraph's inserted runs before the FIRST
        // deleted base paragraph's runs.
        Assert.Contains("Summer mornings", InsertedText(paras[2]));
        Assert.Contains("indented", DeletedText(paras[2]));
        AssertRoundTrip(result, left, right);
    }

    // ---------------------------------------------------------------- blank-spacer anchoring

    [Fact]
    public void FormatDifferingBlanks_DoNotAnchor_AcrossUnrelatedContent()
    {
        // Decoded from Word's compare output: a blank paragraph's only identity IS its formatting,
        // so a format-DIFFERING blank is a different blank — not "the same paragraph, reformatted".
        // In-order pairing such blanks (the FormatOnly pass) manufactures mid-gap anchors between
        // unrelated regions and fragments what Word processes as ONE replace region; Word emits the
        // whole region as inserts followed by deletes, with the blanks marked like any other
        // surplus. Format-EQUAL blanks still anchor (the empty-anchoring pins elsewhere hold).
        var left = Doc(
            ("head anchor", null),
            ("alpha one", null),
            ("", "center"),
            ("beta two", null),
            ("tail anchor", null));
        var right = Doc(
            ("head anchor", null),
            ("gamma three", null),
            ("", null),
            ("delta four", null),
            ("tail anchor", null));
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(
            new[] { "RET", "INS:INS", "EMPTY_INS:INS", "INS:INS", "DEL:DEL", "EMPTY_DEL:DEL", "DEL:DEL", "RET" },
            shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void FormatEqualBlank_BetweenPairedContent_KeepsAnchoring()
    {
        var left = Doc("Shared alpha heading line.", "", "Shared beta body line.");
        var right = Doc("Shared alpha heading line.", "", "Shared beta body line revised.");

        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal("RET", shape[0]);
        Assert.Equal("EMPTY:-", shape[1]);
        AssertRoundTrip(result, left, right);
    }

    // ---------------------------------------------------------------- helpers

    private static WmlDocument Doc(params string[] paraTexts) =>
        Doc(paraTexts.Select(t => (t, (string?)null)).ToArray());

    private static WmlDocument Doc(params (string Text, string? Jc)[] paras)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            foreach (var (text, jc) in paras)
            {
                var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                if (jc is not null)
                {
                    p.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.Justification
                        {
                            Val = jc switch
                            {
                                "center" => DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center,
                                "right" => DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right,
                                _ => DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left,
                            },
                        });
                }
                if (text.Length > 0)
                    p.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text(text)));
                body.Append(p);
            }
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(body);
            doc.Save();
        }
        return new WmlDocument("same-slot.docx", stream.ToArray());
    }

    private static List<XElement> BodyParas(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(word.MainDocumentPart!.GetStream());
        var xdoc = XDocument.Parse(reader.ReadToEnd());
        return xdoc.Root!.Element(W + "body")!.Elements(W + "p").ToList();
    }

    private static string RetainedText(XElement p) => string.Concat(p.Descendants(W + "t")
        .Where(t => t.Ancestors(W + "ins").FirstOrDefault() is null)
        .Select(t => t.Value));

    private static string DeletedText(XElement p) =>
        string.Concat(p.Descendants(W + "delText").Select(t => t.Value));

    private static string InsertedText(XElement p) => string.Concat(p.Descendants(W + "ins")
        .SelectMany(i => i.Descendants(W + "t")).Select(t => t.Value));

    private static string? ParaMark(XElement p)
    {
        var rPr = p.Element(W + "pPr")?.Element(W + "rPr");
        if (rPr?.Element(W + "ins") is not null) return "INS";
        if (rPr?.Element(W + "del") is not null) return "DEL";
        return null;
    }

    /// <summary>Category:mark per body paragraph — INS/DEL/MIXED/EMPTY[_INS|_DEL]/RET, mark INS/DEL/-.</summary>
    private static string[] BodyShape(WmlDocument doc)
    {
        var shapes = new List<string>();
        foreach (var p in BodyParas(doc))
        {
            string insText = InsertedText(p);
            string delText = DeletedText(p);
            string retText = RetainedText(p);
            var mark = ParaMark(p) ?? "-";
            string cat;
            if (insText.Trim().Length > 0 && delText.Trim().Length > 0) cat = "MIXED";
            else if (insText.Trim().Length > 0) cat = "INS";
            else if (delText.Trim().Length > 0) cat = "DEL";
            else if (retText.Trim().Length > 0) cat = "RET";
            else cat = mark switch { "INS" => "EMPTY_INS", "DEL" => "EMPTY_DEL", _ => "EMPTY" };
            shapes.Add(cat == "RET" ? "RET" : $"{cat}:{mark}");
        }
        return shapes.ToArray();
    }

    private static void AssertRoundTrip(WmlDocument redline, WmlDocument left, WmlDocument right)
    {
        Assert.Equal(BodyTextOf(right), BodyTextOf(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(BodyTextOf(left), BodyTextOf(RevisionProcessor.RejectRevisions(redline)));
    }

    private static string BodyTextOf(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        return string.Join(" ", body
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
            .Select(p => string.Concat(p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text))));
    }
}
