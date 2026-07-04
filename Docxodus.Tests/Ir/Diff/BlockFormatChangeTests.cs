#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Phase 0 characterization pins for the block-format-change family
/// (spec: docs/superpowers/specs/2026-07-03-diff-block-format-changes-design.md §1).
/// Each region pins what pairwise Compare/GetRevisions does TODAY for a property-only change with
/// identical text, at the soundness level (is the change visible? tracked? does accept ≡ right and
/// reject ≡ left hold at the property level?). As each implementation phase lands, the corresponding
/// pins are FLIPPED in place to the new tracked behavior — never deleted.
/// </summary>
public class BlockFormatChangeTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly DocxDiffSettings ModeledOnly = new();
    private static readonly DocxDiffSettings Full = new() { FormatComparison = DocxDiffFormatComparison.Full };

    private static XElement BodyOf(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        using var s = wDoc.MainDocumentPart!.GetStream();
        return XElement.Load(s).Element(W + "body")!;
    }

    // ------------------------------------------------------------------ pPr-only (w:jc)

    private static readonly WmlDocument PPrLeft = IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Same text here.</w:t></w:r></w:p>");

    private static readonly WmlDocument PPrRight = IrTestDocuments.FromBodyXml(
        "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Same text here.</w:t></w:r></w:p>");

    [Fact]
    public void PublicTrackBlockFormatChanges_false_suppresses_pPrChange()
    {
        // The public opt-out (DocxDiffSettings.TrackBlockFormatChanges = false) reaches the engine and
        // restores the untracked-right-apply behavior end to end.
        var result = DocxDiff.Compare(PPrLeft, PPrRight,
            new DocxDiffSettings { TrackBlockFormatChanges = false });
        Assert.Empty(BodyOf(result).Descendants(W + "pPrChange"));
        Assert.Empty(DocxDiff.GetRevisions(PPrLeft, PPrRight,
            new DocxDiffSettings { TrackBlockFormatChanges = false }));
    }

    [Fact]
    public void PPrOnly_change_is_tracked_with_native_pPrChange()
    {
        // Phase 1 (flipped Phase 0 pin): a modeled pPr-only change is tracked under BOTH policies.
        foreach (var settings in new[] { ModeledOnly, Full })
        {
            var result = DocxDiff.Compare(PPrLeft, PPrRight, settings);
            var body = BodyOf(result);

            var pPrChange = body.Descendants(W + "pPrChange").Single();
            Assert.Same(pPrChange, pPrChange.Parent!.Elements().Last());                  // last child of pPr
            Assert.Equal(settings.AuthorForRevisions, (string?)pPrChange.Attribute(W + "author"));
            Assert.NotNull(pPrChange.Attribute(W + "id"));
            Assert.NotNull(pPrChange.Attribute(W + "date"));
            var inner = pPrChange.Element(W + "pPr")!;
            Assert.Null(inner.Element(W + "jc"));                                          // old = no jc
            Assert.Empty(inner.Elements(W + "rPr"));                                       // CT_PPrBase: no mark rPr
            Assert.Empty(inner.Elements(W + "sectPr"));                                    //   and no sectPr

            // Output carries the RIGHT pPr (accepted state)…
            Assert.Equal("center",
                (string?)pPrChange.Parent!.Element(W + "jc")?.Attribute(W + "val"));

            // …accept ≡ right, and reject ≡ LEFT at the pPr level (the flipped soundness pin).
            var accepted = RevisionProcessor.AcceptRevisions(result);
            Assert.Equal("center", (string?)BodyOf(accepted).Descendants(W + "jc").Single().Attribute(W + "val"));
            Assert.Empty(BodyOf(accepted).Descendants(W + "pPrChange"));
            var rejected = RevisionProcessor.RejectRevisions(result);
            Assert.Empty(BodyOf(rejected).Descendants(W + "jc"));
            Assert.Empty(BodyOf(rejected).Descendants(W + "pPrChange"));
        }
    }

    [Fact]
    public void PPrOnly_Full_reports_a_FormatChanged_revision()
    {
        var revisions = DocxDiff.GetRevisions(PPrLeft, PPrRight, Full);
        var rev = Assert.Single(revisions);
        Assert.Equal(DocxDiffRevisionType.FormatChanged, rev.Type);
        Assert.NotNull(rev.FormatChange);
    }

    [Fact]
    public void PPrOnly_reports_paragraph_scope_FormatChanged_details()
    {
        var revisions = DocxDiff.GetRevisions(PPrLeft, PPrRight, ModeledOnly);
        var rev = Assert.Single(revisions);
        Assert.Equal(DocxDiffRevisionType.FormatChanged, rev.Type);
        var fc = rev.FormatChange!;
        Assert.Equal(DocxDiffFormatChangeScope.Paragraph, fc.Scope);
        Assert.Equal(new[] { "justification" }, fc.ChangedPropertyNames);
        Assert.Equal("Center", fc.NewProperties["justification"]);
        Assert.False(fc.OldProperties.ContainsKey("justification"));
        Assert.NotNull(rev.LeftAnchor);
        Assert.NotNull(rev.RightAnchor);
    }

    [Fact]
    public void NumberingOnly_change_reports_numId_in_details()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Item text.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Item text.</w:t></w:r></w:p>");
        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(new[] { "numId" }, rev.FormatChange!.ChangedPropertyNames);
        Assert.Equal("1", rev.FormatChange.OldProperties["numId"]);
        Assert.Equal("2", rev.FormatChange.NewProperties["numId"]);
    }

    [Fact]
    public void CompatibleMode_excludes_paragraph_scope_revisions()
    {
        // By-construction exclusion (the hdr/ftr precedent): the compatible granularity is defined as the
        // ORACLE's revision set, and WmlComparer cannot produce block-scope format revisions.
        var settings = new DocxDiffSettings { RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible };
        Assert.Empty(DocxDiff.GetRevisions(PPrLeft, PPrRight, settings));
    }

    [Fact]
    public void RunLevel_FormatChanged_keeps_Run_scope()
    {
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Same text here.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Same text here.</w:t></w:r></w:p>");
        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(DocxDiffFormatChangeScope.Run, rev.FormatChange!.Scope);
        Assert.Contains("bold", rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void TextAndPPr_modify_reports_paragraph_scope_alongside_token_revisions()
    {
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Old words here now.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>New words here now.</w:t></w:r></w:p>");
        var revisions = DocxDiff.GetRevisions(left, right, ModeledOnly);
        var para = Assert.Single(revisions,
            r => r.FormatChange is { } fc && fc.Scope == DocxDiffFormatChangeScope.Paragraph);
        Assert.Equal(new[] { "justification" }, para.FormatChange!.ChangedPropertyNames);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted);
    }

    [Fact]
    public void PPrChange_on_right_paragraph_without_pPr()
    {
        // Right paragraph LOST its pPr (left was centered): a fresh pPr holds only the pPrChange, whose
        // inner carries the OLD (left) jc; reject restores the centering.
        var result = DocxDiff.Compare(PPrRight, PPrLeft, ModeledOnly);
        var body = BodyOf(result);
        var pPrChange = body.Descendants(W + "pPrChange").Single();
        Assert.Equal("center", (string?)pPrChange.Element(W + "pPr")!.Element(W + "jc")?.Attribute(W + "val"));
        Assert.Null(pPrChange.Parent!.Element(W + "jc"));                                  // accepted state: no jc

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("center", (string?)BodyOf(rejected).Descendants(W + "jc").Single().Attribute(W + "val"));
    }

    [Fact]
    public void TextAndPPr_modify_block_also_tracks_pPrChange()
    {
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Old words here now.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>New words here now.</w:t></w:r></w:p>");

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        Assert.Single(body.Descendants(W + "pPrChange"));
        Assert.NotEmpty(body.Descendants(W + "ins"));                                      // the text edit is there too

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Empty(BodyOf(rejected).Descendants(W + "jc"));                              // left pPr restored
        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal("center", (string?)BodyOf(accepted).Descendants(W + "jc").Single().Attribute(W + "val"));
    }

    [Fact]
    public void MarkRPr_change_is_tracked_via_pPr_rPr_rPrChange()
    {
        // Only the paragraph-MARK rPr differs (bold pilcrow) — schema puts this OUTSIDE pPrChange:
        // it tracks as w:pPr/w:rPr/w:rPrChange. Detected under Full (mark rPr rides the unmodeled digest).
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Same text here.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Same text here.</w:t></w:r></w:p>");

        var result = DocxDiff.Compare(left, right, Full);
        var body = BodyOf(result);
        Assert.Empty(body.Descendants(W + "pPrChange"));                                   // no property child changed
        var markChange = body.Descendants(W + "pPr").Elements(W + "rPr").Elements(W + "rPrChange").Single();
        Assert.Empty(markChange.Element(W + "rPr")!.Elements());                           // old mark = empty

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Empty(BodyOf(rejected).Descendants(W + "pPr").Elements(W + "rPr").Elements(W + "b"));
        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Single(BodyOf(accepted).Descendants(W + "pPr").Elements(W + "rPr").Elements(W + "b"));
    }

    [Fact]
    public void UnmodeledOnly_pPr_delta_stays_untracked_under_ModeledOnly()
    {
        // The documented blind spot survives Phase 1: paragraph shading is unmodeled → no markup, right-apply.
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/></w:pPr><w:r><w:t>Same text here.</w:t></w:r></w:p>");
        var result = DocxDiff.Compare(PPrLeft, right, ModeledOnly);
        Assert.Empty(BodyOf(result).Descendants(W + "pPrChange"));
        Assert.Single(BodyOf(result).Descendants(W + "shd"));
    }

    // ------------------------------------------------------------------ note / header pPrChange (follow-up A4)

    private static XElement PartXml(WmlDocument doc, System.Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var part = pick(wd.MainDocumentPart!)!;
        using var s = part.GetStream();
        return XElement.Load(s);
    }

    [Fact]
    public void Footnote_paragraph_pPr_change_is_tracked_with_pPrChange()
    {
        // A4: a changed footnote-definition paragraph already emits w:pPrChange (the note scope routes through
        // the same RenderBlockOp dispatch as the body) — proven here, not a v1 ceiling.
        const string body = "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(body, "<w:r><w:t>Footnote text.</w:t></w:r>");
        var right = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(body, "<w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Footnote text.</w:t></w:r>");

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var fnXml = PartXml(result, m => m.FootnotesPart);
        Assert.Single(fnXml.Descendants(W + "pPrChange"));

        // Round-trips inside the footnotes part.
        Assert.Single(PartXml(RevisionProcessor.AcceptRevisions(result), m => m.FootnotesPart).Descendants(W + "jc"));
        Assert.Empty(PartXml(RevisionProcessor.RejectRevisions(result), m => m.FootnotesPart).Descendants(W + "jc"));

        // Fine revisions report the Paragraph-scope change with a footnote anchor.
        Assert.Contains(DocxDiff.GetRevisions(left, right, ModeledOnly), r =>
            r.FormatChange is { } fc && fc.Scope == DocxDiffFormatChangeScope.Paragraph
            && (r.LeftAnchor?.Contains(":fn") ?? false));
    }

    [Fact]
    public void Header_paragraph_pPr_change_is_tracked_with_pPrChange()
    {
        // A4: a changed header-story paragraph already emits w:pPrChange (same shared dispatch).
        const string body = "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyAndHeaderXml(body, "<w:p><w:r><w:t>HEADER</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyAndHeaderXml(body, "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>HEADER</w:t></w:r></w:p>");

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var hdrXml = PartXml(result, m => m.HeaderParts.First());
        Assert.Single(hdrXml.Descendants(W + "pPrChange"));
        Assert.Single(PartXml(RevisionProcessor.AcceptRevisions(result), m => m.HeaderParts.First()).Descendants(W + "jc"));
        Assert.Empty(PartXml(RevisionProcessor.RejectRevisions(result), m => m.HeaderParts.First()).Descendants(W + "jc"));
    }

    [Fact]
    public void Split_members_do_not_emit_pPrChange_declined_v1()
    {
        // A4 / D1: split (and merge) members are NOT eligible for w:pPrChange — a split's members are
        // brand-new paragraphs already tracked by the inserted pilcrow mark; a pPr "change" against the
        // single left paragraph is not well-defined and would fight the reject-fuse. This pins the decline:
        // whatever the aligner classifies the edit as (split or del/ins), NO pPrChange is emitted, and it
        // still round-trips.
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>The quick brown fox jumps over the lazy dog every single day.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>The quick brown fox jumps</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>over the lazy dog every single day.</w:t></w:r></w:p>");

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        Assert.Empty(BodyOf(result).Descendants(W + "pPrChange"));                          // the decline

        // Content round-trips: accept ≡ right two paragraphs, reject ≡ left one paragraph.
        static string[] Texts(WmlDocument d) => BodyOf(d).Elements(W + "p")
            .Select(p => string.Concat(p.Descendants(W + "t").Select(t => (string?)t)))
            .Where(s => s.Length > 0).ToArray();
        Assert.Equal(Texts(right), Texts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(Texts(left), Texts(RevisionProcessor.RejectRevisions(result)));
    }

    // ------------------------------------------------------------------ Consolidate pPr merge (sub-project B)

    [Fact]
    public void PPrDigest_distinguishes_paragraph_properties_but_ignores_mark_sect_markers()
    {
        var opts = new IrReaderOptions { RetainSources = false };
        IrParagraph P(string pInner) => (IrParagraph)IrReader.Read(
            IrTestDocuments.FromBodyXml($"<w:p>{pInner}<w:r><w:t>Text.</w:t></w:r></w:p>"), opts).Body.Blocks[0];

        var plain = P("");
        var centered = P("<w:pPr><w:jc w:val=\"center\"/></w:pPr>");
        var centered2 = P("<w:pPr><w:jc w:val=\"center\"/></w:pPr>");
        var centeredBoldMark = P("<w:pPr><w:jc w:val=\"center\"/><w:rPr><w:b/></w:rPr></w:pPr>");

        Assert.NotEqual(plain.PPrDigest, centered.PPrDigest);          // a jc change is visible
        Assert.Equal(centered.PPrDigest, centered2.PPrDigest);         // identical pPr → identical digest
        Assert.Equal(centered.PPrDigest, centeredBoldMark.PPrDigest);  // the mark rPr is NOT part of the pPr digest
    }

    // ------------------------------------------------------------------ mid-doc inline sectPr (follow-up A3)

    private const string InlineSectBody =
        "<w:p><w:pPr><w:sectPr>{S}</w:sectPr></w:pPr><w:r><w:t>Section body.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>Next.</w:t></w:r></w:p>";

    [Fact]
    public void MidDoc_inline_sectPr_change_is_tracked_with_native_sectPrChange()
    {
        var left = IrTestDocuments.FromBodyXml(InlineSectBody.Replace("{S}", "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>"));
        var right = IrTestDocuments.FromBodyXml(InlineSectBody.Replace("{S}", "<w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var inlineSect = BodyOf(result).Elements(W + "p").First().Element(W + "pPr")!.Element(W + "sectPr")!;
        var change = inlineSect.Element(W + "sectPrChange")!;
        Assert.Same(change, inlineSect.Elements().Last());                                  // last child of the inline sectPr
        Assert.Equal("12240", (string?)change.Element(W + "sectPr")!.Element(W + "pgSz")?.Attribute(W + "w")); // old
        Assert.Equal("15840", (string?)inlineSect.Element(W + "pgSz")?.Attribute(W + "w"));                    // right applied

        Assert.Equal("15840", InlinePgW(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal("12240", InlinePgW(RevisionProcessor.RejectRevisions(result)));       // reject ≡ left

        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly),
            r => r.FormatChange is { } fc && fc.Scope == DocxDiffFormatChangeScope.Section);
        Assert.Contains("pageWidth", rev.FormatChange!.ChangedPropertyNames);

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(wd)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
            .Select(e => e.Description));
    }

    private static string? InlinePgW(WmlDocument doc) =>
        (string?)BodyOf(doc).Elements(W + "p").First().Element(W + "pPr")?.Element(W + "sectPr")?.Element(W + "pgSz")?.Attribute(W + "w");

    [Fact]
    public void Paragraph_with_both_pPr_and_inline_sectPr_change_is_schema_valid_and_round_trips()
    {
        // Review Note C: exercise the BOTH-fire path (a modeled pPr change AND an inline sectPr change on the
        // same paragraph) — the pPrChange (last pPr child) must sit AFTER the sectPr, and both round-trip.
        const string body = "<w:p><w:pPr>{P}<w:sectPr>{S}</w:sectPr></w:pPr><w:r><w:t>Section body.</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>Next.</w:t></w:r></w:p>";
        var left = IrTestDocuments.FromBodyXml(body.Replace("{P}", "").Replace("{S}", "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>"));
        var right = IrTestDocuments.FromBodyXml(body.Replace("{P}", "<w:jc w:val=\"center\"/>").Replace("{S}", "<w:pgSz w:w=\"15840\" w:h=\"12240\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var pPr = BodyOf(result).Elements(W + "p").First().Element(W + "pPr")!;
        Assert.Single(pPr.Elements(W + "pPrChange"));                                       // pPr change
        Assert.Single(pPr.Element(W + "sectPr")!.Elements(W + "sectPrChange"));             // inline sect change
        // Schema order: sectPr must precede pPrChange (both present).
        var kids = pPr.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.True(kids.IndexOf("sectPr") < kids.IndexOf("pPrChange"),
            $"sectPr must precede pPrChange; order was [{string.Join(",", kids)}]");

        using (var ms = new MemoryStream(result.DocumentByteArray))
        using (var wd = WordprocessingDocument.Open(ms, false))
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(wd)
                .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
                .Select(e => e.Description));

        // Both changes round-trip: reject restores left jc-absence AND left page size.
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Empty(BodyOf(rejected).Elements(W + "p").First().Element(W + "pPr")!.Elements(W + "jc"));
        Assert.Equal("12240", InlinePgW(rejected));
        Assert.Equal("15840", InlinePgW(RevisionProcessor.AcceptRevisions(result)));
    }

    [Fact]
    public void Fresh_tblPrEx_is_placed_before_an_existing_trPr()
    {
        // Review Note C: the removal-direction of the schema-order guard — a fresh w:tblPrEx (right row had
        // none) must land BEFORE an existing w:trPr (CT_Row: tblPrEx, then trPr).
        var left = IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"400\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"double\" w:sz=\"8\"/></w:tblBorders></w:tblPrEx><w:trPr><w:trHeight w:val=\"400\"/></w:trPr>",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var tr = BodyOf(result).Descendants(W + "tr").First();
        var kids = tr.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.True(kids.IndexOf("tblPrEx") >= 0 && kids.IndexOf("tblPrEx") < kids.IndexOf("trPr"),
            $"tblPrEx must precede trPr; order was [{string.Join(",", kids)}]");
        Assert.Single(tr.Descendants(W + "tblPrExChange"));

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(wd)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
            .Select(e => e.Description));
    }

    [Fact]
    public void Inline_sectPr_props_participate_in_paragraph_fingerprint()
    {
        var opts = new IrReaderOptions { RetainSources = false };
        var lp = (IrParagraph)IrReader.Read(IrTestDocuments.FromBodyXml(
            InlineSectBody.Replace("{S}", "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>")), opts).Body.Blocks[0];
        var rp = (IrParagraph)IrReader.Read(IrTestDocuments.FromBodyXml(
            InlineSectBody.Replace("{S}", "<w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/>")), opts).Body.Blocks[0];
        Assert.Equal(lp.ContentHash, rp.ContentHash);                  // text identical
        Assert.NotEqual(lp.FormatFingerprint, rp.FormatFingerprint);   // inline sectPr delta is now visible
        Assert.NotNull(lp.InlineSectionFormat);
    }

    // ------------------------------------------------------------------ tblPrEx (follow-up A2)

    [Fact]
    public void TblPrEx_change_flips_TrPrExDigest_only()
    {
        var left = IrReader.Read(IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPrEx>",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>")), new IrReaderOptions { RetainSources = false });
        var right = IrReader.Read(IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"double\" w:sz=\"8\"/></w:tblBorders></w:tblPrEx>",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>")), new IrReaderOptions { RetainSources = false });
        var lr = ((IrTable)left.Body.Blocks[0]).Rows[0];
        var rr = ((IrTable)right.Body.Blocks[0]).Rows[0];
        Assert.NotEqual(lr.TrPrExDigest, rr.TrPrExDigest);          // tblPrEx delta visible on its own digest
        Assert.Equal(lr.TrPrShellDigest, rr.TrPrShellDigest);       // trPr-only digest unchanged
    }

    // ------------------------------------------------------------------ table family (Phase 2)

    private static string Table(string trPr, string tcPr, string tblPr = "<w:tblW w:w=\"0\" w:type=\"auto\"/>",
                                string grid = "<w:gridCol w:w=\"4000\"/>") =>
        $"<w:tbl><w:tblPr>{tblPr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>" +
        $"<w:tr>{trPr}<w:tc><w:tcPr>{tcPr}</w:tcPr><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><w:t>After.</w:t></w:r></w:p>";

    [Fact]
    public void TrPrOnly_change_is_tracked_with_native_trPrChange()
    {
        var left = IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"400\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"800\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        var trPrChange = body.Descendants(W + "trPrChange").Single();
        Assert.Same(trPrChange, trPrChange.Parent!.Elements().Last());                     // last child of trPr
        Assert.NotNull(trPrChange.Attribute(W + "author"));
        Assert.Equal("400", (string?)trPrChange.Element(W + "trPr")!.Element(W + "trHeight")?.Attribute(W + "val"));
        Assert.Equal("800", (string?)trPrChange.Parent!.Element(W + "trHeight")?.Attribute(W + "val")); // right applied

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal("800", (string?)BodyOf(accepted).Descendants(W + "trHeight").Single().Attribute(W + "val"));
        Assert.Empty(BodyOf(accepted).Descendants(W + "trPrChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("400", (string?)BodyOf(rejected).Descendants(W + "trHeight").Single().Attribute(W + "val")); // reject ≡ left

        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(DocxDiffRevisionType.FormatChanged, rev.Type);
        Assert.Equal(DocxDiffFormatChangeScope.TableRow, rev.FormatChange!.Scope);
        Assert.Equal(new[] { "shell" }, rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void TblPrOnly_change_is_tracked_with_native_tblPrChange()
    {
        var borders = "<w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders>";
        var left = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>", tblPr: borders));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        var tblPrChange = body.Descendants(W + "tblPrChange").Single();
        Assert.Same(tblPrChange, tblPrChange.Parent!.Elements().Last());                   // last child of tblPr
        Assert.Empty(tblPrChange.Element(W + "tblPr")!.Elements(W + "tblBorders"));        // old = no borders
        Assert.Single(body.Descendants(W + "tblBorders"));                                 // right applied

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Single(BodyOf(accepted).Descendants(W + "tblBorders"));
        Assert.Empty(BodyOf(accepted).Descendants(W + "tblPrChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Empty(BodyOf(rejected).Descendants(W + "tblBorders"));                      // reject ≡ left

        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(DocxDiffFormatChangeScope.Table, rev.FormatChange!.Scope);
        Assert.Equal(new[] { "shell" }, rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void TblGridOnly_change_is_tracked_with_native_tblGridChange()
    {
        var left = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>", grid: "<w:gridCol w:w=\"6000\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        var gridChange = body.Descendants(W + "tblGridChange").Single();
        Assert.Equal(W + "tblGrid", gridChange.Parent!.Name);                              // inside the grid
        Assert.NotNull(gridChange.Attribute(W + "id"));
        Assert.Null(gridChange.Attribute(W + "author"));                                   // CT_Markup: id only
        Assert.Equal("4000", (string?)gridChange.Element(W + "tblGrid")!.Element(W + "gridCol")?.Attribute(W + "w"));
        // The applied (right) grid col is 6000; the OLD grid rides only inside the change marker.
        Assert.Equal("6000", (string?)gridChange.Parent!.Elements(W + "gridCol").Single().Attribute(W + "w"));

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal("6000", (string?)BodyOf(accepted).Descendants(W + "gridCol").Single().Attribute(W + "w"));
        Assert.Empty(BodyOf(accepted).Descendants(W + "tblGridChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("4000", (string?)BodyOf(rejected).Descendants(W + "gridCol").Single().Attribute(W + "w"));

        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(DocxDiffFormatChangeScope.Table, rev.FormatChange!.Scope);
        Assert.Equal(new[] { "grid" }, rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void TcPrOnly_change_is_tracked_with_native_tcPrChange()
    {
        var left = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table("",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/>"));

        // The shell digest participates in cell ContentHash → the table pair is Modified; the right tcPr
        // is applied WITH a tcPrChange carrying the old (left) shell — closing the #250-noted gap.
        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        var tcPrChange = body.Descendants(W + "tcPrChange").Single();
        Assert.Same(tcPrChange, tcPrChange.Parent!.Elements().Last());                     // last child of tcPr
        Assert.Empty(tcPrChange.Element(W + "tcPr")!.Elements(W + "shd"));                 // old = no shading
        Assert.Single(body.Descendants(W + "tc").Elements(W + "tcPr").Elements(W + "shd"));

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Single(BodyOf(accepted).Descendants(W + "tc").Elements(W + "tcPr").Elements(W + "shd"));
        Assert.Empty(BodyOf(accepted).Descendants(W + "tcPrChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Empty(BodyOf(rejected).Descendants(W + "tc").Elements(W + "tcPr").Elements(W + "shd")); // reject ≡ left

        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(DocxDiffFormatChangeScope.TableCell, rev.FormatChange!.Scope);
        Assert.Equal(new[] { "shell" }, rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void TableFamily_and_pPr_outputs_are_schema_valid()
    {
        var pairs = new (WmlDocument Left, WmlDocument Right)[]
        {
            (PPrLeft, PPrRight),
            (IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"400\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>")),
             IrTestDocuments.FromBodyXml(Table("<w:trPr><w:trHeight w:val=\"800\"/></w:trPr>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"))),
            (IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>")),
             IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/>", tblPr: "<w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders>", grid: "<w:gridCol w:w=\"6000\"/>"))),
        };
        foreach (var (left, right) in pairs)
        {
            var result = DocxDiff.Compare(left, right, ModeledOnly);
            using var ms = new MemoryStream(result.DocumentByteArray);
            using var wd = WordprocessingDocument.Open(ms, false);
            var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator()
                .Validate(wd)
                .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
                .Select(e => e.Description)
                .ToList();
            Assert.True(errors.Count == 0, string.Join("\n", errors));
        }
    }

    [Fact]
    public void Consolidate_ignores_table_shell_changes_v1_ceiling()
    {
        // Review finding 1: the table-shell REVISION emitters must respect the Consolidate ceiling too — a
        // reviewer's tcPr/trPr/tblPr-only edit produces neither markup nor a consolidated revision.
        var baseDoc = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var reviewerDoc = IrTestDocuments.FromBodyXml(Table("", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>",
            grid: "<w:gridCol w:w=\"6000\"/>", tblPr: "<w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders>"));
        var reviewer = new DocxDiffReviewer { Document = reviewerDoc, Author = "Reviewer A" };

        var merged = DocxDiff.Consolidate(baseDoc, new[] { reviewer });
        var body = BodyOf(merged);
        Assert.Empty(body.Descendants(W + "tblPrChange"));
        Assert.Empty(body.Descendants(W + "tblGridChange"));

        var revs = DocxDiff.GetConsolidatedRevisions(baseDoc, new[] { reviewer });
        Assert.DoesNotContain(revs, r => r.FormatChange is { } fc && fc.Scope != DocxDiffFormatChangeScope.Run);
    }

    [Fact]
    public void TblPrEx_change_is_tracked_with_native_tblPrExChange()
    {
        // Follow-up A2 (flips the former "untracked v1 ceiling" pin): a w:tblPrEx-only row change is now
        // tracked with native w:tblPrExChange — NOT via w:trPrChange (that stays trPr-only), and a distinct
        // TableRow revision with the "tblPrEx" changed-name.
        var left = IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPrEx>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"double\" w:sz=\"8\"/></w:tblBorders></w:tblPrEx>", "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        var change = body.Descendants(W + "tblPrExChange").Single();
        Assert.Same(change, change.Parent!.Elements().Last());                              // last child of tblPrEx
        Assert.Empty(body.Descendants(W + "trPrChange"));                                   // NOT a trPr change
        Assert.Equal("single",
            (string?)change.Element(W + "tblPrEx")!.Element(W + "tblBorders")!.Element(W + "top")?.Attribute(W + "val")); // old (left)
        Assert.Equal("double",
            (string?)change.Parent!.Element(W + "tblBorders")!.Element(W + "top")?.Attribute(W + "val"));                 // right applied

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal("double", (string?)BodyOf(accepted).Descendants(W + "tblPrEx").Single()
            .Element(W + "tblBorders")!.Element(W + "top")?.Attribute(W + "val"));
        Assert.Empty(BodyOf(accepted).Descendants(W + "tblPrExChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("single", (string?)BodyOf(rejected).Descendants(W + "tblPrEx").Single()
            .Element(W + "tblBorders")!.Element(W + "top")?.Attribute(W + "val"));          // reject ≡ left

        var rev = Assert.Single(DocxDiff.GetRevisions(left, right, ModeledOnly));
        Assert.Equal(DocxDiffRevisionType.FormatChanged, rev.Type);
        Assert.Equal(DocxDiffFormatChangeScope.TableRow, rev.FormatChange!.Scope);
        Assert.Equal(new[] { "tblPrEx" }, rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void Fresh_trPr_is_placed_after_an_existing_tblPrEx()
    {
        // Review finding 2: when the RIGHT row has a tblPrEx but no trPr and the LEFT row has a trPr, the
        // emitted (right-cloned) row gains a fresh trPr for the trPrChange — which must land AFTER the
        // tblPrEx (CT_Row orders tblPrEx before trPr), or the output is schema-invalid.
        var left = IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPrEx><w:trPr><w:trHeight w:val=\"400\"/></w:trPr>",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));
        var right = IrTestDocuments.FromBodyXml(Table(
            "<w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPrEx>",
            "<w:tcW w:w=\"4000\" w:type=\"dxa\"/>"));

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        var body = BodyOf(result);
        var tr = body.Descendants(W + "tr").First();
        var kids = tr.Elements().Select(e => e.Name.LocalName).ToList();
        int exIdx = kids.IndexOf("tblPrEx"), trPrIdx = kids.IndexOf("trPr");
        Assert.True(exIdx >= 0 && trPrIdx > exIdx, $"trPr must follow tblPrEx; order was [{string.Join(",", kids)}]");
        Assert.Single(body.Descendants(W + "trPrChange"));

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(wd)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
            .Select(e => e.Description).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    // ------------------------------------------------------------------ trailing sectPr-only (w:pgSz)

    private static readonly WmlDocument SectLeft = IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>");

    private static readonly WmlDocument SectRight = IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/></w:sectPr>");

    [Fact]
    public void TrailingSectPrOnly_change_is_tracked_with_native_sectPrChange()
    {
        var result = DocxDiff.Compare(SectLeft, SectRight, ModeledOnly);
        var body = BodyOf(result);
        var sectPr = body.Elements(W + "sectPr").Single();
        var sectPrChange = sectPr.Element(W + "sectPrChange")!;
        Assert.Same(sectPrChange, sectPr.Elements().Last());                               // last child of sectPr
        Assert.NotNull(sectPrChange.Attribute(W + "author"));
        // Right (accepted-state) page size applied; the OLD (left) page size rides in the change inner.
        Assert.Equal("15840", (string?)sectPr.Element(W + "pgSz")?.Attribute(W + "w"));
        Assert.Equal("12240", (string?)sectPrChange.Element(W + "sectPr")!.Element(W + "pgSz")?.Attribute(W + "w"));
        Assert.Empty(sectPrChange.Element(W + "sectPr")!.Elements(W + "sectPrChange"));     // CT_SectPrBase

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal("15840", (string?)BodyOf(accepted).Elements(W + "sectPr").Single().Element(W + "pgSz")?.Attribute(W + "w"));
        Assert.Empty(BodyOf(accepted).Descendants(W + "sectPrChange"));
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal("12240", (string?)BodyOf(rejected).Elements(W + "sectPr").Single().Element(W + "pgSz")?.Attribute(W + "w")); // reject ≡ left

        var rev = Assert.Single(DocxDiff.GetRevisions(SectLeft, SectRight, ModeledOnly));
        Assert.Equal(DocxDiffRevisionType.FormatChanged, rev.Type);
        Assert.Equal(DocxDiffFormatChangeScope.Section, rev.FormatChange!.Scope);
        Assert.Contains("pageWidth", rev.FormatChange.ChangedPropertyNames);
        Assert.Contains("pageHeight", rev.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void SectPrChange_reject_preserves_header_footer_references()
    {
        // The sectPrChange inner is CT_SectPrBase (no references); rejecting must NOT drop the section's
        // header/footer references (RevisionProcessor fix). Left = a header-referencing section whose margins
        // change on the right.
        var left = IrTestDocuments.FromBodyAndHeaderXml(
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>", "<w:p><w:r><w:t>HEADER</w:t></w:r></w:p>");
        // Build the RIGHT with the SAME header wiring but a different margin by editing the left bytes' sectPr.
        var right = WithSectPrMargin(left, "1440");

        var result = DocxDiff.Compare(left, right, ModeledOnly);
        Assert.Single(BodyOf(result).Descendants(W + "sectPrChange"));

        var rejected = RevisionProcessor.RejectRevisions(result);
        var rejSect = BodyOf(rejected).Elements(W + "sectPr").Single();
        Assert.NotEmpty(rejSect.Elements(W + "headerReference"));                           // reference survives reject
    }

    [Fact]
    public void SectPr_family_output_is_schema_valid()
    {
        var result = DocxDiff.Compare(SectLeft, SectRight, ModeledOnly);
        using var ms = new MemoryStream(result.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator()
            .Validate(wd)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
            .Select(e => e.Description)
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    // Return a copy of <paramref name="doc"/> whose trailing sectPr has a w:pgMar with the given uniform margin.
    private static WmlDocument WithSectPrMargin(WmlDocument doc, string margin)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var wd = WordprocessingDocument.Open(ms, true))
        {
            var xdoc = wd.MainDocumentPart!.GetXDocument();
            var sectPr = xdoc.Root!.Element(W + "body")!.Elements(W + "sectPr").Last();
            sectPr.Elements(W + "pgMar").Remove();
            sectPr.Add(new XElement(W + "pgMar",
                new XAttribute(W + "top", margin), new XAttribute(W + "bottom", margin),
                new XAttribute(W + "left", margin), new XAttribute(W + "right", margin)));
            wd.MainDocumentPart.PutXDocument();
        }
        return new WmlDocument("sect-right.docx", ms.ToArray());
    }

    // ------------------------------------------------------------------ direct numbering is modeled (Phase 1)

    [Fact]
    public void MapParaFormat_models_direct_numbering()
    {
        var pPr = XElement.Parse(
            $"<w:pPr xmlns:w=\"{W}\"><w:numPr><w:ilvl w:val=\"2\"/><w:numId w:val=\"5\"/></w:numPr></w:pPr>");
        var f = IrReader.MapParaFormat(pPr);
        Assert.Equal(5, f.NumId);
        Assert.Equal(2, f.Ilvl);

        var empty = IrReader.MapParaFormat(XElement.Parse($"<w:pPr xmlns:w=\"{W}\"/>"));
        Assert.Null(empty.NumId);
        Assert.Null(empty.Ilvl);

        // numPr is CONSUMED by the modeled fields — it no longer rides the unmodeled digest.
        Assert.Equal(empty.UnmodeledDigest, f.UnmodeledDigest);
    }

    [Fact]
    public void FingerprintParaFormat_distinguishes_direct_numbering_via_modeled_fields()
    {
        var a = IrReader.MapParaFormat(XElement.Parse(
            $"<w:pPr xmlns:w=\"{W}\"><w:numPr><w:numId w:val=\"5\"/></w:numPr></w:pPr>"));
        var b = IrReader.MapParaFormat(XElement.Parse(
            $"<w:pPr xmlns:w=\"{W}\"><w:numPr><w:numId w:val=\"7\"/></w:numPr></w:pPr>"));
        Assert.NotEqual(IrHasher.FingerprintParaFormat(a), IrHasher.FingerprintParaFormat(b));
        Assert.Equal(a.UnmodeledDigest, b.UnmodeledDigest);   // the difference is modeled, not digest-borne
    }

    // ------------------------------------------------------------------ consume-side: inline sectPr survives reject

    [Fact]
    public void RejectRevisions_preserves_inline_sectPr_when_rejecting_a_pPrChange()
    {
        // Word semantics: an inline w:sectPr is OUTSIDE pPrChange scope (CT_PPrBase excludes it), so
        // rejecting the paragraph-property change must NOT delete the section break. Latent consume-side
        // bug found by this campaign — see docs/ooxml_corner_cases.md.
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/>" +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
            "<w:pPrChange w:id=\"1\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"><w:pPr/></w:pPrChange></w:pPr>" +
            "<w:r><w:t>Section-final paragraph.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Next section.</w:t></w:r></w:p>" +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>");

        var rejected = RevisionProcessor.RejectRevisions(doc);
        var body = BodyOf(rejected);
        var firstPara = body.Elements(W + "p").First();
        Assert.Single(firstPara.Descendants(W + "sectPr"));                        // the section break survives
        Assert.Empty(body.Descendants(W + "jc"));                                  // the pPr change is rejected
        Assert.Empty(body.Descendants(W + "pPrChange"));
    }

    // ------------------------------------------------------------------ detection at block pairing (Phase 1)

    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };

    private static IrDocument Ir(WmlDocument doc) => IrReader.Read(doc, NoSources);

    [Fact]
    public void PPrOnly_modeled_change_classifies_FormatOnly_under_both_policies()
    {
        foreach (var cmp in new[] { IrFormatComparison.ModeledOnly, IrFormatComparison.Full })
        {
            var a = IrBlockAligner.Align(Ir(PPrLeft), Ir(PPrRight), new IrDiffSettings { FormatComparison = cmp });
            Assert.Equal(IrAlignmentKind.FormatOnly, a.Entries.Single().Kind);
        }
    }

    [Fact]
    public void NumberingOnly_change_classifies_FormatOnly_under_ModeledOnly()
    {
        var left = Ir(IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Item text.</w:t></w:r></w:p>"));
        var right = Ir(IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Item text.</w:t></w:r></w:p>"));
        var a = IrBlockAligner.Align(left, right, new IrDiffSettings());
        Assert.Equal(IrAlignmentKind.FormatOnly, a.Entries.Single().Kind);
    }

    [Fact]
    public void UnmodeledOnly_pPr_change_stays_Unchanged_under_ModeledOnly()
    {
        // The documented ModeledOnly blind spot, pPr edition: paragraph shading is unmodeled, so the
        // delta reads Unchanged under the default (untracked right-apply) and FormatOnly under Full.
        var left = Ir(PPrLeft);
        var right = Ir(IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/></w:pPr><w:r><w:t>Same text here.</w:t></w:r></w:p>"));

        var modeled = IrBlockAligner.Align(left, right, new IrDiffSettings());
        Assert.Equal(IrAlignmentKind.Unchanged, modeled.Entries.Single().Kind);

        var full = IrBlockAligner.Align(left, right, new IrDiffSettings { FormatComparison = IrFormatComparison.Full });
        Assert.Equal(IrAlignmentKind.FormatOnly, full.Entries.Single().Kind);
    }

    [Fact]
    public void TrackBlockFormatChanges_off_restores_Unchanged_classification()
    {
        var a = IrBlockAligner.Align(Ir(PPrLeft), Ir(PPrRight),
            new IrDiffSettings { TrackBlockFormatChanges = false });
        Assert.Equal(IrAlignmentKind.Unchanged, a.Entries.Single().Kind);
    }

    [Fact]
    public void StyleDefinition_only_difference_with_identical_direct_pPr_stays_Unchanged()
    {
        // Detection uses DIRECT pPr facts: when both paragraphs carry the same pStyle and the DIFFERENCE
        // lives in the style definitions part, no pPrChange-describable delta exists → Unchanged.
        const string body = "<w:p><w:pPr><w:pStyle w:val=\"Quote\"/></w:pPr><w:r><w:t>Styled text.</w:t></w:r></w:p>";
        const string stylesA =
            "<w:style w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:pPr><w:jc w:val=\"left\"/></w:pPr></w:style>";
        const string stylesB =
            "<w:style w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:style>";
        var left = Ir(IrTestDocuments.FromBodyAndStylesXml(body, stylesA));
        var right = Ir(IrTestDocuments.FromBodyAndStylesXml(body, stylesB));
        var a = IrBlockAligner.Align(left, right, new IrDiffSettings());
        Assert.Equal(IrAlignmentKind.Unchanged, a.Entries.Single().Kind);
    }

    [Fact]
    public void Consolidate_ignores_block_format_changes_v1_ceiling()
    {
        // Pinned v1 ceiling (the CompareHeadersFooters precedent): the composite merger forces
        // TrackBlockFormatChanges off for its per-reviewer diffs, so a reviewer's pPr-only edit is
        // ignored by Consolidate — no pPrChange, base pPr preserved, no conflict recorded.
        var baseDoc = PPrLeft;
        var reviewer = new DocxDiffReviewer { Document = PPrRight, Author = "Reviewer A" };

        var merged = DocxDiff.Consolidate(baseDoc, new[] { reviewer });
        var body = BodyOf(merged);
        Assert.Empty(body.Descendants(W + "pPrChange"));
        Assert.Empty(body.Descendants(W + "jc"));            // base (no jc) wins; reviewer's centering ignored
        Assert.Empty(DocxDiff.GetConflicts(baseDoc, new[] { reviewer }));
    }

    // ------------------------------------------------------------------ the WmlComparer oracle

    [Fact]
    public void Oracle_WmlComparer_ignores_a_pPr_only_change()
    {
        // Rationale pin for the differential harness: the blessed oracle reports NOTHING for a
        // pPr-only change (it emits no pPrChange anywhere), so IR-side paragraph-scope format
        // revisions bucket as "IR more correct" / oracle-cannot-produce.
        var settings = new WmlComparerSettings();
        var compared = WmlComparer.Compare(PPrLeft, PPrRight, settings);
        var revisions = WmlComparer.GetRevisions(compared, settings);
        Assert.Empty(revisions);
        Assert.Empty(BodyOf(compared).Descendants(W + "pPrChange"));
    }
}
