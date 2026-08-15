#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.IO;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for the block-metadata read surface on <see cref="DocxSession"/>
/// (<c>GetBlockMetadata</c>, <c>GetBlockMetadatas</c>, <c>GetListMembership</c>,
/// <c>GetSectionInfo</c>). Test IDs follow the <c>BM###</c> prefix convention.
/// </summary>
public class DocxSessionMetadataTests
{
    [Fact]
    public void BM001_GetBlockMetadata_PlainParagraph_ReturnsKindAndScope()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        var meta = session.GetBlockMetadata(anchor.Anchor.Id);

        Assert.NotNull(meta);
        Assert.Equal("p", meta!.Kind);
        Assert.Equal("body", meta.Scope);
        Assert.Null(meta.StyleId);
        Assert.Null(meta.StyleName);
        Assert.Null(meta.OutlineLevel);
        Assert.Null(meta.List);
        Assert.False(meta.HasInlineFormatting);
    }

    [Fact]
    public void BM002_GetListMembership_InlineNumPr_BulletList_ReturnsListFacts()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS002_BulletedList());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "li");

        var list = session.GetListMembership(anchor.Anchor.Id);

        Assert.NotNull(list);
        Assert.Equal(1, list!.NumId);
        Assert.Equal(0, list.AbstractNumId);
        Assert.Equal(0, list.Level);
        Assert.Equal(NumberFormat.Bullet, list.Format);
        Assert.True(list.IsAutoNumbered);
        Assert.False(list.FromStyle);
        Assert.Null(list.StartOverride);
    }

    [Fact]
    public void BM003_GetListMembership_NotAList_ReturnsNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        Assert.Null(session.GetListMembership(anchor.Anchor.Id));
    }

    [Fact]
    public void BM004_GetListMembership_StyleInheritedNumPr_SetsFromStyleTrue()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_StyleInheritedList());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "li");

        var list = session.GetListMembership(anchor.Anchor.Id);

        Assert.NotNull(list);
        Assert.True(list!.FromStyle);
        Assert.Equal(1, list.NumId);
        Assert.Equal(0, list.Level);
        Assert.Equal(NumberFormat.Bullet, list.Format);
    }

    [Fact]
    public void BM005_GetSectionInfo_BodyAnchor_ResolvesLandscapeAndHeaders()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_LandscapeSection());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        var info = session.GetSectionInfo(anchor.Anchor.Id);

        Assert.NotNull(info);
        Assert.Equal(16838, info!.PageWidthTwips);
        Assert.Equal(11906, info.PageHeightTwips);
        Assert.True(info.Landscape);
        Assert.Equal(720, info.MarginTopTwips);
        Assert.Equal(720, info.MarginBottomTwips);
        Assert.Equal(1080, info.MarginLeftTwips);
        Assert.Equal(1080, info.MarginRightTwips);
        Assert.Equal(2, info.Columns);
        Assert.Single(info.HeaderPartUris);
        Assert.Empty(info.FooterPartUris);
    }

    [Fact]
    public void BM006_GetSectionInfo_UnknownAnchor_ReturnsNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.Null(session.GetSectionInfo("p:body:does-not-exist"));
    }

    [Fact]
    public void BM007_GetSectionInfo_NonBodyAnchor_ReturnsNull()
    {
        // The landscape-section fixture has a HeaderPart with one paragraph.
        // That paragraph's anchor lives in scope "hdr1", not "body".
        using var session = new DocxSession(DocxSessionTests.BuildBM_LandscapeSection());
        var hdrAnchor = session.Project().AnchorIndex.Values
            .FirstOrDefault(t => t.Anchor.Scope.StartsWith("hdr", System.StringComparison.Ordinal));
        Assert.NotNull(hdrAnchor);

        Assert.Null(session.GetSectionInfo(hdrAnchor!.Anchor.Id));
    }

    [Fact]
    public void BM008_OutlineLevel_FromHeadingStyle_ResolvesToZeroBasedLevel()
    {
        // BuildDS001 has Heading1..6 styles defined. Apply Heading2 to the first paragraph.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");
        var setStyle = session.SetParagraphStyle(anchor.Anchor.Id, "Heading2");
        Assert.True(setStyle.Success);

        // SetParagraphStyle may have changed the anchor kind from "p" to "h" — re-resolve.
        var freshIndex = session.Project().AnchorIndex;
        var promoted = freshIndex.Values.First(t => t.Anchor.Kind == "h");

        var meta = session.GetBlockMetadata(promoted.Anchor.Id);
        Assert.NotNull(meta);
        Assert.Equal("Heading2", meta!.StyleId);
        Assert.Equal("Heading 2", meta.StyleName);
        Assert.Equal(1, meta.OutlineLevel);  // Heading2 → outlineLvl 1 (0-based)
    }

    [Fact]
    public void BM009_GetBlockMetadatas_Bulk_DedupesAndMapsUnknownToNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind == "p").ToList();
        Assert.True(anchors.Count >= 2);

        var ids = new[] {
            anchors[0].Anchor.Id,
            anchors[0].Anchor.Id,         // duplicate
            anchors[1].Anchor.Id,
            "p:body:does-not-exist",
        };

        var map = session.GetBlockMetadatas(ids);

        Assert.Equal(3, map.Count);  // duplicate dropped
        Assert.NotNull(map[anchors[0].Anchor.Id]);
        Assert.NotNull(map[anchors[1].Anchor.Id]);
        Assert.Null(map["p:body:does-not-exist"]);
    }

    [Fact]
    public void BM010_HasInlineFormatting_DetectsBoldRun()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        Assert.False(session.GetBlockMetadata(anchor.Anchor.Id)!.HasInlineFormatting);

        var apply = session.ApplyFormat(anchor.Anchor.Id, span: null, new FormatOp { Bold = true });
        Assert.True(apply.Success);

        Assert.True(session.GetBlockMetadata(anchor.Anchor.Id)!.HasInlineFormatting);
    }

    [Fact]
    public void BM011_ListStyles_ResolvesInheritanceLatentMetadata_AndIdsRoundTrip()
    {
        using var session = new DocxSession(BuildFormattingIntrospectionDocument());
        var styles = session.ListStyles();

        var child = Assert.Single(styles, s => s.Id == "ChildPara");
        Assert.Equal("BasePara", child.BasedOn);
        Assert.Equal("Normal", child.Next);
        Assert.Equal(7, child.UiPriority);
        Assert.True(child.QuickFormat);
        Assert.Equal(720, child.ResolvedParagraph!.LeftIndentTwips);
        Assert.Equal(200, child.ResolvedParagraph.SpacingAfterTwips);
        Assert.True(child.ResolvedRun!.Bold);
        Assert.True(child.ResolvedRun.Italic);

        var character = Assert.Single(styles, s => s.Id == "StrongCustom");
        Assert.Equal("EmphasisBase", character.BasedOn);
        Assert.Equal(4, character.UiPriority);
        Assert.True(character.SemiHidden);
        Assert.True(character.QuickFormat);
        Assert.True(character.ResolvedRun!.Bold);
        Assert.True(character.ResolvedRun.Italic);

        var table = Assert.Single(styles, s => s.Id == "AgentTable");
        Assert.Equal("center", table.ResolvedTable!.Alignment);
        Assert.Equal(5000, table.ResolvedTable.WidthTwips);
        Assert.True(table.ResolvedTable.HasBorders);

        var anchor = session.Project().AnchorIndex.Values.First(v => v.Anchor.Scope == "body").Anchor.Id;
        Assert.True(session.SetParagraphStyle(anchor, child.Id).Success);
        var span = Assert.Single(session.ListInlineSpans(anchor), s => s.Text == "Alpha");
        Assert.True(session.ApplyFormat(span.AnchorId, span.Span,
            new FormatOp { RunStyle = character.Id }).Success);
    }

    [Fact]
    public void BM012_GetFormatting_DistinguishesDirectFromEffective_AndSpansRoundTrip()
    {
        using var session = new DocxSession(BuildFormattingIntrospectionDocument());
        var anchor = session.Project().AnchorIndex.Values.First(v => v.Anchor.Scope == "body").Anchor.Id;

        var formatting = session.GetFormatting(anchor);

        Assert.NotNull(formatting);
        Assert.Equal("ChildPara", formatting!.DirectParagraph.StyleId);
        Assert.Equal(ParagraphAlignment.Right, formatting.DirectParagraph.Alignment);
        Assert.Null(formatting.DirectParagraph.LeftIndentTwips);
        Assert.Equal(720, formatting.EffectiveParagraph.LeftIndentTwips);
        Assert.Equal(300, formatting.EffectiveParagraph.SpacingAfterTwips);
        Assert.Equal(ParagraphAlignment.Right, formatting.EffectiveParagraph.Alignment);

        var alpha = Assert.Single(formatting.Runs, s => s.Text == "Alpha");
        Assert.Equal("StrongCustom", alpha.Direct.StyleId);
        Assert.False(alpha.Direct.Italic);
        Assert.Null(alpha.Direct.Bold);
        Assert.True(alpha.Effective.Bold);
        Assert.False(alpha.Effective.Italic);
        Assert.Equal(12, alpha.Effective.FontSizePts);
        Assert.Equal(anchor, alpha.AnchorId);

        Assert.True(session.ApplyFormat(alpha.AnchorId, alpha.Span,
            new FormatOp { Underline = true }).Success);
        var refreshed = Assert.Single(session.ListInlineSpans(anchor), s => s.Text == "Alpha");
        Assert.True(refreshed.Direct.Underline);
    }

    [Fact]
    public void BM013_ListMembership_ReportsDefinitionStartIndent_AndMutationAnchor()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_StyleInheritedList());
        var anchor = session.Project().AnchorIndex.Values.Single(v => v.Anchor.Kind == "li").Anchor.Id;

        var list = session.GetListMembership(anchor)!;

        Assert.Equal(anchor, list.AnchorId);
        Assert.True(list.FromStyle);
        Assert.Equal(3, list.Start);
        Assert.Equal("·", list.LevelText);
        Assert.Equal(720, list.LeftIndentTwips);
        Assert.Equal(360, list.HangingIndentTwips);
        Assert.True(session.SetListStartOverride(list.AnchorId, 9).Success);
        Assert.Equal(9, session.GetListMembership(list.AnchorId)!.StartOverride);
    }

    [Fact]
    public void BM014_GetSectionInfo_IsPerAnchor_AndSectionIdsStayStableAcrossMutation()
    {
        using var session = new DocxSession(BuildMixedSectionsDocument());
        var anchors = session.Project().AnchorIndex.Values
            .Where(v => v.Anchor.Scope == "body" && v.Anchor.Kind == "p")
            .Select(v => v.Anchor.Id).ToArray();

        var first = session.GetSectionInfo(anchors[0])!;
        var second = session.GetSectionInfo(anchors[1])!;

        Assert.Equal(anchors[0], first.AnchorId);
        Assert.Equal(anchors[1], second.AnchorId);
        Assert.NotEqual(first.SectionUnid, second.SectionUnid);
        Assert.Equal(10000, first.PageWidthTwips);
        Assert.Equal(14000, second.PageWidthTwips);

        Assert.True(session.SetPageNumbering(second.AnchorId,
            new PageNumberingOp { Start = 4 }).Success);
        Assert.Equal(first.SectionUnid, session.GetSectionInfo(first.AnchorId)!.SectionUnid);
        Assert.Equal(second.SectionUnid, session.GetSectionInfo(second.AnchorId)!.SectionUnid);
    }

    [Fact]
    public void BM015_Introspection_IsBytePure_AndIdsSurviveSaveReopen()
    {
        var input = BuildFormattingIntrospectionDocument();
        string anchorId;
        string runUnid;
        string sectionUnid;
        string[] styleIds;
        byte[] saved;

        using (var session = new DocxSession(input))
        {
            anchorId = session.Project().AnchorIndex.Values
                .First(v => v.Anchor.Scope == "body" && v.Anchor.Kind == "p").Anchor.Id;
            var before = SHA256.HashData(session.Save(persistAnchorIds: true));

            styleIds = session.ListStyles().Select(s => s.Id).ToArray();
            var formatting = session.GetFormatting(anchorId)!;
            var spans = session.ListInlineSpans(anchorId);
            var section = session.GetSectionInfo(anchorId)!;

            Assert.Equal(formatting.Runs.Select(s => s.RunUnid), spans.Select(s => s.RunUnid));
            runUnid = spans[0].RunUnid;
            sectionUnid = section.SectionUnid;

            Assert.Equal(runUnid, session.ListInlineSpans(anchorId)[0].RunUnid);
            Assert.Equal(sectionUnid, session.GetSectionInfo(anchorId)!.SectionUnid);
            Assert.Equal(before, SHA256.HashData(session.Save(persistAnchorIds: true)));

            saved = session.Save();
        }

        using var reopened = new DocxSession(saved);
        Assert.Equal(styleIds, reopened.ListStyles().Select(s => s.Id).ToArray());
        Assert.NotNull(reopened.GetFormatting(anchorId));
        Assert.Equal(runUnid, reopened.ListInlineSpans(anchorId)[0].RunUnid);
        Assert.Equal(sectionUnid, reopened.GetSectionInfo(anchorId)!.SectionUnid);
    }

    [Fact]
    public void BM016_DeterministicInspectionFallbacks_DoNotWriteXml()
    {
        using var stream = new MemoryStream(BuildFormattingIntrospectionDocument());
        using var doc = WordprocessingDocument.Open(stream, true);
        var main = doc.MainDocumentPart!;
        var root = main.GetXDocument().Root!;
        var paragraph = root.Descendants(W.p).First();
        paragraph.SetAttributeValue(PtOpenXml.Unid, "testanchor");
        var run = paragraph.Elements(W.r).First();
        var sectPr = root.Descendants(W.sectPr).First();
        var target = new AnchorTarget
        {
            Anchor = new Anchor("p:body:testanchor", "p", "body", "testanchor"),
            PartUri = main.Uri.ToString(),
            Unid = "testanchor",
            TextPreview = "Alpha beta",
        };
        var before = root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

        _ = FormattingIntrospectionOps.ListStyles(doc);
        _ = FormattingIntrospectionOps.GetFormatting(doc, target);
        var firstSpans = FormattingIntrospectionOps.ListInlineSpans(doc, target);
        var firstSection = BlockMetadataOps.GetSectionInfo(doc, target)!;
        var secondSpans = FormattingIntrospectionOps.ListInlineSpans(doc, target);
        var secondSection = BlockMetadataOps.GetSectionInfo(doc, target)!;

        Assert.Equal(before, root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
        Assert.Null(run.Attribute(PtOpenXml.Unid));
        Assert.Null(sectPr.Attribute(PtOpenXml.Unid));
        Assert.Equal(firstSpans[0].RunUnid, secondSpans[0].RunUnid);
        Assert.Equal(firstSection.SectionUnid, secondSection.SectionUnid);

        Assert.True(UnidHelper.AssignToAllElementsDeterministic(root));
        Assert.Equal(firstSpans[0].RunUnid, (string?)run.Attribute(PtOpenXml.Unid));
        Assert.Equal(firstSection.SectionUnid, (string?)sectPr.Attribute(PtOpenXml.Unid));
    }

    // A pure read must not be able to change a later read's answer. GetListMembership resolves
    // the generated label through ListItemRetriever, which plants ListItemInfo annotations on
    // the LIVE paragraphs; ParagraphStyleRollup folds an extra numbering-level pPr layer in
    // whenever it sees one. GetFormatting must be blind to that.
    [Fact]
    public void BM017_GetFormatting_IsUnaffectedByAnInterveningListMembershipRead()
    {
        var input = DocxSessionTests.BuildBM_StyleInheritedList();
        string anchor;
        using (var probe = new DocxSession(input))
        {
            anchor = probe.Project().AnchorIndex.Values.Single(v => v.Anchor.Kind == "li").Anchor.Id;
        }

        using var session = new DocxSession(input);
        var before = session.GetFormatting(anchor)!;
        Assert.NotNull(session.GetListMembership(anchor));
        var after = session.GetFormatting(anchor)!;

        Assert.Equal(before.EffectiveParagraph, after.EffectiveParagraph);
        Assert.Equal(before.DirectParagraph, after.DirectParagraph);
        Assert.Equal(
            before.Runs.Select(r => r.Effective),
            after.Runs.Select(r => r.Effective));
    }

    // The same invariant across the OTHER channel that plants those annotations: Project()
    // enriches (and annotates), the cheap BuildAnchorIndexOnly path does not. Two sessions over
    // identical bytes must resolve identical effective formatting. This is the assertion that
    // discriminates "resolver is annotation-independent" from "the read happens to clean up".
    [Fact]
    public void BM018_GetFormatting_IsIdenticalOnProjectedAndUnprojectedSessions()
    {
        var input = DocxSessionTests.BuildBM_StyleInheritedList();
        using var projected = new DocxSession(input);
        var anchor = projected.Project().AnchorIndex.Values.Single(v => v.Anchor.Kind == "li").Anchor.Id;
        var fromProjected = projected.GetFormatting(anchor)!;

        using var unprojected = new DocxSession(input);
        var fromUnprojected = unprojected.GetFormatting(anchor)!;

        Assert.Equal(fromProjected.EffectiveParagraph, fromUnprojected.EffectiveParagraph);
        // Pinned direction: BOTH report the cascade documented in docx_mutation_api.md, which
        // excludes the numbering level. See BM021 for the number this costs.
        Assert.Equal(0, fromProjected.EffectiveParagraph.LeftIndentTwips);
    }

    // ST_OnOff is case-INSENSITIVE and its parser is PtUtil.ToBoolean — the one the renderer
    // uses. A writer emitting bool.ToString() ("False") must not be reported as bold, and a
    // value outside ST_OnOff must not be reported as bold either, nor throw.
    [Fact]
    public void BM019_OnOffValues_AreParsedCaseInsensitively_AndGarbageIsNotTrue()
    {
        using var session = new DocxSession(BuildOnOffCaseDocument());
        var anchor = session.Project().AnchorIndex.Values
            .Single(v => v.Anchor.Scope == "body" && v.Anchor.Kind == "p").Anchor.Id;

        var spans = session.ListInlineSpans(anchor);
        var mixedCase = Assert.Single(spans, s => s.Text == "Alpha");
        var upperOnOff = Assert.Single(spans, s => s.Text == "Beta");

        Assert.False(mixedCase.Direct.Bold);
        Assert.False(mixedCase.Effective.Bold);
        Assert.True(mixedCase.Direct.Italic);
        Assert.False(upperOnOff.Direct.Bold);
        Assert.True(upperOnOff.Direct.Italic);

        // A value outside ST_OnOff is UNKNOWN, never "on". (Asserted on a paragraph toggle:
        // FormattingAssembler.CharStyleAttributes.GetBoolProperty throws on an out-of-spec RUN
        // toggle, which is pre-existing engine strictness this read API inherits.)
        var formatting = session.GetFormatting(anchor)!;
        Assert.Null(formatting.DirectParagraph.KeepNext);
        Assert.False(formatting.EffectiveParagraph.KeepNext);
        Assert.True(formatting.DirectParagraph.KeepLines);

        // w:default="True" must resolve for the catalog and for the effective-style lookup the
        // same way FormattingAssembler's own default-style scan (already .ToBoolean()) does.
        var normal = Assert.Single(session.ListStyles(), s => s.Id == "Normal");
        Assert.True(normal.IsDefault);
        Assert.Equal("Normal", session.GetFormatting(anchor)!.EffectiveParagraph.StyleId);
    }

    // w:basedOn is caller data, not a guaranteed tree. ListStyles rolls up EVERY style in the
    // catalog, including ones no content references, so a self- or mutually-based style reaches
    // the walkers. Terminating (and returning what accumulated) is the contract.
    [Fact]
    public void BM020_ListStyles_TerminatesOnCyclicBasedOnChains()
    {
        using var session = new DocxSession(BuildCyclicBasedOnDocument());

        var styles = session.ListStyles();

        Assert.Equal(240, Assert.Single(styles, s => s.Id == "SelfCycle").ResolvedParagraph!.LeftIndentTwips);
        Assert.Equal(60, Assert.Single(styles, s => s.Id == "LoopA").ResolvedParagraph!.SpacingAfterTwips);
        Assert.True(Assert.Single(styles, s => s.Id == "CharCycle").ResolvedRun!.Bold);
        Assert.Equal(4000, Assert.Single(styles, s => s.Id == "TableCycle").ResolvedTable!.WidthTwips);
    }

    // DOCUMENTED LIMITATION (not a desired behaviour): the effective-paragraph cascade is
    // docDefaults + pStyle chain + direct pPr. It does NOT include the numbering level's own
    // w:pPr, which is where a list item's indentation normally lives — so the render of this
    // same paragraph is indented 720 twips and introspection reports 0. GetListMembership
    // surfaces the real numbers separately. Unifying the two cascades is deferred; when it
    // lands, these numbers change and this test is the target.
    [Fact]
    public void BM021_EffectiveParagraph_ExcludesTheNumberingLevelIndent_KnownLimitation()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_StyleInheritedList());
        var anchor = session.Project().AnchorIndex.Values.Single(v => v.Anchor.Kind == "li").Anchor.Id;

        var membership = session.GetListMembership(anchor)!;
        var formatting = session.GetFormatting(anchor)!;

        Assert.Equal(720, membership.LeftIndentTwips);
        Assert.Equal(360, membership.HangingIndentTwips);
        Assert.Equal(0, formatting.EffectiveParagraph.LeftIndentTwips);
        Assert.Equal(0, formatting.EffectiveParagraph.HangingIndentTwips);
    }

    // DOCUMENTED LIMITATION (not a desired behaviour): the effective-run cascade is
    // docDefaults + character/paragraph style chain + direct rPr + theme fonts. It does NOT
    // toggle-merge the table style's conditional rPr (w:tblStylePr), so a run in a firstRow-
    // styled table renders bold but introspects as not bold. Deferred with BM021.
    [Fact]
    public void BM022_EffectiveRun_ExcludesConditionalTableStyleFormatting_KnownLimitation()
    {
        using var session = new DocxSession(BuildConditionalTableStyleDocument());
        var anchor = session.Project().AnchorIndex.Values
            .First(v => v.Anchor.Scope == "body" && v.Anchor.Kind == "p"
                && v.TextPreview.Contains("Header cell")).Anchor.Id;

        var span = Assert.Single(session.ListInlineSpans(anchor));

        Assert.Null(span.Direct.Bold);
        Assert.False(span.Effective.Bold);
    }

    /// <summary>
    /// A document written as literal XML, so <c>ST_OnOff</c> values keep the exact lexical form
    /// under test (the SDK's <c>OnOffValue</c> would normalize "False" to "false").
    /// </summary>
    private static byte[] BuildLiteralXmlDocument(string bodyXml, string stylesXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                    + bodyXml + "</w:document>");
            }

            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(stylesXml);
            }
        }
        return stream.ToArray();
    }

    private static byte[] BuildOnOffCaseDocument() => BuildLiteralXmlDocument(
        """
        <w:body>
          <w:p>
            <w:pPr><w:keepNext w:val="huh"/><w:keepLines w:val="ON"/></w:pPr>
            <w:r><w:rPr><w:b w:val="False"/><w:i w:val="True"/></w:rPr><w:t>Alpha</w:t></w:r>
            <w:r><w:rPr><w:b w:val="off"/><w:i w:val="on"/></w:rPr><w:t>Beta</w:t></w:r>
          </w:p>
          <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
        </w:body>
        """,
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="True" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        </w:styles>
        """);

    private static byte[] BuildCyclicBasedOnDocument() => BuildLiteralXmlDocument(
        """
        <w:body>
          <w:p><w:r><w:t>Body</w:t></w:r></w:p>
          <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
        </w:body>
        """,
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
          <w:style w:type="paragraph" w:styleId="SelfCycle"><w:name w:val="Self Cycle"/><w:basedOn w:val="SelfCycle"/><w:pPr><w:ind w:left="240"/></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="LoopA"><w:name w:val="Loop A"/><w:basedOn w:val="LoopB"/><w:pPr><w:spacing w:after="60"/></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="LoopB"><w:name w:val="Loop B"/><w:basedOn w:val="LoopA"/></w:style>
          <w:style w:type="character" w:styleId="CharCycle"><w:name w:val="Char Cycle"/><w:basedOn w:val="CharCycle"/><w:rPr><w:b/></w:rPr></w:style>
          <w:style w:type="table" w:styleId="TableCycle"><w:name w:val="Table Cycle"/><w:basedOn w:val="TableCycle"/><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr></w:style>
        </w:styles>
        """);

    private static byte[] BuildConditionalTableStyleDocument() => BuildLiteralXmlDocument(
        """
        <w:body>
          <w:tbl>
            <w:tblPr><w:tblStyle w:val="ConditionalTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>
            <w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>
            <w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Header cell</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Body cell</w:t></w:r></w:p></w:tc></w:tr>
          </w:tbl>
          <w:p><w:r><w:t>After table</w:t></w:r></w:p>
          <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
        </w:body>
        """,
        """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
          <w:style w:type="table" w:styleId="ConditionalTable"><w:name w:val="Conditional Table"/><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>
        </w:styles>
        """);

    private static byte[] BuildFormattingIntrospectionDocument()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "ChildPara" },
                        new Justification { Val = JustificationValues.Right },
                        new SpacingBetweenLines { After = "300" }),
                    new Run(
                        new RunProperties(
                            new RunStyle { Val = "StrongCustom" },
                            new Italic { Val = false }),
                        new Text("Alpha")),
                    new Run(new Text(" beta"))),
                new SectionProperties(new PageSize { Width = 12240, Height = 15840 })));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write("""
                    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
                      <w:latentStyles w:defUIPriority="99" w:defSemiHidden="1"><w:lsdException w:name="Child Paragraph" w:uiPriority="7" w:qFormat="1" w:semiHidden="0"/></w:latentStyles>
                      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                      <w:style w:type="paragraph" w:customStyle="1" w:styleId="BasePara"><w:name w:val="Base Paragraph"/><w:pPr><w:ind w:left="720"/><w:spacing w:after="200"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
                      <w:style w:type="paragraph" w:customStyle="1" w:styleId="ChildPara"><w:name w:val="Child Paragraph"/><w:basedOn w:val="BasePara"/><w:next w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
                      <w:style w:type="character" w:styleId="EmphasisBase"><w:name w:val="Emphasis Base"/><w:rPr><w:i/></w:rPr></w:style>
                      <w:style w:type="character" w:customStyle="1" w:styleId="StrongCustom"><w:name w:val="Strong Custom"/><w:basedOn w:val="EmphasisBase"/><w:uiPriority w:val="4"/><w:semiHidden/><w:qFormat/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
                      <w:style w:type="table" w:customStyle="1" w:styleId="AgentTable"><w:name w:val="Agent Table"/><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr></w:style>
                    </w:styles>
                    """);
            }
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] BuildMixedSectionsDocument()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new ParagraphProperties(new SectionProperties(
                        new PageSize { Width = 10000, Height = 12000 })),
                    new Run(new Text("First section"))),
                new Paragraph(new Run(new Text("Second section"))),
                new SectionProperties(new PageSize { Width = 14000, Height = 16000 })));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();
        }
        return stream.ToArray();
    }
}
