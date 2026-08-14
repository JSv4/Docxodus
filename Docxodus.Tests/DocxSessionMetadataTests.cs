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
