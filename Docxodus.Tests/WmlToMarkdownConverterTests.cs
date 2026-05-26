#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace Docxodus.Tests;

/// <summary>
/// Tests for <see cref="WmlToMarkdownConverter"/>. Test IDs follow the <c>MD###</c> prefix
/// convention documented in <c>docs/architecture/markdown_projection.md</c>. Phase numbering
/// in the doc maps to test ID ranges: phase 1 (anchor index) = MD001-MD009, phase 2
/// (paragraphs/headings) = MD010-MD019, phase 3 (inline runs) = MD020-MD029, phase 4 (lists)
/// = MD030-MD039, phase 5 (tables) = MD040-MD049, phase 6 (multipart) = MD050-MD059, phase 7
/// (tracked changes) = MD060-MD069.
/// </summary>
public class WmlToMarkdownConverterTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static WmlDocument LoadFixture(string fixtureName) =>
        new WmlDocument(Path.Combine(TestFilesDir.FullName, fixtureName));

    // ----- Programmatic fixture builders (reused across phases) -----

    /// <summary>Create a minimal WmlDocument with a single Normal-style paragraph.</summary>
    private static WmlDocument BuildSimpleDoc(string text) =>
        BuildDoc(body => body.Append(new Paragraph(new Run(new Text(text)))));

    /// <summary>Create a minimal WmlDocument with a single styled-heading paragraph.</summary>
    private static WmlDocument BuildHeadingDoc(string text, int level)
    {
        return BuildDoc(body =>
        {
            var pPr = new ParagraphProperties(new ParagraphStyleId { Val = $"Heading{level}" });
            body.Append(new Paragraph(pPr, new Run(new Text(text))));
        }, addHeadingStyles: true);
    }

    /// <summary>Create a minimal WmlDocument and configure its body.</summary>
    private static WmlDocument BuildDoc(Action<Body> configureBody, bool addHeadingStyles = false)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            if (addHeadingStyles)
            {
                for (var i = 1; i <= 9; i++)
                {
                    styles.Append(new Style
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = $"Heading{i}",
                        StyleName = new StyleName { Val = $"heading {i}" },
                    });
                }
            }
            stylesPart.Styles = styles;

            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            configureBody(body);
            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    // ----- Phase 1: anchor index -----

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    [InlineData("HC004-ResumeTemplate.docx")]
    public void MD001_AnchorIndexIsExhaustive(string fixtureName)
    {
        var doc = LoadFixture(fixtureName);
        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart!.Document.Body!;
        var expectedBlockCount = body.Descendants()
            .Count(d => d.LocalName == "p" || d.LocalName == "tbl");

        Assert.NotEmpty(projection.AnchorIndex);

        var bodyAnchors = projection.AnchorIndex.Values.Count(t => t.Anchor.Scope == "body");
        Assert.True(bodyAnchors >= expectedBlockCount,
            $"Expected at least {expectedBlockCount} body anchors, got {bodyAnchors}");
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD002_AnchorsAreStable(string fixtureName)
    {
        // Projecting the SAME document twice MUST produce the same anchor ids — the first
        // projection persists Unids into the byte array, so the second projection reuses them.
        // (Two cold loads of the same bytes legitimately get different Unids — the fixture has
        // none stored — and that is not what "stable" means in the spec.)
        var doc = LoadFixture(fixtureName);
        var p1 = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
        var p2 = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        Assert.Equal(
            p1.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray(),
            p2.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray());
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD003_AnchorsResolve(string fixtureName)
    {
        var doc = LoadFixture(fixtureName);
        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);

        Assert.NotEmpty(projection.AnchorIndex);
        foreach (var kvp in projection.AnchorIndex)
        {
            var target = kvp.Value;
            var element = target.Resolve(wdoc);
            Assert.True(element != null, $"Failed to resolve anchor {kvp.Key}");
            Assert.Equal(target.Unid, (string?)element!.Attribute(PtOpenXml.Unid));
        }
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD004_AnchorsSurviveRoundTrip(string fixtureName)
    {
        var doc = LoadFixture(fixtureName);
        var first = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        // Round-trip the in-memory bytes through OpenXmlMemoryStreamDocument.
        WmlDocument roundTripped;
        using (var sm = new OpenXmlMemoryStreamDocument(doc))
        {
            using (var w = sm.GetWordprocessingDocument())
            {
                w.MainDocumentPart!.Document.Save();
            }
            roundTripped = sm.GetModifiedWmlDocument();
        }

        var second = WmlToMarkdownConverter.Convert(roundTripped, new WmlToMarkdownConverterSettings());
        Assert.Equal(
            first.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray(),
            second.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray());
    }

    // ----- Phase 2: paragraphs + headings -----

    [Fact]
    public void MD010_BodyScopeHeaderEmitted()
    {
        var doc = BuildSimpleDoc("Hello");
        var p = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.Body });
        Assert.StartsWith("# Document", p.Markdown);
    }

    [Fact]
    public void MD011_ParagraphRendersWithAnchorOnOwnLine()
    {
        var doc = BuildSimpleDoc("Hello world");
        var p = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
        Assert.Matches(@"\{#p:body:[0-9a-f]{32}\} Hello world", p.Markdown);
    }

    [Fact]
    public void MD012_HeadingRendersAtCorrectLevel()
    {
        var p1 = WmlToMarkdownConverter.Convert(BuildHeadingDoc("Title", 1),
            new WmlToMarkdownConverterSettings());
        Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} # Title", p1.Markdown);

        var p3 = WmlToMarkdownConverter.Convert(BuildHeadingDoc("Sub", 3),
            new WmlToMarkdownConverterSettings());
        Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} ### Sub", p3.Markdown);
    }

    [Fact]
    public void MD013_HeadingLevelOffsetApplies()
    {
        var doc = BuildHeadingDoc("Title", 1);
        var p = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { HeadingLevelOffset = 1 });
        Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} ## Title", p.Markdown);
    }

    [Theory]
    [InlineData(7, "#######")]
    [InlineData(8, "########")]
    [InlineData(9, "#########")]
    public void MD017_HeadingLevelsBeyond6PreserveDepth(int wordLevel, string expectedPrefix)
    {
        // Legal/clause-numbered DOCX files routinely use Word's Heading7-9 styles for
        // sub-sub-clauses; silently clamping to ###### loses the outline depth that callers
        // (LLMs, ToC builders, diff renderers) rely on. Emit 7-9 hashes verbatim so depth
        // survives, even though strict CommonMark caps ATX headings at 6.
        var doc = BuildHeadingDoc("Deep clause", wordLevel);
        var md = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings()).Markdown;
        Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} " + expectedPrefix + @" Deep clause", md);
    }

    [Fact]
    public void MD018_EmptyParagraphAnchorHasNoTrailingSpace()
    {
        // The anchor prefix carries a trailing space so it doesn't collide with inline
        // content, but for paragraphs that have no runs (visual spacers in Word) that
        // leaves a stray `{#p:body:UNID} ` line ending in whitespace. Strip the trailing
        // space when there's no inline content to separate from.
        var doc = BuildDoc(body => body.Append(new Paragraph()));
        var md = MarkdownOf(doc);
        Assert.Matches(@"\{#p:body:[0-9a-f]{32}\}\n", md);
        Assert.DoesNotMatch(@"\{#p:body:[0-9a-f]{32}\} \n", md);
    }

    [Fact]
    public void MD014_AnchorModeNoneOmitsTokens()
    {
        var doc = BuildSimpleDoc("Hello world");
        var p = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { AnchorMode = AnchorRenderMode.None });
        Assert.DoesNotContain("{#p:", p.Markdown);
        Assert.Contains("Hello world", p.Markdown);
    }

    [Fact]
    public void MD015_NumberedHeadingResolvesNumPrefix()
    {
        // Legal docs frequently style each clause as Heading2/3/etc AND attach w:numPr so
        // Word renders "FIRST: …", "1.1 …", etc. The auto-number must survive projection;
        // otherwise the markdown shows only the trailing text and loses ordinal context.
        var doc = BuildNumberedHeadingDoc("The name of this corporation is X.");
        var md = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings()).Markdown;
        Assert.Matches(
            @"\{#h:body:[0-9a-f]{32}\} ## 1\.\s+The name of this corporation is X\.",
            md);
    }

    [Fact]
    public void MD016_NumberedHeadingNumberingPrefixHiddenWhenResolveDisabled()
    {
        // With ResolveNumbering=false we treat the paragraph as a plain heading — no
        // resolved number prefix. The markdown is what's authored, no Word-side math.
        var doc = BuildNumberedHeadingDoc("Body text");
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { ResolveNumbering = false }).Markdown;
        Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} ## Body text", md);
        Assert.DoesNotMatch(@"## 1\.\s+Body text", md);
    }

    private static WmlDocument BuildNumberedHeadingDoc(string text)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;

            // Styles part with Heading2 declared.
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Heading2", StyleName = new StyleName { Val = "heading 2" } });

            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // Numbering definition with a single ordered level.
            var numPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." },
                        new StartNumberingValue { Val = 1 })
                    { LevelIndex = 0 })
                { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });

            // Heading2 paragraph with both Heading2 style AND numPr.
            var pPr = new ParagraphProperties(
                new ParagraphStyleId { Val = "Heading2" },
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            body.Append(new Paragraph(pPr, new Run(new Text(text))));
            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    // ----- Phase 3: inline runs -----

    private static WmlDocument BuildRunsDoc(params (string text, bool bold, bool italic, bool code, bool strike)[] runs) =>
        BuildDoc(body =>
        {
            var paragraph = new Paragraph();
            foreach (var (text, bold, italic, code, strike) in runs)
            {
                var rPr = new RunProperties();
                if (bold) rPr.Append(new Bold());
                if (italic) rPr.Append(new Italic());
                if (code) rPr.Append(new RunStyle { Val = "Code" });
                if (strike) rPr.Append(new Strike());
                paragraph.Append(new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            }
            body.Append(paragraph);
        });

    private static string MarkdownOf(WmlDocument doc) =>
        WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings()).Markdown;

    [Fact]
    public void MD020_BoldRun()
    {
        var md = MarkdownOf(BuildRunsDoc(("hello", true, false, false, false)));
        Assert.Contains("**hello**", md);
    }

    [Fact]
    public void MD021_ItalicRun()
    {
        var md = MarkdownOf(BuildRunsDoc(("hello", false, true, false, false)));
        Assert.Contains("*hello*", md);
    }

    [Fact]
    public void MD022_CodeRun()
    {
        var md = MarkdownOf(BuildRunsDoc(("x", false, false, true, false)));
        Assert.Contains("`x`", md);
    }

    [Fact]
    public void MD023_StrikeRun()
    {
        var md = MarkdownOf(BuildRunsDoc(("x", false, false, false, true)));
        Assert.Contains("~~x~~", md);
    }

    [Fact]
    public void MD024_CombinedBoldItalic()
    {
        var md = MarkdownOf(BuildRunsDoc(("hi", true, true, false, false)));
        Assert.Contains("***hi***", md);
    }

    [Fact]
    public void MD025_BoldCancelsAcrossAdjacentRuns()
    {
        var md = MarkdownOf(BuildRunsDoc(
            ("a", true, false, false, false),
            ("b", true, false, false, false)));
        Assert.Contains("**ab**", md);
        Assert.DoesNotContain("**a****b**", md);
    }

    [Fact]
    public void MD026_HyperlinkRendersAsLink()
    {
        var doc = BuildHyperlinkDoc("click here", "https://example.com");
        var md = MarkdownOf(doc);
        // Uri canonicalizes "https://example.com" to "https://example.com/"; assert either form.
        Assert.Matches(@"\[click here\]\(https://example\.com/?\)", md);
    }

    [Fact]
    public void MD027_EscapesMarkdownMetacharacters()
    {
        // Use a character set that doesn't include the leading "{#" pattern (anchor) the
        // emitter writes for us. The point of MD027 is that user content can't smuggle markup.
        var md = MarkdownOf(BuildRunsDoc(("a*b_c[d]", false, false, false, false)));
        Assert.Contains(@"a\*b\_c\[d\]", md);
    }

    // ----- Phase 4: lists -----

    [Theory]
    [InlineData("HC031-Complicated-Document.docx")]
    public void MD030_FixtureContainsListItemAnchors(string fixture)
    {
        var p = WmlToMarkdownConverter.Convert(LoadFixture(fixture), new WmlToMarkdownConverterSettings());
        // A real-world Word document with list items; at least one anchor with kind=li
        // should appear in the index, and at least one matching line should begin with
        // the corresponding token + indented marker.
        Assert.Contains(p.AnchorIndex.Values, t => t.Anchor.Kind == "li");
        Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+(-|\d+\.|[a-zA-Z]\.)\s", p.Markdown);
    }

    [Fact]
    public void MD031_BulletedListUsesDashMarkers()
    {
        var doc = BuildBulletedListDoc("first", "second");
        var md = MarkdownOf(doc);
        Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+-\s+first", md);
        Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+-\s+second", md);
    }

    [Fact]
    public void MD032_NumberedListUsesResolvedMarkers()
    {
        var doc = BuildNumberedListDoc("first", "second");
        var md = MarkdownOf(doc);
        Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+1\.\s+first", md);
        Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+2\.\s+second", md);
    }

    [Fact]
    public void MD033_ResolveNumberingFalseFallsBackToDashes()
    {
        var doc = BuildNumberedListDoc("first", "second");
        var p = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { ResolveNumbering = false });
        // With resolution disabled, ordered lists still emit "-" markers.
        Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+-\s+first", p.Markdown);
        Assert.DoesNotMatch(@"\{#li:body:[0-9a-f]{32}\}\s+1\.\s", p.Markdown);
    }

    private static WmlDocument BuildBulletedListDoc(params string[] items) => BuildListDoc(items, ordered: false);
    private static WmlDocument BuildNumberedListDoc(params string[] items) => BuildListDoc(items, ordered: true);

    private static WmlDocument BuildListDoc(string[] items, bool ordered)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // NumberingDefinitionsPart with one abstractNum/num.
            var numPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            var fmt = ordered ? NumberFormatValues.Decimal : NumberFormatValues.Bullet;
            var lvlText = ordered ? "%1." : "·";
            numPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = fmt },
                        new LevelText { Val = lvlText },
                        new StartNumberingValue { Val = 1 })
                    { LevelIndex = 0 })
                { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });

            foreach (var item in items)
            {
                var pPr = new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 1 }));
                body.Append(new Paragraph(pPr, new Run(new Text(item))));
            }
            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    // ----- Phase 5: tables -----

    [Fact]
    public void MD040_SimpleTableRendersAsGfmPipeTable()
    {
        var doc = BuildSimpleTable(new[,] { { "a", "b" }, { "c", "d" } });
        var md = MarkdownOf(doc);
        Assert.Contains("| a | b |", md);
        Assert.Contains("| --- | --- |", md);
        Assert.Contains("| c | d |", md);
    }

    [Fact]
    public void MD041_TableWithMergedCellsFallsBackToOpaque()
    {
        var doc = BuildTableWithMergedHeader();
        var md = MarkdownOf(doc);
        Assert.Contains("```table", md);
        Assert.Contains("rows:", md);
        Assert.Contains("cols:", md);
        Assert.DoesNotContain("| --- |", md);
    }

    [Fact]
    public void MD042_TableInlineCellMaxTriggersOpaque()
    {
        var doc = BuildSimpleTable(new[,] { { "abcdef", "x" } });
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { TableInlineCellMax = 4 }).Markdown;
        // Cell "abcdef" is 6 chars and exceeds the limit; expect opaque fallback.
        Assert.Contains("```table", md);
    }

    [Fact]
    public void MD043_AlwaysOpaqueAlwaysEmitsOpaqueBlock()
    {
        var doc = BuildSimpleTable(new[,] { { "a", "b" }, { "c", "d" } });
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { TableMode = TableRenderMode.AlwaysOpaque }).Markdown;
        Assert.Contains("```table", md);
        Assert.DoesNotContain("| --- |", md);
    }

    [Fact]
    public void MD044_TableAnchorPrecedesContent()
    {
        var doc = BuildSimpleTable(new[,] { { "a", "b" } });
        var md = MarkdownOf(doc);
        // The anchor token for the table should appear before the table content.
        Assert.Matches(@"\{#tbl:body:[0-9a-f]{32}\}\s+\|\s+a\s+\|\s+b\s+\|", md);
    }

    private static WmlDocument BuildSimpleTable(string[,] cells)
    {
        var rows = cells.GetLength(0);
        var cols = cells.GetLength(1);
        return BuildDoc(body =>
        {
            var table = new WTable();
            for (var r = 0; r < rows; r++)
            {
                var row = new WTableRow();
                for (var c = 0; c < cols; c++)
                {
                    row.Append(new WTableCell(new Paragraph(new Run(new Text(cells[r, c])))));
                }
                table.Append(row);
            }
            body.Append(table);
        });
    }

    private static WmlDocument BuildTableWithMergedHeader()
    {
        return BuildDoc(body =>
        {
            var row1Cell1 = new WTableCell(
                new TableCellProperties(new GridSpan { Val = 2 }),
                new Paragraph(new Run(new Text("merged header"))));
            var row1 = new WTableRow(row1Cell1);
            var row2 = new WTableRow(
                new WTableCell(new Paragraph(new Run(new Text("a")))),
                new WTableCell(new Paragraph(new Run(new Text("b")))));
            body.Append(new WTable(row1, row2));
        });
    }

    // ----- Phase 6: multipart scopes -----

    [Fact]
    public void MD050_HeaderEmittedWhenPresent()
    {
        var doc = BuildHeaderFooterDoc(headerText: "CONFIDENTIAL", footerText: null);
        var md = MarkdownOf(doc);
        Assert.Contains("# Headers", md);
        Assert.Contains("## hdr1", md);
        Assert.Contains("CONFIDENTIAL", md);
    }

    [Fact]
    public void MD051_FooterEmittedWhenPresent()
    {
        var doc = BuildHeaderFooterDoc(headerText: null, footerText: "Page 1");
        var md = MarkdownOf(doc);
        Assert.Contains("# Footers", md);
        Assert.Contains("## ftr1", md);
        Assert.Contains("Page 1", md);
    }

    [Fact]
    public void MD052_ScopesBodyOnlySkipsOtherSections()
    {
        var doc = BuildHeaderFooterDoc(headerText: "H", footerText: "F");
        var p = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.Body });
        Assert.DoesNotContain("# Headers", p.Markdown);
        Assert.DoesNotContain("# Footers", p.Markdown);
    }

    [Fact]
    public void MD053_FootnoteEmittedAsGfmFootnote()
    {
        var doc = BuildFootnoteDoc("Body text.", "This is a footnote.");
        var md = MarkdownOf(doc);
        // Inline reference in the body paragraph.
        Assert.Matches(@"\[\^fn-[0-9a-f]+\]", md);
        // Definitions section.
        Assert.Contains("# Footnotes", md);
        Assert.Contains("This is a footnote.", md);
    }

    [Fact]
    public void MD054_InterScopeSeparatorEmittedBetweenSections()
    {
        // The projection-spec example shows `---` between # Document / # Headers / # Footers /
        // # Footnotes scope sections — those breaks let downstream parsers cleanly split the
        // markdown stream into per-scope chunks without inspecting heading text.
        var doc = BuildHeaderFooterDoc(headerText: "H", footerText: "F");
        var md = MarkdownOf(doc);
        // The body section ends, then `---` precedes # Headers.
        Assert.Matches(@"(?ms)^---\s*\n\s*\n# Headers", md);
        // The headers section ends, then `---` precedes # Footers.
        Assert.Matches(@"(?ms)^---\s*\n\s*\n# Footers", md);
    }

    [Fact]
    public void MD055_InterScopeSeparatorOmittedWhenOnlyOneScope()
    {
        // No header or footer parts — no scope dividers needed. Avoid emitting trailing `---`.
        var doc = BuildSimpleDoc("Hello");
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.Body }).Markdown;
        Assert.DoesNotContain("\n---\n", md);
    }

    [Fact]
    public void MD056_SectPrEmitsThematicBreakWithSectionAnchor()
    {
        // A w:sectPr inside a paragraph's pPr marks a section break at that paragraph.
        // The projection should emit `---` preceded by a {#sec:body:UNID} anchor so callers
        // can navigate section boundaries.
        var doc = BuildTwoSectionDoc("first section", "second section");
        var p = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
        Assert.Matches(@"\{#sec:body:[0-9a-f]{32}\}\s*\n---", p.Markdown);
        // The AnchorIndex includes a kind=sec entry that resolves to a sectPr element.
        Assert.Contains(p.AnchorIndex.Values, t => t.Anchor.Kind == "sec" && t.Anchor.Scope == "body");
    }

    private static WmlDocument BuildTwoSectionDoc(string first, string second)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // First paragraph ends with a section break in its pPr.
            var firstSect = new SectionProperties();
            var firstPara = new Paragraph(
                new ParagraphProperties(firstSect),
                new Run(new Text(first)));
            body.Append(firstPara);

            // Second paragraph; body ends with a final sectPr (standard OOXML).
            body.Append(new Paragraph(new Run(new Text(second))));
            body.Append(new SectionProperties());

            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    [Fact]
    public void MD057_EmptyHeaderScopeIsOmitted()
    {
        // A real DOCX often defines 6+ header/footer parts for first-page / even-page /
        // default variants, leaving the unused ones blank. Emitting "## hdrN" titles for
        // empty parts pads the projection with noise; skip scopes whose only content is
        // whitespace.
        var doc = BuildHeaderFooterDoc(headerText: "   ", footerText: "real footer");
        var md = MarkdownOf(doc);
        Assert.DoesNotContain("# Headers", md);
        Assert.DoesNotContain("## hdr1", md);
        Assert.Contains("# Footers", md);
        Assert.Contains("real footer", md);
    }

    [Fact]
    public void MD058_NonEmptyHeaderScopeStillEmitsTitle()
    {
        // Sanity check the inverse — the suppression must NOT eat populated scopes.
        var doc = BuildHeaderFooterDoc(headerText: "CONFIDENTIAL", footerText: null);
        var md = MarkdownOf(doc);
        Assert.Contains("# Headers", md);
        Assert.Contains("## hdr1", md);
        Assert.Contains("CONFIDENTIAL", md);
    }

    private static WmlDocument BuildHeaderFooterDoc(string? headerText, string? footerText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            body.Append(new Paragraph(new Run(new Text("body"))));

            HeaderReference? headerRef = null;
            FooterReference? footerRef = null;
            if (headerText != null)
            {
                var hp = mainPart.AddNewPart<HeaderPart>();
                hp.Header = new Header(new Paragraph(new Run(new Text(headerText))));
                headerRef = new HeaderReference { Id = mainPart.GetIdOfPart(hp), Type = HeaderFooterValues.Default };
            }
            if (footerText != null)
            {
                var fp = mainPart.AddNewPart<FooterPart>();
                fp.Footer = new Footer(new Paragraph(new Run(new Text(footerText))));
                footerRef = new FooterReference { Id = mainPart.GetIdOfPart(fp), Type = HeaderFooterValues.Default };
            }
            var sectPr = new SectionProperties();
            if (headerRef != null) sectPr.Append(headerRef);
            if (footerRef != null) sectPr.Append(footerRef);
            body.Append(sectPr);

            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    private static WmlDocument BuildFootnoteDoc(string bodyText, string footnoteText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fp = mainPart.AddNewPart<FootnotesPart>();
            fp.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new Text(footnoteText)))) { Id = 1 });

            var bodyPara = new Paragraph(
                new Run(new Text(bodyText)),
                new Run(new FootnoteReference { Id = 1 }));
            body.Append(bodyPara);

            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    // ----- Phase 7: tracked changes -----

    [Fact]
    public void MD060_AcceptModeKeepsInsKeepsTextDropsDel()
    {
        var doc = BuildTrackedChangesDoc(
            beforeText: "Hello ",
            insertedText: "brave ",
            deletedText: "cruel ",
            afterText: "world");
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { TrackedChanges = TrackedChangeMode.Accept }).Markdown;
        Assert.Contains("Hello", md);
        Assert.Contains("brave", md);
        Assert.Contains("world", md);
        Assert.DoesNotContain("cruel", md);
    }

    [Fact]
    public void MD061_RenderInlineModeEmitsBraceMarkers()
    {
        var doc = BuildTrackedChangesDoc(
            beforeText: "Hello ",
            insertedText: "brave",
            deletedText: "cruel",
            afterText: " world");
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { TrackedChanges = TrackedChangeMode.RenderInline }).Markdown;
        Assert.Contains("{+brave+}", md);
        Assert.Contains("{-cruel-}", md);
    }

    [Fact]
    public void MD062_StripDeletionsKeepsInsDropsDel()
    {
        var doc = BuildTrackedChangesDoc(
            beforeText: "Hello ",
            insertedText: "brave ",
            deletedText: "cruel ",
            afterText: "world");
        var md = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { TrackedChanges = TrackedChangeMode.StripDeletions }).Markdown;
        Assert.Contains("brave", md);
        Assert.DoesNotContain("cruel", md);
        Assert.DoesNotContain("{+", md);
        Assert.DoesNotContain("{-", md);
    }

    private static WmlDocument BuildTrackedChangesDoc(string beforeText, string insertedText, string deletedText, string afterText)
    {
        return BuildDoc(body =>
        {
            var paragraph = new Paragraph(
                new Run(new Text(beforeText) { Space = SpaceProcessingModeValues.Preserve }),
                new InsertedRun(new Run(new Text(insertedText) { Space = SpaceProcessingModeValues.Preserve }))
                    { Id = "1", Author = "tester", Date = DateTime.UtcNow },
                new DeletedRun(new Run(new DeletedText(deletedText) { Space = SpaceProcessingModeValues.Preserve }))
                    { Id = "2", Author = "tester", Date = DateTime.UtcNow },
                new Run(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve }));
            body.Append(paragraph);
        });
    }

    private static WmlDocument BuildHyperlinkDoc(string text, string url)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var rel = mainPart.AddHyperlinkRelationship(new Uri(url, UriKind.Absolute), true);
            var paragraph = new Paragraph(
                new Hyperlink(new Run(new Text(text))) { Id = rel.Id });
            body.Append(paragraph);
            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    // ─── EmptyParagraphMode (#135) ────────────────────────────────────

    private static WmlDocument BuildDocWithSpacerParagraph()
    {
        // First paragraph has visible text, second is an empty spacer, third has text again.
        // The spacer is the test target: an empty <w:p> with no runs at all.
        return BuildDoc(body =>
        {
            body.Append(new Paragraph(new Run(new Text("Before."))));
            body.Append(new Paragraph());          // <-- the empty spacer
            body.Append(new Paragraph(new Run(new Text("After."))));
        });
    }

    [Fact]
    public void MD030_EmptyParagraph_AnchorOnly_Default()
    {
        var md = WmlToMarkdownConverter.Convert(BuildDocWithSpacerParagraph(),
            new WmlToMarkdownConverterSettings()).Markdown;

        // Default mode: the spacer appears as a bare anchor line with no text and no marker.
        Assert.Matches(@"\{#p:body:[0-9a-f]{32}\}\n", md);
        Assert.DoesNotContain('∅', md);
    }

    [Fact]
    public void MD031_EmptyParagraph_MarkedEmpty()
    {
        var md = WmlToMarkdownConverter.Convert(BuildDocWithSpacerParagraph(),
            new WmlToMarkdownConverterSettings { EmptyParagraphs = EmptyParagraphMode.MarkedEmpty }).Markdown;

        // Marked mode: the spacer ends with `∅`.
        Assert.Matches(@"\{#p:body:[0-9a-f]{32}\}\s*∅", md);
    }

    /// <summary>
    /// Heading1 paragraph whose numbering comes from the STYLE (Heading1 declares
    /// its own w:numPr) — not from any inline w:numPr on the paragraph. Mirrors
    /// the NVCA Model COI's "First Article" / "Second Article" headings where the
    /// markdown projector previously emitted no number prefix while the HTML
    /// converter rendered "1." / "2." — issue #141.
    /// </summary>
    private static WmlDocument BuildHeadingWithStyleNumbering(string text)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Body = body;

            var numPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." },
                        new StartNumberingValue { Val = 1 })
                    { LevelIndex = 0 })
                { AbstractNumberId = 7 },
                new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 9 });

            // Heading1 style with numPr in the STYLE — exactly the NVCA pattern.
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(
                    new StyleName { Val = "heading 1" },
                    new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = 0 },
                            new NumberingId { Val = 9 })))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1",
                });

            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // Paragraph has ONLY <w:pStyle val="Heading1"/> — no inline numPr.
            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text(text))));
            mainPart.Document.Save();
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    [Fact]
    public void MD033_HeadingNumberPrefix_ResolvesFromStyleLevelNumPr()
    {
        // Regression for #141: projector was checking inline w:numPr only and
        // short-circuiting for style-inherited numbering. Result: NVCA Heading1
        // paragraphs rendered as "# That the name..." with no "1." prefix, while
        // the HTML converter rendered them with "1.". Fix: let ListItemRetriever
        // decide (it handles style-level numPr correctly).
        var doc = BuildHeadingWithStyleNumbering("That the name of this corporation.");
        var md = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings()).Markdown;
        Assert.Matches(
            @"\{#h:body:[0-9a-f]{32}\} # 1\. That the name of this corporation\.",
            md);
    }

    [Fact]
    public void MD032_EmptyParagraph_Suppress_DropsFromMarkdownAndIndex()
    {
        var proj = WmlToMarkdownConverter.Convert(BuildDocWithSpacerParagraph(),
            new WmlToMarkdownConverterSettings { EmptyParagraphs = EmptyParagraphMode.Suppress });

        // Only the two text-bearing paragraphs survive in the anchor index.
        var bodyParagraphs = proj.AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .ToList();
        Assert.Equal(2, bodyParagraphs.Count);

        // Markdown contains Before. and After. but no orphan {#p:body:…} bare line
        // between them (the spacer is gone, so no anchor for it appears anywhere).
        Assert.Contains("Before.", proj.Markdown);
        Assert.Contains("After.", proj.Markdown);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(proj.Markdown, @"\{#p:body:[0-9a-f]{32}\}").Count);
    }

    [Fact]
    public void MD005_AnchorTargetCarriesTextPreview()
    {
        // A two-paragraph body — each AnchorTarget should expose the first ~80 chars
        // of the element's flat text directly, without needing a separate session walk.
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var wml = new WmlDocument("test.docx", bytes);

        var projection = WmlToMarkdownConverter.Convert(wml, new WmlToMarkdownConverterSettings());

        var bodyParas = projection.AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
            .ToList();

        Assert.NotEmpty(bodyParas);
        foreach (var t in bodyParas)
        {
            // Every body block has a non-null preview; for non-empty paragraphs it's non-empty.
            Assert.NotNull(t.TextPreview);
        }

        // At least one paragraph has the expected literal first chars
        // (BuildDS001 paragraphs start with "First paragraph" and "Second paragraph").
        Assert.Contains(bodyParas, t => t.TextPreview.StartsWith("First paragraph"));
        Assert.Contains(bodyParas, t => t.TextPreview.StartsWith("Second paragraph"));
    }

    [Fact]
    public void MD080_AnchorIdRendering_FullUnid_IsExistingBehavior()
    {
        // Default rendering — anchor tokens carry the full 32-char hex Unid.
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var wml = new WmlDocument("test.docx", bytes);
        var settings = new WmlToMarkdownConverterSettings();   // AnchorIdRendering defaults to FullUnid
        var projection = WmlToMarkdownConverter.Convert(wml, settings);

        // Every {#…} token uses the full 32-char hex form (matched by regex {#kind:scope:UNID}).
        var match = System.Text.RegularExpressions.Regex.Match(
            projection.Markdown, @"\{#[^:]+:[^:]+:([a-f0-9]+)\}");
        Assert.True(match.Success, "expected at least one anchor token in the projection");
        Assert.Equal(32, match.Groups[1].Length);
    }

    [Fact]
    public void MD081_AnchorIdRendering_Abbreviated_UsesShortestUniquePrefixPerScope()
    {
        // Abbreviated mode picks the shortest prefix per (kind, scope) bucket
        // that uniquely identifies each anchor, with a 4-char floor.
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var wml = new WmlDocument("test.docx", bytes);
        var settings = new WmlToMarkdownConverterSettings
        {
            AnchorIdRendering = AnchorIdRendering.Abbreviated,
        };
        var projection = WmlToMarkdownConverter.Convert(wml, settings);

        // All emitted tokens are < 32 chars in their unid portion, and >= 4.
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(projection.Markdown, @"\{#[^:]+:[^:]+:([a-f0-9]+)\}"))
        {
            var unid = m.Groups[1].Value;
            Assert.True(unid.Length >= 4, $"abbreviation '{unid}' shorter than 4-char floor");
            Assert.True(unid.Length < 32, $"abbreviation '{unid}' wasn't actually abbreviated");
        }
    }

    [Fact]
    public void MD082_AnchorIdRendering_Abbreviated_AnchorIndexHasDualKeys()
    {
        // The AnchorIndex must contain entries keyed by BOTH the full Unid and
        // the abbreviated id, both pointing at the same AnchorTarget. Callers
        // can use whichever form they have in hand.
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var wml = new WmlDocument("test.docx", bytes);
        var settings = new WmlToMarkdownConverterSettings
        {
            AnchorIdRendering = AnchorIdRendering.Abbreviated,
        };
        var projection = WmlToMarkdownConverter.Convert(wml, settings);

        // Find any anchor — extract its full Unid from the underlying AnchorTarget,
        // and the abbreviated form from a token in the markdown.
        var firstTarget = projection.AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h");
        var fullKey = firstTarget.Anchor.Id;
        Assert.True(projection.AnchorIndex.ContainsKey(fullKey),
            $"full key '{fullKey}' missing from index");

        // The matching abbreviated key is somewhere in the markdown — extract any one.
        var abbreviatedToken = System.Text.RegularExpressions.Regex.Match(
            projection.Markdown, @"\{#([^:]+:[^:]+:[a-f0-9]+)\}");
        Assert.True(abbreviatedToken.Success);
        var abbreviatedKey = abbreviatedToken.Groups[1].Value;
        Assert.NotEqual(fullKey, abbreviatedKey);  // it's actually abbreviated, not the full Unid
        Assert.True(projection.AnchorIndex.ContainsKey(abbreviatedKey),
            $"abbreviated key '{abbreviatedKey}' missing from index");

        // Both keys resolve to AnchorTarget whose underlying Unid matches.
        var fromFull = projection.AnchorIndex[fullKey];
        var fromAbbreviated = projection.AnchorIndex[abbreviatedKey];
        Assert.Equal(fromFull.Unid, fromAbbreviated.Unid);
    }

    [Fact]
    public void MD006_BoilerplateFootnotesNotInAnchorIndex()
    {
        // Every DOCX with a FootnotesPart includes two Word-reserved boilerplate
        // footnotes (type="separator" and type="continuationSeparator") used to
        // render the horizontal separator lines above footnote text. They are
        // structural plumbing — not editorial content — and must not appear in
        // the agent-facing AnchorIndex.
        var bytes = DocxSessionTests.BuildDocWithFootnotes();
        var wml = new WmlDocument("test.docx", bytes);

        var projection = WmlToMarkdownConverter.Convert(wml, new WmlToMarkdownConverterSettings());

        // Resolve every fn-kind anchor in the fn scope back to its XElement
        // and assert none is a separator/continuationSeparator.
        using var sm = new OpenXmlMemoryStreamDocument(wml);
        using var doc = sm.GetWordprocessingDocument();
        foreach (var t in projection.AnchorIndex.Values
                     .Where(t => t.Anchor.Scope == "fn" && t.Anchor.Kind == "fn"))
        {
            var el = t.Resolve(doc);
            Assert.NotNull(el);
            var type = (string?)el.Attribute(W.type);
            Assert.False(type is "separator" or "continuationSeparator",
                $"Boilerplate fn (type={type}) leaked into AnchorIndex: {t.Anchor.Id}");
        }

        // And descendant paragraphs of boilerplate notes shouldn't appear either —
        // their scope is "fn" but their kind is "p"/"h"/"li". An empty set is also
        // valid (a doc with no real footnotes), so we just verify no leakage.
        foreach (var t in projection.AnchorIndex.Values
                     .Where(t => t.Anchor.Scope == "fn" && t.Anchor.Kind is "p" or "h" or "li"))
        {
            var el = t.Resolve(doc);
            Assert.NotNull(el);
            var ancestorFn = el.Ancestors(W.footnote).FirstOrDefault();
            Assert.NotNull(ancestorFn);
            var type = (string?)ancestorFn.Attribute(W.type);
            Assert.False(type is "separator" or "continuationSeparator",
                $"Descendant of boilerplate fn leaked into AnchorIndex: {t.Anchor.Id}");
        }
    }
}
