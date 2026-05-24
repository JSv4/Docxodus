#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

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
                for (var i = 1; i <= 6; i++)
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

    [Fact]
    public void MD014_AnchorModeNoneOmitsTokens()
    {
        var doc = BuildSimpleDoc("Hello world");
        var p = WmlToMarkdownConverter.Convert(doc,
            new WmlToMarkdownConverterSettings { AnchorMode = AnchorRenderMode.None });
        Assert.DoesNotContain("{#p:", p.Markdown);
        Assert.Contains("Hello world", p.Markdown);
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
}
