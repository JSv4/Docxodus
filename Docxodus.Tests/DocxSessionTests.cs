#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for <see cref="DocxSession"/>. Test IDs follow the <c>DS###</c> prefix convention.
/// Phase ranges: phase 1 (skeleton) = DS001-DS009, phase 2 (parser) = DS010-DS029,
/// phase 3 (text CRUD) = DS030-DS039, phase 4 (structural) = DS040-DS049,
/// phase 5 (formatting) = DS050-DS059, phase 6 (cell + tracked) = DS060-DS069,
/// phase 7 (raw) = DS070-DS079, phase 8 (WASM/npm) = npm/tests/docx-session.spec.ts.
/// </summary>
public class DocxSessionTests
{
    // ─── In-memory fixture builders ───────────────────────────────────────

    /// <summary>
    /// A simple two-paragraph document with Heading1..Heading6 + Quote + Code style
    /// definitions in the styles part. The styles allow later phases (SetParagraphStyle)
    /// to flip the paragraph kind without rebuilding the fixture.
    /// </summary>
    internal static byte[] BuildDS001_SimpleTwoParagraphs()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = BuildHeadingStyles();

            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();

            body.Append(new Paragraph(new Run(new Text("First paragraph."))));
            body.Append(new Paragraph(new Run(new Text("Second paragraph."))));

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>Single paragraph; styles part defines ONLY Normal (no "Code" style).</summary>
    internal static byte[] BuildDocWithoutCodeStyle()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First paragraph.")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new Style(new StyleName { Val = "Normal" })
                { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// 2×2 table with simple text in each cell.
    /// </summary>
    internal static byte[] BuildDS003_TableWithCells()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var table = new DocumentFormat.OpenXml.Wordprocessing.Table();
            for (int row = 0; row < 2; row++)
            {
                var tr = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                for (int col = 0; col < 2; col++)
                {
                    var tc = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    tc.Append(new Paragraph(new Run(new Text($"R{row}C{col}"))));
                    tr.Append(tc);
                }
                table.Append(tr);
            }
            body.Append(table);
            body.Append(new Paragraph(new Run(new Text("After table."))));

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Two-item bulleted list (nested). Includes a NumberingDefinitionsPart
    /// with a single abstractNum (bullets at all levels) and a numId mapping.
    /// </summary>
    internal static byte[] BuildDS002_BulletedList()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = BuildBulletNumbering();

            body.Append(MakeListItem("Top-level item", level: 0, numId: 1));
            body.Append(MakeListItem("Nested item", level: 1, numId: 1));
            body.Append(MakeListItem("Another top", level: 0, numId: 1));

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Single-item list where the paragraph carries only a <c>pStyle</c> pointing
    /// at a custom style that contributes the <c>numPr</c>. Used to verify
    /// <c>FromStyle = true</c> resolution.
    /// </summary>
    internal static byte[] BuildBM_StyleInheritedList()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var styles = BuildHeadingStyles();
            // Add a paragraph style "MyListStyle" that carries numPr → numId=1, ilvl=0.
            var styleNumPr = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            styles.Append(new Style(
                new StyleName { Val = "My List Style" },
                styleNumPr)
            {
                Type = StyleValues.Paragraph,
                StyleId = "MyListStyle",
            });
            stylesPart.Styles = styles;

            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = BuildBulletNumbering();

            // Paragraph carries only the pStyle — no inline numPr.
            var pPr = new ParagraphProperties(new ParagraphStyleId { Val = "MyListStyle" });
            body.Append(new Paragraph(pPr, new Run(new Text("Style-inherited list item"))));

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Style-inherited bullet list — the numPr lives on the "ListBullet" STYLE (not the
    /// paragraph) and the numbering defines ONLY level 0. Mirrors python-docx's default
    /// "List Bullet" output, the real-world case where nesting was a silent no-op: SetListLevel
    /// found no direct numPr, and even with one the abstractNum lacked the nested level.
    /// </summary>
    internal static byte[] BuildStyleListSingleLevel()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            var styles = BuildHeadingStyles();
            styles.Append(new Style(
                new StyleName { Val = "List Bullet" },
                new ParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })))
            {
                Type = StyleValues.Paragraph,
                StyleId = "ListBullet",
            });
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // Single-level (ilvl 0 only) bullet numbering, marked singleLevel — exactly what
            // python-docx's default "List Bullet" emits, and the case where WmlToHtmlConverter
            // forces ilvl=0 unless the nest upgrades multiLevelType.
            var num = new DocumentFormat.OpenXml.Wordprocessing.Numbering();
            var abs = new AbstractNum(
                new MultiLevelType { Val = MultiLevelValues.SingleLevel })
            { AbstractNumberId = 0 };
            abs.Append(new Level(
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "·" })
            { LevelIndex = 0 });
            num.Append(abs);
            num.Append(new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = num;

            // Paragraphs carry only the pStyle — list membership comes from the style.
            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "ListBullet" }),
                new Run(new Text("First bullet"))));
            body.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "ListBullet" }),
                new Run(new Text("Second bullet"))));

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Single paragraph in a landscape A4 section with non-default margins,
    /// two columns, and a header part reference. Used to verify
    /// <c>GetSectionInfo</c> against a richly-configured sectPr.
    /// </summary>
    internal static byte[] BuildBM_LandscapeSection()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // Add a header part so we can verify HeaderPartUris is populated.
            var headerPart = main.AddNewPart<HeaderPart>("rIdH1");
            headerPart.Header = new Header(new Paragraph(new Run(new Text("Header text"))));

            body.Append(new Paragraph(new Run(new Text("Page body"))));

            // Section properties: A4 landscape (16838 x 11906 twips), columns=2, margins set.
            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdH1" },
                new PageSize { Width = 16838u, Height = 11906u, Orient = PageOrientationValues.Landscape },
                new PageMargin { Top = 720, Bottom = 720, Left = 1080, Right = 1080,
                    Header = 0, Footer = 0, Gutter = 0 },
                new Columns { ColumnCount = 2 });
            body.Append(sectPr);

            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static Paragraph MakeListItem(string text, int level, int numId)
    {
        var pPr = new ParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = numId }));
        return new Paragraph(pPr, new Run(new Text(text)));
    }

    private static DocumentFormat.OpenXml.Wordprocessing.Numbering BuildBulletNumbering()
    {
        var n = new DocumentFormat.OpenXml.Wordprocessing.Numbering();
        var abs = new AbstractNum { AbstractNumberId = 0 };
        for (int i = 0; i < 9; i++)
        {
            abs.Append(new Level(
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "·" })
            {
                LevelIndex = i,
            });
        }
        n.Append(abs);
        n.Append(new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
        return n;
    }

    internal static Styles BuildHeadingStyles()
    {
        var styles = new Styles();
        for (int i = 1; i <= 6; i++)
        {
            styles.Append(new Style(
                new StyleName { Val = $"Heading {i}" })
            {
                Type = StyleValues.Paragraph,
                StyleId = $"Heading{i}",
            });
        }
        styles.Append(new Style(new StyleName { Val = "Quote" })
        {
            Type = StyleValues.Paragraph,
            StyleId = "Quote",
        });
        styles.Append(new Style(new StyleName { Val = "Code" })
        {
            Type = StyleValues.Paragraph,
            StyleId = "Code",
        });
        return styles;
    }

    /// <summary>
    /// Document with a FootnotesPart containing the two Word-reserved boilerplate
    /// footnotes (type="separator" and type="continuationSeparator") plus one
    /// user-authored footnote. Used to verify the AnchorIndex filters out the
    /// boilerplate notes.
    /// </summary>
    internal static byte[] BuildDocWithFootnotes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Body text with a footnote reference."),
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 }))));

            var fnPart = main.AddNewPart<DocumentFormat.OpenXml.Packaging.FootnotesPart>();
            var fnXml = """
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="1"><w:p><w:r><w:t>User-authored footnote text.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """;
            using (var s = fnPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s)) w.Write(fnXml);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Five-paragraph document covering the placeholder shapes that <see cref="DocxSession.FindPlaceholders"/>
    /// recognizes — used by the template-fill convenience tests (Issue #163, units A/B/C).
    /// Includes a leading-prefix case (<c>$[___]</c>) so tests can verify that
    /// <see cref="DocxSession.ReplaceInner"/> preserves text outside the brackets.
    /// </summary>
    internal static byte[] BuildDocWithBracketPlaceholders()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("The name of this corporation is [_____]."))),
                new Paragraph(new Run(new Text("The price per share is $[___]."))),
                new Paragraph(new Run(new Text("Located at [____________]."))),
                new Paragraph(new Run(new Text("[Notwithstanding the foregoing, additional terms may apply.]"))),
                new Paragraph(new Run(new Text("[outer [inner] clause]")))));
        }
        return ms.ToArray();
    }

    internal static byte[] BuildDocWithLongClauseBlank()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(
                    "[that would result in at least $_______ in gross proceeds, including or excluding proceeds previously received]")))));
        }
        return ms.ToArray();
    }

    internal static byte[] BuildDocWithAdjacentPlaceholders()
    {
        // One paragraph, three blanks back-to-back like the NVCA address line —
        // the canonical case where 40-char context windows pull in the next blank's
        // text and confuse keyword-based picker logic.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(
                    "The address is [STREET], in the City of [CITY], County of [COUNTY], in the State of Delaware.")))));
        }
        return ms.ToArray();
    }

    internal static byte[] BuildDocWithSentences()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(
                    "This is the first sentence. Here is some text with a [FILL] placeholder. And a third sentence after.")))));
        }
        return ms.ToArray();
    }

    internal static byte[] BuildDocWithCommaSeparatedItems()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(
                    "Item one is [A], item two is [B], and item three is [C].")))));
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Builds the NVCA Model COI repro for issue #188 — an optional outer
    /// wrapper around a name slot that a picker drops by returning <c>""</c>,
    /// plus a few simpler shapes covering each <c>CoalesceWhitespaceAroundEmptyFill</c>
    /// case the option promises:
    ///   <list type="bullet">
    ///     <item>Para 1: <c>"… on March 14, 2024 [under the name [_______________]]."</c>
    ///       after the inner name is filled — outer drop should yield <c>"… 2024."</c>.</item>
    ///     <item>Para 2: <c>"alpha [opt] beta"</c> — both-sides-whitespace case.</item>
    ///     <item>Para 3: <c>"[[opt]]"</c> — bracket-on-both-sides exact-wrap case.</item>
    ///     <item>Para 4: <c>"[front]edge"</c> — no neighbor pattern (no coalesce).</item>
    ///   </list>
    /// </summary>
    internal static byte[] BuildDocWithEmptyFillNeighbors()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(
                    "Filed pursuant to the General Corporation Law on March 14, 2024 [under the name [_______________]]."))),
                new Paragraph(new Run(new Text(
                    "alpha [opt] beta"))),
                new Paragraph(new Run(new Text(
                    "[[opt]]"))),
                new Paragraph(new Run(new Text(
                    "[front]edge")))));
        }
        return ms.ToArray();
    }

    // ─── Phase 1: Skeleton tests ─────────────────────────────────────────

    [Fact]
    public void DS001_OpenAndProject()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var projection = session.Project();
        Assert.Contains("First paragraph.", projection.Markdown);
        Assert.Contains("Second paragraph.", projection.Markdown);
        Assert.True(projection.AnchorIndex.Count >= 2);
    }

    [Fact]
    public void DS002_SaveRoundtrip()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var out1 = session.Save();
        Assert.NotEmpty(out1);

        using var session2 = new DocxSession(out1);
        Assert.Contains("First paragraph.", session2.Project().Markdown);
    }

    [Fact]
    public void DS003_ExistsAndGetAnchorInfo()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var proj = session.Project();

        var firstAnchor = proj.AnchorIndex.Keys.First();
        Assert.True(session.Exists(firstAnchor));
        Assert.False(session.Exists("p:body:deadbeefdeadbeefdeadbeefdeadbeef"));

        var info = session.GetAnchorInfo(firstAnchor);
        Assert.NotNull(info);
        Assert.Contains(info!.Kind, new[] { "p", "h", "li" });
        Assert.False(string.IsNullOrEmpty(info.TextPreview));
    }

    [Fact]
    public void DS004_DisposeDoubleOk()
    {
        var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        session.Dispose();
        session.Dispose();
    }

    [Fact]
    public void DS005_ProjectionCached()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var p1 = session.Project();
        var p2 = session.Project();
        Assert.Same(p1, p2);
    }

    [Fact]
    public void DS220_GetAnchorInfosBulkLookup()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var projection = session.Project();

        // Collect every body paragraph anchor id.
        var ids = projection.AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
            .Select(t => t.Anchor.Id)
            .ToList();
        Assert.NotEmpty(ids);

        // Bulk lookup returns the same data as N individual GetAnchorInfo calls,
        // and an unknown anchor id maps to null in the result.
        var idsWithBogus = ids.Concat(new[] { "p:body:0000000000000000ffffffffffffffff" }).ToList();
        var bulk = session.GetAnchorInfos(idsWithBogus);

        Assert.Equal(idsWithBogus.Count, bulk.Count);
        foreach (var id in ids)
        {
            Assert.True(bulk.ContainsKey(id));
            var info = bulk[id];
            Assert.NotNull(info);
            Assert.Equal(id, info!.Id);

            // Verify it agrees with the existing per-anchor API.
            var singleton = session.GetAnchorInfo(id);
            Assert.NotNull(singleton);
            Assert.Equal(singleton!.TextPreview, info.TextPreview);
            Assert.Equal(singleton.Kind, info.Kind);
            Assert.Equal(singleton.Scope, info.Scope);
        }
        Assert.Null(bulk["p:body:0000000000000000ffffffffffffffff"]);
    }

    [Fact]
    public void DS220b_GetAnchorInfos_DedupesAndSkipsNullIds()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var projection = session.Project();
        var realId = projection.AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
            .Anchor.Id;

        // Dedup: passing the same real id twice yields one entry.
        var dupes = session.GetAnchorInfos(new[] { realId, realId, "p:body:unknown-id-xyz" });
        Assert.Equal(2, dupes.Count);
        Assert.True(dupes.ContainsKey(realId));
        Assert.Null(dupes["p:body:unknown-id-xyz"]);

        // Null-string skip: null strings in the input enumerable are silently skipped.
        IEnumerable<string> withNulls = new[] { realId, null!, realId };
        var bulkWithNulls = session.GetAnchorInfos(withNulls);
        Assert.Single(bulkWithNulls);
        Assert.True(bulkWithNulls.ContainsKey(realId));
    }

    [Fact]
    public void DS221_GetAnchorInfoUsesAnchorTargetTextPreview()
    {
        // Regression: after #162, GetAnchorInfo reads target.TextPreview directly
        // instead of re-walking the element. Confirm the value matches what the
        // projection put on AnchorTarget.
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var projection = session.Project();
        foreach (var t in projection.AnchorIndex.Values
                      .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li"))
        {
            var info = session.GetAnchorInfo(t.Anchor.Id);
            Assert.NotNull(info);
            Assert.Equal(t.TextPreview, info!.TextPreview);
        }
    }

    /// <summary>
    /// Build a one-paragraph doc whose only paragraph has a numbered Heading2 style.
    /// The numbering renders as "1." so callers can verify AutoNumberPrefix resolution.
    /// </summary>
    internal static byte[] BuildNumberedHeadingDoc(string text = "The total number of shares")
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Heading2", StyleName = new StyleName { Val = "heading 2" } });
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

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

            var pPr = new ParagraphProperties(
                new ParagraphStyleId { Val = "Heading2" },
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            body.Append(new Paragraph(pPr, new Run(new Text(text))));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS222_AnchorTarget_AutoNumberPrefix_ResolvedForNumberedHeading()
    {
        using var session = new DocxSession(BuildNumberedHeadingDoc());
        var projection = session.Project();
        var target = projection.AnchorIndex.Values
            .Single(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "h");
        Assert.Equal("1.", target.AutoNumberPrefix);
        Assert.Equal("The total number of shares", target.TextPreview);
        Assert.Equal("1. The total number of shares", target.FullText);
    }

    [Fact]
    public void DS222a_AnchorInfo_CarriesAutoNumberPrefix()
    {
        // GetAnchorInfo() must surface the resolved prefix all the way to the
        // public API so callers don't need to walk AnchorIndex themselves.
        using var session = new DocxSession(BuildNumberedHeadingDoc("My heading"));
        var projection = session.Project();
        var target = projection.AnchorIndex.Values
            .Single(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "h");
        var info = session.GetAnchorInfo(target.Anchor.Id);
        Assert.NotNull(info);
        Assert.Equal("1.", info!.AutoNumberPrefix);
        Assert.Equal("1. My heading", info.FullText);
    }

    [Fact]
    public void DS222b_AnchorTarget_AutoNumberPrefix_NullForUnnumberedParagraph()
    {
        // A plain paragraph without w:numPr has no prefix and FullText == TextPreview.
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var projection = session.Project();
        foreach (var t in projection.AnchorIndex.Values
                      .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p"))
        {
            Assert.Null(t.AutoNumberPrefix);
            Assert.Equal(t.TextPreview, t.FullText);
        }
    }

    [Fact]
    public void DS230_ReplaceInner_StripsBracketsKeepsPrefix()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());

        // The "$[___]" placeholder includes the leading "$" in the match (regex \$?\[…\]).
        // ReplaceInner must keep the "$" prefix and substitute only the bracketed portion.
        var dollarMatch = session.FindPlaceholders(PlaceholderKinds.BlankFill)
            .Single(p => p.Match.Text.StartsWith("$["));

        var r = session.ReplaceInner(dollarMatch.Match, "0.20");
        Assert.True(r.Success);

        // After replacement, the paragraph should read "The price per share is $0.20."
        var anchorId = dollarMatch.Match.EnclosingAnchor.Anchor.Id;
        var info = session.GetAnchorInfo(anchorId);
        Assert.NotNull(info);
        Assert.Equal("The price per share is $0.20.", info!.TextPreview);
    }

    [Fact]
    public void DS231_ReplaceInner_NoBracketsReturnsError()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());

        // Pick any match, then hand-craft a TextMatch whose Text has no brackets.
        var p = session.FindPlaceholders(PlaceholderKinds.BlankFill).First();
        var bogus = new TextMatch
        {
            Text = "no brackets here",
            EnclosingAnchor = p.Match.EnclosingAnchor,
            Span = p.Match.Span,
            Fragments = p.Match.Fragments,
            ContextBefore = "",
            ContextAfter = "",
            Groups = Array.Empty<string>(),
        };

        var r = session.ReplaceInner(bogus, "anything");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.MalformedMarkdown, r.Error?.Code);
    }

    [Fact]
    public void DS232_Classifier_LongClauseWithUnderscoresExposesAlternativeKinds()
    {
        // Long bracketed clauses that contain embedded underscores are ambiguous —
        // strictly speaking they're an AlternativeClause (a real clause), but the
        // current classifier rule fires BlankFill because >= 2 underscores are present.
        // Surface BOTH classifications so callers can decide.
        using var session = new DocxSession(BuildDocWithLongClauseBlank());

        var p = session.FindPlaceholders().Single();
        Assert.Equal(PlaceholderKind.BlankFill, p.Kind);              // primary stays BlankFill (back-compat)
        Assert.Contains(PlaceholderKind.AlternativeClause, p.AlternativeKinds);
    }

    [Fact]
    public void DS233_Classifier_SimpleBlankFillEmptyAlternativeKinds()
    {
        // Plain "[_____]" placeholders are unambiguous — AlternativeKinds should be empty.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var simpleBlankFill = session.FindPlaceholders(PlaceholderKinds.BlankFill)
            .Single(p => p.Match.Text == "[_____]");
        Assert.Empty(simpleBlankFill.AlternativeKinds);
    }

    [Fact]
    public void DS240_FillPlaceholders_BlankFillPickerReplacesValue()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var result = session.FillPlaceholders(p =>
            p.Kind == PlaceholderKind.BlankFill && p.Match.Text == "[_____]"
                ? "ACME, INC."
                : null);
        Assert.True(result.Filled >= 1);
        Assert.Empty(result.Errors);

        // Verify the first paragraph now contains "ACME, INC."
        var projection = session.Project();
        var firstPara = projection.AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h");
        Assert.Contains("ACME, INC.", firstPara.TextPreview);
    }

    [Fact]
    public void DS241_FillPlaceholders_PickerReturningNullSkips()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var result = session.FillPlaceholders(p => null);   // skip everything
        Assert.Equal(0, result.Filled);
        Assert.True(result.Skipped > 0);
        Assert.NotEmpty(result.Unfilled);
    }

    [Fact]
    public void DS242_FillPlaceholders_PreservesDollarPrefix()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var result = session.FillPlaceholders(p =>
            p.Match.Text.Contains("$[___]") ? "0.20" : null);
        Assert.Equal(1, result.Filled);

        // The dollar-prefixed paragraph should now read "$0.20", not "0.20" or "$$0.20".
        var projection = session.Project();
        var allMd = projection.Markdown;
        Assert.Contains("$0.20", allMd);
        Assert.DoesNotContain("$$0.20", allMd);
    }

    [Fact]
    public void DS243_FillPlaceholders_DollarPrefixDisabled()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var options = new FillOptions { PreserveDollarPrefix = false };
        var result = session.FillPlaceholders(
            p => p.Match.Text.Contains("$[___]") ? "0.20" : null,
            options);
        Assert.Equal(1, result.Filled);

        // With PreserveDollarPrefix=false, the "$" is overwritten by the replacement.
        var projection = session.Project();
        Assert.Contains("share is 0.20.", projection.Markdown);
        Assert.DoesNotContain("$0.20", projection.Markdown);
    }

    [Fact]
    public void DS244a_FillPlaceholders_DefaultKindsVisitsAlternativeClauses()
    {
        // After the FillOptions.Kinds default change, the picker should be invoked
        // for AlternativeClause placeholders too — without the caller needing to
        // opt in via Kinds = All. Picker that unwraps every bracketed match should
        // strip every alternative in one pass over BuildDocWithBracketPlaceholders.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var result = session.FillPlaceholders(p => p.Kind == PlaceholderKind.AlternativeClause
            ? p.Match.Text.Trim('[', ']')
            : null);
        Assert.True(result.Filled >= 1, "Expected at least one AlternativeClause unwrapped");

        // No AlternativeClause matches should remain after convergence.
        var leftover = session.FindPlaceholders(PlaceholderKinds.AlternativeClause);
        Assert.Empty(leftover);
    }

    [Fact]
    public void DS244_FillPlaceholders_AlternativeClauseMultiPassStripsNestedBrackets()
    {
        // The "[outer [inner] clause]" placeholder is nested; FindPlaceholders returns
        // INNERMOST first, so stripping has to iterate. FillPlaceholders should converge
        // after multiple passes when picker strips brackets uniformly.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var options = new FillOptions { Kinds = PlaceholderKinds.AlternativeClause };
        var result = session.FillPlaceholders(p =>
        {
            // Strip [ and ] from the match's bracketed portion, preserve any prefix/suffix.
            var t = p.Match.Text;
            int lb = t.IndexOf('['), rb = t.LastIndexOf(']');
            if (lb < 0 || rb <= lb) return null;
            return t[..lb] + t[(lb + 1)..rb] + t[(rb + 1)..];
        }, options);

        Assert.True(result.Passes >= 2);

        // After convergence, no AlternativeClause matches remain.
        var leftover = session.FindPlaceholders(PlaceholderKinds.AlternativeClause);
        Assert.Empty(leftover);
    }

    [Fact]
    public void DS245_FillPlaceholders_MaxPassesRejectsZeroOrNegative()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            session.FillPlaceholders(_ => null, new FillOptions { MaxPasses = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            session.FillPlaceholders(_ => null, new FillOptions { MaxPasses = -1 }));
    }

    [Fact]
    public void DS246_FillPlaceholders_PassesReflectsActualWorkDone()
    {
        // One paragraph, one [_____] placeholder. Picker fills it in pass 1.
        // After pass 1's edit, pass 2's FindPlaceholders returns empty and breaks
        // before recording any work. Passes should be 1, not 2.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var result = session.FillPlaceholders(p =>
            p.Match.Text == "[_____]" ? "FILLED" : null);
        Assert.True(result.Filled >= 1);
        Assert.Equal(1, result.Passes);
    }

    [Fact]
    public void DS247_FillPlaceholders_StillPresentZeroWhenMultiPassFinishesSkippedPlaceholders()
    {
        // Issue #189: Skipped is a misleading "is the template done?" signal
        // because the picker can return null for a nested-outer wrapper in
        // pass 1, then succeed in pass 2 once the inner brackets are stripped.
        // The fixture's "[outer [inner] clause]" is the canonical case:
        //   pass 1: picker sees "[inner]" only (innermost-first) → strips it
        //   pass 2: picker now sees "[outer  clause]" → strips it too
        // BUT the picker below also evaluates a non-bracket placeholder it
        // can't recognize in pass 1, forcing a skip. Multi-pass should still
        // finish all the bracketed work, so StillPresent must be 0 even
        // though Skipped > 0.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var options = new FillOptions { Kinds = PlaceholderKinds.AlternativeClause };
        bool firstNestedSeen = false;
        var result = session.FillPlaceholders(p =>
        {
            var t = p.Match.Text;
            // Force a first-pass skip the first time we see the innermost
            // "[inner]" — but every other call (including the second visit
            // on the surfaced outer) strips brackets normally.
            if (t == "[inner]" && !firstNestedSeen)
            {
                firstNestedSeen = true;
                return null;
            }
            int lb = t.IndexOf('['), rb = t.LastIndexOf(']');
            if (lb < 0 || rb <= lb) return null;
            return t[..lb] + t[(lb + 1)..rb] + t[(rb + 1)..];
        }, options);

        // The first-pass skip of "[inner]" produced Skipped > 0…
        Assert.True(result.Skipped > 0,
            "Expected the contrived first-pass null to count as a skip.");
        // …but multi-pass convergence still emptied every AlternativeClause,
        // so StillPresent must report the trustworthy "template is done" 0.
        Assert.Equal(0, result.StillPresent);
        Assert.Empty(session.FindPlaceholders(PlaceholderKinds.AlternativeClause));
    }

    [Fact]
    public void DS248_FillPlaceholders_StillPresentCountsGenuinelyUnfilled()
    {
        // When the picker truly returns null for everything, every visited
        // placeholder remains in the doc. StillPresent must equal the count
        // of placeholders in the requested kinds/scope after the loop exits.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var options = new FillOptions { Kinds = PlaceholderKinds.BlankFill };
        var result = session.FillPlaceholders(_ => null, options);

        Assert.Equal(0, result.Filled);
        var leftover = session.FindPlaceholders(PlaceholderKinds.BlankFill);
        Assert.Equal(leftover.Count, result.StillPresent);
        Assert.True(result.StillPresent > 0);
    }

    // ─── Issue #188: CoalesceWhitespaceAroundEmptyFill ───────────────────

    [Fact]
    public void DS247a_FillPlaceholders_EmptyPickReturn_LeavesArtifactsByDefault()
    {
        // Baseline: the current (default) behavior preserves surrounding chars
        // verbatim, which is what issue #188 reports as a cosmetic problem.
        using var session = new DocxSession(BuildDocWithEmptyFillNeighbors());
        var result = session.FillPlaceholders(p => p.Kind switch
        {
            PlaceholderKind.BlankFill => "Bluth, Inc.",          // fill inner name slot
            PlaceholderKind.AlternativeClause => string.Empty,    // drop every bracketed wrapper
            _ => null,
        });
        Assert.True(result.Filled >= 1);

        var md = session.Project().Markdown;
        // Pass 1 fills [_______________] → "Bluth, Inc."; pass 2 sees the now-innermost
        // [under the name Bluth, Inc.] and returns "" for it. Without coalescing, the
        // doubled space before "." survives. Match the cosmetic artifact exactly.
        Assert.Contains("on March 14, 2024 .", md);
    }

    [Fact]
    public void DS247b_FillPlaceholders_CoalesceCollapsesSpaceBeforePunctuation()
    {
        // The headline issue #188 case: outer wrapper drops, the leading space
        // between the prior word and the dropped bracket is absorbed so no stray
        // space remains before the period.
        using var session = new DocxSession(BuildDocWithEmptyFillNeighbors());
        var options = new FillOptions { CoalesceWhitespaceAroundEmptyFill = true };
        var result = session.FillPlaceholders(p => p.Kind switch
        {
            PlaceholderKind.BlankFill => "Bluth, Inc.",
            PlaceholderKind.AlternativeClause => string.Empty,
            _ => null,
        }, options);
        Assert.True(result.Filled >= 1);

        var md = session.Project().Markdown;
        Assert.Contains("on March 14, 2024.", md);
        Assert.DoesNotContain("on March 14, 2024 .", md);
        Assert.DoesNotContain("on March 14, 2024  .", md);
    }

    [Fact]
    public void DS247c_FillPlaceholders_CoalesceCollapsesSurroundingSpaces()
    {
        // "alpha [opt] beta" → drop [opt] with coalescing → "alpha beta" (one space).
        using var session = new DocxSession(BuildDocWithEmptyFillNeighbors());
        var options = new FillOptions { CoalesceWhitespaceAroundEmptyFill = true };
        session.FillPlaceholders(p =>
            p.Kind == PlaceholderKind.AlternativeClause ? string.Empty : null, options);

        var md = session.Project().Markdown;
        Assert.Contains("alpha beta", md);
        Assert.DoesNotContain("alpha  beta", md);
    }

    [Fact]
    public void DS247d_FillPlaceholders_CoalesceDropsSurroundingBrackets()
    {
        // "[[opt]]" — innermost match is "[opt]" at offset 1; the chars at offsets
        // 0 ("[") and 6 ("]") are an exact bracket wrap. Returning "" with coalescing
        // should absorb the outer pair too, so the paragraph ends up empty rather
        // than leaving "[]" stragglers.
        using var session = new DocxSession(BuildDocWithEmptyFillNeighbors());
        var options = new FillOptions { CoalesceWhitespaceAroundEmptyFill = true };
        session.FillPlaceholders(p =>
            p.Kind == PlaceholderKind.AlternativeClause ? string.Empty : null, options);

        var md = session.Project().Markdown;
        Assert.DoesNotContain("[opt]", md);
        Assert.DoesNotContain("[]", md);
    }

    [Fact]
    public void DS247e_FillPlaceholders_CoalesceLeavesEdgeMatchesAlone()
    {
        // "[front]edge" — placeholder at offset 0 of its paragraph; no left
        // neighbor. The right neighbor 'e' isn't whitespace/punct/bracket either.
        // No coalescing pattern applies → literal delete → "edge".
        using var session = new DocxSession(BuildDocWithEmptyFillNeighbors());
        var options = new FillOptions { CoalesceWhitespaceAroundEmptyFill = true };
        session.FillPlaceholders(p =>
            p.Kind == PlaceholderKind.AlternativeClause ? string.Empty : null, options);

        var md = session.Project().Markdown;
        Assert.Contains("edge", md);
        Assert.DoesNotContain("[front]edge", md);
    }

    [Fact]
    public void DS247f_FillPlaceholders_CoalesceIsSkippedWhenPreserveDollarPrefixKeepsTheDollar()
    {
        // $-prefix preservation runs first: if picker returns "" for "$[xxx]" with
        // the default PreserveDollarPrefix = true, the replacement becomes "$",
        // which is no longer empty, so coalescing should not fire. Verify the "$"
        // stays put (and surrounding chars are untouched).
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Pay $[___] now.")))));
        }
        using var session = new DocxSession(ms.ToArray());
        var options = new FillOptions { CoalesceWhitespaceAroundEmptyFill = true };
        session.FillPlaceholders(p => string.Empty, options);

        var md = session.Project().Markdown;
        Assert.Contains("Pay $ now.", md);   // $ survives, surrounding spaces untouched
    }

    // ─── Issue #164: boundary-aware Grep context windows ─────────────────

    [Fact]
    public void DS250_Grep_DefaultContextCharsIsEighty()
    {
        using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());
        var matches = session.Grep(@"\[STREET\]");
        var m = Assert.Single(matches);

        // "in the City of [CITY], County of [COUNTY], in the State of Delaware."
        // is 68 chars. Pre-#164 was capped at 40. Post-#164 default 80 → all 68 visible.
        Assert.True(m.ContextAfter.Length > 40,
            $"ContextAfter should exceed 40 chars under widened default; got {m.ContextAfter.Length}: '{m.ContextAfter}'");
    }

    [Fact]
    public void DS251_Grep_BoundaryCharMatchesPreviousBehavior()
    {
        using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());
        var matches = session.Grep(@"\[CITY\]",
            scope: ProjectionScopes.Body, contextChars: 40, boundary: ContextBoundary.Char);
        var m = Assert.Single(matches);

        Assert.True(m.ContextBefore.Length <= 40);
        Assert.True(m.ContextAfter.Length <= 40);
    }

    [Fact]
    public void DS252_Grep_BoundaryBracketStopsAtNearestBracket()
    {
        using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());
        var matches = session.Grep(@"\[CITY\]",
            scope: ProjectionScopes.Body, contextChars: 80, boundary: ContextBoundary.Bracket);
        var m = Assert.Single(matches);

        Assert.DoesNotContain('[', m.ContextBefore);
        Assert.DoesNotContain(']', m.ContextBefore);
        Assert.DoesNotContain('[', m.ContextAfter);
        Assert.DoesNotContain(']', m.ContextAfter);

        Assert.Equal(", in the City of ", m.ContextBefore);
        Assert.Equal(", County of ", m.ContextAfter);
    }

    [Fact]
    public void DS253_Grep_BoundarySentenceStopsAtTerminator()
    {
        // Sentence mode stops at . ! ? : ;
        // Multi-sentence paragraph; assert ContextBefore stops at the previous '.'.
        using var session = new DocxSession(BuildDocWithSentences());
        var matches = session.Grep(@"\[FILL\]",
            scope: ProjectionScopes.Body, contextChars: 200, boundary: ContextBoundary.Sentence);
        var m = Assert.Single(matches);

        Assert.DoesNotContain('.', m.ContextBefore);
        Assert.DoesNotContain('.', m.ContextAfter);
        Assert.DoesNotContain(';', m.ContextBefore);
        Assert.DoesNotContain(';', m.ContextAfter);
    }

    [Fact]
    public void DS254_Grep_BoundaryCommaStopsAtComma()
    {
        // Comma mode for matches inside enumerations.
        using var session = new DocxSession(BuildDocWithCommaSeparatedItems());
        var matches = session.Grep(@"\[B\]",
            scope: ProjectionScopes.Body, contextChars: 200, boundary: ContextBoundary.Comma);
        var m = Assert.Single(matches);

        Assert.DoesNotContain(',', m.ContextBefore);
        Assert.DoesNotContain(',', m.ContextAfter);
    }

    [Fact]
    public void DS255_FindPlaceholders_AcceptsBoundaryAndContextChars()
    {
        using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());

        // FindPlaceholders should plumb the boundary parameter through to the
        // internal Grep call so picker code can rely on bracket-bounded context
        // without dropping down to Grep manually.
        var placeholders = session.FindPlaceholders(
            PlaceholderKinds.All,
            ProjectionScopes.Body,
            contextChars: 80,
            boundary: ContextBoundary.Bracket);

        // We have three placeholders ([STREET], [CITY], [COUNTY]); none of their
        // contexts should contain bracket characters from sibling placeholders.
        Assert.Equal(3, placeholders.Count);
        foreach (var p in placeholders)
        {
            Assert.DoesNotContain('[', p.Match.ContextBefore);
            Assert.DoesNotContain(']', p.Match.ContextBefore);
            Assert.DoesNotContain('[', p.Match.ContextAfter);
            Assert.DoesNotContain(']', p.Match.ContextAfter);
        }
    }

    // ─── Phase 3: text CRUD + undo/redo ──────────────────────────────────

    [Fact]
    public void DS030_ReplaceTextSimple()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var firstAnchor = session.Project().AnchorIndex.Keys.First();

        var result = session.ReplaceText(firstAnchor, "Replaced text.");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Contains(result.Modified, a => a.Id == firstAnchor);
        Assert.NotNull(result.Patch);
        Assert.Contains("Replaced text.", result.Patch!.Markdown);

        Assert.Contains("Replaced text.", session.Project().Markdown);
        Assert.DoesNotContain("First paragraph.", session.Project().Markdown);
    }

    [Fact]
    public void DS031_ReplaceText_AnchorNotFound()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var r = s.ReplaceText("p:body:deadbeef", "x");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, r.Error!.Code);
    }

    [Fact]
    public void DS032_ReplaceText_MalformedMarkdownNull()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ReplaceText(anchor, null!);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.MalformedMarkdown, r.Error!.Code);
    }

    [Fact]
    public void DS033_ReplaceText_RejectsTableSyntax()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ReplaceText(anchor, "| a | b |\n|---|---|\n| 1 | 2 |");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.TableInsertNotSupported, r.Error!.Code);
    }

    [Fact]
    public void DS034_ReplaceText_FailureLeavesDocUnchanged()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var before = s.Project().Markdown;
        s.ReplaceText("p:body:deadbeef", "x");
        Assert.Equal(before, s.Project().Markdown);
    }

    [Fact]
    public void DS035_DeleteBlock()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchors = s.Project().AnchorIndex.Keys.ToList();
        Assert.True(anchors.Count >= 2);
        var toDelete = anchors[0];

        var r = s.DeleteBlock(toDelete);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Removed, a => a.Id == toDelete);
        Assert.False(s.Exists(toDelete));
        Assert.DoesNotContain("First paragraph.", s.Project().Markdown);
        Assert.Contains("Second paragraph.", s.Project().Markdown);
    }

    [Fact]
    public void DS036_UndoReplaceText()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var before = s.Project().Markdown;
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Replaced.");
        Assert.True(s.Undo());
        Assert.Equal(before, s.Project().Markdown);
    }

    [Fact]
    public void DS037_RedoAfterUndo()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Replaced.");
        var afterEdit = s.Project().Markdown;
        s.Undo();
        Assert.True(s.Redo());
        Assert.Equal(afterEdit, s.Project().Markdown);
    }

    [Fact]
    public void DS038_NothingToUndo()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        Assert.False(s.Undo());
    }

    [Fact]
    public void DS039_ReplaceText_WithHyperlink()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ReplaceText(anchor, "See [Docxodus](https://example.com/d).");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("[Docxodus](https://example.com/d)", s.Project().Markdown);
    }

    // ─── Phase 4: structural ops ─────────────────────────────────────────

    [Fact]
    public void DS040_InsertParagraphAfter()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.InsertParagraph(anchor, Position.After, "Inserted paragraph.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Single(r.Created);
        var newAnchor = r.Created[0];
        Assert.Equal("p", newAnchor.Kind);
        Assert.Contains("Inserted paragraph.", s.Project().Markdown);
    }

    [Fact]
    public void DS041_InsertParagraphBefore()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.InsertParagraph(anchor, Position.Before, "First inserted.");
        Assert.True(r.Success);
        Assert.Contains("First inserted.", s.Project().Markdown);
        // The inserted paragraph should appear before the original first paragraph
        var md = s.Project().Markdown;
        Assert.True(md.IndexOf("First inserted.") < md.IndexOf("First paragraph."));
    }

    [Fact]
    public void DS041b_InsertMultiBlockPayload()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.InsertParagraph(anchor, Position.After,
            "# New Heading\n\nA normal paragraph beneath it.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count);
        Assert.Equal("h", r.Created[0].Kind);
        Assert.Equal("p", r.Created[1].Kind);
        Assert.Contains("# New Heading", s.Project().Markdown);
    }

    [Fact]
    public void DS042_SplitParagraph()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        // "First paragraph." → split at offset 5 ("First" | " paragraph.")
        var r = s.SplitParagraph(anchor, 5);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Id == anchor);
        Assert.Single(r.Created);

        var md = s.Project().Markdown;
        Assert.Contains("First", md);
        Assert.Contains("paragraph.", md);
    }

    [Fact]
    public void DS042b_SplitParagraph_OffsetOutOfRange()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.SplitParagraph(anchor, 9999);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, r.Error!.Code);
    }

    [Fact]
    public void DS046_SplitAfterStyledParagraph_AppliesNextStyle()
    {
        // Pressing Enter at the END of a Title/Heading should start a Normal body paragraph (the
        // style's linked w:next), not another Title — matching Word. Regression for the editor
        // "draft from scratch" flow where every block below the letterhead inherited the Title.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var titleId = s.SetParagraphStyle(anchor, "Title").Modified[0].Id;

        // "First paragraph." is 16 chars; split at the end → empty new paragraph.
        var r = s.SplitParagraph(titleId, 16);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("Title", s.GetBlockMetadata(r.Modified[0].Id)?.StyleId);   // original keeps Title
        Assert.Equal("Normal", s.GetBlockMetadata(r.Created[0].Id)?.StyleId);   // new para = next style
    }

    [Fact]
    public void DS046b_SplitStyledParagraphMidText_KeepsStyle()
    {
        // A mid-text split is a continuation of the SAME paragraph, so both halves keep the style
        // (the next-style rebase only applies to an empty Enter-at-end split).
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var hId = s.SetParagraphStyle(anchor, "Heading2").Modified[0].Id;

        var r = s.SplitParagraph(hId, 5); // "First" | " paragraph."
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("Heading2", s.GetBlockMetadata(r.Modified[0].Id)?.StyleId);
        Assert.Equal("Heading2", s.GetBlockMetadata(r.Created[0].Id)?.StyleId);
    }

    [Fact]
    public void DS047_SplitDoesNotPropagatePageBreakBefore()
    {
        // pageBreakBefore is a once-only property: the original paragraph keeps it, the new
        // paragraph created by Enter must not inherit a second page break. Regression for the
        // editor flow where a page-broken "Enclosures" heading pushed every list item onto its own page.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.SetParagraphFormat(anchor, new ParagraphFormatOp { PageBreakBefore = true });

        var r = s.SplitParagraph(anchor, 16);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("pageBreakBefore", s.Raw.GetXml(r.Modified[0].Id));      // original keeps it
        Assert.DoesNotContain("pageBreakBefore", s.Raw.GetXml(r.Created[0].Id)); // new para does not
    }

    [Fact]
    public void DS048_SplitListItemContinuesList()
    {
        // Splitting a list item at its end must produce a continuing list item (same numbering),
        // not a plain Normal paragraph — the next-style rebase must skip list items so the editor's
        // Enter-continuation keeps working.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var p = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
        var li = s.ApplyListFormat(p, ListFormat.Decimal).Modified[0].Id;

        var r = s.SplitParagraph(li, 16);
        Assert.True(r.Success, r.Error?.Message);
        var lm = s.GetListMembership(r.Created[0].Id);
        Assert.NotNull(lm);
        Assert.Equal(NumberFormat.Decimal, lm!.Format);
    }

    [Fact]
    public void DS049_SplitReMintsClonedPropertyUnids()
    {
        // The new paragraph clones the source pPr (to carry indent/list membership). Each cloned
        // property must get a FRESH Unid — otherwise the same Unid lands on two elements. Regression
        // for duplicate property Unids observed across split-created paragraphs.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.SetParagraphFormat(anchor, new ParagraphFormatOp { IndentDelta = 720 });

        var r = s.SplitParagraph(anchor, 5); // mid-split clones the w:ind into the new paragraph
        Assert.True(r.Success, r.Error?.Message);
        var firstUnids = System.Text.RegularExpressions.Regex
            .Matches(s.Raw.GetXml(r.Modified[0].Id), "Unid=\"([0-9a-fA-F]+)\"")
            .Select(m => m.Groups[1].Value).ToHashSet();
        var secondUnids = System.Text.RegularExpressions.Regex
            .Matches(s.Raw.GetXml(r.Created[0].Id), "Unid=\"([0-9a-fA-F]+)\"")
            .Select(m => m.Groups[1].Value).ToHashSet();
        Assert.NotEmpty(secondUnids);
        Assert.Empty(firstUnids.Intersect(secondUnids)); // no Unid shared between the two subtrees
    }

    [Fact]
    public void DS043_MergeParagraphs()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchors = s.Project().AnchorIndex.Keys.ToList();
        var first = anchors[0];
        var second = anchors[1];

        var r = s.MergeParagraphs(first, second);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Id == first);
        Assert.Contains(r.Removed, a => a.Id == second);
        Assert.False(s.Exists(second));
        // MergeParagraphs inserts a single-space separator when both sides end/start
        // with non-whitespace — otherwise sentences jam together. See DS085.
        Assert.Contains("First paragraph. Second paragraph.", s.Project().Markdown);
    }

    [Fact]
    public void DS044_MergeParagraphs_NotAdjacent()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchors = s.Project().AnchorIndex.Keys.ToList();
        s.InsertParagraph(anchors[0], Position.After, "Middle paragraph.");
        var r = s.MergeParagraphs(anchors[0], anchors[1]);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorsNotAdjacent, r.Error!.Code);
    }

    // ─── Phase 5: formatting ──────────────────────────────────────────────

    [Fact]
    public void DS050_SetParagraphStyle()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.SetParagraphStyle(anchor, "Heading2");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Single(r.Modified);
        Assert.Equal("h", r.Modified[0].Kind);
        Assert.Contains("## ", s.Project().Markdown);
    }

    [Fact]
    public void DS051_SetParagraphStyle_UnknownStyle()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.SetParagraphStyle(anchor, "NotARealStyle1234");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.UnknownStyle, r.Error!.Code);
    }

    [Fact]
    public void DS053_SetParagraphStyle_CreatesMissingBuiltInStyle()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        // "Title" is a well-known built-in but BuildHeadingStyles defines only Heading1-6/Code/
        // Quote/MyListStyle, so it is absent. Applying a built-in the document hasn't defined used
        // to fail with UnknownStyle (a silent no-op in the editor); now the style is find-or-created
        // (like the inline "Code" character style) and applied.
        var r = s.SetParagraphStyle(anchor, "Title");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("Title", s.GetBlockMetadata(r.Modified[0].Id)?.StyleId);
    }

    [Fact]
    public void DS054_SetParagraphStyle_CreatesMissingHeadingAsHeading()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        // Heading8 is absent (BuildHeadingStyles stops at 6). The synthesized heading carries an
        // outlineLvl, so the anchor flips p -> h and it projects as a markdown heading.
        var r = s.SetParagraphStyle(anchor, "Heading8");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("h", r.Modified[0].Kind);
        Assert.Equal("Heading8", s.GetBlockMetadata(r.Modified[0].Id)?.StyleId);
    }

    [Fact]
    public void DS052_ApplyFormat_WholeParagraphBold()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ApplyFormat(anchor, span: null, new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("**First paragraph.**", s.Project().Markdown);
    }

    [Fact]
    public void DS053_ApplyFormat_Span()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        // "First paragraph." → bold characters 0..5 ("First")
        var r = s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("**First**", s.Project().Markdown);
    }

    [Fact]
    public void DS200_ApplyFormat_Superscript_EmitsVertAlign()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { VertAlign = "superscript" });
        Assert.True(r.Success, r.Error?.Message);
        var xml = s.Raw.GetXml(anchor);
        Assert.Contains("vertAlign", xml);
        Assert.Contains("superscript", xml);
        // Clear it back to baseline.
        var r2 = s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { VertAlign = "" });
        Assert.True(r2.Success, r2.Error?.Message);
        Assert.DoesNotContain("vertAlign", s.Raw.GetXml(anchor));
    }

    [Fact]
    public void DS201_SetParagraphFormat_Alignment_Center()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.SetParagraphFormat(anchor, new ParagraphFormatOp { Alignment = ParagraphAlignment.Center });
        Assert.True(r.Success, r.Error?.Message);
        var xml = s.Raw.GetXml(anchor);
        Assert.Contains("jc", xml);
        Assert.Contains("center", xml);
    }

    [Fact]
    public void DS202_SetParagraphFormat_PageBreakBefore_And_Indent()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.SetParagraphFormat(anchor, new ParagraphFormatOp { PageBreakBefore = true, IndentDelta = 720 });
        Assert.True(r.Success, r.Error?.Message);
        var xml = s.Raw.GetXml(anchor);
        Assert.Contains("pageBreakBefore", xml);
        Assert.Contains("720", xml);
        // Indent again accumulates; removing the page break leaves the indent.
        var r2 = s.SetParagraphFormat(anchor, new ParagraphFormatOp { PageBreakBefore = false, IndentDelta = 720 });
        Assert.True(r2.Success, r2.Error?.Message);
        var xml2 = s.Raw.GetXml(anchor);
        Assert.DoesNotContain("pageBreakBefore", xml2);
        Assert.Contains("1440", xml2); // 720 + 720
    }

    [Fact]
    public void DS215_SetParagraphFormat_Indent_ToleratesNonIntegerExistingValue()
    {
        // Google-Docs-exported documents emit non-integer twips like w:left="12.996749877929688".
        // SetParagraphFormat's indent delta must read that tolerantly (matching the HTML converter's
        // AttributeToTwips: decimal → truncate) instead of throwing a FormatException, and write back
        // a clean integer. Regression for indent being silently dead on every paragraph of such docs.
        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = wDoc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                main.Document.Body!.Append(new Paragraph(
                    new ParagraphProperties(new Indentation { Left = "12.996749877929688" }),
                    new Run(new Text("Indented paragraph."))));
                main.Document.Save();
            }
            bytes = ms.ToArray();
        }

        using var s = new DocxSession(bytes);
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.SetParagraphFormat(anchor, new ParagraphFormatOp { IndentDelta = 720 });
        Assert.True(r.Success, r.Error?.Message);
        var xml = s.Raw.GetXml(anchor);
        Assert.Contains("732", xml);          // 12 (truncated from 12.996…) + 720
        Assert.DoesNotContain("12.99", xml);  // the non-integer value is normalized away
    }

    [Fact]
    public void DS216_ApplyFormat_Bold_TurnsOnExplicitlyOffRun()
    {
        // Google Docs stamps an explicit <w:b w:val="0"/> (bold OFF) on every run. Toggling bold ON
        // must normalize that to ON (bare <w:b/>), not no-op because "a w:b element already exists".
        // Regression for bold/italic/strike silently doing nothing on such documents.
        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = wDoc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                main.Document.Body!.Append(new Paragraph(new Run(
                    new RunProperties(new Bold { Val = OnOffValue.FromBoolean(false) }),
                    new Text("Hello world"))));
                main.Document.Save();
            }
            bytes = ms.ToArray();
        }

        using var s = new DocxSession(bytes);
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);
        var html = Docxodus.Internal.HtmlConversionOps.RenderBlockHtml(s,
            r.Modified is { Count: > 0 } ? r.Modified[0].Id : anchor,
            new Docxodus.Internal.HtmlConversionOptions { FabricateCssClasses = false });
        Assert.Contains("font-weight: bold", html); // the bolded span actually renders bold
    }

    [Fact]
    public void DS217_RenderBlock_InvisiblePBdr_DoesNotEatIndent()
    {
        // An all-"nil" w:pBdr (Google Docs stamps one on every paragraph) is NOT a visible border.
        // CreateBorderDivs must not wrap such a paragraph in a border <div> and relocate its left
        // indent onto the div — that forces the paragraph's own margin-left to 0, so the editor's
        // single-block re-render silently loses indentation. Regression for indent never showing.
        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = wDoc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                main.Document.Body!.Append(new Paragraph(
                    new ParagraphProperties(
                        new ParagraphBorders(
                            new TopBorder { Val = BorderValues.Nil, Size = 0U, Space = 0U },
                            new LeftBorder { Val = BorderValues.Nil, Size = 0U, Space = 0U },
                            new BottomBorder { Val = BorderValues.Nil, Size = 0U, Space = 0U },
                            new RightBorder { Val = BorderValues.Nil, Size = 0U, Space = 0U }),
                        new Indentation { Left = "1440" }),
                    new Run(new Text("Indented under an invisible border."))));
                main.Document.Save();
            }
            bytes = ms.ToArray();
        }

        using var s = new DocxSession(bytes);
        var anchor = s.Project().AnchorIndex.Keys.First();
        var html = Docxodus.Internal.HtmlConversionOps.RenderBlockHtml(s, anchor,
            new Docxodus.Internal.HtmlConversionOptions { FabricateCssClasses = false });
        // The <p> itself carries the 1-inch indent; it is not eaten by a wrapping border div.
        Assert.Contains("margin-left: 1.00in", html);
    }

    [Fact]
    public void DS213_ApplyListFormat_Bullet_RendersMarker()
    {
        // The single-block render (the editor's incremental path) must show the bullet marker
        // and the list's hanging indent. The Symbol-font glyph U+F0B7 is mapped to Unicode (•).
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var p = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
        var li = s.ApplyListFormat(p, ListFormat.Bullet).Modified[0].Id;
        var html = Docxodus.Internal.HtmlConversionOps.RenderBlockHtml(s, li,
            new Docxodus.Internal.HtmlConversionOptions { FabricateCssClasses = false });
        Assert.Contains("•", html);            // Unicode bullet marker rendered (mapped from Symbol U+F0B7)
        Assert.Contains("text-indent", html); // hanging indent applied
    }

    [Fact]
    public void DS214_ApplyFormat_Code_CreatesMissingCodeCharacterStyle()
    {
        // On a document that does NOT define a "Code" style, inline code (FormatOp.Code)
        // referenced a phantom style → Word silently ignored it (no monospace). The op must
        // find-or-create a real *character* style so the run actually renders as code.
        using var s = new DocxSession(BuildDocWithoutCodeStyle());
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { Code = true });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("Code", s.Raw.GetXml(anchor)); // run references the style

        // Save + reopen: the Code character style is defined and persisted with a monospace font.
        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var codeStyle = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().FirstOrDefault(st => st.StyleId == "Code");
        Assert.NotNull(codeStyle);
        Assert.Equal(StyleValues.Character, codeStyle!.Type!.Value);
        Assert.Equal("Consolas", codeStyle.StyleRunProperties?.RunFonts?.Ascii?.Value);
    }

    [Fact]
    public void DS210_ApplyListFormat_Bullet_PromotesAndReuses()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var ps = s.Project().AnchorIndex.Keys.Where(k => k.StartsWith("p:")).ToList();
        var r1 = s.ApplyListFormat(ps[0], ListFormat.Bullet);
        Assert.True(r1.Success, r1.Error?.Message);
        var li1 = r1.Modified[0].Id;
        Assert.StartsWith("li:", li1); // promoted plain paragraph → list item
        var lm1 = s.GetListMembership(li1);
        Assert.NotNull(lm1);
        Assert.Equal(NumberFormat.Bullet, lm1!.Format);

        // A second bullet paragraph reuses the same synthesized numbering definition.
        var r2 = s.ApplyListFormat(ps[1], ListFormat.Bullet);
        Assert.True(r2.Success, r2.Error?.Message);
        var lm2 = s.GetListMembership(r2.Modified[0].Id);
        Assert.Equal(lm1.NumId, lm2!.NumId);
    }

    [Fact]
    public void DS211_ApplyListFormat_Decimal_ThenNone()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var p = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
        var r = s.ApplyListFormat(p, ListFormat.Decimal);
        Assert.True(r.Success, r.Error?.Message);
        var li = r.Modified[0].Id;
        Assert.Equal(NumberFormat.Decimal, s.GetListMembership(li)!.Format);

        var r2 = s.ApplyListFormat(li, ListFormat.None);
        Assert.True(r2.Success, r2.Error?.Message);
        Assert.StartsWith("p:", r2.Modified[0].Id); // back to a plain paragraph
        Assert.Null(s.GetListMembership(r2.Modified[0].Id));
    }

    [Fact]
    public void DS212_ApplyListFormat_RoundTrips()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var p = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
        s.ApplyListFormat(p, ListFormat.Bullet);
        var saved = s.Save();

        using var s2 = new DocxSession(saved);
        var li = s2.Project().AnchorIndex.Keys.FirstOrDefault(k => k.StartsWith("li:"));
        Assert.NotNull(li);
        Assert.Equal(NumberFormat.Bullet, s2.GetListMembership(li!)!.Format);
    }

    [Fact]
    public void DS054_SetListLevelIndent()
    {
        using var s = new DocxSession(BuildDS002_BulletedList());
        var firstLi = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("li:"));
        var r = s.SetListLevel(firstLi, +1);
        Assert.True(r.Success, r.Error?.Message);
    }

    [Fact]
    public void DS054b_SetListLevel_StyleInheritedSingleLevel_MaterializesNumPrAndSynthesizesLevel()
    {
        // Real-world regression: a style-inherited bullet (numPr on the style, abstractNum with
        // only level 0) used to be a silent no-op on nest. After the fix SetListLevel materializes
        // a direct numPr at the new level AND synthesizes the missing abstractNum level.
        using var s = new DocxSession(BuildStyleListSingleLevel());
        var firstLi = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("li:"));
        var r = s.SetListLevel(firstLi, +1);
        Assert.True(r.Success, r.Error?.Message);

        var bytes = s.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // The styled paragraph now carries a DIRECT numPr at ilvl=1 (materialized from the style).
        var firstP = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(w + "p").First();
        var directNumPr = firstP.Element(w + "pPr")?.Element(w + "numPr");
        Assert.NotNull(directNumPr);
        Assert.Equal("1", (string?)directNumPr!.Element(w + "ilvl")?.Attribute(w + "val"));
        Assert.Equal("1", (string?)directNumPr.Element(w + "numId")?.Attribute(w + "val"));

        // The abstractNum now DEFINES level 1 (synthesized — only level 0 existed before)...
        var abstractNum = doc.MainDocumentPart!.NumberingDefinitionsPart!.GetXDocument().Root!
            .Elements(w + "abstractNum").First();
        var levels = abstractNum.Elements(w + "lvl").Select(l => (string?)l.Attribute(w + "ilvl")).ToList();
        Assert.Contains("0", levels);
        Assert.Contains("1", levels);
        // ...and multiLevelType is upgraded off singleLevel, else the converter would force ilvl=0.
        Assert.NotEqual("singleLevel", (string?)abstractNum.Element(w + "multiLevelType")?.Attribute(w + "val"));

        // End-to-end: the converter must render the nested item with LEVEL 1's bullet glyph
        // ("o"), not collapse it to level 0's "·". This guards BOTH the singleLevel→ilvl=0 force
        // AND the bullet-continuation collapse — either bug renders the nested item as "· First…".
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("x.docx", bytes),
            new WmlToHtmlConverterSettings { StampAnchors = true }).ToString(SaveOptions.DisableFormatting);
        var text = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", string.Empty)
            .Replace("&#x00a0;", " ");
        Assert.Contains("o First bullet", text, StringComparison.Ordinal);   // nested → level 1 glyph
        Assert.Contains("· Second bullet", text, StringComparison.Ordinal);  // sibling stays level 0
    }

    [Fact]
    public void DS055_RemoveListMembership()
    {
        using var s = new DocxSession(BuildDS002_BulletedList());
        var firstLi = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("li:"));
        var r = s.RemoveListMembership(firstLi);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Kind == "p");
    }

    // ─── Phase 6: cell content + tracked-change mode ─────────────────────

    [Fact]
    public void DS060_ReplaceCellContent()
    {
        using var s = new DocxSession(BuildDS003_TableWithCells());
        var cellAnchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("tc:"));
        var r = s.ReplaceCellContent(cellAnchor, "New cell text.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("New cell text.", s.Project().Markdown);
    }

    [Fact]
    public void DS061_ReplaceText_Tracked()
    {
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "test-agent",
        };
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(), settings);
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.ReplaceText(anchor, "New text.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Id == anchor);

        // Round-trip to byte form and inspect the XML for w:ins/w:del markers
        var bytes = s.Save();
        using var ms = new MemoryStream(bytes);
        using var verify = WordprocessingDocument.Open(ms, isEditable: false);
        var docXml = verify.MainDocumentPart!.GetXDocument().Root!.ToString();
        Assert.Contains("w:ins", docXml);
        Assert.Contains("w:del", docXml);
        Assert.Contains("test-agent", docXml);
    }

    [Fact]
    public void DS062_DeleteBlock_Tracked()
    {
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "tester",
        };
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(), settings);
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.DeleteBlock(anchor);
        Assert.True(r.Success, r.Error?.Message);
        // In tracked mode, anchor stays live (modified, not removed)
        Assert.Empty(r.Removed);
        Assert.Contains(r.Modified, a => a.Id == anchor);
    }

    // ─── Phase 7: raw escape hatch ───────────────────────────────────────

    [Fact]
    public void DS070_RawGetXml()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var xml = s.Raw.GetXml(anchor);
        Assert.Contains("First paragraph.", xml);
        Assert.Contains("w:p", xml);
    }

    [Fact]
    public void DS071_RawInsertXml()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var newP = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Raw inserted.</w:t></w:r></w:p>";
        var r = s.Raw.InsertXml(anchor, Position.After, newP);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Single(r.Created);
        Assert.Contains("Raw inserted.", s.Project().Markdown);
    }

    [Fact]
    public void DS072_RawInsertXml_MalformedRejected()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.Raw.InsertXml(anchor, Position.After, "<not-closed");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.MalformedXml, r.Error!.Code);
    }

    [Fact]
    public void DS073_RawInsertXml_DisallowedNs()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.Raw.InsertXml(anchor, Position.After, "<foo xmlns=\"http://evil/\"/>");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.DisallowedNamespace, r.Error!.Code);
    }

    [Fact]
    public void DS074_RawReplaceXml()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var newP = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Raw replacement.</w:t></w:r></w:p>";
        var r = s.Raw.ReplaceXml(anchor, newP);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("Raw replacement.", s.Project().Markdown);
        Assert.Contains(r.Removed, a => a.Id == anchor);
    }

    [Fact]
    public void DS075_Raw_GetThenReplaceRoundtrip()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var xml = s.Raw.GetXml(anchor);
        // Naive mutation: prefix the text with "EDITED: "
        var mutated = xml.Replace("First paragraph.", "EDITED: First paragraph.");
        var r = s.Raw.ReplaceXml(anchor, mutated);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("EDITED: First paragraph.", s.Project().Markdown);
    }

    // ─── Bug-fix regressions (post-PR-131 verification) ──────────────────

    [Fact]
    public void DS080_StalePrefix_FallsBackToUnidLookup()
    {
        // A cached anchor id whose kind-prefix has gone stale (e.g., `p:body:abcd`
        // after promoting the paragraph to `h:body:abcd` via SetParagraphStyle)
        // must still resolve via FindAnchor's Unid fallback, as promised by
        // `docs/architecture/docx_mutation_api.md`.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var oldId = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body").Anchor.Id;
        Assert.StartsWith("p:body:", oldId);

        var promote = s.SetParagraphStyle(oldId, "Heading2");
        Assert.True(promote.Success, promote.Error?.Message);
        Assert.Equal("h", promote.Modified[0].Kind);
        Assert.NotEqual(oldId, promote.Modified[0].Id);

        // Operations using the stale id should still succeed via fallback.
        Assert.True(s.Exists(oldId), "Exists() must accept stale-prefix id");
        Assert.NotNull(s.GetAnchorInfo(oldId));

        var r = s.ReplaceText(oldId, "Replaced via stale id.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("Replaced via stale id.", s.Project().Markdown);
    }

    [Fact]
    public void DS081_StalePrefix_UnknownIdStillReturnsAnchorNotFound()
    {
        // Sanity guard: the fallback must NOT make every malformed id resolve —
        // a totally unknown Unid still fails AnchorNotFound.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var bogus = "p:body:" + new string('0', 32);
        Assert.False(s.Exists(bogus));
        var r = s.ReplaceText(bogus, "anything");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, r.Error!.Code);
    }

    [Fact]
    public void DS082_ValidateRawOps_SucceedsWhenNoNewErrors()
    {
        // ValidateRawOps must use delta semantics, not "zero errors total" —
        // every Project() call adds PtOpenXml:Unid attributes which are not in
        // the OOXML schema and would otherwise trip Sch_UndeclaredAttribute on
        // every op. Filtering those + counting deltas is what the doc promises.
        var settings = new DocxSessionSettings { ValidateRawOps = true };
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(), settings);
        var anchor = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p").Anchor.Id;

        var ok = """
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:r><w:t xml:space="preserve">VALIDATED</w:t></w:r>
            </w:p>
            """;
        var r = s.Raw.ReplaceXml(anchor, ok);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("VALIDATED", s.Project().Markdown);
    }

    [Fact]
    public void DS083_ValidateRawOps_RollsBackWhenSchemaInvalid()
    {
        // A fragment with an undeclared element in the w: namespace must
        // increment the validator error count and trigger rollback.
        var settings = new DocxSessionSettings { ValidateRawOps = true };
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(), settings);
        var anchor = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p").Anchor.Id;
        var before = s.Project().Markdown;

        // A w:jc with an unknown alignment enum value trips
        // Sch_AttributeValueDataTypeDetailed (Enumeration constraint failed).
        var bad = """
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:pPr><w:jc w:val="NOT_A_REAL_ALIGNMENT"/></w:pPr>
            </w:p>
            """;
        var r = s.Raw.ReplaceXml(anchor, bad);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.ValidationFailed, r.Error!.Code);
        Assert.Equal(before, s.Project().Markdown);
    }

    // ─── Phase 10: Grep primitive (#143) ─────────────────────────────────

    /// <summary>
    /// Three-paragraph fixture: paragraph 1 has a single plain run "Once upon a time, in a faraway land.";
    /// paragraph 2 has formatting boundaries that split runs ("Plain " + bold "BOLD" + " plain again");
    /// paragraph 3 has a hyperlink in the middle ("Visit " + hyperlink("Anthropic") + " for more.").
    /// </summary>
    internal static byte[] BuildDS100_GrepFixture()
    {
        XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var rel = main.AddHyperlinkRelationship(new System.Uri("https://www.anthropic.com"), true);

            var body = new XElement(W + "body",
                new XElement(W + "p",
                    new XElement(W + "r",
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                            "Once upon a time, in a faraway land."))),
                new XElement(W + "p",
                    new XElement(W + "r",
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Plain ")),
                    new XElement(W + "r",
                        new XElement(W + "rPr", new XElement(W + "b")),
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "BOLD")),
                    new XElement(W + "r",
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " plain again."))),
                new XElement(W + "p",
                    new XElement(W + "r",
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Visit ")),
                    new XElement(W + "hyperlink",
                        new XAttribute(R + "id", rel.Id),
                        new XElement(W + "r",
                            new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Anthropic"))),
                    new XElement(W + "r",
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " for more."))));

            var doc = new XElement(W + "document",
                new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                body);
            main.PutXDocument(new XDocument(doc));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS100_Grep_SingleRunMatch()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep("faraway land");
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal("faraway land", m.Text);
        Assert.Equal("p", m.EnclosingAnchor.Anchor.Kind);
        Assert.Single(m.Fragments);
        Assert.Equal("faraway land", m.Fragments[0].Text);
        Assert.False(m.Fragments[0].Formatting.Bold);
        Assert.Null(m.Fragments[0].Formatting.HyperlinkUrl);
    }

    [Fact]
    public void DS101_Grep_MatchSpanningFormattingBoundary()
    {
        // "ain BOLD pl" crosses two formatting boundaries: plain → bold → plain.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep("ain BOLD pl");
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal("ain BOLD pl", m.Text);
        Assert.Equal(3, m.Fragments.Count);

        Assert.Equal("ain ", m.Fragments[0].Text);
        Assert.False(m.Fragments[0].Formatting.Bold);

        Assert.Equal("BOLD", m.Fragments[1].Text);
        Assert.True(m.Fragments[1].Formatting.Bold);

        Assert.Equal(" pl", m.Fragments[2].Text);
        Assert.False(m.Fragments[2].Formatting.Bold);

        // The three fragments must reference three distinct runs (distinct Unids).
        var unids = m.Fragments.Select(f => f.Unid).Distinct().ToList();
        Assert.Equal(3, unids.Count);
    }

    [Fact]
    public void DS102_Grep_MatchSpanningHyperlink()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep("Visit Anthropic for");
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal(3, m.Fragments.Count);
        Assert.Equal("Visit ", m.Fragments[0].Text);
        Assert.Null(m.Fragments[0].Formatting.HyperlinkUrl);

        Assert.Equal("Anthropic", m.Fragments[1].Text);
        Assert.Equal("https://www.anthropic.com/", m.Fragments[1].Formatting.HyperlinkUrl);

        Assert.Equal(" for", m.Fragments[2].Text);
        Assert.Null(m.Fragments[2].Formatting.HyperlinkUrl);
    }

    [Fact]
    public void DS103_Grep_RegexWithGroups()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep(@"(?<who>Once) upon a (?<when>time)");
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal("Once upon a time", m.Text);
        // Group 0 == whole match; named groups appear at their index.
        Assert.Equal("Once", m.Groups[1]);
        Assert.Equal("time", m.Groups[2]);
    }

    [Fact]
    public void DS104_Grep_NoMatchReturnsEmpty()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep("string that absolutely does not appear anywhere");
        Assert.Empty(matches);
    }

    [Fact]
    public void DS105_Grep_ContextBeforeAndAfter()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep("BOLD");
        Assert.Single(matches);

        var m = matches[0];
        Assert.EndsWith("Plain ", m.ContextBefore);
        Assert.StartsWith(" plain again.", m.ContextAfter);
    }

    [Fact]
    public void DS106_Grep_MultipleMatchesInDocumentOrder()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var matches = s.Grep("a"); // very common letter, will hit many places
        Assert.True(matches.Count >= 3);

        // Document order: first match comes from paragraph 1, then 2, then 3.
        var firstThreeAnchors = matches.Take(3).Select(x => x.EnclosingAnchor.Anchor.Id).ToList();
        var bodyAnchors = s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p").Select(t => t.Anchor.Id).ToList();
        // Each successive match's anchor index must be >= the previous (matches are emitted in doc order).
        var positions = firstThreeAnchors.Select(a => bodyAnchors.IndexOf(a)).ToList();
        for (int i = 1; i < positions.Count; i++)
            Assert.True(positions[i] >= positions[i - 1], "matches must be in document order");
    }

    [Fact]
    public void DS107_Grep_RegexOptionsRespected()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        // Case-insensitive should find "BOLD" via lowercase "bold".
        var insensitive = s.Grep("bold", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        Assert.Single(insensitive);
        Assert.Equal("BOLD", insensitive[0].Text);

        // Case-sensitive (default) must not match.
        var sensitive = s.Grep("bold");
        Assert.Empty(sensitive);
    }

    [Fact]
    public void DS108_Grep_SpanInsideElement_PointsAtMatchingSubstring()
    {
        // For the bold-spanning case, the middle fragment's text is the WHOLE w:r
        // text ("BOLD") and SpanInElement covers 0..4 of that run.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var m = s.Grep("ain BOLD pl").Single();

        var bold = m.Fragments[1];
        Assert.Equal("BOLD", bold.Text);
        Assert.Equal(0, bold.SpanInElement.Start);
        Assert.Equal(4, bold.SpanInElement.Length);

        // The trailing-plain fragment starts at offset 0 of " plain again." and covers " pl" (3 chars).
        var trailing = m.Fragments[2];
        Assert.Equal(0, trailing.SpanInElement.Start);
        Assert.Equal(3, trailing.SpanInElement.Length);
    }

    // ─── Phase 11: ReplaceTextRange (#139) ───────────────────────────────

    private static string FlatBodyText(DocxSession s)
    {
        return string.Join("\n", s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && (t.Anchor.Kind == "p" || t.Anchor.Kind == "h" || t.Anchor.Kind == "li"))
            .Select(t =>
            {
                var xml = s.Raw.GetXml(t.Anchor.Id);
                var el = XElement.Parse(xml);
                return string.Concat(el.Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                    .Select(tn => (string)tn));
            }));
    }

    [Fact]
    public void DS110_ReplaceTextRange_SingleFragmentReplacement()
    {
        // First paragraph of the Grep fixture: one plain run with
        // "Once upon a time, in a faraway land."
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var anchor = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p").Anchor.Id;

        var results = s.ReplaceTextRange(anchor, "faraway", "distant");

        Assert.Single(results);
        Assert.True(results[0].Success, results[0].Error?.Message);
        Assert.Contains("Once upon a time, in a distant land.", FlatBodyText(s));
        Assert.DoesNotContain("faraway", FlatBodyText(s));
    }

    [Fact]
    public void DS111_ReplaceTextRange_MultiFragmentPreservesRemainingFormatting()
    {
        // Second paragraph: "Plain " + bold "BOLD" + " plain again."
        // Replacing "ain BOLD pl" must:
        //   - drop the participating slice from each of the 3 runs
        //   - inject the replacement into the FIRST fragment's run
        //   - leave the bold run's formatting intact for any text that survives
        //     (in this case the bold run's slice is the whole run, so the bold run
        //      ends up empty but still present with bold rPr)
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var paragraphs = s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p").ToList();
        var second = paragraphs[1].Anchor.Id;

        var results = s.ReplaceTextRange(second, "ain BOLD pl", "REPL");

        Assert.Single(results);
        Assert.True(results[0].Success, results[0].Error?.Message);

        // Resulting flat text: "PlREPLain again."  ("Pl" from before "ain" + REPL + "ain again." after " pl")
        var xml = s.Raw.GetXml(second);
        var el = XElement.Parse(xml);
        var Wt = XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var Wr = XName.Get("r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var Wb = XName.Get("b", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var concat = string.Concat(el.Descendants(Wt).Select(t => (string)t));
        Assert.Equal("PlREPLain again.", concat);

        // The bold run must still exist with rPr/<w:b/> even after losing all its text,
        // because preserving formatting is the whole point. (Future op could prune
        // empty runs but the contract is "formatting survives".)
        var boldRun = el.Descendants(Wr).FirstOrDefault(r =>
            r.Element(XName.Get("rPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))?.Element(Wb) is not null);
        Assert.NotNull(boldRun);
    }

    [Fact]
    public void DS112_ReplaceTextRange_MultipleMatchesInSameParagraph()
    {
        // First paragraph contains the letter "a" multiple times.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var first = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p").Anchor.Id;
        var beforeCount = FlatBodyText(s).Count(c => c == 'a');

        var results = s.ReplaceTextRange(first, "a", "@");

        Assert.True(results.Count >= 3);
        Assert.All(results, r => Assert.True(r.Success, r.Error?.Message));

        var after = FlatBodyText(s);
        // Every 'a' in the first paragraph is now '@'; no leftovers in that paragraph.
        var firstParaXml = s.Raw.GetXml(first);
        var firstParaText = string.Concat(XElement.Parse(firstParaXml)
            .Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
            .Select(t => (string)t));
        Assert.DoesNotContain("a", firstParaText);
    }

    [Fact]
    public void DS113_ReplaceTextRange_FindNotFound_ReturnsEmpty()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var anchor = s.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p").Anchor.Id;
        var before = FlatBodyText(s);

        var results = s.ReplaceTextRange(anchor, "nope-this-is-not-in-the-doc", "irrelevant");

        Assert.Empty(results);
        Assert.Equal(before, FlatBodyText(s));
    }

    [Fact]
    public void DS114_ReplaceTextRange_IgnoreCase()
    {
        // "BOLD" is uppercase in the fixture; case-insensitive find with lowercase needle should hit.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var second = s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p").Skip(1).First().Anchor.Id;

        var results = s.ReplaceTextRange(second, "bold", "calm", new ReplaceOptions { IgnoreCase = true });
        Assert.Single(results);
        Assert.True(results[0].Success);
        Assert.Contains("Plain calm plain again.", FlatBodyText(s));
    }

    [Fact]
    public void DS115_ReplaceTextRange_MaxReplacementsHonored()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var first = s.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p").Anchor.Id;

        var results = s.ReplaceTextRange(first, "a", "@", new ReplaceOptions { MaxReplacements = 2 });
        Assert.Equal(2, results.Count);

        // Exactly two 'a's became '@'; subsequent 'a's still present in the paragraph.
        var firstParaXml = s.Raw.GetXml(first);
        var firstParaText = string.Concat(XElement.Parse(firstParaXml)
            .Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
            .Select(t => (string)t));
        Assert.Contains("a", firstParaText);
        Assert.Equal(2, firstParaText.Count(c => c == '@'));
    }

    [Fact]
    public void DS116_ReplaceTextRange_EmptyReplaceDeletesText()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var second = s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p").Skip(1).First().Anchor.Id;

        var results = s.ReplaceTextRange(second, "BOLD", "");
        Assert.Single(results);
        Assert.True(results[0].Success);
        Assert.Contains("Plain  plain again.", FlatBodyText(s));   // double space where BOLD was
    }

    [Fact]
    public void DS117_ReplaceMatch_FromGrepResult()
    {
        // ReplaceMatch is the convenience overload that takes a TextMatch directly,
        // so the caller doesn't pay for re-scanning the anchor.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var match = s.Grep("faraway").Single();

        var r = s.ReplaceMatch(match, "nearby");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("nearby land", FlatBodyText(s));
    }

    [Fact]
    public void DS118_ReplaceTextRange_AnchorNotFound()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var results = s.ReplaceTextRange("p:body:deadbeefdeadbeefdeadbeefdeadbeef", "anything", "else");
        Assert.Single(results);
        Assert.False(results[0].Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, results[0].Error!.Code);
    }

    [Fact]
    public void DS119_ReplaceTextRange_UndoRestoresPriorState()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var anchor = s.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p").Anchor.Id;
        var before = FlatBodyText(s);

        s.ReplaceTextRange(anchor, "faraway", "distant");
        Assert.True(s.Undo());
        Assert.Equal(before, FlatBodyText(s));
    }

    // ─── Phase 12: FindPlaceholders (#142) ───────────────────────────────

    /// <summary>
    /// Three-paragraph fixture covering each PlaceholderKind:
    ///   1. BlankFill — "Name: [_______]" + "Price: $[___]"
    ///   2. AlternativeClause — "[There shall be no cumulative voting.]"
    ///   3. Instruction — "[insert percentage] of the outstanding shares" + "[*specify name*]"
    /// </summary>
    internal static byte[] BuildDS120_PlaceholderFixture()
    {
        XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            XElement Para(string text) => new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));

            var body = new XElement(W + "body",
                Para("Name: [_______] and price: $[___]"),
                Para("[There shall be no cumulative voting.]"),
                Para("Holders of at least [insert percentage] of the outstanding shares of [*specify name*]."));

            var doc = new XElement(W + "document",
                new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                body);
            main.PutXDocument(new XDocument(doc));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS120_FindPlaceholders_DetectsBlankFill()
    {
        using var s = new DocxSession(BuildDS120_PlaceholderFixture());
        var placeholders = s.FindPlaceholders(PlaceholderKinds.BlankFill);

        Assert.Equal(2, placeholders.Count);
        Assert.All(placeholders, p => Assert.Equal(PlaceholderKind.BlankFill, p.Kind));
        Assert.Contains(placeholders, p => p.Match.Text == "[_______]");
        Assert.Contains(placeholders, p => p.Match.Text == "$[___]");
    }

    [Fact]
    public void DS121_FindPlaceholders_DetectsAlternativeClause()
    {
        using var s = new DocxSession(BuildDS120_PlaceholderFixture());
        var alts = s.FindPlaceholders(PlaceholderKinds.AlternativeClause);

        Assert.Single(alts);
        Assert.Equal(PlaceholderKind.AlternativeClause, alts[0].Kind);
        Assert.Equal("[There shall be no cumulative voting.]", alts[0].Match.Text);
    }

    [Fact]
    public void DS122_FindPlaceholders_DetectsInstruction_ExtractsHint()
    {
        using var s = new DocxSession(BuildDS120_PlaceholderFixture());
        var instructions = s.FindPlaceholders(PlaceholderKinds.Instruction);

        Assert.Equal(2, instructions.Count);
        Assert.All(instructions, i => Assert.Equal(PlaceholderKind.Instruction, i.Kind));
        Assert.Contains(instructions, i => i.Hint == "insert percentage");
        // Italic-styled instructions ([*specify name*]) strip the asterisks for the hint.
        Assert.Contains(instructions, i => i.Hint == "specify name");
    }

    [Fact]
    public void DS123_FindPlaceholders_DefaultKindsReturnsAll()
    {
        using var s = new DocxSession(BuildDS120_PlaceholderFixture());
        var all = s.FindPlaceholders();
        // 2 BlankFill + 1 AlternativeClause + 2 Instruction = 5
        Assert.Equal(5, all.Count);
    }

    [Fact]
    public void DS124_FindPlaceholders_KindsFilterCombines()
    {
        using var s = new DocxSession(BuildDS120_PlaceholderFixture());
        var blankAndInstr = s.FindPlaceholders(PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction);
        Assert.Equal(4, blankAndInstr.Count);
        Assert.DoesNotContain(blankAndInstr, p => p.Kind == PlaceholderKind.AlternativeClause);
    }

    [Fact]
    public void DS125_FindPlaceholders_EmptyDocReturnsEmpty()
    {
        // BuildDS001 is two paragraphs with no brackets at all.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        Assert.Empty(s.FindPlaceholders());
    }

    // ─── Phase 13: ApplyFormat substring + TextMatch overloads (#138) ────

    [Fact]
    public void DS130_ApplyFormat_BySubstring_FormatsFirstOccurrence()
    {
        // Avoids the offset-arithmetic trap from #138: caller passes the visible
        // text they want bolded; we resolve to a CharSpan internally.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var first = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p").Anchor.Id;

        var r = s.ApplyFormatToSubstring(first, "faraway", new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);
        // After projection, "faraway" wraps in **…** markers.
        Assert.Contains("**faraway**", s.Project().Markdown);
    }

    [Fact]
    public void DS131_ApplyFormat_BySubstring_NotFound_ReturnsOffsetOutOfRange()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var first = s.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p").Anchor.Id;
        var r = s.ApplyFormatToSubstring(first, "this-string-not-in-doc", new FormatOp { Bold = true });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, r.Error!.Code);
    }

    // ─── Phase 14: DeleteBlock for fn/en/cmt (#133) ──────────────────────

    /// <summary>
    /// Single-paragraph body with a footnote reference (id=1) plus a Footnotes part
    /// containing the required Word-reserved separators (id=-1, id=0) AND a user
    /// footnote (id=1) the test will delete.
    /// </summary>
    internal static byte[] BuildDS140_FootnoteFixture()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new FootnoteReferenceMark()))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark()))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 },
                new Footnote(new Paragraph(new Run(new Text("This footnote will be deleted.")))) { Type = FootnoteEndnoteValues.Normal, Id = 1 });
            footnotesPart.Footnotes.Save();

            var body = new Body();
            main.Document.Body = body;
            body.Append(new Paragraph(
                new Run(new Text("Main text") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FootnoteReference() { Id = 1 }),
                new Run(new Text(" continued.") { Space = SpaceProcessingModeValues.Preserve })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    internal static byte[] BuildDS141_EndnoteFixture()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var endnotesPart = main.AddNewPart<EndnotesPart>();
            endnotesPart.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new FootnoteReferenceMark()))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Endnote(new Paragraph(new Run(new ContinuationSeparatorMark()))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 },
                new Endnote(new Paragraph(new Run(new Text("This endnote will be deleted.")))) { Type = FootnoteEndnoteValues.Normal, Id = 1 });
            endnotesPart.Endnotes.Save();

            var body = new Body();
            main.Document.Body = body;
            body.Append(new Paragraph(
                new Run(new Text("Body text") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new EndnoteReference() { Id = 1 })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    internal static byte[] BuildDS142_CommentFixture()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments(
                new Comment(new Paragraph(new Run(new Text("This comment will be deleted."))))
                {
                    Id = "1", Initials = "TC", Author = "Test", Date = System.DateTime.UtcNow,
                });
            commentsPart.Comments.Save();

            var body = new Body();
            main.Document.Body = body;
            body.Append(new Paragraph(
                new CommentRangeStart() { Id = "1" },
                new Run(new Text("Body text") { Space = SpaceProcessingModeValues.Preserve }),
                new CommentRangeEnd() { Id = "1" },
                new Run(new CommentReference() { Id = "1" })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>True iff the anchor's serialized XML contains <paramref name="snippet"/>.</summary>
    private static bool AnchorXmlContains(DocxSession s, string anchorId, string snippet) =>
        s.Raw.GetXml(anchorId).Contains(snippet, System.StringComparison.Ordinal);

    [Fact]
    public void DS140_DeleteBlock_Footnote_RemovesDefinitionAndBodyReference()
    {
        using var s = new DocxSession(BuildDS140_FootnoteFixture());

        // Find specifically the user footnote (the only one with "will be deleted" text;
        // the two separator footnotes don't have user text).
        var userFootnote = s.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "fn"
                      && AnchorXmlContains(s, t.Anchor.Id, "will be deleted"));

        var r = s.DeleteBlock(userFootnote.Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);

        // Both the definition and the body's <w:footnoteReference w:id="1"/> are gone.
        var saved = s.Save();
        using var d = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var bodyXml = d.MainDocumentPart!.GetXDocument().Root!.ToString();
        Assert.DoesNotContain("footnoteReference", bodyXml);
        if (d.MainDocumentPart.FootnotesPart is not null)
        {
            var fnXml = d.MainDocumentPart.FootnotesPart.GetXDocument().Root!.ToString();
            Assert.DoesNotContain("This footnote will be deleted", fnXml);
        }
    }

    [Fact]
    public void DS140b_DeleteBlock_Footnote_RefusesSeparator()
    {
        // Word's id=-1 (separator) and id=0 (continuationSeparator) footnotes are
        // page-rendering scaffolding — deleting them corrupts the document.
        //
        // As of issue #162, these notes are filtered out of the AnchorIndex
        // entirely, so callers can no longer discover them via projection. The
        // new contract is stronger than the previous AnchorWrongKind refusal:
        // they're invisible to projection-driven editors. Verify (a) no fn-kind
        // anchor in the index resolves to a separator/continuationSeparator
        // footnote, and (b) any hand-synthesized anchor id pointing at one is
        // refused by DeleteBlock (the universal AnchorNotFound safety net).
        using var s = new DocxSession(BuildDS140_FootnoteFixture());

        var fnAnchors = s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "fn")
            .ToList();
        Assert.DoesNotContain(fnAnchors,
            t => AnchorXmlContains(s, t.Anchor.Id, "w:type=\"separator\"")
              || AnchorXmlContains(s, t.Anchor.Id, "w:type=\"continuationSeparator\""));

        // A synthetic anchor id targeting the boilerplate (whose Unid the caller
        // has no legitimate way to obtain) is rejected as not found.
        var r = s.DeleteBlock("fn:fn:00000000000000000000000000000000");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, r.Error!.Code);
    }

    [Fact]
    public void DS141_DeleteBlock_Endnote_RemovesDefinitionAndBodyReference()
    {
        using var s = new DocxSession(BuildDS141_EndnoteFixture());
        var userEndnote = s.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "en"
                      && AnchorXmlContains(s, t.Anchor.Id, "will be deleted"));

        var r = s.DeleteBlock(userEndnote.Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);

        var saved = s.Save();
        using var d = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var bodyXml = d.MainDocumentPart!.GetXDocument().Root!.ToString();
        Assert.DoesNotContain("endnoteReference", bodyXml);
    }

    [Fact]
    public void DS142_DeleteBlock_Comment_RemovesDefinitionAndAllRangeMarkers()
    {
        using var s = new DocxSession(BuildDS142_CommentFixture());
        var comment = s.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "cmt" && t.Anchor.Scope == "cmt");

        var r = s.DeleteBlock(comment.Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);

        var saved = s.Save();
        using var d = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var bodyXml = d.MainDocumentPart!.GetXDocument().Root!.ToString();
        // All three comment markers (commentReference, commentRangeStart, commentRangeEnd)
        // must be gone — leaving any of them behind makes Word render a broken marker.
        Assert.DoesNotContain("commentReference", bodyXml);
        Assert.DoesNotContain("commentRangeStart", bodyXml);
        Assert.DoesNotContain("commentRangeEnd", bodyXml);
    }

    // ─── Phase 15: WhitespaceMode + FindBy helpers + SmartQuotes (#136/#137/#140) ────

    /// <summary>Single body paragraph where the only whitespace between "First" and ":" is a NBSP.</summary>
    internal static byte[] BuildDS150_NbspFixture()
    {
        XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // The whitespace between "First" and ":" is a NON-BREAKING SPACE (U+00A0) —
            // exactly what NVCA legal templates use for ordinal: headings. Written as an
            // explicit escape so the source character isn't lost in editor round trips.
            var body = new XElement(W + "body",
                new XElement(W + "p",
                    new XElement(W + "r",
                        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                            "First\u00A0: The name of this corporation."))));
            main.PutXDocument(new XDocument(new XElement(W + "document",
                new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                body)));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS150_Grep_WhitespacePreserve_DoesNotMatchSpaceWhenSourceIsNbsp()
    {
        // Default behavior: NBSP is preserved as-is, so a needle with regular space misses.
        using var s = new DocxSession(BuildDS150_NbspFixture());
        var matches = s.Grep("First :");
        Assert.Empty(matches);
    }

    [Fact]
    public void DS151_Grep_WhitespaceNormalize_MatchesSpaceWhenSourceIsNbsp()
    {
        // Normalize mode: NBSP folds to regular space so the needle hits.
        using var s = new DocxSession(BuildDS150_NbspFixture());
        var matches = s.Grep(
            "First :",
            scope: ProjectionScopes.Body,
            whitespace: WhitespaceMode.Normalize);
        Assert.Single(matches);
        Assert.Equal("First :", matches[0].Text);
    }

    [Fact]
    public void DS152_Grep_WhitespaceNormalize_FragmentSpansPointAtOriginalText()
    {
        // Critical correctness check: even though we MATCHED on the normalized text,
        // the returned Span must address the same character positions in the original.
        // Otherwise a follow-up ReplaceMatch would land in the wrong place.
        using var s = new DocxSession(BuildDS150_NbspFixture());
        var m = s.Grep("First :", whitespace: WhitespaceMode.Normalize).Single();

        var r = s.ReplaceMatch(m, "Section 1 :");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("Section 1 : The name of this corporation.", s.Project().Markdown);
    }

    [Fact]
    public void DS160_FindByText_FirstMatch()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var t = s.FindByText("BOLD");
        Assert.NotNull(t);
        Assert.Equal("p", t!.Anchor.Kind);
    }

    [Fact]
    public void DS161_FindByText_IgnoreCase()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        Assert.Null(s.FindByText("bold"));  // case-sensitive default misses uppercase BOLD
        Assert.NotNull(s.FindByText("bold", new FindOptions { IgnoreCase = true }));
    }

    [Fact]
    public void DS162_FindByText_IgnoreWhitespace_HandlesNbsp()
    {
        using var s = new DocxSession(BuildDS150_NbspFixture());
        Assert.Null(s.FindByText("First :"));  // NBSP in source, regular space in needle
        Assert.NotNull(s.FindByText("First :", new FindOptions { IgnoreWhitespace = true }));
    }

    [Fact]
    public void DS163_FindAllByText_ReturnsDeduplicatedAnchorsInDocumentOrder()
    {
        // The Grep fixture has 3 paragraphs that all contain lowercase "n" ("Once"/"in"/
        // "Plain"/"again"/"Anthropic"); FindAllByText must collapse the many in-paragraph
        // hits into one entry per paragraph.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var all = s.FindAllByText("n");
        Assert.Equal(3, all.Count);
        Assert.Equal(all.Count, all.Select(a => a.Anchor.Id).Distinct().Count());
    }

    [Fact]
    public void DS164_FindByRegex_ReturnsAllMatchingAnchors()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var anchors = s.FindByRegex(@"\b\w*BOLD\w*\b");
        Assert.NotEmpty(anchors);
    }

    [Fact]
    public void DS165_FindByKind_FiltersByKindAndScope()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        Assert.Equal(2, s.FindByKind("p", "body").Count);
        Assert.Empty(s.FindByKind("p", "hdr1"));
        Assert.True(s.FindByKind("p").Count >= 2);
    }

    [Fact]
    public void DS166_FindOptions_KindFilter_Narrows()
    {
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var headingsOnly = s.FindAllByText("n", new FindOptions { KindFilter = "h" });
        Assert.All(headingsOnly, a => Assert.Equal("h", a.Anchor.Kind));
    }

    [Fact]
    public void DS170_SmartQuotes_OffByDefault_PassesThrough()
    {
        // Default: straight quotes survive unchanged.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Hello \"world\".");
        Assert.Contains("\"world\"", s.Project().Markdown);
        Assert.DoesNotContain('“', s.Project().Markdown);
    }

    [Fact]
    public void DS171_SmartQuotes_On_DoubleQuotesBecomeCurly()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { SmartQuotes = true });
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Hello \"world\".");
        var md = s.Project().Markdown;
        Assert.Contains("Hello “world”.", md);
        Assert.DoesNotContain("\"world\"", md);
    }

    [Fact]
    public void DS172_SmartQuotes_On_HandlesApostrophes()
    {
        // Mid-word ' becomes the right-single-quote (apostrophe) since the preceding
        // char isn't whitespace/open-punct — the same rule that "close quote" follows.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { SmartQuotes = true });
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Don't worry.");
        Assert.Contains("Don’t worry.", s.Project().Markdown);
    }

    [Fact]
    public void DS173_SmartQuotes_On_OpenQuoteAfterOpenBracket()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { SmartQuotes = true });
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "She said (\"hi\")");
        Assert.Contains("She said \\(“hi”\\)", s.Project().Markdown);
    }

    [Fact]
    public void DS174_SmartQuotes_PropagatesToReplaceTextRange()
    {
        // The SmartQuotes setting must apply to the surgical-edit path too,
        // not just ReplaceText.
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { SmartQuotes = true });
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceTextRange(anchor, "First", "\"first\"");
        Assert.Contains("“first”", s.Project().Markdown);
    }

    [Fact]
    public void DS143_DeleteBlock_Footnote_UndoRestores()
    {
        // The single-snapshot undo must roll back both the definition AND every
        // cross-reference that was stripped in one shot.
        using var s = new DocxSession(BuildDS140_FootnoteFixture());
        var userFootnote = s.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "fn"
                      && AnchorXmlContains(s, t.Anchor.Id, "will be deleted"));

        var savedBefore = s.Save();
        var delResult = s.DeleteBlock(userFootnote.Anchor.Id);
        Assert.True(delResult.Success, delResult.Error?.Message);
        Assert.True(s.Undo(), "Undo should restore the pre-delete state");
        var savedAfterUndo = s.Save();

        // Undo restores both the definition AND the in-body reference.
        using var dBefore = WordprocessingDocument.Open(new MemoryStream(savedBefore), false);
        using var dAfter = WordprocessingDocument.Open(new MemoryStream(savedAfterUndo), false);
        var bodyBefore = dBefore.MainDocumentPart!.GetXDocument().Root!.ToString();
        var bodyAfter = dAfter.MainDocumentPart!.GetXDocument().Root!.ToString();
        Assert.Contains("footnoteReference", bodyAfter);
        Assert.Equal(
            System.Text.RegularExpressions.Regex.Matches(bodyBefore, "footnoteReference").Count,
            System.Text.RegularExpressions.Regex.Matches(bodyAfter, "footnoteReference").Count);
    }

    [Fact]
    public void DS132_ApplyFormat_FromTextMatch_AddressesExactSpan()
    {
        // The TextMatch overload is the pair to ReplaceMatch — same shape: take the
        // exact (anchor, span) from a Grep result and apply formatting to it.
        using var s = new DocxSession(BuildDS100_GrepFixture());
        var match = s.Grep("faraway").Single();

        var r = s.ApplyFormat(match, new FormatOp { Italic = true });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("*faraway*", s.Project().Markdown);
    }

    [Fact]
    public void DS126_FindPlaceholders_ProvidesEnoughInfoForReplaceMatch()
    {
        // End-to-end: enumerate placeholders, fill each via ReplaceMatch.
        // This is the canonical agentic template-fill recipe.
        using var s = new DocxSession(BuildDS120_PlaceholderFixture());

        // Fill each BlankFill in reverse offset order (per the ReplaceTextRange contract).
        var fills = s.FindPlaceholders(PlaceholderKinds.BlankFill)
            .OrderByDescending(p => p.Match.Span.Start);
        foreach (var p in fills)
            s.ReplaceMatch(p.Match, p.Match.Text == "$[___]" ? "$42" : "Bluth Co.");

        var flat = FlatBodyText(s);
        Assert.Contains("Name: Bluth Co. and price: $42", flat);
        // Remaining placeholders untouched.
        Assert.Contains("[There shall be no cumulative voting.]", flat);
        Assert.Contains("[insert percentage]", flat);
    }

    // ─── Annotation-based anchor discovery (#132) ────────────────────────

    /// <summary>
    /// Builds <c>BuildDS001_SimpleTwoParagraphs</c> and adds annotations declared in
    /// <paramref name="specs"/>. Each spec is <c>(id, labelId, label, searchText)</c>
    /// — the search runs against the existing fixture text, so the supplied <c>searchText</c>
    /// must appear once. Returns bytes ready to feed to <see cref="DocxSession"/>.
    /// </summary>
    private static byte[] BuildDS180_TwoParaWithAnnotations(
        params (string Id, string LabelId, string Label, string SearchText)[] specs)
    {
        var wml = new WmlDocument("DS180.docx", BuildDS001_SimpleTwoParagraphs());
        foreach (var (id, labelId, label, searchText) in specs)
        {
            wml = AnnotationManager.AddAnnotation(
                wml,
                new DocumentAnnotation(id, labelId, label, "#FFEB3B"),
                AnnotationRange.FromSearch(searchText));
        }
        return wml.DocumentByteArray;
    }

    [Fact]
    public void DS180_FindByAnnotation_ResolvesBookmarkInsideSingleParagraph()
    {
        // Annotation covers "First paragraph." inside the first paragraph. The bookmark
        // sits between two markers within that paragraph; ResolveBookmarkAnchors should
        // return exactly that one paragraph anchor (no table/cell wrappers in this doc).
        using var s = new DocxSession(
            BuildDS180_TwoParaWithAnnotations(("a1", "FIRST", "First clause", "First paragraph")));

        var anchors = s.FindByAnnotation("a1");
        Assert.Single(anchors);
        Assert.Equal("p", anchors[0].Anchor.Kind);
        Assert.Equal("body", anchors[0].Anchor.Scope);

        // The returned anchor is usable for a follow-up edit — the canonical agentic
        // recipe: find by annotation, ReplaceText. Verify the round-trip works.
        var r = s.ReplaceText(anchors[0].Anchor.Id, "Replaced first.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("Replaced first.", s.Project().Markdown);
    }

    [Fact]
    public void DS181_FindByAnnotation_MissingIdReturnsEmpty()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        Assert.Empty(s.FindByAnnotation("does-not-exist"));
        Assert.Empty(s.FindByAnnotation(""));
    }

    [Fact]
    public void DS182_FindByLabel_GroupsSameLabelAcrossMultipleAnnotations()
    {
        // Two annotations both tagged "WARRANTY", one per paragraph. FindByLabel must
        // keep them disambiguated by annotation id — same label, distinct regions.
        using var s = new DocxSession(BuildDS180_TwoParaWithAnnotations(
            ("w1", "WARRANTY", "Warranty A", "First paragraph"),
            ("w2", "WARRANTY", "Warranty B", "Second paragraph")));

        var grouped = s.FindByLabel("WARRANTY");
        Assert.Equal(2, grouped.Count);
        Assert.True(grouped.ContainsKey("w1"));
        Assert.True(grouped.ContainsKey("w2"));
        // Each annotation resolves to exactly one paragraph anchor, and they are distinct.
        var w1Anchor = Assert.Single(grouped["w1"]);
        var w2Anchor = Assert.Single(grouped["w2"]);
        Assert.NotEqual(w1Anchor.Anchor.Id, w2Anchor.Anchor.Id);
    }

    [Fact]
    public void DS183_FindByLabel_UnknownLabel_ReturnsEmptyDict()
    {
        using var s = new DocxSession(
            BuildDS180_TwoParaWithAnnotations(("a1", "ONLY_LABEL", "Only", "First paragraph")));
        Assert.Empty(s.FindByLabel("MISSING"));
        Assert.Empty(s.FindByLabel(""));
    }

    [Fact]
    public void DS184_FindByBookmark_ResolvesArbitraryBookmarkName()
    {
        // FindByBookmark accepts any bookmark name, not just the Docxodus-managed
        // _Docxodus_Ann_* ones — this is the lower-level escape hatch the issue calls out.
        using var s = new DocxSession(
            BuildDS180_TwoParaWithAnnotations(("a1", "L", "L", "First paragraph")));

        var byManagedName = s.FindByBookmark(AnnotationManager.BookmarkPrefix + "a1");
        Assert.Single(byManagedName);
        Assert.Equal("p", byManagedName[0].Anchor.Kind);

        Assert.Empty(s.FindByBookmark("no_such_bookmark"));
        Assert.Empty(s.FindByBookmark(""));
    }

    [Fact]
    public void DS185_ListAnnotations_ReturnsEveryPersistedAnnotation()
    {
        using var s = new DocxSession(BuildDS180_TwoParaWithAnnotations(
            ("a1", "L1", "Label 1", "First paragraph"),
            ("a2", "L2", "Label 2", "Second paragraph")));

        var all = s.ListAnnotations();
        Assert.Equal(2, all.Count);
        Assert.Contains(all, a => a.Id == "a1" && a.LabelId == "L1");
        Assert.Contains(all, a => a.Id == "a2" && a.LabelId == "L2");
        // AnnotatedText is populated from the bookmark — sanity check it's the right slice.
        Assert.Equal("First paragraph", all.First(a => a.Id == "a1").AnnotatedText);
    }

    [Fact]
    public void DS185b_ListAnnotations_EmptyDocument_ReturnsEmpty()
    {
        using var s = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        Assert.Empty(s.ListAnnotations());
    }

    [Fact]
    public void DS186_FindByAnnotation_MultiParagraph_ReturnsAnchorsInDocOrder()
    {
        // An index-based annotation that spans paragraph 0 → paragraph 1 — exercises the
        // "bookmark covers multiple block elements" branch where the returned anchor list
        // must include both, document-order.
        var wml = new WmlDocument("DS186.docx", BuildDS001_SimpleTwoParagraphs());
        wml = AnnotationManager.AddAnnotation(
            wml,
            new DocumentAnnotation("span", "SECTION", "Section", "#FFEB3B"),
            AnnotationRange.FromParagraphs(0, 1));

        using var s = new DocxSession(wml.DocumentByteArray);
        var anchors = s.FindByAnnotation("span");
        Assert.Equal(2, anchors.Count);
        Assert.All(anchors, a => Assert.Equal("p", a.Anchor.Kind));

        // Doc order check: the first anchor's preview matches the first paragraph text.
        var firstInfo = s.GetAnchorInfo(anchors[0].Anchor.Id);
        var secondInfo = s.GetAnchorInfo(anchors[1].Anchor.Id);
        Assert.NotNull(firstInfo);
        Assert.NotNull(secondInfo);
        Assert.Contains("First paragraph", firstInfo!.TextPreview);
        Assert.Contains("Second paragraph", secondInfo!.TextPreview);
    }

    [Fact]
    public void DS187_FindByAnnotation_BookmarkInsideCell_IncludesEnclosingBlocks()
    {
        // Annotation lands inside a table cell paragraph — the returned anchor list
        // should include the cell paragraph plus the enclosing tbl/tr/tc anchors so an
        // agent can see "this annotation lives in a table" without re-walking the tree.
        var wml = new WmlDocument("DS187.docx", BuildDS003_TableWithCells());
        wml = AnnotationManager.AddAnnotation(
            wml,
            new DocumentAnnotation("cellAnn", "CELL", "Cell", "#FFEB3B"),
            AnnotationRange.FromSearch("R0C0"));

        using var s = new DocxSession(wml.DocumentByteArray);
        var anchors = s.FindByAnnotation("cellAnn");

        Assert.NotEmpty(anchors);
        // The paragraph inside the cell holds the bookmark; check it's present.
        Assert.Contains(anchors, a => a.Anchor.Kind == "p");
        // Table-level anchors are emitted for the enclosing structure.
        Assert.Contains(anchors, a => a.Anchor.Kind == "tc");
        Assert.Contains(anchors, a => a.Anchor.Kind == "tbl");
    }

    // ─── Phase: GrepCrossBlock (#146) ──────────────────────────────────────
    //
    // Fixture layout (body paragraphs in order):
    //   P0: "Section 1.1. Hello "
    //   P1: "world! Trailing text."
    //   P2: ""                                  (empty paragraph between clauses)
    //   P3: "Indemnification. The Company "
    //   P4: "shall indemnify all officers."
    //   [Table here] — interrupts the run; body paragraphs after table
    //                  belong to a separate group.
    //   P5: "Post-table paragraph."
    // The table cell contains two paragraphs of its own so cross-block can be
    // tested inside a cell, and confirms body→cell never bridges.

    internal static byte[] BuildDS200_CrossBlockFixture()
    {
        XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            XElement Para(params XElement[] runs) => new XElement(W + "p", runs);
            XElement Run(string text, bool bold = false)
            {
                var r = new XElement(W + "r");
                if (bold) r.Add(new XElement(W + "rPr", new XElement(W + "b")));
                r.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
                return r;
            }

            var cellPara1 = Para(Run("Cell line one. "));
            var cellPara2 = Para(Run("Cell line two."));

            var table = new XElement(W + "tbl",
                new XElement(W + "tr",
                    new XElement(W + "tc", cellPara1, cellPara2)));

            var body = new XElement(W + "body",
                Para(Run("Section 1.1. Hello ")),
                Para(Run("world! Trailing text.")),
                Para(), // empty paragraph
                Para(Run("Indemnification. The Company ")),
                Para(Run("shall ", bold: true), Run("indemnify all officers.")),
                table,
                Para(Run("Post-table paragraph.")));

            var doc = new XElement(W + "document",
                new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                body);
            main.PutXDocument(new XDocument(doc));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS200_GrepCrossBlock_MatchSpanningTwoAdjacentParagraphs()
    {
        // "Hello \nworld!" with a single-newline separator between adjacent blocks.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var matches = s.GrepCrossBlock("Hello \nworld!");
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal(2, m.Slices.Count);
        Assert.Equal(2, m.EnclosingAnchors.Count);
        Assert.NotEqual(m.EnclosingAnchors[0].Anchor.Id, m.EnclosingAnchors[1].Anchor.Id);

        Assert.Equal("Hello ", m.Slices[0].Fragments.Single().Text);
        Assert.Equal("world!", m.Slices[1].Fragments.Single().Text);
        Assert.Equal(13, m.Slices[0].SpanInBlock.Start); // "Section 1.1. " = 13 chars
        Assert.Equal(6, m.Slices[0].SpanInBlock.Length);
        Assert.Equal(0, m.Slices[1].SpanInBlock.Start);
        Assert.Equal(6, m.Slices[1].SpanInBlock.Length);
    }

    [Fact]
    public void DS201_GrepCrossBlock_DotDoesNotCrossBoundaryByDefault()
    {
        // Without Singleline, "." does not match the '\n' boundary character.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var noFlag = s.GrepCrossBlock(@"Hello.*world");
        Assert.Empty(noFlag);

        var withSingleline = s.GrepCrossBlock(@"Hello.*world", System.Text.RegularExpressions.RegexOptions.Singleline);
        Assert.Single(withSingleline);
        Assert.Equal(2, withSingleline[0].Slices.Count);
    }

    [Fact]
    public void DS202_GrepCrossBlock_IncludesSingleBlockMatchesAsSuperset()
    {
        // A purely intra-block match still surfaces, with exactly one slice.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var m = s.GrepCrossBlock("Section 1.1").Single();
        Assert.Single(m.Slices);
        Assert.Single(m.EnclosingAnchors);
    }

    [Fact]
    public void DS203_GrepCrossBlock_PreservesFormattingPerSlice()
    {
        // Match crosses bold + non-bold runs within P4, AND extends back into P3.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var matches = s.GrepCrossBlock(@"Company \nshall indemnify");
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal(2, m.Slices.Count);

        // P3 slice: "Company " trailing — single run, not bold.
        Assert.Equal("Company ", m.Slices[0].Fragments.Single().Text);
        Assert.False(m.Slices[0].Fragments.Single().Formatting.Bold);

        // P4 slice: "shall indemnify" — spans the bold "shall " and the plain "indemnify".
        Assert.Equal(2, m.Slices[1].Fragments.Count);
        Assert.Equal("shall ", m.Slices[1].Fragments[0].Text);
        Assert.True(m.Slices[1].Fragments[0].Formatting.Bold);
        Assert.Equal("indemnify", m.Slices[1].Fragments[1].Text);
        Assert.False(m.Slices[1].Fragments[1].Formatting.Bold);
    }

    [Fact]
    public void DS204_GrepCrossBlock_DoesNotBridgeAcrossTable()
    {
        // P4 ends with "officers." and P5 (after the table) starts with "Post-table".
        // The table interrupts the run, so a cross-block match across them is impossible.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var matches = s.GrepCrossBlock(@"officers\.\nPost-table", System.Text.RegularExpressions.RegexOptions.Singleline);
        Assert.Empty(matches);
    }

    [Fact]
    public void DS205_GrepCrossBlock_DoesNotBridgeBodyAndCellParagraphs()
    {
        // The body paragraph before the table and the first cell paragraph share no
        // direct parent — a match across them must not appear.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var matches = s.GrepCrossBlock(@"officers\.\nCell line", System.Text.RegularExpressions.RegexOptions.Singleline);
        Assert.Empty(matches);
    }

    [Fact]
    public void DS206_GrepCrossBlock_MatchesAcrossEmptyParagraph()
    {
        // P1 → P2 (empty) → P3: the empty paragraph contributes a recorded slice
        // (with zero fragments) so the agent sees the boundary the match crossed.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var matches = s.GrepCrossBlock(@"Trailing text\.\n\nIndemnification", System.Text.RegularExpressions.RegexOptions.Singleline);
        Assert.Single(matches);

        var m = matches[0];
        Assert.Equal(3, m.Slices.Count);
        Assert.Equal("Trailing text.", m.Slices[0].Fragments.Single().Text);
        Assert.Empty(m.Slices[1].Fragments); // empty paragraph contributed nothing
        Assert.Equal(0, m.Slices[1].SpanInBlock.Length);
        Assert.Equal("Indemnification", m.Slices[2].Fragments.Single().Text);
    }

    [Fact]
    public void DS207_GrepCrossBlock_MatchesWithinTableCell()
    {
        // Two paragraphs inside the same cell are siblings — cross-block within
        // the cell must work.
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var matches = s.GrepCrossBlock(@"Cell line one\. \nCell line two");
        Assert.Single(matches);
        Assert.Equal(2, matches[0].Slices.Count);
    }

    [Fact]
    public void DS208_GrepCrossBlock_ContextSurroundsTheMatch()
    {
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        var m = s.GrepCrossBlock("Hello \nworld!").Single();
        Assert.EndsWith("Section 1.1. ", m.ContextBefore);
        Assert.StartsWith(" Trailing text.", m.ContextAfter);
    }

    [Fact]
    public void DS209_GrepCrossBlock_EmptyPatternReturnsEmpty()
    {
        using var s = new DocxSession(BuildDS200_CrossBlockFixture());
        Assert.Empty(s.GrepCrossBlock(""));
    }

    // ─── DeleteRange (issue #165 — DS260-DS265) ───────────────────────────

    internal static byte[] BuildDocFiveBodyParagraphs()
    {
        // Five top-level body paragraphs: para A through E. Used by DeleteRange
        // tests that need a deterministic forward sequence of sibling blocks.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Paragraph A"))),
                new Paragraph(new Run(new Text("Paragraph B"))),
                new Paragraph(new Run(new Text("Paragraph C"))),
                new Paragraph(new Run(new Text("Paragraph D"))),
                new Paragraph(new Run(new Text("Paragraph E")))));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS260_DeleteRange_RemovesContiguousBlocksInOneCall()
    {
        using var session = new DocxSession(BuildDocFiveBodyParagraphs());
        var body = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
            .ToList();
        // Delete B, C, D (indices 1..3 inclusive). DeleteRange's "to" is exclusive,
        // so pass anchor[1] as from and anchor[4] as toExclusive.
        var r = session.DeleteRange(body[1].Anchor.Id, body[4].Anchor.Id);
        Assert.True(r.Success);
        Assert.Equal(3, r.Removed.Count);

        var afterMd = session.Project().Markdown;
        Assert.Contains("Paragraph A", afterMd);
        Assert.DoesNotContain("Paragraph B", afterMd);
        Assert.DoesNotContain("Paragraph C", afterMd);
        Assert.DoesNotContain("Paragraph D", afterMd);
        Assert.Contains("Paragraph E", afterMd);
    }

    [Fact]
    public void DS261_DeleteRange_FromMustPrecedeToInDocumentOrder()
    {
        using var session = new DocxSession(BuildDocFiveBodyParagraphs());
        var body = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
            .ToList();
        // Reversed args: from = E, to = B (from comes after to)
        var r = session.DeleteRange(body[4].Anchor.Id, body[1].Anchor.Id);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidPosition, r.Error?.Code);
    }

    [Fact]
    public void DS262_DeleteRange_UnknownFromAnchorReturnsAnchorNotFound()
    {
        using var session = new DocxSession(BuildDocFiveBodyParagraphs());
        var body = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.DeleteRange("p:body:0000000000000000ffffffffffffffff", body.Anchor.Id);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, r.Error?.Code);
    }

    [Fact]
    public void DS263_DeleteRange_WrongKindReturnsAnchorWrongKind()
    {
        using var session = new DocxSession(BuildDocFiveBodyParagraphs());
        var body = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var sec = session.Project().AnchorIndex.Values
            .FirstOrDefault(t => t.Anchor.Kind == "sec");
        if (sec is null) return;   // Five-paragraph fixture may have no sectPr; skip the assertion if so.
        var r = session.DeleteRange(body.Anchor.Id, sec.Anchor.Id);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, r.Error?.Code);
    }

    [Fact]
    public void DS264_DeleteRange_AnchorsInDifferentPartsRefused()
    {
        // Use the FootnotesPart fixture from #162 to get a fn-scope anchor.
        using var session = new DocxSession(BuildDocWithFootnotes());
        var bodyAnchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p").Anchor.Id;
        var fnAnchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "fn" && t.Anchor.Kind == "fn").Anchor.Id;
        var r = session.DeleteRange(bodyAnchor, fnAnchor);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorsNotAdjacent, r.Error?.Code);
    }

    [Fact]
    public void DS265_DeleteRange_UndoRestoresEntireRange()
    {
        using var session = new DocxSession(BuildDocFiveBodyParagraphs());
        var body = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
            .ToList();
        var r = session.DeleteRange(body[1].Anchor.Id, body[4].Anchor.Id);
        Assert.True(r.Success);

        var undo = session.Undo();
        Assert.True(undo);

        var afterMd = session.Project().Markdown;
        Assert.Contains("Paragraph A", afterMd);
        Assert.Contains("Paragraph B", afterMd);
        Assert.Contains("Paragraph C", afterMd);
        Assert.Contains("Paragraph D", afterMd);
        Assert.Contains("Paragraph E", afterMd);
    }

    [Fact]
    public void DS266_HeadingLevelHelpersExposedAsInternal()
    {
        // Regression guard: DocxSession.DeleteSection (Task 5) needs to consume these
        // helpers from WmlToMarkdownConverter. If a future refactor demotes them back
        // to private, the build of DeleteSection breaks — this test fails to build.
        var p = new XElement(W.p,
            new XElement(W.pPr,
                new XElement(W.pStyle, new XAttribute(W.val, "Heading2"))));
        Assert.True(WmlToMarkdownConverter.IsHeading(p));
        Assert.Equal(2, WmlToMarkdownConverter.HeadingLevel(p));
    }

    // ─── DeleteSection (issue #165 — DS267-DS270) ─────────────────────────

    internal static byte[] BuildDocWithHeadingSections()
    {
        // Two top-level sections under Heading1, with a Heading2 inside each.
        //
        //   # Section One
        //   para 1.1
        //   ## Section One.A
        //   para 1.A.1
        //   # Section Two
        //   para 2.1
        //
        // DeleteSection on "Section One" should remove the heading and everything down
        // to (but not including) "Section Two". DeleteSection on "Section Two"
        // (the last heading) should remove "Section Two" + "para 2.1".
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var stylesXml = """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
                  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>
                </w:styles>
                """;
            using (var s = stylesPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s)) w.Write(stylesXml);

            Paragraph H(string text, string styleId) => new(
                new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
                new Run(new Text(text)));
            Paragraph P(string text) => new(new Run(new Text(text)));

            main.Document = new Document(new Body(
                H("Section One", "Heading1"),
                P("para 1.1"),
                H("Section One.A", "Heading2"),
                P("para 1.A.1"),
                H("Section Two", "Heading1"),
                P("para 2.1")));
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS267_DeleteSection_RemovesHeadingThroughNextSameOrHigherLevel()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var sectionOne = session.Project().AnchorIndex.Values
            .Single(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section One");
        var r = session.DeleteSection(sectionOne.Anchor.Id);
        Assert.True(r.Success);
        // Removed: "Section One", "para 1.1", "Section One.A", "para 1.A.1" = 4 blocks.
        Assert.True(r.Removed.Count >= 4);

        var afterMd = session.Project().Markdown;
        Assert.DoesNotContain("Section One.A", afterMd);
        Assert.DoesNotContain("para 1.A.1", afterMd);
        Assert.DoesNotContain("para 1.1", afterMd);
        Assert.Contains("Section Two", afterMd);
        Assert.Contains("para 2.1", afterMd);
    }

    [Fact]
    public void DS268_DeleteSection_LastSectionExtendsToEndOfParent()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var sectionTwo = session.Project().AnchorIndex.Values
            .Single(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section Two");
        var r = session.DeleteSection(sectionTwo.Anchor.Id);
        Assert.True(r.Success);

        var afterMd = session.Project().Markdown;
        Assert.Contains("Section One", afterMd);
        Assert.Contains("para 1.1", afterMd);
        Assert.DoesNotContain("Section Two", afterMd);
        Assert.DoesNotContain("para 2.1", afterMd);
    }

    [Fact]
    public void DS269_DeleteSection_NonHeadingAnchorReturnsAnchorWrongKind()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var nonHeading = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.DeleteSection(nonHeading.Anchor.Id);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, r.Error?.Code);
    }

    [Fact]
    public void DS270_DeleteSection_NestedHeadingDoesNotEatNextSameLevelSection()
    {
        // Deleting "Section One.A" (Heading2) should remove ONLY the H2 and its child
        // ("para 1.A.1"), not bleed into the next Heading1 ("Section Two") which is a
        // higher-level boundary.
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var sectionOneA = session.Project().AnchorIndex.Values
            .Single(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section One.A");
        var r = session.DeleteSection(sectionOneA.Anchor.Id);
        Assert.True(r.Success);

        var afterMd = session.Project().Markdown;
        Assert.Contains("Section One", afterMd);
        Assert.Contains("para 1.1", afterMd);
        Assert.DoesNotContain("Section One.A", afterMd);
        Assert.DoesNotContain("para 1.A.1", afterMd);
        Assert.Contains("Section Two", afterMd);
        Assert.Contains("para 2.1", afterMd);
    }

    // ─── Tracked-change-aware bulk delete (issue #177 — DS271-DS273) ─────

    [Fact]
    public void DS271_DeleteRange_TrackedChangeWrapsRunsInDel()
    {
        // In RenderInline mode, DeleteRange should wrap every run in every removed
        // paragraph in w:del AND mark each paragraph-mark deleted via w:pPr/w:rPr/w:del.
        // The blocks must stay in the document tree so callers can re-address them
        // before accepting the changes — EditResult.Modified, not Removed.
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "range-deleter",
        };
        using var session = new DocxSession(BuildDocFiveBodyParagraphs(), settings);
        var body = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
            .ToList();
        // Delete B, C, D (indices 1..3 inclusive).
        var r = session.DeleteRange(body[1].Anchor.Id, body[4].Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);

        // Tracked-mode contract: anchors stay live, so Modified is populated and
        // Removed is empty. Three top-level paragraphs were marked.
        Assert.Empty(r.Removed);
        Assert.Equal(3, r.Modified.Count);
        Assert.Contains(r.Modified, a => a.Id == body[1].Anchor.Id);
        Assert.Contains(r.Modified, a => a.Id == body[2].Anchor.Id);
        Assert.Contains(r.Modified, a => a.Id == body[3].Anchor.Id);

        // Round-trip to bytes and inspect the XML directly.
        var bytes = session.Save();
        using var ms = new MemoryStream(bytes);
        using var verify = WordprocessingDocument.Open(ms, isEditable: false);
        var docXml = verify.MainDocumentPart!.GetXDocument().Root!.ToString();

        // Every run-text from B/C/D should land in w:delText, not w:t.
        Assert.Contains("Paragraph B", docXml);
        Assert.Contains("Paragraph C", docXml);
        Assert.Contains("Paragraph D", docXml);
        Assert.Contains("w:delText", docXml);
        Assert.Contains("range-deleter", docXml);

        // Paragraph-mark deletion: each of the three deleted paragraphs must have a
        // w:pPr/w:rPr/w:del element (counted by DescendantsAndSelf below).
        var root = verify.MainDocumentPart.GetXDocument().Root!;
        var paraMarkDels = root
            .Descendants(W.p)
            .Where(p => p.Element(W.pPr)?.Element(W.rPr)?.Element(W.del) is not null)
            .ToList();
        Assert.Equal(3, paraMarkDels.Count);

        // A (untouched) and E (the exclusive endpoint) still have plain w:t runs.
        Assert.Contains("<w:t>Paragraph A</w:t>", docXml);
        Assert.Contains("<w:t>Paragraph E</w:t>", docXml);
    }

    [Fact]
    public void DS272_DeleteSection_TrackedChangeWrapsRunsInDel()
    {
        // DeleteSection in tracked mode should inherit DeleteRange's behavior via the
        // shared DeleteSiblingRangeCore helper: heading + every sibling under it through
        // the next same-or-higher heading get w:del-wrapped, not structurally removed.
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "section-deleter",
        };
        using var session = new DocxSession(BuildDocWithHeadingSections(), settings);
        var sectionOne = session.Project().AnchorIndex.Values
            .Single(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section One");

        var r = session.DeleteSection(sectionOne.Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);

        // Tracked mode: Modified is populated (anchors live), Removed is empty.
        // Section One spans 4 blocks (heading + 3 followers up to the next Heading1).
        Assert.Empty(r.Removed);
        Assert.Equal(4, r.Modified.Count);

        // The projection still sees the heading text (anchors are live) — content is
        // technically still there, just wrapped in tracked-deletion markup.
        var afterIndex = session.Project().AnchorIndex;
        Assert.True(afterIndex.ContainsKey(sectionOne.Anchor.Id));

        // Round-trip and inspect XML. Section One + para 1.1 + Section One.A + para 1.A.1
        // should all be wrapped; Section Two + para 2.1 must remain untouched.
        var bytes = session.Save();
        using var ms = new MemoryStream(bytes);
        using var verify = WordprocessingDocument.Open(ms, isEditable: false);
        var root = verify.MainDocumentPart!.GetXDocument().Root!;
        var docXml = root.ToString();

        Assert.Contains("w:delText", docXml);
        Assert.Contains("section-deleter", docXml);
        // Section Two is the boundary marker (exclusive) — its run text must not be in w:delText.
        Assert.Contains("<w:t>Section Two</w:t>", docXml);
        Assert.Contains("<w:t>para 2.1</w:t>", docXml);

        // Each of the 4 deleted paragraphs has a paragraph-mark del.
        var paraMarkDels = root
            .Descendants(W.p)
            .Where(p => p.Element(W.pPr)?.Element(W.rPr)?.Element(W.del) is not null)
            .ToList();
        Assert.Equal(4, paraMarkDels.Count);
    }

    [Fact]
    public void DS273_DeleteRange_TrackedChange_TableRowsMarkedDeleted()
    {
        // Tables in the range get w:trPr/w:del on every row plus run-wrapping inside
        // every cell paragraph — Word's native convention for tracked row deletion.
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "table-deleter",
        };
        using var session = new DocxSession(BuildDocFiveBlocksWithTable(), settings);
        var blocks = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "tbl")
            .OrderBy(t => t.Anchor.Id)
            .ToList();
        // Layout: para0, para1, tbl, para2, para3.
        // Delete [para1, tbl, para2] — exclusive endpoint = para3.
        var tbl = session.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "tbl");
        var paras = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .ToList();

        // Resolve by text preview for stability across Unid renumbering.
        var p0 = paras.Single(t => session.GetAnchorInfo(t.Anchor.Id)!.TextPreview == "Para Zero");
        var p1 = paras.Single(t => session.GetAnchorInfo(t.Anchor.Id)!.TextPreview == "Para One");
        var p2 = paras.Single(t => session.GetAnchorInfo(t.Anchor.Id)!.TextPreview == "Para Two");
        var p3 = paras.Single(t => session.GetAnchorInfo(t.Anchor.Id)!.TextPreview == "Para Three");

        var r = session.DeleteRange(p1.Anchor.Id, p3.Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);

        // Modified: p1, tbl, p2.
        Assert.Empty(r.Removed);
        Assert.Equal(3, r.Modified.Count);
        Assert.Contains(r.Modified, a => a.Id == p1.Anchor.Id);
        Assert.Contains(r.Modified, a => a.Id == tbl.Anchor.Id);
        Assert.Contains(r.Modified, a => a.Id == p2.Anchor.Id);

        var bytes = session.Save();
        using var ms = new MemoryStream(bytes);
        using var verify = WordprocessingDocument.Open(ms, isEditable: false);
        var root = verify.MainDocumentPart!.GetXDocument().Root!;
        var docXml = root.ToString();

        // Every row gets a trPr/del marker.
        var rowDels = root
            .Descendants(W.tr)
            .Where(tr => tr.Element(W.trPr)?.Element(W.del) is not null)
            .ToList();
        Assert.Equal(2, rowDels.Count);  // two rows in BuildDocFiveBlocksWithTable

        // Cell text wrapped in w:delText.
        Assert.Contains("w:delText", docXml);
        Assert.Contains("Cell A1", docXml);
        Assert.Contains("Cell B2", docXml);
        Assert.Contains("table-deleter", docXml);

        // Untouched paragraphs at the ends remain.
        Assert.Contains("<w:t>Para Zero</w:t>", docXml);
        Assert.Contains("<w:t>Para Three</w:t>", docXml);
    }

    internal static byte[] BuildDocFiveBlocksWithTable()
    {
        // para0, para1, table (2x2), para2, para3 — for tracked-delete table coverage.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            body.Append(new Paragraph(new Run(new Text("Para Zero"))));
            body.Append(new Paragraph(new Run(new Text("Para One"))));

            var table = new DocumentFormat.OpenXml.Wordprocessing.Table();
            string[][] cellText = { new[] { "Cell A1", "Cell A2" }, new[] { "Cell B1", "Cell B2" } };
            foreach (var rowText in cellText)
            {
                var tr = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                foreach (var cellContent in rowText)
                {
                    var tc = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    tc.Append(new Paragraph(new Run(new Text(cellContent))));
                    tr.Append(tc);
                }
                table.Append(tr);
            }
            body.Append(table);

            body.Append(new Paragraph(new Run(new Text("Para Two"))));
            body.Append(new Paragraph(new Run(new Text("Para Three"))));
        }
        return ms.ToArray();
    }

    // ─── FindOptions.Scopes + AnchorsByScope ────────────────────────────

    /// <summary>
    /// Build a doc with one body paragraph + one header part + one footer part,
    /// so scope-filter tests can confirm narrowing behavior.
    /// </summary>
    internal static byte[] BuildDocWithHeaderAndFooter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            body.Append(new Paragraph(new Run(new Text("BODY-NEEDLE marker"))));

            var hp = mainPart.AddNewPart<HeaderPart>();
            hp.Header = new Header(new Paragraph(new Run(new Text("HEADER-NEEDLE confidential"))));
            var hRef = new HeaderReference { Id = mainPart.GetIdOfPart(hp), Type = HeaderFooterValues.Default };

            var fp = mainPart.AddNewPart<FooterPart>();
            fp.Footer = new Footer(new Paragraph(new Run(new Text("FOOTER-NEEDLE page x"))));
            var fRef = new FooterReference { Id = mainPart.GetIdOfPart(fp), Type = HeaderFooterValues.Default };

            var sectPr = new SectionProperties(hRef, fRef);
            body.Append(sectPr);
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS290_AnchorsByScope_BodyByDefault()
    {
        using var session = new DocxSession(BuildDocWithHeaderAndFooter());
        var bodyAnchors = session.AnchorsByScope(ProjectionScopes.Body);
        Assert.NotEmpty(bodyAnchors);
        Assert.All(bodyAnchors, t => Assert.Equal("body", t.Anchor.Scope));
    }

    [Fact]
    public void DS291_AnchorsByScope_HeadersFootersCompose()
    {
        using var session = new DocxSession(BuildDocWithHeaderAndFooter());
        var hdrFtr = session.AnchorsByScope(ProjectionScopes.Headers | ProjectionScopes.Footers);
        Assert.NotEmpty(hdrFtr);
        Assert.All(hdrFtr, t =>
            Assert.True(t.Anchor.Scope.StartsWith("hdr") || t.Anchor.Scope.StartsWith("ftr"),
                $"Expected hdr*/ftr*, got {t.Anchor.Scope}"));
        // Body anchors must NOT appear.
        Assert.DoesNotContain(hdrFtr, t => t.Anchor.Scope == "body");
    }

    [Fact]
    public void DS292_FindOptions_Scopes_NarrowsFindAllByText()
    {
        // Default Scopes = All. FindAllByText sees the body, header, and footer needles.
        using var session = new DocxSession(BuildDocWithHeaderAndFooter());
        var allHits = session.FindAllByText("NEEDLE");
        Assert.Equal(3, allHits.Count);

        // Narrow to Headers only — only the header needle survives.
        var headerOnly = session.FindAllByText("NEEDLE", new FindOptions { Scopes = ProjectionScopes.Headers });
        Assert.Single(headerOnly);
        Assert.StartsWith("hdr", headerOnly[0].Anchor.Scope);

        // Compose Body | Footers — two hits, neither is in a header.
        var bodyAndFooters = session.FindAllByText("NEEDLE",
            new FindOptions { Scopes = ProjectionScopes.Body | ProjectionScopes.Footers });
        Assert.Equal(2, bodyAndFooters.Count);
        Assert.DoesNotContain(bodyAndFooters, t => t.Anchor.Scope.StartsWith("hdr"));
    }

    [Fact]
    public void DS293_FindOptions_ScopeFilter_StringStillNarrowsWithinScopes()
    {
        // Scopes selects Headers + Footers; ScopeFilter pins one specific part.
        // The fine-grained string ScopeFilter is still respected as a further narrow.
        using var session = new DocxSession(BuildDocWithHeaderAndFooter());
        var hdr1Only = session.FindAllByText("NEEDLE", new FindOptions
        {
            Scopes = ProjectionScopes.Headers | ProjectionScopes.Footers,
            ScopeFilter = "hdr1",
        });
        Assert.Single(hdr1Only);
        Assert.Equal("hdr1", hdr1Only[0].Anchor.Scope);
    }

    [Fact]
    public void DS294_ProjectionScopesExtensions_IncludesScopeRoundtrip()
    {
        // Cheap unit test for the helper — body, hdr*, ftr*, fn, en, cmt all map.
        Assert.True(ProjectionScopes.Body.IncludesScope("body"));
        Assert.True(ProjectionScopes.Headers.IncludesScope("hdr1"));
        Assert.True(ProjectionScopes.Headers.IncludesScope("hdr27"));
        Assert.False(ProjectionScopes.Headers.IncludesScope("body"));
        Assert.True(ProjectionScopes.Footers.IncludesScope("ftr2"));
        Assert.True(ProjectionScopes.Footnotes.IncludesScope("fn"));
        Assert.True(ProjectionScopes.Endnotes.IncludesScope("en"));
        Assert.True(ProjectionScopes.Comments.IncludesScope("cmt"));
        Assert.True(ProjectionScopes.All.IncludesScope("anything-goes"));
        // Composition.
        var bodyHdr = ProjectionScopes.Body | ProjectionScopes.Headers;
        Assert.True(bodyHdr.IncludesScope("body"));
        Assert.True(bodyHdr.IncludesScope("hdr3"));
        Assert.False(bodyHdr.IncludesScope("fn"));
    }

    // ─── GetEditSummary (issue #166 — DS280-DS283) ────────────────────────

    [Fact]
    public void DS280_GetEditSummary_CountsPlaceholdersOnUneditedDoc()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var summary = session.GetEditSummary();
        Assert.True(summary.RemainingPlaceholders.Count >= 4);
        Assert.True(summary.TotalAnchors > 0);
    }

    [Fact]
    public void DS280b_GetEditSummary_BareUnderscoresExcludeBracketPlaceholders()
    {
        // BuildDocWithBracketPlaceholders contains "[_____]" and "[____________]"
        // — both are bracketed placeholders, not bare underscore runs.
        // EditSummary.BareUnderscoreRuns must be empty for this doc.
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var summary = session.GetEditSummary();
        Assert.Empty(summary.BareUnderscoreRuns);
    }

    [Fact]
    public void DS280c_GetEditSummary_BareUnderscoresPositiveCase()
    {
        // A doc with both a bare ____ and a bracketed [___] — only the bare run should match.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Fill ____ here. Also [___] should not count.")))));
        }
        using var session = new DocxSession(ms.ToArray());
        var summary = session.GetEditSummary();
        Assert.Single(summary.BareUnderscoreRuns);   // exactly the bare ____
    }

    [Fact]
    public void DS281_GetEditSummary_RemainingPlaceholdersShrinksAfterFill()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var before = session.GetEditSummary().RemainingPlaceholders.Count;
        var result = session.FillPlaceholders(p => "FILLED", new FillOptions
        {
            Kinds = PlaceholderKinds.BlankFill,
        });
        Assert.True(result.Filled > 0);
        var after = session.GetEditSummary().RemainingPlaceholders.Count;
        Assert.True(after < before, $"expected after < before; got after={after} before={before}");
    }

    [Fact]
    public void DS282_RemainingPlaceholders_MatchesFindPlaceholders()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var fromAlias = session.RemainingPlaceholders().Count;
        var fromCanonical = session.FindPlaceholders().Count;
        Assert.Equal(fromCanonical, fromAlias);
    }

    [Fact]
    public void DS283_GetEditSummary_FootnoteCountsExcludeBoilerplate()
    {
        using var session = new DocxSession(BuildDocWithFootnotes());
        var summary = session.GetEditSummary();
        Assert.Equal(1, summary.FootnoteCount);
        Assert.Equal(1, summary.InlineFootnoteRefCount);
    }

    // ─── GetDiff (issue #166 — DS284-DS289) ───────────────────────────────

    [Fact]
    public void DS284_GetDiff_OnUneditedDoc_IsEmptyArray()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var diffJson = session.GetDiff();
        // JSON empty array (no mutations have happened).
        Assert.Equal("[]", diffJson);
    }

    [Fact]
    public void DS285_GetDiff_AfterDelete_ShowsDeleteOp()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.DeleteBlock(firstP.Anchor.Id);
        Assert.True(r.Success);

        var diffJson = session.GetDiff();
        // Expect at least one entry with op=delete pointing at firstP's anchor id.
        Assert.Contains("\"delete\"", diffJson);
        Assert.Contains(firstP.Anchor.Id, diffJson);
    }

    [Fact]
    public void DS285b_GetDiff_DetectsKindChangeWithoutTextChange()
    {
        // SetParagraphStyle flips anchor kind (p→h) without changing text.
        // ComputeDiff must detect this as a modify, not silently miss it.
        // Uses BuildDS001_SimpleTwoParagraphs because it ships with the heading
        // style definitions that SetParagraphStyle needs.
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");

        var r = session.SetParagraphStyle(firstP.Anchor.Id, "Heading2");
        Assert.True(r.Success);

        var diffJson = session.GetDiff();
        Assert.Contains("\"modify\"", diffJson);
    }

    [Fact]
    public void DS286_GetDiff_AfterReplaceText_ShowsModifyOp()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        session.ReplaceText(firstP.Anchor.Id, "NEW TEXT");

        var diffJson = session.GetDiff();
        Assert.Contains("\"modify\"", diffJson);
        Assert.Contains("NEW TEXT", diffJson);
    }

    [Fact]
    public void DS287_GetDiff_AfterInsertParagraph_ShowsInsertOp()
    {
        using var session = new DocxSession(BuildDocWithBracketPlaceholders());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.InsertParagraph(firstP.Anchor.Id, Position.After, "Inserted text");
        Assert.True(r.Success);

        var diffJson = session.GetDiff();
        Assert.Contains("\"insert\"", diffJson);
        Assert.Contains("Inserted text", diffJson);
    }

    [Fact]
    public void DS288_GetDiff_WithoutInitialCapture_Throws()
    {
        var settings = new DocxSessionSettings { CaptureInitialProjection = false };
        using var session = new DocxSession(BuildDocWithBracketPlaceholders(), settings);
        var ex = Assert.Throws<InvalidOperationException>(() => session.GetDiff());
        Assert.Contains("CaptureInitialProjection", ex.Message);
    }

    // ─── GetDiff line-based formats (issue #178 — DS289-DS289d) ──────────

    [Fact]
    public void DS289_GetDiff_UnifiedFormat_AfterReplaceText_ProducesPatchCompatibleDiff()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.ReplaceText(firstP.Anchor.Id, "REPLACED PARAGRAPH");
        Assert.True(r.Success);

        var unified = session.GetDiff(DiffFormat.Unified);

        // Standard unified-diff structural markers — what patch(1) parses.
        Assert.StartsWith("--- initial\n+++ current\n", unified);
        Assert.Contains("@@ ", unified);
        // Hunk header shape: @@ -<a>,<b> +<c>,<d> @@
        Assert.Matches(@"@@ -\d+,\d+ \+\d+,\d+ @@", unified);
        Assert.Contains("REPLACED PARAGRAPH", unified);

        // Must have at least one '-' (deleted) line and one '+' (added) line in
        // the body of the hunk (excluding the '---' / '+++' headers).
        var bodyLines = unified.Split('\n').Skip(2).ToList();
        Assert.Contains(bodyLines, l => l.StartsWith("-", StringComparison.Ordinal));
        Assert.Contains(bodyLines, l => l.StartsWith("+", StringComparison.Ordinal));
    }

    [Fact]
    public void DS289a_GetDiff_UnifiedFormat_OnUneditedDoc_IsEmptyString()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var unified = session.GetDiff(DiffFormat.Unified);
        Assert.Equal(string.Empty, unified);
    }

    [Fact]
    public void DS289b_GetDiff_UnifiedFormat_AfterInsertParagraph_HasInsertHunk()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.InsertParagraph(firstP.Anchor.Id, Position.After, "Brand new paragraph");
        Assert.True(r.Success);

        var unified = session.GetDiff(DiffFormat.Unified);

        Assert.Contains("Brand new paragraph", unified);
        // Inserted line must appear with a leading '+'.
        var bodyLines = unified.Split('\n').Skip(2);
        Assert.Contains(bodyLines, l => l.StartsWith("+", StringComparison.Ordinal)
            && l.Contains("Brand new paragraph"));
    }

    [Fact]
    public void DS289c_GetDiff_SideBySideFormat_AfterReplaceText_ShowsModifyMarker()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.ReplaceText(firstP.Anchor.Id, "REPLACED PARAGRAPH");
        Assert.True(r.Success);

        var sideBySide = session.GetDiff(DiffFormat.SideBySide);

        Assert.Contains("REPLACED PARAGRAPH", sideBySide);
        // Adjacent delete/insert collapses to a '|' modify row.
        Assert.Contains(" | ", sideBySide);

        // Each row's left column should be padded to a fixed width so the marker
        // column lines up — verify by checking the marker is at the same offset on
        // every non-empty row.
        var rows = sideBySide.Split('\n').Where(l => l.Length > 0).ToList();
        Assert.NotEmpty(rows);
        var markerOffsets = rows.Select(r => r.IndexOfAny(new[] { ' ', '|', '<', '>' }, 70)).Distinct().ToList();
        // All rows should have a marker at column 72 (the column width), surrounded
        // by spaces — meaning index 72 is a space and index 73 is the marker.
        Assert.All(rows, row =>
        {
            Assert.True(row.Length >= 74, $"row too short: {row}");
            Assert.Equal(' ', row[72]);
            Assert.Contains(row[73], " |<>");
            Assert.Equal(' ', row[74]);
        });
    }

    [Fact]
    public void DS289d_GetDiff_SideBySideFormat_AfterInsertParagraph_HasInsertMarker()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        var r = session.InsertParagraph(firstP.Anchor.Id, Position.After, "Inserted block");
        Assert.True(r.Success);

        var sideBySide = session.GetDiff(DiffFormat.SideBySide);

        Assert.Contains("Inserted block", sideBySide);
        // Pure insert (no matching delete) renders as '>' on the right.
        Assert.Contains(" > ", sideBySide);
    }

    [Fact]
    public void DS289e_GetDiff_InvalidDiffFormat_Throws()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        // Out-of-range enum values still throw, just like an unrecognized switch arm.
        Assert.Throws<NotSupportedException>(() => session.GetDiff((DiffFormat)99));
    }

    /// <summary>
    /// Build a doc whose three header parts each contain a single empty paragraph.
    /// The deterministic Unid scheme seeds each scope's root with the root element's
    /// local name ("hdr" for every header part), so the first-paragraph children of
    /// these structurally-identical headers all hash to the same raw Unid — across
    /// scopes. Reproduces issue #187 on a minimal in-memory fixture.
    /// </summary>
    private static byte[] BuildDocWithCollidingHeaderUnids()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            body.Append(new Paragraph(new Run(new Text("Body content"))));

            // Three header parts, each with an identically-structured empty paragraph.
            // Same parent-root local name ("hdr") + same child tag (p) + same content
            // signature (empty paragraph) + same dup index (0, only child) ⇒ identical
            // deterministic Unid across all three parts.
            HeaderReference? firstRef = null;
            for (var i = 0; i < 3; i++)
            {
                var hp = mainPart.AddNewPart<HeaderPart>();
                hp.Header = new Header(new Paragraph());
                var hRef = new HeaderReference
                {
                    Id = mainPart.GetIdOfPart(hp),
                    Type = i == 0 ? HeaderFooterValues.First : i == 1 ? HeaderFooterValues.Default : HeaderFooterValues.Even,
                };
                firstRef ??= hRef;
                body.Append(new SectionProperties(hRef));
            }
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS289a_GetDiff_DoesNotThrowOnCrossScopeUnidCollision()
    {
        // Regression for #187: ComputeDiff used to ToDictionary(t => t.Unid) which
        // threw ArgumentException when two scopes legitimately produced the same
        // raw Unid for paragraphs at the same position in identically-structured
        // header parts. Reproduced on the NVCA Model COI; this synthetic fixture
        // triggers the same hash collision deterministically.
        using var session = new DocxSession(BuildDocWithCollidingHeaderUnids());

        // Sanity: confirm the fixture actually produces a Unid collision across scopes.
        var hdrTargets = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope.StartsWith("hdr"))
            .ToList();
        Assert.True(hdrTargets.Count >= 2, "Fixture must have multiple header anchors.");
        Assert.True(hdrTargets.Select(t => t.Unid).Distinct().Count() < hdrTargets.Count,
            "Fixture must produce at least one Unid collision across header scopes.");

        // The bug: GetDiff() used to throw ArgumentException here. Must return "[]"
        // (no mutations) without throwing.
        var diff = session.GetDiff();
        Assert.Equal("[]", diff);
    }

    [Fact]
    public void DS289b_GetDiff_MutationOnCollidingScope_IsolatesToEditedScope()
    {
        // Correctness companion to DS289a: when several anchors share a raw Unid
        // across scopes and we edit only one, the diff must surface exactly that
        // one anchor — not silently miss the change because another scope still
        // matches the Unid, and not flag unedited siblings as modified.
        using var session = new DocxSession(BuildDocWithCollidingHeaderUnids());
        var hdrTargets = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope.StartsWith("hdr") && t.Anchor.Kind == "p")
            .ToList();
        // Pick two distinct scopes that share the same Unid.
        var collidingPair = hdrTargets
            .GroupBy(t => t.Unid).First(g => g.Count() >= 2).ToList();
        var edited = collidingPair[0];
        var untouched = collidingPair[1];

        var r = session.ReplaceText(edited.Anchor.Id, "edited header line");
        Assert.True(r.Success);

        var diff = session.GetDiff();
        Assert.Contains("\"modify\"", diff);
        Assert.Contains(edited.Anchor.Id, diff);
        Assert.DoesNotContain(untouched.Anchor.Id, diff);
    }

    [Fact]
    public void DS289c_GetDiff_DoesNotThrowUnderAbbreviatedRendering()
    {
        // Defense-in-depth for #187: the AnchorIndex is dual-keyed under
        // Abbreviated/Sequential rendering — the same AnchorTarget is reachable
        // via its full Unid key and its rendered alias key, so AnchorIndex.Values
        // enumerates each target twice. ComputeDiff must dedupe before building
        // its lookup or ToDictionary throws on every call under these modes.
        var settings = new DocxSessionSettings
        {
            CaptureInitialProjection = true,
            ProjectionSettings = new WmlToMarkdownConverterSettings
            {
                AnchorIdRendering = AnchorIdRendering.Abbreviated,
            },
        };
        using var session = new DocxSession(BuildDocWithBracketPlaceholders(), settings);

        // Unedited: must return "[]" without throwing.
        Assert.Equal("[]", session.GetDiff());

        // Edited: must produce a well-formed diff (no duplicate-key crash).
        var firstP = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
        session.ReplaceText(firstP.Anchor.Id, "rendered-id diff path works");
        var diff = session.GetDiff();
        Assert.Contains("\"modify\"", diff);
        Assert.Contains("rendered-id diff path works", diff);
    }

    // ─── CompactRuns ─────────────────────────────────────────────────────

    /// <summary>
    /// Build a doc whose first paragraph has one empty bold run plus one real
    /// text run. The empty run is the kind of residue you get after deleting a
    /// footnote/endnote/comment reference that was the only child of a styled
    /// run, or after a refactoring pass that pulled the text out of a run but
    /// left the run shell in place.
    /// </summary>
    internal static byte[] BuildDocWithEmptyBoldRun()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            body.Append(new Paragraph(
                new Run(new RunProperties(new Bold())),                    // empty bold run
                new Run(new RunProperties(new Bold()), new Text("hi"))));  // real bold run
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS295_CompactRuns_RemovesEmptyBoldRun()
    {
        using var session = new DocxSession(BuildDocWithEmptyBoldRun());
        var r = session.CompactRuns();
        Assert.Equal(1, r.RunsRemoved);

        // After compaction, projection still contains "hi" — the real run survived.
        var md = session.Project().Markdown;
        Assert.Contains("hi", md);
    }

    [Fact]
    public void DS296_CompactRuns_Undoable()
    {
        using var session = new DocxSession(BuildDocWithEmptyBoldRun());
        var r = session.CompactRuns();
        Assert.Equal(1, r.RunsRemoved);

        Assert.True(session.Undo());

        // Second CompactRuns should find the empty run again.
        var r2 = session.CompactRuns();
        Assert.Equal(1, r2.RunsRemoved);
    }

    [Fact]
    public void DS297_CompactRuns_HonorsScope()
    {
        // Doc with one empty bold run in body. CompactRuns(Headers) shouldn't touch it.
        using var session = new DocxSession(BuildDocWithEmptyBoldRun());
        var r = session.CompactRuns(ProjectionScopes.Headers);
        Assert.Equal(0, r.RunsRemoved);

        // But CompactRuns(Body) does.
        var r2 = session.CompactRuns(ProjectionScopes.Body);
        Assert.Equal(1, r2.RunsRemoved);
    }

    [Fact]
    public void DS298_CompactRuns_ReturnsZeroWhenAlreadyCompact()
    {
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        var r = session.CompactRuns();
        Assert.Equal(0, r.RunsRemoved);
    }

    // ─── ProjectAnchor (issue #167 — renumbered to DS310-DS313 to avoid
    // collision with DS295-DS298 CompactRuns that landed on main) ─────────

    [Fact]
    public void DS310_ProjectAnchor_OnParagraph_ReturnsJustThatParagraph()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var firstPara = session.Project().AnchorIndex.Values
            .First(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "para 1.1");

        var sub = session.ProjectAnchor(firstPara.Anchor.Id);
        Assert.Contains("para 1.1", sub.Markdown);
        Assert.DoesNotContain("Section One.A", sub.Markdown);
        Assert.DoesNotContain("Section Two", sub.Markdown);
    }

    [Fact]
    public void DS311_ProjectAnchor_OnHeading_ReturnsTheWholeSection()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var sectionOne = session.Project().AnchorIndex.Values
            .First(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section One");

        var sub = session.ProjectAnchor(sectionOne.Anchor.Id);
        // The Heading1 section runs from "Section One" through everything before the next Heading1.
        Assert.Contains("Section One", sub.Markdown);
        Assert.Contains("para 1.1", sub.Markdown);
        Assert.Contains("Section One.A", sub.Markdown);
        Assert.Contains("para 1.A.1", sub.Markdown);
        // …but does NOT include "Section Two" (the next Heading1 boundary).
        Assert.DoesNotContain("Section Two", sub.Markdown);
        Assert.DoesNotContain("para 2.1", sub.Markdown);
    }

    [Fact]
    public void DS312_ProjectAnchor_OnLastHeading_ExtendsToEndOfParent()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var sectionTwo = session.Project().AnchorIndex.Values
            .First(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section Two");

        var sub = session.ProjectAnchor(sectionTwo.Anchor.Id);
        Assert.Contains("Section Two", sub.Markdown);
        Assert.Contains("para 2.1", sub.Markdown);
        Assert.DoesNotContain("Section One", sub.Markdown);
    }

    [Fact]
    public void DS313_ProjectAnchor_UnknownAnchorThrows()
    {
        using var session = new DocxSession(BuildDocWithHeadingSections());
        var ex = Assert.Throws<InvalidOperationException>(() =>
            session.ProjectAnchor("p:body:0000000000000000ffffffffffffffff"));
        Assert.Contains("not found", ex.Message);
    }

    // ─── Deterministic Unids ──────────────────────────────────────────────

    [Fact]
    public void DS300_DeterministicUnids_SameBytesProduceSameAnchorIds()
    {
        // Two independent sessions over the same bytes must observe the same
        // anchor ids on the same content. Closes the cross-session non-determinism
        // foot-gun where a CLI script's anchor ids from run 1 wouldn't resolve in run 2.
        var bytes = BuildDS001_SimpleTwoParagraphs();

        using var s1 = new DocxSession(bytes);
        using var s2 = new DocxSession(bytes);

        var ids1 = s1.Project().AnchorIndex.Keys.OrderBy(k => k).ToArray();
        var ids2 = s2.Project().AnchorIndex.Keys.OrderBy(k => k).ToArray();

        Assert.Equal(ids1, ids2);
        Assert.NotEmpty(ids1);
    }

    [Fact]
    public void DS301_DeterministicUnids_DuplicateContentSiblingsGetDistinctIds()
    {
        // Two paragraphs with identical text must still get distinct Unids — the
        // dup_index disambiguator handles same-content siblings.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Hello world"))),
                new Paragraph(new Run(new Text("Hello world")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
        }
        using var session = new DocxSession(ms.ToArray());
        var ids = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id)
            .ToList();
        Assert.Equal(2, ids.Count);
        Assert.NotEqual(ids[0], ids[1]);
    }

    [Fact]
    public void DS302_DeterministicUnids_UniqueContentInsertionDoesNotShiftSiblings()
    {
        // Build a doc with three paragraphs: A, B, C. Capture their Unids.
        // Then build a sibling doc with A, NEW, B, C — A's, B's, and C's Unids
        // should be identical to the first doc. The inserted NEW has unique
        // content so it shouldn't shift dup_indices.
        byte[] BuildDoc(params string[] paras)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var body = new Body();
                foreach (var p in paras) body.Append(new Paragraph(new Run(new Text(p))));
                main.Document = new Document(body);
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            }
            return ms.ToArray();
        }

        using var s1 = new DocxSession(BuildDoc("A", "B", "C"));
        using var s2 = new DocxSession(BuildDoc("A", "NEW", "B", "C"));

        Dictionary<string, string> TextToId(DocxSession s) =>
            s.Project().AnchorIndex.Values
                .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
                .ToDictionary(t => t.TextPreview, t => t.Anchor.Id);

        var map1 = TextToId(s1);
        var map2 = TextToId(s2);

        Assert.Equal(map1["A"], map2["A"]);
        Assert.Equal(map1["B"], map2["B"]);
        Assert.Equal(map1["C"], map2["C"]);
        Assert.Contains("NEW", map2.Keys);
    }

    [Fact]
    public void DS303_DeterministicUnids_LooksLike32HexString()
    {
        // The Unid format hasn't changed — still 32 lowercase hex characters.
        using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
        foreach (var t in session.Project().AnchorIndex.Values)
        {
            Assert.Equal(32, t.Unid.Length);
            Assert.Matches("^[0-9a-f]{32}$", t.Unid);
        }
    }

    [Fact]
    public void DS304_DeterministicUnids_EditingTextChangesItsIdButNotSiblings()
    {
        // Two paragraphs A and B. Open session 1, capture both Unids.
        // Build session 2 over a doc with A and B' (B's text edited) — A's Unid
        // should be unchanged; B's must differ.
        byte[] BuildDoc(string second)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("A"))),
                    new Paragraph(new Run(new Text(second)))));
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            }
            return ms.ToArray();
        }

        using var s1 = new DocxSession(BuildDoc("B"));
        using var s2 = new DocxSession(BuildDoc("B-edited"));

        var s1A = s1.Project().AnchorIndex.Values.Single(t => t.TextPreview == "A").Anchor.Id;
        var s1B = s1.Project().AnchorIndex.Values.Single(t => t.TextPreview == "B").Anchor.Id;
        var s2A = s2.Project().AnchorIndex.Values.Single(t => t.TextPreview == "A").Anchor.Id;
        var s2B = s2.Project().AnchorIndex.Values.Single(t => t.TextPreview == "B-edited").Anchor.Id;

        Assert.Equal(s1A, s2A);    // A unchanged → same id
        Assert.NotEqual(s1B, s2B); // B's text changed → new id
    }

    // ─── DS314/DS315/DS316: note references survive ReplaceText (issue B3) ──────
    //
    // ReplaceText replaces a whole block's visible text. A footnote/endnote
    // reference is a zero-width, semantically-significant marker that points into
    // the FootnotesPart/EndnotesPart — dropping it on a text edit silently orphans
    // the note definition (the destructive-accept symptom seen in the consolidate
    // footnotes-survive battery). These assert the body-side <w:footnoteReference>
    // (and endnote) survive the edit, on both the default (accept) and tracked paths.

    private static int BodyNoteRefCount<T>(byte[] bytes)
        where T : OpenXmlElement
    {
        using var s = new MemoryStream(bytes);
        using var d = WordprocessingDocument.Open(s, false);
        return d.MainDocumentPart!.Document.Body!.Descendants<T>().Count();
    }

    [Fact]
    public void DS314_ReplaceText_PreservesFootnoteReference_AcceptMode()
    {
        var baseBytes = BuildDS140_FootnoteFixture(); // "Main text"[fnRef]" continued."
        Assert.Equal(1, BodyNoteRefCount<FootnoteReference>(baseBytes));

        using var s = new DocxSession(baseBytes);
        var para = s.Project().AnchorIndex.Values.Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body");
        var r = s.ReplaceText(para.Anchor.Id, "Main text EDITED continued.");
        Assert.True(r.Success, r.Error?.Message);

        var saved = s.Save();
        Assert.Equal("Main text EDITED continued.", string.Join("\n",
            BodyParagraphTexts(saved)));
        // The body-side footnote reference survives the whole-block replace.
        Assert.Equal(1, BodyNoteRefCount<FootnoteReference>(saved));
    }

    [Fact]
    public void DS315_ReplaceText_PreservesFootnoteReference_TrackedMode()
    {
        var baseBytes = BuildDS140_FootnoteFixture();

        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Tester",
        };
        using var s = new DocxSession(baseBytes, settings);
        var para = s.Project().AnchorIndex.Values.Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body");
        var r = s.ReplaceText(para.Anchor.Id, "Main text EDITED continued.");
        Assert.True(r.Success, r.Error?.Message);

        var saved = s.Save();
        // The note ref must NOT be swept into the w:del: it survives on both sides.
        Assert.Equal(1, BodyNoteRefCount<FootnoteReference>(saved));

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", saved));
        var rejected = RevisionProcessor.RejectRevisions(new WmlDocument("r.docx", saved));
        Assert.Equal("Main text EDITED continued.", string.Join("\n",
            BodyParagraphTexts(accepted.DocumentByteArray)));
        Assert.Equal("Main text continued.", string.Join("\n",
            BodyParagraphTexts(rejected.DocumentByteArray)));
        Assert.Equal(1, BodyNoteRefCount<FootnoteReference>(accepted.DocumentByteArray));
        Assert.Equal(1, BodyNoteRefCount<FootnoteReference>(rejected.DocumentByteArray));
    }

    [Fact]
    public void DS316_ReplaceText_PreservesEndnoteReference_AcceptMode()
    {
        var baseBytes = BuildDS141_EndnoteFixture(); // "Body text"[enRef]
        Assert.Equal(1, BodyNoteRefCount<EndnoteReference>(baseBytes));

        using var s = new DocxSession(baseBytes);
        var para = s.Project().AnchorIndex.Values.Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body");
        var r = s.ReplaceText(para.Anchor.Id, "Body text EDITED");
        Assert.True(r.Success, r.Error?.Message);

        var saved = s.Save();
        Assert.Equal("Body text EDITED", string.Join("\n", BodyParagraphTexts(saved)));
        Assert.Equal(1, BodyNoteRefCount<EndnoteReference>(saved));
    }

    private static System.Collections.Generic.List<string> BodyParagraphTexts(byte[] bytes)
    {
        using var s = new MemoryStream(bytes);
        using var d = WordprocessingDocument.Open(s, false);
        return d.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Select(p => p.InnerText).ToList();
    }
}
