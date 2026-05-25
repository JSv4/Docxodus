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
    public void DS054_SetListLevelIndent()
    {
        using var s = new DocxSession(BuildDS002_BulletedList());
        var firstLi = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("li:"));
        var r = s.SetListLevel(firstLi, +1);
        Assert.True(r.Success, r.Error?.Message);
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
        // page-rendering scaffolding — deleting them corrupts the document. The op
        // must refuse with AnchorWrongKind.
        using var s = new DocxSession(BuildDS140_FootnoteFixture());
        var separator = s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "fn"
                     && AnchorXmlContains(s, t.Anchor.Id, "w:type=\"separator\""));

        var r = s.DeleteBlock(separator.Anchor.Id);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, r.Error!.Code);
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
}
