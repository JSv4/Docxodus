#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
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
}
