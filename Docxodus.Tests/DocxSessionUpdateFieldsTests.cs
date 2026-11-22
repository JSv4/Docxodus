#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary><see cref="DocxSession.UpdateFields"/>: a table of contents built from the headings,
/// REF / PAGEREF / SEQ results made current, one undo, a range, the update-on-open request.</summary>
public class DocxSessionUpdateFieldsTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind is "p" or "h" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();

    private static XElement Body(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        return XDocument.Parse(new StreamReader(doc.MainDocumentPart!.GetStream()).ReadToEnd()).Root!.Element(W + "body")!;
    }

    private static bool AsksForUpdate(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        var part = doc.MainDocumentPart!.DocumentSettingsPart;
        if (part is null) return false;
        return XDocument.Parse(new StreamReader(part.GetStream()).ReadToEnd()).Descendants(W + "updateFields").Any();
    }

    private static string Text(XElement paragraph) =>
        string.Concat(paragraph.Descendants().Where(e => e.Name == W + "t" || (e.Name == W + "tab" && e.Parent?.Name == W + "r"))
            .Select(e => e.Name == W + "tab" ? "\t" : e.Value));

    /// <summary>Two headings with a table of contents above them.</summary>
    private static (DocxSession Session, string[] Headings) WithContents()
    {
        // (outline levels on the paragraphs themselves: the bare test document has no heading styles)
        var stream = new MemoryStream();
        stream.Write(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var ps = xml.Descendants(W + "p").ToList();
            for (int i = 0; i < 2; i++)
            {
                var pPr = ps[i].Element(W + "pPr");
                if (pPr is null) { pPr = new XElement(W + "pPr"); ps[i].AddFirst(pPr); }
                pPr.Add(new XElement(W + "outlineLvl", new XAttribute(W + "val", i.ToString())));
            }
            part.PutXDocument();
        }
        var session = new DocxSession(stream.ToArray());
        var headings = BodyParagraphs(session);
        Assert.True(session.InsertTableOfContents(headings[0], Position.Before, new TableOfContentsOptions()).Success);
        return (session, session.ListOutline().Select(h => h.AnchorId).ToArray());
    }

    [Fact]
    public void UF001_TableOfContents_IsBuiltFromTheHeadings_WithPagesLinksAndBookmarks()
    {
        var (session, headings) = WithContents();
        using var _ = session;
        Assert.True(AsksForUpdate(session));
        var texts = session.ListOutline().Select(h => h.Text).ToArray();

        var result = session.UpdateFields(new UpdateFieldsOptions
        {
            Pages = new Dictionary<string, string> { [headings[0]] = "2", [headings[1]] = "5" },
        });

        Assert.True(result.Success, result.Error?.Message);
        var body = Body(session);
        var entries = body.Descendants(W + "p")
            .Where(p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") is "TOC1" or "TOC2").ToList();
        Assert.Equal(2, entries.Count);
        Assert.EndsWith(texts[0] + "\t2", Text(entries[0]));
        Assert.EndsWith(texts[1] + "\t5", Text(entries[1]));
        Assert.Equal("TOC2", (string?)entries[1].Element(W + "pPr")!.Element(W + "pStyle")!.Attribute(W + "val"));
        // each entry links to a _Toc bookmark on its heading, and its page is a PAGEREF to it
        var link = entries[0].Descendants(W + "hyperlink").Single();
        var bookmark = (string)link.Attribute(W + "anchor")!;
        Assert.Matches(new Regex(@"^_Toc\d{9}$"), bookmark);
        Assert.Contains(session.ListBookmarks(), b => b.Name == bookmark && b.Text == texts[0]);
        Assert.Contains(entries[0].Descendants(W + "instrText"), t => t.Value.Contains("PAGEREF " + bookmark));
        // one TOC field still spans the entries, and the document no longer asks Word to update it
        Assert.Single(body.Descendants(W + "instrText"), t => t.Value.TrimStart().StartsWith("TOC"));
        Assert.False(AsksForUpdate(session));
        Assert.True(session.ListOutline().Count == 2);

        // again: the same bookmarks, the new pages
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            Pages = new Dictionary<string, string> { [headings[0]] = "3", [headings[1]] = "4" },
        }).Success);
        Assert.Equal(2, session.ListBookmarks().Count(b => b.Name.StartsWith("_Toc", StringComparison.Ordinal)));
        Assert.Contains(Body(session).Descendants(W + "p"), p => Text(p).EndsWith(texts[1] + "\t4", StringComparison.Ordinal));

        // one undo takes the last update back
        Assert.True(session.Undo());
        Assert.Contains(Body(session).Descendants(W + "p"), p => Text(p).EndsWith(texts[1] + "\t5", StringComparison.Ordinal));
    }

    [Fact]
    public void UF002_TableOfContents_WithoutHeadings_SaysSo()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(session.InsertTableOfContents(BodyParagraphs(session)[0], Position.Before, new TableOfContentsOptions()).Success);
        Assert.True(session.UpdateFields().Success);
        Assert.Contains(Body(session).Descendants(W + "p"), p => Text(p) == "No table of contents entries found.");
    }

    [Fact]
    public void UF003_Ref_TakesTheBookmarksTextAsItIsNow()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);
        Assert.True(session.InsertCrossReference(paragraphs[1], 0, "target").Success);
        var before = Body(session).Descendants(W + "fldSimple").Single().Value;
        Assert.True(session.ReplaceTextAtSpan(paragraphs[0], 0, 1, "Z").Success);

        Assert.True(session.UpdateFields().Success);

        var after = Body(session).Descendants(W + "fldSimple").Single().Value;
        Assert.Equal("Z" + before[1..], after);
    }

    [Fact]
    public void UF004_PageRef_TakesThePageFromTheMap_AndAMissingBookmarkSaysSo()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);
        Assert.True(session.InsertCrossReference(paragraphs[1], 0, "target", new CrossReferenceOptions { PageNumber = true }).Success);
        var unid = paragraphs[0][(paragraphs[0].LastIndexOf(':') + 1)..];
        Assert.True(session.UpdateFields(new UpdateFieldsOptions { Pages = new Dictionary<string, string> { [unid] = "9" } }).Success);
        var field = Body(session).Descendants(W + "fldSimple").Single();
        Assert.Equal("9", field.Value);
        Assert.Null(field.Attribute(W + "dirty"));
    }

    [Fact]
    public void UF005_Seq_IsNumberedAgainInDocumentOrder_AndARangeUpdatesOnlyItsFields()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var ps = xml.Descendants(W + "p").ToList();
            XElement Seq(string instr, string shown) => new XElement(W + "fldSimple", new XAttribute(W + "instr", instr),
                new XElement(W + "r", new XElement(W + "t", shown)));
            ps[0].Add(Seq(" SEQ Figure \\* ARABIC ", "7"), Seq(" SEQ Figure ", "7"));
            ps[1].Add(Seq(" SEQ Figure \\* ROMAN ", "7"), Seq(" SEQ Table ", "7"));
            part.PutXDocument();
        }
        using var session = new DocxSession(stream.ToArray());
        var paragraphs = BodyParagraphs(session);

        // only the second paragraph: its fields numbered as the whole document counts them
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            Range = DocumentRange.In(paragraphs[1], new CharSpan(0, 0)),
        }).Success);
        Assert.Equal(new[] { "7", "7", "III", "1" }, Body(session).Descendants(W + "fldSimple").Select(f => f.Value).ToArray());

        Assert.True(session.UpdateFields().Success);
        Assert.Equal(new[] { "1", "2", "III", "1" }, Body(session).Descendants(W + "fldSimple").Select(f => f.Value).ToArray());
    }

    [Fact]
    public void UF007_TablesOfContentsOnly_LeavesOtherFields()
    {
        var (session, _) = WithContents();
        using var _s = session;
        var paragraphs = BodyParagraphs(session);
        var last = paragraphs[^1];
        Assert.True(session.AddBookmark("target", DocumentRange.In(last, new CharSpan(0, 2))).Success);
        Assert.True(session.InsertCrossReference(last, 0, "target", new CrossReferenceOptions { PageNumber = true, CachedPage = "8" }).Success);
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            TablesOfContentsOnly = true,
            Pages = new Dictionary<string, string> { [last] = "3" },
        }).Success);
        var body = Body(session);
        Assert.Equal("8", body.Descendants(W + "fldSimple").Single().Value);
        Assert.Contains(body.Descendants(W + "p"), p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "TOC1");
        // (part of an update: the document still asks Word for the rest)
        Assert.True(AsksForUpdate(session));
    }

    [Fact]
    public void UF008_ARangeInAContentsEntry_RebuildsTheTable()
    {
        var (session, headings) = WithContents();
        using var _s = session;
        Assert.True(session.UpdateFields().Success);
        var entry = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p" && session.Raw.GetXml(t.Anchor.Id).Contains("TOC1")).Anchor.Id;
        Assert.True(session.ReplaceTextAtSpan(headings[0], 0, 1, "Z").Success);
        var result = session.UpdateFields(new UpdateFieldsOptions { Range = DocumentRange.In(entry, new CharSpan(1, 0)) });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Contains(Body(session).Descendants(W + "p"),
            p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "TOC1" && Text(p).StartsWith("Z", StringComparison.Ordinal));
    }

    [Fact]
    public void UF006_RefusedWhileChangesAreTracked()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        Assert.False(session.UpdateFields().Success);
    }
}
