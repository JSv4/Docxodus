// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #959: the list numbers <see cref="DocxSession.Project"/> reports must match what a
/// save-and-reopen of the same document reports after every edit, not the counters computed
/// before the edit.
/// </summary>
public class DocxSessionListNumberingRefreshTests
{
    [Fact]
    public void DS959a_DeleteFirstItem_RenumbersSurvivors()
    {
        using var session = ThreeItemList(out var items);
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "3. Item 2" }, ListItems(session));

        Assert.True(session.DeleteBlock(items[0]).Success);

        Assert.Equal(new[] { "1. Item 1", "2. Item 2" }, ListItems(session));
        AssertMatchesReopen(session);
    }

    [Fact]
    public void DS959b_InsertItemBeforeFirst_RenumbersFollowers()
    {
        using var session = ThreeItemList(out var items);
        _ = ListItems(session);

        var inserted = session.InsertParagraph(items[0], Position.Before, "1. Item new");
        Assert.True(inserted.Success, inserted.Error?.Message);

        AssertMatchesReopen(session);
    }

    [Fact]
    public void DS959c_SetListLevel_RenumbersFollowingItems()
    {
        using var session = ThreeItemList(out var items);
        _ = ListItems(session);

        var result = session.SetListLevel(items[1], +1);
        Assert.True(result.Success, result.Error?.Message);

        AssertMatchesReopen(session);
        Assert.Equal("2. Item 2", ListItems(session)[2]);
    }

    [Fact]
    public void DS959d_ApplyListFormatToParagraphAbove_RenumbersTheList()
    {
        using var session = ThreeItemList(out var items);
        var before = BodyAnchors(session).First(a => a.Kind == "p").Id;
        _ = ListItems(session);

        var result = session.ApplyListFormat(before, ListFormat.Decimal);
        Assert.True(result.Success, result.Error?.Message);

        AssertMatchesReopen(session);
    }

    [Fact]
    public void DS959e_UndoAndRedoOfDelete_ReportCurrentNumbers()
    {
        using var session = ThreeItemList(out var items);
        _ = ListItems(session);
        Assert.True(session.DeleteBlock(items[0]).Success);
        _ = ListItems(session);

        Assert.True(session.Undo());
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "3. Item 2" }, ListItems(session));
        AssertMatchesReopen(session);

        Assert.True(session.Redo());
        Assert.Equal(new[] { "1. Item 1", "2. Item 2" }, ListItems(session));
        AssertMatchesReopen(session);
    }

    /// <summary>
    /// The session clears numbering only on the parts the retriever numbers. That is safe only
    /// while a numbered paragraph anywhere else (here, a comment that reuses the body list's
    /// numId) is never stamped with counters. If the retriever starts numbering comments,
    /// this fails and the clear set must grow with it.
    /// </summary>
    [Fact]
    public void DS959f_NumberedCommentParagraph_IsNeverStampedWithCounters()
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var session = ThreeItemList(out _);
        using var stream = new MemoryStream();
        stream.Write(session.Save());
        using var doc = WordprocessingDocument.Open(stream, true);
        var main = doc.MainDocumentPart!;
        var bodyItem = main.GetXDocument().Descendants(w + "p").First(p => p.Descendants(w + "numPr").Any());
        var commentItem = new XElement(w + "p",
            new XElement(w + "pPr", new XElement(bodyItem.Element(w + "pPr")!.Element(w + "numPr")!)),
            new XElement(w + "r", new XElement(w + "t", "Comment item")));
        var comments = main.AddNewPart<WordprocessingCommentsPart>();
        comments.PutXDocument(new XDocument(new XElement(w + "comments",
            new XElement(w + "comment", new XAttribute(w + "id", "0"), commentItem))));
        commentItem = comments.GetXDocument().Descendants(w + "p").Single();

        Assert.Equal("1.", ListItemRetriever.RetrieveListItem(doc, bodyItem)?.TrimEnd());
        Assert.Null(ListItemRetriever.RetrieveListItem(doc, commentItem));

        Assert.NotNull(bodyItem.Annotation<ListItemRetriever.ListItemInfo>());
        Assert.Null(commentItem.Annotation<ListItemRetriever.ListItemInfo>());
    }

    [Fact]
    public void DS959g_MoveLastItemFirst_RenumbersEveryItem()
    {
        using var session = ThreeItemList(out var items);
        _ = ListItems(session);

        var moved = session.MoveBlock(items[2], items[0], Position.Before);
        Assert.True(moved.Success, moved.Error?.Message);

        Assert.Equal(new[] { "1. Item 2", "2. Item 0", "3. Item 1" }, ListItems(session));
        AssertMatchesReopen(session);
    }

    [Fact]
    public void DS959h_TrackedDeleteOfFirstItem_MatchesReopen()
    {
        using var session = ThreeItemList(out var items);
        _ = ListItems(session);
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);

        Assert.True(session.DeleteBlock(items[0]).Success);

        AssertMatchesReopen(session);
    }

    [Fact]
    public void DS959i_StyleNumberedList_DeleteFirstItem_RenumbersSurvivors()
    {
        using var session = new DocxSession(StyleNumberedList(keepStyle: true));
        var before = BodyTexts(session);
        Assert.Contains("1. Item 0", before);
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.TextPreview.Contains("Item 0")).Anchor.Id;

        Assert.True(session.DeleteBlock(first).Success);

        Assert.Contains("1. Item 1", BodyTexts(session));
        using var reopened = new DocxSession(session.Save());
        Assert.Equal(BodyTexts(reopened), BodyTexts(session));
    }

    /// <summary>
    /// Each numbering input the stale check reads, changed in place on a live document that
    /// has already been counted: the recount after <c>ClearStaleAnnotations</c> must match a
    /// fresh count of the same change.
    /// </summary>
    [Theory]
    [MemberData(nameof(InPlaceNumberingChanges))]
    public void DS959j_InPlaceNumberingChange_IsRecounted(string change)
    {
        var (bytes, mutate) = Changes()[change];

        using var liveStream = Editable(bytes);
        using var live = WordprocessingDocument.Open(liveStream, true);
        ListItemRetriever.ClearStaleAnnotations(live);
        var before = Count(live);
        mutate(live);
        ListItemRetriever.ClearStaleAnnotations(live);
        var after = Count(live);

        using var freshStream = Editable(bytes);
        using var fresh = WordprocessingDocument.Open(freshStream, true);
        mutate(fresh);
        var expected = Count(fresh);

        Assert.NotEqual(before, expected);
        Assert.Equal(expected, after);
    }

    public static IEnumerable<object[]> InPlaceNumberingChanges() => Changes().Keys.Select(k => new object[] { k });

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static Dictionary<string, (byte[] Bytes, Action<WordprocessingDocument> Mutate)> Changes()
    {
        var direct = DirectList();
        var styled = StyleNumberedList(keepStyle: true);
        var plain = StyleNumberedList(keepStyle: false);
        return new()
        {
            ["start override in the numbering part"] = (direct, doc =>
            {
                var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.GetXDocument().Root!;
                var numId = (string)FirstItem(doc).Descendants(W + "numId").First().Attribute(W + "val")!;
                var num = numbering.Elements(W + "num").First(n => (string?)n.Attribute(W + "numId") == numId);
                num.Elements(W + "lvlOverride").Where(o => (string?)o.Attribute(W + "ilvl") == "0").Remove();
                num.Add(new XElement(W + "lvlOverride", new XAttribute(W + "ilvl", "0"),
                    new XElement(W + "startOverride", new XAttribute(W + "val", "5"))));
            }),
            ["numbering of the paragraph style"] = (styled, doc =>
                NumberedParaStyle(doc).Descendants(W + "numId").First().SetAttributeValue(W + "val", "0")),
            ["default paragraph style"] = (plain, doc =>
            {
                foreach (var style in doc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument().Root!.Elements(W + "style"))
                    style.Attribute(W + "default")?.Remove();
                NumberedParaStyle(doc).SetAttributeValue(W + "default", "1");
            }),
            ["emptied section-break paragraph"] = (direct, doc =>
            {
                var item = FirstItem(doc);
                item.Elements(W + "r").Remove();
                item.Element(W + "pPr")!.Add(new XElement(W + "sectPr"));
            }),
            ["deleted paragraph mark"] = (direct, doc =>
                FirstItem(doc).Element(W + "pPr")!.Add(new XElement(W + "rPr", new XElement(W + "del",
                    new XAttribute(W + "id", "901"), new XAttribute(W + "author", "a"))))),
            ["declared level numbers"] = (direct, doc =>
                FirstItem(doc).SetAttributeValue(PtOpenXml.LevelNumbers, "7")),
            ["paragraph switched to another numbered style"] = (styled, doc =>
                Item(doc, "Item 1").Descendants(W + "pStyle").First().SetAttributeValue(W + "val", "NumberedParaLevel2")),
            ["item repointed at another list instance"] = (direct, doc =>
                Item(doc, "Item 1").Descendants(W + "numId").First().SetAttributeValue(W + "val", "99")),
            ["last item's numbering style removed"] = (styled, doc =>
                Item(doc, "Item 2").Descendants(W + "pStyle").Remove()),
            ["declared continuation"] = (Nested(direct), doc =>
                Item(doc, "Item 1").SetAttributeValue(PtOpenXml.ListContinuation, "true")),
        };
    }

    private static XElement FirstItem(WordprocessingDocument doc) => Item(doc, "Item 0");

    private static XElement Item(WordprocessingDocument doc, string text) =>
        doc.MainDocumentPart!.GetXDocument().Descendants(W + "p")
            .First(p => p.Descendants(W + "t").Any(t => t.Value == text));

    private static XElement NumberedParaStyle(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument().Root!.Elements(W + "style")
            .First(s => (string?)s.Attribute(W + "styleId") == "NumberedPara");

    /// <summary><paramref name="bytes"/> with "Item 1" moved to the list's second level.</summary>
    private static byte[] Nested(byte[] bytes)
    {
        using var stream = Editable(bytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var numPr = Item(doc, "Item 1").Descendants(W + "numPr").First();
            numPr.Elements(W + "ilvl").Remove();
            numPr.AddFirst(new XElement(W + "ilvl", new XAttribute(W + "val", "1")));
            doc.MainDocumentPart!.PutXDocument();
        }
        return stream.ToArray();
    }

    /// <summary>Each body paragraph's list text, its text, and whether the retriever counts it as
    /// continuing its parent level's sequence (which the HTML renderer reads).</summary>
    private static string[] Count(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.GetXDocument().Descendants(W + "p")
            .Select(p => (ListItemRetriever.RetrieveListItem(doc, p) ?? string.Empty) + "|" +
                string.Concat(p.Descendants(W + "t").Select(t => t.Value)) +
                (p.Annotation<ListItemRetriever.ContinuationInfo>()?.IsContinuation == true ? "|continues" : string.Empty))
            .ToArray();

    private static MemoryStream Editable(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// The three-item list, plus a second list instance (numId 99) of the same definition that
    /// starts at 5 and no paragraph uses yet.
    /// </summary>
    private static byte[] DirectList()
    {
        using var session = ThreeItemList(out _);
        using var stream = Editable(session.Save());
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!;
            var root = numbering.GetXDocument().Root!;
            var numId = (string)Item(doc, "Item 0").Descendants(W + "numId").First().Attribute(W + "val")!;
            var num = root.Elements(W + "num").First(n => (string?)n.Attribute(W + "numId") == numId);
            root.Add(new XElement(W + "num", new XAttribute(W + "numId", "99"),
                new XElement(num.Element(W + "abstractNumId")!),
                new XElement(W + "lvlOverride", new XAttribute(W + "ilvl", "0"),
                    new XElement(W + "startOverride", new XAttribute(W + "val", "5")))));
            numbering.PutXDocument();
        }
        return stream.ToArray();
    }

    /// <summary>
    /// The three-item list, numbered through a "NumberedPara" paragraph style instead of direct
    /// numbering, plus an unused "NumberedParaLevel2" style that numbers the same list at its second
    /// level. With <paramref name="keepStyle"/> false the items carry no style at all.
    /// </summary>
    private static byte[] StyleNumberedList(bool keepStyle)
    {
        using var stream = Editable(DirectList());
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument();
            var items = body.Descendants(W + "p").Where(p => p.Descendants(W + "numPr").Any()).ToList();
            var numPr = items[0].Descendants(W + "numPr").First();
            var styles = main.StyleDefinitionsPart!.GetXDocument();
            styles.Root!.Add(new XElement(W + "style",
                new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", "NumberedPara"),
                new XElement(W + "name", new XAttribute(W + "val", "Numbered Para")),
                new XElement(W + "pPr", new XElement(W + "numPr", numPr.Element(W + "numId")))));
            styles.Root!.Add(new XElement(W + "style",
                new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", "NumberedParaLevel2"),
                new XElement(W + "name", new XAttribute(W + "val", "Numbered Para Level 2")),
                new XElement(W + "pPr", new XElement(W + "numPr",
                    new XElement(W + "ilvl", new XAttribute(W + "val", "1")), numPr.Element(W + "numId")))));
            foreach (var item in items)
            {
                var pPr = item.Element(W + "pPr")!;
                pPr.Element(W + "numPr")!.Remove();
                pPr.Element(W + "pStyle")?.Remove();
                if (keepStyle) pPr.AddFirst(new XElement(W + "pStyle", new XAttribute(W + "val", "NumberedPara")));
            }
            main.PutXDocument();
            main.StyleDefinitionsPart.PutXDocument();
        }
        return stream.ToArray();
    }

    private static string[] BodyTexts(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body")
            .Select(t => t.FullText)
            .ToArray();

    private static DocxSession ThreeItemList(out string[] items)
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = BodyAnchors(session).First().Id;
        var created = session.InsertParagraph(first, Position.After, "1. Item 0\n2. Item 1\n3. Item 2");
        Assert.True(created.Success, created.Error?.Message);
        items = created.Created.Select(a => a.Id).ToArray();
        Assert.Equal(3, items.Length);
        return session;
    }

    private static Anchor[] BodyAnchors(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body")
            .Select(t => t.Anchor)
            .ToArray();

    private static string[] ListItems(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "li")
            .Select(t => t.FullText)
            .ToArray();

    private static void AssertMatchesReopen(DocxSession session)
    {
        var live = ListItems(session);
        using var reopened = new DocxSession(session.Save());
        Assert.Equal(ListItems(reopened), live);
    }
}
