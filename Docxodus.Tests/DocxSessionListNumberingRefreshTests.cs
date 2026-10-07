// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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
