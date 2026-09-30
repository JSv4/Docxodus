// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="BookmarkIds"/> renumbers bookmarks that share an id, keeping each start with its own end
/// (issue #840).
/// </summary>
public class BookmarkIdsTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement Body(string innerXml) =>
        XElement.Parse($"<w:body xmlns:w=\"{W}\" " +
                       "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">" +
                       $"{innerXml}</w:body>");

    private static string IdOf(XElement root, string markerName, string? bookmarkName = null, int index = 0) =>
        (string)root.Descendants(W + markerName)
            .Where(e => bookmarkName is null || (string?)e.Attribute(W + "name") == bookmarkName)
            .ElementAt(index).Attribute(W + "id")!;

    [Fact]
    public void OverlappingDeletedAndInsertedBookmarks_EachKeepsItsOwnEnd()
    {
        // The deleted bookmark opens first and closes last; the inserted one lies inside it. Pairing by
        // position alone would give the deleted start the inserted end.
        var body = Body(
            "<w:p><w:del><w:bookmarkStart w:id=\"0\" w:name=\"Old\"/></w:del>" +
            "<w:ins><w:bookmarkStart w:id=\"0\" w:name=\"New\"/><w:r><w:t>new</w:t></w:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></w:ins>" +
            "<w:del><w:r><w:delText>old</w:delText></w:r><w:bookmarkEnd w:id=\"0\"/></w:del></w:p>");

        BookmarkIds.MakeUnique(new[] { body });

        var ends = body.Descendants(W + "bookmarkEnd").ToList();
        Assert.Equal(IdOf(body, "bookmarkStart", "New"), (string)ends[0].Attribute(W + "id")!);
        Assert.Equal(IdOf(body, "bookmarkStart", "Old"), (string)ends[1].Attribute(W + "id")!);
        Assert.NotEqual(IdOf(body, "bookmarkStart", "Old"), IdOf(body, "bookmarkStart", "New"));
    }

    [Fact]
    public void KeptBookmarkSplitAcrossRevisions_DeletedEndNeverClosesInsertedStart()
    {
        // The original's bookmark runs from a deleted paragraph to a deleted one; the revised document's opens in
        // a paragraph both kept and closes in an inserted one. Neither end has a start on its own side.
        var body = Body(
            "<w:p><w:del><w:bookmarkStart w:id=\"0\" w:name=\"Old\"/></w:del></w:p>" +
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"New\"/><w:r><w:t>kept</w:t></w:r></w:p>" +
            "<w:p><w:ins><w:r><w:t>new</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:ins></w:p>" +
            "<w:p><w:del><w:r><w:delText>old</w:delText></w:r><w:bookmarkEnd w:id=\"0\"/></w:del></w:p>");

        BookmarkIds.MakeUnique(new[] { body });

        Assert.Equal(IdOf(body, "bookmarkStart", "New"), IdOf(body, "bookmarkEnd", index: 0));
        Assert.Equal(IdOf(body, "bookmarkStart", "Old"), IdOf(body, "bookmarkEnd", index: 1));
        Assert.NotEqual(IdOf(body, "bookmarkStart", "Old"), IdOf(body, "bookmarkStart", "New"));
    }

    [Fact]
    public void UnchangedBookmark_KeepsItsIdOverAnInsertedOne()
    {
        var body = Body(
            "<w:p><w:ins><w:bookmarkStart w:id=\"0\" w:name=\"Inserted\"/><w:r><w:t>x</w:t></w:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></w:ins></w:p>" +
            "<w:p><m:oMath><w:bookmarkStart w:id=\"0\" w:name=\"Kept\"/><m:r><m:t>y</m:t></m:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></m:oMath></w:p>");

        BookmarkIds.MakeUnique(new[] { body });

        Assert.Equal("0", IdOf(body, "bookmarkStart", "Kept"));
        Assert.Equal("1", IdOf(body, "bookmarkStart", "Inserted"));
        Assert.Equal("1", IdOf(body, "bookmarkEnd"));
    }

    [Fact]
    public void BookmarksInDifferentStories_ShareOneIdSpace()
    {
        var body = Body("<w:p><w:bookmarkStart w:id=\"3\" w:name=\"A\"/><w:bookmarkEnd w:id=\"3\"/></w:p>");
        var notes = Body("<w:p><w:ins><w:bookmarkStart w:id=\"3\" w:name=\"B\"/><w:bookmarkEnd w:id=\"3\"/></w:ins></w:p>");

        var changed = BookmarkIds.MakeUnique(new[] { body, notes });

        Assert.Equal(new[] { notes }, changed);
        Assert.Equal("3", IdOf(body, "bookmarkStart"));
        Assert.Equal("4", IdOf(notes, "bookmarkStart"));
        Assert.Equal("4", IdOf(notes, "bookmarkEnd"));
    }

    [Fact]
    public void UnchangedBookmarksInDifferentStories_KeepTheIdTheySharedInTheSource()
    {
        var body = Body("<w:p><w:bookmarkStart w:id=\"0\" w:name=\"A\"/><w:bookmarkEnd w:id=\"0\"/></w:p>");
        var notes = Body("<w:p><w:bookmarkStart w:id=\"0\" w:name=\"B\"/><w:bookmarkEnd w:id=\"0\"/></w:p>");

        Assert.Empty(BookmarkIds.MakeUnique(new[] { body, notes }));
        Assert.Equal("0", IdOf(notes, "bookmarkStart"));
    }

    [Fact]
    public void EndWithoutStart_IsDropped()
    {
        var body = Body(
            "<w:p><w:del><w:bookmarkStart w:id=\"0\" w:name=\"Old\"/><w:r><w:delText>old</w:delText></w:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></w:del><w:ins><w:bookmarkEnd w:id=\"1\"/></w:ins></w:p>");

        BookmarkIds.MakeUnique(new[] { body });

        Assert.Equal(new[] { "0" }, body.Descendants(W + "bookmarkEnd").Select(e => (string)e.Attribute(W + "id")!));
    }

    [Fact]
    public void UniqueIds_AreLeftAlone()
    {
        var body = Body(
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"A\"/><w:bookmarkEnd w:id=\"0\"/>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"B\"/><w:bookmarkEnd w:id=\"1\"/></w:p>");
        var before = body.ToString();

        Assert.Empty(BookmarkIds.MakeUnique(new[] { body }));
        Assert.Equal(before, body.ToString());
    }
}
