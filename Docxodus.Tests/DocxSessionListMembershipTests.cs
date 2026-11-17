// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxSession.SetListMembership"/>: a paragraph joins an existing list at a level, as
/// Word's Format Painter makes a paragraph painted from a numbered one number on in its list.
/// </summary>
public class DocxSessionListMembershipTests
{
    [Fact]
    public void LM001_PlainParagraph_JoinsTheList_AndNumbersOn()
    {
        using var session = ListThenPlain(out var items, out var plain);
        int numId = session.GetListMembership(items[0])!.NumId;

        var res = session.SetListMembership(plain, numId, 0);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "3. Item 2", "4. Plain" }, ListItems(session));
        Assert.Equal(numId, session.GetListMembership(res.Modified[0].Id)!.NumId);
        using var reopened = new DocxSession(session.Save());
        Assert.Equal(ListItems(session), ListItems(reopened));
    }

    [Fact]
    public void LM002_Level_IsSet_AndUndoIsOneStep()
    {
        using var session = ListThenPlain(out var items, out var plain);
        int numId = session.GetListMembership(items[0])!.NumId;

        var res = session.SetListMembership(plain, numId, 1);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(1, session.GetListMembership(res.Modified[0].Id)!.Level);
        Assert.Equal(4, ListItems(session).Length);
        Assert.True(session.Undo());
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "3. Item 2" }, ListItems(session));
    }

    [Fact]
    public void LM003_UnknownList_AndBadLevel_AreRefused_WithoutAChange()
    {
        using var session = ListThenPlain(out var items, out var plain);
        int numId = session.GetListMembership(items[0])!.NumId;

        var unknown = session.SetListMembership(plain, numId + 1000, 0);
        var level = session.SetListMembership(plain, numId, 9);

        Assert.False(unknown.Success);
        Assert.Equal(EditErrorCode.InvalidListLevel, unknown.Error!.Code);
        Assert.False(level.Success);
        Assert.Equal(EditErrorCode.InvalidListLevel, level.Error!.Code);
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "3. Item 2" }, ListItems(session));
    }

    [Fact]
    public void LM004_ListItem_MovesToAnotherList()
    {
        using var session = ListThenPlain(out var items, out var plain);
        int first = session.GetListMembership(items[0])!.NumId;
        var restarted = session.SetListStartOverride(items[2], 1);
        Assert.True(restarted.Success, restarted.Error?.Message);
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "1. Item 2" }, ListItems(session));
        var moved = ListItemAnchors(session).Last();

        var res = session.SetListMembership(moved, first, 0);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "1. Item 0", "2. Item 1", "3. Item 2" }, ListItems(session));
    }

    private static DocxSession ListThenPlain(out string[] items, out string plain)
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = session.Project().AnchorIndex.Values.First(t => t.Anchor.Scope == "body").Anchor.Id;
        var created = session.InsertParagraph(first, Position.After, "1. Item 0\n2. Item 1\n3. Item 2");
        Assert.True(created.Success, created.Error?.Message);
        items = created.Created.Select(a => a.Id).ToArray();
        var after = session.InsertParagraph(items[^1], Position.After, "Plain");
        Assert.True(after.Success, after.Error?.Message);
        plain = after.Created.Single().Id;
        Assert.Null(session.GetListMembership(plain));
        return session;
    }

    private static string[] ListItemAnchors(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "li")
            .Select(t => t.Anchor.Id)
            .ToArray();

    private static string[] ListItems(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "li")
            .Select(t => t.FullText)
            .ToArray();
}
