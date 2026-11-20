// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

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

    [Fact]
    public void LM005_LevelTheListDoesNotDefine_IsDefinedFirst()
    {
        // A list whose definition has level 0 only: joining it at level 3 must define level 3,
        // or the paragraph would point at a level with no number.
        using var session = new DocxSession(LevelZeroOnlyList());
        var plain = BodyAnchor(session, "Plain");
        Assert.DoesNotContain(3, DefinedLevels(session));

        var res = session.SetListMembership(plain, 1, 3);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Contains(3, DefinedLevels(session));
        var membership = session.GetListMembership(res.Modified[0].Id)!;
        Assert.Equal((1, 3), (membership.NumId, membership.Level));
        Assert.Equal(2, ListItems(session).Length);
        using var reopened = new DocxSession(session.Save());
        Assert.Contains(3, DefinedLevels(reopened));
        Assert.Equal(ListItems(session), ListItems(reopened));
        Assert.True(session.Undo());
        Assert.DoesNotContain(3, DefinedLevels(session));
    }

    [Fact]
    public void LM006_TrackedChanges_RecordTheNewNumbering_AndRejectRestoresIt()
    {
        // A plain paragraph joining a list: Word records the added numbering as w:numPr/w:ins.
        using var session = ListThenPlain(out var items, out var plain);
        int numId = session.GetListMembership(items[0])!.NumId;
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);

        var res = session.SetListMembership(plain, numId, 0);

        Assert.True(res.Success, res.Error?.Message);
        var numPr = Paragraph(session.Save(), "Plain").Element(W.pPr)!.Element(W.numPr)!;
        Assert.Equal(numId, (int?)numPr.Element(W.numId)!.Attribute(W.val));
        Assert.NotNull(numPr.Element(W.ins));
        Assert.NotEmpty(session.ListRevisions());
        Assert.True(session.RejectAllRevisions().Success);
        Assert.Null(Paragraph(session.Save(), "Plain").Element(W.pPr)?.Element(W.numPr));
    }

    [Fact]
    public void LM006b_TrackedChanges_ListItemMovingLists_KeepsTheOldListInPPrChange()
    {
        // A list item moving to another list: a w:pPrChange holds its old numbering.
        using var session = ListThenPlain(out var items, out _);
        int first = session.GetListMembership(items[0])!.NumId;
        Assert.True(session.SetListStartOverride(items[2], 1).Success);
        var moved = ListItemAnchors(session).Last();
        int restarted = session.GetListMembership(moved)!.NumId;
        Assert.NotEqual(first, restarted);
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);

        var res = session.SetListMembership(moved, first, 0);

        Assert.True(res.Success, res.Error?.Message);
        var pPr = Paragraph(session.Save(), "Item 2").Element(W.pPr)!;
        Assert.Equal(first, (int?)pPr.Element(W.numPr)!.Element(W.numId)!.Attribute(W.val));
        var old = pPr.Element(W.pPrChange)?.Element(W.pPr)?.Element(W.numPr);
        Assert.Equal(restarted, (int?)old?.Element(W.numId)?.Attribute(W.val));
        Assert.True(session.RejectAllRevisions().Success);
        Assert.Equal(restarted, (int?)Paragraph(session.Save(), "Item 2")
            .Element(W.pPr)!.Element(W.numPr)!.Element(W.numId)!.Attribute(W.val));
    }

    [Fact]
    public void LM007_TrackedChanges_RefuseALevelThatWouldHaveToBeDefined()
    {
        // Defining a level changes the numbering part, which has no tracked-change encoding.
        using var session = new DocxSession(LevelZeroOnlyList());
        var plain = BodyAnchor(session, "Plain");
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);

        var res = session.SetListMembership(plain, 1, 3);

        Assert.False(res.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, res.Error!.Code);
        Assert.DoesNotContain(3, DefinedLevels(session));
        Assert.Null(session.GetListMembership(plain));
    }

    private static XElement Paragraph(byte[] docx, string text)
    {
        using var zip = new ZipArchive(new MemoryStream(docx));
        using var stream = zip.GetEntry("word/document.xml")!.Open();
        return XDocument.Load(stream).Descendants(W.p).Single(p => string.Concat(p.Descendants(W.t).Select(t => t.Value)) == text);
    }

    private static string BodyAnchor(DocxSession session, string text) =>
        session.Project().AnchorIndex.Values.First(t => t.Anchor.Scope == "body" && t.FullText == text).Anchor.Id;

    private static int[] DefinedLevels(DocxSession session)
    {
        using var zip = new ZipArchive(new MemoryStream(session.Save()));
        using var stream = zip.GetEntry("word/numbering.xml")!.Open();
        return XDocument.Load(stream).Descendants(W.abstractNum).Single().Elements(W.lvl)
            .Select(l => (int)l.Attribute(W.ilvl)!).ToArray();
    }

    /// <summary>One numbered item (list 1, whose definition has level 0 only), then "Plain".</summary>
    private static byte[] LevelZeroOnlyList()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(
                    new Wp.ParagraphProperties(new Wp.NumberingProperties(
                        new Wp.NumberingLevelReference { Val = 0 }, new Wp.NumberingId { Val = 1 })),
                    new Wp.Run(new Wp.Text("Item"))),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Plain"))),
                new Wp.SectionProperties()));
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Wp.Numbering(
                new Wp.AbstractNum(
                    new Wp.Level(
                        new Wp.StartNumberingValue { Val = 1 },
                        new Wp.NumberingFormat { Val = Wp.NumberFormatValues.Decimal },
                        new Wp.LevelText { Val = "%1." }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new Wp.NumberingInstance(new Wp.AbstractNumId { Val = 0 }) { NumberID = 1 });
        }
        return ms.ToArray();
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
