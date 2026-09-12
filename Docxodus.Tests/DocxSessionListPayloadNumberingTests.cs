// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #759: markdown list payloads create native Word numbering through the same owner
/// <see cref="DocxSession.ApplyListFormat"/> uses, instead of bare paragraphs that only
/// project as list items.
/// </summary>
public class DocxSessionListPayloadNumberingTests
{
    private static readonly XNamespace Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void DS759a_BulletPayload_CreatesNativeBulletList()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);

        var result = session.InsertParagraph(first, Position.After, "- one\n- two");

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { "li", "li" }, result.Created.Select(a => a.Kind).ToArray());
        var one = NumPr(session, result.Created[0]);
        var two = NumPr(session, result.Created[1]);
        Assert.Equal(one, two);
        Assert.Equal(0, one.Ilvl);
        var saved = session.Save();
        Assert.Equal(NumberFormat.Bullet, NumFmt(saved, one.NumId, 0));
        Assert.Empty(ValidationErrors(saved));

        using var reopened = new DocxSession(saved);
        Assert.Equal(new[] { "p", "li", "li", "p" },
            reopened.Project().AnchorIndex.Values.Where(a => a.Anchor.Scope == "body")
                .Select(a => a.Anchor.Kind).ToArray());
    }

    [Fact]
    public void DS759b_OrderedPayload_StartsWhereItsFirstMarkerSays()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);

        var result = session.InsertParagraph(first, Position.After, "3. three\n4. four");

        Assert.True(result.Success, result.Error?.Message);
        var three = NumPr(session, result.Created[0]);
        Assert.Equal(three, NumPr(session, result.Created[1]));
        Assert.Equal(NumberFormat.Decimal, NumFmt(session.Save(), three.NumId, 0));
        Assert.Equal(3, StartOverride(session.Save(), three.NumId, 0));
        Assert.Equal("3. three", session.GetAnchorInfo(result.Created[0].Id)!.FullText);
        Assert.Equal("4. four", session.GetAnchorInfo(result.Created[1].Id)!.FullText);
    }

    [Fact]
    public void DS759c_NestedPayload_MapsIndentToLevelsAndKeepsOneOrderedSequence()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);

        var result = session.InsertParagraph(first, Position.After,
            "1. a\n  - b\n    1. c\n2. d");

        Assert.True(result.Success, result.Error?.Message);
        var (a, b, c, d) = (NumPr(session, result.Created[0]), NumPr(session, result.Created[1]),
            NumPr(session, result.Created[2]), NumPr(session, result.Created[3]));
        Assert.Equal(new[] { 0, 1, 2, 0 }, new[] { a.Ilvl, b.Ilvl, c.Ilvl, d.Ilvl });
        Assert.Equal(a.NumId, c.NumId);
        Assert.Equal(a.NumId, d.NumId);
        Assert.NotEqual(a.NumId, b.NumId);
        var saved = session.Save();
        Assert.Equal(NumberFormat.Bullet, NumFmt(saved, b.NumId, 1));
        Assert.Equal(NumberFormat.Decimal, NumFmt(saved, c.NumId, 2));
        Assert.Equal("1. a", session.GetAnchorInfo(result.Created[0].Id)!.FullText);
        Assert.Equal("2. d", session.GetAnchorInfo(result.Created[3].Id)!.FullText);
        Assert.Empty(ValidationErrors(saved));
    }

    [Fact]
    public void DS759d_SeparateListsInOnePayload_RestartIndependently()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);

        var result = session.InsertParagraph(first, Position.After,
            "1. a\n2. b\n\nplain\n\n1. c");

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { "li", "li", "p", "li" }, result.Created.Select(x => x.Kind).ToArray());
        Assert.NotEqual(NumPr(session, result.Created[0]).NumId, NumPr(session, result.Created[3]).NumId);
        Assert.Equal("2. b", session.GetAnchorInfo(result.Created[1].Id)!.FullText);
        Assert.Equal("1. c", session.GetAnchorInfo(result.Created[3].Id)!.FullText);
    }

    [Theory]
    [InlineData(Position.After)]
    [InlineData(Position.Before)]
    public void DS759e_FirstListContinuesACompatibleNeighborOnly(Position position)
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);
        Assert.True(session.ApplyListFormat(first, ListFormat.Bullet).Success);
        var existing = NumPr(session, session.Project().AnchorIndex.Values
            .Single(a => a.Anchor.Kind == "li").Anchor);

        var bullets = session.InsertParagraph(FirstBodyParagraph(session), position, "- joined\n- also");
        Assert.True(bullets.Success, bullets.Error?.Message);
        Assert.All(bullets.Created, created => Assert.Equal(existing.NumId, NumPr(session, created).NumId));

        // A numbered payload next to a bullet item is a different family and starts its own list.
        var numbered = session.InsertParagraph(FirstBodyParagraph(session), position, "1. separate");
        Assert.True(numbered.Success, numbered.Error?.Message);
        var separate = NumPr(session, numbered.Created[0]);
        Assert.NotEqual(existing.NumId, separate.NumId);
        Assert.Equal(NumberFormat.Decimal, NumFmt(session.Save(), separate.NumId, 0));
    }

    [Fact]
    public void DS759f_ReplaceText_ListMarkerPromotesPlainParagraphsButNotExistingItems()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);

        var promoted = session.ReplaceText(first, "- now a bullet");
        Assert.True(promoted.Success, promoted.Error?.Message);
        var item = Assert.Single(promoted.Modified);
        Assert.Equal("li", item.Kind);
        Assert.Equal("now a bullet", session.GetAnchorInfo(item.Id)!.TextPreview);
        var bullet = NumPr(session, item);
        Assert.Equal(NumberFormat.Bullet, NumFmt(session.Save(), bullet.NumId, 0));

        // An existing item keeps its numbering; the marker is the projection's spelling.
        var renamed = session.ReplaceText(item.Id, "1. renamed");
        Assert.True(renamed.Success, renamed.Error?.Message);
        Assert.Equal(bullet, NumPr(session, renamed.Modified[0]));
        Assert.Equal("renamed", session.GetAnchorInfo(renamed.Modified[0].Id)!.TextPreview);

        // A backslash-escaped marker is literal text, exactly as before.
        var second = session.Project().AnchorIndex.Values
            .Single(a => a.Anchor.Scope == "body" && a.Anchor.Kind == "p").Anchor.Id;
        var literal = session.ReplaceText(second, "\\- not a list");
        Assert.True(literal.Success, literal.Error?.Message);
        Assert.Equal("p", literal.Modified[0].Kind);
        Assert.Equal("- not a list", session.GetAnchorInfo(literal.Modified[0].Id)!.TextPreview);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS759g_TrackedInsert_RejectRemovesTheNumberingItBroughtIn(bool accept)
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        Assert.False(HasNumberingPart(baseline));
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "List Author",
        });
        var first = FirstBodyParagraph(session);

        var result = session.InsertParagraph(first, Position.After, "1. a\n2. b");

        Assert.True(result.Success, result.Error?.Message);
        Assert.True(HasNumberingPart(session.Save()));
        Assert.NotEmpty(session.ListRevisions());
        var resolved = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.Empty(session.ListRevisions());

        var saved = session.Save();
        if (accept)
        {
            Assert.True(HasNumberingPart(saved));
            Assert.Equal(new[] { "p", "li", "li", "p" }, BodyKinds(saved));
        }
        else
        {
            Assert.False(HasNumberingPart(saved));
            Assert.True(XNode.DeepEquals(Normalized(baseline), Normalized(saved)),
                FirstDifference(Normalized(baseline), Normalized(saved)));
        }
        Assert.True(session.Undo());
        Assert.NotEmpty(session.ListRevisions());
        Assert.True(session.Redo());
        Assert.Empty(session.ListRevisions());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS759h_TrackedReplaceText_RecordsPromotionAsNumberingInsertion(bool accept)
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "List Author",
        });
        var first = FirstBodyParagraph(session);

        var result = session.ReplaceText(first, "- promoted");

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("li", result.Modified[0].Kind);
        Assert.Contains(session.ListRevisions(),
            r => r.Family == RevisionFamily.NumberingPropertiesInsert);
        var resolved = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(resolved.Success, resolved.Error?.Message);

        var saved = session.Save();
        if (accept)
        {
            Assert.Equal(new[] { "li", "p" }, BodyKinds(saved));
            Assert.Equal("promoted", session.GetAnchorInfo(result.Modified[0].Id)!.TextPreview);
        }
        else
        {
            Assert.False(HasNumberingPart(saved));
            Assert.True(XNode.DeepEquals(Normalized(baseline), Normalized(saved)),
                FirstDifference(Normalized(baseline), Normalized(saved)));
        }
    }

    [Fact]
    public void DS759i_ReplaceCellContent_NumbersCellParagraphs()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = FirstBodyParagraph(session);
        Assert.True(session.InsertTable(first, Position.After, 1, 1).Success);
        var cell = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "tc").Anchor.Id;

        var result = session.ReplaceCellContent(cell, "- a\n- b");

        Assert.True(result.Success, result.Error?.Message);
        var cellXml = XElement.Parse(session.Raw.GetXml(cell));
        var numbering = cellXml.Elements(Wns + "p")
            .Select(p => (int?)p.Element(Wns + "pPr")?.Element(Wns + "numPr")?.Element(Wns + "numId")?.Attribute(Wns + "val"))
            .ToArray();
        Assert.Equal(2, numbering.Length);
        Assert.All(numbering, id => Assert.NotNull(id));
        Assert.Equal(numbering[0], numbering[1]);
        Assert.Empty(ValidationErrors(session.Save()));
    }

    [Fact]
    public void DS759j_PreviewAndApply_ReportTheSameListItems()
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(baseline);
        var first = FirstBodyParagraph(session);
        MutationBatchStep Step() => new("docxodus_create", "insert_paragraph",
            s => new[] { s.InsertParagraph(first, Position.After, "1. a\n2. b") });

        var preview = session.PreviewBatch(new[] { Step() });
        Assert.True(XNode.DeepEquals(Normalized(baseline), Normalized(session.Save())));
        var applied = session.ExecuteBatch(new[] { Step() });

        Assert.True(preview.Success && applied.Success);
        Assert.Equal(
            preview.Steps.Single().Results.Single().Created.Select(a => a.Kind),
            applied.Steps.Single().Results.Single().Created.Select(a => a.Kind));
        Assert.Equal(new[] { "p", "li", "li", "p" }, BodyKinds(session.Save()));
    }

    // ─── helpers ────────────────────────────────────────────────────────

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(a => a.Anchor.Scope == "body" && a.Anchor.Kind is "p" or "li").Anchor.Id;

    private static (int NumId, int Ilvl) NumPr(DocxSession session, Anchor anchor)
    {
        var numPr = XElement.Parse(session.Raw.GetXml(anchor.Id))
            .Element(Wns + "pPr")?.Element(Wns + "numPr");
        Assert.NotNull(numPr);
        return ((int)numPr!.Element(Wns + "numId")!.Attribute(Wns + "val")!,
            (int?)numPr.Element(Wns + "ilvl")?.Attribute(Wns + "val") ?? 0);
    }

    private static string[] BodyKinds(byte[] bytes)
    {
        using var session = new DocxSession(bytes);
        return session.Project().AnchorIndex.Values
            .Where(a => a.Anchor.Scope == "body" && a.Anchor.Kind is "p" or "li" or "h")
            .Select(a => a.Anchor.Kind).ToArray();
    }

    private static XElement? NumberingRoot(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var part = document.MainDocumentPart!.NumberingDefinitionsPart;
        return part is null ? null : new XElement(part.GetXDocument().Root!);
    }

    private static bool HasNumberingPart(byte[] bytes) => NumberingRoot(bytes) is not null;

    private static NumberFormat? NumFmt(byte[] bytes, int numId, int ilvl)
    {
        var root = NumberingRoot(bytes);
        var num = root?.Elements(Wns + "num").Single(n => (int)n.Attribute(Wns + "numId")! == numId);
        var abstractId = (string?)num?.Element(Wns + "abstractNumId")?.Attribute(Wns + "val");
        var token = (string?)root?.Elements(Wns + "abstractNum")
            .Single(a => (string?)a.Attribute(Wns + "abstractNumId") == abstractId)
            .Elements(Wns + "lvl").Single(l => (int)l.Attribute(Wns + "ilvl")! == ilvl)
            .Element(Wns + "numFmt")?.Attribute(Wns + "val");
        return token switch
        {
            null => null,
            "bullet" => NumberFormat.Bullet,
            "decimal" => NumberFormat.Decimal,
            _ => throw new InvalidOperationException("unexpected numFmt " + token),
        };
    }

    private static int? StartOverride(byte[] bytes, int numId, int ilvl) =>
        (int?)NumberingRoot(bytes)?.Elements(Wns + "num")
            .Single(n => (int)n.Attribute(Wns + "numId")! == numId)
            .Elements(Wns + "lvlOverride").SingleOrDefault(o => (int)o.Attribute(Wns + "ilvl")! == ilvl)
            ?.Element(Wns + "startOverride")?.Attribute(Wns + "val");

    private static XElement Normalized(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var root = new XElement(document.MainDocumentPart!.GetXDocument().Root!);
        // Anchor ids are session state, and xml:space on a restored w:t is a serialization
        // detail of the delText round trip; neither is document content.
        root.DescendantsAndSelf().Attributes()
            .Where(a => a.Name.Namespace == PtOpenXml.pt || a.IsNamespaceDeclaration
                || a.Name == XNamespace.Xml + "space")
            .Remove();
        return root;
    }

    private static string FirstDifference(XElement expected, XElement actual)
    {
        var expectedNodes = expected.DescendantsAndSelf().ToList();
        var actualNodes = actual.DescendantsAndSelf().ToList();
        static string Shallow(XElement e) => e.Name + "[" + string.Join(" ",
            e.Attributes().OrderBy(a => a.Name.ToString()).Select(a => a.Name + "=" + a.Value))
            + "]" + string.Concat(e.Nodes().OfType<XText>().Select(t => t.Value));
        for (int i = 0; i < Math.Min(expectedNodes.Count, actualNodes.Count); i++)
        {
            if (Shallow(expectedNodes[i]) != Shallow(actualNodes[i]))
                return $"first difference at element {i}: expected {Shallow(expectedNodes[i])}, actual {Shallow(actualNodes[i])}";
        }
        return $"element counts differ: expected {expectedNodes.Count}, actual {actualNodes.Count}";
    }

    private static string[] ValidationErrors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(document)
            .Select(e => $"{e.Id}|{e.Description}|{e.Path?.XPath}")
            .ToArray();
    }
}
