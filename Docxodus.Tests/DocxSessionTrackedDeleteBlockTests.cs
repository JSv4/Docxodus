// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionTrackedDeleteBlockTests
{
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    public static IEnumerable<object[]> ContentCases =>
        from shape in new[] { "hyperlink", "field", "sdt", "instruction", "insertion", "textbox" }
        from resolver in new[] { "session", "diff", "individual" }
        from operation in new[] { "block", "range" }
        from position in new[] { "middle", "last" }
        where operation == "block" || position == "middle"
        select new object[] { shape, resolver, operation, position };

    public static IEnumerable<object[]> TableCases =>
        from shape in new[] { "wrapped", "only-wrapped", "row-format", "row-exceptions", "row-insertion", "nested" }
        from resolver in new[] { "session", "diff", "individual" }
        select new object[] { shape, resolver };

    public static IEnumerable<object[]> StoryCases =>
        from story in new[] { "header", "footer", "footnote", "endnote", "cell", "sdt" }
        from resolver in new[] { "session", "diff", "individual" }
        select new object[] { story, resolver };

    public static IEnumerable<object[]> RoundTripCases =>
        from kind in new[] { "p", "h", "li", "empty", "tbl", "nested-table" }
        from resolver in new[] { "session", "diff", "individual" }
        select new object[] { kind, resolver };

    [Theory]
    [MemberData(nameof(RoundTripCases))]
    public void DeleteBlock_AcceptsToCleanDeletionAndRejectsToOriginal(string kind, string resolver)
    {
        var original = BuildNumberedDocument(P("before"), TargetBlock(kind), NumberedParagraph("after"));
        byte[] clean;
        using (var session = new DocxSession(original))
        {
            Assert.True(session.DeleteBlock(TargetAnchor(session)).Success);
            clean = session.Save();
        }

        byte[] tracked;
        using (var session = new DocxSession(original))
        {
            // Match the public repro: recording and author can be switched after opening.
            session.SetTrackedChanges(TrackedChangeMode.RenderInline);
            session.SetRevisionAuthor("Reviewer");
            var anchor = TargetAnchor(session);
            Assert.Equal(kind is "empty" ? "p" : kind is "nested-table" ? "tbl" : kind,
                session.GetAnchorInfo(anchor)!.Kind);
            var anchorsBefore = session.Project().AnchorIndex.Keys.ToHashSet();

            var result = session.DeleteBlock(anchor);

            Assert.True(result.Success, result.Error?.Message);
            Assert.Equal(anchor, Assert.Single(result.Modified).Id);
            Assert.Empty(result.Removed);
            Assert.Empty(result.Created);
            Assert.True(anchorsBefore.SetEquals(session.Project().AnchorIndex.Keys));
            Assert.Equal(1, session.UndoCount);
            tracked = session.Save();
        }

        AssertValid(tracked);
        var target = Body(tracked).Elements().ElementAt(1);
        Assert.All(target.DescendantsAndSelf(W.p), paragraph =>
            Assert.NotNull(paragraph.Element(W.pPr)?.Element(W.rPr)?.Element(W.del)));
        Assert.All(target.Descendants(W.tr), row =>
            Assert.NotNull(row.Element(W.trPr)?.Element(W.del)));
        Assert.NotEmpty(target.Descendants(W.del));
        Assert.All(target.Descendants(W.del), deletion =>
        {
            Assert.Equal("Reviewer", (string?)deletion.Attribute(W.author));
            Assert.NotNull(deletion.Attribute(W.id));
            Assert.NotNull(deletion.Attribute(W.date));
        });

        var accepted = Resolve(tracked, resolver, accept: true);
        var rejected = Resolve(tracked, resolver, accept: false);
        // Whole-body XML includes paragraph/table counts and numbering, which a text-only
        // comparison misses when acceptance strands an empty numbered paragraph (#784).
        AssertSameBody(clean, accepted);
        AssertSameBody(original, rejected);
        AssertValid(accepted);
        AssertValid(rejected);
    }

    [Theory]
    [InlineData("p")]
    [InlineData("tbl")]
    public void DeleteBlock_LastSibling_TracksOnlyTheTargetAndSupportsUndoRedo(string kind)
    {
        var original = BuildNumberedDocument(P("before"), TargetBlock(kind));
        using var session = Open(original);
        Assert.True(session.DeleteBlock(TargetAnchor(session)).Success);
        Assert.NotEmpty(session.ListRevisions());
        var tracked = session.Save();

        Assert.True(session.Undo());
        AssertSameBody(original, session.Save());
        Assert.Empty(session.ListRevisions());
        Assert.True(session.Redo());
        AssertSameBody(tracked, session.Save());
        Assert.True(session.AcceptAllRevisions().Success);
        Assert.Equal("before", Assert.Single(Body(session.Save()).Elements(W.p)).Value);
        Assert.Empty(Body(session.Save()).Descendants(W.tbl));
    }

    [Theory]
    [InlineData("p")]
    [InlineData("tbl")]
    public void DeleteBlock_PreservesReferencedBookmarksOrRefusesBeforeRecording(string kind)
    {
        var target = TargetBlock(kind);
        var paragraph = target.Name == W.p ? target : target.Descendants(W.p).First();
        paragraph.AddFirst(E("bookmarkStart", A("id", "1"), A("name", "TableBookmark")));
        paragraph.Add(E("bookmarkEnd", A("id", "1")));
        var original = BuildNumberedDocument(
            P("before"), target,
            E("p", E("hyperlink", A("anchor", "TableBookmark"), Run("see table"))));

        using (var untracked = new DocxSession(original))
        {
            var refused = untracked.DeleteBlock(TargetAnchor(untracked));
            Assert.Equal(EditErrorCode.BookmarkInUse, refused.Error?.Code);
            Assert.Equal(0, untracked.UndoCount);
        }

        using var session = Open(original);
        var result = session.DeleteBlock(TargetAnchor(session));
        if (kind == "tbl")
        {
            Assert.Equal(EditErrorCode.BookmarkInUse, result.Error?.Code);
            Assert.Equal(0, session.UndoCount);
            Assert.Empty(session.ListRevisions());
        }
        else Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(result.Removed);
        Assert.Equal("TableBookmark", Assert.Single(session.ListBookmarks()).Name);
        Assert.Equal(kind == "tbl" ? 1 : 0, Body(session.Save()).Descendants(W.tbl).Count());
        Assert.True(session.RejectAllRevisions().Success);
        AssertSameBody(original, session.Save());
        AssertValid(session.Save());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeleteBlock_InlineCustomXml_RefusesBeforeMutation(bool inTable)
    {
        var paragraph = E("p", E("customXml", A("uri", "urn:docxodus:test"), A("element", "inline"),
            E("customXmlPr"), Run("unsupported")));
        var target = inTable ? Table(E("tr", E("tc", paragraph))) : paragraph;
        using var session = Open(BuildNumberedDocument(P("before"), target, P("after")));
        var anchor = TargetAnchor(session);
        var before = session.Save();

        var result = session.DeleteBlock(anchor);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.IncompatibleElementType, result.Error?.Code);
        Assert.Contains("run-level w:customXml", result.Error?.Message);
        Assert.Equal(0, session.UndoCount);
        Assert.Empty(session.ListRevisions());
        AssertSameBody(before, session.Save());
    }

    [Theory]
    [MemberData(nameof(ContentCases))]
    public void DeleteNestedContent_AcceptsWithoutLeakingAndRejectsToBaseline(
        string shape, string resolver, string operation, string position)
    {
        var payload = shape switch
        {
            "hyperlink" => E("hyperlink", new XAttribute(R.id, "rIdLink"), Run("delete")),
            "field" => E("fldSimple", A("instr", "DATE"), Run("delete")),
            "sdt" => E("sdt", E("sdtPr", E("id", A("val", "42"))), E("sdtContent", Run("delete"))),
            "instruction" => E("r", E("instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), " DATE ")),
            "insertion" => E("ins", Stamp(), Run("delete")),
            "textbox" => TextBox(),
            _ => throw new ArgumentOutOfRangeException(nameof(shape)),
        };
        var blocks = new List<XElement> { P("before"), E("p", Run("ordinary"), payload) };
        if (position == "middle") blocks.Add(P("after"));
        var input = Build(blocks.ToArray(),
            (main, _) => main.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rIdLink"));
        AssertValid(input);
        using var session = Open(input);
        var anchor = Anchor(session, "p", 1);
        var anchorsBefore = session.Project().AnchorIndex.Keys.ToHashSet();
        var result = operation == "block" ? session.DeleteBlock(anchor)
            : session.DeleteRange(anchor, session.Project().AnchorIndex.Values
                .Single(a => a.Anchor.Kind == "p" && session.GetAnchorInfo(a.Anchor.Id)!.TextPreview == "after").Anchor.Id);
        Assert.True(result.Success, result.Error?.Message);
        Assert.True(anchorsBefore.SetEquals(session.Project().AnchorIndex.Keys),
            "recording deletion must keep descendant anchors live until review");
        var tracked = session.Save();
        AssertValid(tracked);
        Assert.Empty(Body(tracked).Descendants(W.instrText));
        if (shape == "instruction") Assert.Single(Body(tracked).Descendants(W.delInstrText));
        var accepted = Resolve(tracked, resolver, true);
        var expected = position == "middle" ? E("body", P("before"), P("after")) : E("body", P("before"));
        AssertXmlEqual(expected, Body(accepted));
        AssertValid(accepted);
        var rejected = Resolve(tracked, resolver, false);
        AssertXmlEqual(Body(Resolve(input, resolver, false)), Body(rejected));
        AssertValid(rejected);
    }

    [Fact]
    public void RejectingOnlyNewDeletion_RestoresAnEarlierInsertion()
    {
        using var session = Open(Build(new[] { P("before"), P("after") }));
        session.SetRevisionAuthor("Earlier");
        var insertion = session.InsertParagraph(Anchor(session, "p", 0), Position.After, "inserted");
        Assert.True(insertion.Success, insertion.Error?.Message);
        var before = session.Save();
        session.SetRevisionAuthor("Deleter");
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var deletions = session.ListRevisions().Where(r => r.Author == "Deleter").ToList();
        Assert.NotEmpty(deletions);
        foreach (var deletion in deletions)
            Assert.True(session.RejectRevision(deletion.Id).Success);
        AssertXmlEqual(Body(before), Body(session.Save()));
        Assert.NotEmpty(session.ListRevisions());
        Assert.True(session.AcceptAllRevisions().Success);
        Assert.Equal(new[] { "before", "inserted", "after" }, Body(session.Save()).Elements(W.p).Select(p => p.Value));
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void InsertThenDelete_ResolvesToTheOriginalBlocks(string resolver)
    {
        var original = Build(new[] { P("before"), P("after") });
        using var session = Open(original);
        Assert.True(session.InsertParagraph(Anchor(session, "p", 0), Position.After, "inserted").Success);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var tracked = session.Save();
        AssertValid(tracked);
        AssertXmlEqual(Body(original), Body(Resolve(tracked, resolver, true)));
        AssertXmlEqual(Body(original), Body(Resolve(tracked, resolver, false)));
    }

    [Theory]
    [MemberData(nameof(TableCases))]
    public void DeleteTable_MarksWrappedRowsInSchemaOrderAndRoundTrips(string shape, string resolver)
    {
        var wrapped = E("sdt", E("sdtPr", E("id", A("val", "43"))), E("sdtContent", Row("wrapped")));
        var table = shape switch
        {
            "wrapped" => Table(Row("direct"), wrapped),
            "only-wrapped" => Table(wrapped),
            "row-format" => Table(E("tr", E("trPr", E("trPrChange", Stamp(), E("trPr"))), E("tc", P("cell")))),
            "row-exceptions" => Table(E("tr", E("tblPrEx", E("shd", A("val", "clear"), A("fill", "EEEEEE"))), E("tc", P("cell")))),
            "row-insertion" => Table(E("tr", E("trPr", E("ins", Stamp())), E("tc", P("cell")))),
            "nested" => Table(E("tr", E("tc", Table(wrapped), P("cell")))),
            _ => throw new ArgumentOutOfRangeException(nameof(shape)),
        };
        var input = Build(new[] { P("before"), table, P("after") });
        AssertValid(input);
        using var session = Open(input);
        var anchorsBefore = session.Project().AnchorIndex.Keys.ToHashSet();
        Assert.True(session.DeleteBlock(Anchor(session, "tbl", 0)).Success);
        Assert.True(anchorsBefore.SetEquals(session.Project().AnchorIndex.Keys),
            "recording deletion must keep wrapped row and cell anchors live until review");
        var tracked = session.Save();
        var rows = Body(tracked).Descendants(W.tr).ToList();
        Assert.NotEmpty(rows);
        Assert.All(rows, row => Assert.NotNull(row.Element(W.trPr)?.Element(W.del)));
        AssertValid(tracked);
        var accepted = Resolve(tracked, resolver, true);
        Assert.Empty(Body(accepted).Descendants(W.tbl));
        Assert.Equal(new[] { "before", "after" }, Body(accepted).Elements(W.p).Select(p => p.Value));
        AssertValid(accepted);
        var rejected = Resolve(tracked, resolver, false);
        AssertXmlEqual(Body(Resolve(input, resolver, false)), Body(rejected));
        AssertValid(rejected);
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void DeletingADirectRow_PreservesAnUnmarkedWrappedRow(string resolver)
    {
        var input = Build(new[] { Table(Row("remove"),
            E("sdt", E("sdtPr"), E("sdtContent", Row("keep")))) });
        using var session = Open(input);
        Assert.True(session.DeleteTableRow(Anchor(session, "tc", 0)).Success);
        var accepted = Resolve(session.Save(), resolver, true);
        Assert.Equal("keep", Assert.Single(Body(accepted).Descendants(W.tr)).Value);
        AssertValid(accepted);
    }

    [Fact]
    public void RowDeletion_PlacesPropertiesAfterTableExceptions()
    {
        var input = Build(new[] { Table(E("tr", E("tblPrEx"), E("tc", P("cell"))), Row("keep")) });
        AssertValid(input);
        using var session = Open(input);
        Assert.True(session.DeleteTableRow(Anchor(session, "tc", 0)).Success);
        AssertValid(session.Save());
    }

    [Theory]
    [InlineData("block")]
    [InlineData("range")]
    public void ReferencedTableBookmark_IsRefusedBeforeRecording(string operation)
    {
        var input = Build(new[] { P("before"), Table(E("tr", E("tc", E("p",
            E("bookmarkStart", A("id", "1"), A("name", "Target")), Run("delete"), E("bookmarkEnd", A("id", "1")))))),
            E("p", E("hyperlink", A("anchor", "Target"), Run("see target"))) });
        using var session = Open(input);
        var before = session.Save();
        var target = Anchor(session, "tbl", 0);
        var result = operation == "block" ? session.DeleteBlock(target)
            : session.DeleteRange(target, Anchor(session, "p", 2));
        Assert.Equal(EditErrorCode.BookmarkInUse, result.Error?.Code);
        Assert.Equal(0, session.UndoCount);
        Assert.Empty(session.ListRevisions());
        AssertXmlEqual(Body(before), Body(session.Save()));
    }

    [Theory]
    [MemberData(nameof(StoryCases))]
    public void SoleStoryParagraph_KeepsTheSameEmptyFormattedParagraph(string story, string resolver)
    {
        var paragraph = E("p", Properties(), Run("delete"));
        string partName = "word/document.xml";
        byte[] input;
        if (story == "cell") input = Build(new[] { Table(E("tr", E("tc", paragraph))) });
        else if (story == "sdt") input = Build(new[] { E("sdt", E("sdtPr"), E("sdtContent", paragraph)) });
        else input = Build(new[] { P("body") }, (main, body) =>
        {
            var section = E("sectPr");
            if (story == "header")
            {
                var part = main.AddNewPart<HeaderPart>();
                part.PutXDocument(new XDocument(E("hdr", paragraph)));
                section.Add(E("headerReference", A("type", "default"), new XAttribute(R.id, main.GetIdOfPart(part))));
                partName = "word/header1.xml";
            }
            else if (story == "footer")
            {
                var part = main.AddNewPart<FooterPart>();
                part.PutXDocument(new XDocument(E("ftr", paragraph)));
                section.Add(E("footerReference", A("type", "default"), new XAttribute(R.id, main.GetIdOfPart(part))));
                partName = "word/footer1.xml";
            }
            else if (story == "footnote")
            {
                main.AddNewPart<FootnotesPart>().PutXDocument(new XDocument(E("footnotes", E("footnote", A("id", "1"), paragraph))));
                body.Elements(W.p).First().Add(E("r", E("footnoteReference", A("id", "1"))));
                partName = "word/footnotes.xml";
            }
            else
            {
                main.AddNewPart<EndnotesPart>().PutXDocument(new XDocument(E("endnotes", E("endnote", A("id", "1"), paragraph))));
                body.Elements(W.p).First().Add(E("r", E("endnoteReference", A("id", "1"))));
                partName = "word/endnotes.xml";
            }
            body.Add(section);
        });
        AssertValid(input);
        using var session = Open(input);
        var target = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p"
            && session.GetAnchorInfo(a.Anchor.Id)!.TextPreview == "delete");
        Assert.True(session.DeleteBlock(target.Anchor.Id).Success);
        var tracked = session.Save();
        var accepted = Resolve(tracked, resolver, true);
        var actual = Assert.Single(Root(accepted, partName).Descendants(W.p),
            p => story is not ("cell" or "sdt") || p.Ancestors().Any(a => a.Name == W.tc || a.Name == W.sdtContent));
        AssertXmlEqual(E("p", Properties()), actual);
        AssertXmlEqual(Root(input, partName), Root(Resolve(tracked, resolver, false), partName));
        AssertValid(accepted);
    }

    [Fact]
    public void IdenticalHeaders_ReportTheAddressedStory()
    {
        var input = Build(new[] { P("body") }, (main, body) =>
        {
            var section = E("sectPr");
            foreach (var type in new[] { "default", "first" })
            {
                var part = main.AddNewPart<HeaderPart>();
                part.PutXDocument(new XDocument(E("hdr", P("identical"))));
                section.Add(E("headerReference", A("type", type), new XAttribute(R.id, main.GetIdOfPart(part))));
            }
            section.Add(E("titlePg")); body.Add(section);
        });
        using var session = Open(input);
        var target = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p" && a.Anchor.Scope == "hdr2");
        var result = session.DeleteBlock(target.Anchor.Id);
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(target.Anchor.Id, Assert.Single(result.Modified).Id);
        AssertXmlEqual(Root(input, "word/header1.xml"), Root(session.Save(), "word/header1.xml"));
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void SuccessiveCellParagraphDeletions_KeepTheFinalFormatting(string resolver)
    {
        var input = Build(new[] { Table(E("tr", E("tc", P("first"), E("p", Properties(), Run("last"))))) });
        using var session = Open(input);
        var paragraphs = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "p")
            .Select(a => a.Anchor.Id).ToList();
        foreach (var paragraph in paragraphs) Assert.True(session.DeleteBlock(paragraph).Success);
        var accepted = Resolve(session.Save(), resolver, true);
        AssertXmlEqual(E("p", Properties()), Assert.Single(Body(accepted).Descendants(W.p)));
        AssertValid(accepted);
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void DeletingAnotherParagraph_PreservesAnUnrevisedEmptyField(string resolver)
    {
        var field = E("p", E("fldSimple", A("instr", "DATE")));
        var input = Build(new[] { P("before"), P("delete"), field });
        AssertValid(input);
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var accepted = Resolve(session.Save(), resolver, true);
        AssertXmlEqual(field, Body(accepted).Elements(W.p).Last());
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void DeletingTableAndItsReferenceTogether_IsAllowed(string resolver)
    {
        var input = Build(new[] { Table(E("tr", E("tc", E("p",
            E("bookmarkStart", A("id", "1"), A("name", "Target")), Run("delete"), E("bookmarkEnd", A("id", "1")))))),
            E("p", E("hyperlink", A("anchor", "Target"), Run("see target"))), P("after") });
        using var session = Open(input);
        var result = session.DeleteRange(Anchor(session, "tbl", 0), Anchor(session, "p", 2));
        Assert.True(result.Success, result.Error?.Message);
        var accepted = Resolve(session.Save(), resolver, true);
        AssertXmlEqual(E("body", P("after")), Body(accepted));
        AssertValid(accepted);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphDeletion_RefusesBookmarksThatCannotMigrate(bool inlineControl)
    {
        var marked = E("p", E("bookmarkStart", A("id", "1"), A("name", "Target")),
            Run("delete"), E("bookmarkEnd", A("id", "1")));
        if (inlineControl) marked = E("p", E("sdt", E("sdtPr"), E("sdtContent", marked.Nodes())));
        var blocks = new List<XElement>
        {
            E("p", E("hyperlink", A("anchor", "Target"), Run("see target"))), marked,
        };
        if (inlineControl) blocks.Add(P("after"));
        var input = Build(blocks.ToArray());
        AssertValid(input);
        using var session = Open(input);
        Assert.Equal(EditErrorCode.BookmarkInUse, session.DeleteBlock(Anchor(session, "p", 1)).Error?.Code);
        Assert.Equal(0, session.UndoCount);
        AssertXmlEqual(Body(input), Body(session.Save()));
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void RecordingAndRejecting_PreserveUnrelatedRelationships(string resolver)
    {
        var input = Build(new[] { P("before"), P("delete"), P("after") },
            (main, _) => main.AddHyperlinkRelationship(new Uri("https://example.com/orphan"), true, "rIdOrphan"));
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var tracked = session.Save();
        Assert.Equal(RelationshipIds(input), RelationshipIds(tracked));
        Assert.Equal(RelationshipIds(input), RelationshipIds(Resolve(tracked, resolver, false)));
    }

    public static IEnumerable<object[]> TrailingParagraphCases =>
        from story in new[] { "cell", "header" }
        from resolver in new[] { "session", "diff", "individual" }
        select new object[] { story, resolver };

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void DeletingASectionBreakParagraph_AcceptsAcrossEveryResolver(string resolver)
    {
        var input = Build(new[] { P("first section") }, (main, body) =>
        {
            var header = main.AddNewPart<HeaderPart>();
            header.PutXDocument(new XDocument(E("hdr", P("section one header"))));
            body.Add(E("p", E("pPr", E("sectPr", E("headerReference", A("type", "default"),
                new XAttribute(R.id, main.GetIdOfPart(header))))), Run("section break")));
            body.Add(P("second section"));
            body.Add(E("sectPr"));
        });
        AssertValid(input);
        byte[] clean;
        using (var untracked = new DocxSession(input))
        {
            Assert.True(untracked.DeleteBlock(BodyParagraph(untracked, 1)).Success);
            clean = untracked.Save();
        }
        using var session = Open(input);
        Assert.True(session.DeleteBlock(BodyParagraph(session, 1)).Success);
        var tracked = session.Save();
        AssertValid(tracked);
        var accepted = Resolve(tracked, resolver, true);
        AssertSameBody(clean, accepted);
        AssertValid(accepted);
        AssertSameBody(input, Resolve(tracked, resolver, false));
    }

    [Theory]
    [MemberData(nameof(TrailingParagraphCases))]
    public void LastParagraphAfterATable_KeepsItsPilcrow(string story, string resolver)
    {
        var paragraph = E("p", Properties(), Run("delete"));
        string partName = "word/document.xml";
        var input = story == "cell"
            ? Build(new[] { Table(E("tr", E("tc", Table(Row("inner")), paragraph))) })
            : Build(new[] { P("body") }, (main, body) =>
            {
                var part = main.AddNewPart<HeaderPart>();
                part.PutXDocument(new XDocument(E("hdr", Table(Row("inner")), paragraph)));
                body.Add(E("sectPr", E("headerReference", A("type", "default"), new XAttribute(R.id, main.GetIdOfPart(part)))));
                partName = "word/header1.xml";
            });
        AssertValid(input);
        using var session = Open(input);
        Assert.True(session.DeleteBlock(ParagraphWithText(session, "delete")).Success);
        var tracked = session.Save();
        var accepted = Resolve(tracked, resolver, true);
        var container = story == "cell" ? Root(accepted, partName).Descendants(W.tc).First() : Root(accepted, partName);
        AssertXmlEqual(E("p", Properties()), container.Elements().Last());
        AssertXmlEqual(Root(input, partName), Root(Resolve(tracked, resolver, false), partName));
        AssertValid(accepted);
    }

    [Fact]
    public void TrackedRangeDeletion_ScalesWithTheRange()
    {
        var blocks = Enumerable.Range(0, 4000).Select(i => P($"paragraph {i}")).Append(P("survivor")).ToArray();
        using var session = Open(Build(blocks));
        var paragraphs = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "p").Select(a => a.Anchor.Id).ToList();
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        Assert.True(session.DeleteRange(paragraphs[0], paragraphs[^1]).Success);
        stopwatch.Stop();
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10), $"tracked DeleteRange over 4000 paragraphs took {stopwatch.Elapsed}");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BookmarkGuard_SeesThroughAnAlreadyDeletedNeighbour(bool deleteFollowingFirst)
    {
        var input = Build(new[]
        {
            E("p", E("hyperlink", A("anchor", "Target"), Run("see target"))),
            E("p", E("bookmarkStart", A("id", "1"), A("name", "Target")), Run("bookmarked"), E("bookmarkEnd", A("id", "1"))),
            P("following"), Table(Row("cell")), P("after"),
        });
        using var session = Open(input);
        var first = deleteFollowingFirst ? "following" : "bookmarked";
        var second = deleteFollowingFirst ? "bookmarked" : "following";
        Assert.True(session.DeleteBlock(ParagraphWithText(session, first)).Success);
        var before = session.Save();
        var result = session.DeleteBlock(ParagraphWithText(session, second));
        Assert.Equal(EditErrorCode.BookmarkInUse, result.Error?.Code);
        Assert.Equal(1, session.UndoCount);
        AssertXmlEqual(Body(before), Body(session.Save()));
    }

    [Fact]
    public void DeletingAMoveDestination_KeepsTheParagraphMarkInSchemaOrder()
    {
        using var session = Open(Build(new[] { P("one"), P("two"), P("three") }));
        session.SetRevisionAuthor("Mover");
        var move = session.MoveBlock(Anchor(session, "p", 0), Anchor(session, "p", 2), Position.After);
        Assert.True(move.Success, move.Error?.Message);
        session.SetRevisionAuthor("Deleter");
        var destination = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "p").Last().Anchor.Id;
        var result = session.DeleteBlock(destination);
        Assert.True(result.Success, result.Error?.Message);
        var tracked = session.Save();
        AssertValid(tracked);
        var mark = Body(tracked).Elements(W.p).Last().Element(W.pPr)!.Element(W.rPr)!;
        Assert.Equal(new[] { W.del, W.moveTo }, mark.Elements().Select(e => e.Name).ToArray());
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void Retention_IgnoresSiblingsAlreadyDeletedInsideWrappers(string resolver)
    {
        var input = Build(new[] { Table(E("tr", E("tc", E("p", Properties(), Run("keep formatting")), P("x"),
            E("sdt", E("sdtPr"), E("sdtContent", P("a"))), P("c")))) });
        AssertValid(input);
        using var session = Open(input);
        Assert.True(session.DeleteRange(ParagraphWithText(session, "x"), ParagraphWithText(session, "c")).Success);
        Assert.True(session.DeleteBlock(ParagraphWithText(session, "c")).Success);
        Assert.True(session.DeleteBlock(ParagraphWithText(session, "keep formatting")).Success);
        var accepted = Resolve(session.Save(), resolver, true);
        AssertXmlEqual(E("p", Properties()), Assert.Single(Body(accepted).Descendants(W.p)));
        AssertValid(accepted);
    }

    [Fact]
    public void ComparingAwayATableWithAWrappedRow_AcceptsToTheRightDocument()
    {
        var left = Build(new[] { P("x"), Table(Row("a"), E("sdt", E("sdtPr"), E("sdtContent", Row("b")))) });
        var right = Build(new[] { P("x") });
        var redline = DocxCompare.Compare(new WmlDocument("left.docx", left), new WmlDocument("right.docx", right)).DocumentByteArray;
        foreach (var resolver in new[] { "diff", "session", "individual" })
        {
            var accepted = Body(Resolve(redline, resolver, true));
            Assert.Empty(accepted.Descendants(W.tbl));
            Assert.Equal(new[] { "x" }, accepted.Elements(W.p).Select(p => p.Value));
        }
    }

    [Theory]
    [InlineData("session")]
    [InlineData("individual")]
    public void AcceptingADeletedCitation_SweepsThePrunedNotesMedia(string resolver)
    {
        var input = Build(new[] { P("before"), E("p", Run("cite"), E("r", E("footnoteReference", A("id", "1")))), P("after") },
            (main, _) =>
            {
                var notes = main.AddNewPart<FootnotesPart>();
                var image = notes.AddImagePart(ImagePartType.Png, "rIdNoteImage");
                using (var bytes = new MemoryStream(new byte[] { 137, 80, 78, 71 })) image.FeedData(bytes);
                notes.PutXDocument(new XDocument(E("footnotes", E("footnote", A("id", "1"), E("p", Run("note"),
                    E("r", E("pict", new XElement(V + "shape", new XElement(V + "imagedata", new XAttribute(R.id, "rIdNoteImage"))))))))));
            });
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        using var resolved = WordprocessingDocument.Open(new MemoryStream(Resolve(session.Save(), resolver, true)), false);
        var notesPart = resolved.MainDocumentPart!.FootnotesPart!;
        Assert.Empty(notesPart.GetXDocument().Descendants(W.footnote));
        Assert.Empty(notesPart.ImageParts);
    }

    [Fact]
    public void AcceptingOnlyTheDeletionOfInsertedContent_LeavesNoInsertionHusk()
    {
        using var session = Open(Build(new[] { P("before"), P("after") }));
        session.SetRevisionAuthor("Earlier");
        Assert.True(session.InsertParagraph(Anchor(session, "p", 0), Position.After, "inserted").Success);
        session.SetRevisionAuthor("Deleter");
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        foreach (var deletion in session.ListRevisions().Where(r => r.Author == "Deleter").ToList())
            Assert.True(session.AcceptRevision(deletion.Id).Success);
        Assert.Empty(session.ListRevisions());
        AssertXmlEqual(E("body", P("before"), P("after")), Body(session.Save()));
    }

    [Fact]
    public void DeletingAnOwnInsertedParagraph_PrunesTheNoteInsertedWithIt()
    {
        using var session = Open(Build(new[] { P("before"), P("after") }));
        Assert.True(session.InsertParagraph(Anchor(session, "p", 0), Position.After, "inserted").Success);
        var note = session.InsertFootnote(Anchor(session, "p", 1), 8, "own note");
        Assert.True(note.Success, note.Error?.Message);
        Assert.Single(BodyNotes(session.Save()));

        var deletion = session.DeleteBlock(Anchor(session, "p", 1));

        Assert.True(deletion.Success, deletion.Error?.Message);
        var tracked = session.Save();
        AssertValid(tracked);
        // The author's own inserted text is un-inserted outright, so the note it carried has no
        // reference left anywhere, goes with it, and nothing in its body is listed as pending.
        Assert.Empty(Root(tracked).Descendants(W.footnoteReference));
        Assert.Empty(BodyNotes(tracked));
        Assert.All(session.ListRevisions(), r => Assert.DoesNotContain("own note", r.Text));
        Assert.True(session.AcceptAllRevisions().Success);
        AssertXmlEqual(E("body", P("before"), P("after")), Body(session.Save()));

        static IEnumerable<XElement> BodyNotes(byte[] bytes) => Root(bytes, "word/footnotes.xml")
            .Descendants(W.footnote).Where(f => f.Attribute(W.w + "type") is null);
    }

    [Fact]
    public void InlineControlHoldingATextBox_IsDeletedWithItsParagraph()
    {
        var control = E("sdt", E("sdtPr", E("id", A("val", "7"))), E("sdtContent", TextBox()));
        var input = Build(new[] { P("before"), E("p", Run("ordinary"), control), P("after") });
        AssertValid(input);
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var tracked = session.Save();
        AssertValid(tracked);
        foreach (var resolver in new[] { "session", "diff", "individual" })
            AssertXmlEqual(E("body", P("before"), P("after")), Body(Resolve(tracked, resolver, true)));
    }

    [Theory]
    [InlineData("session")]
    [InlineData("diff")]
    [InlineData("individual")]
    public void DeletingAParagraphWithMath_RemovesTheMathOnAcceptance(string resolver)
    {
        var math = new XElement(M.oMath, new XElement(M.r, new XElement(M.t, "x+y")));
        var input = Build(new[] { P("before"), E("p", Run("eq: "), math), P("after") });
        AssertValid(input);
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var tracked = session.Save();
        AssertValid(tracked);
        AssertXmlEqual(E("body", P("before"), P("after")), Body(Resolve(tracked, resolver, true)));
        AssertXmlEqual(Body(input), Body(Resolve(tracked, resolver, false)));
    }

    [Fact]
    public void DeletingAParagraphWithAnEmptyField_IsRefusedBeforeMutation()
    {
        var input = Build(new[] { P("before"), E("p", Run("delete"), E("fldSimple", A("instr", " PAGE "))), P("after") });
        using var session = Open(input);
        var result = session.DeleteBlock(Anchor(session, "p", 1));
        Assert.Equal(EditErrorCode.IncompatibleElementType, result.Error?.Code);
        Assert.Equal(0, session.UndoCount);
        AssertXmlEqual(Body(input), Body(session.Save()));
    }

    [Fact]
    public void AcceptedView_HidesATableWhoseWrappedRowsAreAllDeleted()
    {
        var input = Build(new[] { P("before"), Table(E("sdt", E("sdtPr"), E("sdtContent", Row("wrapped")))), P("after") });
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "tbl", 0)).Success);
        Assert.DoesNotContain(session.ListBlocks(renderTrackedChanges: false).Body, unit => unit.Kind == "tbl");
    }

    [Fact]
    public void UntrackedRowDeletion_KeepsAWrappedRowAlive()
    {
        var input = Build(new[] { Table(Row("remove"), E("sdt", E("sdtPr"), E("sdtContent", Row("keep")))) });
        using var session = new DocxSession(input);
        Assert.True(session.DeleteTableRow(Anchor(session, "tc", 0)).Success);
        var saved = session.Save();
        Assert.Equal("keep", Assert.Single(Body(saved).Descendants(W.tr)).Value);
        AssertValid(saved);
    }

    [Theory]
    [InlineData("dir")]
    [InlineData("bdo")]
    public void DeletingBidirectionalRuns_RoundTripsThroughEveryResolver(string container)
    {
        // The SDK validator's default schema version predates w:dir/w:bdo as paragraph
        // children, so this shape is checked by round trip rather than by AssertValid.
        var input = Build(new[] { P("before"), E("p", E(container, A("val", container == "dir" ? "ltr" : "rtl"), Run("text"))), P("after") });
        using var session = Open(input);
        Assert.True(session.DeleteBlock(Anchor(session, "p", 1)).Success);
        var tracked = session.Save();
        foreach (var resolver in new[] { "session", "diff", "individual" })
        {
            AssertXmlEqual(E("body", P("before"), P("after")), Body(Resolve(tracked, resolver, true)));
            AssertXmlEqual(Body(input), Body(Resolve(tracked, resolver, false)));
        }
    }

    private static string BodyParagraph(DocxSession session, int skip) => session.Project().AnchorIndex.Values
        .Where(a => a.Anchor.Kind == "p" && a.Anchor.Scope == "body").Skip(skip).First().Anchor.Id;

    private static string ParagraphWithText(DocxSession session, string text) => session.Project().AnchorIndex.Values
        .Single(a => a.Anchor.Kind == "p" && session.GetAnchorInfo(a.Anchor.Id)!.TextPreview == text).Anchor.Id;

    private static string TargetAnchor(DocxSession session) => session.Project().AnchorIndex.Values
        .Where(target => target.Anchor.Kind is "p" or "h" or "li" or "tbl")
        .Skip(1).First().Anchor.Id;

    private static XElement TargetBlock(string kind) => kind switch
    {
        "p" => P("delete me"),
        "h" => NumberedParagraph("delete me", heading: true),
        "li" => NumberedParagraph("delete me"),
        "empty" => E("p"),
        "tbl" => Table(Row("Cell A"), Row("Cell B")),
        "nested-table" => Table(E("tr", E("tc", P("Cell A"),
            Table(Row("Cell A"), Row("Cell B")), E("p"))), Row("Cell B")),
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    private static XElement NumberedParagraph(string text, bool heading = false) => E("p", E("pPr",
        heading ? E("pStyle", A("val", "Heading2")) : null,
        E("numPr", E("ilvl", A("val", "0")), E("numId", A("val", "1"))),
        E("rPr", E("sz", A("val", "28")))), Run(text));

    private static byte[] BuildNumberedDocument(params XElement[] blocks) => Build(blocks, (main, _) =>
    {
        main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(E("styles", E("docDefaults"),
            E("style", A("type", "paragraph"), A("styleId", "Heading2"), E("name", A("val", "heading 2"))))));
        main.AddNewPart<NumberingDefinitionsPart>().PutXDocument(new XDocument(E("numbering",
            E("abstractNum", A("abstractNumId", "0"), E("lvl", A("ilvl", "0"),
                E("start", A("val", "1")), E("numFmt", A("val", "decimal")), E("lvlText", A("val", "%1.")))),
            E("num", A("numId", "1"), E("abstractNumId", A("val", "0"))))));
    });

    private static void AssertSameBody(byte[] expected, byte[] actual) => AssertXmlEqual(Body(expected), Body(actual));

    private static XElement E(string name, params object?[] content) => new(W.w + name, content);
    private static XAttribute A(string name, string value) => new(W.w + name, value);
    private static XElement Run(string text) => E("r", E("t", text));
    private static XElement P(string text) => E("p", Run(text));
    private static XElement Row(string text) => E("tr", E("tc", P(text)));
    private static XElement Table(params XElement[] rows) => E("tbl", E("tblPr"),
        E("tblGrid", E("gridCol", A("w", "2500"))), rows);
    private static object[] Stamp() => new object[]
    {
        A("id", "99"), A("author", "Earlier"), A("date", "2026-01-01T00:00:00Z"),
    };
    private static XElement Properties() => E("pPr", E("spacing", A("after", "200")), E("jc", A("val", "right")));

    private static XElement TextBox() => E("r", E("pict", new XElement(V + "shape",
        new XAttribute("id", "TextBox1"), new XAttribute("style", "width:100pt;height:30pt"),
        new XElement(V + "textbox", E("txbxContent", E("customXml", A("element", "block"), E("customXmlPr"), P("text box")))))));

    private static DocxSession Open(byte[] bytes) => new(bytes, new DocxSessionSettings
    {
        TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Reviewer",
    });
    private static string Anchor(DocxSession session, string kind, int skip) => session.Project().AnchorIndex.Values
        .Where(a => a.Anchor.Kind == kind).Skip(skip).First().Anchor.Id;

    private static byte[] Build(XElement[] blocks, Action<MainDocumentPart, XElement>? setup = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = E("body", blocks);
            setup?.Invoke(main, body);
            main.PutXDocument(new XDocument(E("document", new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "r", R.r), body)));
            document.Save();
        }
        return stream.ToArray();
    }
    private static XElement Root(byte[] bytes, string part = "word/document.xml")
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(part)!.Open();
        return XElement.Load(stream);
    }
    private static XElement Body(byte[] bytes) => Root(bytes).Element(W.body)!;
    private static string[] RelationshipIds(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return document.MainDocumentPart!.HyperlinkRelationships.Select(r => r.Id).Order().ToArray();
    }
    private static byte[] Resolve(byte[] bytes, string resolver, bool accept)
    {
        if (resolver == "diff") return accept ? DocxDiffOps.AcceptRevisions(bytes) : DocxDiffOps.RejectRevisions(bytes);
        using var session = new DocxSession(bytes);
        if (resolver == "session")
        {
            var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
            Assert.True(result.Success, result.Error?.Message);
        }
        else
        {
            for (var revisions = session.ListRevisions(); revisions.Count > 0; revisions = session.ListRevisions())
            {
                var result = accept ? session.AcceptRevision(revisions[0].Id) : session.RejectRevision(revisions[0].Id);
                Assert.True(result.Success, result.Error?.Message);
                Assert.True(session.ListRevisions().Count < revisions.Count);
            }
        }
        Assert.Empty(session.ListRevisions());
        return session.Save();
    }
    private static void AssertXmlEqual(XElement expected, XElement actual)
    {
        static string Normalize(XElement value)
        {
            value = new XElement(value);
            value.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration || a.Name == PtOpenXml.Unid).Remove();
            foreach (var node in value.Descendants().Reverse().ToList())
                if ((node.Name == W.pPr || node.Name == W.rPr || node.Name == W.trPr)
                    && !node.HasAttributes && !node.HasElements) node.Remove();
            return value.ToString(SaveOptions.DisableFormatting);
        }
        Assert.Equal(Normalize(expected), Normalize(actual));
    }
    private static void AssertValid(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }
}
