// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A bookmark or comment range written between blocks — a direct child of the body, a table cell or a block
/// wrapper, a sibling of <c>w:p</c> — keeps its place through a comparison and through accepting or
/// rejecting revisions: accepting the redline gives the revised document's markers at the revised document's
/// block boundaries, rejecting gives the original's, and a document with no revisions keeps them (issue #1149).
/// </summary>
public class BlockLevelMarkerDiffTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Start(string name, int id = 7) => $"<w:bookmarkStart w:id=\"{id}\" w:name=\"{name}\"/>";

    private static string End(int id = 7) => $"<w:bookmarkEnd w:id=\"{id}\"/>";

    private static string Table(params string[] cellContents) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(cellContents.Select(_ => "<w:gridCol w:w=\"3000\"/>")) + "</w:tblGrid><w:tr>" +
        string.Concat(cellContents.Select(c => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{c}</w:tc>")) +
        "</w:tr></w:tbl>";

    private static string CustomXml(string content) =>
        $"<w:customXml w:uri=\"urn:test\" w:element=\"clause\">{content}</w:customXml>";

    private static string Sdt(string content) =>
        $"<w:sdt><w:sdtPr><w:alias w:val=\"Clause\"/><w:id w:val=\"42\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>";

    private static WmlDocument Doc(string body) => IrTestDocuments.FromParts(body);

    private static WmlDocument CommentDoc(string body) =>
        IrTestDocuments.WithComment("Reviewer", "R", "2026-01-01T00:00:00Z", "Check this", body);

    private const string CommentReference =
        "<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"0\"/></w:r>";

    // ------------------------------------------------------------------------------------- the layout

    private static XDocument Part(WmlDocument document, string path)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        var entry = package.GetEntry(path);
        if (entry is null)
            return new XDocument();
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static XElement Body(WmlDocument document) => Part(document, "word/document.xml").Root!.Element(W + "body")!;

    /// <summary>The block structure of the body: each paragraph as its text, each block-level bookmark or
    /// comment-range marker by kind and name (a bookmark end and a comment marker resolved to the bookmark's
    /// name or the comment's text, since ids are renumbered), and tables, cells and wrappers as nested groups.
    /// Markers inside paragraphs are left out: this is the layout at block boundaries.</summary>
    private static string Layout(WmlDocument document) => StreamLayout(document, Body(document));

    /// <summary>The block structure (see <see cref="Layout"/>) of any block stream: a body, header or footer.</summary>
    private static string StreamLayout(WmlDocument document, XElement body)
    {
        var bookmarkNames = body.Descendants(W + "bookmarkStart")
            .GroupBy(s => (string?)s.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(W + "name") ?? "");
        var commentTexts = (Part(document, "word/comments.xml").Root?.Elements(W + "comment") ?? Enumerable.Empty<XElement>())
            .GroupBy(c => (string?)c.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => string.Concat(g.First().Descendants(W + "t").Select(t => t.Value)));

        string Id(XElement e) => (string?)e.Attribute(W + "id") ?? "";

        IEnumerable<string> Blocks(XElement container)
        {
            foreach (var child in container.Elements())
            {
                var n = child.Name;
                if (n == W + "p")
                    yield return "P:" + string.Concat(child.Descendants(W + "t").Select(t => t.Value));
                else if (n == W + "bookmarkStart")
                    yield return "BS:" + (string?)child.Attribute(W + "name");
                else if (n == W + "bookmarkEnd")
                    yield return "BE:" + bookmarkNames.GetValueOrDefault(Id(child), "?" + Id(child));
                else if (n == W + "commentRangeStart")
                    yield return "CS:" + commentTexts.GetValueOrDefault(Id(child), "?" + Id(child));
                else if (n == W + "commentRangeEnd")
                    yield return "CE:" + commentTexts.GetValueOrDefault(Id(child), "?" + Id(child));
                else if (n == W + "tbl")
                    yield return "TBL[" + string.Join(" | ", child.Elements(W + "tr").Select(tr =>
                        string.Join(" || ", tr.Elements(W + "tc").Select(tc => string.Join(" ", Blocks(tc)))))) + "]";
                else if (n == W + "sdt")
                    yield return "SDT{" + string.Join(" ", Blocks(child.Element(W + "sdtContent") ?? new XElement("x"))) + "}";
                else if (n == W + "customXml")
                    yield return "CX{" + string.Join(" ", Blocks(child)) + "}";
                else if (n == W + "ins" || n == W + "del")
                    yield return n.LocalName + "(" + string.Join(" ", Blocks(child)) + ")";
            }
        }

        return string.Join(" ", Blocks(body));
    }

    private static WmlDocument Accept(WmlDocument d) => RevisionProcessor.AcceptRevisions(d);

    private static WmlDocument Reject(WmlDocument d) => RevisionProcessor.RejectRevisions(d);

    public static TheoryData<string> Engines => new() { "DocxCompare", "DocxDiff" };

    private static WmlDocument Compare(string engine, WmlDocument left, WmlDocument right) =>
        engine == "DocxCompare" ? DocxCompare.Compare(left, right) : DocxDiff.Compare(left, right);

    /// <summary>Accepting the redline gives the revised layout, rejecting it the original's, and the redline
    /// adds no validation errors.</summary>
    private static WmlDocument AssertRoundTrip(string engine, WmlDocument left, WmlDocument right)
    {
        var redline = Compare(engine, left, right);
        Assert.Equal(Layout(right), Layout(Accept(redline)));
        Assert.Equal(Layout(left), Layout(Reject(redline)));
        AssertNoNewValidationErrors(redline, left, right);
        return redline;
    }

    /// <summary>Every validation error of <paramref name="redline"/> also occurs in one of the inputs, compared by
    /// description: the comparison introduced none. (A path-keyed comparison would flag an input's own error, such
    /// as the validator not knowing <c>w:customXml</c>'s attributes, again at a replaced wrapper's second copy.)</summary>
    private static void AssertNoNewValidationErrors(WmlDocument redline, params WmlDocument[] inputs) =>
        AssertNoNewValidationErrors(redline, description => description, inputs);

    private static void AssertNoNewValidationErrors(WmlDocument redline, Func<string, string> normalize,
        params WmlDocument[] inputs)
    {
        HashSet<string> Errors(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            return new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
                .Validate(package).Select(e => normalize(e.Description)).ToHashSet(StringComparer.Ordinal);
        }
        var known = inputs.SelectMany(Errors).ToHashSet(StringComparer.Ordinal);
        Assert.All(Errors(redline), e => Assert.Contains(e, known));
    }

    /// <summary>The validator looks for a header's or footer's comments through a comments part related to that
    /// story part, which Word never writes (comments live in the main part's), so it reports every comment marker in
    /// a header as missing its comment. An input reports that for its own ids; a redline that gives one view's copy
    /// a fresh id reports it for that id too. Compare those reports without the id.</summary>
    private static string WithoutStoryCommentId(string description) =>
        description.StartsWith("Element 'w:comment' referenced by", StringComparison.Ordinal) &&
        description.Contains("does not exist in part 'WordprocessingCommentsPart'. The reference value is '", StringComparison.Ordinal)
            ? System.Text.RegularExpressions.Regex.Replace(description, @"The reference value is '\d+'\.$", "The reference value is '#'.")
            : description;

    // ------------------------------------------------------------- accept/reject of a revision-free document

    public static TheoryData<string> RevisionFreeBodies => new()
    {
        Start("Article") + P("Alpha") + P("Bravo") + End(),
        P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo") + End()),
        P("Lead") + CustomXml(Start("Wrapped") + P("Alpha") + P("Bravo") + End()),
        P("Lead") + Sdt(Start("Controlled") + P("Alpha") + P("Bravo") + End()),
    };

    [Theory]
    [MemberData(nameof(RevisionFreeBodies))]
    public void AcceptAndReject_OfARevisionFreeDocument_KeepBlockLevelBookmarks(string body)
    {
        var document = Doc(body);
        Assert.Contains("BS:", Layout(document));

        Assert.Equal(Layout(document), Layout(Accept(document)));
        Assert.Equal(Layout(document), Layout(Reject(document)));
    }

    [Fact]
    public void AcceptAndReject_OfARevisionFreeDocument_KeepBlockLevelCommentRanges()
    {
        var document = CommentDoc(
            "<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo") + "<w:commentRangeEnd w:id=\"0\"/>" +
            $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");
        Assert.Contains("CS:Check this", Layout(document));

        Assert.Equal(Layout(document), Layout(Accept(document)));
        Assert.Equal(Layout(document), Layout(Reject(document)));
    }

    [Fact]
    public void Accept_KeepsABlockLevelBookmark_NextToARevisedParagraph()
    {
        // A revision elsewhere in the document must not take the bookmark out either.
        var document = Doc(Start("Article") + P("Alpha") +
            "<w:p><w:ins w:id=\"1\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:r><w:t>Bravo</w:t></w:r></w:ins></w:p>" + End());

        Assert.Equal("BS:Article P:Alpha P:Bravo BE:Article", Layout(Accept(document)));
        Assert.Equal("BS:Article P:Alpha P: BE:Article", Layout(Reject(document)));
    }

    // ------------------------------------------------------------------------------------ the body

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        // The issue's reproducer.
        var left = Doc(Start("Article") + P("Alpha") + P("Bravo monthly") + End());
        var right = Doc(Start("Article") + P("Alpha") + P("Bravo weekly") + End());

        var redline = AssertRoundTrip(engine, left, right);
        Assert.Equal("BS:Article", Layout(redline).Split(' ')[0]);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AroundAnInsertedAndADeletedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + Start("Article") + P("Alpha") + P("Bravo") + End() + P("Tail"));
        var right = Doc(P("Lead") + Start("Article") + P("Alpha") + P("Charlie inserted") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AddedOnlyInTheRevisedDocument_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Lead") + P("Alpha") + P("Bravo monthly") + P("Tail"));
        var right = Doc(P("Lead") + Start("Added") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_RemovedInTheRevisedDocument_AcceptDropsItAndRejectKeepsIt(string engine)
    {
        var left = Doc(P("Lead") + Start("Removed") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var right = Doc(P("Lead") + P("Alpha") + P("Bravo weekly") + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AddedWithNoOtherChange_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Lead") + P("Alpha") + P("Bravo") + P("Tail"));
        var right = Doc(P("Lead") + Start("Added") + P("Alpha") + P("Bravo") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_WhoseEndMovesToALaterBoundary_FollowsEachSide(string engine)
    {
        var left = Doc(Start("Range") + P("Alpha") + P("Bravo monthly") + End() + P("Charlie") + P("Tail"));
        var right = Doc(Start("Range") + P("Alpha") + P("Bravo weekly") + P("Charlie") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_WhoseStartMovesToAnEarlierBoundary_FollowsEachSide(string engine)
    {
        var left = Doc(P("Lead") + P("Alpha") + Start("Range") + P("Bravo monthly") + End() + P("Tail"));
        var right = Doc(P("Lead") + Start("Range") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_InsideAReplacedGap_FollowsEachSide(string engine)
    {
        // The original's bookmark opens before its second deleted paragraph and the revised document's
        // before its second inserted one: one-sided positions inside a gap of whole-block changes.
        var left = Doc(P("Lead") + P("Old one") + Start("Mid") + P("Old two") + End() + P("Tail"));
        var right = Doc(P("Lead") + P("New first") + Start("Mid", 9) + P("New second") + End(9) + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    // ---------------------------------------------------------- folding a replaced gap around a marker (#1183)

    // A heading whose bookmark starts inside it and ends between blocks, as Word writes a TOC heading bookmark,
    // followed by two more paragraphs, all of it replaced. The bookmark end sits in front of the second deleted
    // paragraph, which the paragraph-mark pairing never fuses: the inserted paragraph still folds into the
    // deleted heading, as Word's compare does.
    private const string HeadingWithBookmarkEndBetweenBlocks =
        "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"_Toc1\"/><w:r><w:t>Quarterly Planning Notes</w:t></w:r></w:p>" +
        "<w:bookmarkEnd w:id=\"0\"/>";

    private static readonly string ReplacedGapLeft =
        HeadingWithBookmarkEndBetweenBlocks + P("Budget review is scheduled for Monday.") + P("Hiring plan is attached.");

    private static string TwoByTwoTable()
    {
        static string Cell(string text) => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{P(text)}</w:tc>";
        return "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/>" +
            "<w:gridCol w:w=\"3000\"/></w:tblGrid>" + $"<w:tr>{Cell("R1C1")}{Cell("R1C2")}</w:tr><w:tr>{Cell("R2C1")}{Cell("R2C2")}</w:tr></w:tbl>";
    }

    private static int BodyParagraphs(WmlDocument document) => Body(document).Descendants(W + "p").Count();

    private static int BookmarkMarkers(WmlDocument document) =>
        Body(document).Descendants().Count(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd");

    /// <summary>The round trip, plus the bookmark checks <see cref="Layout"/> cannot see: accepting leaves no
    /// bookmark anywhere (the revised document has none, not even inside a coalesced paragraph), and rejecting
    /// restores the heading's start inside it.</summary>
    private static WmlDocument AssertReplacedGapRoundTrip(string engine, WmlDocument left, WmlDocument right)
    {
        var redline = AssertRoundTrip(engine, left, right);
        Assert.Equal(0, BookmarkMarkers(Accept(redline)));
        var rejected = Reject(redline);
        Assert.Equal(2, BookmarkMarkers(rejected));
        Assert.Equal("_Toc1", (string?)Body(rejected).Elements(W + "p").First()
            .Elements(W + "bookmarkStart").Single().Attribute(W + "name"));
        return redline;
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkEndBetweenItsDeletedBlocks_StillFoldsTheInsertedParagraph(string engine)
    {
        var left = Doc(ReplacedGapLeft);
        var right = Doc(P("Draft agenda for the offsite."));

        var redline = AssertReplacedGapRoundTrip(engine, left, right);

        // The inserted text shares the deleted heading's paragraph, and the bookmark end stays between blocks.
        Assert.Equal(3, BodyParagraphs(redline));
        var first = Body(redline).Elements(W + "p").First();
        Assert.Contains("Draft agenda for the offsite.", string.Concat(first.Descendants(W + "t").Select(t => t.Value)));
        Assert.Contains("Quarterly Planning Notes", string.Concat(first.Descendants(W + "delText").Select(t => t.Value)));
        Assert.Equal("p", Body(redline).Elements().Single(e => e.Descendants(W + "bookmarkEnd").Any() &&
            e.Name != W + "p").ElementsBeforeSelf().Last().Name.LocalName);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkEndBetweenItsDeletedBlocks_StillFoldsTheParagraphAfterAnInsertedTable(string engine)
    {
        var left = Doc(ReplacedGapLeft);
        var right = Doc(TwoByTwoTable() + "<w:p/>");

        AssertReplacedGapRoundTrip(engine, left, right);

        // Four cell paragraphs and the three deleted ones: the inserted empty paragraph after the table folds
        // into the last deleted paragraph rather than standing on its own.
        Assert.Equal(7, BodyParagraphs(Compare(engine, left, right)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedCellGap_WithABookmarkEndBetweenItsDeletedBlocks_StillFoldsTheInsertedParagraph(string engine)
    {
        var left = Doc(P("Lead") + Table(ReplacedGapLeft));
        var right = Doc(P("Lead") + Table(P("Draft agenda for the offsite.")));

        var redline = AssertRoundTrip(engine, left, right);

        var cell = Body(redline).Descendants(W + "tc").Single();
        Assert.Equal(new[] { "p", "del", "p", "p" },
            cell.Elements().Where(e => e.Name != W + "tcPr").Select(e => e.Name.LocalName));
        Assert.Equal(0, BookmarkMarkers(Accept(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkInFrontOfAnInsertedParagraph_StillFoldsTheFinalPair(string engine)
    {
        // The revised side's marker, in front of an inserted paragraph the pairing leaves alone.
        var left = Doc(P("Quarterly budget summary."));
        var right = Doc(P("Draft agenda.") + Start("Added") + P("Venue options.") + P("Travel logistics.") + End());

        var redline = AssertRoundTrip(engine, left, right);

        // The last inserted paragraph shares the deleted one's paragraph mark.
        Assert.Equal(3, BodyParagraphs(redline));
        Assert.Equal(new[] { "BS:Added@1", "BE:Added@3" }, BodyMarkerPositions(Accept(redline)));
        Assert.Equal(0, BookmarkMarkers(Reject(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkEndInFrontOfTheParagraphThatKeepsItsMark_StillFolds(string engine)
    {
        // Here the marker sits in front of the deleted paragraph whose mark the pairing keeps: the pair still forms,
        // and the marker goes in front of the shared paragraph that paragraph's runs end up in.
        var left = Doc(HeadingWithBookmarkEndBetweenBlocks + P("Hiring plan is attached."));
        var right = Doc(P("Draft agenda for the offsite."));

        var redline = AssertReplacedGapRoundTrip(engine, left, right);

        Assert.Equal(2, BodyParagraphs(redline));
        var children = Body(redline).Elements().Where(e => e.Name != W + "sectPr").ToList();
        Assert.Equal(new[] { "p", "del", "p" }, children.Select(e => e.Name.LocalName));
        Assert.Equal("bookmarkEnd", children[1].Elements().Single().Name.LocalName);
        Assert.Contains("Draft agenda for the offsite.", string.Concat(children[0].Descendants(W + "t").Select(t => t.Value)));
        Assert.Contains("Quarterly Planning Notes", string.Concat(children[0].Descendants(W + "delText").Select(t => t.Value)));
        Assert.Contains("Hiring plan is attached.", string.Concat(children[2].Descendants(W + "delText").Select(t => t.Value)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedInteriorGap_WithABookmarkInFrontOfAnInsertedBlank_PairsAsWithoutIt(string engine)
    {
        // Between unchanged paragraphs a wordful-then-blank original faces three blanks, the first of which holds a
        // heading bookmark that ends after it. The paragraph-mark chain decides exactly as it does without the
        // bookmark: the same paragraphs come out, and the bookmark keeps its place in each view.
        string Gap(bool bookmark) =>
            (bookmark ? "<w:p><w:bookmarkStart w:id=\"21\" w:name=\"_TocR\"/></w:p><w:bookmarkEnd w:id=\"21\"/>" : "<w:p/>") +
            "<w:p/><w:p/>";
        var left = Doc(P("Lead") + P("Original wordful paragraph.") + "<w:p/>" + P("Tail"));
        var right = Doc(P("Lead") + Gap(bookmark: true) + P("Tail"));
        var unmarked = Doc(P("Lead") + Gap(bookmark: false) + P("Tail"));

        var redline = AssertRoundTrip(engine, left, right);

        Assert.Equal(BodyParagraphs(Compare(engine, left, unmarked)), BodyParagraphs(redline));
    }

    // ------------------------------------------------------------------------------------- a table cell

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellBookmark_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo monthly") + End(), P("Other")));
        var right = Doc(P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo weekly") + End(), P("Other")));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellBookmark_StillBracketsItsParagraphsInTheRedline(string engine)
    {
        var left = Doc(P("Lead") + Table(P("Before") + Start("InCell") + P("Alpha") + P("Bravo monthly") + End() + P("After")));
        var right = Doc(P("Lead") + Table(P("Before") + Start("InCell") + P("Alpha") + P("Bravo weekly") + End() + P("After")));

        var cell = Body(Compare(engine, left, right)).Descendants(W + "tc").Single();
        Assert.Equal(new[] { "p", "bookmarkStart", "p", "p", "bookmarkEnd", "p" },
            cell.Elements().Where(e => e.Name != W + "tcPr").Select(e => e.Name.LocalName));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellBookmark_AddedOnlyInTheRevisedDocument_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Lead") + Table(P("Alpha") + P("Bravo monthly")));
        var right = Doc(P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo weekly") + End()));

        AssertRoundTrip(engine, left, right);
    }

    // A marker-only difference inside a container the IR finds equal: the IR leaves markers out (rule N3), so
    // nothing else tells the two sides apart and the container is emitted as one side's copy.

    private static string Row(params string[] cellContents) =>
        "<w:tr>" + string.Concat(cellContents.Select(c => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{c}</w:tc>")) + "</w:tr>";

    private static string Rows(int columns, params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(Enumerable.Range(0, columns).Select(_ => "<w:gridCol w:w=\"3000\"/>")) + "</w:tblGrid>" +
        string.Concat(rows) + "</w:tbl>";

    private static readonly string Bracketed = P("Alpha") + Start("Old") + P("Bravo") + End();
    private static readonly string Plain = P("Alpha") + P("Bravo");

    public static TheoryData<string, string, string> MarkerOnlyContainerChanges()
    {
        var data = new TheoryData<string, string, string>();
        foreach (var engine in new[] { "DocxCompare", "DocxDiff" })
            foreach (var (withMarker, without) in new[] { (Bracketed, Plain) })
            {
                // An entirely unchanged table; a body edit elsewhere.
                data.Add(engine, P("Lead one") + Table(withMarker), P("Lead two") + Table(without));
                data.Add(engine, P("Lead one") + Table(without), P("Lead two") + Table(withMarker));
                // Another cell of the same row changes.
                data.Add(engine, Rows(2, Row(withMarker, P("Other one"))), Rows(2, Row(without, P("Other two"))));
                data.Add(engine, Rows(2, Row(without, P("Other one"))), Rows(2, Row(withMarker, P("Other two"))));
                // Another row changes.
                data.Add(engine, Rows(1, Row(withMarker), Row(P("Other one"))), Rows(1, Row(without), Row(P("Other two"))));
                data.Add(engine, Rows(1, Row(without), Row(P("Other one"))), Rows(1, Row(withMarker), Row(P("Other two"))));
                // An unchanged block wrapper and an unchanged custom-XML wrapper.
                data.Add(engine, P("Lead one") + Sdt(withMarker), P("Lead two") + Sdt(without));
                data.Add(engine, P("Lead one") + Sdt(without), P("Lead two") + Sdt(withMarker));
                data.Add(engine, P("Lead one") + CustomXml(withMarker), P("Lead two") + CustomXml(without));
                // A table nested in an unchanged table.
                data.Add(engine, P("Lead one") + Table(P("Outer") + Table(withMarker) + P("After")),
                    P("Lead two") + Table(P("Outer") + Table(without) + P("After")));
            }
        return data;
    }

    [Theory]
    [MemberData(nameof(MarkerOnlyContainerChanges))]
    public void Compare_MarkerOnlyChange_InAnEqualContainer_FollowsEachSide(string engine, string leftBody, string rightBody)
    {
        AssertRoundTrip(engine, Doc(leftBody), Doc(rightBody));
    }

    // ------------------------------------------------------------------- accepting a merged paragraph mark

    private const string DeletedMark = "<w:pPr><w:rPr><w:del w:id=\"91\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>";

    [Fact]
    public void Accept_DeletedParagraphMark_MergesAcrossABookmarkStart_KeepsItBetweenTheTexts()
    {
        var document = Doc($"<w:p>{DeletedMark}<w:r><w:t>Alpha</w:t></w:r></w:p>" + Start("Mid") + P("Bravo") + End());

        var body = Body(Accept(document));

        Assert.Equal(new[] { "p", "bookmarkEnd" }, body.Elements().Where(e => e.Name != W + "sectPr").Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "r", "bookmarkStart", "r" }, body.Element(W + "p")!.Elements().Select(e => e.Name.LocalName));
    }

    [Fact]
    public void Accept_DeletedParagraphMark_MergesAcrossAnAltChunk_KeepsTheAltChunkOutOfTheParagraph()
    {
        var source = Doc($"<w:p>{DeletedMark}<w:r><w:t>Alpha</w:t></w:r></w:p>" +
            "<w:altChunk xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdX\"/>" + P("Bravo"));
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, true))
        {
            var chunk = package.MainDocumentPart!
                .AddNewPart<DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPart>("text/html", "rIdX");
            using var writer = new StreamWriter(chunk.GetStream(FileMode.Create));
            writer.Write("<html><body><p>Chunk</p></body></html>");
        }
        var document = new WmlDocument("altchunk.docx", stream.ToArray());

        var accepted = Accept(document);

        Assert.Empty(Body(accepted).Descendants(W + "p").Elements(W + "altChunk"));
        Assert.Single(Body(accepted).Elements(W + "altChunk"));
        AssertNoNewValidationErrors(accepted, document);
    }

    // ---------------------------------------------------------------------------------- comment ranges

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRange_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo monthly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");
        var right = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo weekly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellCommentRange_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = CommentDoc(P("Lead") + Table("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo monthly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}</w:p>"));
        var right = CommentDoc(P("Lead") + Table("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo weekly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}</w:p>"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRange_AddedOnlyInTheRevisedDocument_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Alpha") + P("Bravo monthly") + P("Charlie"));
        var right = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo weekly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRange_RemovedInTheRevisedDocument_AcceptDropsItAndRejectKeepsIt(string engine)
    {
        var left = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo monthly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");
        var right = Doc(P("Alpha") + P("Bravo weekly") + P("Charlie"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AroundAMovedParagraph_FollowsEachSide(string engine)
    {
        const string moved = "This whole paragraph moves further down the document unchanged";
        var left = Doc(P("Lead") + Start("Moved") + P(moved) + End() + P("Alpha paragraph stays") + P("Bravo paragraph stays") + P("Tail"));
        var right = Doc(P("Lead") + P("Alpha paragraph stays") + P("Bravo paragraph stays") + Start("Moved") + P(moved) + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    // ------------------------------------------------------------------------------------------ notes

    private const string Separators =
        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static WmlDocument FootnoteDoc(string noteContent) => IrTestDocuments.FromParts(
        "<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>",
        footnotesInnerXml: Separators + $"<w:footnote w:id=\"1\">{noteContent}</w:footnote>");

    /// <summary>The children of footnote 1: paragraphs by text, bookmark markers by kind and name.</summary>
    private static string FootnoteLayout(WmlDocument document)
    {
        var note = Part(document, "word/footnotes.xml").Root!.Elements(W + "footnote")
            .Single(n => (string?)n.Attribute(W + "id") == "1");
        var names = note.Descendants(W + "bookmarkStart").GroupBy(b => (string?)b.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(W + "name") ?? "");
        return string.Join(" ", note.Elements().Select(e =>
            e.Name == W + "p" ? "P:" + string.Concat(e.Descendants(W + "t").Select(t => t.Value))
            : e.Name == W + "bookmarkStart" ? "BS:" + (string?)e.Attribute(W + "name")
            : e.Name == W + "bookmarkEnd" ? "BE:" + names.GetValueOrDefault((string?)e.Attribute(W + "id") ?? "", "?")
            : e.Name.LocalName));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_FootnoteBookmark_AroundAnEditedParagraph_StaysAroundItsParagraphs(string engine)
    {
        var left = FootnoteDoc(P("Note lead") + Start("InNote") + P("Alpha") + P("Bravo monthly") + End());
        var right = FootnoteDoc(P("Note lead") + Start("InNote") + P("Alpha") + P("Bravo weekly") + End());

        var redline = Compare(engine, left, right);

        Assert.Equal("P:Note lead BS:InNote P:Alpha P:Bravo weekly BE:InNote", FootnoteLayout(Accept(redline)));
        Assert.Equal("P:Note lead BS:InNote P:Alpha P:Bravo monthly BE:InNote", FootnoteLayout(Reject(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_FootnoteBookmark_MarkerOnlyChange_FollowsEachSide(string engine)
    {
        var left = FootnoteDoc(P("Note lead") + Start("InNote") + P("Alpha") + P("Bravo") + End());
        var right = FootnoteDoc(P("Note lead") + P("Alpha") + P("Bravo"));

        var redline = Compare(engine, left, right);

        Assert.Equal(FootnoteLayout(right), FootnoteLayout(Accept(redline)));
        Assert.Equal(FootnoteLayout(left), FootnoteLayout(Reject(redline)));
    }

    // ------------------------------------------------------------------------------------ block wrappers

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BookmarkBetweenAReplacedWrappersBlocks_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + Sdt(P("Alpha") + Start("Inner") + P("Bravo monthly") + End()) + P("Tail"));
        var right = Doc(P("Lead") + Sdt(P("Alpha") + Start("Inner") + P("Bravo weekly") + End()) + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BookmarkBetweenAReplacedCustomXmlsBlocks_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + CustomXml(P("Alpha") + Start("Inner") + P("Bravo monthly") + End()) + P("Tail"));
        var right = Doc(P("Lead") + CustomXml(P("Alpha") + Start("Inner") + P("Bravo weekly") + End()) + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CommentRangeOpeningBeforeAReplacedWrapper_SurvivesAcceptAndReject(string engine)
    {
        var left = CommentDoc(P("Lead") + "<w:commentRangeStart w:id=\"0\"/>" + Sdt(P("Alpha") + P("Bravo monthly")) +
            $"<w:p><w:r><w:t>Charlie</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>{CommentReference}</w:p>");
        var right = CommentDoc(P("Lead") + "<w:commentRangeStart w:id=\"0\"/>" + Sdt(P("Alpha") + P("Bravo weekly")) +
            $"<w:p><w:r><w:t>Charlie</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>{CommentReference}</w:p>");

        AssertRoundTrip(engine, left, right);
    }

    // ------------------------------------------------------------------------------- headers and footers

    private static WmlDocument HeaderFooterDoc(string headerContent, string footerContent, string body = "",
        string? comments = null)
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string w = W.NamespaceName;
        using var stream = new MemoryStream();
        using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            main.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentFormat.OpenXml.Packaging.DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            static void Write(DocumentFormat.OpenXml.Packaging.OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            if (comments is not null)
                Write(main.AddNewPart<DocumentFormat.OpenXml.Packaging.WordprocessingCommentsPart>(),
                    $"<w:comments xmlns:w=\"{w}\">{comments}</w:comments>");
            Write(main.AddNewPart<DocumentFormat.OpenXml.Packaging.HeaderPart>("rIdHdr"), $"<w:hdr xmlns:w=\"{w}\">{headerContent}</w:hdr>");
            Write(main.AddNewPart<DocumentFormat.OpenXml.Packaging.FooterPart>("rIdFtr"), $"<w:ftr xmlns:w=\"{w}\">{footerContent}</w:ftr>");
            Write(main, $"<w:document xmlns:w=\"{w}\" xmlns:r=\"{R}\"><w:body>{(body.Length > 0 ? body : P("Body"))}" +
                "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHdr\"/><w:footerReference w:type=\"default\" r:id=\"rIdFtr\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("header-footer.docx", stream.ToArray());
    }

    /// <summary>The block layout (see <see cref="Layout"/>) of the default header (or footer) of section
    /// <paramref name="section"/> (the last one when null): the story its <c>w:sectPr</c> references, or the one
    /// it inherits from an earlier section.</summary>
    private static string StoryLayout(WmlDocument document, bool header, int? section = null) =>
        StreamLayout(document, StoryRootOf(document, header, section));

    /// <summary>The root of the default header (or footer) of section <paramref name="section"/> (see
    /// <see cref="StoryLayout"/>).</summary>
    private static XElement StoryRootOf(WmlDocument document, bool header, int? section = null)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        var sections = Body(document).Descendants(W + "sectPr").ToList();
        var reference = sections.Take((section ?? sections.Count - 1) + 1).Reverse()
            .Select(sectPr => sectPr.Elements(W + (header ? "headerReference" : "footerReference"))
                .FirstOrDefault(r => (string?)r.Attribute(W + "type") == "default"))
            .First(r => r is not null)!;
        var part = main.GetPartById((string)reference.Attribute(XNamespace.Get(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id")!);
        using var partStream = part.GetStream();
        return XDocument.Load(partStream).Root!;
    }

    /// <summary>A two-section document with the given header parts (no footer); section <c>i</c> references
    /// <c>headers[sectionHeaders[i]]</c> as its default header, or inherits the previous one when that is null.</summary>
    private static WmlDocument SectionedHeaderDoc(string[] headers, int?[] sectionHeaders)
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string w = W.NamespaceName;
        using var stream = new MemoryStream();
        using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            main.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentFormat.OpenXml.Packaging.DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            static void Write(DocumentFormat.OpenXml.Packaging.OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            for (int h = 0; h < headers.Length; h++)
                Write(main.AddNewPart<DocumentFormat.OpenXml.Packaging.HeaderPart>($"rIdHdr{h}"), $"<w:hdr xmlns:w=\"{w}\">{headers[h]}</w:hdr>");
            string SectPr(int? h) => "<w:sectPr>" + (h is { } i ? $"<w:headerReference w:type=\"default\" r:id=\"rIdHdr{i}\"/>" : "") +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>";
            var body = string.Concat(sectionHeaders.Take(sectionHeaders.Length - 1).Select((h, i) =>
                    $"<w:p><w:pPr>{SectPr(h)}</w:pPr><w:r><w:t>Section {i}</w:t></w:r></w:p>")) +
                P("Last section") + SectPr(sectionHeaders[^1]);
            Write(main, $"<w:document xmlns:w=\"{w}\" xmlns:r=\"{R}\"><w:body>{body}</w:body></w:document>");
        }
        return new WmlDocument("sectioned-header.docx", stream.ToArray());
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_MarkerOnlyChange_InOneOfTwoSectionsSharingAHeader_FollowsEachSection(string engine)
    {
        // Both sections share one header in the original; the revised document gives the second section its own
        // copy of it, which only adds a bookmark. The first section's story is unchanged.
        var left = SectionedHeaderDoc(new[] { StoryWithoutBookmark }, new int?[] { 0, null });
        var right = SectionedHeaderDoc(new[] { StoryWithoutBookmark, StoryWithBookmark }, new int?[] { 0, 1 });

        var redline = Compare(engine, left, right);

        var (accepted, rejected) = (Accept(redline), Reject(redline));
        for (int section = 0; section < 2; section++)
        {
            Assert.Equal(StoryLayout(right, header: true, section), StoryLayout(accepted, header: true, section));
            Assert.Equal(StoryLayout(left, header: true, section), StoryLayout(rejected, header: true, section));
        }
        Assert.NotEqual(StoryLayout(accepted, header: true, 0), StoryLayout(accepted, header: true, 1));
        AssertNoNewValidationErrors(redline, left, right);
    }

    /// <summary>Accepting the redline gives the revised header and footer layout, rejecting it the original's.</summary>
    private static void AssertStoryRoundTrip(string engine, WmlDocument left, WmlDocument right,
        bool rejectKeepsReferences = true)
    {
        var redline = Compare(engine, left, right);
        var (accepted, rejected) = (Accept(redline), Reject(redline));
        foreach (var header in new[] { true, false })
        {
            Assert.Equal(StoryLayout(right, header), StoryLayout(accepted, header));
            Assert.Equal(StoryLayout(left, header), StoryLayout(rejected, header));
            // Each view keeps its own side's comment references, and every comment marker resolves.
            Assert.Equal(CommentReferences(right, header), CommentReferences(accepted, header));
            if (rejectKeepsReferences)
                Assert.Equal(CommentReferences(left, header), CommentReferences(rejected, header));
            foreach (var document in new[] { redline, accepted, rejected })
                AssertStoryCommentMarkersResolve(document, header);
        }
        AssertNoNewValidationErrors(redline, WithoutStoryCommentId, left, right);
    }

    /// <summary>The text of each comment a story's comment references point at, in order.</summary>
    private static List<string> CommentReferences(WmlDocument document, bool header)
    {
        var texts = CommentTexts(document);
        return StoryRootOf(document, header).Descendants(W + "commentReference")
            .Select(r => texts.GetValueOrDefault((string?)r.Attribute(W + "id") ?? "", "?")).ToList();
    }

    private static Dictionary<string, string> CommentTexts(WmlDocument document) =>
        (Part(document, "word/comments.xml").Root?.Elements(W + "comment") ?? Enumerable.Empty<XElement>())
            .GroupBy(c => (string?)c.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => string.Concat(g.First().Descendants(W + "t").Select(t => t.Value)));

    private static void AssertStoryCommentMarkersResolve(WmlDocument document, bool header)
    {
        var defined = CommentTexts(document);
        Assert.All(StoryRootOf(document, header).Descendants()
                .Where(e => e.Name == W + "commentRangeStart" || e.Name == W + "commentRangeEnd" || e.Name == W + "commentReference"),
            m => Assert.Contains((string?)m.Attribute(W + "id") ?? "", defined.Keys));
    }

    private static readonly string StoryWithBookmark = P("Alpha") + Start("Story") + P("Bravo") + End();
    private static readonly string StoryWithoutBookmark = P("Alpha") + P("Bravo");

    public static TheoryData<string, string, string, string, string> MarkerOnlyStoryChanges()
    {
        var data = new TheoryData<string, string, string, string, string>();
        foreach (var engine in new[] { "DocxCompare", "DocxDiff" })
        {
            // Header added, header removed, footer added, footer removed; then with the body edited too.
            data.Add(engine, StoryWithoutBookmark, StoryWithBookmark, P("Footer"), P("Footer"));
            data.Add(engine, StoryWithBookmark, StoryWithoutBookmark, P("Footer"), P("Footer"));
            data.Add(engine, P("Header"), P("Header"), StoryWithoutBookmark, StoryWithBookmark);
            data.Add(engine, P("Header"), P("Header"), StoryWithBookmark, StoryWithoutBookmark);
            // A bookmark whose end moves to a later block boundary.
            data.Add(engine, Start("Story") + P("Alpha") + End() + P("Bravo"), Start("Story") + P("Alpha") + P("Bravo") + End(),
                P("Footer"), P("Footer"));
            // Inside a table cell of an otherwise unchanged header.
            data.Add(engine, Table(StoryWithoutBookmark), Table(StoryWithBookmark), P("Footer"), P("Footer"));
        }
        return data;
    }

    private static string HeaderComment(int id, string text) =>
        $"<w:comment w:id=\"{id}\" w:author=\"A\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>";

    private static string CommentRef(int id) =>
        $"<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"{id}\"/></w:r>";

    private static string CommentedStory(int id, bool endAfterBravo, string tail = "") =>
        $"<w:commentRangeStart w:id=\"{id}\"/>" + P("Alpha") + (endAfterBravo ? "" : $"<w:commentRangeEnd w:id=\"{id}\"/>") +
        $"<w:p><w:r><w:t>Bravo</w:t></w:r>{CommentRef(id)}</w:p>" + (endAfterBravo ? $"<w:commentRangeEnd w:id=\"{id}\"/>" : "") + tail;

    public static TheoryData<string, string, string, string?, string?> CommentRangeStoryChanges()
    {
        var data = new TheoryData<string, string, string, string?, string?>();
        var comment = HeaderComment(0, "Check");
        foreach (var engine in new[] { "DocxCompare", "DocxDiff" })
            foreach (var (leftTail, rightTail) in new[] { ("", ""), (P("Charlie one"), P("Charlie two")) })
            {
                // A comment range added around the header's paragraphs, then one removed: alone, then with a text
                // edit elsewhere in the header.
                data.Add(engine, P("Alpha") + P("Bravo") + leftTail, CommentedStory(0, endAfterBravo: false, rightTail), null, comment);
                data.Add(engine, CommentedStory(0, endAfterBravo: false, leftTail), P("Alpha") + P("Bravo") + rightTail, comment, null);
            }
        return data;
    }

    [Theory]
    [MemberData(nameof(CommentRangeStoryChanges))]
    public void Compare_CommentRangeChange_InAHeader_FollowsEachSide(
        string engine, string leftHeader, string rightHeader, string? leftComments, string? rightComments)
    {
        var left = HeaderFooterDoc(leftHeader, P("Footer"), comments: leftComments);
        var right = HeaderFooterDoc(rightHeader, P("Footer"), comments: rightComments);

        // A removed comment's reference sits inside a paragraph, which the header diff takes from the revised
        // document whether that paragraph is equal or edited (the reader keeps no comment reference in a header's
        // IR), so rejecting cannot restore it. When the header paragraph is edited, main loses it the same way; for
        // a removal alone, main kept the whole original header instead, so accepting kept the comment. Its range
        // between blocks follows each side.
        AssertStoryRoundTrip(engine, left, right, rejectKeepsReferences: leftComments is null);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_HeaderCommentRangeWhoseEndMoves_FollowsEachSide(string engine)
    {
        // The same comment with its end moved one paragraph on, alone and with a text edit elsewhere: each view
        // keeps its own range rather than one range serving both.
        var comments = HeaderComment(0, "Check");
        foreach (var (leftTail, rightTail) in new[] { ("", ""), (P("Charlie one"), P("Charlie two")) })
            AssertStoryRoundTrip(engine,
                HeaderFooterDoc(CommentedStory(0, endAfterBravo: false, leftTail), P("Footer"), comments: comments),
                HeaderFooterDoc(CommentedStory(0, endAfterBravo: true, rightTail), P("Footer"), comments: comments));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRangeWhoseEndMoves_FollowsEachSide(string engine)
    {
        // The body counterpart: a comment both documents have whose end moves to a later block boundary.
        var left = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + "<w:commentRangeEnd w:id=\"0\"/>" +
            $"<w:p><w:r><w:t>Bravo</w:t></w:r>{CommentReference}</w:p>");
        var right = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") +
            $"<w:p><w:r><w:t>Bravo</w:t></w:r>{CommentReference}</w:p>" + "<w:commentRangeEnd w:id=\"0\"/>");

        AssertRoundTrip(engine, left, right);
    }

    private static WmlDocument BaseCommentDoc(string body) =>
        IrTestDocuments.WithComment("Base", "B", "2026-01-01T00:00:00Z", "Base comment", body);

    private const string CommentedTail =
        "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Tail</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>" + CommentReference + "</w:p>";

    private const string ReferencedTail = "<w:p><w:r><w:t>Tail</w:t></w:r>" + CommentReference + "</w:p>";

    public static TheoryData<string, string, string> MovedCommentRanges()
    {
        var data = new TheoryData<string, string, string>();
        foreach (var engine in new[] { "DocxCompare", "DocxDiff" })
        {
            // An existing comment's range moves from inside its paragraph to bracket a cell's paragraphs, then two
            // body paragraphs; a body edit elsewhere keeps it a real comparison.
            data.Add(engine, P("Lead") + Table(P("Alpha") + P("Bravo")) + CommentedTail,
                P("Lead one") + Table("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo") + "<w:commentRangeEnd w:id=\"0\"/>") + ReferencedTail);
            data.Add(engine, P("Lead") + P("Alpha") + P("Bravo") + CommentedTail,
                P("Lead one") + "<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo") + "<w:commentRangeEnd w:id=\"0\"/>" + ReferencedTail);
        }
        return data;
    }

    [Theory]
    [MemberData(nameof(MovedCommentRanges))]
    public void Compare_ExistingCommentRangeMovedBetweenBlocks_FollowsEachSide(string engine, string originalBody, string revisedBody)
    {
        AssertRoundTrip(engine, BaseCommentDoc(originalBody), BaseCommentDoc(revisedBody));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_HeaderCommentWhoseIdAloneChanges_IsUnchanged(string engine)
    {
        // The revised document numbers the same comment 5: that is no change to the header, which keeps the
        // original's story (and its own reference) in both views.
        var left = HeaderFooterDoc(CommentedStory(0, endAfterBravo: false), P("Footer"), comments: HeaderComment(0, "Check"));
        var right = HeaderFooterDoc(CommentedStory(5, endAfterBravo: false), P("Footer"), comments: HeaderComment(5, "Check"));

        Assert.Null(EditScript(left, right).HeaderFooterOps);
        AssertStoryRoundTrip(engine, left, right);
    }

    private static string CommentedAlpha(bool startInside) =>
        (startInside ? "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Alpha</w:t></w:r></w:p>"
            : "<w:commentRangeStart w:id=\"0\"/>" + P("Alpha")) +
        "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p><w:r><w:t>Bravo</w:t></w:r>{CommentReference}</w:p>";

    /// <summary>The body's comment markers and text in document order, each marker by kind and the text of the
    /// comment it points at.</summary>
    private static string CommentView(WmlDocument document)
    {
        var texts = CommentTexts(document);
        return string.Join(" ", Body(document).Descendants()
            .Where(e => e.Name == W + "t" || e.Name.LocalName is "commentRangeStart" or "commentRangeEnd" or "commentReference")
            .Select(e => e.Name == W + "t" ? e.Value : e.Name.LocalName + ":" + texts.GetValueOrDefault((string?)e.Attribute(W + "id") ?? "", "?")));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CommentStartMovedFromInsideAParagraphToBetweenBlocks_KeepsARangeInEachView(string engine)
    {
        // The original's start inside the paragraph is part of a paragraph the revised copy stands in for, so only
        // the end is left of the original's range between blocks; rather than lose it, both views share one range.
        var left = CommentDoc(CommentedAlpha(startInside: true));
        var right = CommentDoc(CommentedAlpha(startInside: false));

        var redline = Compare(engine, left, right);

        Assert.Equal(Layout(right), Layout(Accept(redline)));
        Assert.Equal("commentRangeStart:Check this Alpha commentRangeEnd:Check this Bravo commentReference:Check this",
            CommentView(Reject(redline)));
        AssertNoNewValidationErrors(redline, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CommentStartMovedFromBetweenBlocksToInsideAParagraph_FollowsEachSide(string engine)
    {
        var left = CommentDoc(CommentedAlpha(startInside: false));
        var right = CommentDoc(CommentedAlpha(startInside: true));

        var redline = AssertRoundTrip(engine, left, right);

        Assert.Equal(CommentView(right), CommentView(Accept(redline)));
        Assert.Equal(CommentView(left), CommentView(Reject(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_HeaderCommentRenumberedInTheRevisedDocument_ResolvesInEachView(string engine)
    {
        // The revised document added a body comment first, so the header's (unchanged) comment is its second.
        var left = HeaderFooterDoc(CommentedStory(0, endAfterBravo: false), P("Footer"), comments: HeaderComment(0, "Check"));
        var right = HeaderFooterDoc(CommentedStory(1, endAfterBravo: false), P("Footer"),
            $"<w:commentRangeStart w:id=\"0\"/><w:p><w:r><w:t>Body</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>{CommentRef(0)}</w:p>",
            HeaderComment(0, "New") + HeaderComment(1, "Check"));

        // The same comment is no change to the header, so each view keeps its reference to it.
        AssertStoryRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(MarkerOnlyStoryChanges))]
    public void Compare_MarkerOnlyChange_InAHeaderOrFooter_FollowsEachSide(
        string engine, string leftHeader, string rightHeader, string leftFooter, string rightFooter)
    {
        AssertStoryRoundTrip(engine, HeaderFooterDoc(leftHeader, leftFooter), HeaderFooterDoc(rightHeader, rightFooter));
        AssertStoryRoundTrip(engine, HeaderFooterDoc(leftHeader, leftFooter, P("Body one")),
            HeaderFooterDoc(rightHeader, rightFooter, P("Body two")));
    }

    private static Docxodus.Ir.Diff.IrEditScript EditScript(WmlDocument left, WmlDocument right) =>
        Docxodus.Ir.Diff.IrEditScriptBuilder.Build(Docxodus.Ir.IrReader.Read(left), Docxodus.Ir.IrReader.Read(right), new());

    [Fact]
    public void EditScript_HeaderWithTheSameBlockLevelBookmarkOnBothSides_HasNoStoryDiff()
    {
        // A story both documents bracket the same way is unchanged: the comparison leaves the original's part alone.
        var script = EditScript(HeaderFooterDoc(StoryWithBookmark, P("Footer"), P("Body one")),
            HeaderFooterDoc(StoryWithBookmark, P("Footer"), P("Body two")));

        Assert.Null(script.HeaderFooterOps);
    }

    [Fact]
    public void EditScript_HeaderWhoseOnlyChangeIsABlockLevelBookmark_KeepsItsEqualOps()
    {
        var script = EditScript(HeaderFooterDoc(StoryWithoutBookmark, P("Footer")), HeaderFooterDoc(StoryWithBookmark, P("Footer")));

        var story = Assert.Single(script.HeaderFooterOps!);
        Assert.True(story.IsHeader);
        Assert.Equal(2, story.Ops.Count);
        Assert.All(story.Ops, op => Assert.Equal(Docxodus.Ir.Diff.IrEditOpKind.EqualBlock, op.Kind));
    }

    [Fact]
    public void SemanticDiff_HeaderWhoseOnlyChangeIsABlockLevelBookmark_ReportsNoHeaderChange()
    {
        // Like the body's, a marker between blocks is not a semantic change: the story's projected value is equal.
        var changes = Docxodus.Verification.SemanticDiff.Compare(
            HeaderFooterDoc(StoryWithoutBookmark, P("Footer")), HeaderFooterDoc(StoryWithBookmark, P("Footer"))).Changes;

        Assert.DoesNotContain(changes, c => c.Family == Docxodus.Verification.SemanticChangeFamily.Header);
    }

    private static XElement Story(string content) =>
        XDocument.Parse($"<w:hdr xmlns:w=\"{W.NamespaceName}\">{content}</w:hdr>").Root!;

    public static TheoryData<string, string, bool> StoryMarkerPairs() => new()
    {
        { StoryWithBookmark, StoryWithBookmark, false },
        // The same bookmark under another id: a bookmark is matched by name.
        { StoryWithBookmark, P("Alpha") + Start("Story", 3) + P("Bravo") + End(3), false },
        // A marker inside a paragraph is not between blocks, so it is not compared here. (The reader leaves
        // in-paragraph markers out of the IR too, rule N3, so a story differing only in one reads as unchanged.)
        { $"<w:p>{Start("Inline")}<w:r><w:t>Alpha</w:t></w:r>{End()}</w:p>", P("Alpha"), false },
        { StoryWithoutBookmark, StoryWithBookmark, true },
        { Start("Story") + P("Alpha") + End() + P("Bravo"), Start("Story") + P("Alpha") + P("Bravo") + End(), true },
        { Table(StoryWithoutBookmark), Table(StoryWithBookmark), true },
        { Sdt(StoryWithBookmark), Sdt(StoryWithoutBookmark), true },
        // A cell's marker is not the story's: the same marker one level out is a difference.
        { Start("Story") + Table(P("Alpha")) + End(), Table(Start("Story") + P("Alpha") + End()), true },
    };

    [Theory]
    [MemberData(nameof(StoryMarkerPairs))]
    public void BlockLevelMarkersDiffer_ComparesEachBlockStreamsMarkersByBoundary(string left, string right, bool differ)
    {
        Assert.Equal(differ, Docxodus.Ir.Diff.IrMarkupRenderer.BlockLevelMarkersDiffer(Story(left), Story(right)));
        Assert.Equal(differ, Docxodus.Ir.Diff.IrMarkupRenderer.BlockLevelMarkersDiffer(Story(right), Story(left)));
    }

    // ---------------------------------------------------------------------------------- a real document

    /// <summary>Each body-level bookmark marker as kind, bookmark name and the number of body blocks before it.</summary>
    private static List<string> BodyMarkerPositions(WmlDocument document)
    {
        var body = Body(document);
        var names = body.Descendants(W + "bookmarkStart").GroupBy(s => (string?)s.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(W + "name") ?? "");
        var positions = new List<string>();
        int blocks = 0;
        foreach (var child in body.Elements())
        {
            if (child.Name == W + "p" || child.Name == W + "tbl" || child.Name == W + "sdt")
                blocks++;
            else if (child.Name == W + "bookmarkStart")
                positions.Add($"BS:{(string?)child.Attribute(W + "name")}@{blocks}");
            else if (child.Name == W + "bookmarkEnd")
                positions.Add($"BE:{names.GetValueOrDefault((string?)child.Attribute(W + "id") ?? "", "?")}@{blocks}");
        }
        return positions;
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_RealDocumentWithBodyLevelBookmarks_KeepsThemAtTheirBlockBoundaries(string engine)
    {
        // A Word document whose revised version holds 916 body-level bookmark markers (the WmlComparer-era
        // fixture rule N3 was first written for).
        var dir = Path.Combine("..", "..", "..", "..", "TestFiles", "WC");
        var left = new WmlDocument(Path.Combine(dir, "WC-BodyBookmarks-Before.docx"));
        var right = new WmlDocument(Path.Combine(dir, "WC-BodyBookmarks-After.docx"));

        var redline = Compare(engine, left, right);

        Assert.Equal(916, BodyMarkerPositions(right).Count);
        Assert.Equal(BodyMarkerPositions(right), BodyMarkerPositions(Accept(redline)));
        // Rejecting puts no marker anywhere the original does not have one. (Three of the original's six are ends
        // of bookmarks that start inside a paragraph the revised document restyled; that paragraph is emitted
        // from the revised side, which lacks the start, so those bookmarks are not restored on reject. That is
        // the inline-bookmark limitation, not the block-level placement.)
        var rejected = BodyMarkerPositions(Reject(redline));
        var original = BodyMarkerPositions(left);
        Assert.NotEmpty(rejected);
        Assert.All(rejected, marker => Assert.Contains(marker, original));
    }

    // -------------------------------------------------------------------------------------- consolidate

    [Fact]
    public void Consolidate_BodyBookmark_AroundAReviewersEditedParagraph_SurvivesAcceptAndReject()
    {
        var original = Doc(Start("Article") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var reviewer = Doc(Start("Article") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        var consolidated = DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = reviewer } });

        Assert.Equal(Layout(reviewer), Layout(Accept(consolidated)));
        Assert.Equal(Layout(original), Layout(Reject(consolidated)));
    }

    [Fact]
    public void Consolidate_BodyBookmark_AddedByOneOfTwoReviewers_AcceptKeepsItAndRejectDropsIt()
    {
        var original = Doc(P("Lead") + P("Alpha") + P("Bravo monthly") + P("Tail"));
        var first = Doc(P("Lead") + Start("Added") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var second = Doc(P("Lead") + P("Alpha") + P("Bravo weekly") + P("Tail"));

        var consolidated = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "First", Document = first },
            new DocxDiffReviewer { Author = "Second", Document = second },
        });

        Assert.Equal("P:Lead BS:Added P:Alpha P:Bravo weekly BE:Added P:Tail", Layout(Accept(consolidated)));
        Assert.Equal(Layout(original), Layout(Reject(consolidated)));
    }

    [Fact]
    public void Consolidate_BodyBookmark_RemovedByOneReviewer_AcceptDropsItAndRejectKeepsIt()
    {
        var original = Doc(P("Lead") + Start("Removed") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var first = Doc(P("Lead") + P("Alpha") + P("Bravo monthly") + P("Tail"));
        var second = Doc(P("Lead") + Start("Removed") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        var consolidated = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "First", Document = first },
            new DocxDiffReviewer { Author = "Second", Document = second },
        });

        Assert.Equal("P:Lead P:Alpha P:Bravo weekly P:Tail", Layout(Accept(consolidated)));
        Assert.Equal(Layout(original), Layout(Reject(consolidated)));
        Assert.Equal("First", (string?)Body(consolidated).Elements(W + "del")
            .Single(d => d.Element(W + "bookmarkStart") is not null).Attribute(W + "author"));
    }
}
