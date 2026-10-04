// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// Which relocations the comparison reports as moves (<c>w:moveFrom</c>/<c>w:moveTo</c>) rather than an
/// unrelated delete + insert (issue #844). A relocated paragraph is a move whatever its length or distance,
/// including when it was also lightly edited or re-formatted; a relocated TABLE is one too, drawn the way
/// Word's compare draws it (rows deleted at the source and inserted at the destination, cell content
/// carrying the move). Every case keeps accept ≡ revised and reject ≡ original.
/// </summary>
public class DocxDiffMoveDetectionTests
{
    private const string A = "Alpha paragraph sets out the parties and the effective date of this agreement.";
    private const string B = "Bravo paragraph describes the services to be performed by the contractor.";
    private const string C = "Charlie paragraph covers payment terms, invoices and late fees in detail.";
    private const string D = "Delta paragraph addresses confidentiality obligations of both parties.";
    private const string E = "Echo paragraph limits liability and excludes consequential damages.";
    private const string F = "Foxtrot paragraph governs termination for convenience and for cause.";

    private static string P(string text, string pPr = "") =>
        $"<w:p>{pPr}<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Table(params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        string.Concat(rows.Select(row => $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>{P(row)}</w:tc></w:tr>")) +
        "</w:tbl>";

    private static WmlDocument Doc(string body) => IrTestDocuments.FromParts(body);

    private static WmlDocument Paragraphs(params string[] texts) => Doc(string.Concat(texts.Select(t => P(t))));

    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>Every paragraph's visible text in document order, table cells included.</summary>
    private static string Text(byte[] bytes) => string.Join("|", Body(bytes).Descendants(W.p)
        .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));

    private static void AssertRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline)
    {
        Assert.Equal(Text(revised.DocumentByteArray), Text(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray)));
        Assert.Equal(Text(original.DocumentByteArray), Text(DocxDiffOps.RejectRevisions(redline.DocumentByteArray)));
    }

    private static void AssertMoved(WmlDocument redline, string text, string? arrivedAs = null)
    {
        var body = Body(redline.DocumentByteArray);
        Assert.Contains(body.Descendants(W.moveFrom), m => string.Concat(m.Descendants(W.delText).Select(t => t.Value)) == text);
        Assert.Contains(body.Descendants(W.moveTo), m => string.Concat(m.Descendants(W.t).Select(t => t.Value)) == (arrivedAs ?? text));
    }

    // ---- paragraphs ------------------------------------------------------------------------------

    public static TheoryData<string, string[], string[], string> RelocatedParagraphs => new()
    {
        { "first to last", new[] { A, B, C, D, E, F }, new[] { B, C, D, E, F, A }, A },
        { "adjacent swap", new[] { A, B, C, D, E, F }, new[] { A, C, B, D, E, F }, B },
        { "one of a moved block", new[] { A, B, C, D, E, F }, new[] { A, D, E, F, B, C }, C },
        { "one word long", new[] { A, "Notice.", C, D, E, F }, new[] { A, C, D, E, F, "Notice." }, "Notice." },
        { "next to an insertion", new[] { A, B, C, D, E, F }, new[] { B, C, "A brand new paragraph appears here.", D, E, F, A }, A },
        { "in a two-paragraph document", new[] { A, B }, new[] { B, A }, A },
    };

    [Theory]
    [MemberData(nameof(RelocatedParagraphs))]
    public void RelocatedParagraph_IsAMove(string shape, string[] original, string[] revised, string moved)
    {
        _ = shape;
        var (left, right) = (Paragraphs(original), Paragraphs(revised));

        var redline = DocxCompare.Compare(left, right);

        AssertMoved(redline, moved);
        AssertRoundTrip(left, right, redline);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void RelocatedAndLightlyEditedParagraph_IsAMove(int editedWords)
    {
        var edited = editedWords == 1
            ? A.Replace("parties", "signatories")
            : A.Replace("parties", "signatories").Replace("effective", "commencement");
        var (left, right) = (Paragraphs(A, B, C, D, E, F), Paragraphs(B, C, D, E, F, edited));

        var redline = DocxCompare.Compare(left, right);

        // A moved-and-edited paragraph moves as complete halves: the original text leaves, the edited arrives.
        AssertMoved(redline, A, arrivedAs: edited);
        AssertRoundTrip(left, right, redline);
    }

    [Fact]
    public void RelocatedAndReformattedParagraph_IsAMove()
    {
        var left = Doc(P(A) + P(B) + P(C) + P(D));
        var right = Doc(P(B) + P(C) + P(D) + P(A, "<w:pPr><w:jc w:val=\"center\"/></w:pPr>"));

        var redline = DocxCompare.Compare(left, right);

        AssertMoved(redline, A);
        AssertRoundTrip(left, right, redline);
    }

    // ---- tables ----------------------------------------------------------------------------------

    [Fact]
    public void RelocatedTable_IsAMove_DrawnTheWayWordDrawsIt()
    {
        var left = Doc(Table(A, B) + P(C) + P(D));
        var right = Doc(P(C) + P(D) + Table(A, B));

        var redline = DocxCompare.Compare(left, right);

        var tables = Body(redline.DocumentByteArray).Elements(W.tbl).ToList();
        Assert.Equal(2, tables.Count);
        var (source, destination) = (tables[0], tables[1]);
        Assert.All(source.Elements(W.tr), row => Assert.NotNull(row.Element(W.trPr)?.Element(W.del)));
        Assert.All(destination.Elements(W.tr), row => Assert.NotNull(row.Element(W.trPr)?.Element(W.ins)));
        Assert.Empty(source.Descendants(W.t));
        Assert.Equal(new[] { A, B }, source.Descendants(W.moveFrom).Select(m => string.Concat(m.Descendants(W.delText).Select(t => t.Value))));
        Assert.Equal(new[] { A, B }, destination.Descendants(W.moveTo).Select(m => string.Concat(m.Descendants(W.t).Select(t => t.Value))));
        // One move range per half, inside the table, both halves under the same name.
        var from = Assert.Single(source.Elements(W.moveFromRangeStart));
        var to = Assert.Single(destination.Elements(W.moveToRangeStart));
        Assert.Equal((string?)from.Attribute(W.name), (string?)to.Attribute(W.name));
        Assert.Equal((string?)from.Attribute(W.id), (string?)source.Elements(W.moveFromRangeEnd).Single().Attribute(W.id));
        Assert.Equal((string?)to.Attribute(W.id), (string?)destination.Elements(W.moveToRangeEnd).Single().Attribute(W.id));

        AssertRoundTrip(left, right, redline);
        Assert.Single(Body(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray)).Elements(W.tbl));
        Assert.Single(Body(DocxDiffOps.RejectRevisions(redline.DocumentByteArray)).Elements(W.tbl));
        NoNewValidationErrors(left.DocumentByteArray, redline.DocumentByteArray);
    }

    [Fact]
    public void RelocatedTable_AgreesWithTheRevisionList()
    {
        var left = Doc(Table(A, B) + P(C) + P(D));
        var right = Doc(P(C) + P(D) + Table(A, B));

        var revisions = DocxDiff.GetRevisions(left, right);

        Assert.Equal(2, revisions.Count);
        Assert.All(revisions, r => Assert.Equal(DocxDiffRevisionType.Moved, r.Type));
        Assert.Single(revisions.Select(r => r.MoveGroupId).Distinct());
    }

    [Fact]
    public void RelocatedTable_WithMovesNotReported_IsADeleteAndAnInsert()
    {
        var left = Doc(Table(A, B) + P(C) + P(D));
        var right = Doc(P(C) + P(D) + Table(A, B));

        var redline = DocxCompare.Compare(left, right, new DocxDiffSettings { DetectMoves = false });

        var body = Body(redline.DocumentByteArray);
        Assert.Empty(body.Descendants(W.moveFrom));
        Assert.Empty(body.Descendants(W.moveTo));
        AssertRoundTrip(left, right, redline);
    }

    [Fact]
    public void RelocatedTable_InConsolidate_RoundTrips()
    {
        var original = Doc(Table(A, B) + P(C) + P(D));
        var reviewer = new DocxDiffReviewer { Author = "Reviewer", Document = Doc(P(C) + P(D) + Table(A, B)) };

        var consolidated = DocxDiff.Consolidate(original, new[] { reviewer });

        Assert.Equal(Text(reviewer.Document.DocumentByteArray), Text(DocxDiffOps.AcceptRevisions(consolidated.DocumentByteArray)));
        Assert.Equal(Text(original.DocumentByteArray), Text(DocxDiffOps.RejectRevisions(consolidated.DocumentByteArray)));
        NoNewValidationErrors(original.DocumentByteArray, consolidated.DocumentByteArray);
    }

    // ---- moves in different scopes (issue #924) ----------------------------------------------------

    /// <summary>A one-cell table whose cell holds the given paragraphs.</summary>
    private static string OneCellTable(params string[] paragraphs) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>{string.Concat(paragraphs.Select(t => P(t)))}</w:tc></w:tr></w:tbl>";

    /// <summary>A move in the body and an unrelated move inside a table cell, each the first of its scope.</summary>
    private static (WmlDocument Left, WmlDocument Right) BodyAndCellMoves(string cellMoverArrivesAs = E) =>
        (Doc(P(A) + P(B) + P(C) + OneCellTable(D, E, F)),
         Doc(P(B) + P(C) + P(A) + OneCellTable(D, F, cellMoverArrivesAs)));

    [Fact]
    public void MovesInDifferentScopes_HaveDistinctMoveNames()
    {
        var (left, right) = BodyAndCellMoves();

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline.DocumentByteArray);
        AssertMoved(redline, A);
        AssertMoved(redline, E);
        string NameOf(XName rangeStart, bool inTable) => (string)body.Descendants(rangeStart)
            .Single(e => e.Ancestors(W.tbl).Any() == inTable).Attribute(W.name)!;
        Assert.Equal(NameOf(W.moveFromRangeStart, inTable: false), NameOf(W.moveToRangeStart, inTable: false));
        Assert.Equal(NameOf(W.moveFromRangeStart, inTable: true), NameOf(W.moveToRangeStart, inTable: true));
        Assert.NotEqual(NameOf(W.moveFromRangeStart, inTable: false), NameOf(W.moveFromRangeStart, inTable: true));
        AssertRoundTrip(left, right, redline);
    }

    [Fact]
    public void MovesInDifferentScopes_HaveDistinctMoveGroupIds()
    {
        var (left, right) = BodyAndCellMoves();

        var moved = DocxDiff.GetRevisions(left, right).Where(r => r.Type == DocxDiffRevisionType.Moved).ToList();

        Assert.Equal(4, moved.Count);
        Assert.Equal(2, moved.Select(r => r.MoveGroupId).Distinct().Count());
        Assert.All(moved.GroupBy(r => r.MoveGroupId), group =>
            Assert.Single(group.Select(r => r.Text.Trim()).Distinct()));
    }

    [Fact]
    public void MovedAndEditedCellParagraph_ReportsItsOwnDeletedWord()
    {
        // The cell paragraph moves and one word changes. Its deleted word must be read from the cell's own
        // source paragraph, not from the body paragraph that moved under the same scope-local group id.
        var (left, right) = BodyAndCellMoves(cellMoverArrivesAs: E.Replace("liability", "exposure"));

        var revisions = DocxDiff.GetRevisions(left, right);

        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted && r.Text.Trim() == "liability");
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Trim() == "exposure");
    }

    // ---- table rows and cells (issue #887) ---------------------------------------------------------

    [Fact]
    public void ReorderedTableRow_IsAMove_DrawnTheWayWordDrawsIt()
    {
        var left = Doc(Table(A, B, C) + P(D));
        var right = Doc(Table(B, C, A) + P(D));

        var redline = DocxCompare.Compare(left, right);

        // One table: the row leaves its old position (row deleted, content moved from) and arrives at
        // its new one (row inserted, content moved to); the other rows are untouched.
        var table = Assert.Single(Body(redline.DocumentByteArray).Elements(W.tbl));
        var rows = table.Elements(W.tr).ToList();
        Assert.Equal(4, rows.Count);
        var source = Assert.Single(rows, r => r.Element(W.trPr)?.Element(W.del) != null);
        var destination = Assert.Single(rows, r => r.Element(W.trPr)?.Element(W.ins) != null);
        Assert.True(rows.IndexOf(source) < rows.IndexOf(destination));
        Assert.Empty(source.Descendants(W.t));
        Assert.Equal(new[] { A }, source.Descendants(W.moveFrom).Select(m => string.Concat(m.Descendants(W.delText).Select(t => t.Value))));
        Assert.Equal(new[] { A }, destination.Descendants(W.moveTo).Select(m => string.Concat(m.Descendants(W.t).Select(t => t.Value))));
        // Each half sits in its own move range, among the table's rows, and both share one name.
        var from = Assert.Single(table.Elements(W.moveFromRangeStart));
        var to = Assert.Single(table.Elements(W.moveToRangeStart));
        Assert.Same(source, from.ElementsAfterSelf().First());
        Assert.Same(destination, to.ElementsAfterSelf().First());
        Assert.Equal((string?)from.Attribute(W.id), (string?)source.ElementsAfterSelf().First().Attribute(W.id));
        Assert.Equal(W.moveFromRangeEnd, source.ElementsAfterSelf().First().Name);
        Assert.Equal(W.moveToRangeEnd, destination.ElementsAfterSelf().First().Name);
        Assert.Equal((string?)from.Attribute(W.name), (string?)to.Attribute(W.name));

        AssertRoundTrip(left, right, redline);
        NoNewValidationErrors(left.DocumentByteArray, redline.DocumentByteArray);
    }

    [Fact]
    public void ReorderedTableRow_AgreesWithTheRevisionList()
    {
        var left = Doc(Table(A, B, C) + P(D));
        var right = Doc(Table(B, C, A) + P(D));

        var redline = DocxCompare.Compare(left, right);
        var revisions = DocxDiff.GetRevisions(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(null));

        var moved = revisions.Where(r => r.Type == DocxDiffRevisionType.Moved).ToList();
        Assert.Equal(2, moved.Count);
        Assert.Single(moved.Select(r => r.MoveGroupId).Distinct());
        Assert.All(moved, r => Assert.Equal(A, r.Text.Trim()));
        Assert.Single(Body(redline.DocumentByteArray).Descendants(W.moveFromRangeStart));
    }

    [Fact]
    public void ReorderedTableRow_WithMovesNotReported_IsADeleteAndAnInsert()
    {
        var left = Doc(Table(A, B, C) + P(D));
        var right = Doc(Table(B, C, A) + P(D));

        var redline = DocxCompare.Compare(left, right, new DocxDiffSettings { DetectMoves = false });

        var body = Body(redline.DocumentByteArray);
        Assert.Empty(body.Descendants(W.moveFrom));
        Assert.Empty(body.Descendants(W.moveTo));
        Assert.Empty(body.Descendants(W.moveFromRangeStart));
        AssertRoundTrip(left, right, redline);
    }

    [Fact]
    public void RelocatedAndEditedTable_IsAMove_WithCompleteHalves()
    {
        // The table moves to the end and one cell gains a word. Like a moved-and-edited paragraph, the old
        // table leaves whole and the edited table arrives whole (nested revisions inside a move range are
        // not interoperable).
        var edited = B.Replace("services", "consulting services");
        var left = Doc(Table(A, B) + P(C) + P(D) + P(E));
        var right = Doc(P(C) + P(D) + P(E) + Table(A, edited));

        var redline = DocxCompare.Compare(left, right);

        var tables = Body(redline.DocumentByteArray).Elements(W.tbl).ToList();
        Assert.Equal(2, tables.Count);
        Assert.Equal(new[] { A, B }, tables[0].Descendants(W.moveFrom).Select(m => string.Concat(m.Descendants(W.delText).Select(t => t.Value))));
        Assert.Equal(new[] { A, edited }, tables[1].Descendants(W.moveTo).Select(m => string.Concat(m.Descendants(W.t).Select(t => t.Value))));
        Assert.Equal((string?)tables[0].Element(W.moveFromRangeStart)?.Attribute(W.name),
            (string?)tables[1].Element(W.moveToRangeStart)?.Attribute(W.name));
        AssertRoundTrip(left, right, redline);
        NoNewValidationErrors(left.DocumentByteArray, redline.DocumentByteArray);

        var moved = DocxDiff.GetRevisions(left, right, DocxCompare.ApplyFrontDoorRevisionPolicy(null))
            .Where(r => r.Type == DocxDiffRevisionType.Moved).ToList();
        Assert.Equal(2, moved.Count);
        Assert.Single(moved.Select(r => r.MoveGroupId).Distinct());
    }

    [Fact]
    public void ReplacedTable_IsNotAMove()
    {
        // A table in the same place whose text is mostly rewritten stays a table edit, and a table elsewhere
        // sharing little text with a removed one is not taken for it.
        var left = Doc(Table(A, B) + P(C) + P(D));
        var right = Doc(P(C) + P(D) + Table(E, F));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline.DocumentByteArray);
        Assert.Empty(body.Descendants(W.moveFrom));
        Assert.Empty(body.Descendants(W.moveTo));
        AssertRoundTrip(left, right, redline);
    }
}
