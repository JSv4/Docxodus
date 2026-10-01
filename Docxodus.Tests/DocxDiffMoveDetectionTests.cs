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
}
