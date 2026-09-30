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
/// Two revision constructs the comparison writes that Word's own compare never does (issue #842):
/// <c>w:numberingChange</c>, which records the label a list item showed before an edit elsewhere in
/// the list renumbered it, and <c>w:cellIns</c>/<c>w:cellDel</c>, which track a single cell added to
/// or removed from a paired row. Both stay on by default; <see cref="DocxDiffSettings.TrackNumberingChanges"/>
/// and <see cref="DocxDiffSettings.TrackCellInsertionsAndDeletions"/> turn each off for output limited to
/// the constructs Word writes, without breaking the accept ≡ revised / reject ≡ original round trip.
/// </summary>
public class DocxDiffOptionalRevisionConstructsTests
{
    private const string DecimalLevel =
        "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>";

    private static WmlDocument List(string level, params string[] items) => IrTestDocuments.FromParts(
        string.Concat(items.Select(text =>
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            $"<w:r><w:t>{text}</w:t></w:r></w:p>")),
        numberingInnerXml:
            $"<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">{level}</w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>");

    private static string Cell(string text) =>
        $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static WmlDocument Table(params string[][] rows) => IrTestDocuments.FromParts(
        "<w:p><w:r><w:t>Before the table.</w:t></w:r></w:p>" +
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2000\"/>", rows.Max(r => r.Length))) + "</w:tblGrid>" +
        string.Concat(rows.Select(row => "<w:tr>" + string.Concat(row.Select(Cell)) + "</w:tr>")) +
        "</w:tbl><w:p><w:r><w:t>After the table.</w:t></w:r></w:p>");

    private static XDocument Body(WmlDocument document)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(document.DocumentByteArray), false);
        return package.MainDocumentPart!.GetXDocument();
    }

    private static string VisibleText(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return string.Join("|", package.MainDocumentPart!.GetXDocument().Descendants(W.p)
            .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));
    }

    /// <summary>accept(redline) reads as the revised document and reject(redline) as the original.</summary>
    private static void AssertRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline)
    {
        Assert.Equal(VisibleText(revised.DocumentByteArray), VisibleText(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray)));
        Assert.Equal(VisibleText(original.DocumentByteArray), VisibleText(DocxDiffOps.RejectRevisions(redline.DocumentByteArray)));
    }

    // ---- w:numberingChange ------------------------------------------------------------------------

    [Fact]
    public void NumberingChange_IsWrittenByDefault()
    {
        var redline = DocxCompare.Compare(List(DecimalLevel, "Alpha", "Bravo"), List(DecimalLevel, "New", "Alpha", "Bravo"));

        Assert.NotEmpty(Body(redline).Descendants(W.numberingChange));
    }

    [Fact]
    public void NumberingChange_Off_WritesNone_AndStillRoundTrips()
    {
        var original = List(DecimalLevel, "Alpha", "Bravo", "Charlie", "Delta");
        var revised = List(DecimalLevel, "New", "Alpha", "Charlie", "Delta", "Bravo"); // insert, move, shift

        var redline = DocxCompare.Compare(original, revised, new DocxDiffSettings { TrackNumberingChanges = false });

        Assert.Empty(Body(redline).Descendants(W.numberingChange));
        Assert.NotEmpty(Body(redline).Descendants(W.ins)); // the edits themselves are still tracked
        AssertRoundTrip(original, revised, redline);
    }

    [Fact]
    public void NumberingChange_Off_ReachesTheSharedWireSurface()
    {
        var original = List(DecimalLevel, "Alpha", "Bravo");
        var revised = List(DecimalLevel, "New", "Alpha", "Bravo");

        var bytes = DocxDiffOps.Compare(original.DocumentByteArray, revised.DocumentByteArray,
            "{\"trackNumberingChanges\":false}");

        Assert.Empty(Body(new WmlDocument("r.docx", bytes)).Descendants(W.numberingChange));
    }

    [Fact]
    public void NumberingChange_Off_AppliesToConsolidateToo()
    {
        var original = List(DecimalLevel, "Alpha", "Bravo", "Charlie");
        var reviewer = new DocxDiffReviewer { Author = "Reviewer", Document = List(DecimalLevel, "New", "Alpha", "Charlie") };

        var consolidated = DocxDiff.Consolidate(original, new[] { reviewer },
            new DocxDiffConsolidateSettings { Diff = new DocxDiffSettings { TrackNumberingChanges = false } });

        Assert.Empty(Body(consolidated).Descendants(W.numberingChange));
    }

    [Fact]
    public void NumberingChange_IsNeverWrittenForAListLevelWithNoLabel()
    {
        // A level whose lvlText is empty displays no label, so there is nothing to record — and an
        // empty w:original is what the validator rejects.
        const string unlabeled = "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"\"/>";
        var original = List(unlabeled, "Alpha", "Bravo", "Charlie");

        var redline = DocxCompare.Compare(original, List(unlabeled, "New", "Alpha", "Charlie"));

        Assert.DoesNotContain(Body(redline).Descendants(W.numberingChange),
            change => string.IsNullOrEmpty((string?)change.Attribute(W.original)));
        NoNewValidationErrors(original.DocumentByteArray, redline.DocumentByteArray);
    }

    // ---- w:cellIns / w:cellDel ---------------------------------------------------------------------

    [Fact]
    public void CellRevisions_AreWrittenByDefault_ForAnAddedColumn()
    {
        var redline = DocxCompare.Compare(
            Table(new[] { "A1", "B1" }, new[] { "A2", "B2" }),
            Table(new[] { "A1", "B1", "C1" }, new[] { "A2", "B2", "C2" }));

        Assert.NotEmpty(Body(redline).Descendants(W.cellIns));
    }

    [Theory]
    [InlineData(true)]  // a column added
    [InlineData(false)] // a column removed
    public void CellRevisions_Off_TrackTheTableWhole_AndStillRoundTrip(bool columnAdded)
    {
        var narrow = Table(new[] { "A1", "B1" }, new[] { "A2", "B2" });
        var wide = Table(new[] { "A1", "B1", "C1" }, new[] { "A2", "B2", "C2" });
        var (original, revised) = columnAdded ? (narrow, wide) : (wide, narrow);
        var settings = new DocxDiffSettings { TrackCellInsertionsAndDeletions = false };

        var redline = DocxCompare.Compare(original, revised, settings);

        var body = Body(redline);
        Assert.Empty(body.Descendants(W.cellIns));
        Assert.Empty(body.Descendants(W.cellDel));
        // Word's construct: the original table deleted row by row, the revised one inserted.
        var tables = body.Descendants(W.tbl).ToList();
        Assert.Equal(2, tables.Count);
        Assert.All(tables[0].Elements(W.tr), row => Assert.NotNull(row.Element(W.trPr)?.Element(W.del)));
        Assert.All(tables[1].Elements(W.tr), row => Assert.NotNull(row.Element(W.trPr)?.Element(W.ins)));
        AssertRoundTrip(original, revised, redline);
        NoNewValidationErrors(original.DocumentByteArray, redline.DocumentByteArray);

        // The revision list describes the same whole-table change the markup draws.
        var revisions = DocxDiff.GetRevisions(original, revised, settings);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted && r.Text.Contains("A1"));
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains("A1"));
    }
}
