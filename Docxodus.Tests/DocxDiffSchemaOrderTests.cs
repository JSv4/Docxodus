// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// A comparison writes tables, rows and properties in the order the WordprocessingML schema requires,
/// and never leaves a table without rows (issue #837).
/// </summary>
public class DocxDiffSchemaOrderTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private const string Tail = "<w:p><w:r><w:t>After the table</w:t></w:r></w:p>";

    /// <summary>Row-level table property exceptions, and no row properties of its own.</summary>
    private const string Exceptions =
        "<w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>" +
        "</w:tblBorders></w:tblPrEx>";

    private static string Row(string text, string leading = "") =>
        $"<w:tr>{leading}<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
        $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr>";

    private static string Table(params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
        string.Concat(rows) + "</w:tbl>";

    private static readonly WmlDocument TwoRows =
        IrTestDocuments.FromBodyXml(Table(Row("kept"), Row("exceptional", Exceptions)) + Tail);

    private static readonly WmlDocument OneRow = IrTestDocuments.FromBodyXml(Table(Row("kept")) + Tail);

    /// <summary>A table every row of which is tracked-deleted, the text moved out of it — as Word writes a
    /// table whose content was cut and pasted elsewhere with Track Changes on.</summary>
    private static readonly WmlDocument MovedAwayTable = IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
        Table("<w:tr><w:trPr><w:del w:id=\"1\" w:author=\"A\"/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
              "<w:p><w:pPr><w:rPr><w:del w:id=\"2\" w:author=\"A\"/></w:rPr></w:pPr>" +
              "<w:moveFrom w:id=\"3\" w:author=\"A\"><w:del w:id=\"4\" w:author=\"A\"><w:r><w:delText>Moved</w:delText></w:r>" +
              "</w:del></w:moveFrom></w:p></w:tc></w:tr>") +
        Tail);

    /// <summary>The same table inside a move-from range that opens in the paragraph before it and closes
    /// as the table's last child, as Word wrote it in <c>RA001-Tracked-Revisions-02.docx</c> — the range
    /// covers the table's properties and grid as well as its rows, but not the table element itself.</summary>
    private static readonly WmlDocument MovedAwayRange = IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Before</w:t></w:r><w:moveFromRangeStart w:id=\"10\" w:author=\"A\" w:name=\"move1\"/></w:p>" +
        Table("<w:tr><w:trPr><w:del w:id=\"11\" w:author=\"A\"/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
              "<w:p><w:pPr><w:rPr><w:del w:id=\"12\" w:author=\"A\"/></w:rPr></w:pPr>" +
              "<w:moveFrom w:id=\"13\" w:author=\"A\"><w:r><w:t>Moved</w:t></w:r></w:moveFrom></w:p></w:tc></w:tr>",
              "<w:moveFromRangeEnd w:id=\"10\"/>") +
        Tail);

    private static readonly WmlDocument NoTable =
        IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Before</w:t></w:r></w:p>" + Tail);

    [Theory]
    [InlineData(true)]  // the row is deleted
    [InlineData(false)] // the row is inserted
    public void Compare_WholeRowWithPropertyExceptions_KeepsThemFirstInTheRow(bool deleted)
    {
        var (left, right) = deleted ? (TwoRows, OneRow) : (OneRow, TwoRows);

        var row = Body(DocxCompare.Compare(left, right)).Descendants(W + "tr").Last();

        Assert.Equal(new[] { "tblPrEx", "trPr", "tc" }, row.Elements().Select(e => e.Name.LocalName).ToArray());
    }

    [Fact]
    public void Compare_WholeRowWithPropertyExceptions_AddsNoValidationErrors() =>
        NoNewValidationErrors(TwoRows.DocumentByteArray, DocxCompare.Compare(TwoRows, OneRow).DocumentByteArray);

    [Fact]
    public void Consolidate_WholeRowWithPropertyExceptions_AddsNoValidationErrors()
    {
        var reviewer = new DocxDiffReviewer { Author = "Reviewer", Document = OneRow };

        NoNewValidationErrors(
            TwoRows.DocumentByteArray, DocxDiff.Consolidate(TwoRows, new[] { reviewer }).DocumentByteArray);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void Compare_TableWhoseRowsWereAllMovedAway_LeavesNoRowlessTable(bool movedAwayOnTheRight, bool inMoveRange)
    {
        var movedAway = inMoveRange ? MovedAwayRange : MovedAwayTable;
        var (left, right) = movedAwayOnTheRight ? (NoTable, movedAway) : (movedAway, NoTable);

        var output = DocxCompare.Compare(left, right);

        Assert.DoesNotContain(Body(output).Descendants(W + "tbl"), table => !table.Elements(W + "tr").Any());
        NoNewValidationErrors(NoTable.DocumentByteArray, output.DocumentByteArray);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AcceptRevisions_TableWhoseRowsWereAllMovedAway_RemovesTheTable(bool inMoveRange) =>
        Assert.Empty(Body(RevisionProcessor.AcceptRevisions(inMoveRange ? MovedAwayRange : MovedAwayTable))
            .Descendants(W + "tbl"));

    [Fact]
    public void AcceptRevisions_TableThatArrivedWithoutRows_IsLeftAlone() =>
        Assert.Single(Body(RevisionProcessor.AcceptRevisions(IrTestDocuments.FromBodyXml(Table() + Tail)))
            .Descendants(W + "tbl"));

    private static XElement Body(WmlDocument document)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var part = package.GetEntry("word/document.xml")!.Open();
        return XElement.Load(part).Element(W + "body")!;
    }
}
