// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Accepting a document that contains a move drops each table row the move emptied. A row whose
/// cell content sits inside a block wrapper (<c>w:customXml</c>, <c>w:sdt</c>), at cell or at row
/// level, is not empty and must survive, even when the move is somewhere else in the document
/// (issue #862).
/// </summary>
public class RevisionProcessorMoveEmptiedRowTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = WNs;

    private const string Rev = "w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"";

    // A paragraph moved from the top of the body to the bottom, unrelated to any table.
    private const string MovedFrom =
        "<w:moveFromRangeStart w:id=\"1\" w:name=\"move1\" " + Rev + "/>" +
        "<w:p><w:pPr><w:rPr><w:moveFrom w:id=\"2\" " + Rev + "/></w:rPr></w:pPr>" +
        "<w:moveFrom w:id=\"3\" " + Rev + "><w:r><w:t>moved paragraph</w:t></w:r></w:moveFrom></w:p>" +
        "<w:moveFromRangeEnd w:id=\"1\"/>";

    private const string MovedTo =
        "<w:moveToRangeStart w:id=\"4\" w:name=\"move1\" " + Rev + "/>" +
        "<w:p><w:pPr><w:rPr><w:moveTo w:id=\"5\" " + Rev + "/></w:rPr></w:pPr>" +
        "<w:moveTo w:id=\"6\" " + Rev + "><w:r><w:t>moved paragraph</w:t></w:r></w:moveTo></w:p>" +
        "<w:moveToRangeEnd w:id=\"4\"/>";

    private static string Table(params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" + string.Concat(rows) + "</w:tbl>";

    private static string Cell(string content) =>
        "<w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" + content + "</w:tc>";

    private static string Para(string text) => "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>";

    private static string CustomXml(string content) =>
        "<w:customXml w:uri=\"urn:example\" w:element=\"clause\">" + content + "</w:customXml>";

    private static string Sdt(string content) =>
        "<w:sdt><w:sdtPr><w:id w:val=\"9\"/></w:sdtPr><w:sdtContent>" + content + "</w:sdtContent></w:sdt>";

    private static XElement AcceptBody(string bodyXml)
    {
        var docXml =
            "<w:document xmlns:w=\"" + WNs + "\"><w:body>" + bodyXml +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";

        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var partStream = main.GetStream(FileMode.Create))
            using (var writer = new StreamWriter(partStream))
                writer.Write(docXml);
            main.AddNewPart<StyleDefinitionsPart>().Styles =
                new DocumentFormat.OpenXml.Wordprocessing.Styles(new DocumentFormat.OpenXml.Wordprocessing.DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("d.docx", stream.ToArray()));
        using var acceptedStream = new MemoryStream(accepted.DocumentByteArray);
        using var acceptedDoc = WordprocessingDocument.Open(acceptedStream, false);
        using var reader = new StreamReader(acceptedDoc.MainDocumentPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd()).Root!.Element(W + "body")!;
    }

    private static string[] RowTexts(XElement body) =>
        body.Descendants(W + "tr")
            .Select(tr => string.Concat(tr.Descendants(W + "t").Select(t => t.Value)))
            .ToArray();

    // Today accept strips the cell-level w:customXml wrapper before the row check runs (#913), so
    // this passes on the old predicate too. It pins the row once the wrapper survives accept.
    [Fact]
    public void Accept_KeepsRowWhoseCellParagraphIsWrappedInCustomXml_WhenAMoveIsElsewhere()
    {
        var body = AcceptBody(
            MovedFrom +
            Table(
                "<w:tr>" + Cell(CustomXml(Para("wrapped clause"))) + "</w:tr>",
                "<w:tr>" + Cell(Para("plain row")) + "</w:tr>") +
            MovedTo);

        Assert.Equal(new[] { "wrapped clause", "plain row" }, RowTexts(body));
    }

    [Fact]
    public void Accept_KeepsRowWhoseCellsAreWrappedInCustomXmlAtRowLevel_WhenAMoveIsElsewhere()
    {
        var body = AcceptBody(
            MovedFrom +
            Table("<w:tr>" + CustomXml(Cell(Para("row-level wrapper"))) + "</w:tr>") +
            MovedTo);

        Assert.Equal(new[] { "row-level wrapper" }, RowTexts(body));
    }

    [Fact]
    public void Accept_KeepsRowWhoseCellsAreWrappedInAContentControlAtRowLevel_WhenAMoveIsElsewhere()
    {
        var body = AcceptBody(
            MovedFrom +
            Table("<w:tr>" + Sdt(Cell(Para("cell control"))) + "</w:tr>") +
            MovedTo);

        Assert.Equal(new[] { "cell control" }, RowTexts(body));
    }

    [Fact]
    public void Accept_KeepsTheMovedParagraphAndTheUnrelatedTableTogether()
    {
        var body = AcceptBody(
            MovedFrom +
            Table("<w:tr>" + Cell(CustomXml(Para("wrapped clause"))) + "</w:tr>") +
            MovedTo);

        var text = string.Concat(body.Descendants(W + "t").Select(t => t.Value));
        Assert.Equal("wrapped clausemoved paragraph", text);
        Assert.Single(body.Elements(W + "tbl"));
    }

    [Fact]
    public void Accept_StillDropsARowTheMoveEmptied_EvenInsideAWrapper()
    {
        // The second row's only paragraph moved out of the table: that row really is empty after
        // accept, wrapper or not, and the first row stays.
        var movedRow =
            "<w:tr>" + Cell(
                "<w:moveFromRangeStart w:id=\"11\" w:name=\"move2\" " + Rev + "/>" +
                CustomXml(
                    "<w:p><w:pPr><w:rPr><w:moveFrom w:id=\"12\" " + Rev + "/></w:rPr></w:pPr>" +
                    "<w:moveFrom w:id=\"13\" " + Rev + "><w:r><w:t>leaving row</w:t></w:r></w:moveFrom></w:p>") +
                "<w:moveFromRangeEnd w:id=\"11\"/>") + "</w:tr>";
        var body = AcceptBody(
            Table("<w:tr>" + Cell(Para("staying row")) + "</w:tr>", movedRow) +
            "<w:moveToRangeStart w:id=\"14\" w:name=\"move2\" " + Rev + "/>" +
            "<w:p><w:pPr><w:rPr><w:moveTo w:id=\"15\" " + Rev + "/></w:rPr></w:pPr>" +
            "<w:moveTo w:id=\"16\" " + Rev + "><w:r><w:t>leaving row</w:t></w:r></w:moveTo></w:p>" +
            "<w:moveToRangeEnd w:id=\"14\"/>");

        Assert.Equal(new[] { "staying row" }, RowTexts(body));
    }
}
