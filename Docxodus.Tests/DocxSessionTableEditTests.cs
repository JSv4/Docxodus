#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for post-insert table editing on <see cref="DocxSession"/>: insert/delete row,
/// insert/delete column — addressed by a cell-paragraph anchor. Test IDs use the DT2xx range.
/// </summary>
public class DocxSessionTableEditTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement DocumentXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!;
    }

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    /// <summary>Insert a rows×cols table and return its created cell-paragraph anchors (row-major).</summary>
    private static (DocxSession session, string[] cells) NewTable(int rows, int cols)
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var r = session.InsertTable(anchor, Position.After, rows, cols);
        Assert.True(r.Success, r.Error?.Message);
        return (session, r.Created.Select(a => a.Id).ToArray());
    }

    private static XElement SingleTable(DocxSession session) =>
        DocumentXml(session.Save()).Descendants(W + "tbl").Single();

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(wDoc)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
        Assert.True(errors.Count == 0, "OOXML schema errors:\n" + string.Join("\n", errors));
    }

    [Fact]
    public void DT201_InsertTableRow_After_AddsRowWithSameColumnCount()
    {
        var (session, cells) = NewTable(2, 2); // cells row-major: r0c0,r0c1,r1c0,r1c1
        var r = session.InsertTableRow(cells[0], Position.After); // after row 0
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // the new row's two cell paragraphs

        var tbl = SingleTable(session);
        Assert.Equal(3, tbl.Elements(W + "tr").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(2, tr.Elements(W + "tc").Count()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT202_InsertTableColumn_After_AddsColumnToEveryRow()
    {
        var (session, cells) = NewTable(2, 2);
        var r = session.InsertTableColumn(cells[0], Position.After); // after column 0
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // one new cell per row

        var tbl = SingleTable(session);
        Assert.Equal(3, tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(3, tr.Elements(W + "tc").Count()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT203_DeleteTableRow_RemovesTheRow()
    {
        var (session, cells) = NewTable(3, 2);
        var r = session.DeleteTableRow(cells[2]); // a cell in row 1
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(2, tbl.Elements(W + "tr").Count());
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT204_DeleteTableColumn_RemovesTheColumnFromEveryRow()
    {
        var (session, cells) = NewTable(2, 3);
        var r = session.DeleteTableColumn(cells[1]); // column 1
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(2, tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(2, tr.Elements(W + "tc").Count()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT205_DeleteLastRow_RemovesTheWholeTable()
    {
        var (session, cells) = NewTable(1, 2);
        var r = session.DeleteTableRow(cells[0]);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Empty(DocumentXml(session.Save()).Descendants(W + "tbl"));
    }

    [Fact]
    public void DT206_DeleteLastColumn_RemovesTheWholeTable()
    {
        var (session, cells) = NewTable(2, 1);
        var r = session.DeleteTableColumn(cells[0]);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Empty(DocumentXml(session.Save()).Descendants(W + "tbl"));
    }

    [Fact]
    public void DT207_InsertRow_NonCellAnchor_IsRejected()
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var bodyP = FirstBodyParagraph(session);
        var r = session.InsertTableRow(bodyP, Position.After);
        Assert.False(r.Success); // a body paragraph is not in a table
    }
}
