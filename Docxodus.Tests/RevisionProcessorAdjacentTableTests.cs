using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Accepting or rejecting revisions joins two tables only when resolving a revision is what made them
/// adjacent, such as a deleted paragraph that sat between them. Tables that were already adjacent in the
/// input stay separate and keep their own properties, as Word keeps them (issue #1341).
/// </summary>
public class RevisionProcessorAdjacentTableTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = WNs;
    private const string Rev = "w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"";

    private const string PlainTable =
        "<w:tbl><w:tblPr><w:tblW w:w=\"4800\" w:type=\"dxa\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"4800\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:p><w:r><w:t>First table.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private const string BorderedTable =
        "<w:tbl><w:tblPr><w:tblW w:w=\"5000\" w:type=\"dxa\"/><w:tblBorders>" +
        "<w:top w:val=\"single\" w:sz=\"8\"/><w:left w:val=\"single\" w:sz=\"8\"/>" +
        "<w:bottom w:val=\"single\" w:sz=\"8\"/><w:right w:val=\"single\" w:sz=\"8\"/></w:tblBorders>" +
        "<w:tblCellMar><w:left w:w=\"720\" w:type=\"dxa\"/><w:right w:w=\"180\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:p><w:r><w:t>Second table.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private static string Para(string text) => "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>";

    private const string DeletedPara =
        "<w:p><w:pPr><w:rPr><w:del w:id=\"1\" " + Rev + "/></w:rPr></w:pPr>" +
        "<w:del w:id=\"2\" " + Rev + "><w:r><w:delText>Between.</w:delText></w:r></w:del></w:p>";

    private const string InsertedPara =
        "<w:p><w:pPr><w:rPr><w:ins w:id=\"3\" " + Rev + "/></w:rPr></w:pPr>" +
        "<w:ins w:id=\"4\" " + Rev + "><w:r><w:t>Between.</w:t></w:r></w:ins></w:p>";

    private const string ParaWithInsertion =
        "<w:p><w:r><w:t>Before</w:t></w:r><w:ins w:id=\"5\" " + Rev + "><w:r><w:t> more</w:t></w:r></w:ins></w:p>";

    [Fact]
    public void Accept_keeps_adjacent_tables_without_revisions_apart()
    {
        var body = Resolve(Para("Before.") + PlainTable + BorderedTable + Para("After."), accept: true);
        AssertTwoSeparateTables(body);
    }

    [Fact]
    public void Reject_keeps_adjacent_tables_without_revisions_apart()
    {
        var body = Resolve(Para("Before.") + PlainTable + BorderedTable + Para("After."), accept: false);
        AssertTwoSeparateTables(body);
    }

    [Fact]
    public void Accept_keeps_adjacent_tables_apart_when_another_paragraph_has_a_revision()
    {
        var body = Resolve(ParaWithInsertion + PlainTable + BorderedTable + Para("After."), accept: true);
        AssertTwoSeparateTables(body);
    }

    [Fact]
    public void Accept_joins_tables_once_the_deleted_paragraph_between_them_is_gone()
    {
        var body = Resolve(Para("Before.") + PlainTable + DeletedPara + BorderedTable + Para("After."), accept: true);
        AssertJoinedUnderFirstTableProperties(body);
    }

    [Fact]
    public void Reject_joins_tables_once_the_inserted_paragraph_between_them_is_gone()
    {
        var body = Resolve(Para("Before.") + PlainTable + InsertedPara + BorderedTable + Para("After."), accept: false);
        AssertJoinedUnderFirstTableProperties(body);
    }

    [Fact]
    public void Accept_joins_only_the_tables_a_deleted_paragraph_separated()
    {
        // The plain and bordered tables were already adjacent; the deleted paragraph separated the
        // bordered table from a third. Only that second boundary is the revision's to close.
        var third = PlainTable.Replace("First table.", "Third table.");
        var body = Resolve(PlainTable + BorderedTable + DeletedPara + third + Para("After."), accept: true);

        var tables = body.Elements(W + "tbl").ToList();
        Assert.Equal(2, tables.Count);
        Assert.Equal(new[] { "First table." }, RowTexts(tables[0]));
        Assert.Equal(new[] { "Second table.", "Third table." }, RowTexts(tables[1]));
        Assert.NotNull(tables[1].Element(W + "tblPr")!.Element(W + "tblBorders"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Resolution_joins_tables_after_removing_an_entire_intervening_table(bool accept)
    {
        var mark = accept ? "del" : "ins";
        var removed = PlainTable.Replace("First table.", "Removed table.")
            .Replace("<w:tr>", "<w:tr><w:trPr><w:" + mark + " w:id=\"17\" " + Rev + "/></w:trPr>");
        var body = Resolve(PlainTable + removed + BorderedTable + Para("After."), accept);

        AssertJoinedUnderFirstTableProperties(body);
    }

    private static void AssertTwoSeparateTables(XElement body)
    {
        var tables = body.Elements(W + "tbl").ToList();
        Assert.Equal(2, tables.Count);
        Assert.Equal(new[] { "First table." }, RowTexts(tables[0]));
        Assert.Equal(new[] { "Second table." }, RowTexts(tables[1]));
        Assert.Null(tables[0].Element(W + "tblPr")!.Element(W + "tblBorders"));
        var second = tables[1].Element(W + "tblPr")!;
        Assert.NotNull(second.Element(W + "tblBorders"));
        Assert.Equal("5000", (string?)second.Element(W + "tblW")!.Attribute(W + "w"));
        Assert.Equal("720", (string?)second.Element(W + "tblCellMar")!.Element(W + "left")!.Attribute(W + "w"));
    }

    private static void AssertJoinedUnderFirstTableProperties(XElement body)
    {
        var table = Assert.Single(body.Elements(W + "tbl"));
        Assert.Equal(new[] { "First table.", "Second table." }, RowTexts(table));
        Assert.Equal("4800", (string?)table.Element(W + "tblPr")!.Element(W + "tblW")!.Attribute(W + "w"));
        Assert.DoesNotContain(body.Elements(W + "p"), p => p.Value == "Between.");
    }

    private static string[] RowTexts(XElement table) =>
        table.Elements(W + "tr").Select(tr => string.Concat(tr.Descendants(W + "t").Select(t => t.Value))).ToArray();

    private static XElement Resolve(string bodyXml, bool accept)
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
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
        }

        var input = new WmlDocument("d.docx", stream.ToArray());
        var resolved = accept ? RevisionProcessor.AcceptRevisions(input) : RevisionProcessor.RejectRevisions(input);
        using var resolvedStream = new MemoryStream(resolved.DocumentByteArray);
        using var resolvedDoc = WordprocessingDocument.Open(resolvedStream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(resolvedDoc));
        using var reader = new StreamReader(resolvedDoc.MainDocumentPart!.GetStream());
        var body = XDocument.Parse(reader.ReadToEnd()).Root!.Element(W + "body")!;
        Assert.DoesNotContain(body.DescendantsAndSelf().Attributes(), a =>
            a.Name.LocalName is "FollowsTable" or "TableBoundaryId");
        return body;
    }
}
