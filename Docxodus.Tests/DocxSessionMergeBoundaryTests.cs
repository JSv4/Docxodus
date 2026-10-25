using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionMergeBoundaryTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void UntrackedMerge_KeepsTheSecondParagraphsSectionBoundary(bool firstEndsSection, bool secondEndsSection)
    {
        using var session = new DocxSession(Build(
            Paragraph("First.", "center", firstEndsSection ? Section(10000) : null),
            Paragraph("Second.", "right", secondEndsSection ? Section(14000) : null),
            Paragraph("Next section.")));
        var first = session.FindByText("First.")!.Anchor.Id;
        var second = session.FindByText("Second.")!.Anchor.Id;
        var expectedSection = Body(session.Save()).Elements(W.p).Skip(1).First().Element(W.pPr)?.Element(W.sectPr);

        var result = session.MergeParagraphs(first, second);

        Assert.True(result.Success, result.Error?.Message);
        var paragraphs = Body(session.Save()).Elements(W.p).ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.Equal("First. Second.", string.Concat(paragraphs[0].Descendants(W.t).Select(t => t.Value)));
        Assert.Equal("center", (string?)paragraphs[0].Element(W.pPr)?.Element(W.jc)?.Attribute(W.val));
        Assert.True(XNode.DeepEquals(expectedSection, paragraphs[0].Element(W.pPr)?.Element(W.sectPr)),
            paragraphs[0].ToString());
        AssertSchemaValid(session.Save());
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, false)]
    [InlineData(TrackedChangeMode.Accept, true)]
    [InlineData(TrackedChangeMode.RenderInline, false)]
    [InlineData(TrackedChangeMode.RenderInline, true)]
    public void Merge_WithATableAnchor_IsRefusedWithoutMutation(TrackedChangeMode tracking, bool tableFirst)
    {
        using var session = new DocxSession(Build(Paragraph("Before."), Table(), Paragraph("After.")),
            new DocxSessionSettings { TrackedChanges = tracking, RevisionAuthor = "Reviewer" });
        var table = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "tbl").Anchor.Id;
        var paragraph = session.FindByText(tableFirst ? "After." : "Before.")!.Anchor.Id;
        var before = Body(session.Save()).ToString();
        var version = session.Version;
        var undoCount = session.UndoCount;

        var result = session.MergeParagraphs(tableFirst ? table : paragraph, tableFirst ? paragraph : table);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, result.Error?.Code);
        Assert.Equal(table, result.Error?.AnchorId);
        Assert.Equal(before, Body(session.Save()).ToString());
        Assert.Equal(version, session.Version);
        Assert.Equal(undoCount, session.UndoCount);
        AssertSchemaValid(session.Save());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void UntrackedMerge_KeepsTheSectionsPendingChangeAndPropertyOrder(bool firstHasPendingFormatting, bool accept)
    {
        var first = Paragraph("First.");
        if (firstHasPendingFormatting)
            first.AddFirst(new XElement(W.pPr,
                new XElement(W.jc, new XAttribute(W.val, "center")),
                new XElement(W.pPrChange, RevisionAttributes(30),
                    new XElement(W.pPr, new XElement(W.jc, new XAttribute(W.val, "left"))))));
        var section = Section(14000);
        section.Add(new XElement(W.sectPrChange, RevisionAttributes(31), Section(10000)));
        using var session = new DocxSession(Build(first, Paragraph("Second.", "right", section), Paragraph("Next section.")));
        var a = session.FindByText("First.")!.Anchor.Id;
        var b = session.FindByText("Second.")!.Anchor.Id;

        Assert.True(session.MergeParagraphs(a, b).Success);

        var merged = Body(session.Save()).Elements(W.p).First();
        Assert.True(XNode.DeepEquals(section, merged.Element(W.pPr)?.Element(W.sectPr)), merged.ToString());
        AssertSchemaValid(session.Save());
        Assert.True((accept ? session.AcceptAllRevisions() : session.RejectAllRevisions()).Success);
        var pPr = Body(session.Save()).Elements(W.p).First().Element(W.pPr)!;
        Assert.Equal(accept ? "14000" : "10000", (string?)pPr.Element(W.sectPr)?.Element(W.pgSz)?.Attribute(W._w));
        Assert.Equal(firstHasPendingFormatting ? accept ? "center" : "left" : null,
            (string?)pPr.Element(W.jc)?.Attribute(W.val));
        AssertSchemaValid(session.Save());
    }

    private static object[] RevisionAttributes(int id) =>
    [
        new XAttribute(W.id, id), new XAttribute(W.author, "Reviewer"),
        new XAttribute(W.date, "2026-10-23T12:00:00Z"),
    ];

    private static XElement Paragraph(string text, string? alignment = null, XElement? section = null) =>
        new(W.p, alignment is not null || section is not null
            ? new XElement(W.pPr, alignment is not null ? new XElement(W.jc, new XAttribute(W.val, alignment)) : null, section)
            : null, new XElement(W.r, new XElement(W.t, text)));

    private static XElement Section(int width) => new(W.sectPr,
        new XElement(W.pgSz, new XAttribute(W._w, width), new XAttribute(W.h, 15840)));

    private static XElement Table() => new(W.tbl,
        new XElement(W.tblPr, new XElement(W.tblW, new XAttribute(W._w, 2400), new XAttribute(W.type, "dxa"))),
        new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, 2400))),
        new XElement(W.tr, new XElement(W.tc, Paragraph("Cell."))));

    private static byte[] Build(params XElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w), new XElement(W.body, blocks, Section(12240)))));
        }
        var bytes = stream.ToArray();
        AssertSchemaValid(bytes);
        return bytes;
    }

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(document));
    }
}
