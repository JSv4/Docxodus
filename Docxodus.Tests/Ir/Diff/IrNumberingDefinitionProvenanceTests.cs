#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Numbering definitions live in a package part rather than on each paragraph.  When two otherwise-identical
/// paragraphs retain the same numId but that numId resolves to a different definition, the redline must still
/// switch the displayed numbering on Accept and restore it on Reject.
/// </summary>
public class IrNumberingDefinitionProvenanceTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Fact]
    public void Compare_SameNumIdWithChangedDefinition_RoundTripsResolvedNumbering()
    {
        var left = Numbered("decimal");
        var right = Numbered("bullet");

        var redline = DocxDiff.Compare(left, right);

        var currentPPr = MainBody(redline).Element(W + "p")!.Element(W + "pPr")!;
        Assert.Equal("2", (string?)currentPPr.Element(W + "numPr")!.Element(W + "numId")!.Attribute(W + "val"));
        Assert.Equal("1", (string?)currentPPr.Element(W + "pPrChange")!.Element(W + "pPr")!
            .Element(W + "numPr")!.Element(W + "numId")!.Attribute(W + "val"));
        AssertSchemaValid(redline);
        Assert.Equal("bullet", ResolvedFormat(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("decimal", ResolvedFormat(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_CollidingDefinitionWithConcurrentPPrChange_RestoresBothLeftValues()
    {
        var left = Numbered("decimal", pPrSuffix: "<w:jc w:val=\"left\"/>");
        var right = Numbered("bullet", pPrSuffix: "<w:jc w:val=\"center\"/>");

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        Assert.Equal("bullet", ResolvedFormat(accepted));
        Assert.Equal("center", ResolvedJustification(accepted));

        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal("decimal", ResolvedFormat(rejected));
        Assert.Equal("left", ResolvedJustification(rejected));
    }

    [Fact]
    public void Compare_CollidingDefinitionOnInsertedParagraph_DoesNotAddRedundantPPrHistory()
    {
        var left = WithNumbering("decimal", "");
        var right = Numbered("bullet");

        var redline = DocxDiff.Compare(left, right);
        var pPr = MainBody(redline).Element(W + "p")!.Element(W + "pPr")!;

        Assert.Equal("2", (string?)pPr.Element(W + "numPr")!.Element(W + "numId")!.Attribute(W + "val"));
        Assert.Null(pPr.Element(W + "pPrChange"));
        AssertSchemaValid(redline);
        Assert.Equal("bullet", ResolvedFormat(RevisionProcessor.AcceptRevisions(redline)));
    }

    [Fact]
    public void Compare_CollidingDefinitionOnDeletedParagraph_PreservesLeftDefinition()
    {
        var left = Numbered("decimal");
        var right = WithNumbering("bullet", "");

        var redline = DocxDiff.Compare(left, right);
        var pPr = MainBody(redline).Element(W + "p")!.Element(W + "pPr")!;

        Assert.Equal("1", (string?)pPr.Element(W + "numPr")!.Element(W + "numId")!.Attribute(W + "val"));
        Assert.Null(pPr.Element(W + "pPrChange"));
        AssertSchemaValid(redline);
        Assert.Equal("decimal", ResolvedFormat(RevisionProcessor.RejectRevisions(redline)));
    }

    [Fact]
    public void Compare_CollidingDefinitionInFootnote_RoundTripsResolvedNumbering()
    {
        var left = NumberedFootnote("decimal");
        var right = NumberedFootnote("bullet");

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);

        Assert.Equal("bullet", ResolvedFootnoteFormat(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("decimal", ResolvedFootnoteFormat(RevisionProcessor.RejectRevisions(redline)));
    }

    private static WmlDocument Numbered(string format, string pPrSuffix = "") => WithNumbering(format,
        "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>" +
        pPrSuffix + "</w:pPr><w:r><w:t>List item</w:t></w:r></w:p>");

    private static WmlDocument WithNumbering(string format, string bodyInnerXml) => IrTestDocuments.FromParts(
        bodyInnerXml,
        numberingInnerXml:
            "<w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\">" +
            "<w:start w:val=\"1\"/><w:numFmt w:val=\"" + format + "\"/>" +
            "<w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num>");

    private static WmlDocument NumberedFootnote(string format)
    {
        using var stream = new MemoryStream();
        using (var wdoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = wdoc.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            WritePartXml(styles, $"<w:styles xmlns:w=\"{W}\"/>");
            var settings = main.AddNewPart<DocumentSettingsPart>();
            WritePartXml(settings, $"<w:settings xmlns:w=\"{W}\"/>");
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            WritePartXml(numbering,
                $"<w:numbering xmlns:w=\"{W}\"><w:abstractNum w:abstractNumId=\"1\">" +
                $"<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{format}\"/>" +
                "<w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>" +
                "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>");
            var footnotes = main.AddNewPart<FootnotesPart>();
            WritePartXml(footnotes,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"1\"><w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/>" +
                "</w:numPr></w:pPr><w:r><w:t>Footnote item</w:t></w:r></w:p></w:footnote></w:footnotes>");
            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:footnoteReference w:id=\"1\"/>" +
                "</w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("numbered-footnote.docx", stream.ToArray());
    }

    private static XElement MainBody(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        return new XElement(wdoc.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!);
    }

    private static string? ResolvedFormat(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var main = wdoc.MainDocumentPart!;
        var body = main.GetXDocument().Root!.Element(W + "body")!;
        int numId = int.Parse((string)body.Elements(W + "p").Single()
            .Element(W + "pPr")!.Element(W + "numPr")!.Element(W + "numId")!.Attribute(W + "val")!);
        var numbering = main.NumberingDefinitionsPart!.GetXDocument().Root!;
        int abstractId = int.Parse((string)numbering.Elements(W + "num")
            .Single(element => (string?)element.Attribute(W + "numId") == numId.ToString())
            .Element(W + "abstractNumId")!.Attribute(W + "val")!);
        return (string?)numbering.Elements(W + "abstractNum")
            .Single(element => (string?)element.Attribute(W + "abstractNumId") == abstractId.ToString())
            .Elements(W + "lvl").Single(element => (string?)element.Attribute(W + "ilvl") == "0")
            .Element(W + "numFmt")?.Attribute(W + "val");
    }

    private static string? ResolvedFootnoteFormat(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var main = wdoc.MainDocumentPart!;
        var paragraph = main.FootnotesPart!.GetXDocument().Root!.Elements(W + "footnote")
            .Single(note => (string?)note.Attribute(W + "id") == "1").Element(W + "p")!;
        int numId = int.Parse((string)paragraph.Element(W + "pPr")!.Element(W + "numPr")!
            .Element(W + "numId")!.Attribute(W + "val")!);
        var numbering = main.NumberingDefinitionsPart!.GetXDocument().Root!;
        int abstractId = int.Parse((string)numbering.Elements(W + "num")
            .Single(element => (string?)element.Attribute(W + "numId") == numId.ToString())
            .Element(W + "abstractNumId")!.Attribute(W + "val")!);
        return (string?)numbering.Elements(W + "abstractNum")
            .Single(element => (string?)element.Attribute(W + "abstractNumId") == abstractId.ToString())
            .Elements(W + "lvl").Single(element => (string?)element.Attribute(W + "ilvl") == "0")
            .Element(W + "numFmt")?.Attribute(W + "val");
    }

    private static string? ResolvedJustification(WmlDocument doc) =>
        (string?)MainBody(doc).Element(W + "p")?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val");

    private static void AssertSchemaValid(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(wdoc)
            .Select(error => $"{error.Id}@{error.Path?.XPath}: {error.Description}")
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(partStream);
        writer.Write(xml);
    }
}
