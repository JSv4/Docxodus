using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffDefaultStyleRemovalTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData("unchanged", false)]
    [InlineData("unchanged", true)]
    [InlineData("edit", false)]
    [InlineData("insert", false)]
    [InlineData("delete", false)]
    [InlineData("split", false)]
    [InlineData("merge", false)]
    [InlineData("override", true)]
    [InlineData("retained-style", false)]
    [InlineData("builtin-defaults", false)]
    [InlineData("id-collision", false)]
    public void RemovingDefaultRole_PreservesBothFormattingEndpoints(string change, bool table)
    {
        var left = Create(false, change, table);
        var right = Create(true, change, table);
        var redline = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(Presentation(right), Presentation(accepted));
        Assert.Equal(Presentation(left), Presentation(rejected));
        foreach (var document in new[] { left, right, redline, accepted, rejected })
        {
            using var memory = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(memory, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
    }

    private static WmlDocument Create(bool revised, string change, bool table)
    {
        string P(string text) => "<w:p>" + (change == "override" ? "<w:pPr><w:spacing w:after='80'/></w:pPr>" : "") +
            "<w:r>" + (change == "override" ? "<w:rPr><w:sz w:val='18'/></w:rPr>" : "") +
            $"<w:t xml:space='preserve'>{text}</w:t></w:r></w:p>";
        var body = P("Amber receiver records weather.");
        if (change == "edit" && revised) body = P("Amber receiver records temperature.");
        if (change == "insert" && revised || change == "delete" && !revised) body += P("Violet sensor stores a reading.");
        if (change == "split" && revised || change == "merge" && !revised)
            body = P("Amber receiver ") + P("records weather.");
        if (table)
            body = "<w:tbl><w:tblPr><w:tblW w:w='4000' w:type='dxa'/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w='4000'/></w:tblGrid><w:tr><w:tc>" +
                "<w:tcPr><w:tcW w:w='4000' w:type='dxa'/></w:tcPr>" + body + "</w:tc></w:tr></w:tbl><w:p/>";
        var defaults = revised && change == "builtin-defaults" ? "" :
            "<w:docDefaults><w:rPrDefault><w:rPr>" +
            $"<w:rFonts w:ascii='{(revised ? "Liberation Serif" : "Liberation Mono")}'/>" +
            $"<w:sz w:val='{(revised ? "32" : "24")}'/></w:rPr></w:rPrDefault>" +
            $"<w:pPrDefault><w:pPr><w:spacing w:after='{(revised ? "160" : "120")}'/></w:pPr></w:pPrDefault></w:docDefaults>";
        var original = "<w:style w:type='paragraph' " + (!revised ? "w:default='1' " : "") +
            "w:styleId='ArchiveBody'><w:name w:val='Archive body'/><w:pPr><w:spacing w:after='200'/></w:pPr>" +
            "<w:rPr><w:b/><w:sz w:val='28'/></w:rPr></w:style>";
        var styles = defaults + (!revised || change == "retained-style" ? original : "") +
            "<w:style w:type='paragraph' w:styleId='PlainBody'><w:name w:val='Plain body'/></w:style>" +
            (change == "id-collision" ? "<w:style w:type='paragraph' w:styleId='DocxodusImplicitParagraph'><w:name w:val='Existing custom style'/></w:style>" : "");
        return IrTestDocuments.FromBodyAndStylesXml(body, styles);
    }

    private static string[] Presentation(WmlDocument document)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings
        {
            CreateHtmlConverterAnnotationAttributes = false,
        });
        using var memory = new MemoryStream(assembled.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        using var stream = package.MainDocumentPart!.GetStream();
        return XDocument.Load(stream).Descendants(W + "p").Select(paragraph =>
            ((string?)paragraph.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "after") ?? "0") + "|" +
            string.Concat(paragraph.Descendants(W + "r").SelectMany(run => run.Descendants(W + "t").SelectMany(text =>
                text.Value.Select(character =>
                    (string?)run.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii") + ":" +
                    (string?)run.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val") + ":" +
                    (run.Element(W + "rPr")?.Element(W + "b") is { } bold &&
                        (string?)bold.Attribute(W + "val") is not ("0" or "false" or "off")) + ":" + character))))).ToArray();
    }
}
