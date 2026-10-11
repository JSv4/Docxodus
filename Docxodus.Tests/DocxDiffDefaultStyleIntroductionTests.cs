using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffDefaultStyleIntroductionTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("unchanged", false)]
    [InlineData("unchanged", true)]
    [InlineData("edit", false)]
    [InlineData("insert", false)]
    [InlineData("delete", false)]
    [InlineData("split", false)]
    [InlineData("merge", false)]
    [InlineData("override", true)]
    public void IntroducingDefaultStyle_PreservesBothFormattingEndpoints(string change, bool table)
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
        var body = P("Copper sensor records sunlight.");
        if (change == "edit" && revised) body = P("Copper sensor records moonlight.");
        if (change == "insert" && revised || change == "delete" && !revised) body += P("Jade receiver stores readings.");
        if (change == "split" && revised || change == "merge" && !revised)
            body = P("Copper sensor ") + P("records sunlight.");
        if (table)
            body = "<w:tbl><w:tblPr><w:tblW w:w='4000' w:type='dxa'/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w='4000'/></w:tblGrid><w:tr><w:tc>" +
                "<w:tcPr><w:tcW w:w='4000' w:type='dxa'/></w:tcPr>" + body + "</w:tc></w:tr></w:tbl><w:p/>";
        var size = revised ? "32" : "24";
        var after = revised ? "160" : "120";
        var font = revised ? "Liberation Mono" : "Liberation Serif";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr>" +
            $"<w:rFonts w:ascii='{font}' w:hAnsi='{font}'/>" +
            $"<w:sz w:val='{size}'/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" +
            $"<w:spacing w:after='{after}'/></w:pPr></w:pPrDefault></w:docDefaults>" +
            "<w:style w:type='paragraph' w:styleId='Body'><w:name w:val='Body'/></w:style>" +
            (revised ? "<w:style w:type='paragraph' w:default='1' w:styleId='FreshBody'><w:name w:val='Fresh body'/></w:style>" : "");
        using var memory = new MemoryStream();
        using (var package = WordprocessingDocument.Create(memory, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var stream = main.GetStream())
                XDocument.Parse($"<w:document xmlns:w='{W}'><w:body>{body}</w:body></w:document>").Save(stream);
            using (var stream = main.AddNewPart<StyleDefinitionsPart>().GetStream())
                XDocument.Parse($"<w:styles xmlns:w='{W}'>{styles}</w:styles>").Save(stream);
        }
        return new WmlDocument("synthetic-style-introduction.docx", memory.ToArray());
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
            (string?)paragraph.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "after") + "|" +
            string.Concat(paragraph.Descendants(W + "r").SelectMany(run => run.Descendants(W + "t").SelectMany(text =>
                text.Value.Select(character => (string?)run.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii") + ":" +
                    (string?)run.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val") + ":" + character))))).ToArray();
    }
}
