using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffConsumerFormattingTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("tab-stop", false)]
    [InlineData("tab-stop", true)]
    [InlineData("tracking", false)]
    [InlineData("compatibility", false)]
    [InlineData("absent-original", false)]
    [InlineData("absent-revised", false)]
    public void TableDefaults_RemainReversibleWhenSettingsDiffer(string settingsChange, bool directOverride)
    {
        var left = CreateDocument(false, settingsChange, directOverride);
        var right = CreateDocument(true, settingsChange, directOverride);
        var redline = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);

        Assert.Equal(Presentation(right), Presentation(accepted));
        Assert.Equal(Presentation(left), Presentation(rejected));
        foreach (var document in new[] { left, right, redline, accepted, rejected })
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
    }

    private static WmlDocument CreateDocument(bool revised, string settingsChange, bool directOverride)
    {
        var direct = directOverride ? "<w:pPr><w:spacing w:after='80'/></w:pPr>" : "";
        var run = directOverride ? "<w:rPr><w:sz w:val='18'/></w:rPr>" : "";
        var body = "<w:tbl><w:tblPr><w:tblW w:w='4000' w:type='dxa'/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w='4000'/></w:tblGrid><w:tr><w:tc>" +
            "<w:tcPr><w:tcW w:w='4000' w:type='dxa'/></w:tcPr><w:p>" + direct +
            "<w:r>" + run + "<w:t>Copper sensor records sunlight.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:t>Jade receiver stores readings.</w:t></w:r></w:p>";
        var size = revised ? "32" : "24";
        var after = revised ? "160" : "120";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr>" +
            "<w:rFonts w:ascii='Liberation Serif' w:hAnsi='Liberation Serif'/>" +
            $"<w:sz w:val='{size}'/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" +
            $"<w:spacing w:after='{after}'/></w:pPr></w:pPrDefault></w:docDefaults>" +
            "<w:style w:type='paragraph' w:default='1' w:styleId='Body'><w:name w:val='Body'/></w:style>";
        using var memory = new MemoryStream();
        using (var package = WordprocessingDocument.Create(memory, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var stream = main.GetStream())
                XDocument.Parse($"<w:document xmlns:w='{W}'><w:body>{body}</w:body></w:document>").Save(stream);
            using (var stream = main.AddNewPart<StyleDefinitionsPart>().GetStream())
                XDocument.Parse($"<w:styles xmlns:w='{W}'>{styles}</w:styles>").Save(stream);
            var settings = settingsChange switch
            {
                "tracking" => revised ? "<w:trackRevisions/>" : "",
                "compatibility" => revised ? "<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>" : "<w:compat/>",
                _ => $"<w:defaultTabStop w:val='{(revised ? "480" : "720")}'/>",
            };
            if (!(settingsChange == "absent-original" && !revised || settingsChange == "absent-revised" && revised))
            {
                using var stream = main.AddNewPart<DocumentSettingsPart>().GetStream();
                XDocument.Parse($"<w:settings xmlns:w='{W}'>{settings}</w:settings>").Save(stream);
            }
        }
        return new WmlDocument("synthetic-consumer.docx", memory.ToArray());
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
        var root = XDocument.Load(stream);
        return root.Descendants(W + "p").Select(paragraph =>
            (string?)paragraph.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "after") + "|" +
            string.Concat(paragraph.Descendants(W + "r").Select(run =>
                (string?)run.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val") + ":" +
                string.Concat(run.Descendants(W + "t").Select(text => text.Value))))).ToArray();
    }
}
