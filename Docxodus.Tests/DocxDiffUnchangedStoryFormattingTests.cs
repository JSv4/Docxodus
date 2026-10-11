using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffUnchangedStoryFormattingTests
{
    private static readonly XNamespace W = IrTestDocuments.W;
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Theory]
    [InlineData(true, "default", false)]
    [InlineData(true, "first", false)]
    [InlineData(true, "even", false)]
    [InlineData(false, "default", false)]
    [InlineData(false, "first", false)]
    [InlineData(false, "even", false)]
    [InlineData(true, "default", true)]
    [InlineData(false, "default", true)]
    public void DefaultRoleChange_ReachesUnchangedStories(bool header, string type, bool table)
    {
        var left = Create(false, header, type, table);
        var right = Create(true, header, type, table);
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
            var main = package.MainDocumentPart!;
            var reference = main.GetXDocument().Descendants(W + (header ? "headerReference" : "footerReference")).Single();
            var part = main.GetPartById((string)reference.Attribute(R + "id")!);
            Assert.Equal(header, part is HeaderPart);
            Assert.Equal(type, (string?)reference.Attribute(W + "type"));
        }
        using var resultMemory = new MemoryStream(redline.DocumentByteArray);
        using var resultPackage = WordprocessingDocument.Open(resultMemory, false);
        var resultPart = header ? (OpenXmlPart)resultPackage.MainDocumentPart!.HeaderParts.Single() :
            resultPackage.MainDocumentPart!.FooterParts.Single();
        var caption = resultPart.GetXDocument().Descendants(W + "p").Last();
        Assert.Equal("Caption", (string?)caption.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.Null(caption.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    private static WmlDocument Create(bool revised, bool header, string type, bool table)
    {
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Liberation Serif'/>" +
            $"<w:sz w:val='{(revised ? "32" : "24")}'/></w:rPr></w:rPrDefault>" +
            $"<w:pPrDefault><w:pPr><w:spacing w:after='{(revised ? "160" : "120")}'/></w:pPr></w:pPrDefault></w:docDefaults>" +
            $"<w:style w:type='paragraph' w:default='1' w:styleId='{(revised ? "NewBody" : "OldBody")}'><w:name w:val='Body'/></w:style>" +
            "<w:style w:type='paragraph' w:styleId='Caption'><w:name w:val='Caption'/><w:pPr><w:spacing w:after='60'/></w:pPr>" +
            "<w:rPr><w:rFonts w:ascii='Liberation Mono'/><w:sz w:val='18'/></w:rPr></w:style>";
        var document = IrTestDocuments.FromBodyAndStylesXml("<w:p><w:r><w:t>Copper station records a signal.</w:t></w:r></w:p>", styles);
        using var stream = new OpenXmlMemoryStreamDocument(document);
        using (var package = stream.GetWordprocessingDocument())
        {
            var main = package.MainDocumentPart!;
            var part = header ? (OpenXmlPart)main.AddNewPart<HeaderPart>("rIdStationStory") :
                main.AddNewPart<FooterPart>("rIdStationStory");
            var paragraph = XElement.Parse($"<w:p xmlns:w='{W}'><w:r><w:t>Amber station reports rainfall.</w:t></w:r></w:p>");
            var content = table ? XElement.Parse($"<w:tbl xmlns:w='{W}'><w:tblPr><w:tblW w:w='4000' w:type='dxa'/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w='4000'/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w='4000' w:type='dxa'/></w:tcPr>" +
                paragraph + "</w:tc></w:tr></w:tbl>") : paragraph;
            part.GetXDocument().Add(new XElement(W + (header ? "hdr" : "ftr"), content,
                XElement.Parse($"<w:p xmlns:w='{W}'><w:pPr><w:pStyle w:val='Caption'/></w:pPr><w:r><w:t>Fixed station label.</w:t></w:r></w:p>")));
            part.PutXDocument();
            var body = main.GetXDocument().Root!.Element(W + "body")!;
            var section = body.Element(W + "sectPr");
            if (section is null) { section = new XElement(W + "sectPr"); body.Add(section); }
            section.AddFirst(new XElement(W + (header ? "headerReference" : "footerReference"),
                new XAttribute(W + "type", type), new XAttribute(R + "id", "rIdStationStory")));
            if (type == "first") section.Add(new XElement(W + "titlePg"));
            if (type == "even")
            {
                var settings = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
                var xml = settings.GetXDocument();
                if (xml.Root is null) xml.Add(new XElement(W + "settings"));
                xml.Root!.Add(new XElement(W + "evenAndOddHeaders"));
            }
            main.DocumentSettingsPart?.PutXDocument();
            main.PutXDocument();
        }
        return stream.GetModifiedWmlDocument();
    }

    private static string[] Presentation(WmlDocument document)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings { CreateHtmlConverterAnnotationAttributes = false });
        using var memory = new MemoryStream(assembled.DocumentByteArray);
        using var package = WordprocessingDocument.Open(memory, false);
        return package.ContentParts().SelectMany(part => part.GetXDocument().Descendants(W + "p")).Select(paragraph =>
            (string?)paragraph.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "after") + "|" +
            string.Concat(paragraph.Descendants(W + "r").SelectMany(run => run.Descendants(W + "t").SelectMany(text =>
                text.Value.Select(character => (string?)run.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val") + ":" + character))))).ToArray();
    }
}
