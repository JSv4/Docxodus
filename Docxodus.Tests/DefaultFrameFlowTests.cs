using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class DefaultFrameFlowTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static byte[] Fixture(string frame, bool direct, bool implicitStyle)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var styles = new Wp.Styles();
            styles.InnerXml = $"<w:style xmlns:w=\"{WNs}\" w:type=\"paragraph\" w:styleId=\"TernBase\">" +
                $"<w:name w:val=\"Tern Base\"/><w:pPr><w:framePr {frame}/></w:pPr></w:style>" +
                $"<w:style xmlns:w=\"{WNs}\" w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Tern\">" +
                "<w:name w:val=\"Tern\"/><w:basedOn w:val=\"TernBase\"/></w:style>";
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var paragraph = new Wp.Paragraph(new Wp.Run(new Wp.Text("Tern amber sail")));
            if (direct || !implicitStyle)
            {
                paragraph.PrependChild(new Wp.ParagraphProperties());
                paragraph.ParagraphProperties!.InnerXml =
                    (implicitStyle ? "" : $"<w:pStyle xmlns:w=\"{WNs}\" w:val=\"Tern\"/>") +
                    (direct ? $"<w:framePr xmlns:w=\"{WNs}\"/>" : "");
            }
            main.Document = new Wp.Document(new Wp.Body(paragraph, new Wp.SectionProperties(
                new Wp.PageSize { Width = 12240, Height = 15840 },
                new Wp.PageMargin { Top = 1440, Left = 1440, Bottom = 1440, Right = 1440 })));
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("", true)]
    [InlineData("w:wrap=\"auto\"", false)]
    [InlineData("w:wrap=\"auto\"", true)]
    public void InheritedFrameDefaults_DoNotCreateAFloatingStory(string frame, bool implicitStyle)
    {
        AssertFrame(Fixture(frame, false, implicitStyle), false);
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("w:wrap=\"auto\"", true)]
    [InlineData("w:w=\"5000\" w:x=\"300\" w:y=\"200\"", false)]
    [InlineData("w:w=\"5000\" w:x=\"300\" w:y=\"200\"", true)]
    [InlineData("w:wrap=\"around\"", false)]
    public void DeclaredFramesAndNondefaultStyleProperties_RemainFrames(string frame, bool direct)
    {
        AssertFrame(Fixture(frame, direct, false), true);
    }

    [Fact]
    public void AssembledFrameDefaults_RemainInactiveWhenRenderedAgain()
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("w:wrap=\"auto\"", false, false));
        using (var doc = WordprocessingDocument.Open(stream, true))
            FormattingAssembler.AssembleFormatting(doc, new FormattingAssemblerSettings
            {
                ClearStyles = false,
                RemoveStyleNamesFromParagraphAndRunProperties = false,
            });
        AssertFrame(stream.ToArray(), false);
    }

    private static void AssertFrame(byte[] bytes, bool expected)
    {
        using (var stream = new MemoryStream(bytes))
        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var main = doc.MainDocumentPart!;
            var paragraph = main.GetXDocument().Descendants(W.p).Single();
            var before = paragraph.ToString(SaveOptions.DisableFormatting);
            var stylesPart = main.StyleDefinitionsPart!;
            var stylesBefore = stylesPart.GetXDocument().ToString();
            var properties = FormattingAssembler.ResolveEffectiveParagraphProperties(doc, paragraph);
            // Effective formatting still reports the style's source properties. Only
            // layout decides whether these defaults activate a separate text story.
            Assert.NotNull(properties.Element(W.framePr));
            Assert.Equal(before, paragraph.ToString(SaveOptions.DisableFormatting));
            Assert.Equal(stylesBefore, stylesPart.GetXDocument().ToString());
        }
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("tern-frame.docx", bytes),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = PaginationMode.Paginated });
        var content = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "Tern amber sail");
        Assert.Equal(expected, content.Ancestors().Any(e => e.Attribute("data-docx-text-frame") != null));
    }
}
