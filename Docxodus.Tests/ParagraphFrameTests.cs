using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class ParagraphFrameTests
{
    private static byte[] Fixture(string story, bool framed, bool multiple = false, string? heightRule = null,
        string? inlineWrap = null, string frameAttributes = "")
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(
                new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" },
                    new Wp.FontSize { Val = "24" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var first = new Wp.Paragraph(new Wp.Run(new Wp.Text("CEDAR RECORD 731")));
            if (framed)
            {
                first.PrependChild(new Wp.ParagraphProperties());
                var height = heightRule == null ? "" : " w:h=\"400\"" +
                    (heightRule == "omitted" ? "" : " w:hRule=\"" + heightRule + "\"");
                var placement = inlineWrap == null
                    ? "w:w=\"10080\" w:hAnchor=\"page\" w:vAnchor=\"text\" w:x=\"840\" w:y=\"20\" w:wrap=\"auto\""
                    : "w:w=\"3600\" w:hAnchor=\"text\" w:vAnchor=\"text\" w:x=\"240\" w:yAlign=\"inline\" w:wrap=\"" + inlineWrap + "\"";
                first.ParagraphProperties!.InnerXml = "<w:framePr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " + placement + height + frameAttributes + "/>";
            }
            var second = new Wp.Paragraph(new Wp.ParagraphProperties(new Wp.Justification { Val = Wp.JustificationValues.Center }),
                new Wp.Run(new Wp.Text("Cedar page control")));
            var section = new Wp.SectionProperties(new Wp.PageSize { Width = 12240, Height = 15840 },
                new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 });
            main.Document = new Wp.Document(new Wp.Body());
            if (story == "footer")
            {
                var footer = main.AddNewPart<FooterPart>();
                footer.Footer = new Wp.Footer(first, second);
                section.PrependChild(new Wp.FooterReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
            }
            else if (story == "header")
            {
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new Wp.Header(first, second);
                section.PrependChild(new Wp.HeaderReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(header) });
            }
            else main.Document.Body!.Append(first, second);
            if (story != "body") main.Document.Body!.Append(new Wp.Paragraph(new Wp.Run(new Wp.Text("Cedar body control"))));
            if (multiple)
                first.InsertAfterSelf(new Wp.Paragraph(first.ParagraphProperties!.CloneNode(true),
                    new Wp.Run(new Wp.Text("CEDAR FRAME SECOND"))));
            main.Document.Body!.Append(section);
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("footer", false)]
    [InlineData("footer", true)]
    [InlineData("header", true)]
    [InlineData("body", true)]
    public void FrameOffsetsAndWidth_ArePreservedOutsideNormalFlow(string story, bool paginated)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-frame.docx", Fixture(story, true)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true,
                RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None });
        var paragraph = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "CEDAR RECORD 731");
        var frame = Assert.Single(paragraph.Ancestors().Where(e => (string?)e.Attribute("data-docx-text-frame") == "true"));
        var style = (string?)frame.Attribute("style") ?? "";
        Assert.Contains("position: absolute", style);
        Assert.Contains("width: 504pt", style);
        Assert.Contains("left: 42pt", style);
        Assert.Contains("top: 1pt", style);
        var control = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "Cedar page control");
        Assert.DoesNotContain("position: absolute", (string?)control.Attribute("style") ?? "");
    }

    [Theory]
    [InlineData("footer")]
    [InlineData("body")]
    public void AdjacentParagraphsWithIdenticalProperties_ShareOneFrame(string story)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-shared-frame.docx", Fixture(story, true, true)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true,
                RenderPagination = PaginationMode.Paginated });
        var first = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "CEDAR RECORD 731");
        var second = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "CEDAR FRAME SECOND");
        var frame = Assert.Single(first.Ancestors().Where(e =>
            ((string?)e.Attribute("style") ?? "").Contains("position: absolute")));
        Assert.Contains(second, frame.Descendants());
        Assert.DoesNotContain("Cedar page control", frame.Value);
        Assert.DoesNotContain("position: absolute", (string?)second.Attribute("style") ?? "");
    }

    [Theory]
    [InlineData("auto", null)]
    [InlineData("exact", "height")]
    [InlineData("atLeast", "min-height")]
    [InlineData("omitted", "min-height")]
    public void FrameHeight_RespectsItsRule(string rule, string? property)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-height-frame.docx", Fixture("footer", true, heightRule: rule)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true });
        var frame = html.Descendants().Single(e => (string?)e.Attribute("data-docx-text-frame") == "true");
        var heights = ((string?)frame.Attribute("style") ?? "").Split(';').Select(d => d.Split(':', 2))
            .Where(parts => parts.Length == 2 && parts[0].Trim() is "height" or "min-height").ToArray();
        if (property == null) Assert.Empty(heights);
        else
        {
            var declaration = Assert.Single(heights);
            Assert.Equal(property, declaration[0].Trim());
            Assert.Equal("20pt", declaration[1].Trim());
        }
    }

    [Fact]
    public void OrdinaryFooter_RemainsInFlow()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-flow.docx", Fixture("footer", false)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true,
                RenderPagination = PaginationMode.Paginated });
        var paragraph = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "CEDAR RECORD 731");
        Assert.DoesNotContain("position: absolute", (string?)paragraph.Attribute("style") ?? "");
    }

    [Theory]
    [InlineData("body", "notBeside", true, false)]
    [InlineData("header", "notBeside", true, false)]
    [InlineData("footer", "notBeside", true, false)]
    [InlineData("body", "around", false, true)]
    [InlineData("body", "auto", false, true)]
    [InlineData("body", "tight", false, true)]
    [InlineData("body", "through", false, true)]
    [InlineData("body", "none", false, false)]
    public void InlineFrame_PreservesOneGroupedStoryAndItsWrappingContract(
        string story, string wrap, bool inFlow, bool excludes)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("inline-frame.docx",
            Fixture(story, true, multiple: true, inlineWrap: wrap)), new WmlToHtmlConverterSettings
            { FabricateCssClasses = false, RenderHeadersAndFooters = true, RenderPagination = PaginationMode.Paginated });
        var frame = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-text-frame") == "true");
        Assert.Contains("CEDAR RECORD 731", frame.Value);
        Assert.Contains("CEDAR FRAME SECOND", frame.Value);
        Assert.DoesNotContain("Cedar page control", frame.Value);
        var slot = Assert.Single(frame.Ancestors(), e => (string?)e.Attribute("data-docx-frame-container") == "true");
        Assert.Equal(inFlow, slot.Attribute("data-docx-frame-flow") != null);
        Assert.Equal(excludes, frame.Ancestors().Any(e => e.Attribute("data-docx-frame-exclusion") != null));
        Assert.Equal(!inFlow, slot.Attribute("data-keep-with-next") != null);
        Assert.Contains(inFlow || excludes ? "position: relative" : "position: absolute", (string?)frame.Attribute("style"));
    }

    [Theory]
    [InlineData("around", " w:xAlign=\"center\"", true)]
    [InlineData("around", " w:xAlign=\"right\"", true)]
    [InlineData("around", " w:vSpace=\"80\"", true)]
    [InlineData("notBeside", " w:vSpace=\"80\"", true)]
    [InlineData("notBeside", " w:xAlign=\"right\"", false)]
    public void InlineFrame_UnsupportedWrappingPreservesContentWithADiagnostic(string wrap, string attributes, bool approximate)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("inline-frame.docx",
            Fixture("body", true, multiple: true, inlineWrap: wrap, frameAttributes: attributes)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = PaginationMode.Paginated,
                RenderUnsupportedContentPlaceholders = true, IncludeUnsupportedContentMetadata = true });
        var frame = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-text-frame") == "true");
        Assert.Contains("CEDAR RECORD 731", frame.Value);
        Assert.Contains("CEDAR FRAME SECOND", frame.Value);
        var pending = wrap == "around" && !attributes.Contains("vSpace");
        Assert.Contains(frame.Ancestors(), e => e.Attribute(pending ? "data-docx-frame-exclusion" : "data-docx-frame-flow") != null);
        Assert.Equal(pending, frame.Attribute("data-docx-frame-pending") != null);
        Assert.Equal(approximate ? "partial" : null, (string?)frame.Attribute("data-docx-shape"));
        Assert.Equal(approximate ? "Other" : null, (string?)frame.Attribute("data-content-type"));
        Assert.Equal(approximate ? "framePr" : null, (string?)frame.Attribute("data-element-name"));
    }

    [Theory]
    [InlineData("", true)]
    [InlineData(" w:xAlign=\"left\"", false)]
    public void InlineFrame_InteriorOffsetReportsConservativeWrapping(string attributes, bool approximate)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("inline-frame.docx",
            Fixture("body", true, inlineWrap: "around", frameAttributes: attributes)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = PaginationMode.Paginated,
                RenderUnsupportedContentPlaceholders = true, IncludeUnsupportedContentMetadata = true });
        var frame = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-text-frame") == "true");
        Assert.Contains(frame.Ancestors(), e => e.Attribute("data-docx-frame-exclusion") != null);
        Assert.Equal(approximate ? "partial" : null, (string?)frame.Attribute("data-docx-shape"));
        Assert.Contains(approximate ? "margin-left: 12pt;" : "margin-left: 0;", (string?)frame.Attribute("style"));
    }
}
