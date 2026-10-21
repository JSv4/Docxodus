using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

/// <summary>
/// A paragraph's own <c>w:framePr</c> overrides its style's frame attribute by attribute, so a
/// paragraph that restates only <c>w:wrap</c> keeps the style's width, anchors and offsets.
/// </summary>
public class FramePrStyleMergeTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StyleFrame =
        "<w:framePr w:w=\"8640\" w:wrap=\"auto\" w:vAnchor=\"text\" w:hAnchor=\"page\" w:x=\"1008\" w:y=\"40\"/>";

    // directFrame: the stamp paragraph's own w:framePr attributes, or null for none.
    private static byte[] Fixture(string? directFrame)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var styles = new Wp.Styles();
            styles.InnerXml =
                $"<w:style xmlns:w=\"{WNs}\" w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"StampFrame\">" +
                $"<w:name w:val=\"Stamp Frame\"/><w:pPr>{StyleFrame.Replace("<w:framePr", $"<w:framePr xmlns:w=\"{WNs}\"")}</w:pPr></w:style>";
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var stamp = new Wp.Paragraph(new Wp.ParagraphProperties(), new Wp.Run(new Wp.Text("REF 55-0912 rev3")));
            stamp.ParagraphProperties!.InnerXml = $"<w:pStyle xmlns:w=\"{WNs}\" w:val=\"StampFrame\"/>" +
                (directFrame == null ? "" : $"<w:framePr xmlns:w=\"{WNs}\" {directFrame}/>");
            var leaf = new Wp.Paragraph(new Wp.ParagraphProperties(new Wp.Justification { Val = Wp.JustificationValues.Center }),
                new Wp.Run(new Wp.Text("Leaf 4")));
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Wp.Footer(stamp, leaf);
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Body text"))),
                new Wp.SectionProperties(
                    new Wp.FooterReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) },
                    new Wp.PageSize { Width = 12240, Height = 15840 },
                    new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 })));
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    private static XElement EffectiveFrame(string? directFrame)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture(directFrame));
        using var doc = WordprocessingDocument.Open(stream, false);
        var stamp = doc.MainDocumentPart!.FooterParts.Single().GetXDocument().Descendants(W.p).First();
        var pPr = FormattingAssembler.ResolveEffectiveParagraphProperties(doc, stamp);
        return Assert.Single(pPr.Elements(W.framePr));
    }

    private static string? Attr(XElement e, XName name) => (string?)e.Attribute(name);

    [Fact]
    public void DirectFramePr_KeepsTheStyleAttributesItDoesNotSet()
    {
        var frame = EffectiveFrame("w:wrap=\"auto\"");
        Assert.Equal("8640", Attr(frame, W._w));
        Assert.Equal("page", Attr(frame, W.hAnchor));
        Assert.Equal("text", Attr(frame, W.vAnchor));
        Assert.Equal("1008", Attr(frame, W.x));
        Assert.Equal("40", Attr(frame, W.y));
        Assert.Equal("auto", Attr(frame, W.wrap));
    }

    [Fact]
    public void DirectFramePr_StillOverridesTheAttributesItSets()
    {
        var frame = EffectiveFrame("w:wrap=\"around\" w:x=\"2000\"");
        Assert.Equal("around", Attr(frame, W.wrap));
        Assert.Equal("2000", Attr(frame, W.x));
        Assert.Equal("8640", Attr(frame, W._w));
        Assert.Equal("40", Attr(frame, W.y));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("w:wrap=\"auto\"")]
    public void StyleDefinedFooterFrame_KeepsItsWidthAndPosition(string? directFrame)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("stamp-frame.docx", Fixture(directFrame)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true,
                RenderPagination = PaginationMode.Paginated });
        var stamp = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "REF 55-0912 rev3");
        var frame = Assert.Single(stamp.Ancestors(), e => (string?)e.Attribute("data-docx-text-frame") == "true");
        var style = (string?)frame.Attribute("style") ?? "";
        Assert.Contains("width: 432pt", style);
        Assert.Contains("left: 50.4pt", style);
        Assert.Contains("top: 2pt", style);
        Assert.Equal("page", (string?)frame.Attribute("data-docx-frame-h-anchor"));
        var leaf = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "Leaf 4");
        Assert.DoesNotContain(leaf, frame.Descendants());
    }
}
