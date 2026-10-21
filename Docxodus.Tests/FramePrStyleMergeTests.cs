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
    private static byte[] Fixture(string? directFrame, string styleFrame = StyleFrame)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var styles = new Wp.Styles();
            styles.InnerXml =
                $"<w:style xmlns:w=\"{WNs}\" w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"StampFrame\">" +
                $"<w:name w:val=\"Stamp Frame\"/><w:pPr>{styleFrame.Replace("<w:framePr", $"<w:framePr xmlns:w=\"{WNs}\"")}</w:pPr></w:style>";
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

    private static XElement EffectiveFrame(string? directFrame, string styleFrame = StyleFrame)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture(directFrame, styleFrame));
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

    // Style frames that place the stamp by alignment, or as a drop cap, on a page-anchored 2in frame.
    private const string CenteredStyleFrame =
        "<w:framePr w:w=\"2880\" w:wrap=\"around\" w:vAnchor=\"page\" w:hAnchor=\"page\" w:xAlign=\"center\" w:y=\"4000\"/>";
    private const string MiddleStyleFrame =
        "<w:framePr w:w=\"2880\" w:wrap=\"around\" w:vAnchor=\"page\" w:hAnchor=\"page\" w:x=\"2000\" w:yAlign=\"center\"/>";
    private const string OffsetStyleFrame =
        "<w:framePr w:w=\"2880\" w:wrap=\"around\" w:vAnchor=\"page\" w:hAnchor=\"page\" w:x=\"2000\" w:y=\"4000\"/>";
    private const string DropCapStyleFrame =
        "<w:framePr w:dropCap=\"drop\" w:lines=\"3\" w:wrap=\"around\" w:vAnchor=\"text\" w:hAnchor=\"text\"/>";
    private const string PageFrame = "w:wrap=\"around\" w:hAnchor=\"page\" w:x=\"4000\" w:vAnchor=\"page\" w:y=\"4000\"";

    // The merged frame's attributes, sorted, as "name=value" pairs.
    private static string FrameAttributes(string directFrame, string styleFrame) =>
        string.Join(" ", EffectiveFrame(directFrame, styleFrame).Attributes().Where(a => !a.IsNamespaceDeclaration)
            .Select(a => $"{a.Name.LocalName}={a.Value}").OrderBy(a => a, System.StringComparer.Ordinal));

    // x/xAlign and y/yAlign are one setting per axis: a paragraph that sets either replaces both. The
    // anchors are separate settings, and a style's drop cap never reaches a paragraph with its own frame.
    // LibreOffice 25.8.7.3 renders each case this way (docs/ooxml_corner_cases.md).
    [Theory]
    [InlineData(CenteredStyleFrame, "w:wrap=\"around\" w:x=\"2000\"",
        "hAnchor=page vAnchor=page w=2880 wrap=around x=2000 y=4000")]
    [InlineData(CenteredStyleFrame, "w:w=\"2880\" w:wrap=\"around\" w:vAnchor=\"page\" w:hAnchor=\"page\" w:x=\"2000\" w:y=\"4000\"",
        "hAnchor=page vAnchor=page w=2880 wrap=around x=2000 y=4000")]
    [InlineData(MiddleStyleFrame, "w:wrap=\"around\" w:y=\"4000\"",
        "hAnchor=page vAnchor=page w=2880 wrap=around x=2000 y=4000")]
    [InlineData(OffsetStyleFrame, "w:wrap=\"around\" w:xAlign=\"right\"",
        "hAnchor=page vAnchor=page w=2880 wrap=around xAlign=right y=4000")]
    [InlineData(OffsetStyleFrame, "w:wrap=\"around\" w:hAnchor=\"margin\"",
        "hAnchor=margin vAnchor=page w=2880 wrap=around x=2000 y=4000")]
    [InlineData(DropCapStyleFrame, PageFrame, "hAnchor=page vAnchor=page wrap=around x=4000 y=4000")]
    [InlineData(DropCapStyleFrame, "w:lines=\"2\"", "hAnchor=text lines=2 vAnchor=text wrap=around")]
    public void DirectFramePr_ReplacesWholeAxesAndNeverInheritsADropCap(string styleFrame, string directFrame, string expected) =>
        Assert.Equal(expected, FrameAttributes(directFrame, styleFrame));

    private static XElement RenderedFrame(string? directFrame, string styleFrame = StyleFrame)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("stamp-frame.docx", Fixture(directFrame, styleFrame)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true,
                RenderPagination = PaginationMode.Paginated });
        var stamp = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "REF 55-0912 rev3");
        return Assert.Single(stamp.Ancestors(), e => (string?)e.Attribute("data-docx-text-frame") == "true");
    }

    [Theory]
    [InlineData(CenteredStyleFrame, "w:wrap=\"around\" w:x=\"2000\"")]
    [InlineData(CenteredStyleFrame, "w:w=\"2880\" w:wrap=\"around\" w:vAnchor=\"page\" w:hAnchor=\"page\" w:x=\"2000\" w:y=\"4000\"")]
    [InlineData(MiddleStyleFrame, "w:wrap=\"around\" w:y=\"4000\"")]
    public void DirectOffset_IsNotRealignedByTheStyle(string styleFrame, string directFrame)
    {
        var frame = RenderedFrame(directFrame, styleFrame);
        Assert.Contains("left: 100pt", (string?)frame.Attribute("style"));
        Assert.Contains("top: 200pt", (string?)frame.Attribute("style"));
        Assert.Null(frame.Attribute("data-docx-frame-xalign"));
        Assert.Null(frame.Attribute("data-docx-frame-yalign"));
    }

    [Fact]
    public void StyleDropCap_DoesNotSuppressADirectPageFrame()
    {
        var frame = RenderedFrame(PageFrame, DropCapStyleFrame);
        Assert.Contains("left: 200pt", (string?)frame.Attribute("style"));
        Assert.Contains("top: 200pt", (string?)frame.Attribute("style"));
        Assert.Equal("page", (string?)frame.Attribute("data-docx-frame-h-anchor"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("w:wrap=\"auto\"")]
    public void StyleDefinedFooterFrame_KeepsItsWidthAndPosition(string? directFrame)
    {
        var frame = RenderedFrame(directFrame);
        var html = frame.AncestorsAndSelf().Last();
        var style = (string?)frame.Attribute("style") ?? "";
        Assert.Contains("width: 432pt", style);
        Assert.Contains("left: 50.4pt", style);
        Assert.Contains("top: 2pt", style);
        Assert.Equal("page", (string?)frame.Attribute("data-docx-frame-h-anchor"));
        var leaf = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == "Leaf 4");
        Assert.DoesNotContain(leaf, frame.Descendants());
    }
}
