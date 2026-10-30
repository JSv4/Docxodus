using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class TextBoxClipProvenanceTests
{
    [Theory]
    [InlineData("drawing", "fixed", true)]
    [InlineData("drawing", "implicit", false)]
    [InlineData("drawing", "overflow", false)]
    [InlineData("drawing", "ellipsis", false)]
    [InlineData("drawing", "clip", true)]
    [InlineData("drawing", "auto", false)]
    [InlineData("drawing", "shrink", false)]
    [InlineData("drawing", "linked", false)]
    [InlineData("drawing", "zero-height", false)]
    [InlineData("vml", "fixed", true)]
    [InlineData("vml", "auto", false)]
    [InlineData("vml", "spaced-auto", false)]
    [InlineData("vml", "shrink", false)]
    [InlineData("vml", "linked-start", false)]
    [InlineData("vml", "linked-end", false)]
    [InlineData("vml", "type-auto", false)]
    [InlineData("vml", "zero-height", false)]
    public void Only_proven_fixed_unlinked_boxes_can_claim_source_clipping(string kind, string mode, bool eligible)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("quail-clip-provenance.docx", Fixture(kind, mode)),
            new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated, FabricateCssClasses = false });
        var box = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-textbox") == "true" && e.Value.Contains("Quail marker"));
        Assert.Equal(eligible ? "true" : null, (string?)box.Attribute("data-docx-textbox-source-clip"));
        if (mode is "auto" or "spaced-auto")
            Assert.DoesNotContain(((string?)box.Attribute("style") ?? "").Split(';'),
                declaration => declaration.TrimStart().StartsWith("height:", System.StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("implicit", true)]
    [InlineData("fixed", true)]
    [InlineData("auto", false)]
    [InlineData("shrink", false)]
    [InlineData("linked", false)]
    public void Supported_placement_can_be_proven_without_claiming_implicit_internal_clipping(string mode, bool eligible)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("drawing", mode));
        using var document = WordprocessingDocument.Open(stream, true);
        var part = document.MainDocumentPart!;
        var source = part.GetXDocument();
        foreach (var inline in source.Descendants(WP.inline).ToList())
        {
            inline.Name = WP.anchor;
            inline.SetAttributeValue("behindDoc", true);
            inline.AddFirst(new XElement(WP.positionH, new XAttribute("relativeFrom", "page"), new XElement(WP.posOffset, -2794000)),
                new XElement(WP.positionV, new XAttribute("relativeFrom", "page"), new XElement(WP.posOffset, 762000)));
            inline.Add(new XElement(WP.wrapNone));
        }
        part.PutXDocument();
        var html = WmlToHtmlConverter.ConvertToHtml(document,
            new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated, FabricateCssClasses = false });
        var box = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-textbox") == "true" && e.Value.Contains("Quail marker"));
        Assert.Equal(eligible ? "true" : null, (string?)box.Attribute("data-docx-textbox-source-position"));
        if (mode == "implicit") Assert.Null(box.Attribute("data-docx-textbox-source-clip"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Vml_clipping_metadata_is_rechecked_after_source_edits(bool inheritedFit)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("vml", "fixed"));
        using var document = WordprocessingDocument.Open(stream, true);
        var settings = new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated };
        if (inheritedFit)
        {
            var part = document.MainDocumentPart!;
            var source = part.GetXDocument();
            source.Descendants(VML.shape).First().SetAttributeValue("type", "#quail-fit-type");
            source.Descendants(W.pict).First().AddFirst(new XElement(VML.shapetype,
                new XAttribute("id", "quail-fit-type"), new XElement(VML.textbox, new XAttribute("style", "mso-fit-shape-to-text:false"))));
            part.PutXDocument();
        }
        for (var pass = 0; pass < 3; pass++)
        {
            var html = WmlToHtmlConverter.ConvertToHtml(document, settings);
            var box = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-textbox") == "true" && e.Value.Contains("Quail marker"));
            Assert.Equal(pass == 1 ? null : "true", (string?)box.Attribute("data-docx-textbox-source-clip"));
            var part = document.MainDocumentPart!;
            var source = part.GetXDocument();
            if (inheritedFit)
                source.Descendants(VML.shapetype).Elements(VML.textbox).Single()
                    .SetAttributeValue("style", pass == 0 ? "mso-fit-shape-to-text:true" : "mso-fit-shape-to-text:false");
            else if (pass == 0)
                source.Descendants(W.pict).First().Add(new XElement(VML.shape, new XAttribute("id", "quail-link-source"),
                    new XElement(VML.textbox, new XAttribute("style", "mso-next-textbox:#quail-box"))));
            else
                source.Descendants(VML.shape).Where(e => (string?)e.Attribute("id") == "quail-link-source").Remove();
            part.PutXDocument();
        }
    }

    private static byte[] Fixture(string kind, string mode)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            const string content = "<w:txbxContent><w:p><w:r><w:t>Quail marker</w:t></w:r></w:p></w:txbxContent>";
            var height = mode == "zero-height" ? 0 : 355600;
            var bodyPr = mode is "fixed" or "linked" or "zero-height" or "overflow" or "ellipsis" or "clip" ? "<a:noAutofit/>" : mode == "auto" ? "<a:spAutoFit/>"
                : mode == "shrink" ? "<a:normAutofit/>" : "";
            string Drawing(string body) => "<w:r><w:drawing><wp:inline><wp:extent cx=\"2032000\" cy=\"" + height + "\"/>" +
                "<wp:docPr id=\"1\" name=\"Quail window\"/><a:graphic><a:graphicData uri=\"" + WPS.wps + "\"><wps:wsp>" +
                "<wps:cNvSpPr txBox=\"1\"/><wps:spPr/><wps:txbx id=\"9\">" + body + "</wps:txbx>" +
                "<wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"" +
                (mode is "overflow" or "ellipsis" or "clip" ? " vertOverflow=\"" + mode + "\"" : "") + ">" + bodyPr + "</wps:bodyPr>" +
                "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
            var fit = mode == "auto" ? "mso-fit-shape-to-text:true" : mode == "spaced-auto" ? "mso-fit-shape-to-text: true"
                : mode == "shrink" ? "mso-fit-text-to-shape:true" : mode == "linked-start" ? "mso-next-textbox:#quail-next" : "";
            var legacy = "<w:r><w:pict>" + (mode == "type-auto" ?
                "<v:shapetype id=\"quail-type\"><v:textbox style=\"mso-fit-shape-to-text:true\"/></v:shapetype>" : "") +
                "<v:shape" + (mode == "type-auto" ? " type=\"#quail-type\"" : "") + " style=\"width:160pt;height:" +
                (mode == "zero-height" ? "0" : "28") + "pt\"><v:textbox id=\"quail-box\" style=\"" + fit + "\" inset=\"0,0,0,0\">" +
                content + "</v:textbox></v:shape>" + (mode == "linked-end" ?
                "<v:shape style=\"width:10pt;height:10pt\"><v:textbox style=\"mso-next-textbox:#quail-box\"/></v:shape>" : "") + "</w:pict></w:r>";
            var continuation = mode == "linked" ? Drawing("").Replace("<wps:txbx id=\"9\"></wps:txbx>", "<wps:linkedTxbx id=\"9\" seq=\"1\"/>") : "";
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}" xmlns:wp="{WP.wp}" xmlns:wps="{WPS.wps}" xmlns:a="{A.a}" xmlns:v="{VML.vml}">
                  <w:body><w:p>{(kind == "drawing" ? Drawing(content) + continuation : legacy)}</w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:left="720" w:bottom="720" w:right="720"/></w:sectPr>
                  </w:body>
                </w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
        }
        return stream.ToArray();
    }
}
