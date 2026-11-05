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
    public void Supported_placement_requires_effectively_fixed_sizing(string mode, bool eligible)
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
    [InlineData("implicit", "spAutoFit", "noAutofit", "noAutofit", false, true)]
    [InlineData("implicit", "", "spAutoFit", "noAutofit", false, true)]
    [InlineData("implicit", "", "", "spAutoFit", false, true)]
    [InlineData("implicit", "noAutofit", "spAutoFit", "normAutofit", true, false)]
    [InlineData("implicit", "normAutofit", "noAutofit", "noAutofit", false, false)]
    [InlineData("fixed", "spAutoFit", "spAutoFit", "spAutoFit", true, false)]
    [InlineData("auto", "noAutofit", "noAutofit", "noAutofit", false, true)]
    [InlineData("shrink", "noAutofit", "noAutofit", "noAutofit", false, false)]
    public void Drawing_fit_uses_shape_then_theme_text_line_and_shape_defaults(
        string local, string text, string line, string shape, bool fixedFit, bool autoFit)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("drawing", local));
        using var document = WordprocessingDocument.Open(stream, true);
        var defaults = new XElement(A.objectDefaults);
        foreach (var (name, fit) in new[] { (A.spDef, shape), (A.lnDef, line), (A.txDef, text) })
            defaults.Add(new XElement(name, new XElement(A.spPr),
                new XElement(A.bodyPr, fit == "" ? null : new XElement(A.a + fit)), new XElement(A.lstStyle)));
        document.MainDocumentPart!.AddNewPart<ThemePart>().PutXDocument(
            new XDocument(new XElement(A.theme, defaults)));
        var inline = document.MainDocumentPart.GetXDocument().Descendants(WP.inline).Single();
        inline.Name = WP.anchor;
        inline.AddFirst(new XElement(WP.positionH, new XAttribute("relativeFrom", "page"), new XElement(WP.posOffset, 762000)),
            new XElement(WP.positionV, new XAttribute("relativeFrom", "page"), new XElement(WP.posOffset, 762000)));
        inline.Add(new XElement(WP.wrapNone));
        document.MainDocumentPart.PutXDocument();
        var html = WmlToHtmlConverter.ConvertToHtml(document,
            new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated, FabricateCssClasses = false });
        var box = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-textbox") == "true");
        Assert.Equal(fixedFit ? "true" : null, (string?)box.Attribute("data-docx-textbox-source-position"));
        Assert.Equal(!autoFit, ((string?)box.Attribute("style") ?? "").Split(';')
            .Any(declaration => declaration.TrimStart().StartsWith("height:", System.StringComparison.Ordinal)));
    }

    [Fact]
    public void Theme_fit_is_rechecked_after_source_edits()
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("drawing", "implicit"));
        using var document = WordprocessingDocument.Open(stream, true);
        var theme = document.MainDocumentPart!.AddNewPart<ThemePart>();
        foreach (var mode in new[] { A.noAutofit, A.spAutoFit, A.normAutofit, A.noAutofit })
        {
            theme.PutXDocument(new XDocument(new XElement(A.theme, new XElement(A.objectDefaults,
                new XElement(A.txDef, new XElement(A.bodyPr, new XElement(mode)))))));
            var html = WmlToHtmlConverter.ConvertToHtml(document,
                new WmlToHtmlConverterSettings { FabricateCssClasses = false });
            var box = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-textbox") == "true");
            Assert.Equal(mode != A.spAutoFit, ((string?)box.Attribute("style") ?? "").Split(';')
                .Any(declaration => declaration.TrimStart().StartsWith("height:", System.StringComparison.Ordinal)));
        }
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("choice")]
    [InlineData("body")]
    [InlineData("theme-body")]
    [InlineData("theme-defaults")]
    public void Missing_or_ambiguous_fit_context_cannot_claim_source_provenance(string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("drawing", "fixed"));
        using var document = WordprocessingDocument.Open(stream, true);
        var main = document.MainDocumentPart!;
        var bodyPr = main.GetXDocument().Descendants(WPS.bodyPr).Single();
        if (mode == "missing") bodyPr.Remove();
        else if (mode == "choice") bodyPr.Add(new XElement(A.normAutofit));
        else if (mode == "body") bodyPr.AddAfterSelf(new XElement(WPS.bodyPr, new XElement(A.spAutoFit)));
        else
        {
            var properties = new XElement(A.bodyPr, new XElement(A.noAutofit));
            var definition = new XElement(A.txDef, properties);
            var defaults = new XElement(A.objectDefaults, definition);
            if (mode == "theme-body") properties.AddAfterSelf(new XElement(A.bodyPr, new XElement(A.spAutoFit)));
            else definition.AddAfterSelf(new XElement(A.txDef, new XElement(A.bodyPr, new XElement(A.spAutoFit))));
            main.AddNewPart<ThemePart>().PutXDocument(new XDocument(new XElement(A.theme, defaults)));
        }
        main.PutXDocument();
        var html = WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings());
        var box = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-textbox") == "true");
        Assert.Null(box.Attribute("data-docx-textbox-source-clip"));
        Assert.Null(box.Attribute("data-docx-textbox-source-position"));
    }

    [Fact]
    public void Nested_autofit_box_does_not_change_its_fixed_outer_story()
    {
        using var innerStream = new MemoryStream(Fixture("drawing", "auto"));
        using var innerDocument = WordprocessingDocument.Open(innerStream, false);
        var inner = new XElement(innerDocument.MainDocumentPart!.GetXDocument().Descendants(W.drawing).Single());
        inner.Descendants(W.t).Single().Value = "Nested marker";
        inner.Descendants(WP.docPr).Single().SetAttributeValue("id", 2);
        using var stream = new MemoryStream();
        stream.Write(Fixture("drawing", "fixed"));
        using var document = WordprocessingDocument.Open(stream, true);
        var main = document.MainDocumentPart!;
        main.GetXDocument().Descendants(W.txbxContent).Single().Element(W.p)!.Add(new XElement(W.r, inner));
        main.PutXDocument();
        var html = WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var boxes = html.Descendants().Where(e => (string?)e.Attribute("data-docx-textbox") == "true").ToList();
        Assert.Equal(2, boxes.Count);
        Assert.Equal("true", (string?)boxes[0].Attribute("data-docx-textbox-source-clip"));
        Assert.Contains("height:", (string?)boxes[0].Attribute("style"));
        Assert.DoesNotContain(((string?)boxes[1].Attribute("style") ?? "").Split(';'),
            declaration => declaration.TrimStart().StartsWith("height:", System.StringComparison.Ordinal));
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
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        }
        return stream.ToArray();
    }
}
