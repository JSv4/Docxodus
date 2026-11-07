using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlLineGridTests
{
    [Theory]
    [InlineData("single", true)]
    [InlineData("direct-on", true)]
    [InlineData("inherited-off", false)]
    [InlineData("direct-off", false)]
    [InlineData("exact", false)]
    [InlineData("multiple", false)]
    [InlineData("no-type", false)]
    [InlineData("none", false)]
    [InlineData("no-grid", false)]
    [InlineData("cell", false)]
    [InlineData("header", false)]
    [InlineData("content-control", false)]
    [InlineData("revision", false)]
    public void OriginalSectionGridAndResolvedFlagsAgreeAcrossRenderPaths(string kind, bool enabled)
    {
        var bytes = Create(kind);
        using (var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false))
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(doc));
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true,
            PaginationMode = (int)PaginationMode.Paginated, RenderHeadersAndFooters = true };
        var full = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, options));
        var paragraphs = full.Descendants().Where(e => e.Name.LocalName == "p" && e.Value.StartsWith("Cedar", StringComparison.Ordinal)).ToArray();
        Assert.Equal(2, paragraphs.Length);
        var anchors = paragraphs.Select(p => (string)p.Attribute("data-anchor")!).ToArray();
        using var session = new DocxSession(bytes);
        using var batch = JsonDocument.Parse(HtmlConversionOps.RenderBlocksHtml(session, anchors, options));
        using var window = JsonDocument.Parse(HtmlConversionOps.RenderBlocksRangeHtml(session, anchors, options));
        for (int i = 0; i < 2; i++)
        {
            var expected = $"--docx-line-grid: {(i == 0 ? 18 : 12)}pt;";
            foreach (var html in new[] { paragraphs[i].ToString(),
                HtmlConversionOps.RenderBlockHtml(bytes, anchors[i], options),
                HtmlConversionOps.RenderBlockHtml(session, anchors[i], options),
                batch.RootElement.GetProperty(anchors[i]).GetString()!,
                window.RootElement.EnumerateArray().Select(e => e.GetString()!)
                    .Single(s => s.Contains(i == 0 ? "CedarFirst" : "CedarSecond", StringComparison.Ordinal)) })
            {
                var paragraph = XElement.Parse(html).DescendantsAndSelf().Single(e => e.Name.LocalName == "p"
                    && e.Value == (i == 0 ? "CedarFirst" : "CedarSecond"));
                var style = (string?)paragraph.Attribute("style") ?? "";
                if (enabled) Assert.Contains(expected, style);
                else Assert.DoesNotContain("--docx-line-grid", style);
            }
        }
    }

    private static byte[] Create(string kind)
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSession.CreateBlankDocxBytes());
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            XElement Section(int pitch) => new(W.sectPr, new XElement(W.pgSz, new XAttribute(W._w, 12240), new XAttribute(W.h, 15840)),
                kind == "no-grid" ? null : new XElement(W.docGrid,
                    kind == "no-type" ? null : new XAttribute(W.type, kind == "none" ? "default" : "lines"), new XAttribute(W.linePitch, pitch)));
            XElement Paragraph(string text, XElement? section) => new(W.p, new XElement(W.pPr,
                kind is "direct-on" or "direct-off" ? new XElement(W.snapToGrid, new XAttribute(W.val, kind == "direct-on" ? "1" : "0")) : null,
                new XElement(W.spacing, new XAttribute(W.before, 0), new XAttribute(W.after, 0),
                    new XAttribute(W.line, kind == "multiple" ? 480 : 240), new XAttribute(W.lineRule, kind == "exact" ? "exact" : "auto")), section),
                new XElement(W.r, new XElement(W.t, text)));
            main.GetXDocument().Root!.Element(W.body)!.ReplaceNodes(Paragraph("CedarFirst", Section(360)),
                Paragraph("CedarSecond", null), Section(240));
            var body = main.GetXDocument().Root!.Element(W.body)!;
            if (kind is "cell" or "header" or "content-control")
            {
                var paragraphs = body.Elements(W.p).ToArray();
                paragraphs.Elements(W.pPr).Elements(W.sectPr).Remove();
                paragraphs.Remove();
                if (kind == "header")
                {
                    var header = main.AddNewPart<HeaderPart>();
                    header.PutXDocument(new XDocument(new XElement(W.hdr, paragraphs)));
                    body.AddFirst(Paragraph("Body control", null));
                    body.Element(W.sectPr)!.AddFirst(new XElement(W.headerReference,
                        new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(header))));
                }
                else body.AddFirst(kind == "cell"
                    ? new XElement(W.tbl, new XElement(W.tblPr), new XElement(W.tblGrid,
                        new XElement(W.gridCol, new XAttribute(W._w, 9360))), new XElement(W.tr, new XElement(W.tc, paragraphs)))
                    : new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, paragraphs)));
            }
            if (kind == "revision") body.AddFirst(new XElement(W.p, new XElement(W.ins,
                new XAttribute(W.id, 1), new XAttribute(W.author, "Editor"), new XElement(W.r, new XElement(W.t, "Revised control")))));
            var styles = main.StyleDefinitionsPart!;
            var runDefaults = new XElement(W.rPrDefault, new XElement(W.rPr, new XElement(W.rFonts,
                new XAttribute(W.ascii, "Liberation Serif"), new XAttribute(W.hAnsi, "Liberation Serif")),
                new XElement(W.sz, new XAttribute(W.val, 24))));
            var paragraphDefaults = new XElement(W.pPrDefault, new XElement(W.pPr,
                kind is "inherited-off" or "direct-on" ? new XElement(W.snapToGrid, new XAttribute(W.val, "0")) : null));
            styles.PutXDocument(new XDocument(new XElement(W.styles,
                new XElement(W.docDefaults, runDefaults, paragraphDefaults))));
            main.PutXDocument();
        }
        return stream.ToArray();
    }
}
