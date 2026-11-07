using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlAutomaticSpacingTests
{
    [Theory]
    [InlineData(0, 0, "14pt")]
    [InlineData(0, 1, "14pt")]
    [InlineData(1, 0, "14pt")]
    [InlineData(1, 1, "0")]
    [InlineData(1, 2, "14pt")]
    public void AutomaticGapsUseTheOriginalAdjacentListInEveryRenderPath(int firstList, int secondList, string gap)
        => VerifyGaps(Create(firstList, secondList), gap, gap);

    [Theory]
    [InlineData("inherited", "14pt", "14pt")]
    [InlineData("direct-off", "5pt", "14pt")]
    [InlineData("explicit-neighbor", "14pt", "30pt")]
    [InlineData("contextual", "0", "14pt")]
    [InlineData("zero-number", "14pt", "14pt")]
    [InlineData("compatibility", "5pt", "5pt")]
    [InlineData("compatibility-off", "14pt", "14pt")]
    [InlineData("revision", "5pt", "5pt")]
    [InlineData("table", "5pt", "5pt")]
    [InlineData("bookmark", "5pt", "5pt")]
    [InlineData("contextual-bookmark", "5pt", "5pt")]
    [InlineData("section", "5pt", "5pt")]
    [InlineData("inherited-border", "5pt", "5pt")]
    [InlineData("inherited-frame", "5pt", "5pt")]
    [InlineData("level-border", "5pt", "5pt")]
    public void FlagsAndOriginalBoundariesStayConsistent(string kind, string after, string before)
    {
        var bytes = Create(kind == "level-border" ? 1 : 0, kind == "level-border" ? 1 : 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            var first = body.Elements(W.p).ElementAt(1);
            var second = body.Elements(W.p).ElementAt(2);
            var firstProps = first.Element(W.pPr)!;
            var secondProps = second.Element(W.pPr)!;
            var styleProps = new XElement(W.pPr);
            main.StyleDefinitionsPart!.GetXDocument().Descendants(W.style).Single().Add(styleProps);
            var border = new XElement(W.pBdr, new XElement(W.bottom,
                new XAttribute(W.val, "single"), new XAttribute(W.sz, "8")));
            switch (kind)
            {
                case "inherited":
                    styleProps.Add(new XElement(W.spacing, new XAttribute(W.beforeAutospacing, "1"),
                        new XAttribute(W.afterAutospacing, "1")));
                    foreach (var p in new[] { firstProps, secondProps })
                        p.Element(W.spacing)!.Attributes().Where(a => a.Name == W.beforeAutospacing || a.Name == W.afterAutospacing).Remove();
                    break;
                case "direct-off": firstProps.Element(W.spacing)!.SetAttributeValue(W.afterAutospacing, "0"); break;
                case "explicit-neighbor":
                    secondProps.Element(W.spacing)!.SetAttributeValue(W.beforeAutospacing, "0");
                    secondProps.Element(W.spacing)!.SetAttributeValue(W.before, "600");
                    break;
                case "contextual": firstProps.Add(new XElement(W.contextualSpacing)); break;
                case "zero-number":
                    foreach (var p in new[] { firstProps, secondProps })
                        p.AddFirst(new XElement(W.numPr, new XElement(W.numId, new XAttribute(W.val, "0"))));
                    break;
                case "compatibility":
                case "compatibility-off":
                    main.DocumentSettingsPart!.GetXDocument().Root!.Add(new XElement(W.compat,
                        new XElement(W.doNotUseHTMLParagraphAutoSpacing, new XAttribute(W.val, kind == "compatibility" ? "1" : "0"))));
                    break;
                case "revision":
                    body.Elements(W.p).Last().Add(new XElement(W.ins, new XAttribute(W.id, "1"),
                        new XAttribute(W.author, "Editor"), new XElement(W.r, new XElement(W.t, " revised"))));
                    break;
                case "table": first.AddAfterSelf(new XElement(W.tbl, new XElement(W.tblPr),
                    new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "1200"))),
                    new XElement(W.tr, new XElement(W.tc, new XElement(W.p))))); break;
                case "bookmark":
                case "contextual-bookmark":
                    if (kind == "contextual-bookmark") firstProps.Add(new XElement(W.contextualSpacing));
                    first.AddAfterSelf(new XElement(W.bookmarkStart, new XAttribute(W.id, "1"),
                        new XAttribute(W.name, "AshMark")), new XElement(W.bookmarkEnd, new XAttribute(W.id, "1"))); break;
                case "section": firstProps.Add(new XElement(W.sectPr)); break;
                case "inherited-border": styleProps.Add(border); break;
                case "inherited-frame": styleProps.Add(new XElement(W.framePr,
                    new XAttribute(W._w, "3600"), new XAttribute(W.hAnchor, "text"), new XAttribute(W.vAnchor, "text"))); break;
                case "level-border": main.NumberingDefinitionsPart!.GetXDocument().Descendants(W.lvl).Single()
                    .Element(W.pPr)!.AddFirst(border); break;
            }
            main.PutXDocument();
            main.StyleDefinitionsPart.PutXDocument();
            main.DocumentSettingsPart!.PutXDocument();
            main.NumberingDefinitionsPart!.PutXDocument();
        });
        VerifyGaps(bytes, after, before, windowed: !kind.EndsWith("border") && kind != "inherited-frame");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CellAndHeaderClonesDoNotBecomeOrdinaryBodyParagraphs(bool header)
    {
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            var paragraphs = body.Elements(W.p).Skip(1).Take(2).ToArray();
            paragraphs.Remove();
            if (header)
            {
                var part = main.AddNewPart<HeaderPart>();
                part.PutXDocument(new XDocument(new XElement(W.hdr, paragraphs)));
                body.Element(W.sectPr)!.AddFirst(new XElement(W.headerReference,
                    new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(part))));
            }
            else body.AddFirst(new XElement(W.tbl, new XElement(W.tblPr), new XElement(W.tblGrid,
                new XElement(W.gridCol, new XAttribute(W._w, "6000"))), new XElement(W.tr, new XElement(W.tc, paragraphs))));
            main.PutXDocument();
        });
        VerifyGaps(bytes, "5pt", "5pt");
    }

    private static void VerifyGaps(byte[] bytes, string after, string before, bool windowed = true)
    {
        using (var validation = WordprocessingDocument.Open(new MemoryStream(bytes), false))
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(validation));
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true,
            RenderHeadersAndFooters = true, PaginationMode = (int)PaginationMode.Paginated };
        var full = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, options));
        var paragraphs = full.Descendants().Where(e => e.Name.LocalName == "p" &&
            (e.Value.Contains("AshFirst") || e.Value.Contains("AshSecond"))).ToArray();
        Assert.Equal(2, paragraphs.Length);
        using var session = new DocxSession(bytes);
        var anchors = paragraphs.Select(p => (string)p.Attribute("data-anchor")!).ToArray();
        using var batch = JsonDocument.Parse(HtmlConversionOps.RenderBlocksHtml(session, anchors, options));
        using var window = JsonDocument.Parse(HtmlConversionOps.RenderBlocksRangeHtml(session, anchors, options));
        for (int i = 0; i < paragraphs.Length; i++)
        {
            var side = i == 0 ? "bottom" : "top";
            var expected = $"margin-{side}: {(i == 0 ? after : before)};";
            Assert.Contains(expected, (string?)paragraphs[i].Attribute("style"));
            foreach (var html in new[] {
                HtmlConversionOps.RenderBlockHtml(bytes, anchors[i], options),
                HtmlConversionOps.RenderBlockHtml(session, anchors[i], options),
                batch.RootElement.GetProperty(anchors[i]).GetString()! })
                Assert.Contains(expected, (string?)XElement.Parse(html).DescendantsAndSelf()
                    .Single(e => e.Name.LocalName == "p" && e.Value.Contains(i == 0 ? "AshFirst" : "AshSecond")).Attribute("style"));
            if (windowed)
                Assert.Contains(expected, (string?)window.RootElement.EnumerateArray().Select(e => XElement.Parse(e.GetString()!))
                    .SelectMany(e => e.DescendantsAndSelf()).Single(e => e.Name.LocalName == "p"
                        && e.Value.Contains(i == 0 ? "AshFirst" : "AshSecond")).Attribute("style"));
        }
    }

    [Fact]
    public void UnmeasuredDocumentEdgesKeepTheirExplicitSpacing()
    {
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var paragraphs = main.GetXDocument().Descendants(W.p).ToArray();
            paragraphs[0].Remove();
            paragraphs[^1].Remove();
            main.PutXDocument();
        });
        VerifyGaps(bytes, "14pt", "14pt");
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true };
        var paragraphs = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, options))
            .Descendants().Where(e => e.Name.LocalName == "p").ToArray();
        foreach (var (paragraph, side) in new[] { (paragraphs[0], "top"), (paragraphs[1], "bottom") })
        {
            var expected = $"margin-{side}: 5pt;";
            Assert.Contains(expected, (string?)paragraph.Attribute("style"));
            Assert.Contains(expected, HtmlConversionOps.RenderBlockHtml(bytes,
                (string)paragraph.Attribute("data-anchor")!, options));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DirectListIdentityDoesNotRequireAStylesPart(bool emptyStyles)
    {
        var bytes = Create(1, 1, doc =>
        {
            var main = doc.MainDocumentPart!;
            if (emptyStyles) main.StyleDefinitionsPart!.PutXDocument(new XDocument(new XElement(W.styles)));
            else main.DeletePart(main.StyleDefinitionsPart!);
        });
        VerifyGaps(bytes, "0", "0");
    }

    [Fact]
    public void RangeContextClonesAContentControlsRequestedChildrenOnlyOnce()
    {
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            var children = body.Elements(W.p).Skip(1).Take(2).ToArray();
            children.Remove();
            body.Elements(W.p).First().AddAfterSelf(new XElement(W.sdt, new XElement(W.sdtPr),
                new XElement(W.sdtContent, children)));
            main.PutXDocument();
        });
        var options = new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false };
        var anchors = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, options)).Descendants()
            .Where(e => e.Name.LocalName == "p" && !e.Value.Contains("AshClosing"))
            .Select(e => (string)e.Attribute("data-anchor")!).ToArray();
        using var session = new DocxSession(bytes);
        using var range = JsonDocument.Parse(HtmlConversionOps.RenderBlocksRangeHtml(session, anchors, options));
        var observed = range.RootElement.EnumerateArray().SelectMany(e => XElement.Parse(e.GetString()!).DescendantsAndSelf())
            .Select(e => (string?)e.Attribute("data-anchor")).Where(a => a != null).ToArray();
        Assert.Equal(anchors, observed);
    }

    private static byte[] Create(int firstList, int secondList, Action<WordprocessingDocument>? configure = null)
    {
        static string Paragraph(string text, int list, bool automatic = false) =>
            "<w:p><w:pPr>" + (list == 0 ? "" : $"<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"{list}\"/></w:numPr>") +
            "<w:spacing w:before=\"100\" w:after=\"100\" w:line=\"320\" w:lineRule=\"exact\"" +
            (automatic ? " w:beforeAutospacing=\"1\" w:afterAutospacing=\"1\"" : "") +
            $"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            void Write(OpenXmlPart part, string root, string content)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write($"<w:{root} xmlns:w=\"{W.w}\">{content}</w:{root}>");
            }
            Write(main, "document", "<w:body>" + Paragraph("AshOpening", 0) +
                Paragraph("AshFirst", firstList, true) + Paragraph("AshSecond", secondList, true) +
                Paragraph("AshClosing", 0) + "<w:sectPr/></w:body>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), "styles", "<w:docDefaults><w:rPrDefault><w:rPr>" +
                "<w:rFonts w:ascii=\"Liberation Serif\" w:hAnsi=\"Liberation Serif\"/><w:sz w:val=\"22\"/>" +
                "</w:rPr></w:rPrDefault></w:docDefaults>" +
                "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>");
            Write(main.AddNewPart<NumberingDefinitionsPart>(), "numbering", "<w:abstractNum w:abstractNumId=\"0\">" +
                "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>" +
                "<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/>" +
                "<w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum>" +
                "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>" +
                "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"0\"/></w:num>");
            Write(main.AddNewPart<DocumentSettingsPart>(), "settings", "");
            configure?.Invoke(doc);
        }
        return stream.ToArray();
    }
}
