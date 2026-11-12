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
    [InlineData("contextual", "0", "0")]
    [InlineData("zero-number", "14pt", "14pt")]
    [InlineData("compatibility", "5pt", "10pt")]
    [InlineData("compatibility-off", "14pt", "14pt")]
    [InlineData("revision", "14pt", "14pt")]
    [InlineData("table", "14pt", "14pt")]
    [InlineData("bookmark", "14pt", "14pt")]
    [InlineData("contextual-bookmark", "0", "0")]
    [InlineData("content-control", "14pt", "14pt")]
    [InlineData("section", "14pt", "14pt")]
    // Inside a border box the gap is the upper paragraph's padding, which its shading fills (#1692).
    [InlineData("inherited-border", "padding-bottom: 14pt;", "margin-top: 0;")]
    [InlineData("inherited-frame", "5pt", "5pt")]
    [InlineData("level-border", "0", "0")]
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
                case "content-control":
                    first.ReplaceWith(new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, first))); break;
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

    // Under w:doNotUseHTMLParagraphAutoSpacing Word adds a paragraph's space after to the next
    // one's space before (#1553). Every renderer collapses margins to the larger one, so the
    // second paragraph's margin carries the sum, except below a section break onto a new page.
    [Theory]
    [InlineData("style", "10pt", "15pt")]
    [InlineData("unspaced", "5pt", "5pt")]
    [InlineData("after-lines", "12pt", "17pt")]
    [InlineData("before-lines", "5pt", "11pt")]
    [InlineData("bookmark", "5pt", "10pt")]
    [InlineData("content-control", "5pt", "10pt")]
    [InlineData("continuous-section", "5pt", "10pt")]
    [InlineData("next-page-section", "5pt", "5pt")]
    [InlineData("table", "5pt", "5pt")]
    [InlineData("contextual-first", "0", "5pt")]
    [InlineData("contextual-second", "5pt", "0")]
    public void CompatibilitySpacingAddsTheSpaceAfterAbove(string kind, string after, string before)
    {
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            var first = body.Elements(W.p).ElementAt(1);
            main.DocumentSettingsPart!.GetXDocument().Root!.Add(new XElement(W.compat,
                new XElement(W.doNotUseHTMLParagraphAutoSpacing)));
            switch (kind)
            {
                case "style":
                    first.Element(W.pPr)!.Element(W.spacing)!.Attribute(W.after)!.Remove();
                    main.StyleDefinitionsPart!.GetXDocument().Descendants(W.style).Single().Add(
                        new XElement(W.pPr, new XElement(W.spacing, new XAttribute(W.after, "200"))));
                    main.StyleDefinitionsPart.PutXDocument();
                    break;
                case "unspaced": body.Elements(W.p).ElementAt(2).Element(W.pPr)!.Element(W.spacing)!.Remove(); break;
                // Line counts replace the twips; a line is 12pt without a line grid.
                case "after-lines": first.Element(W.pPr)!.Element(W.spacing)!.SetAttributeValue(W.afterLines, "100"); break;
                case "before-lines":
                    body.Elements(W.p).ElementAt(2).Element(W.pPr)!.Element(W.spacing)!.SetAttributeValue(W.beforeLines, "50"); break;
                case "bookmark":
                    first.AddAfterSelf(new XElement(W.bookmarkStart, new XAttribute(W.id, "1"),
                        new XAttribute(W.name, "AshMark")), new XElement(W.bookmarkEnd, new XAttribute(W.id, "1"))); break;
                case "content-control":
                    first.ReplaceWith(new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, first))); break;
                case "continuous-section":
                case "next-page-section":
                    first.Element(W.pPr)!.Add(new XElement(W.sectPr));
                    if (kind == "continuous-section")
                        body.Element(W.sectPr)!.Add(new XElement(W.type, new XAttribute(W.val, "continuous")));
                    break;
                case "table": first.AddAfterSelf(new XElement(W.tbl, new XElement(W.tblPr),
                    new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "1200"))),
                    new XElement(W.tr, new XElement(W.tc, new XElement(W.p))))); break;
                case "contextual-first": first.Element(W.pPr)!.Add(new XElement(W.contextualSpacing)); break;
                case "contextual-second":
                    body.Elements(W.p).ElementAt(2).Element(W.pPr)!.Add(new XElement(W.contextualSpacing)); break;
            }
            main.PutXDocument();
            main.DocumentSettingsPart.PutXDocument();
        });
        VerifyGaps(bytes, after, before);
    }

    // In a cell the table style's space after outranks the document defaults, so the sum uses the
    // space after the cell paragraph above was rendered with. A cell paragraph rendered on its own
    // has no table style, so only the document render is compared.
    [Theory]
    [InlineData(false, "10pt", "15pt")]
    [InlineData(true, "0", "5pt")]
    public void CompatibilitySpacingInACellAddsTheRenderedSpaceAfter(bool tableStyle, string after, string before)
    {
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            var paragraphs = body.Elements(W.p).Skip(1).Take(2).ToArray();
            paragraphs.Remove();
            paragraphs[0].Element(W.pPr)!.Element(W.spacing)!.Attribute(W.after)!.Remove();
            body.AddFirst(new XElement(W.tbl,
                new XElement(W.tblPr, tableStyle ? new XElement(W.tblStyle, new XAttribute(W.val, "Grid")) : null),
                new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "6000"))),
                new XElement(W.tr, new XElement(W.tc, paragraphs))));
            var styles = main.StyleDefinitionsPart!.GetXDocument().Root!;
            styles.Element(W.docDefaults)!.Add(new XElement(W.pPrDefault, new XElement(W.pPr,
                new XElement(W.spacing, new XAttribute(W.after, "200")))));
            styles.Add(new XElement(W.style, new XAttribute(W.type, "table"), new XAttribute(W.styleId, "Grid"),
                new XElement(W.name, new XAttribute(W.val, "Grid")),
                new XElement(W.pPr, new XElement(W.spacing, new XAttribute(W.after, "0")))));
            main.DocumentSettingsPart!.GetXDocument().Root!.Add(new XElement(W.compat,
                new XElement(W.doNotUseHTMLParagraphAutoSpacing)));
            main.PutXDocument();
            main.StyleDefinitionsPart.PutXDocument();
            main.DocumentSettingsPart.PutXDocument();
        });
        var styles = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions { FabricateCssClasses = false }))
            .Descendants().Where(e => e.Name.LocalName == "p").ToDictionary(e => e.Value, e => (string?)e.Attribute("style"));
        Assert.Contains($"margin-bottom: {after};", styles["AshFirst"]);
        Assert.Contains($"margin-top: {before};", styles["AshSecond"]);
    }

    // Word 2024 prints 14pt between the paragraphs of a cell, header or footer, nothing before a
    // story's first paragraph, and 14pt after the last paragraph of a header or footer, where a
    // cell, like the body, has none (#1553). Under doNotUseHTMLParagraphAutoSpacing a header keeps
    // its explicit values and adds them.
    [Theory]
    [InlineData("cell", "14pt", "14pt", "0", "0")]
    [InlineData("header", "14pt", "14pt", "0", "14pt")]
    [InlineData("footer", "14pt", "14pt", "0", "14pt")]
    [InlineData("header-cell", "14pt", "14pt", "0", "0")]
    [InlineData("compatibility-header", "5pt", "10pt", "5pt", "5pt")]
    public void StoryParagraphsUseTheirStoryNeighborsAndEdges(string story, string after, string before, string first, string last)
    {
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W.body)!;
            var paragraphs = body.Elements(W.p).Skip(1).Take(2).ToArray();
            paragraphs.Remove();
            var table = new XElement(W.tbl, new XElement(W.tblPr), new XElement(W.tblGrid,
                new XElement(W.gridCol, new XAttribute(W._w, "6000"))), new XElement(W.tr, new XElement(W.tc, paragraphs)));
            if (story == "cell") body.AddFirst(table);
            else
            {
                var footer = story == "footer";
                OpenXmlPart part = footer ? main.AddNewPart<FooterPart>() : main.AddNewPart<HeaderPart>();
                part.PutXDocument(new XDocument(new XElement(footer ? W.ftr : W.hdr,
                    story == "header-cell" ? new object[] { table, new XElement(W.p) } : paragraphs)));
                body.Element(W.sectPr)!.AddFirst(new XElement(footer ? W.footerReference : W.headerReference,
                    new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(part))));
            }
            if (story.StartsWith("compatibility"))
                main.DocumentSettingsPart!.PutXDocument(new XDocument(new XElement(W.settings,
                    new XElement(W.compat, new XElement(W.doNotUseHTMLParagraphAutoSpacing)))));
            main.PutXDocument();
        });
        VerifyGaps(bytes, after, before);
        VerifyEdges(bytes, first, last);
    }

    private static void VerifyEdges(byte[] bytes, string first, string last)
    {
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true,
            RenderHeadersAndFooters = true, PaginationMode = (int)PaginationMode.Paginated };
        var paragraphs = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, options)).Descendants()
            .Where(e => e.Name.LocalName == "p" && e.Value is "AshFirst" or "AshSecond").ToArray();
        foreach (var (paragraph, expected) in new[] { (paragraphs[0], $"margin-top: {first};"), (paragraphs[1], $"margin-bottom: {last};") })
        {
            Assert.Contains(expected, (string?)paragraph.Attribute("style"));
            Assert.Contains(expected, HtmlConversionOps.RenderBlockHtml(bytes, (string)paragraph.Attribute("data-anchor")!, options));
        }
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
            var value = i == 0 ? after : before;
            var expected = value.Contains(':') ? value : $"margin-{side}: {value};";
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DocumentAndCellEdgesHaveNoAutomaticSpacing(bool cell)
    {
        // Word prints no automatic space before the first or after the last paragraph of the
        // document or of a table cell (#1553).
        var bytes = Create(0, 0, doc =>
        {
            var main = doc.MainDocumentPart!;
            var paragraphs = main.GetXDocument().Descendants(W.p).ToArray();
            if (cell)
                paragraphs[0].AddAfterSelf(new XElement(W.tbl, new XElement(W.tblPr), new XElement(W.tblGrid,
                    new XElement(W.gridCol, new XAttribute(W._w, "6000"))),
                    new XElement(W.tr, new XElement(W.tc, paragraphs[1..3].Select(p => { p.Remove(); return p; }).ToArray()))));
            else
            {
                paragraphs[0].Remove();
                paragraphs[^1].Remove();
            }
            main.PutXDocument();
        });
        VerifyGaps(bytes, "14pt", "14pt");
        VerifyEdges(bytes, "0", "0");
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
