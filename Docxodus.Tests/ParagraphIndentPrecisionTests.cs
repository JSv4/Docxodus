// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests
{
    /// <summary>
    /// Issue #1128: paragraph indents keep twip precision in the HTML. They used to be written as
    /// inches rounded to 0.01in (0.72pt), so <c>w:left="1417"</c> (70.85pt) became <c>0.98in</c>
    /// (70.56pt).
    /// </summary>
    public class ParagraphIndentPrecisionTests
    {
        private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static XElement ToHtml(string bodyXml, string? numberingXml = null, bool semanticLists = false)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                Write(main, $"<w:document xmlns:w=\"{WNs}\"><w:body>{bodyXml}" +
                    "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>" +
                    "</w:body></w:document>");
                Write(main.AddNewPart<StyleDefinitionsPart>(),
                    $"<w:styles xmlns:w=\"{WNs}\"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"24\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>");
                Write(main.AddNewPart<DocumentSettingsPart>(),
                    $"<w:settings xmlns:w=\"{WNs}\"><w:defaultTabStop w:val=\"720\"/></w:settings>");
                if (numberingXml != null)
                    Write(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering xmlns:w=\"{WNs}\">{numberingXml}</w:numbering>");
            }
            return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("indent.docx", ms.ToArray()),
                new WmlToHtmlConverterSettings { FabricateCssClasses = false, SemanticLists = semanticLists });
        }

        private static void Write(OpenXmlPart part, string xml)
        {
            using var stream = part.GetStream(FileMode.Create);
            var bytes = Encoding.UTF8.GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string Para(string pPr, string text) =>
            $"<w:p><w:pPr>{pPr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

        private static XElement Block(XElement html, string name, string text) =>
            html.Descendants().First(e => e.Name.LocalName == name && e.Value.Contains(text) &&
                !e.Descendants().Any(c => c.Name.LocalName == name && c.Value.Contains(text)));

        /// <summary>The value of <paramref name="property"/> in the element's inline style, or null.</summary>
        private static string? Css(XElement element, string property) =>
            ((string?)element.Attribute("style") ?? "").Split(';')
                .Select(d => d.Split(':', 2))
                .Where(d => d.Length == 2 && d[0].Trim() == property)
                .Select(d => d[1].Trim())
                .FirstOrDefault();

        [Fact]
        public void Leading_indent_keeps_twip_precision()
        {
            var paragraph = Block(ToHtml(Para("<w:ind w:left=\"1417\"/>", "Lead")), "p", "Lead");
            Assert.Equal("70.85pt", Css(paragraph, "margin-left"));
        }

        [Fact]
        public void Trailing_indent_keeps_twip_precision()
        {
            var paragraph = Block(ToHtml(Para("<w:ind w:right=\"1417\"/>", "Trail")), "p", "Trail");
            Assert.Equal("70.85pt", Css(paragraph, "margin-right"));
        }

        [Fact]
        public void First_line_indent_keeps_twip_precision()
        {
            var paragraph = Block(ToHtml(Para("<w:ind w:firstLine=\"283\"/>", "First")), "p", "First");
            Assert.Equal("14.15pt", Css(paragraph, "text-indent"));
        }

        [Fact]
        public void Hanging_indent_keeps_twip_precision()
        {
            var paragraph = Block(ToHtml(Para("<w:ind w:left=\"1417\" w:hanging=\"283\"/>", "Hang")), "p", "Hang");
            Assert.Equal("70.85pt", Css(paragraph, "margin-left"));
            Assert.Equal("-14.15pt", Css(paragraph, "text-indent"));
        }

        [Fact]
        public void Right_to_left_paragraph_maps_precise_indents_to_the_opposite_sides()
        {
            var paragraph = Block(ToHtml(Para("<w:bidi/><w:ind w:left=\"1417\" w:right=\"283\"/>", "Rtl")), "p", "Rtl");
            Assert.Equal("70.85pt", Css(paragraph, "margin-right"));
            Assert.Equal("14.15pt", Css(paragraph, "margin-left"));
        }

        [Fact]
        public void Bordered_group_keeps_twip_precision_on_its_wrapper_and_paragraphs()
        {
            // A top-only border has no left outset, so the wrapper's margin is the indent less the hang.
            var pPr = "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"000000\"/></w:pBdr>" +
                "<w:ind w:left=\"1417\" w:hanging=\"283\"/>";
            var html = ToHtml(Para(pPr, "Boxed"));
            var paragraph = Block(html, "p", "Boxed");
            var wrapper = paragraph.Parent!;
            Assert.Equal("div", wrapper.Name.LocalName);
            Assert.Equal("56.7pt", Css(wrapper, "margin-left"));
            Assert.Equal("14.15pt", Css(paragraph, "margin-left"));
            Assert.Equal("-14.15pt", Css(paragraph, "text-indent"));
        }

        [Theory]
        [InlineData("1417")]
        [InlineData("1")]
        [InlineData("7")]
        [InlineData("2125")]
        public void Bordered_group_leaves_no_residual_indent_on_its_paragraphs(string left)
        {
            // The wrapper takes the whole leading indent; the paragraph keeps none of it.
            var pPr = "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"000000\"/></w:pBdr>" +
                $"<w:ind w:left=\"{left}\"/>";
            var paragraph = Block(ToHtml(Para(pPr, "Boxed")), "p", "Boxed");
            Assert.Equal("0", Css(paragraph, "margin-left"));
        }

        [Fact]
        public void Semantic_list_item_keeps_twip_precision_relative_to_its_parent()
        {
            var numbering = "<w:abstractNum w:abstractNumId=\"0\">" +
                "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/>" +
                "<w:pPr><w:ind w:left=\"1417\" w:hanging=\"283\"/></w:pPr></w:lvl>" +
                "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%2.\"/><w:lvlJc w:val=\"left\"/>" +
                "<w:pPr><w:ind w:left=\"2125\" w:hanging=\"283\"/></w:pPr></w:lvl></w:abstractNum>" +
                "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
            string Item(string text, int level) =>
                Para($"<w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"1\"/></w:numPr>", text);
            var html = ToHtml(Item("Outer", 0) + Item("Inner", 1), numbering, semanticLists: true);

            Assert.Equal("70.85pt", Css(Block(html, "li", "Outer"), "margin-left"));
            // 2125 - 1417 twips = 35.4pt inside the parent item.
            Assert.Equal("35.4pt", Css(Block(html, "li", "Inner"), "margin-left"));
        }
    }
}
