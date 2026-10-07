// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace OxPt
{
    /// <summary>
    /// Issue #895: with <see cref="WmlToHtmlConverterSettings.SemanticLists"/> on, Word list
    /// paragraphs convert to <c>ol</c>/<c>ul</c>/<c>li</c> instead of paragraphs with a generated
    /// marker.
    /// </summary>
    public class WmlSemanticListsTests
    {
        private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static byte[] BuildDocx(string bodyXml, string numberingXml, string stylesXml = "")
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                Write(main, $"<w:document xmlns:w=\"{WNs}\"><w:body>{bodyXml}" +
                    "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>" +
                    "</w:body></w:document>");
                Write(main.AddNewPart<StyleDefinitionsPart>(),
                    $"<w:styles xmlns:w=\"{WNs}\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/>" +
                    "<w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
                    $"<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>{stylesXml}</w:styles>");
                Write(main.AddNewPart<DocumentSettingsPart>(),
                    $"<w:settings xmlns:w=\"{WNs}\"><w:defaultTabStop w:val=\"720\"/></w:settings>");
                Write(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering xmlns:w=\"{WNs}\">{numberingXml}</w:numbering>");
            }
            return ms.ToArray();
        }

        private static void Write(OpenXmlPart part, string xml)
        {
            using var stream = part.GetStream(FileMode.Create);
            var bytes = Encoding.UTF8.GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>An abstract list whose levels are (numFmt, lvlText) pairs, each indented a further half inch.</summary>
        private static string AbstractNum(int id, params (string Format, string Text)[] levels) =>
            $"<w:abstractNum w:abstractNumId=\"{id}\">" +
            string.Concat(levels.Select((l, i) =>
                $"<w:lvl w:ilvl=\"{i}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{l.Format}\"/><w:lvlText w:val=\"{l.Text}\"/>" +
                $"<w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{720 * (i + 1)}\" w:hanging=\"360\"/></w:pPr></w:lvl>")) +
            "</w:abstractNum>";

        private static string Num(int numId, int abstractNumId, string overrides = "") =>
            $"<w:num w:numId=\"{numId}\"><w:abstractNumId w:val=\"{abstractNumId}\"/>{overrides}</w:num>";

        private static string Item(string text, int numId = 1, int level = 0, string extraPPr = "") =>
            $"<w:p><w:pPr>{extraPPr}<w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{numId}\"/></w:numPr></w:pPr>" +
            $"<w:r><w:t>{text}</w:t></w:r></w:p>";

        private static string Plain(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";

        private static readonly string Decimal = AbstractNum(0, ("decimal", "%1."), ("lowerLetter", "%2."));

        private static XElement Convert(byte[] docx, bool semanticLists = true, PaginationMode pagination = PaginationMode.None) =>
            WmlToHtmlConverter.ConvertToHtml(new WmlDocument("t.docx", docx), new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                SemanticLists = semanticLists,
                RenderPagination = pagination,
            });

        private static XElement Body(XElement html) => html.Descendants().First(e => e.Name.LocalName == "body");

        private static string Style(XElement e) => (string?)e.Attribute("style") ?? "";

        private static bool HasMarkerSpan(XElement li) =>
            li.Elements().Any(e => (string?)e.Attribute("data-list-marker") == "true");

        /// <summary>The item's own text, without any nested list.</summary>
        private static string OwnText(XElement li) =>
            string.Concat(li.Nodes().Where(n => n is not XElement e || e.Name.LocalName is not ("ol" or "ul"))
                .Select(n => n is XElement e ? e.Value : (n as XText)?.Value)).Trim();

        [Fact]
        public void Adjacent_numbered_items_become_one_ordered_list_drawn_by_css()
        {
            var body = Body(Convert(BuildDocx(Item("One") + Item("Two") + Item("Three"), Decimal + Num(1, 0))));

            var list = Assert.Single(body.Descendants().Where(e => e.Name.LocalName == "ol"));
            Assert.Contains("list-style-type: decimal", Style(list));
            Assert.Equal(new[] { "One", "Two", "Three" }, list.Elements().Select(OwnText));
            Assert.All(list.Elements(), li => Assert.Equal("li", li.Name.LocalName));
            Assert.All(list.Elements(), li => Assert.False(HasMarkerSpan(li)));
            Assert.Null(list.Attribute("start"));
            Assert.DoesNotContain(body.Descendants(), e => e.Name.LocalName == "p");
        }

        [Fact]
        public void Off_by_default_list_items_stay_paragraphs()
        {
            var body = Body(Convert(BuildDocx(Item("One") + Item("Two"), Decimal + Num(1, 0)), semanticLists: false));

            Assert.DoesNotContain(body.Descendants(), e => e.Name.LocalName is "ol" or "ul" or "li");
            Assert.All(body.Descendants().Where(e => e.Name.LocalName == "p"), p => Assert.True(HasMarkerSpan(p)));
        }

        [Fact]
        public void Bullets_become_an_unordered_list()
        {
            var bullets = AbstractNum(0, ("bullet", "•"));
            var body = Body(Convert(BuildDocx(Item("Apples") + Item("Pears"), bullets + Num(1, 0))));

            var list = Assert.Single(body.Descendants().Where(e => e.Name.LocalName == "ul"));
            Assert.Contains("list-style-type: disc", Style(list));
            Assert.Equal(2, list.Elements().Count());
        }

        [Fact]
        public void A_deeper_level_nests_inside_the_preceding_item_with_its_indent_measured_from_it()
        {
            var body = Body(Convert(BuildDocx(
                Item("One") + Item("One a", level: 1) + Item("One b", level: 1) + Item("Two"), Decimal + Num(1, 0))));

            var outer = body.Descendants().First(e => e.Name.LocalName == "ol");
            var items = outer.Elements().ToList();
            Assert.Equal(new[] { "One", "Two" }, items.Select(OwnText));
            var inner = items[0].Elements().Single(e => e.Name.LocalName == "ol");
            Assert.Contains("list-style-type: lower-alpha", Style(inner));
            Assert.Equal(new[] { "One a", "One b" }, inner.Elements().Select(OwnText));

            // Level 0 sits at 0.5 in, level 1 at 1.0 in: the nested item is 0.5 in inside its parent.
            Assert.Contains("margin-left: 0.50in", Style(items[0]));
            Assert.All(inner.Elements(), li => Assert.Contains("margin-left: 0.50in", Style(li)));
            Assert.All(inner.Elements(), li => Assert.Contains("text-indent: 0", Style(li)));
        }

        [Fact]
        public void A_marker_css_cannot_draw_keeps_its_span_and_no_css_marker()
        {
            var parenthesized = AbstractNum(0, ("lowerLetter", "(%1)"));
            var body = Body(Convert(BuildDocx(Item("First") + Item("Second"), parenthesized + Num(1, 0))));

            var list = Assert.Single(body.Descendants().Where(e => e.Name.LocalName == "ol"));
            Assert.Contains("list-style-type: none", Style(list));
            Assert.All(list.Elements(), li => Assert.True(HasMarkerSpan(li)));
            Assert.All(list.Elements(), li => Assert.Contains("display: block", Style(li)));
            Assert.All(list.Elements(), li => Assert.Contains("text-indent: -0.25in", Style(li)));
        }

        [Fact]
        public void A_list_resumed_after_an_interruption_starts_at_its_number()
        {
            var body = Body(Convert(BuildDocx(
                Item("One") + Item("Two") + Plain("Between") + Item("Three"), Decimal + Num(1, 0))));

            var lists = body.Descendants().Where(e => e.Name.LocalName == "ol").ToList();
            Assert.Equal(2, lists.Count);
            Assert.Null(lists[0].Attribute("start"));
            Assert.Equal("3", (string?)lists[1].Attribute("start"));
        }

        [Fact]
        public void A_start_override_becomes_start()
        {
            var overridden = Num(1, 0, "<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride>");
            var body = Body(Convert(BuildDocx(Item("Five") + Item("Six"), Decimal + overridden)));

            var list = Assert.Single(body.Descendants().Where(e => e.Name.LocalName == "ol"));
            Assert.Equal("5", (string?)list.Attribute("start"));
            Assert.All(list.Elements(), li => Assert.Null(li.Attribute("value")));
        }

        [Fact]
        public void Two_adjacent_lists_stay_separate()
        {
            var body = Body(Convert(BuildDocx(
                Item("A1") + Item("A2") + Item("B1", numId: 2) + Item("B2", numId: 2),
                Decimal + Num(1, 0) + Num(2, 0, "<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride>"))));

            var lists = body.Descendants().Where(e => e.Name.LocalName == "ol").ToList();
            Assert.Equal(2, lists.Count);
            Assert.Equal(new[] { "B1", "B2" }, lists[1].Elements().Select(OwnText));
        }

        [Fact]
        public void A_numbered_heading_stays_a_heading()
        {
            var styles = "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/>" +
                "<w:pPr><w:outlineLvl w:val=\"0\"/></w:pPr></w:style>";
            var body = Body(Convert(BuildDocx(
                Item("Chapter", extraPPr: "<w:pStyle w:val=\"Heading1\"/>") + Item("Point"), Decimal + Num(1, 0), styles)));

            Assert.Contains(body.Descendants(), e => e.Name.LocalName == "h1" && e.Value.Contains("Chapter"));
            Assert.Equal("Point", OwnText(body.Descendants().Single(e => e.Name.LocalName == "li")));
        }

        [Fact]
        public void A_list_in_a_table_cell_stays_in_the_cell()
        {
            var cell = "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
                $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>{Item("In cell")}{Item("Also in cell")}</w:tc></w:tr></w:tbl>";
            var body = Body(Convert(BuildDocx(Item("Before") + cell, Decimal + Num(1, 0))));

            var td = body.Descendants().Single(e => e.Name.LocalName == "td");
            var inCell = Assert.Single(td.Elements().Where(e => e.Name.LocalName == "ol"));
            Assert.Equal(2, inCell.Elements().Count());
            Assert.Equal("Before", OwnText(body.Descendants().First(e => e.Name.LocalName == "ol").Elements().Single()));
        }

        [Fact]
        public void An_items_space_after_moves_above_its_nested_list()
        {
            // Within one numbered group the converter already drops space after (Word's numbered
            // paragraph grouping); an item followed by a deeper level of ANOTHER list keeps it.
            var spaced = "<w:spacing w:after=\"240\"/>";
            var body = Body(Convert(BuildDocx(
                Item("One", extraPPr: spaced) + Item("Nested", numId: 2, level: 1),
                Decimal + AbstractNum(1, ("decimal", "%1."), ("decimal", "%2.")) + Num(1, 0) + Num(2, 1))));

            var first = body.Descendants().First(e => e.Name.LocalName == "li");
            var nested = first.Elements().Single(e => e.Name.LocalName == "ol");
            Assert.DoesNotContain("margin-bottom", Style(first), System.StringComparison.Ordinal);
            Assert.Contains("margin-top: 12pt", Style(nested), System.StringComparison.Ordinal);
        }

        [Fact]
        public void A_marker_in_another_font_keeps_its_span()
        {
            // The marker run's own size makes its line box the tallest on the first line; a ::marker,
            // which takes the item's font, would shorten the line.
            var bigMarker = "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>" +
                "<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr>" +
                "<w:rPr><w:sz w:val=\"40\"/></w:rPr></w:lvl></w:abstractNum>";
            var body = Body(Convert(BuildDocx(Item("One") + Item("Two"), bigMarker + Num(1, 0))));

            var list = Assert.Single(body.Descendants().Where(e => e.Name.LocalName == "ol"));
            Assert.Contains("list-style-type: none", Style(list));
            Assert.All(list.Elements(), li => Assert.True(HasMarkerSpan(li)));
        }

        [Fact]
        public void A_right_to_left_item_measures_its_indent_on_the_right()
        {
            var body = Body(Convert(BuildDocx(
                Item("One", extraPPr: "<w:bidi/>") + Item("One a", level: 1, extraPPr: "<w:bidi/>"), Decimal + Num(1, 0))));

            var nested = body.Descendants().Where(e => e.Name.LocalName == "li").Last();
            Assert.Contains("margin-right: 0.50in", Style(nested));
        }

        [Fact]
        public void Paginated_output_ignores_the_setting()
        {
            var docx = BuildDocx(Item("One") + Item("Two"), Decimal + Num(1, 0));
            Assert.Equal(
                Convert(docx, semanticLists: false, pagination: PaginationMode.Paginated).ToString(),
                Convert(docx, semanticLists: true, pagination: PaginationMode.Paginated).ToString());
        }

        [Fact]
        public void The_facade_carries_the_setting_and_anchors_land_on_items()
        {
            var html = HtmlConversionOps.ConvertToHtml(BuildDocx(Item("One") + Item("Two"), Decimal + Num(1, 0)),
                new HtmlConversionOptions { SemanticLists = true, StampAnchors = true });
            var items = XElement.Parse(html).Descendants().Where(e => e.Name.LocalName == "li").ToList();

            Assert.Equal(2, items.Count);
            Assert.All(items, li => Assert.NotNull(li.Attribute("data-anchor")));
        }
    }
}
