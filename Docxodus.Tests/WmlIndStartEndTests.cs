// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests
{
    /// <summary>
    /// Issue #894: ECMA-376 spells a paragraph's indent edges <c>w:start</c>/<c>w:end</c> as well as
    /// the transitional <c>w:left</c>/<c>w:right</c>. LibreOffice writes the former, and the HTML
    /// converter read only the latter, so every indent of such a document converted to 0.
    /// </summary>
    public class WmlIndStartEndTests
    {
        private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace W = WNs;

        private static byte[] BuildDocx(string bodyXml, string stylesXml = "", string? numberingXml = null)
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
                    $"<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>{stylesXml}</w:styles>");
                Write(main.AddNewPart<DocumentSettingsPart>(),
                    $"<w:settings xmlns:w=\"{WNs}\"><w:defaultTabStop w:val=\"720\"/></w:settings>");
                if (numberingXml != null)
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

        private static string Para(string ind, string text, string extraPPr = "") =>
            $"<w:p><w:pPr>{extraPPr}{ind}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

        private static string ListNumbering(string lvlInd) =>
            "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>" +
            $"<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/><w:pPr>{lvlInd}</w:pPr></w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";

        private const string ListPPr = "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>";

        private static XElement ToHtml(byte[] docx) =>
            WmlToHtmlConverter.ConvertToHtml(new WmlDocument("t.docx", docx), new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
            });

        /// <summary>The inline style of the HTML paragraph whose text contains <paramref name="text"/>.</summary>
        private static string StyleOf(XElement html, string text) =>
            (string?)html.Descendants().First(e => e.Name.LocalName == "p" && e.Value.Contains(text)).Attribute("style") ?? "";

        [Theory]
        [InlineData("left")]
        [InlineData("start")]
        public void Paragraph_leading_indent_becomes_margin_left(string spelling)
        {
            var html = ToHtml(BuildDocx(Para($"<w:ind w:{spelling}=\"1440\"/>", "Indented")));
            Assert.Contains("margin-left: 1.00in", StyleOf(html, "Indented"));
        }

        [Theory]
        [InlineData("right")]
        [InlineData("end")]
        public void Paragraph_trailing_indent_becomes_margin_right(string spelling)
        {
            var html = ToHtml(BuildDocx(Para($"<w:ind w:{spelling}=\"720\"/>", "Indented")));
            Assert.Contains("margin-right: 0.50in", StyleOf(html, "Indented"));
        }

        [Theory]
        [InlineData("left")]
        [InlineData("start")]
        public void List_level_indent_survives_conversion(string spelling)
        {
            var html = ToHtml(BuildDocx(
                Para("", "First item", ListPPr),
                numberingXml: ListNumbering($"<w:ind w:{spelling}=\"720\" w:hanging=\"360\"/>")));
            var style = StyleOf(html, "First item");
            Assert.Contains("margin-left: 0.50in", style);
            Assert.Contains("text-indent: -0.25in", style);
        }

        [Theory]
        [InlineData("start", "left")]
        [InlineData("left", "start")]
        public void Paragraph_overrides_its_style_in_the_other_spelling(string styleSpelling, string paraSpelling)
        {
            var styles = "<w:style w:type=\"paragraph\" w:styleId=\"Deep\"><w:name w:val=\"Deep\"/>" +
                $"<w:pPr><w:ind w:{styleSpelling}=\"2880\"/></w:pPr></w:style>";
            var html = ToHtml(BuildDocx(
                Para($"<w:ind w:{paraSpelling}=\"720\"/>", "Override", "<w:pStyle w:val=\"Deep\"/>"), styles));
            Assert.Contains("margin-left: 0.50in", StyleOf(html, "Override"));
        }

        [Fact]
        public void Start_wins_when_one_ind_carries_both_spellings()
        {
            var html = ToHtml(BuildDocx(Para("<w:ind w:left=\"2880\" w:start=\"720\" w:right=\"2880\" w:end=\"360\"/>", "Both")));
            var style = StyleOf(html, "Both");
            Assert.Contains("margin-left: 0.50in", style);
            Assert.Contains("margin-right: 0.25in", style);
        }

        /// <summary>
        /// Covers every converter read site at once — paragraph margins, the tab layout that starts at
        /// the leading indent, a list level, and a bordered paragraph group — and requires the
        /// <c>w:start</c>/<c>w:end</c> document to convert exactly as its <c>w:left</c>/<c>w:right</c> twin.
        /// </summary>
        [Fact]
        public void Strict_spelling_converts_exactly_like_transitional_spelling()
        {
            string Body(string lead, string trail) =>
                Para($"<w:ind w:{lead}=\"1440\" w:{trail}=\"720\"/>", "Plain") +
                $"<w:p><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"2880\"/></w:tabs><w:ind w:{lead}=\"1440\" w:hanging=\"720\"/></w:pPr>" +
                "<w:r><w:t>Term</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Definition</w:t></w:r></w:p>" +
                Para("", "Listed", ListPPr) +
                Para($"<w:ind w:{lead}=\"720\" w:hanging=\"360\"/>", "Boxed",
                    "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"000000\"/></w:pBdr>");

            string Convert(string lead, string trail) => ToHtml(BuildDocx(
                Body(lead, trail),
                numberingXml: ListNumbering($"<w:ind w:{lead}=\"720\" w:hanging=\"360\"/>"))).ToString();

            Assert.Equal(Convert("left", "right"), Convert("start", "end"));
        }

        [Theory]
        [InlineData("left")]
        [InlineData("start")]
        public void AssembleFormatting_puts_the_list_tab_stop_at_the_leading_indent(string spelling)
        {
            var docx = BuildDocx(
                Para("", "Listed", ListPPr),
                numberingXml: ListNumbering($"<w:ind w:{spelling}=\"720\" w:hanging=\"360\"/>"));
            var assembled = FormattingAssembler.AssembleFormatting(new WmlDocument("t.docx", docx), new FormattingAssemblerSettings());

            using var ms = new MemoryStream(assembled.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var body = XDocument.Load(wDoc.MainDocumentPart!.GetStream()).Root!;
            var positions = body.Descendants(W + "p").First().Elements(W + "pPr").Elements(W + "tabs").Elements(W + "tab")
                .Select(t => (string?)t.Attribute(W + "pos")).ToList();
            Assert.Contains("720", positions);
        }

        [Fact]
        public void GetFormatting_reads_the_edge_the_renderer_reads()
        {
            using var s = new DocxSession(BuildDocx(Para("<w:ind w:left=\"2880\" w:start=\"720\" w:end=\"360\"/>", "Both")));
            var anchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
            var paragraph = s.GetFormatting(anchor)!.DirectParagraph;
            Assert.Equal(720, paragraph.LeftIndentTwips);
            Assert.Equal(360, paragraph.RightIndentTwips);
        }

        [Theory]
        [InlineData("start", "left")]
        [InlineData("left", "start")]
        public void Effective_formatting_takes_the_paragraph_override_in_the_other_spelling(string styleSpelling, string paraSpelling)
        {
            var styles = "<w:style w:type=\"paragraph\" w:styleId=\"Deep\"><w:name w:val=\"Deep\"/>" +
                $"<w:pPr><w:ind w:{styleSpelling}=\"2880\"/></w:pPr></w:style>";
            using var s = new DocxSession(BuildDocx(
                Para($"<w:ind w:{paraSpelling}=\"720\"/>", "Override", "<w:pStyle w:val=\"Deep\"/>"), styles));
            var anchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
            Assert.Equal(720, s.GetFormatting(anchor)!.EffectiveParagraph.LeftIndentTwips);
        }

        [Theory]
        [InlineData("start", "left")]
        [InlineData("left", "start")]
        public void List_paragraph_overrides_its_level_indent_in_the_other_spelling(string levelSpelling, string paraSpelling)
        {
            var html = ToHtml(BuildDocx(
                Para($"<w:ind w:{paraSpelling}=\"1440\" w:hanging=\"360\"/>", "Moved item", ListPPr),
                numberingXml: ListNumbering($"<w:ind w:{levelSpelling}=\"720\" w:hanging=\"360\"/>")));
            Assert.Contains("margin-left: 1.00in", StyleOf(html, "Moved item"));
        }

        [Fact]
        public void GetListMembership_reads_a_strict_spelling_level_indent()
        {
            using var s = new DocxSession(BuildDocx(
                Para("", "Listed", ListPPr),
                numberingXml: ListNumbering("<w:ind w:left=\"2880\" w:start=\"720\" w:hanging=\"360\"/>")));
            var anchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("li:"));
            Assert.Equal(720, s.GetListMembership(anchor)!.LeftIndentTwips);
        }

        [Fact]
        public void DocxDiff_reads_the_edge_the_renderer_reads()
        {
            var format = Docxodus.Ir.IrReader.MapParaFormat(XElement.Parse(
                $"<w:pPr xmlns:w=\"{WNs}\"><w:ind w:left=\"2880\" w:start=\"720\" w:right=\"2880\" w:end=\"360\"/></w:pPr>"));
            Assert.Equal(720, format.IndentLeftTwips);
            Assert.Equal(360, format.IndentRightTwips);
        }

        [Theory]
        [InlineData("left")]
        [InlineData("start")]
        public void IndentDelta_adjusts_the_spelling_the_paragraph_uses(string spelling)
        {
            using var s = new DocxSession(BuildDocx(Para($"<w:ind w:{spelling}=\"720\"/>", "Shift")));
            var anchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));
            Assert.True(s.SetParagraphFormat(anchor, new ParagraphFormatOp { IndentDelta = 720 }).Success);

            var ind = XElement.Parse(s.Raw.GetXml(anchor)).Descendants(W + "ind").Single();
            Assert.Equal("1440", (string?)ind.Attribute(W + spelling));
            Assert.Single(ind.Attributes(), a => a.Name == W + "left" || a.Name == W + "start");
            Assert.Contains("margin-left: 1.00in", StyleOf(ToHtml(s.Save()), "Shift"));
        }
    }
}
