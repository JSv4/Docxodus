using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class FontTableFallbackTests
{
    [Theory]
    [InlineData("Oriel One", "<w:family w:val='swiss'/>", "'Oriel One', sans-serif")]
    [InlineData("Oriel Two", "<w:family w:val='roman'/>", "'Oriel Two', serif")]
    [InlineData("Oriel Three", "<w:family w:val='modern'/>", "'Oriel Three', monospace")]
    [InlineData("Oriel Four", "<w:family w:val='swiss'/><w:pitch w:val='fixed'/>", "'Oriel Four', monospace")]
    [InlineData("Oriel Five", "<w:panose1 w:val='020B0604020202020204'/>", "'Oriel Five', sans-serif")]
    [InlineData("Oriel Six", "<w:panose1 w:val='02020603050405020304'/>", "'Oriel Six', serif")]
    [InlineData("Oriel Seven", "<w:panose1 w:val='02070309020205020404'/>", "'Oriel Seven', monospace")]
    [InlineData("Oriel Eight", "<w:panose1 w:val='02070309020205020404'/><w:family w:val='swiss'/>", "'Oriel Eight', sans-serif")]
    [InlineData("Oriel Sans", "<w:family w:val='roman'/>", "'Oriel Sans', serif")]
    [InlineData("Oriel Mono", "<w:family w:val='swiss'/>", "'Oriel Mono', sans-serif")]
    [InlineData("Oriel Sans", "<w:panose1 w:val='02020603050405020304'/><w:family w:val='script'/>", "'Oriel Sans', sans-serif")]
    [InlineData("Oriel Mono", "<w:panose1 w:val='020B0604020202020204'/><w:family w:val='decorative'/>", "'Oriel Mono', monospace")]
    [InlineData("Oriel Sans", "<w:panose1 w:val='050B0604020202020204'/>", "'Oriel Sans', sans-serif")]
    [InlineData("Oriel Sans", "<w:panose1 w:val='00000000000000000000'/>", "'Oriel Sans', sans-serif")]
    [InlineData("Oriel Sans", "<w:panose1 w:val='not-a-panose'/>", "'Oriel Sans', sans-serif")]
    [InlineData("Oriel Mono", "", "'Oriel Mono', monospace")]
    [InlineData("Oriel Nine", "", "'Oriel Nine', serif")]
    [InlineData("Oriel Ten", "<w:altName w:val='Liberation Sans'/><w:family w:val='swiss'/>", "'Oriel Ten', 'Liberation Sans', sans-serif")]
    [InlineData("Oriel's Serif", "<w:altName w:val='Oriel\\Text'/><w:family w:val='roman'/>", "'Oriel\\'s Serif', 'Oriel\\\\Text', serif")]
    public void BodyAndRunningStoriesUseDeclaredFontClass(string family, string properties, string expected)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("fonts.docx", Document(family, properties)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true });
        foreach (var text in new[] { "Body", "Header", "Footer" })
        {
            var run = Assert.Single(html.Descendants(Xhtml.span), e => e.Value == text);
            var style = run.AncestorsAndSelf().Select(e => (string?)e.Attribute("style"))
                .First(s => s?.Contains("font-family:") == true)!;
            Assert.Contains("font-family: " + expected + ";", style);
        }
    }

    [Fact]
    public void ThemeFontsResolveBeforeTheFontTableLookup()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("theme.docx",
            Document("Oriel Theme", "<w:family w:val='swiss'/>", theme: true)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        Assert.Contains("'Oriel Theme', sans-serif", html.ToString());
    }

    [Fact]
    public void DeclaredFixedPitchAndNamedMonospaceChooseTheSameTabStop()
    {
        string Width(string family, string properties)
        {
            var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("tabs.docx",
                Document(family, properties, tabs: true)), new WmlToHtmlConverterSettings { FabricateCssClasses = false });
            var segment = Assert.Single(html.Descendants(), e => e.Value.Contains("iiii")
                && ((string?)e.Attribute("style"))?.Contains("display: inline-flex") == true);
            return ((string)segment.Attribute("style")!).Split(';').Single(s => s.TrimStart().StartsWith("width:")).Trim();
        }
        var control = Width("Oriel Mono", "");
        Assert.Equal("width: 36pt", control);
        Assert.Equal(control, Width("Oriel Declared", "<w:pitch w:val='fixed'/>"));
    }

    [Theory]
    [InlineData("swiss", "sans-serif")]
    [InlineData("roman", "serif")]
    [InlineData("modern", "monospace")]
    public void BalloonAndCommentMetadataCarryTheSameFontStack(string family, string generic)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("balloons.docx",
            Document("Oriel Markup", $"<w:family w:val='{family}'/>", markup: true)),
            new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated,
                RenderComments = true, CommentRenderMode = CommentRenderMode.Margin,
                RevisionPresentation = RevisionPresentation.Word });
        Assert.Equal("'Oriel Markup', " + generic,
            Assert.Single(html.Descendants().Attributes("data-balloon-font-stack")).Value);
        Assert.Equal("'Oriel Markup', " + generic,
            Assert.Single(html.Descendants().Attributes("data-font-stack")).Value);
    }

    private static byte[] Document(string family, string properties, bool theme = false, bool markup = false, bool tabs = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var fonts = XDocument.Parse($"<w:fonts xmlns:w='{W.w}'><w:font>{properties}</w:font></w:fonts>");
            fonts.Root!.Element(W.font)!.SetAttributeValue(W.name, family);
            main.AddNewPart<FontTablePart>().PutXDocument(fonts);
            if (theme)
            {
                var xml = XDocument.Parse(WordStockTheme.Xml);
                foreach (var latin in xml.Descendants(A.latin)) latin.SetAttributeValue("typeface", family);
                main.AddNewPart<ThemePart>().PutXDocument(xml);
            }
            Wp.Paragraph Paragraph(string text) => new(new Wp.Run(
                new Wp.RunProperties(theme
                    ? new Wp.RunFonts { AsciiTheme = Wp.ThemeFontValues.MinorHighAnsi, HighAnsiTheme = Wp.ThemeFontValues.MinorHighAnsi }
                    : new Wp.RunFonts { Ascii = family, HighAnsi = family }), new Wp.Text(text)));
            main.AddNewPart<HeaderPart>("header").Header = new Wp.Header(Paragraph("Header"));
            main.AddNewPart<FooterPart>("footer").Footer = new Wp.Footer(Paragraph("Footer"));
            var bodyParagraph = Paragraph("Body");
            if (tabs)
            {
                bodyParagraph.PrependChild(new Wp.ParagraphProperties(new Wp.Tabs(
                    new Wp.TabStop { Val = Wp.TabStopValues.Left, Position = 360 },
                    new Wp.TabStop { Val = Wp.TabStopValues.Left, Position = 720 })));
                var run = bodyParagraph.GetFirstChild<Wp.Run>()!;
                run.RunProperties!.Append(new Wp.FontSize { Val = "22" });
                run.GetFirstChild<Wp.Text>()!.Text = "iiii";
                run.Append(new Wp.TabChar(), new Wp.Text("After"));
            }
            if (markup)
            {
                main.StyleDefinitionsPart!.Styles!.Append(new Wp.Style(new Wp.StyleName { Val = "Balloon Text" },
                    new Wp.StyleRunProperties(new Wp.RunFonts { Ascii = family, HighAnsi = family }))
                    { Type = Wp.StyleValues.Paragraph, StyleId = "BalloonText" });
                main.AddNewPart<WordprocessingCommentsPart>().Comments = new Wp.Comments(
                    new Wp.Comment(Paragraph("Comment")) { Id = "0", Author = "Oriel Reviewer", Initials = "OR" });
                bodyParagraph.PrependChild(new Wp.CommentRangeStart { Id = "0" });
                bodyParagraph.Append(new Wp.CommentRangeEnd { Id = "0" }, new Wp.Run(new Wp.CommentReference { Id = "0" }));
            }
            main.Document = new Wp.Document(new Wp.Body(bodyParagraph, new Wp.SectionProperties(
                new Wp.HeaderReference { Id = "header", Type = Wp.HeaderFooterValues.Default },
                new Wp.FooterReference { Id = "footer", Type = Wp.HeaderFooterValues.Default })));
        }
        return stream.ToArray();
    }
}
