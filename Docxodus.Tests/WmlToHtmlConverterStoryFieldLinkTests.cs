using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>A complex field converts as a field in every story the converter renders, not only in the body: a
/// HYPERLINK field in a footnote, an endnote, a header or a footer is a link (issue #1876).</summary>
public class WmlToHtmlConverterStoryFieldLinkTests
{
    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Link(string name) => $"""
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.test/{name}" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>{Run(name)}<w:r><w:fldChar w:fldCharType="end"/></w:r>
        """;

    private const string Notes = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static void Put(OpenXmlPart part, string xml) => part.PutXDocument(XDocument.Parse(xml));

    private static byte[] Docx(string headerPrefix = "")
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var ns = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\"";
            Put(main.AddNewPart<HeaderPart>("rIdHeader"), $"<w:hdr {ns}><w:p>{headerPrefix}{Link("HeaderLink")}</w:p></w:hdr>");
            Put(main.AddNewPart<FooterPart>("rIdFooter"), $"<w:ftr {ns}><w:p>{Link("FooterLink")}</w:p></w:ftr>");
            Put(main.AddNewPart<FootnotesPart>(), $"<w:footnotes {ns}>{Notes}<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r>{Link("FootnoteLink")}</w:p></w:footnote></w:footnotes>");
            Put(main.AddNewPart<EndnotesPart>(), $"<w:endnotes {ns}>{Notes.Replace("footnote", "endnote")}<w:endnote w:id=\"1\"><w:p><w:r><w:endnoteRef/></w:r>{Link("EndnoteLink")}</w:p></w:endnote></w:endnotes>");
            Put(main, $"""
                <w:document {ns}><w:body>
                  <w:p>{Link("BodyLink")}<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>
                  <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/>
                    <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
                </w:body></w:document>
                """);
            Put(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}/>");
            Put(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>");
        }
        return stream.ToArray();
    }

    private static XElement Render(PaginationMode mode, string headerPrefix = "") =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("stories.docx", Docx(headerPrefix)), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            RenderPagination = mode,
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = true,
        });

    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void AHyperlinkFieldInAnyStoryIsALink(PaginationMode mode)
    {
        var html = Render(mode);
        foreach (var name in new[] { "BodyLink", "FootnoteLink", "EndnoteLink", "HeaderLink", "FooterLink" })
        {
            var text = html.Descendants().Where(e => e.Name.LocalName == "span" && !e.HasElements && e.Value == name).ToList();
            Assert.NotEmpty(text);
            Assert.All(text, span => Assert.Equal($"https://example.test/{name}",
                (string?)span.Ancestors().FirstOrDefault(a => a.Name.LocalName == "a" && a.Attribute("href") != null)?.Attribute("href")));
        }
    }

    // A stray field character with no field open (malformed markup) in one story is ignored: the rest still converts,
    // and the next field is still a link.
    [Theory]
    [InlineData("end")]
    [InlineData("separate")]
    public void AStrayFieldCharacterInAStoryDoesNotStopItsFields(string type)
    {
        var html = Render(PaginationMode.None, $"<w:r><w:fldChar w:fldCharType=\"{type}\"/></w:r>{Run("Stray ")}");
        var span = Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "span" && !e.HasElements && e.Value == "HeaderLink"));
        Assert.Equal("https://example.test/HeaderLink", (string?)span.Ancestors().First(a => a.Name.LocalName == "a").Attribute("href"));
        Assert.Contains("Stray", html.Value);
    }

    // Each note and text box is a story of its own: a field left open in one (malformed markup) does not run into the
    // next note, or out of the text box into the paragraphs after it.
    private const string Unclosed = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        + "<w:r><w:instrText xml:space=\"preserve\"> HYPERLINK \"https://example.test/open\" </w:instrText></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>Open</w:t></w:r>";

    private static string TextBox(string paragraph) => "<w:r><w:pict><v:shape style=\"width:120pt;height:40pt\"><v:textbox>"
        + $"<w:txbxContent><w:p>{paragraph}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>";

    private static XElement RenderStories(string body, string firstNote, string secondNote, string header, PaginationMode mode)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var ns = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\" xmlns:v=\"urn:schemas-microsoft-com:vml\"";
            Put(main.AddNewPart<HeaderPart>("rIdHeader"), $"<w:hdr {ns}>{header}</w:hdr>");
            Put(main.AddNewPart<FootnotesPart>(), $"<w:footnotes {ns}>{Notes}"
                + $"<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r>{firstNote}</w:p></w:footnote>"
                + $"<w:footnote w:id=\"2\"><w:p><w:r><w:footnoteRef/></w:r>{secondNote}</w:p></w:footnote></w:footnotes>");
            Put(main, $"""
                <w:document {ns}><w:body>{body}
                  <w:p>{Run("Cited")}<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
                  <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>
                    <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
                </w:body></w:document>
                """);
            Put(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}/>");
            Put(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>");
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("stories.docx", stream.ToArray()), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            RenderPagination = mode,
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = true,
        });
    }

    private static void AssertNotLinked(XElement html, string text)
    {
        var spans = html.Descendants().Where(e => e.Name.LocalName == "span" && !e.HasElements && e.Value == text).ToList();
        Assert.NotEmpty(spans);
        Assert.All(spans, span => Assert.DoesNotContain(span.Ancestors(), a => a.Name.LocalName == "a" && a.Attribute("href") != null));
    }

    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void AFieldLeftOpenInOneStoryDoesNotRunIntoTheNext(PaginationMode mode)
    {
        var html = RenderStories(
            body: $"<w:p>{TextBox(Unclosed)}</w:p><w:p>{Run("BodyAfterBox")}</w:p>",
            firstNote: Unclosed,
            secondNote: Run("SecondNoteBefore") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" + Run("SecondNoteAfter"),
            header: $"<w:p>{TextBox(Unclosed)}{Run("HeaderHost")}</w:p><w:p>{Run("HeaderAfterBox")}</w:p>",
            mode);
        foreach (var text in new[] { "BodyAfterBox", "SecondNoteBefore", "SecondNoteAfter", "HeaderHost", "HeaderAfterBox", "Cited" })
            AssertNotLinked(html, text);
    }

    // A document with a stray field character in its body, as a rejected field insertion can leave it, converts.
    [Fact]
    public void ADocumentWithAStrayFieldCharacterConverts()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument(Path.Combine("..", "..", "..", "..", "TestFiles", "RP", "RP020-Inserted-Field-Code-Rejected.docx")),
            new WmlToHtmlConverterSettings());
        Assert.Contains("Test", html.Descendants().First(e => e.Name.LocalName == "body").Value);
    }
}
