using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>A HYPERLINK field's result is a link in the same way a <c>w:hyperlink</c> is: it shows its
/// runs' own formatting, not the browser's default link colour and underline (issue #1869).</summary>
public class WmlToHtmlConverterFieldLinkStyleTests
{
    private const string FieldUrl = "https://example.test/field";
    private const string ExternalUrl = "https://example.test/external";

    private static string Field(string result) => $"""
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> HYPERLINK "{FieldUrl}" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>
        """;

    private static string Run(string text, string rPr = "") => $"<w:r>{rPr}<w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private const string HyperlinkStyle = "<w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr>";

    // A field link; one whose result spans a tab (the converter keeps it whole); the two w:hyperlink
    // forms (a relationship target and a bookmark target) for comparison; and the bookmark.
    private static readonly string Links =
        $"<w:p>{Field(Run("FieldLink"))}</w:p>"
        + $"<w:p>{Field(Run("Tabbed") + "<w:r><w:tab/></w:r>" + Run("Link"))}</w:p>"
        + $"<w:p><w:hyperlink r:id=\"rIdLink\">{Run("ExternalLink")}</w:hyperlink></w:p>"
        + $"<w:p><w:hyperlink w:anchor=\"Target\">{Run("BookmarkLink")}</w:hyperlink></w:p>"
        + $"<w:p><w:bookmarkStart w:id=\"1\" w:name=\"Target\"/>{Run("Target")}<w:bookmarkEnd w:id=\"1\"/></w:p>";

    private static byte[] Docx(string body)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddHyperlinkRelationship(new Uri(ExternalUrl), true, "rIdLink");
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}" xmlns:r="{R.r}"><w:body>{body}
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
                    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"""
                <w:styles xmlns:w="{W.w}">
                  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>
                    <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
                </w:styles>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        return stream.ToArray();
    }

    private static XElement Render(string body, bool fabricateCssClasses, PaginationMode mode) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", Docx(body)), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = fabricateCssClasses,
            RenderPagination = mode,
            RenderTrackedChanges = true,
        });

    /// <summary>The declarations an element gets from its fabricated classes and its style attribute, sorted.</summary>
    private static string Declarations(XElement html, XElement element)
    {
        var css = string.Concat(html.Descendants(Xhtml.style).Select(s => s.Value));
        var fromClasses = ((string?)element.Attribute("class") ?? "")
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .SelectMany(name => Regex.Matches(css, $@"(?m)^{element.Name.LocalName}\.{Regex.Escape(name)} \{{([^}}]*)\}}")
                .SelectMany(rule => rule.Groups[1].Value.Split(';')));
        var fromStyle = ((string?)element.Attribute("style") ?? "").Split(';');
        return string.Join("; ", fromClasses.Concat(fromStyle).Select(d => Regex.Replace(d.Trim(), @"\s*:\s*", ": "))
            .Where(d => d.Length > 0).OrderBy(d => d, StringComparer.Ordinal));
    }

    /// <summary>A link that leaves colour and underline to its runs, as a w:hyperlink's does.</summary>
    private static Action<XElement> AssertNeutral(XElement html) => link =>
    {
        var declarations = Declarations(html, link);
        Assert.Contains("color: inherit", declarations);
        Assert.Contains("text-decoration: none", declarations);
    };

    private static XElement Link(XElement html, string href) =>
        Assert.Single(html.Descendants(Xhtml.a), a => (string?)a.Attribute("href") == href);

    private static XElement TextSpan(XElement link, string text) =>
        link.Descendants(Xhtml.span).Single(s => !s.HasElements && s.Value == text);

    [Theory]
    [InlineData(true, PaginationMode.None)]
    [InlineData(true, PaginationMode.Paginated)]
    [InlineData(false, PaginationMode.None)]
    [InlineData(false, PaginationMode.Paginated)]
    public void AHyperlinkFieldsLinkIsStyledAsAWordHyperlinkIs(bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render(Links, fabricateCssClasses, mode);
        var fieldLinks = html.Descendants(Xhtml.a).Where(a => (string?)a.Attribute("href") == FieldUrl).ToList();
        Assert.Equal(2, fieldLinks.Count);
        Assert.Equal("FieldLink", fieldLinks[0].Value);
        Assert.Matches("^Tabbed.*Link$", fieldLinks[1].Value);
        var external = Link(html, ExternalUrl);
        var bookmark = Link(html, "#Target");
        Assert.All(new[] { external, bookmark }.Concat(fieldLinks), AssertNeutral(html));
        // One mechanism: the same class (or the same inline style) as the w:hyperlink's.
        Assert.All(fieldLinks, a => Assert.Equal(Declarations(html, bookmark), Declarations(html, a)));
        Assert.All(fieldLinks, a => Assert.Equal((string?)bookmark.Attribute(fabricateCssClasses ? "class" : "style"),
            (string?)a.Attribute(fabricateCssClasses ? "class" : "style")));
    }

    [Theory]
    [InlineData(true, PaginationMode.None)]
    [InlineData(true, PaginationMode.Paginated)]
    [InlineData(false, PaginationMode.None)]
    [InlineData(false, PaginationMode.Paginated)]
    public void AHyperlinkFieldsRunsKeepTheirOwnFormatting(bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render($"<w:p>{Field(Run("Styled", HyperlinkStyle) + Run("Plain"))}</w:p>", fabricateCssClasses, mode);
        var link = Link(html, FieldUrl);
        // The link leaves colour and underline to its runs, so the browser's own link style does not paint the plain one.
        AssertNeutral(html)(link);
        var styled = Declarations(html, TextSpan(link, "Styled"));
        Assert.Contains("color: #0563C1", styled);
        Assert.Contains("text-decoration: underline", styled);
        var plain = Declarations(html, TextSpan(link, "Plain"));
        Assert.DoesNotContain("color:", plain);
        Assert.DoesNotContain("text-decoration:", plain);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AnInsertionInAHyperlinkFieldsResultStaysInsideTheStyledLink(bool fabricateCssClasses)
    {
        var html = Render($"""
            <w:p>{Field(Run("Kept ") + "<w:ins w:id=\"5\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\">" + Run("Inserted") + "</w:ins>")}</w:p>
            """, fabricateCssClasses, PaginationMode.None);
        var link = Link(html, FieldUrl);
        AssertNeutral(html)(link);
        Assert.Equal("Inserted", Assert.Single(link.Descendants(Xhtml.ins)).Value);
    }

    // The editor renders one block at a time, with references stamped; its field link is styled the same.
    [Fact]
    public void ABlockRenderStylesAHyperlinkFieldsLinkAsTheFullRenderDoes()
    {
        var bytes = Docx(Links);
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true, StampReferenceFields = true };
        using var session = new DocxSession(bytes);
        var anchor = session.Project().AnchorIndex.Single(kv => kv.Value.TextPreview.Contains("FieldLink")).Key;
        var block = XElement.Parse(HtmlConversionOps.RenderBlockHtml(session, anchor, options));
        AssertNeutral(block)(Link(block, FieldUrl));
    }

    private static string FieldWith(string instruction, string result) => $"""
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>
        """;

    // A HYPERLINK field's \l names a place in the document, so the link goes there: after the address if there is one.
    [Theory]
    [InlineData("HYPERLINK \\l \"Target\"", "#Target")]
    [InlineData("HYPERLINK \"https://example.test/doc\" \\l \"Section\"", "https://example.test/doc#Section")]
    [InlineData("HYPERLINK \"https://example.test/doc\" \\o \"A tip\"", "https://example.test/doc")] // control: \o is a tooltip
    [InlineData("HYPERLINK \\l \"Target\" \\o \"A tip\"", "#Target")]           // a tooltip after the location
    [InlineData("HYPERLINK \\o \"A tip\" \"https://example.test/doc\"", "https://example.test/doc")] // a tooltip first
    [InlineData("HYPERLINK \\l \"Target\" \\* MERGEFORMAT", "#Target")]           // a general switch's value
    [InlineData("HYPERLINK \"\\\\server\\share\\a.docx\"", "\\\\server\\share\\a.docx")]          // a network path, not a switch
    public void AHyperlinkFieldLinksToTheLocationItNames(string instruction, string href)
    {
        var html = Render($"<w:p>{FieldWith(instruction, Run("Go"))}</w:p>", fabricateCssClasses: false, PaginationMode.None);
        Assert.Equal("Go", Link(html, href).Value);
    }

    // A HYPERLINK field in another's result gives up its own link there too, as a w:hyperlink does: an anchor
    // cannot hold another.
    [Fact]
    public void AHyperlinkFieldInAnothersResultIsNotALinkInsideALink()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx",
                Docx($"<w:p>{Field(Run("Outer ") + FieldWith("HYPERLINK \"https://example.test/inner\"", Run("Inner")))}</w:p>")),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, StampReferenceFields = true });
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => a.Ancestors(Xhtml.a).Any());
    }
}
