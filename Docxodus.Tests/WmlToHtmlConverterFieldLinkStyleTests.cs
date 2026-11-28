using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir;
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
    // The IR reads the same target (one parser serves both), so the markdown projection and the diff agree with the HTML.
    [Theory]
    [InlineData("HYPERLINK \\l \"Target\"", "#Target")]
    [InlineData("HYPERLINK \"https://example.test/doc\" \\l \"Section\"", "https://example.test/doc#Section")]
    [InlineData("HYPERLINK \"https://example.test/doc\" \\o \"A tip\"", "https://example.test/doc")] // control: \o is a tooltip
    [InlineData("HYPERLINK \\l \"Target\" \\o \"A tip\"", "#Target")]           // a tooltip after the location
    [InlineData("HYPERLINK \\o \"A tip\" \"https://example.test/doc\"", "https://example.test/doc")] // a tooltip first
    [InlineData("HYPERLINK \\l \"Target\" \\* MERGEFORMAT", "#Target")]           // a general switch's value
    [InlineData("HYPERLINK \"\\\\\\\\server\\\\share\\\\a.docx\"", "\\\\server\\share\\a.docx")] // a network path as Word escapes it
    [InlineData("HYPERLINK \\o \"say \\\"hi\\\"\" \"https://example.test/doc\"", "https://example.test/doc")] // an escaped quote
    [InlineData("HYPERLINK \\m \"https://example.test/doc\"", "https://example.test/doc")] // \m takes no value
    [InlineData("HYPERLINK 'https://example.test/doc'", "https://example.test/doc")] // single quotes, as on main
    [InlineData("HYPERLINK \\o \"\" \"https://example.test/doc\"", "https://example.test/doc")]  // an empty tooltip first
    [InlineData("HYPERLINK \"https://example.test/doc\" \\l \"\"", "https://example.test/doc")]  // an empty location
    [InlineData("HYPERLINK \\l \"Target\" \\*MERGEFORMAT", "#Target")]           // a switch with its value glued on
    [InlineData("HYPERLINK \\l\"Target\"", "#Target")]
    [InlineData("HYPERLINK \\o\"A tip\" \"https://example.test/doc\"", "https://example.test/doc")]
    [InlineData("HYPERLINK \\t\"_blank\" \"https://example.test/doc\"", "https://example.test/doc")]
    public void AHyperlinkFieldLinksToTheLocationItNames(string instruction, string href)
    {
        var body = $"<w:p>{FieldWith(instruction, Run("Go"))}</w:p>";
        var html = Render(body, fabricateCssClasses: false, PaginationMode.None);
        Assert.Equal("Go", Link(html, href).Value);
        var paragraph = Assert.Single(IrReader.Read(new WmlDocument("field.docx", Docx(body))).Body.Blocks.OfType<IrParagraph>());
        Assert.Equal(href, Assert.Single(paragraph.Inlines.OfType<IrHyperlink>()).Target);
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

    // Issue #1870: the simple field form, w:fldSimple, is a link too, and a link written as a field inside a table of
    // contents is drawn as Word draws a w:hyperlink there.

    private static string Simple(string instruction, string result) =>
        $"<w:fldSimple w:instr=\"{instruction.Replace("\"", "&quot;")}\">{result}</w:fldSimple>";

    /// <summary>A complex field's start: its begin character, instruction and separator.</summary>
    private static string Begin(string instruction) => $"""
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        """;

    private const string EndOfField = "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    /// <summary>The one link that holds the text, which no other link holds.</summary>
    private static XElement LinkOf(XElement html, string text) =>
        Assert.Single(TextSpanIn(html, text).Ancestors(Xhtml.a));

    private static XElement TextSpanIn(XElement html, string text) =>
        Assert.Single(html.Descendants(Xhtml.span), s => !s.HasElements && s.Value == text);

    [Theory]
    [InlineData(true, PaginationMode.None)]
    [InlineData(true, PaginationMode.Paginated)]
    [InlineData(false, PaginationMode.None)]
    [InlineData(false, PaginationMode.Paginated)]
    public void ASimpleHyperlinkFieldIsALinkStyledAsAWordHyperlinkIs(bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render(Links + $"<w:p>{Simple($" HYPERLINK \"{FieldUrl}\" ", Run("Styled", HyperlinkStyle) + Run("Plain"))}</w:p>",
            fabricateCssClasses, mode);
        var link = LinkOf(html, "Plain");
        Assert.Equal(FieldUrl, (string?)link.Attribute("href"));
        Assert.Equal("StyledPlain", link.Value);
        // The same link as the w:hyperlink's, and the runs keep their own formatting inside it.
        Assert.Equal(Declarations(html, Link(html, "#Target")), Declarations(html, link));
        var styled = Declarations(html, TextSpan(link, "Styled"));
        Assert.Contains("color: #0563C1", styled);
        Assert.Contains("text-decoration: underline", styled);
        var plain = Declarations(html, TextSpan(link, "Plain"));
        Assert.DoesNotContain("color:", plain);
        Assert.DoesNotContain("text-decoration:", plain);
    }

    // The simple form reads its target with the complex form's parser, and the IR agrees with the HTML.
    [Theory]
    [InlineData(" HYPERLINK \\l \"Target\" ", "#Target")]
    [InlineData(" HYPERLINK \"https://example.test/doc\" \\l \"Section\" \\o \"A tip\" ", "https://example.test/doc#Section")]
    public void ASimpleHyperlinkFieldLinksToTheLocationItNames(string instruction, string href)
    {
        var body = $"<w:p>{Simple(instruction, Run("Go"))}</w:p>";
        var html = Render(body, fabricateCssClasses: false, PaginationMode.None);
        Assert.Equal("Go", Link(html, href).Value);
        var paragraph = Assert.Single(IrReader.Read(new WmlDocument("field.docx", Docx(body))).Body.Blocks.OfType<IrParagraph>());
        Assert.Equal(href, Assert.Single(paragraph.Inlines.OfType<IrHyperlink>()).Target);
    }

    // Control: a simple field of another kind is not a link.
    [Fact]
    public void ASimpleFieldOfAnotherKindIsNotALink()
    {
        var html = Render($"<w:p>{Simple(" AUTHOR ", Run("Someone"))}</w:p>", fabricateCssClasses: false, PaginationMode.None);
        Assert.Empty(html.Descendants(Xhtml.a));
        Assert.Equal("Someone", TextSpanIn(html, "Someone").Value);
    }

    // An anchor cannot hold another, whichever form each link is written in; a bookmark in a simple field's link keeps
    // its id on a span inside it.
    [Theory]
    [InlineData("<w:hyperlink r:id=\"rIdLink\">{0}</w:hyperlink>")]
    [InlineData("{1}")]
    [InlineData("<w:fldSimple w:instr=\" HYPERLINK &quot;https://example.test/outer&quot; \">{0}</w:fldSimple>")]
    public void ASimpleHyperlinkFieldInsideAnotherLinkIsNotALinkInsideALink(string outer)
    {
        var inner = Simple(" HYPERLINK \"https://example.test/inner\" ",
            "<w:bookmarkStart w:id=\"7\" w:name=\"Inside\"/>" + Run("Inner") + "<w:bookmarkEnd w:id=\"7\"/>");
        var body = $"<w:p>{string.Format(outer, Run("Outer ") + inner, Field(Run("Outer ") + inner))}</w:p>";
        foreach (var stamp in new[] { false, true })
        {
            var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", Docx(body)),
                new WmlToHtmlConverterSettings { FabricateCssClasses = false, StampReferenceFields = stamp });
            Assert.DoesNotContain(html.Descendants(Xhtml.a), a => a.Ancestors(Xhtml.a).Any());
            var link = Assert.Single(html.Descendants(Xhtml.a), a => a.Attribute("href") != null);
            Assert.Equal("Outer Inner", link.Value);
            Assert.DoesNotContain("inner", (string?)link.Attribute("href"));
            Assert.Single(link.Descendants(Xhtml.span), s => (string?)s.Attribute("id") == "Inside");
        }
    }

    // A w:hyperlink inside a simple HYPERLINK field's result gives up its own link, as it does in the complex form's.
    [Fact]
    public void AWordHyperlinkInsideASimpleHyperlinkFieldIsNotALinkInsideALink()
    {
        var html = Render($"<w:p>{Simple($" HYPERLINK \"{FieldUrl}\" ", Run("Outer ") + $"<w:hyperlink r:id=\"rIdLink\">{Run("Inner")}</w:hyperlink>")}</w:p>",
            fabricateCssClasses: false, PaginationMode.None);
        Assert.Equal("Outer Inner", Assert.Single(html.Descendants(Xhtml.a)).Value);
        Assert.Equal(FieldUrl, (string?)html.Descendants(Xhtml.a).Single().Attribute("href"));
    }

    private const string PageNumber = "<w:rPr><w:webHidden/></w:rPr>";

    // A contents entry as Word writes it: the heading's text and a leader tab in the Hyperlink style, then a PAGEREF
    // to the heading's page.
    private static string EntryContent(string text) =>
        Run(text, HyperlinkStyle) + $"<w:r>{HyperlinkStyle}<w:tab/></w:r>" + FieldWith("PAGEREF _Toc1 \\h", Run("1", PageNumber));

    private const string TocTabs = "<w:pPr><w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/></w:tabs></w:pPr>";

    // A table of contents over three paragraphs with one entry in each form (a HYPERLINK field, a w:hyperlink, and a
    // simple HYPERLINK field); then the field forms outside it, and the heading the entries lead to.
    private static readonly string Contents =
        $"<w:p>{TocTabs}{Begin("TOC \\o \"1-3\" \\h \\z \\u")}{FieldWith("HYPERLINK \\l \"_Toc1\"", EntryContent("FieldEntry"))}</w:p>"
        + $"<w:p>{TocTabs}<w:hyperlink w:anchor=\"_Toc1\" w:history=\"1\">{EntryContent("ElementEntry")}</w:hyperlink></w:p>"
        + $"<w:p>{TocTabs}{Simple(" HYPERLINK \\l \"_Toc1\" ", Run("SimpleEntry", HyperlinkStyle))}{EndOfField}</w:p>"
        + $"<w:p>{FieldWith("HYPERLINK \\l \"_Toc1\"", Run("FieldOutside", HyperlinkStyle))}</w:p>"
        + $"<w:p>{Simple(" HYPERLINK \\l \"_Toc1\" ", Run("SimpleOutside", HyperlinkStyle))}</w:p>"
        + $"<w:p><w:bookmarkStart w:id=\"1\" w:name=\"_Toc1\"/>{Run("Heading")}<w:bookmarkEnd w:id=\"1\"/></w:p>";

    [Theory]
    [InlineData(true, PaginationMode.None)]
    [InlineData(true, PaginationMode.Paginated)]
    [InlineData(false, PaginationMode.None)]
    [InlineData(false, PaginationMode.Paginated)]
    public void AContentsEntryWrittenAsAFieldIsALinkDrawnAsAWordHyperlinkEntryIs(bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render(Contents, fabricateCssClasses, mode);
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => a.Ancestors(Xhtml.a).Any());
        var element = Declarations(html, TextSpanIn(html, "ElementEntry"));
        Assert.DoesNotContain("color:", element);
        Assert.DoesNotContain("text-decoration:", element);
        foreach (var entry in new[] { "FieldEntry", "ElementEntry", "SimpleEntry" })
        {
            var link = LinkOf(html, entry);
            Assert.Equal("#_Toc1", (string?)link.Attribute("href"));
            AssertNeutral(html)(link);
            // The entry's text is drawn as the w:hyperlink entry's is: no link colour, no underline.
            Assert.Equal(element, Declarations(html, TextSpanIn(html, entry)));
            // Nor does the leader tab draw the link's underline.
            Assert.DoesNotContain(link.DescendantsAndSelf(), e => Declarations(html, e).Contains("border-top:"));
        }
        // The whole entry is one link, its leader tab too. (Word marks the page number web-hidden, so flow HTML hides it.)
        Assert.All(new[] { "FieldEntry", "ElementEntry" }, entry =>
            Assert.Single(LinkOf(html, entry).Descendants(), e => e.Attribute("data-docx-tab") != null));
        // Control: the same links outside the table of contents keep the Hyperlink style's colour and underline.
        foreach (var outside in new[] { "FieldOutside", "SimpleOutside" })
        {
            Assert.Equal("#_Toc1", (string?)LinkOf(html, outside).Attribute("href"));
            var declarations = Declarations(html, TextSpanIn(html, outside));
            Assert.Contains("color: #0563C1", declarations);
            Assert.Contains("text-decoration: underline", declarations);
        }
    }

    // Word reads a field's name in any case; so do the converter's two field forms, and the IR.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AHyperlinkFieldNamedInLowerCaseIsALink(bool simple)
    {
        var instruction = $"hyperlink \"{FieldUrl}\"";
        var body = $"<w:p>{(simple ? Simple($" {instruction} ", Run("Go")) : FieldWith(instruction, Run("Go")))}</w:p>";
        Assert.Equal("Go", Link(Render(body, fabricateCssClasses: false, PaginationMode.None), FieldUrl).Value);
        var paragraph = Assert.Single(IrReader.Read(new WmlDocument("field.docx", Docx(body))).Body.Blocks.OfType<IrParagraph>());
        Assert.Equal(FieldUrl, Assert.Single(paragraph.Inlines.OfType<IrHyperlink>()).Target);
    }

    // A HYPERLINK field in another's result is converted in every mode now, and still gives up its own link there.
    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void AHyperlinkFieldInAnothersResultIsNotALinkInsideALinkInAnyMode(PaginationMode mode)
    {
        var html = Render($"<w:p>{Field(Run("Outer ") + FieldWith("HYPERLINK \"https://example.test/inner\"", Run("Inner")))}</w:p>",
            fabricateCssClasses: false, mode);
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => a.Ancestors(Xhtml.a).Any());
        Assert.Equal("Outer Inner", Link(html, FieldUrl).Value);
    }

    // Fields nested in an IF field's instruction only feed its condition: Word shows the IF's result. They make no link
    // and no page-number marker, whatever becomes of their cached text (control: the same fields in the IF's result).
    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void AFieldInAnotherFieldsInstructionIsNoLinkAndNoPageMarker(PaginationMode mode)
    {
        var inInstruction = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> IF </w:instrText></w:r>"
            + FieldWith("HYPERLINK \"https://example.test/condition\"", Run("Left"))
            + "<w:r><w:instrText xml:space=\"preserve\"> = </w:instrText></w:r>" + FieldWith("PAGE", Run("7"))
            + "<w:r><w:instrText xml:space=\"preserve\"> \"Shown\" \"Hidden\" </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Run("Shown") + EndOfField;
        var inResult = FieldWith("QUOTE \"x\"", FieldWith($"HYPERLINK \"{FieldUrl}\"", Run("Linked")) + Run(" page ") + FieldWith("PAGE", Run("3")));
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", Docx($"<w:p>{inInstruction}</w:p><w:p>{inResult}</w:p>")),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = mode });
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => ((string?)a.Attribute("href"))?.Contains("condition") == true);
        Assert.Equal("Linked", Link(html, FieldUrl).Value);
        var pageMarkers = html.Descendants().Where(e => (string?)e.Attribute("data-field") == "PAGE").ToList();
        // In paginated HTML a PAGE field in another's result is marked for the paginator, as at the top level.
        Assert.Equal(mode == PaginationMode.Paginated ? new[] { "3" } : Array.Empty<string>(), pageMarkers.Select(e => e.Value));
    }

    private const string Inserted = "w:id=\"{0}\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\"";

    private static string Ins(int id, string content) => $"<w:ins {string.Format(Inserted, id)}>{content}</w:ins>";

    // An old contents entry, deleted: its field code is w:delInstrText and its text w:delText.
    private static string DeletedEntry(int id, string text) => $"""
        <w:del {string.Format(Inserted, id)}>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:delInstrText xml:space="preserve"> HYPERLINK \l "_Toc1" </w:delInstrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r>{HyperlinkStyle}<w:delText>{text}</w:delText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:del>
        """;

    // A table of contents updated with changes tracked, as Word writes it: the first paragraph's insertion holds the TOC
    // field's start and the first entry, each later entry sits in its own insertion beside the deleted old one, and the
    // last insertion holds the last entry and the field's end.
    [Theory]
    [InlineData(true, PaginationMode.None)]
    [InlineData(true, PaginationMode.Paginated)]
    [InlineData(false, PaginationMode.None)]
    [InlineData(false, PaginationMode.Paginated)]
    public void AContentsEntryWrittenAsAFieldInsideTrackedChangesIsALink(bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render(
            $"<w:p>{TocTabs}{Ins(11, Begin("TOC \\o \"1-3\" \\h \\z \\u") + FieldWith("HYPERLINK \\l \"_Toc1\"", EntryContent("FirstEntry")))}</w:p>"
            + $"<w:p>{TocTabs}{DeletedEntry(12, "OldEntry")}{Ins(13, FieldWith("HYPERLINK \\l \"_Toc1\"", EntryContent("MiddleEntry")))}</w:p>"
            + $"<w:p>{TocTabs}{Ins(14, FieldWith("HYPERLINK \\l \"_Toc1\"", EntryContent("LastEntry")) + EndOfField)}</w:p>"
            + $"<w:p><w:bookmarkStart w:id=\"1\" w:name=\"_Toc1\"/>{Run("Heading")}<w:bookmarkEnd w:id=\"1\"/></w:p>",
            fabricateCssClasses, mode);
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => a.Ancestors(Xhtml.a).Any());
        foreach (var entry in new[] { "FirstEntry", "MiddleEntry", "LastEntry" })
        {
            var link = LinkOf(html, entry);
            Assert.Equal("#_Toc1", (string?)link.Attribute("href"));
            Assert.NotEmpty(link.Ancestors(Xhtml.ins));
            var declarations = Declarations(html, TextSpanIn(html, entry));
            Assert.DoesNotContain("#0563C1", declarations);
        }
        // The deleted old entry stays a link too, as a deleted w:hyperlink entry does.
        Assert.NotEmpty(TextSpanIn(html, "OldEntry").Ancestors(Xhtml.del));
        Assert.Equal("#_Toc1", (string?)LinkOf(html, "OldEntry").Attribute("href"));
    }

    // An insertion that starts a field and ends in its result: a field nested in the started field's instruction still
    // only feeds it, and the result's own field converts.
    [Fact]
    public void AnInsertionStartingAFieldLeavesFieldsInItsInstructionInert()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", Docx(
                "<w:p>" + Ins(21, "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> IF </w:instrText></w:r>"
                    + FieldWith("HYPERLINK \"https://example.test/condition\"", Run("Left"))
                    + "<w:r><w:instrText xml:space=\"preserve\"> = \"Left\" \"Shown\" \"Hidden\" </w:instrText></w:r>"
                    + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + FieldWith($"HYPERLINK \"{FieldUrl}\"", Run("Shown")))
                + EndOfField + "</w:p>")),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = true });
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => ((string?)a.Attribute("href"))?.Contains("condition") == true);
        Assert.Equal("Shown", Link(html, FieldUrl).Value);
    }

    // A revision mark that starts in a nested field's code, with two fields open at its start: it groups only the whole
    // fields above both, so the PAGE field around it gets one page-number marker, not a second inside the mark (#1272).
    [Fact]
    public void ARevisionMarkInsideANestedFieldsCodeMarksThatFieldOnce()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", Docx(
                $"<w:p>{Begin("QUOTE \"x\"")}<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
                + "<w:r><w:instrText xml:space=\"preserve\"> PA</w:instrText></w:r>"
                + Ins(31, "<w:r><w:instrText xml:space=\"preserve\">GE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
                    + Run("8") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>")
                + EndOfField + "</w:p>")),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = true, RenderPagination = PaginationMode.Paginated });
        Assert.Equal("8", Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-field") == "PAGE").Value);
    }
}
