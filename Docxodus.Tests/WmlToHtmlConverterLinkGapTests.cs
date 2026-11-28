using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Three lesser gaps around links (issue #1878): a contents entry keeps the colour and underline set on its
/// own runs; a complex field inside a simple field converts as a field; and a note reference or comment marker
/// inside a link leaves no anchor inside another.</summary>
public class WmlToHtmlConverterLinkGapTests
{
    private const string Url = "https://example.test/link";

    private static string Run(string text, string rPr = "") => $"<w:r>{rPr}<w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Props(string properties) => $"<w:rPr>{properties}</w:rPr>";

    private const string HyperlinkStyle = "<w:rStyle w:val=\"Hyperlink\"/>";

    private static string Complex(string instruction, string result) => $"""
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>
        """;

    private static string Simple(string instruction, string result) =>
        $"<w:fldSimple w:instr=\"{instruction.Replace("\"", "&quot;")}\">{result}</w:fldSimple>";

    private const string BeginContents = """
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        """;

    private const string EndOfField = "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private const string TocTabs = "<w:pPr><w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/></w:tabs></w:pPr>";

    private const string Heading = "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"_Toc1\"/><w:r><w:t>Heading</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";

    /// <summary>The three forms a link takes in a document, around <paramref name="content"/>: a w:hyperlink, a complex
    /// HYPERLINK field and a simple one. <paramref name="href"/> is an address, or a <c>#</c> and a bookmark.</summary>
    private static string LinkIn(string form, string href, string content)
    {
        var bookmark = href.StartsWith('#');
        var instruction = bookmark ? $"HYPERLINK \\l \"{href[1..]}\"" : $"HYPERLINK \"{href}\"";
        return form switch
        {
            "element" => bookmark
                ? $"<w:hyperlink w:anchor=\"{href[1..]}\" w:history=\"1\">{content}</w:hyperlink>"
                : $"<w:hyperlink r:id=\"rIdLink\">{content}</w:hyperlink>",
            "complex" => Complex(instruction, content),
            "simple" => Simple($" {instruction} ", content),
            _ => throw new ArgumentOutOfRangeException(nameof(form)),
        };
    }

    private static void Put(OpenXmlPart part, string xml) => part.PutXDocument(XDocument.Parse(xml));

    private static byte[] Docx(string body)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var ns = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\"";
            main.AddHyperlinkRelationship(new Uri(Url), true, "rIdLink");
            Put(main, $"""
                <w:document {ns}><w:body>{body}
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
                    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """);
            Put(main.AddNewPart<StyleDefinitionsPart>(), $"""
                <w:styles {ns}>
                  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>
                    <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
                  <w:style w:type="character" w:styleId="Accent"><w:name w:val="Accent"/>
                    <w:rPr><w:color w:val="00B050"/><w:u w:val="wave"/></w:rPr></w:style>
                </w:styles>
                """);
            Put(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>");
            const string separators = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";
            Put(main.AddNewPart<FootnotesPart>(), $"<w:footnotes {ns}>{separators}<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r>{Run(" A note.")}</w:p></w:footnote></w:footnotes>");
            Put(main.AddNewPart<EndnotesPart>(), $"<w:endnotes {ns}>{separators.Replace("footnote", "endnote")}<w:endnote w:id=\"1\"><w:p><w:r><w:endnoteRef/></w:r>{Run(" An endnote.")}</w:p></w:endnote></w:endnotes>");
            Put(main.AddNewPart<WordprocessingCommentsPart>(), $"""
                <w:comments {ns}><w:comment w:id="0" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="R">
                  <w:p><w:r><w:t>A comment.</w:t></w:r></w:p></w:comment></w:comments>
                """);
        }
        return stream.ToArray();
    }

    private static XElement Render(string body, WmlToHtmlConverterSettings settings) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("links.docx", Docx(body)), settings);

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

    private static XElement TextSpan(XElement html, string text) =>
        Assert.Single(html.Descendants(Xhtml.span), s => !s.HasElements && s.Value == text);

    private static XElement? NearestLink(XElement element) => element.Ancestors(Xhtml.a).FirstOrDefault();

    private static bool HasClass(XElement element, string name) =>
        ((string?)element.Attribute("class") ?? "").Split(' ').Contains(name);

    public static TheoryData<string, bool, PaginationMode> FormsAndModes()
    {
        var data = new TheoryData<string, bool, PaginationMode>();
        foreach (var form in new[] { "element", "complex", "simple" })
            foreach (var fabricate in new[] { true, false })
                foreach (var mode in new[] { PaginationMode.None, PaginationMode.Paginated })
                    data.Add(form, fabricate, mode);
        return data;
    }

    // Item 1. Word drops the Hyperlink style's colour and underline from a contents entry (#427), not what the entry's
    // runs set themselves: in Word reference PDFs a Hyperlink-styled run with a direct red colour prints red and not
    // underlined, a run with a direct double underline keeps it, and a Hyperlink-styled run with nothing of its own
    // prints with neither. The page-number tab of a directly underlined run draws that underline too.

    private const string DirectRed = "<w:rStyle w:val=\"Hyperlink\"/><w:color w:val=\"FF0000\"/>";
    private const string DirectDouble = "<w:u w:val=\"double\"/>";

    // (Runs drawn alike are drawn as one span, so the runs drawn plain are kept apart.)
    private static string DirectEntry(string name) =>
        Run(name + "Styled", Props(HyperlinkStyle)) + Run(name + "Red", Props(DirectRed))
        + Run(name + "Accent", Props("<w:rStyle w:val=\"Accent\"/>")) + Run(name + "Double", Props(DirectDouble))
        + $"<w:r>{Props(DirectDouble)}<w:tab/></w:r>" + Run("1");

    private static string DirectContents(string form) =>
        $"<w:p>{TocTabs}{BeginContents}{LinkIn(form, "#_Toc1", DirectEntry("Entry"))}{EndOfField}</w:p>"
        + $"<w:p>{LinkIn(form, "#_Toc1", DirectEntry("Outside"))}</w:p>" + Heading;

    [Theory]
    [MemberData(nameof(FormsAndModes))]
    public void AContentsEntryKeepsTheColourAndUnderlineSetOnItsOwnRuns(string form, bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render(DirectContents(form), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = fabricateCssClasses,
            RenderPagination = mode,
        });
        var red = Declarations(html, TextSpan(html, "EntryRed"));
        Assert.Contains("color: #FF0000", red);
        Assert.DoesNotContain("text-decoration", red);
        Assert.Contains("text-decoration: underline double", Declarations(html, TextSpan(html, "EntryDouble")));
        var styled = Declarations(html, TextSpan(html, "EntryStyled"));
        Assert.DoesNotContain("color:", styled);
        Assert.DoesNotContain("text-decoration", styled);
        // Another character style's colour and underline still go, as before (Word's behaviour there is unmeasured).
        var accent = Declarations(html, TextSpan(html, "EntryAccent"));
        Assert.DoesNotContain("color:", accent);
        Assert.DoesNotContain("text-decoration", accent);
        // The entry is one link, and the tab of its directly double-underlined run draws the run's line.
        var entryLink = NearestLink(TextSpan(html, "EntryRed"));
        Assert.Equal("#_Toc1", (string?)entryLink?.Attribute("href"));
        var tab = Assert.Single(entryLink!.Descendants(), e => e.Attribute("data-docx-tab") != null);
        Assert.Contains(tab.Descendants(), e => Declarations(html, e).Contains("border-top: 3px double"));

        // Control: outside the contents every run keeps what its styles give it.
        Assert.Contains("text-decoration: underline", Declarations(html, TextSpan(html, "OutsideRed")));
        Assert.Contains("color: #0563C1", Declarations(html, TextSpan(html, "OutsideStyled")));
        Assert.Contains("color: #00B050", Declarations(html, TextSpan(html, "OutsideAccent")));
    }

    // Item 2. A simple field's content converts with the field grouping a paragraph's does, so a complex field in it is
    // a field: a HYPERLINK there is a link, and with StampReferenceFields a PAGEREF there is marked once.

    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void AComplexHyperlinkFieldInsideASimpleFieldIsALink(PaginationMode mode)
    {
        var html = Render($"<w:p>{Simple(" QUOTE \"x\" ", Run("Quoted ") + Complex($"HYPERLINK \"{Url}\"", Run("Linked")))}</w:p>",
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = mode });
        var link = Assert.Single(html.Descendants(Xhtml.a));
        Assert.Equal(Url, (string?)link.Attribute("href"));
        Assert.Equal("Linked", link.Value);
        Assert.Null(NearestLink(TextSpan(html, "Quoted ")));
    }

    // (An entry with a tab was already marked, by the tab layout a link's content gets; the control.)
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AComplexPageReferenceInASimpleContentsEntryIsMarkedOnce(bool tabbed)
    {
        var entry = Run("Entry", Props(HyperlinkStyle)) + (tabbed ? $"<w:r>{Props(HyperlinkStyle)}<w:tab/></w:r>" : "")
            + Complex("PAGEREF _Toc1 \\h", Run("3"));
        var html = Render($"<w:p>{TocTabs}{BeginContents}{Simple(" HYPERLINK \\l \"_Toc1\" ", entry)}{EndOfField}</w:p>{Heading}",
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, StampReferenceFields = true });
        var marker = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-field") == "PAGEREF");
        Assert.Equal("3", marker.Value);
        Assert.Equal("#_Toc1", (string?)NearestLink(marker)?.Attribute("href"));
    }

    // The fields open around a simple field are converted by the content around it, not again inside it.
    [Fact]
    public void ASimpleFieldInsideAReferencesResultLeavesTheReferenceMarkedOnce()
    {
        var html = Render($"<w:p>{Complex("REF _Toc1 \\h", Simple(" QUOTE \"x\" ", Complex("PAGEREF _Toc1 \\h", Run("3"))))}</w:p>{Heading}",
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, StampReferenceFields = true });
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-field") == "REF");
        Assert.Equal("3", Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-field") == "PAGEREF").Value);
    }

    // A simple field in another field's instruction only feeds it: a HYPERLINK field inside it is no link.
    [Fact]
    public void AComplexHyperlinkFieldInsideASimpleFieldInAnInstructionIsNoLink()
    {
        var html = Render("<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> IF </w:instrText></w:r>"
            + Simple(" QUOTE \"x\" ", Complex($"HYPERLINK \"{Url}\"", Run("Left")))
            + "<w:r><w:instrText xml:space=\"preserve\"> = \"Left\" \"Shown\" \"Hidden\" </w:instrText></w:r>"
            + $"<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>{Run("Shown")}{EndOfField}</w:p>",
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        Assert.Empty(html.Descendants(Xhtml.a));
    }

    // Item 3. The converter's own anchor for a note reference or a comment marker inside a link: an HTML parser closes
    // the link at it, so the link's text after it would be no link. The link is split around it instead.

    private const string FootnoteReference =
        "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r>";

    private static void AssertNoAnchorInsideAnother(XElement html, string href, params string[] texts)
    {
        Assert.DoesNotContain(html.Descendants(Xhtml.a), a => a.Ancestors(Xhtml.a).Any());
        foreach (var text in texts)
            Assert.Equal(href, (string?)NearestLink(TextSpan(html, text))?.Attribute("href"));
        var ids = html.Descendants().Select(e => (string?)e.Attribute("id")).OfType<string>().ToList();
        Assert.Equal(ids.Distinct().Count(), ids.Count);
        // Every piece of the link leaves colour and underline to its runs, as the link did.
        Assert.All(html.Descendants(Xhtml.a).Where(a => (string?)a.Attribute("href") == href), piece =>
        {
            Assert.Contains("color: inherit", Declarations(html, piece));
            Assert.Contains("text-decoration: none", Declarations(html, piece));
        });
    }

    // A footnote reference in the link's own runs, and an endnote reference inside a tracked insertion in the link, with
    // inserted text on either side of it.
    [Theory]
    [MemberData(nameof(FormsAndModes))]
    public void ANoteReferenceInsideALinkLeavesTheTextAfterItInTheLink(string form, bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render($"<w:p>{LinkIn(form, Url, Run("Before ") + FootnoteReference + Run(" Middle ")
                + "<w:ins w:id=\"5\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\">" + Run("Inserted ")
                + "<w:r><w:endnoteReference w:id=\"1\"/></w:r>" + Run(" Kept") + "</w:ins>" + Run(" After"))}</w:p>",
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = fabricateCssClasses,
                RenderPagination = mode,
                RenderFootnotesAndEndnotes = true,
                RenderTrackedChanges = true,
            });
        AssertNoAnchorInsideAnother(html, Url, "Before ", " Middle ", "Inserted ", " Kept", " After");
        Assert.Equal("Inserted i Kept", Assert.Single(html.Descendants(Xhtml.ins)).Value);
        // The note's own anchors are as they were.
        var footnote = Assert.Single(html.Descendants(Xhtml.a), a => HasClass(a, "footnote-ref"));
        Assert.Equal("fn-ref-1", (string?)footnote.Attribute("id"));
        Assert.Equal("1", (string?)footnote.Attribute("data-footnote-id"));
        Assert.Equal(mode == PaginationMode.Paginated ? null : "#fn-1", (string?)footnote.Attribute("href"));
        Assert.Equal("1", footnote.Value);
        var endnote = Assert.Single(html.Descendants(Xhtml.a), a => HasClass(a, "endnote-ref"));
        Assert.Equal("en-ref-1", (string?)endnote.Attribute("id"));
        // The text reads in its order.
        var paragraph = TextSpan(html, "Before ").Ancestors(Xhtml.p).First();
        Assert.Equal("Before 1 Middle Inserted i Kept After", paragraph.Value);
    }

    // A contents entry with a footnote reference in its label: the link is split inside the box that lays the entry out at
    // its leader tab, and the box that the leader fills stays an item of that box, so it still reaches the tab stop.
    [Theory]
    [MemberData(nameof(FormsAndModes))]
    public void ANoteReferenceInAContentsEntryLeavesTheLeaderTabLaidOut(string form, bool fabricateCssClasses, PaginationMode mode)
    {
        var entry = Run("Label", Props(HyperlinkStyle)) + FootnoteReference + Run(" More", Props(HyperlinkStyle))
            + $"<w:r>{Props(HyperlinkStyle)}<w:tab/></w:r>" + Run("7");
        var html = Render($"<w:p>{TocTabs}{BeginContents}{LinkIn(form, "#_Toc1", entry)}{EndOfField}</w:p>{Heading}",
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = fabricateCssClasses,
                RenderPagination = mode,
                RenderFootnotesAndEndnotes = true,
            });
        AssertNoAnchorInsideAnother(html, "#_Toc1", "Label", " More", "7");
        var tab = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-tab") != null);
        Assert.Equal("#_Toc1", (string?)NearestLink(tab)?.Attribute("href"));
        var leaderBox = tab.Ancestors().First(e => Declarations(html, e).Contains("flex-grow: 1"));
        Assert.Contains("display: inline-flex", Declarations(html, leaderBox.Parent!));
    }

    [Theory]
    [MemberData(nameof(FormsAndModes))]
    public void ACommentMarkerInsideALinkLeavesTheTextAfterItInTheLink(string form, bool fabricateCssClasses, PaginationMode mode)
    {
        var html = Render($"<w:p>{LinkIn(form, Url, Run("Before ") + "<w:commentRangeStart w:id=\"0\"/>" + Run("Commented")
                + "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r>" + Run(" After"))}</w:p>",
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = fabricateCssClasses,
                RenderPagination = mode,
                RenderComments = true,
            });
        AssertNoAnchorInsideAnother(html, Url, "Before ", "Commented", " After");
        var marker = Assert.Single(html.Descendants(Xhtml.a), a => (string?)a.Attribute("id") == "comment-ref-0");
        Assert.True(HasClass(marker, "comment-marker"));
        Assert.Equal("#comment-0", (string?)marker.Attribute("href"));
        Assert.Equal("[0]", marker.Value);
        var paragraph = TextSpan(html, "Before ").Ancestors(Xhtml.p).First();
        Assert.Equal("Before Commented[0] After", paragraph.Value);
    }
}
