using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Lesser gaps around links (issue #1878): a contents entry keeps the colour and underline set on its own
/// runs, and a complex field inside a simple field converts as a field.</summary>
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
    // prints with neither.

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
        // The entry is one link.
        Assert.Equal("#_Toc1", (string?)NearestLink(TextSpan(html, "EntryRed"))?.Attribute("href"));

        // Control: outside the contents every run keeps what its styles give it.
        Assert.Contains("text-decoration: underline", Declarations(html, TextSpan(html, "OutsideRed")));
        Assert.Contains("color: #0563C1", Declarations(html, TextSpan(html, "OutsideStyled")));
        Assert.Contains("color: #00B050", Declarations(html, TextSpan(html, "OutsideAccent")));
    }
}
