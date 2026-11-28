using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Word shows a field's result only: whatever sits between its begin and separate characters is its code, a
/// field nested there included, and a nested field's cached result feeds the code rather than showing (issue #1875).
/// So <c>{ IF { PAGE } &gt; 1 "Later" "First" }</c> reads "First", and <c>{ HYPERLINK "{ MERGEFIELD Url }" }</c> links
/// to the merged address and shows only its own result.</summary>
public class WmlToHtmlConverterFieldInstructionTests
{
    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Code(string code) => $"<w:r><w:instrText xml:space=\"preserve\">{code}</w:instrText></w:r>";

    private const string Begin = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>";
    private const string Separate = "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>";
    private const string End = "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    /// <summary>A complex field whose code is <paramref name="instruction"/>: instruction runs and nested fields.</summary>
    private static string Field(string instruction, string result) => Begin + instruction + Separate + result + End;

    private static string Ins(int id, string content) =>
        $"<w:ins w:id=\"{id}\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\">{content}</w:ins>";

    private static readonly string NestedPage = Field(Code(" PAGE "), Run("1"));

    private static readonly string IfCode = Code(" IF ") + NestedPage + Code(" &gt; 1 \"Later\" \"First\" ");

    private static string MergeCode(string merged) =>
        Code(" HYPERLINK \"") + Field(Code(" MERGEFIELD Url "), merged.Length == 0 ? "" : Run(merged)) + Code("\" ");

    private static string Labelled(string story, string label, string content) =>
        $"<w:p>{Run($"{story} {label}: ")}{content}</w:p>";

    // The issue's shapes, a bookmark in a field's code, and the controls: the same nested fields in a result, and a
    // plain link.
    private static string Paragraphs(string story) =>
        Labelled(story, "If", Field(IfCode, Run("First")))
        + Labelled(story, "Merge", Field(MergeCode("https://example.test/merged"), Run("Visit")))
        + Labelled(story, "Quote", Field(Code(" HYPERLINK \"") + Field(Code(" QUOTE \"https://example.test/quoted\" "),
            Run("https://example.test/quoted")) + Code("\" "), Run("Go")))
        + Labelled(story, "Mark", Field(Code(" IF ") + $"<w:bookmarkStart w:id=\"7\" w:name=\"{story}Mark\"/>" + NestedPage
            + "<w:bookmarkEnd w:id=\"7\"/>" + Code(" &gt; 1 \"Later\" \"Marked\" "), Run("Marked")))
        + Labelled(story, "Ref", Field(Code(" IF ") + Field(Code(" REF Target \\h "), Run("x"))
            + Code(" = \"x\" \"Matched\" \"Missed\" "), Run("Matched")))
        + Labelled(story, "Cite", Field(Code(" REF Mark") + NestedPage + Code(" \\h "), Run("Cited")))
        + Labelled(story, "Result", Field(Code(" QUOTE \"x\" "), Field(Code(" PAGE "), Run("3")) + Run(" ")
            + Field(Code(" HYPERLINK \"https://example.test/result\" "), Run("Linked"))))
        + Labelled(story, "Plain", Field(Code(" HYPERLINK \"https://example.test/plain\" "), Run("Plain")));

    private const string Notes = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static void Put(OpenXmlPart part, string xml) => part.PutXDocument(XDocument.Parse(xml));

    /// <summary>A document whose body, footnote and header each hold <paramref name="paragraphs"/> for their story.</summary>
    private static byte[] Docx(Func<string, string> paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var ns = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\"";
            Put(main.AddNewPart<HeaderPart>("rIdHeader"), $"<w:hdr {ns}>{paragraphs("Header")}</w:hdr>");
            Put(main.AddNewPart<FootnotesPart>(), $"<w:footnotes {ns}>{Notes}<w:footnote w:id=\"1\">"
                + $"<w:p><w:r><w:footnoteRef/></w:r>{Run("Note")}</w:p>{paragraphs("Footnote")}<w:p>{Run("End")}</w:p></w:footnote></w:footnotes>");
            Put(main, $"""
                <w:document {ns}><w:body>
                  {paragraphs("Body")}<w:p>{Run("Reference")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>
                  <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>
                    <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
                </w:body></w:document>
                """);
            Put(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}/>");
            Put(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>");
        }
        return stream.ToArray();
    }

    private static XElement Render(Func<string, string> paragraphs, PaginationMode mode, bool stampReferences, bool stampPages,
        bool trackedChanges = false) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("fields.docx", Docx(paragraphs)), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            RenderPagination = mode,
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = true,
            StampReferenceFields = stampReferences,
            StampPageNumberFields = stampPages,
            RenderTrackedChanges = trackedChanges,
        });

    /// <summary>What each paragraph labelled <paramref name="label"/> in <paramref name="story"/> shows after its
    /// label; the header and note can render in more than one place.</summary>
    private static List<string> Shown(XElement html, string story, string label)
    {
        var prefix = $"{story} {label}: ";
        var shown = html.Descendants(Xhtml.p).Where(p => p.Value.StartsWith(prefix, StringComparison.Ordinal))
            .Select(p => p.Value.Substring(prefix.Length)).ToList();
        Assert.NotEmpty(shown);
        return shown;
    }

    /// <summary>The target of every link that shows <paramref name="text"/>, which must all agree.</summary>
    private static string? Href(XElement html, string text)
    {
        var targets = html.Descendants(Xhtml.a).Where(a => a.Value == text).Select(a => (string?)a.Attribute("href")).ToList();
        Assert.NotEmpty(targets);
        return Assert.Single(targets.Distinct());
    }

    private static readonly string[] Stories = { "Body", "Footnote", "Header" };

    private static IEnumerable<XElement> PageMarkers(XElement html) =>
        html.Descendants().Where(e => (string?)e.Attribute("data-field") == "PAGE");

    public static TheoryData<PaginationMode, bool, bool> Modes()
    {
        var data = new TheoryData<PaginationMode, bool, bool>();
        foreach (var mode in new[] { PaginationMode.None, PaginationMode.Paginated })
            foreach (var stampReferences in new[] { false, true })
                foreach (var stampPages in new[] { false, true })
                    data.Add(mode, stampReferences, stampPages);
        return data;
    }

    [Theory]
    [MemberData(nameof(Modes))]
    public void AFieldInAnotherFieldsCodeShowsNothingOfItsOwn(PaginationMode mode, bool stampReferences, bool stampPages)
    {
        var html = Render(Paragraphs, mode, stampReferences, stampPages);
        foreach (var story in Stories)
        {
            Assert.All(Shown(html, story, "If"), text => Assert.Equal("First", text));
            Assert.All(Shown(html, story, "Merge"), text => Assert.Equal("Visit", text));
            Assert.All(Shown(html, story, "Quote"), text => Assert.Equal("Go", text));
            Assert.All(Shown(html, story, "Mark"), text => Assert.Equal("Marked", text));
            // A bookmark in a field's code is still a target.
            Assert.Contains(html.Descendants(), e => (string?)e.Attribute("id") == $"{story}Mark");
            Assert.All(Shown(html, story, "Ref"), text => Assert.Equal("Matched", text));
            Assert.All(Shown(html, story, "Cite"), text => Assert.Equal("Cited", text));
            Assert.All(Shown(html, story, "Result"), text => Assert.Equal("3 Linked", text));
            Assert.All(Shown(html, story, "Plain"), text => Assert.Equal("Plain", text));
        }
        // A reference in a field's code is not marked for an editor either; a reference whose code holds a field is.
        Assert.Equal(stampReferences ? new[] { "Cited", "Cited", "Cited" } : Array.Empty<string>(),
            html.Descendants().Where(e => (string?)e.Attribute("data-field") is { } kind && kind != "PAGE").Select(e => e.Value));
        // A link goes to the address the field in its code gives; a field in another's code is no link.
        Assert.Equal("https://example.test/merged", Href(html, "Visit"));
        Assert.Equal("https://example.test/quoted", Href(html, "Go"));
        Assert.Equal("https://example.test/result", Href(html, "Linked"));
        Assert.Equal("https://example.test/plain", Href(html, "Plain"));
        Assert.All(html.Descendants(Xhtml.a).Select(a => (string?)a.Attribute("href")).Where(href => href?.StartsWith("#fn") == false),
            href => Assert.Contains(href, new[] { "https://example.test/merged", "https://example.test/quoted",
                "https://example.test/result", "https://example.test/plain" }));
        // Only a PAGE field in a result is marked for the paginator: none in a field's code, whose marker it would fill.
        var marked = mode == PaginationMode.Paginated || stampPages;
        Assert.Equal(marked ? new[] { "3", "3", "3" } : Array.Empty<string>(), PageMarkers(html).Select(e => e.Value));
    }

    // A simple field in another's code, the only field nested in a code in its document.
    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void ASimpleFieldInAnotherFieldsCodeShowsNothingOfItsOwn(PaginationMode mode)
    {
        var html = Render(story => Labelled(story, "If", Field(Code(" IF ") + $"<w:fldSimple w:instr=\" PAGE \">{Run("1")}</w:fldSimple>"
            + Code(" &gt; 1 \"Later\" \"First\" "), Run("First"))), mode, stampReferences: false, stampPages: false);
        foreach (var story in Stories)
            Assert.All(Shown(html, story, "If"), text => Assert.Equal("First", text));
    }

    // The field in another's code sits in an insertion, or an insertion starts the outer field.
    private static string Revised(string story) =>
        Labelled(story, "If", Field(Code(" IF ") + Ins(41, NestedPage) + Code(" &gt; 1 \"Later\" \"First\" "), Run("First")))
        + Labelled(story, "Inserted", Ins(42, Begin + IfCode + Separate + Run("Started")) + End)
        + Labelled(story, "Merge", Field(Code(" HYPERLINK \"") + Ins(43, Field(Code(" MERGEFIELD Url "), Run("https://example.test/merged")))
            + Code("\" "), Run("Visit")));

    [Theory]
    [MemberData(nameof(Modes))]
    public void AFieldInAnInsertionInAnotherFieldsCodeShowsNothingOfItsOwn(PaginationMode mode, bool stampReferences, bool stampPages)
    {
        var html = Render(Revised, mode, stampReferences, stampPages, trackedChanges: true);
        foreach (var story in Stories)
        {
            Assert.All(Shown(html, story, "If"), text => Assert.Equal("First", text));
            Assert.All(Shown(html, story, "Inserted"), text => Assert.Equal("Started", text));
            Assert.All(Shown(html, story, "Merge"), text => Assert.Equal("Visit", text));
        }
        Assert.Equal("https://example.test/merged", Href(html, "Visit"));
        Assert.Empty(PageMarkers(html));
    }

    // A link's address from an IF, itself fed by a PAGE field: only the IF's result is the address. A nested field with
    // no result leaves an empty address, so no target. A deleted link reads its deleted nested result.
    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void ALinkGoesToTheResultOfTheFieldInItsCode(PaginationMode mode)
    {
        var html = Render(story => story != "Body" ? "" :
            Labelled(story, "Two", Field(Code(" HYPERLINK \"") + Field(Code(" IF ") + NestedPage
                    + Code(" &gt; 1 \"https://example.test/later\" \"https://example.test/first\" "), Run("https://example.test/first"))
                + Code("\" "), Run("Twice")))
            + Labelled(story, "Empty", Field(MergeCode(""), Run("Nowhere")))
            + Labelled(story, "Deleted", $"<w:del w:id=\"51\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\">{Field(
                MergeCode("https://example.test/deleted"), Run("Gone"))}</w:del>"
                .Replace("w:instrText", "w:delInstrText").Replace("<w:t ", "<w:delText ").Replace("</w:t>", "</w:delText>")),
            mode, stampReferences: false, stampPages: false, trackedChanges: true);
        Assert.All(Shown(html, "Body", "Two"), text => Assert.Equal("Twice", text));
        Assert.Equal("https://example.test/first", Href(html, "Twice"));
        Assert.All(Shown(html, "Body", "Empty"), text => Assert.Equal("Nowhere", text));
        Assert.Null(Href(html, "Nowhere"));
        Assert.All(Shown(html, "Body", "Deleted"), text => Assert.Equal("Gone", text));
        Assert.Equal("https://example.test/deleted", Href(html, "Gone"));
    }
}
