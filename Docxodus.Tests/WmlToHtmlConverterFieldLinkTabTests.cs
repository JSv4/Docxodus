using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Word writes a table-of-contents entry as a <c>w:hyperlink</c> or as a HYPERLINK field, the entry's text, leader
/// tab and page number inside the link either way. Both forms lay the tab out alike, so the page number ends on the
/// right tab stop in either (issue #1877).</summary>
public class WmlToHtmlConverterFieldLinkTabTests
{
    private const string Hyperlink = "<w:rPr><w:rStyle w:val=\"Hyperlink\"/><w:noProof/></w:rPr>";
    private const string WebHidden = "<w:rPr><w:noProof/><w:webHidden/></w:rPr>";

    private static string Run(string rPr, string content) => $"<w:r>{rPr}{content}</w:r>";

    private static string Text(string text) => $"<w:t xml:space=\"preserve\">{text}</w:t>";

    private static string Field(string instruction, string result, string rPr = "") =>
        Run(rPr, "<w:fldChar w:fldCharType=\"begin\"/>")
        + Run(rPr, $"<w:instrText xml:space=\"preserve\"> {instruction} </w:instrText>")
        + Run(rPr, "<w:fldChar w:fldCharType=\"separate\"/>") + result + Run(rPr, "<w:fldChar w:fldCharType=\"end\"/>");

    /// <summary>An entry as Word writes it: the heading's text and a leader tab in the Hyperlink style, then a PAGEREF
    /// to the heading's page, which Word marks web-hidden.</summary>
    private static string Entry(string pageNumber) => Run(Hyperlink, Text("Same entry")) + Run(Hyperlink, "<w:tab/>") + pageNumber;

    private static readonly string PageReference = Field("PAGEREF _Toc1 \\h", Run(WebHidden, Text("1")), WebHidden);

    // The entry in each of the three forms of a link to the heading.
    private static string AsElement(string entry) => $"<w:hyperlink w:anchor=\"_Toc1\" w:history=\"1\">{entry}</w:hyperlink>";

    private static string AsField(string entry) => Field("HYPERLINK \\l \"_Toc1\"", entry);

    private static string AsSimpleField(string entry) =>
        $"<w:fldSimple w:instr=\" HYPERLINK \\l &quot;_Toc1&quot; \">{entry}</w:fldSimple>";

    private static string Paragraph(string content, string style = "TOC1") =>
        $"<w:p><w:pPr><w:pStyle w:val=\"{style}\"/><w:rPr><w:noProof/></w:rPr></w:pPr>{content}</w:p>";

    private static readonly string TocBegin = Run("", "<w:fldChar w:fldCharType=\"begin\"/>")
        + Run("", "<w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText>")
        + Run("", "<w:fldChar w:fldCharType=\"separate\"/>");

    private static readonly string TocEnd = Run("", "<w:fldChar w:fldCharType=\"end\"/>");

    private const string Heading =
        "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:bookmarkStart w:id=\"1\" w:name=\"_Toc1\"/>"
        + "<w:r><w:t>Same entry</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>";

    /// <summary>A table of contents holding the same entry three times: as a w:hyperlink, a HYPERLINK field and a simple
    /// HYPERLINK field.</summary>
    private static string Contents(string entry) =>
        Paragraph(TocBegin + AsElement(entry)) + Paragraph(AsField(entry)) + Paragraph(AsSimpleField(entry) + TocEnd) + Heading;

    private static byte[] Docx(string body)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}"><w:body>{body}
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
                    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"""
                <w:styles xmlns:w="{W.w}">
                  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
                    <w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
                  <w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>
                    <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:spacing w:after="100"/></w:pPr></w:style>
                  <w:style w:type="paragraph" w:styleId="Centred"><w:name w:val="Centred"/><w:basedOn w:val="Normal"/>
                    <w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/></w:tabs></w:pPr></w:style>
                  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>
                    <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
                </w:styles>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        return stream.ToArray();
    }

    /// <summary>The paragraphs that hold a tab, in order.</summary>
    private static List<XElement> Entries(string body, PaginationMode mode, bool stampReferenceFields = false) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("contents.docx", Docx(body)), new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                RenderPagination = mode,
                StampReferenceFields = stampReferenceFields,
            })
            .Descendants(Xhtml.p).Where(p => p.Descendants().Any(IsTab)).ToList();

    private static bool IsTab(XElement e) => e.Attribute("data-docx-tab") != null;

    /// <summary>How the paragraph lays an element out: the element and its ancestors up to the paragraph, with their
    /// attributes (the inline styles among them).</summary>
    private static string Layout(XElement element) => string.Join(" < ", element.AncestorsAndSelf()
        .TakeWhile(e => e.Name != Xhtml.p)
        .Select(e => e.Name.LocalName + string.Concat(e.Attributes().Select(a => $" {a.Name.LocalName}=\"{a.Value}\""))));

    private static string TabLayout(XElement entry) => Layout(Assert.Single(entry.Descendants(), IsTab));

    private static string? PageNumberLayout(XElement entry) =>
        entry.Descendants(Xhtml.span).SingleOrDefault(s => !s.HasElements && s.Value == "1") is { } number ? Layout(number) : null;

    private static string? PageReferenceLayout(XElement entry) =>
        entry.Descendants().SingleOrDefault(e => (string?)e.Attribute("data-field") == "PAGEREF") is { } marker ? Layout(marker) : null;

    [Theory]
    [InlineData(PaginationMode.None, false)]
    [InlineData(PaginationMode.None, true)]
    [InlineData(PaginationMode.Paginated, false)]
    [InlineData(PaginationMode.Paginated, true)]
    public void AContentsEntryWrittenAsAFieldGetsTheRightTabSegmentAWordHyperlinkEntryGets(PaginationMode mode, bool stamp)
    {
        var entries = Entries(Contents(Entry(PageReference)), mode, stamp);
        Assert.Equal(3, entries.Count);
        // The w:hyperlink entry's leader tab fills a segment that ends on the right stop, not a width fixed up front.
        Assert.Contains("width: 100%", TabLayout(entries[0]));
        Assert.Equal(TabLayout(entries[0]), TabLayout(entries[1]));
        Assert.Equal(TabLayout(entries[0]), TabLayout(entries[2]));
    }

    [Theory]
    [InlineData(PaginationMode.None, false)]
    [InlineData(PaginationMode.None, true)]
    [InlineData(PaginationMode.Paginated, false)]
    [InlineData(PaginationMode.Paginated, true)]
    public void AContentsEntryWrittenAsAFieldKeepsItsPageNumberAtTheEndOfTheSegmentInsideItsOneLink(PaginationMode mode, bool stamp)
    {
        var entries = Entries(Contents(Entry(PageReference)), mode, stamp);
        foreach (var entry in entries)
        {
            var link = Assert.Single(entry.Descendants(Xhtml.a));
            Assert.Equal("#_Toc1", (string?)link.Attribute("href"));
            Assert.Contains(link.Descendants(), IsTab);
            // A page number shown (paginated HTML; flow HTML hides a web-hidden one) or marked for an editor is in the
            // link, after the tab in the same segment, as the w:hyperlink entry's is.
            Assert.Equal(PageNumberLayout(entries[0]), PageNumberLayout(entry));
            Assert.Equal(PageReferenceLayout(entries[0]), PageReferenceLayout(entry));
        }
        Assert.Equal(mode == PaginationMode.Paginated, PageNumberLayout(entries[1]) != null);
        if (stamp)
            Assert.All(entries, entry => Assert.Single(entry.Descendants(Xhtml.a).Single().Descendants(),
                e => (string?)e.Attribute("data-field") == "PAGEREF"));
        else
            Assert.Null(PageReferenceLayout(entries[1]));
    }

    // Outside a table of contents, too: the link's own field is not a field the tab's segment would cut.
    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void AHyperlinkFieldOutsideATableOfContentsGetsTheRightTabSegmentAWordHyperlinkGets(PaginationMode mode)
    {
        var entries = Entries(Paragraph(AsElement(Entry(PageReference))) + Paragraph(AsField(Entry(PageReference))) + Heading, mode);
        Assert.Contains("width: 100%", TabLayout(entries[0]));
        Assert.Equal(TabLayout(entries[0]), TabLayout(entries[1]));
    }

    // A tab the segment does not take (a centred one) falls back as the w:hyperlink's does, and a field after it in the
    // link's result still converts as a field: a PAGE field keeps its page-number marker.
    [Fact]
    public void AHyperlinkFieldsTabOutsideTheSegmentFallsBackAsAWordHyperlinksDoes()
    {
        var entry = Entry(Field("PAGE", Run("", Text("1"))));
        var entries = Entries(Paragraph(AsElement(entry), "Centred") + Paragraph(AsField(entry), "Centred") + Heading,
            PaginationMode.Paginated);
        // (the fallback still sizes the tab in the browser: it fills a box as wide as the text and the tab)
        Assert.Contains("width: 100%", TabLayout(entries[0]));
        Assert.Equal(TabLayout(entries[0]), TabLayout(entries[1]));
        var marker = Assert.Single(entries[1].Descendants(Xhtml.a).Single().Descendants(), e => (string?)e.Attribute("data-field") == "PAGE");
        // (the fallback: the text after the tab is not in the tab's box)
        Assert.DoesNotContain(marker.Ancestors().TakeWhile(e => e.Name != Xhtml.a), e => e.Descendants().Any(IsTab));
    }

    // The shapes a link's tabs keep out of the segment for. A form names how the content is written: in a w:hyperlink, a
    // HYPERLINK field (on its own or in a TOC field), a simple HYPERLINK field, or as the paragraph's own runs.
    private static string In(string form, string content) => form switch
    {
        "element" => AsElement(content),
        "field" => AsField(content),
        "field in TOC" => TocBegin + AsField(content) + TocEnd,
        "simple" => AsSimpleField(content),
        "plain" => content,
        _ => throw new ArgumentOutOfRangeException(nameof(form)),
    };

    private static List<XElement> Render(string body, WmlToHtmlConverterSettings settings)
    {
        settings.FabricateCssClasses = false;
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("contents.docx", Docx(body)), settings)
            .Descendants(Xhtml.p).Where(p => p.Descendants().Any(IsTab)).ToList();
    }

    private static string? Declaration(XElement element, string property) =>
        ((string?)element.Attribute("style") ?? "").Split(';').Select(d => d.Split(':', 2))
            .Where(d => d.Length == 2 && d[0].Trim() == property).Select(d => d[1].Trim()).SingleOrDefault();

    /// <summary>The elements a node sits in below its link, or its paragraph.</summary>
    private static IEnumerable<XElement> Boxes(XElement element) =>
        element.Ancestors().TakeWhile(e => e.Name != Xhtml.a && e.Name != Xhtml.p);

    private static XElement Marker(XElement entry, string kind) =>
        Assert.Single(entry.Descendants(), e => (string?)e.Attribute("data-field") == kind);

    // A line break before the link's first tab: the tab and the page number go on from the line after the break. Boxed
    // with the text before the break, they would sit on its first line, past the stop.
    [Theory]
    [InlineData("element", PaginationMode.None)]
    [InlineData("field", PaginationMode.None)]
    [InlineData("field in TOC", PaginationMode.None)]
    [InlineData("simple", PaginationMode.None)]
    [InlineData("element", PaginationMode.Paginated)]
    [InlineData("field", PaginationMode.Paginated)]
    [InlineData("field in TOC", PaginationMode.Paginated)]
    [InlineData("simple", PaginationMode.Paginated)]
    public void ALineBreakBeforeALinksTabLeavesTheTabAndPageNumberOnTheLineAfterIt(string form, PaginationMode mode)
    {
        var entry = Assert.Single(Render(Paragraph(In(form, Run(Hyperlink, Text("Line one")) + Run(Hyperlink, "<w:br/>")
                + Run(Hyperlink, Text("two")) + Run(Hyperlink, "<w:tab/>") + Field("PAGEREF _Toc1 \\h", Run("", Text("12")))))
            + Heading, new WmlToHtmlConverterSettings { RenderPagination = mode }));
        var number = Assert.Single(entry.Descendants(Xhtml.span), s => !s.HasElements && s.Value == "12");
        foreach (var element in new[] { Assert.Single(entry.Descendants(), IsTab), number })
            Assert.DoesNotContain(Boxes(element), box => box.Value.Contains("Line one"));
        Assert.Single(entry.Descendants(Xhtml.a));
    }

    // A field the converter keeps whole that holds the link's tab is not cut in two at the tab: a cross-reference marked
    // for an editor keeps one marker, its cached text the whole result.
    [Theory]
    [InlineData("element", PaginationMode.None)]
    [InlineData("field", PaginationMode.None)]
    [InlineData("field in TOC", PaginationMode.None)]
    [InlineData("field", PaginationMode.Paginated)]
    public void AMarkedReferenceHoldingALinksTabIsMarkedOnce(string form, PaginationMode mode)
    {
        var entry = Assert.Single(Render(Paragraph(In(form, Run(Hyperlink, Text("A"))
                + Field("REF _Ref9 \\h", Run("", Text("x")) + Run("", "<w:tab/>") + Run("", Text("y"))) + Run(Hyperlink, Text("z"))))
            + Heading, new WmlToHtmlConverterSettings { RenderPagination = mode, StampReferenceFields = true }));
        Assert.Equal("x\ty", (string?)Marker(entry, "REF").Attribute("data-field-cached"));
    }

    // A page-number field holding a HYPERLINK field's tab gets one marker for the paginator to fill in.
    [Theory]
    [InlineData(false)] // paginated HTML
    [InlineData(true)]  // a block render for an editor that paginates it
    public void APageFieldHoldingAHyperlinkFieldsTabGetsOneMarker(bool stampPageNumberFields)
    {
        var entry = Assert.Single(Render(Paragraph(AsField(Run(Hyperlink, Text("A"))
                + Field("PAGE", Run("", Text("1")) + Run("", "<w:tab/>") + Run("", Text("2"))))) + Heading,
            new WmlToHtmlConverterSettings
            {
                RenderPagination = stampPageNumberFields ? PaginationMode.None : PaginationMode.Paginated,
                StampPageNumberFields = stampPageNumberFields,
            }));
        Assert.Equal("12", Marker(entry, "PAGE").Value);
    }

    // A link that ends after its tab while the paragraph's text goes on: the right tab aligns that text too, but the
    // tab's box holds only the link's runs, so it must not take that text's width as well, which would push the text
    // past the stop. The same entry as the paragraph's own runs is one segment, ending on the stop.
    [Theory]
    [InlineData("element", "")]
    [InlineData("field", "")]
    [InlineData("simple", "")]
    [InlineData("element", "Value")]
    [InlineData("field", "Value")]
    [InlineData("simple", "Value")]
    public void ALinkEndingBeforeTheTextAfterItsTabDoesNotPushThatTextPastTheStop(string form, string inLink)
    {
        string Entry(string f) => Paragraph(In(f, Run(Hyperlink, Text("Name")) + Run(Hyperlink, "<w:tab/>")
            + (inLink.Length > 0 ? Run(Hyperlink, Text(inLink)) : "")) + Run("", Text(" 12")));
        var entries = Render(Entry("plain") + Entry(form) + Heading, new WmlToHtmlConverterSettings());
        decimal BoxWidth(XElement entry) => decimal.Parse(Declaration(Assert.Single(entry.Descendants(), IsTab).Ancestors()
            .First(e => Declaration(e, "display") == "inline-flex"), "width")!.Replace("pt", ""), CultureInfo.InvariantCulture);
        Assert.True(BoxWidth(entries[1]) < BoxWidth(entries[0]), $"{BoxWidth(entries[1])}pt, the stop at {BoxWidth(entries[0])}pt");
        Assert.DoesNotContain(Boxes(Assert.Single(entries[1].Descendants(Xhtml.span), s => !s.HasElements && s.Value == " 12")),
            box => box.Descendants().Any(IsTab));
    }

    // The same for a tracked insertion that ends after its tab: the tab is not given a segment as wide as the text after
    // the insertion, which the segment does not hold.
    [Fact]
    public void ATrackedInsertionEndingAfterItsTabDoesNotPushTheTextAfterItPastTheStop()
    {
        var entry = Assert.Single(Render(Paragraph(Run("", Text("Name"))
                + "<w:ins w:id=\"5\" w:author=\"Reviewer\" w:date=\"2026-01-01T00:00:00Z\">" + Run("", "<w:tab/>") + "</w:ins>"
                + Run("", Text("12"))) + Heading,
            new WmlToHtmlConverterSettings { RenderTrackedChanges = true }));
        var tab = Assert.Single(entry.Descendants(), IsTab);
        Assert.NotEmpty(tab.Ancestors(Xhtml.ins));
        Assert.DoesNotContain(tab.Ancestors().TakeWhile(e => e.Name != Xhtml.p), e => Declaration(e, "display") == "inline-flex");
    }
}
