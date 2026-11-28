using System;
using System.Collections.Generic;
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
}
