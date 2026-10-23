using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

// Issue #1272: a PAGE / NUMPAGES field deleted (or inserted) whole under tracked changes keeps its
// page-number marker in paginated markup, so the paginator can show each page's own number.
public class RevisedPageFieldRenderingTests
{
    private const string Rev = "w:id=\"31\" w:author=\"Review Example\" w:date=\"2026-09-12T10:00:00Z\"";

    private static string Field(string instr, string result, string instrTag = "instrText", string textTag = "t") =>
        $"<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:{instrTag} xml:space=\"preserve\"> {instr} </w:{instrTag}></w:r>"
        + $"<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:{textTag}>{result}</w:{textTag}></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static string Text(string text, string tag = "t") =>
        $"<w:r><w:{tag} xml:space=\"preserve\">{text}</w:{tag}></w:r>";

    private static readonly string LiveFooter =
        Text("Folio ") + Field("PAGE", "1") + Text(" / ") + Field("NUMPAGES", "2");

    private static readonly string DeletedFooter =
        $"<w:del {Rev}>" + Text("Folio ", "delText") + Field("PAGE", "1", "delInstrText", "delText")
        + Text(" / ", "delText") + Field("NUMPAGES", "2", "delInstrText", "delText") + "</w:del>";

    // Authored here from original literal content; no template or input DOCX is loaded.
    private static byte[] Document(string bodyParagraphContent, string footerParagraphContent)
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        const string type = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
        var parts = new (string Name, string Xml)[]
        {
            ("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + $"<Override PartName=\"/word/document.xml\" ContentType=\"{type}document.main+xml\"/>"
                + $"<Override PartName=\"/word/comments.xml\" ContentType=\"{type}comments+xml\"/>"
                + $"<Override PartName=\"/word/footnotes.xml\" ContentType=\"{type}footnotes+xml\"/>"
                + $"<Override PartName=\"/word/footer1.xml\" ContentType=\"{type}footer+xml\"/></Types>"),
            ("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + $"<Relationship Id=\"rId1\" Type=\"{rel}officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ("word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + $"<Relationship Id=\"rIdComments\" Type=\"{rel}comments\" Target=\"comments.xml\"/>"
                + $"<Relationship Id=\"rIdFootnotes\" Type=\"{rel}footnotes\" Target=\"footnotes.xml\"/>"
                + $"<Relationship Id=\"rIdFooter\" Type=\"{rel}footer\" Target=\"footer1.xml\"/></Relationships>"),
            ("word/comments.xml", $"<w:comments {ns}><w:comment w:id=\"7\" w:author=\"Review Example\" "
                + "w:date=\"2026-09-12T10:00:00Z\"><w:p><w:r><w:t>Check the desk</w:t></w:r></w:p></w:comment></w:comments>"),
            ("word/footnotes.xml", $"<w:footnotes {ns}>"
                + "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:id=\"1\"><w:p><w:r><w:t>Desk note</w:t></w:r></w:p></w:footnote></w:footnotes>"),
            ("word/document.xml", $"<w:document {ns}><w:body><w:p>{bodyParagraphContent}</w:p>"
                + "<w:sectPr><w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
                + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
                + "</w:sectPr></w:body></w:document>"),
            ("word/footer1.xml", $"<w:ftr {ns}><w:p>{footerParagraphContent}</w:p></w:ftr>"),
        };
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, xml) in parts)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(xml);
            }
        }
        return stream.ToArray();
    }

    private static XElement Footer(string footerParagraphContent, bool renderTrackedChanges = true,
        bool showDeletedContent = true, int revisionPresentation = 0)
    {
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(
            Document(Text("Lunar receiver inspection"), footerParagraphContent),
            new HtmlConversionOptions
            {
                PaginationMode = (int)PaginationMode.Paginated, RenderHeadersAndFooters = true,
                RenderTrackedChanges = renderTrackedChanges, ShowDeletedContent = showDeletedContent,
                RevisionPresentation = revisionPresentation,
            }));
        return html.Descendants().Single(e => (string?)e.Attribute("data-hf-type") == "footer-default");
    }

    // The body paragraph, unpaginated, with notes rendered and comments highlighted in the text
    // (CommentRenderMode 1), or listed after it (0).
    private static XElement Body(string bodyParagraphContent, int commentRenderMode = 1)
    {
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(Document(bodyParagraphContent, Text("Folio")),
            new HtmlConversionOptions
            {
                RenderTrackedChanges = true, CommentRenderMode = commentRenderMode, RenderFootnotesAndEndnotes = true,
            }));
        return html.Descendants().First(e => e.Name.LocalName == "p");
    }

    private static string HyperlinkField(string result, string instrTag = "instrText") =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        + $"<w:r><w:{instrTag} xml:space=\"preserve\"> HYPERLINK \"https://example.com/\" </w:{instrTag}></w:r>"
        + $"<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>{result}<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static XElement[] Named(XElement root, string localName) =>
        root.Descendants().Where(e => e.Name.LocalName == localName).ToArray();

    // Each page-number marker as "KIND=cached result".
    private static string[] FieldMarkers(XElement footer) => footer.Descendants()
        .Where(e => e.Attribute("data-field") != null)
        .Select(e => $"{(string?)e.Attribute("data-field")}={e.Value}").ToArray();

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void DeletedPageFieldsKeepTheirMarkersInsideTheDeletion(int revisionPresentation)
    {
        var footer = Footer(DeletedFooter, revisionPresentation: revisionPresentation);

        var del = Assert.Single(footer.Descendants(), e => e.Name.LocalName == "del");
        Assert.Equal(new[] { "PAGE=1", "NUMPAGES=2" }, FieldMarkers(del));
        Assert.Equal(FieldMarkers(del), FieldMarkers(footer));
    }

    // Word often writes each deleted run in a revision of its own, with its own id and sometimes
    // its own date; a move source's runs keep w:t and w:instrText. The marker moves inside the
    // revision that holds the result, so the paginator's per-page number is still shown deleted.
    // With deleted content hidden the field shows nothing, so it has no marker.
    [Theory]
    [InlineData("del", "delText", "delInstrText", false, true)]
    [InlineData("del", "delText", "delInstrText", true, true)]
    [InlineData("moveFrom", "t", "instrText", false, true)]
    [InlineData("del", "delText", "delInstrText", false, false)]
    [InlineData("moveFrom", "t", "instrText", false, false)]
    public void APageFieldDeletedRunByRunIsMarkedInsideItsDeletion(
        string wrapper, string textTag, string instrTag, bool laterResultDate, bool showDeletedContent)
    {
        var runs = new[]
        {
            Text("Folio ", textTag), "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
            $"<w:r><w:{instrTag} xml:space=\"preserve\"> PAGE </w:{instrTag}></w:r>",
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>", Text("1", textTag),
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
        };
        var footer = Footer(string.Concat(runs.Select((run, i) =>
            $"<w:{wrapper} w:id=\"{40 + i}\" w:author=\"Review Example\" "
            + $"w:date=\"2026-09-{(laterResultDate && i >= 3 ? 13 : 12)}T10:00:00Z\">{run}</w:{wrapper}>"))
            + Text(" tail"), showDeletedContent: showDeletedContent);

        if (!showDeletedContent)
        {
            Assert.Empty(FieldMarkers(footer));
            Assert.Equal(" tail", footer.Value.Replace('\u00a0', ' '));
            return;
        }
        Assert.Equal(new[] { "PAGE=1" }, FieldMarkers(footer));
        var marker = footer.Descendants().Single(e => e.Attribute("data-field") != null);
        Assert.Equal("del", marker.Parent!.Name.LocalName);
    }

    [Fact]
    public void InsertedPageFieldsKeepTheirMarkersInsideTheInsertion()
    {
        var footer = Footer($"<w:ins {Rev}>{LiveFooter}</w:ins>");

        var ins = Assert.Single(footer.Descendants(), e => e.Name.LocalName == "ins");
        Assert.Equal(new[] { "PAGE=1", "NUMPAGES=2" }, FieldMarkers(ins));
    }

    // A move source keeps its text as w:delText or, as Word writes it, as w:t.
    [Theory]
    [InlineData("moveFrom", "delText", "delInstrText")]
    [InlineData("moveFrom", "t", "instrText")]
    [InlineData("moveTo", "t", "instrText")]
    public void MovedPageFieldsKeepTheirMarkersInsideTheMove(string wrapper, string textTag, string instrTag)
    {
        var footer = Footer($"<w:{wrapper} {Rev}>{Text("Folio ", textTag)}{Field("PAGE", "1", instrTag, textTag)}</w:{wrapper}>");

        var move = Assert.Single(footer.Descendants(),
            e => (string?)e.Attribute("class") == (wrapper == "moveFrom" ? "rev-move-from" : "rev-move-to"));
        Assert.Equal(new[] { "PAGE=1" }, FieldMarkers(move));
    }

    [Fact]
    public void LivePageFieldsAreUnchanged()
    {
        var footer = Footer(LiveFooter);

        Assert.Equal(new[] { "PAGE=1", "NUMPAGES=2" }, FieldMarkers(footer));
    }

    [Fact]
    public void DeletedPageFieldsDisappearFromTheAcceptedView()
    {
        var footer = Footer(DeletedFooter, renderTrackedChanges: false);

        Assert.Empty(FieldMarkers(footer));
        Assert.DoesNotContain("Folio", footer.Value);
    }

    [Fact]
    public void DeletedPageFieldsDisappearWhenDeletedContentIsHidden()
    {
        var footer = Footer(DeletedFooter, showDeletedContent: false);

        Assert.Empty(FieldMarkers(footer));
        Assert.DoesNotContain("Folio", footer.Value);
    }

    // The same grouping reaches every field, so a hyperlink field inserted whole is now a link.
    [Fact]
    public void AHyperlinkFieldInsertedWholeIsALinkInsideTheInsertion()
    {
        var footer = Footer($"<w:ins {Rev}>" + Field("HYPERLINK \"https://example.com/\"", "Lunar desk") + "</w:ins>");

        var ins = Assert.Single(footer.Descendants(), e => e.Name.LocalName == "ins");
        var link = Assert.Single(ins.Descendants(), e => e.Name.LocalName == "a");
        Assert.Equal("https://example.com/", (string?)link.Attribute("href"));
        Assert.Equal("Lunar desk", link.Value);
    }

    [Fact]
    public void AnEditedFieldCodeMeansItsLiveInstruction()
    {
        // The old code survives as w:delInstrText beside the new w:instrText.
        var footer = Footer("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + $"<w:del {Rev}><w:r><w:delInstrText xml:space=\"preserve\"> PAGE </w:delInstrText></w:r></w:del>"
            + $"<w:ins {Rev}><w:r><w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText></w:r></w:ins>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>2</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");

        Assert.Equal(new[] { "NUMPAGES=2" }, FieldMarkers(footer));
    }

    // The paginator writes the page number into the marker, so the marker sits inside the deletion.
    [Fact]
    public void ALivePageFieldWhoseResultAloneIsDeletedIsMarkedInsideTheDeletion()
    {
        var footer = Footer("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + "<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + $"<w:del {Rev}>{Text("1", "delText")}</w:del><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");

        var marker = footer.Descendants().Single(e => e.Attribute("data-field") != null);
        Assert.Equal("PAGE=1", $"{(string?)marker.Attribute("data-field")}={marker.Value}");
        Assert.Equal("rev-del", (string?)marker.Parent!.Attribute("class"));
    }

    // Only the field's code was deleted; the field itself is live and now has no code.
    [Fact]
    public void ALiveFieldWhoseCodeWasDeletedIsNotAPageField()
    {
        var footer = Footer("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + $"<w:del {Rev}><w:r><w:delInstrText xml:space=\"preserve\"> PAGE </w:delInstrText></w:r></w:del>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>1</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");

        Assert.Empty(FieldMarkers(footer));
    }

    // A field that starts inside the deletion and ends after it, and one that starts before it
    // and ends inside it: the paragraph already groups the field, so the deletion must not again.
    [Theory]
    [InlineData("<w:del {0}><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:delInstrText> PAGE </w:delInstrText></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r></w:del><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>")]
    [InlineData("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:del {0}><w:r><w:delText>1</w:delText></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:del>")]
    public void AFieldCrossingTheDeletionIsMarkedOnce(string footerTemplate)
    {
        var footer = Footer(string.Format(footerTemplate, Rev));

        Assert.Equal(new[] { "PAGE=1" }, FieldMarkers(footer));
    }

    [Fact]
    public void ADeletedHyperlinkFieldIsALinkInsideTheDeletion()
    {
        var body = Body(Text("See ") + $"<w:del {Rev}>{HyperlinkField(Text("Lunar desk", "delText"), "delInstrText")}</w:del>");

        var link = Assert.Single(Named(Assert.Single(Named(body, "del")), "a"));
        Assert.Equal("https://example.com/", (string?)link.Attribute("href"));
        Assert.Equal("Lunar desk", link.Value);
    }

    // A tracked deletion inside a hyperlink field's result stays a deletion inside the link,
    // whether the field itself is live or inserted.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ADeletionInsideAHyperlinkFieldResultStaysADeletion(bool fieldInserted)
    {
        var field = HyperlinkField(Text("Lunar ") + $"<w:del w:id=\"40\" w:author=\"Second Reviewer\" "
            + $"w:date=\"2026-09-13T10:00:00Z\">{Text("old ", "delText")}</w:del>" + Text("desk"));
        var body = Body(Text("See ") + (fieldInserted ? $"<w:ins {Rev}>{field}</w:ins>" : field));

        var link = Assert.Single(Named(body, "a"));
        Assert.Equal("old ", Assert.Single(Named(link, "del")).Value);
        Assert.Equal(fieldInserted, link.Parent!.Name.LocalName == "ins");
    }

    [Fact]
    public void ABookmarkAndACommentInsideADeletedHyperlinkFieldStillRender()
    {
        var result = "<w:commentRangeStart w:id=\"7\"/><w:bookmarkStart w:id=\"9\" w:name=\"desk\"/>"
            + Text("Lunar", "delText") + "<w:bookmarkEnd w:id=\"9\"/><w:commentRangeEnd w:id=\"7\"/>"
            + "<w:r><w:commentReference w:id=\"7\"/></w:r>" + Text(" desk", "delText");
        var body = Body(Text("See ") + $"<w:del {Rev}>{HyperlinkField(result, "delInstrText")}</w:del>");

        // The bookmark target keeps its id on a span: an anchor cannot hold another.
        Assert.Contains(Named(body, "span"), span => (string?)span.Attribute("id") == "desk");
        Assert.DoesNotContain(Named(body, "a"), a => (string?)a.Attribute("id") == "desk");
        Assert.Contains(body.Descendants(), e => (string?)e.Attribute("class") == "comment-highlight" && e.Value == "Lunar");
    }

    [Fact]
    public void ACommentOpenedInsideADeletedHyperlinkFieldHighlightsTheTextAfterIt()
    {
        var result = "<w:commentRangeStart w:id=\"7\"/>" + Text("Lunar desk", "delText");
        var body = Body(Text("See ") + $"<w:del {Rev}>{HyperlinkField(result, "delInstrText")}</w:del>"
            + Text(" after") + "<w:commentRangeEnd w:id=\"7\"/><w:r><w:commentReference w:id=\"7\"/></w:r>" + Text(" tail"));

        var highlighted = body.Descendants().Where(e => (string?)e.Attribute("class") == "comment-highlight")
            .Select(e => e.Value.Replace('\u00a0', ' ')).ToArray();
        Assert.Contains(" after", highlighted);
    }

    [Fact]
    public void AnInsertedHyperlinkFieldInsideAHyperlinkIsNotASecondLink()
    {
        var body = Body($"<w:hyperlink w:anchor=\"desk\"><w:ins {Rev}>{HyperlinkField(Text("inner"))}</w:ins></w:hyperlink>");

        var link = Assert.Single(Named(body, "a"));
        Assert.Equal("#desk", (string?)link.Attribute("href"));
        Assert.Equal("inner", link.Value);
    }

    [Fact]
    public void AHyperlinkInsideAHyperlinkFieldResultIsNotASecondLink()
    {
        var body = Body(HyperlinkField($"<w:hyperlink w:anchor=\"desk\">{Text("inner")}</w:hyperlink>"));

        var link = Assert.Single(Named(body, "a"));
        Assert.Equal("https://example.com/", (string?)link.Attribute("href"));
        Assert.Equal("inner", link.Value);
    }

    // The anchors the converter generates inside a link, for a note reference or a comment marker,
    // stay anchors: the paginator and the editor find them by their attributes.
    [Fact]
    public void ANoteReferenceInsideAHyperlinkFieldKeepsItsAnchor()
    {
        var body = Body(HyperlinkField(Text("Lunar") + "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" + Text(" desk")));

        var link = Named(body, "a").First();
        var reference = Assert.Single(Named(link, "a"));
        Assert.Equal("1", (string?)reference.Attribute("data-footnote-id"));
        Assert.Equal("#fn-1", (string?)reference.Attribute("href"));
    }

    [Fact]
    public void ACommentMarkerInsideAHyperlinkFieldKeepsItsAnchor()
    {
        var body = Body(HyperlinkField("<w:commentRangeStart w:id=\"7\"/>" + Text("Lunar") + "<w:commentRangeEnd w:id=\"7\"/>"
            + "<w:r><w:commentReference w:id=\"7\"/></w:r>" + Text(" desk")), commentRenderMode: 0);

        var marker = Assert.Single(Named(Named(body, "a").First(), "a"));
        Assert.Contains("comment-marker", (string?)marker.Attribute("class"));
        Assert.Equal("#comment-7", (string?)marker.Attribute("href"));
    }
}
