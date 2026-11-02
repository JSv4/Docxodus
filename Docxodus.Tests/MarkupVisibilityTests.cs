using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class MarkupVisibilityTests
{
    private const string Tracked = "w:id='1' w:author='Spruce Reviewer' w:date='2026-01-02T09:00:00Z'";
    private const string Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Numbering = "<w:abstractNum w:abstractNumId='0'><w:multiLevelType w:val='singleLevel'/><w:lvl w:ilvl='0'><w:start w:val='1'/>" +
        "<w:numFmt w:val='decimal'/><w:lvlText w:val='%1.'/><w:pPr><w:tabs><w:tab w:val='num' w:pos='720'/></w:tabs>" +
        "<w:ind w:left='720' w:hanging='360'/></w:pPr></w:lvl></w:abstractNum><w:num w:numId='1'><w:abstractNumId w:val='0'/></w:num>";
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    private static byte[] Document(string body, string? comments = null, string styles = "", string? numbering = null,
        string? footnotes = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"<w:document xmlns:w='{Namespace}'><w:body>{body}" +
                "<w:sectPr><w:pgSz w:w='12240' w:h='15840'/><w:pgMar w:top='1440' w:left='1440' w:right='1440' w:bottom='1440'/></w:sectPr>" +
                "</w:body></w:document>"));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w='{Namespace}'>" +
                "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Liberation Sans' w:hAnsi='Liberation Sans'/><w:sz w:val='24'/></w:rPr></w:rPrDefault></w:docDefaults>" + styles + "</w:styles>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w='{Namespace}'/>"));
            if (comments != null)
                main.AddNewPart<WordprocessingCommentsPart>().PutXDocument(XDocument.Parse($"<w:comments xmlns:w='{Namespace}'>{comments}</w:comments>"));
            if (numbering != null)
                main.AddNewPart<NumberingDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:numbering xmlns:w='{Namespace}'>{numbering}</w:numbering>"));
            if (footnotes != null)
                main.AddNewPart<FootnotesPart>().PutXDocument(XDocument.Parse($"<w:footnotes xmlns:w='{Namespace}'>{footnotes}</w:footnotes>"));
        }
        return stream.ToArray();
    }

    private static XElement Render(byte[] source, bool formatting, bool text, int comments = -1) => XElement.Parse(
        HtmlConversionOps.ConvertToHtml(source, new HtmlConversionOptions
        {
            RenderTrackedChanges = true, ShowFormattingChanges = formatting,
            ShowInsertionsAndDeletions = text, RevisionPresentation = 1, StampAnchors = true,
            PaginationMode = 1, FabricateCssClasses = false,
            CommentRenderMode = comments,
            RenderFootnotesAndEndnotes = true,
        }));

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(false, false)]
    public void FormattingAndTextVisibilityAreIndependentWithoutChangingThePackage(bool formatting, bool text)
    {
        var source = Document($"<w:p><w:pPr><w:ind w:left='360'/><w:pPrChange {Tracked}><w:pPr/></w:pPrChange></w:pPr>" +
            $"<w:r><w:rPr><w:b/><w:rPrChange {Tracked}><w:rPr/></w:rPrChange></w:rPr><w:t>SpruceBold</w:t></w:r>" +
            $"<w:ins {Tracked}><w:r><w:t>SpruceAdded</w:t></w:r></w:ins>" +
            $"<w:del {Tracked}><w:r><w:delText>SpruceRemoved</w:delText></w:r></w:del></w:p>");
        var snapshot = source.ToArray();
        var html = Render(source, formatting, text);
        var body = html.Descendants(Html + "body").Single();
        Assert.Equal(snapshot, source);
        Assert.Contains("SpruceAdded", body.Value);
        Assert.Contains("SpruceBold", body.Value);
        Assert.Equal(text, body.Value.Contains("SpruceRemoved"));
        Assert.Equal(text, body.Descendants(Html + "ins").Any());
        Assert.Equal(text, body.Descendants(Html + "del").Any());
        Assert.Equal(formatting, body.Descendants().Attributes("data-format-description").Any());
        Assert.Equal(formatting || text, body.Descendants().Attributes("class").Any(a => a.Value.Split(' ').Contains("rev-changed-line")));
        using var stream = new MemoryStream(source);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.True(RevisionAccepter.HasTrackedRevisions(document));
    }

    [Theory]
    [InlineData("ins")]
    [InlineData("moveTo")]
    public void HiddenTextMarkupStillGroupsInsertedPageFields(string revision)
    {
        var source = Document($"<w:p><w:{revision} {Tracked}><w:r><w:fldChar w:fldCharType='begin'/></w:r>" +
            "<w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType='separate'/></w:r>" +
            $"<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r></w:{revision}></w:p>");
        var html = Render(source, true, false);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-field") == "PAGE");
        Assert.Empty(html.Descendants(Html + "ins"));
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(3, false)]
    [InlineData(1, true)]
    [InlineData(3, true)]
    public void HiddenDeletedCellsUseTheCurrentGridAndRetainSurvivorIdentity(int firstSpan, bool contentControl)
    {
        string Cell(string text, int span, bool deleted = false) => "<w:tc><w:tcPr>" +
            $"<w:gridSpan w:val='{span}'/>" + (deleted ? $"<w:cellDel {Tracked}/>" : "") +
            $"</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        var source = Document("<w:tbl><w:tblPr><w:tblW w:type='dxa' w:w='9600'/></w:tblPr><w:tblGrid>" +
            string.Concat(Enumerable.Repeat("<w:gridCol w:w='2400'/>", 4)) + "</w:tblGrid><w:tr>" +
            (contentControl ? "<w:sdt><w:sdtPr/><w:sdtContent>" : "") + Cell("SpruceA", firstSpan) +
            (contentControl ? "</w:sdtContent></w:sdt>" : "") + Cell("SpruceDeleted", 2, true) + Cell("SpruceC", 1) + "</w:tr></w:tbl>");
        var visible = Render(source, true, true);
        var hidden = Render(source, true, false);
        var cells = hidden.Descendants(Html + "td").ToArray();
        Assert.Equal(2, cells.Length);
        Assert.Equal("3", (string?)cells[0].Attribute("colspan"));
        Assert.DoesNotContain("SpruceDeleted", hidden.Descendants(Html + "body").Single().Value);
        foreach (var cell in cells)
            Assert.Equal((string?)visible.Descendants(Html + "td").Single(c => c.Value == cell.Value).Attribute("data-source-anchor-id"),
                (string?)cell.Attribute("data-source-anchor-id"));
    }

    [Fact]
    public void HiddenDeletedRowsLeaveTheInsertedRowAndItsFormattingAnnotation()
    {
        string Row(string label, string revision) => $"<w:tr><w:trPr><w:{revision} {Tracked}/>" +
            $"<w:trPrChange {Tracked}><w:trPr/></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:tc></w:tr>";
        var source = Document("<w:tbl><w:tblPr><w:tblW w:type='dxa' w:w='2400'/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w='2400'/></w:tblGrid>" + Row("SpruceGone", "del") + Row("SpruceCurrent", "ins") + "</w:tbl>");
        var hidden = Render(source, true, false);
        var row = Assert.Single(hidden.Descendants(Html + "tr"));
        Assert.Equal("SpruceCurrent", row.Value);
        Assert.Equal("rev-row-format-change", (string?)row.Attribute("class"));
    }

    [Fact]
    public void HiddenDeletedParagraphMarkJoinsTextUsingTheLastParagraphProperties()
    {
        var source = Document($"<w:p><w:pPr><w:ind w:left='0'/><w:rPr><w:del {Tracked}/></w:rPr></w:pPr>" +
            "<w:r><w:t>SpruceFirst</w:t></w:r></w:p><w:p><w:pPr><w:ind w:left='720'/></w:pPr>" +
            "<w:r><w:t>SpruceSecond</w:t></w:r></w:p>");
        var hidden = Render(source, true, false);
        var paragraph = Assert.Single(hidden.Descendants(Html + "p"), p => p.Value.Contains("SpruceFirst"));
        Assert.Equal("SpruceFirstSpruceSecond", paragraph.Value);
        Assert.Contains("margin-left: 36pt", (string?)paragraph.Attribute("style"));
        Assert.Equal(2, paragraph.DescendantsAndSelf().Attributes("data-source-anchor-id").Count());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CommentRangeBoundaryInsideHiddenTextDoesNotLoseOrLeakHighlighting(bool hiddenStart)
    {
        const string start = "<w:commentRangeStart w:id='0'/>";
        const string end = "<w:commentRangeEnd w:id='0'/>";
        string Deleted(string boundary) => $"<w:del {Tracked}>{boundary}<w:r><w:delText>SpruceRemoved</w:delText></w:r></w:del>";
        var body = "<w:p>" + (hiddenStart ? Deleted(start) : start) + "<w:r><w:t>SpruceCovered</w:t></w:r>" +
            (hiddenStart ? end : Deleted(end)) + "<w:r><w:t>SpruceOutside</w:t></w:r>" +
            "<w:r><w:commentReference w:id='0'/></w:r></w:p>";
        var source = Document(body, $"<w:comment w:id='0' w:author='Spruce Reviewer'><w:p><w:r><w:t>SpruceNote</w:t></w:r></w:p></w:comment>");
        var hidden = Render(source, true, false, comments: 1);
        Assert.Contains(hidden.Descendants(Html + "span"), span => (string?)span.Attribute("data-comment-id") == "0" && span.Value == "SpruceCovered");
        Assert.DoesNotContain(hidden.Descendants(Html + "span"), span => span.Attribute("data-comment-id") != null && span.Value.Contains("SpruceOutside"));
    }

    [Fact]
    public void DeletedOnlyCommentCannotReappearThroughTheUnreferencedCommentFallback()
    {
        var source = Document($"<w:p><w:del {Tracked}><w:r><w:delText>SpruceRemoved</w:delText></w:r>" +
            "<w:r><w:commentReference w:id='0'/></w:r></w:del><w:r><w:t>SpruceLive</w:t></w:r></w:p>",
            "<w:comment w:id='0' w:author='Spruce Reviewer'><w:p><w:r><w:t>SpruceHiddenNote</w:t></w:r></w:p></w:comment>");
        var hidden = Render(source, true, false, comments: 0);
        Assert.DoesNotContain("SpruceHiddenNote", hidden.Descendants(Html + "body").Single().Value);
        Assert.Empty(hidden.Descendants().Attributes("data-comment-id"));
        Assert.Contains("SpruceHiddenNote", Render(source, true, true, comments: 0).Descendants(Html + "body").Single().Value);
    }

    [Fact]
    public void HiddenParagraphDeletionUsesTheCurrentParagraphStyleForInheritedRunFormatting()
    {
        var source = Document($"<w:p><w:pPr><w:pStyle w:val='Heading1'/><w:rPr><w:del {Tracked}/></w:rPr></w:pPr>" +
            "<w:r><w:t>SpruceFirst</w:t></w:r></w:p><w:p><w:r><w:t>SpruceSecond</w:t></w:r></w:p>",
            styles: "<w:style w:type='paragraph' w:styleId='Heading1'><w:name w:val='heading 1'/><w:pPr><w:outlineLvl w:val='0'/></w:pPr>" +
                "<w:rPr><w:b/><w:sz w:val='40'/></w:rPr></w:style>");
        var hidden = Render(source, true, false);
        var first = Assert.Single(hidden.Descendants(Html + "span"), span => span.Value == "SpruceFirst");
        Assert.Contains("font-size: 12pt", (string?)first.Attribute("style"));
        Assert.Contains("font-weight: normal", (string?)first.Attribute("style"));
        Assert.Contains(hidden.Descendants(Html + "p"), p => p.Value == "SpruceFirstSpruceSecond");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void JoinedParagraphUsesOneCurrentListLabelAtItsStart(bool numberedLast)
    {
        const string num = "<w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/></w:numPr>";
        var source = Document($"<w:p><w:pPr>{(numberedLast ? "" : num)}<w:rPr><w:del {Tracked}/></w:rPr></w:pPr>" +
            $"<w:r><w:t>SpruceFirst</w:t></w:r></w:p><w:p><w:pPr>{(numberedLast ? num : "")}</w:pPr>" +
            "<w:r><w:t>SpruceSecond</w:t></w:r></w:p>", numbering:
            "<w:abstractNum w:abstractNumId='0'><w:multiLevelType w:val='singleLevel'/><w:lvl w:ilvl='0'><w:start w:val='1'/>" +
            "<w:numFmt w:val='decimal'/><w:lvlText w:val='%1.'/></w:lvl></w:abstractNum><w:num w:numId='1'><w:abstractNumId w:val='0'/></w:num>");
        var hidden = Render(source, true, false);
        var paragraph = Assert.Single(hidden.Descendants(Html + "p"), p => p.Value.Contains("SpruceFirst"));
        var text = string.Concat(paragraph.Value.Where(c => !char.IsWhiteSpace(c)));
        Assert.Equal((numberedLast ? "1." : "") + "SpruceFirstSpruceSecond", text);
    }

    [Fact]
    public void EmptyDeletedMemberDoesNotInsertWhitespaceBeforeTheSurvivingText()
    {
        var source = Document($"<w:p><w:pPr><w:rPr><w:del {Tracked}/></w:rPr></w:pPr>" +
            $"<w:del {Tracked}><w:r><w:delText>SpruceRemoved</w:delText></w:r></w:del></w:p>" +
            "<w:p><w:r><w:t>SpruceCurrent</w:t></w:r></w:p>");
        var hidden = Render(source, true, false);
        Assert.Equal("SpruceCurrent", Assert.Single(hidden.Descendants(Html + "p")).Value);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    [InlineData("center")]
    [InlineData("decimal")]
    public void TabAfterDeletedParagraphMarkMeasuresTheCombinedParagraph(string alignment)
    {
        string properties = $"<w:tabs><w:tab w:val='{alignment}' w:pos='2880'/></w:tabs>";
        const string prefix = "<w:r><w:t>SpruceFirst</w:t></w:r>";
        const string suffix = "<w:r><w:tab/></w:r><w:r><w:t>12.34</w:t></w:r>";
        var joined = Render(Document($"<w:p><w:pPr><w:rPr><w:del {Tracked}/></w:rPr></w:pPr>{prefix}</w:p>" +
            $"<w:p><w:pPr>{properties}</w:pPr>{suffix}</w:p>"), true, false);
        var control = Render(Document($"<w:p><w:pPr>{properties}</w:pPr>{prefix}{suffix}</w:p>"), true, false);
        var tab = Assert.Single(joined.Descendants(), e => e.Attribute("data-docx-tab") != null);
        var expected = Assert.Single(control.Descendants(), e => e.Attribute("data-docx-tab") != null);
        Assert.Equal((string?)expected.Attribute("data-docx-tab-width"), (string?)tab.Attribute("data-docx-tab-width"));
        if (alignment is "left" or "right")
            Assert.Contains(tab.Ancestors(), e => ((string?)e.Attribute("style"))?.Contains("width: 144pt") == true && e.Value.Contains("SpruceFirst"));
        Assert.Equal(2, joined.Descendants().Attributes("data-source-anchor-id").Select(a => a.Value).Distinct().Count());
    }

    [Fact]
    public void PageFieldAcrossDeletedParagraphMarkHasOneCurrentResult()
    {
        var source = Document($"<w:p><w:pPr><w:rPr><w:del {Tracked}/></w:rPr></w:pPr>" +
            "<w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r></w:p>" +
            "<w:p><w:r><w:fldChar w:fldCharType='separate'/></w:r><w:r><w:t>1</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType='end'/></w:r></w:p>");
        var hidden = Render(source, true, false);
        var field = Assert.Single(hidden.Descendants(), e => (string?)e.Attribute("data-field") == "PAGE");
        Assert.Equal("1", field.Value);
        Assert.Empty(field.Descendants().Attributes("data-source-anchor-id"));
        Assert.Equal(2, field.Ancestors().Attributes("data-source-anchor-id").Count());
        Assert.Single(hidden.Descendants(Html + "p"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnrelatedNoteCannotResurrectADeletedOnlyComment(bool secondNote)
    {
        var source = Document("<w:p><w:r><w:t>SpruceBody</w:t><w:footnoteReference w:id='1'/>" +
            (secondNote ? "<w:footnoteReference w:id='2'/>" : "") + "</w:r></w:p>",
            comments: "<w:comment w:id='0' w:author='Spruce Reviewer'><w:p><w:r><w:t>SpruceHiddenNote</w:t></w:r></w:p></w:comment>",
            footnotes: "<w:footnote w:id='1'><w:p><w:r><w:footnoteRef/></w:r><w:commentRangeStart w:id='0'/>" +
                $"<w:del {Tracked}><w:r><w:delText>SpruceRemoved</w:delText></w:r>" +
                "<w:r><w:commentReference w:id='0'/></w:r></w:del></w:p></w:footnote>" +
                (secondNote ? "<w:footnote w:id='2'><w:p><w:r><w:footnoteRef/><w:t>SpruceLiveNote</w:t></w:r></w:p></w:footnote>" : ""));
        var hidden = Render(source, true, false, comments: 0);
        Assert.DoesNotContain("SpruceHiddenNote", hidden.Descendants(Html + "body").Single().Value);
        Assert.Equal(secondNote, hidden.Descendants(Html + "body").Single().Value.Contains("SpruceLiveNote"));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    [InlineData(1)]
    public void NumberingHistoryRemainsAFormattingAnnotationWithCurrentLabelGeometry(int changedMember)
    {
        string Properties(bool history, bool deletedMark = false) =>
            "<w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/>" +
            (history ? $"<w:numberingChange {Tracked} w:original='9.'/>" : "") + "</w:numPr>" +
            (deletedMark ? $"<w:rPr><w:del {Tracked}/></w:rPr>" : "");
        string Paragraph(string text, string properties) => $"<w:p><w:pPr>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";
        var source = Document(changedMember == -1 ? Paragraph("SpruceCurrent", Properties(true)) :
            Paragraph("SpruceFirst", Properties(changedMember == 0, true)) + Paragraph("SpruceSecond", Properties(changedMember == 1)),
            numbering: Numbering);
        var hidden = Render(source, true, false);
        Assert.Single(hidden.Descendants(), e => (string?)e.Attribute("data-format-description") == "Numbering changed");
        Assert.Contains(hidden.Descendants().Attributes("data-author"), a => a.Value == "Spruce Reviewer");
        Assert.DoesNotContain("9.", hidden.Descendants(Html + "body").Single().Value);
        Assert.Empty(hidden.Descendants(Html + "ins"));
        Assert.Empty(hidden.Descendants(Html + "del"));
        var control = Render(Document(Paragraph("SpruceControl", Properties(false)), numbering: Numbering), true, false);
        Assert.Equal((string?)Assert.Single(control.Descendants(), e => e.Attribute("data-docx-tab") != null).Attribute("data-docx-tab-width"),
            (string?)Assert.Single(hidden.Descendants(), e => e.Attribute("data-docx-tab") != null).Attribute("data-docx-tab-width"));
        Assert.Empty(Render(source, false, false).Descendants().Attributes("data-format-description"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyContinuationKeepsItsFormattingAnnotationWithoutInventingAVisibleSource(bool sectionChange)
    {
        var properties = sectionChange
            ? $"<w:sectPr><w:sectPrChange {Tracked}><w:sectPr/></w:sectPrChange></w:sectPr>"
            : $"<w:ind w:left='720'/><w:pPrChange {Tracked}><w:pPr/></w:pPrChange>";
        var source = Document($"<w:p><w:pPr><w:rPr><w:del {Tracked}/></w:rPr></w:pPr><w:r><w:t>SpruceCurrent</w:t></w:r></w:p>" +
            $"<w:p><w:pPr>{properties}</w:pPr></w:p>");
        var hidden = Render(source, true, false);
        Assert.Single(hidden.Descendants().Attributes("data-author"), a => a.Value == "Spruce Reviewer");
        Assert.Single(hidden.Descendants().Attributes("data-source-anchor-id"));
        Assert.Equal("SpruceCurrent", Assert.Single(hidden.Descendants(Html + "p")).Value);
    }
}
