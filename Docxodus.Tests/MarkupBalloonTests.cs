using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The data the Word revision presentation records for Word's All Markup balloons (issue #1300):
/// formatting changes in Word's wording, author colours, and comment labels and fonts. The
/// Docxodus presentation records none of it.
/// </summary>
public class MarkupBalloonTests
{
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private const string Tracked = "w:id='31' w:author='Ines' w:date='2025-03-04T10:00:00Z'";

    private const string StylesXml =
        "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Liberation Serif' w:hAnsi='Liberation Serif'/>" +
        "<w:sz w:val='22'/></w:rPr></w:rPrDefault></w:docDefaults>" +
        "<w:style w:type='paragraph' w:default='1' w:styleId='Body'><w:name w:val='Body'/></w:style>" +
        "<w:style w:type='paragraph' w:styleId='ClauseHeading'><w:name w:val='Clause Heading'/>" +
        "<w:pPr><w:jc w:val='center'/></w:pPr></w:style>" +
        "<w:style w:type='character' w:styleId='Emphatic'><w:name w:val='Emphatic'/><w:rPr><w:i/></w:rPr></w:style>" +
        "<w:style w:type='paragraph' w:styleId='CommentText'><w:name w:val='annotation text'/>" +
        "<w:rPr><w:rFonts w:ascii='Cambria' w:hAnsi='Cambria'/></w:rPr></w:style>";

    private const string BalloonStyleXml =
        "<w:style w:type='paragraph' w:styleId='BalloonText'><w:name w:val='Balloon Text'/>" +
        "<w:rPr><w:rFonts w:ascii='Verdana' w:hAnsi='Verdana'/><w:sz w:val='14'/></w:rPr></w:style>";

    private const string NumberingXml =
        "<w:abstractNum w:abstractNumId='3'>" +
        "<w:lvl w:ilvl='0'><w:start w:val='2'/><w:numFmt w:val='upperRoman'/><w:lvlText w:val='%1.'/><w:lvlJc w:val='right'/>" +
        "<w:pPr><w:ind w:left='1080' w:hanging='540'/></w:pPr></w:lvl>" +
        "<w:lvl w:ilvl='1'><w:start w:val='1'/><w:numFmt w:val='bullet'/><w:lvlText w:val='o'/><w:lvlJc w:val='left'/>" +
        "<w:pPr><w:ind w:left='1800' w:hanging='360'/></w:pPr></w:lvl></w:abstractNum>" +
        "<w:num w:numId='7'><w:abstractNumId w:val='3'/></w:num>";

    private const string Listed = "<w:numPr><w:ilvl w:val='0'/><w:numId w:val='7'/></w:numPr>";

    [Theory]
    // Issue #1300's reproduction: bold added to a run.
    [InlineData("<w:b/>", "", "Font: Bold")]
    [InlineData("", "<w:b/><w:i/>", "Font: Not Bold, Not Italic")]
    [InlineData("<w:sz w:val='18'/><w:smallCaps/>", "", "Font: 9 pt, Small caps")]
    [InlineData("<w:rFonts w:ascii='Courier New' w:hAnsi='Courier New'/><w:dstrike/>", "", "Font: (Default) Courier New, Double strikethrough")]
    [InlineData("<w:color w:val='7030A0'/><w:u w:val='single'/>", "", "Font color: Purple, Underline")]
    [InlineData("<w:color w:val='00FF7F'/>", "", "Font color: Custom Color(RGB(0,255,127))")]
    [InlineData("<w:vertAlign w:val='subscript'/>", "<w:highlight w:val='cyan'/>", "Font: Subscript, Not Highlight")]
    // A theme colour's shade or tint is part of the colour.
    [InlineData("<w:color w:val='2F5496' w:themeColor='accent1' w:themeShade='BF'/>", "<w:color w:val='4472C4' w:themeColor='accent1'/>",
        "Font color: Accent 1, Darker 25%")]
    [InlineData("<w:color w:val='8EAADB' w:themeColor='accent1' w:themeTint='99'/>", "<w:color w:val='4472C4' w:themeColor='accent1'/>",
        "Font color: Accent 1, Lighter 40%")]
    // A character style is described by what it changes.
    [InlineData("<w:rStyle w:val='Emphatic'/>", "", "Font: Italic")]
    [InlineData("<w:caps/>", "<w:caps/>", "")]
    public void RunChange_ReadsAsWordsBalloon(string current, string recorded, string wording)
    {
        var html = Render($"<w:p><w:r><w:rPr>{current}<w:rPrChange {Tracked}><w:rPr>{recorded}</w:rPr></w:rPrChange></w:rPr>" +
            "<w:t>terms</w:t></w:r></w:p>");

        var change = html.Descendants(Html + "span").Single(e => HasClass(e, "rev-format-change"));
        Assert.Equal(wording, Data(change, "format-description"));
        Assert.Equal("#0078D4", Data(change, "author-color"));
        // The tooltip is the balloon's text: one describer.
        Assert.Equal(BalloonText(wording), (string?)change.Attribute("title"));
    }

    [Theory]
    // Issue #1300's reproduction: a half-inch left indent.
    [InlineData("<w:ind w:left='720'/>", "", "Indent: Left: 0.5\"")]
    [InlineData("<w:ind w:left='2160' w:firstLine='360'/><w:spacing w:after='120' w:line='300' w:lineRule='auto'/>", "",
        "Indent: Left: 1.5\", First line: 0.25\", Space After: 6 pt, Line spacing: Multiple 1.25 li")]
    [InlineData("<w:jc w:val='right'/><w:pageBreakBefore/>", "<w:keepNext/>", "Right, Don't keep with next, Page break before")]
    [InlineData("<w:pStyle w:val='ClauseHeading'/>", "", "Clause Heading, Centered")]
    [InlineData(Listed, "",
        "Numbered + Level: 1 + Numbering Style: I, II, III, … + Start at: 2 + Alignment: Right + Aligned at: 0.38\" + Indent at: 0.75\"")]
    [InlineData("<w:numPr><w:ilvl w:val='1'/><w:numId w:val='7'/></w:numPr>", "", "Bulleted + Level: 2 + Aligned at: 1\" + Indent at: 1.25\"")]
    [InlineData("<w:ind w:left='360'/>", Listed, "Indent: Left: 0.25\", No bullets or numbering")]
    // Moving to another level of the same list is left undescribed, its indent included.
    [InlineData("<w:numPr><w:ilvl w:val='1'/><w:numId w:val='7'/></w:numPr>", Listed, "")]
    public void ParagraphChange_ReadsAsWordsBalloon(string current, string recorded, string wording)
    {
        var html = Render($"<w:p><w:pPr>{current}<w:pPrChange {Tracked}><w:pPr>{recorded}</w:pPr></w:pPrChange></w:pPr>" +
            "<w:r><w:t>Recital.</w:t></w:r></w:p>");

        var paragraph = html.Descendants().Single(e => HasClass(e, "rev-para-format-change"));
        Assert.Equal(wording, Data(paragraph, "format-description"));
        Assert.Equal("#0078D4", Data(paragraph, "author-color"));
        Assert.Equal(BalloonText(wording), (string?)paragraph.Attribute("title"));
    }

    [Fact]
    public void MarkChangeOfAnEmptyParagraph_IsWorded_ButAListNumberIsNot()
    {
        var mark = $"<w:rPr><w:u w:val='single'/><w:rPrChange {Tracked}><w:rPr/></w:rPrChange></w:rPr>";
        var html = Render($"<w:p><w:pPr>{mark}</w:pPr></w:p>" +
            $"<w:p><w:pPr>{Listed}{mark}</w:pPr><w:r><w:t>Numbered recital.</w:t></w:r></w:p>");

        var changes = html.Descendants(Html + "span").Where(e => HasClass(e, "rev-format-change")).ToList();
        Assert.Equal("Underline", Data(changes[0], "format-description"));
        var number = changes.Single(e => e.Descendants().Any(d => d.Attribute("data-list-marker") != null));
        Assert.Null(Data(number, "format-description"));
    }

    [Fact]
    public void ChangesShareAKey_OnlyWhenTheyRecordedTheSameProperties()
    {
        var italic = $"<w:rPr><w:i/><w:rPrChange {Tracked}><w:rPr/></w:rPrChange></w:rPr>";
        var html = Render($"<w:p><w:r>{italic}<w:t>one</w:t></w:r><w:r><w:t xml:space='preserve'> - </w:t></w:r>" +
            $"<w:r>{italic}<w:t>two</w:t></w:r>" +
            $"<w:r><w:rPr><w:i/><w:color w:val='FF0000'/><w:rPrChange {Tracked}><w:rPr><w:color w:val='FF0000'/></w:rPr>" +
            "</w:rPrChange></w:rPr><w:t>three</w:t></w:r></w:p>");

        var keys = html.Descendants(Html + "span").Where(e => HasClass(e, "rev-format-change"))
            .ToDictionary(e => e.Value, e => Data(e, "format-key"));
        Assert.Equal(keys["one"], keys["two"]);
        Assert.NotEqual(keys["one"], keys["three"]);
    }

    [Fact]
    public void CommentNotes_CarryLabelColourAndFont_WithCommentAuthorsNumberedSeparately()
    {
        // Zoe makes the document's first revision but its second comment.
        var body = "<w:p><w:del w:id='40' w:author='Zoe' w:date='2025-03-04T10:00:00Z'><w:r><w:delText>old </w:delText></w:r></w:del>" +
            "<w:commentRangeStart w:id='5'/><w:r><w:t>alpha</w:t></w:r><w:commentRangeEnd w:id='5'/><w:r><w:commentReference w:id='5'/></w:r>" +
            "<w:commentRangeStart w:id='6'/><w:r><w:t>beta</w:t></w:r><w:commentRangeEnd w:id='6'/><w:r><w:commentReference w:id='6'/></w:r></w:p>";
        var comments = CommentXml(5, "Omar Quist", "OQ") + CommentXml(6, "Zoe", null);

        var notes = Notes(Render(body, comments));
        var recoloured = Notes(Render(body, comments, colors: new Dictionary<string, string> { ["Omar Quist"] = "olive" }));

        Assert.Equal(("OQ", "#0078D4", "Cambria"), (Data(notes["5"], "initials"), Data(notes["5"], "author-color"), Data(notes["5"], "font-family")));
        Assert.Equal(((string?)null, "#D13438"), (Data(notes["6"], "initials"), Data(notes["6"], "author-color")));
        Assert.Equal("olive", Data(recoloured["5"], "author-color"));
    }

    [Fact]
    public void PaginatedStaging_NamesTheBalloonFont_FromTheStyleOrWordsDefault()
    {
        string?[] Font(bool styled)
        {
            var staging = Render("<w:p><w:r><w:t>Text</w:t></w:r></w:p>", balloonStyle: styled, paginated: true)
                .Descendants(Html + "div").Single(d => (string?)d.Attribute("id") == "pagination-staging");
            return new[] { Data(staging, "balloon-font-family"), Data(staging, "balloon-font-size") };
        }

        Assert.Equal(new string?[] { "Verdana", "7" }, Font(true));
        Assert.Equal(new string?[] { "Segoe UI", "9" }, Font(false));
    }

    [Fact]
    public void DocxodusPresentation_RecordsNoBalloonData_AndKeepsItsTooltips()
    {
        var body = $"<w:p><w:r><w:rPr><w:b/><w:rPrChange {Tracked}><w:rPr/></w:rPrChange></w:rPr><w:t>buyer</w:t></w:r>" +
            "<w:commentRangeStart w:id='5'/><w:r><w:t>pays</w:t></w:r><w:commentRangeEnd w:id='5'/>" +
            "<w:r><w:commentReference w:id='5'/></w:r></w:p>" +
            $"<w:p><w:pPr><w:ind w:left='720'/><w:pPrChange {Tracked}><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>Clause.</w:t></w:r></w:p>";

        var html = Render(body, CommentXml(5, "Omar Quist", "OQ"), balloonStyle: true, paginated: true,
            presentation: RevisionPresentation.Docxodus);

        var dataNames = html.DescendantsAndSelf().Attributes().Select(a => a.Name.LocalName).ToHashSet();
        foreach (var name in new[] { "format-description", "format-key", "author-color", "initials", "font-family",
                     "balloon-font-family", "balloon-font-size" })
            Assert.DoesNotContain("data-" + name, dataNames);
        var titles = html.Descendants().Select(e => (string?)e.Attribute("title")).ToList();
        Assert.Contains("Bold added", titles);
        Assert.Contains("Paragraph formatting changed", titles);
    }

    private static string BalloonText(string wording) => wording.Length == 0 ? "Formatted" : $"Formatted: {wording}";

    private static string? Data(XElement element, string name) => (string?)element.Attribute("data-" + name);

    private static bool HasClass(XElement element, string name) =>
        ((string?)element.Attribute("class") ?? "").Split(' ').Contains(name);

    private static Dictionary<string, XElement> Notes(XElement html) =>
        html.Descendants(Html + "div").Where(d => HasClass(d, "comment-margin-note"))
            .ToDictionary(d => (string)d.Attribute("data-comment-id")!);

    private static string CommentXml(int id, string author, string? initials) =>
        $"<w:comment w:id='{id}' w:author='{author}' w:date='2025-03-04T10:00:00Z'" +
        (initials == null ? ">" : $" w:initials='{initials}'>") +
        "<w:p><w:pPr><w:pStyle w:val='CommentText'/></w:pPr><w:r><w:annotationRef/></w:r><w:r><w:t>Please check.</w:t></w:r></w:p></w:comment>";

    /// <summary>Converts a one-section document built from <paramref name="body"/> and the parts above.</summary>
    private static XElement Render(string body, string? comments = null, bool balloonStyle = false, bool paginated = false,
        RevisionPresentation presentation = RevisionPresentation.Word, Dictionary<string, string>? colors = null)
    {
        const string ns = "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'";
        var bytes = new MemoryStream();
        using (var package = WordprocessingDocument.Create(bytes, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            var parts = new List<(OpenXmlPart Part, string Xml)>
            {
                (main, $"<w:document {ns}><w:body>{body}<w:sectPr/></w:body></w:document>"),
                (main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}>{StylesXml}{(balloonStyle ? BalloonStyleXml : "")}</w:styles>"),
                (main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>"),
                (main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering {ns}>{NumberingXml}</w:numbering>"),
            };
            if (comments != null)
                parts.Add((main.AddNewPart<WordprocessingCommentsPart>(), $"<w:comments {ns}>{comments}</w:comments>"));
            foreach (var (part, xml) in parts)
            {
                using var stream = part.GetStream(FileMode.Create);
                stream.Write(System.Text.Encoding.UTF8.GetBytes(xml));
            }
        }
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            RenderTrackedChanges = true,
            RevisionPresentation = presentation,
            AuthorColors = colors,
            RenderComments = comments != null,
            CommentRenderMode = CommentRenderMode.Margin,
            RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None,
        };
        return XElement.Parse(WmlToHtmlConverter.ConvertToHtml(new WmlDocument("balloons.docx", bytes.ToArray()), settings).ToString());
    }
}
