using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word sets an empty paragraph that only carries a section break on the preceding paragraph's
/// last line: it takes no line and no spacing of its own (#1572, measured in
/// docs/ooxml_corner_cases.md). With no paragraph line before it, it keeps its line.
/// </summary>
public class SectionMarkParagraphTests
{
    private const string Page = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" "
        + "w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>";
    private const string Spaced = "<w:spacing w:before=\"240\" w:after=\"160\"/>";
    private const string Above = "<w:p><w:r><w:t>Line above.</w:t></w:r></w:p>";
    private const string Below = "<w:p><w:r><w:t>Line below.</w:t></w:r></w:p>";
    private const string Table = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>"
        + "<w:p><w:r><w:t>Cell text.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private static string Mark(string props = "", string content = "") =>
        $"<w:p><w:pPr>{props}<w:sectPr>{Page}</w:sectPr></w:pPr>{content}</w:p>";

    private static XElement Convert(string body, bool paginated = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            using var part = main.GetStream(FileMode.Create);
            XDocument.Parse($"<w:document xmlns:w=\"{W.w}\"><w:body>{body}"
                + $"<w:sectPr><w:type w:val=\"continuous\"/>{Page}</w:sectPr></w:body></w:document>").Save(part);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("section-mark.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None,
            });
    }

    /// <summary>The body's HTML paragraphs outside tables, in order.</summary>
    private static XElement[] Paragraphs(XElement html) =>
        html.Descendants(Xhtml.p).Where(p => !p.Ancestors(Xhtml.td).Any()).ToArray();

    private static string[] Style(XElement paragraph) =>
        ((string?)paragraph.Attribute("style") ?? "").Replace(" ", "").Split(';');

    private static bool IsOneTwipBox(XElement paragraph) => Style(paragraph).Contains("line-height:0.05pt");

    [Theory]
    [InlineData("")]
    [InlineData(Spaced)]
    [InlineData("<w:rPr><w:sz w:val=\"56\"/></w:rPr>")]
    public void EmptyMarkAfterAParagraph_HasNoLineOrSpacing(string props)
    {
        var mark = Paragraphs(Convert(Above + Mark(props) + Below))[1];
        Assert.True(IsOneTwipBox(mark), mark.ToString());
        Assert.Contains("margin-top:0", Style(mark));
        Assert.Contains("margin-bottom:0", Style(mark));
        Assert.Equal("", mark.Value);
    }

    [Fact]
    public void BookmarkedEmptyMark_HasNoLine()
    {
        var bookmark = "<w:bookmarkStart w:id=\"0\" w:name=\"here\"/><w:bookmarkEnd w:id=\"0\"/>";
        Assert.True(IsOneTwipBox(Paragraphs(Convert(Above + bookmark + Mark(content: bookmark) + Below))[1]));
    }

    [Theory]
    [InlineData(Above + "<w:p/>" + Below, 1)]                    // a plain empty paragraph
    [InlineData(Above + Table + "MARK" + Below, 1)]              // nothing to ride on after a table
    [InlineData("MARK" + Below, 0)]                              // first in the body
    [InlineData(Above + "MARK" + "MARK" + Below, 2)]             // opens its own section's first line
    [InlineData(Above + "<w:p/>", 1)]                            // the body's final sectPr
    public void ParagraphWithNoLineToRideOn_KeepsItsLine(string body, int index)
    {
        var paragraph = Paragraphs(Convert(body.Replace("MARK", Mark())))[index];
        Assert.False(IsOneTwipBox(paragraph), paragraph.ToString());
        Assert.Equal(" ", paragraph.Value);
    }

    [Fact]
    public void MarkWithText_KeepsItsLine() =>
        Assert.False(IsOneTwipBox(Paragraphs(Convert(Above + Mark(content: "<w:r><w:t>Text.</w:t></w:r>") + Below))[1]));

    [Fact]
    public void PaginatedEmptyMark_IsFlaggedForThePaginatorAlone()
    {
        var flagged = Convert(Above + Table + Mark() + Below + Mark() + "<w:p/>", paginated: true)
            .Descendants().Where(e => (string?)e.Attribute("data-section-mark") == "true").ToArray();
        Assert.Equal(new[] { "" }, flagged.Select(e => e.Value));
        Assert.Same(Paragraphs(flagged[0].AncestorsAndSelf().Last())[3], flagged[0]);
    }
}
