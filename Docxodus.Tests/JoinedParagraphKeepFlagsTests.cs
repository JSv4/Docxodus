using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// A paragraph whose mark is hidden (w:pPr/w:rPr/w:vanish) renders on one line with the paragraph
// after it. In paginated output the joined block takes its keep flags from the LAST source
// paragraph, whose visible mark closes the line; its page break still comes from the FIRST.
public class JoinedParagraphKeepFlagsTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string HiddenMark = "<w:rPr><w:vanish/></w:rPr>";

    private static string Para(string pPr, string text) =>
        $"<w:p><w:pPr>{pPr}</w:pPr><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static XElement Convert(string bodyXml, string stylesXml = "")
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var s = main.GetStream(FileMode.Create))
                XDocument.Parse($"<w:document xmlns:w=\"{WNs}\"><w:body>{bodyXml}<w:sectPr/></w:body></w:document>").Save(s);
            using (var s = main.AddNewPart<StyleDefinitionsPart>().GetStream(FileMode.Create))
                XDocument.Parse($"<w:styles xmlns:w=\"{WNs}\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>{stylesXml}</w:styles>").Save(s);
            using (var s = main.AddNewPart<DocumentSettingsPart>().GetStream(FileMode.Create))
                XDocument.Parse($"<w:settings xmlns:w=\"{WNs}\"/>").Save(s);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("joined-keep-flags.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = PaginationMode.Paginated });
    }

    private static XElement Block(XElement html, string text) =>
        html.Descendants().Single(e => e.Name.LocalName is "p" or "h2" && e.Value.Contains(text));

    private static string Flags(XElement block) =>
        $"keepNext={(string?)block.Attribute("data-keep-with-next") ?? "-"} keepLines={(string?)block.Attribute("data-keep-lines") ?? "-"}";

    private static bool AvoidsBreakInside(XElement block) =>
        ((string?)block.Attribute("style") ?? "").Replace(" ", "").Contains("break-inside:avoid");

    [Theory]
    [InlineData("", "keepNext=- keepLines=-")]
    [InlineData("<w:keepNext/>", "keepNext=true keepLines=-")]
    [InlineData("<w:keepLines/>", "keepNext=- keepLines=true")]
    [InlineData("<w:keepNext w:val=\"0\"/>", "keepNext=- keepLines=-")]
    public void JoinedBlock_TakesKeepFlagsFromTheParagraphThatClosesIt(string secondPPr, string expected)
    {
        var html = Convert(Para("<w:keepNext/><w:keepLines/>" + HiddenMark, "Escrow") +
            Para(secondPPr, ". The agent holds the deposit until both parties sign.") +
            Para("", "Next clause."));
        var block = Block(html, "Escrow");
        Assert.Contains("The agent holds the deposit", block.Value);
        Assert.Equal(expected, Flags(block));
        Assert.Equal(expected.EndsWith("keepLines=true"), AvoidsBreakInside(block));
    }

    [Fact]
    public void ChainOfHiddenMarks_UsesTheLastParagraphNotTheSecond()
    {
        var html = Convert(Para("<w:keepLines/>" + HiddenMark, "Escrow") +
            Para("<w:keepLines/>" + HiddenMark, ": deposit") +
            Para("<w:keepNext/>", " held until both parties sign.") +
            Para("", "Next clause."));
        Assert.Equal("keepNext=true keepLines=-", Flags(Block(html, "Escrow")));
    }

    [Fact]
    public void HeadingStyleKeepFlags_DoNotCarryIntoTheRunInBody()
    {
        // The realistic shape: a run-in heading whose style keeps it with the next paragraph.
        var html = Convert(Para("<w:pStyle w:val=\"Heading2\"/>" + HiddenMark, "Escrow.") +
                Para("", " The agent holds the deposit until both parties sign.") +
                Para("", "Next clause."),
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/>" +
            "<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val=\"1\"/></w:pPr></w:style>");
        var block = Block(html, "Escrow.");
        Assert.Equal("h2", block.Name.LocalName);
        Assert.Contains("The agent holds the deposit", block.Value);
        Assert.Equal("keepNext=- keepLines=-", Flags(block));
        Assert.False(AvoidsBreakInside(block));
    }

    [Fact]
    public void PageBreakBefore_StillComesFromTheFirstParagraph()
    {
        var html = Convert(Para("<w:pageBreakBefore/>" + HiddenMark, "Escrow") +
            Para("", ". The agent holds the deposit."));
        Assert.Equal("true", (string?)Block(html, "Escrow").Attribute("data-page-break-before"));
    }

    [Fact]
    public void Control_VisibleMark_KeepsEachParagraphsOwnFlags()
    {
        var html = Convert(Para("<w:keepNext/><w:keepLines/>", "Escrow") +
            Para("", "The agent holds the deposit."));
        Assert.Equal("keepNext=true keepLines=true", Flags(Block(html, "Escrow")));
        Assert.Equal("keepNext=- keepLines=-", Flags(Block(html, "The agent holds")));
    }

    [Fact]
    public void Control_HiddenMarkBeforeATable_KeepsTheFirstParagraphsFlags()
    {
        var html = Convert(Para("<w:keepNext/><w:keepLines/>" + HiddenMark, "Escrow") +
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        Assert.Equal("keepNext=true keepLines=true", Flags(Block(html, "Escrow")));
    }
}
