using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

// Issue #1122: w:hRule decides what a row height bounds. A minimum height bounds the content and the
// collapsed borders are added to it; an exact height is the whole row pitch and clips overflow.
public class TableRowHeightRuleTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original ledger; no external document or template is read.
    private static byte[] AuthoredDocument(string? rule, string extraCell = "", string vAlign = "", string size = "6", string twips = "720")
    {
        var height = rule == null ? $"<w:trHeight w:val=\"{twips}\"/>" : $"<w:trHeight w:val=\"{twips}\" w:hRule=\"{rule}\"/>";
        var borders = "<w:tblBorders>" + string.Concat(new[] { "top", "left", "bottom", "right", "insideH", "insideV" }
            .Select(side => $"<w:{side} w:val=\"single\" w:sz=\"{size}\" w:color=\"444444\"/>")) + "</w:tblBorders>";
        var cell = $"<w:tc><w:tcPr><w:tcW w:w=\"3600\" w:type=\"dxa\"/>{vAlign}</w:tcPr><w:p><w:r><w:t>Birch exact record.</w:t></w:r></w:p>{extraCell}</w:tc>";
        var content = $"<w:document xmlns:w=\"{W}\"><w:body><w:tbl><w:tblPr><w:tblW w:w=\"3600\" w:type=\"dxa\"/>{borders}</w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"3600\"/></w:tblGrid>"
            + $"<w:tr><w:trPr>{height}</w:trPr>{cell}</w:tr></w:tbl>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            main.FeedData(xml);
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(document)
                .Select(e => e.Description));
        }
        return stream.ToArray();
    }

    private static XElement Render(string? rule, string extraCell = "", string vAlign = "", string size = "6", string twips = "720") =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Birch.docx", AuthoredDocument(rule, extraCell, vAlign, size, twips)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    private static Dictionary<string, string> Style(XElement? element) =>
        ((string?)element?.Attribute("style") ?? "").Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2)).ToDictionary(d => d[0].Trim(), d => d[1].Trim());

    private static XElement Row(XElement html) => html.Descendants(Xhtml.tr).Single();

    private static XElement? Clip(XElement html) =>
        html.Descendants(Xhtml.div).SingleOrDefault(d => (string?)d.Attribute("data-docx-exact-row-cell") == "true");

    [Theory]
    [InlineData("atLeast")]
    [InlineData(null)]
    public void MinimumHeightBoundsTheContentAndAddsTheCollapsedBorders(string? rule)
    {
        var html = Render(rule);

        // 36pt of content plus half of each 0.75pt horizontal border.
        Assert.Equal("36.75pt", Style(Row(html))["height"]);
        Assert.Null(Clip(html));
    }

    [Fact]
    public void ExactHeightIsTheWholeRowAndClipsItsCells()
    {
        var html = Render("exact", "<w:p><w:r><w:t>Overflow Birch continuation line.</w:t></w:r></w:p>");

        Assert.Equal("36pt", Style(Row(html))["height"]);
        var clip = Assert.IsType<XElement>(Clip(html));
        Assert.Equal("35.25pt", Style(clip)["height"]);
        Assert.Equal("hidden", Style(clip)["overflow"]);
        // Every paragraph stays in the document; the clip only hides what does not fit.
        Assert.Contains("Overflow Birch continuation line.", clip.Value);
    }

    [Fact]
    public void ExactHeightKeepsCellVerticalAlignmentInsideTheClip()
    {
        var clip = Clip(Render("exact", vAlign: "<w:vAlign w:val=\"bottom\"/>"));
        Assert.Equal("safe flex-end", Style(clip)["justify-content"]);
    }

    // Issue #1339: Word lays out an exact row of height 0 at its content height, as if it were auto.
    [Theory]
    [InlineData("exact")]
    [InlineData("atLeast")]
    public void ZeroHeightAddsNoConstraint(string rule)
    {
        var html = Render(rule, "<w:p><w:r><w:t>Second Birch line.</w:t></w:r></w:p>", twips: "0");
        Assert.False(Style(Row(html)).ContainsKey("height"));
        Assert.Null(Clip(html));
    }

    [Fact]
    public void OneTwipExactHeightStillClips()
    {
        // A positive exact height remains a constraint, even at one twip.
        var html = Render("exact", twips: "1");
        Assert.Equal("0.05pt", Style(Row(html))["height"]);
        Assert.Equal("0pt", Style(Clip(html))["height"]);
    }

    [Theory]
    [InlineData("tables-row-height-exact-zero-content-times-new-roman", 0)]
    [InlineData("tables-row-height-exact-tiny-control-times-new-roman", 6)]
    public void CapturedWordInputs_KeepZeroRowsLiveAndPositiveRowsClipped(string name, int clips)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "../../../../TestFiles/ZeroExactRows", name + ".docx"));
        var bytes = File.ReadAllBytes(path);
        using var observations = JsonDocument.Parse(File.ReadAllText(
            Path.Combine(Path.GetDirectoryName(path)!, "word-observations.json")));
        var observed = observations.RootElement.GetProperty("cases").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == name);
        Assert.Equal(observed.GetProperty("inputSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        PackageValidation.AssertValid(bytes, DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument(name + ".docx", bytes),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        Assert.Equal(clips, html.Descendants().Count(e => e.Attribute("data-docx-exact-row-cell") != null));
        if (clips == 0)
            foreach (var word in new[] { "AlphaMid1L", "BravoMid1L", "CharlieMid1L", "CharlieMid2L", "CharlieMid3L", "KiloMid1L" })
                Assert.Contains(word, html.Value);
    }

    [Fact]
    public void AutoHeightAddsNoConstraint() =>
        Assert.False(Style(Row(Render("auto"))).ContainsKey("height"));

    [Fact]
    public void ThinBordersUseTheirDeclaredWidthForPitchAndTheirLaidOutWidthForTheClip()
    {
        // w:sz="4" (0.5pt) is laid out as one CSS pixel (0.75pt) by the browser.
        var html = Render("exact", size: "4");
        Assert.Equal("36pt", Style(Row(html))["height"]);
        Assert.Equal("35.25pt", Style(Clip(html))["height"]);
        Assert.Equal("36.5pt", Style(Row(Render("atLeast", size: "4")))["height"]);
    }

    [Fact]
    public void ClippedCellsKeepTheirRunFormatting()
    {
        // The clip wraps the cell's converted content without dropping the styles its runs carry.
        var exact = Render("exact").Descendants(Xhtml.td).Single();
        var minimum = Render("atLeast").Descendants(Xhtml.td).Single();
        static string RunStyle(XElement cell) => (string?)cell.Descendants(Xhtml.span).First().Attribute("style") ?? "";
        Assert.NotEqual("", RunStyle(minimum));
        Assert.Equal(RunStyle(minimum), RunStyle(exact));
    }
}
