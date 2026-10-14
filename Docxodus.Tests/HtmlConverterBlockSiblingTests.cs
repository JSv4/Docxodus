using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

/// <summary>
/// The two converter decisions that look at a block's preceding sibling: a paragraph that follows a
/// style-separator paragraph is folded into it, and a table's top margin depends on what precedes it.
/// Both now read a per-parent sibling index instead of walking siblings from the start each time.
/// </summary>
public class HtmlConverterBlockSiblingTests
{
    [Fact]
    public void Style_separator_chain_folds_following_paragraphs_into_the_first()
    {
        var html = Convert(
            SeparatorParagraph("First"),
            SeparatorParagraph("Second"),
            new Wp.Paragraph(new Wp.Run(new Wp.Text("Third"))),
            new Wp.Paragraph(new Wp.Run(new Wp.Text("Fourth"))));

        var blocks = html.Descendants().Where(e => e.Name.LocalName == "p").ToList();
        var first = Assert.Single(blocks, b => b.Value.Contains("First"));
        Assert.Contains("Second", first.Value);
        Assert.Contains("Third", first.Value);
        Assert.DoesNotContain("Fourth", first.Value);
        Assert.Single(blocks, b => b.Value.Contains("Fourth"));
        Assert.DoesNotContain(blocks, b => b != first && (b.Value.Contains("Second") || b.Value.Contains("Third")));
    }

    [Fact]
    public void Table_top_margin_depends_on_the_preceding_sibling()
    {
        var html = Convert(
            Table("T1"),
            new Wp.Paragraph(new Wp.Run(new Wp.Text("unspaced"))),
            Table("T2"),
            new Wp.Paragraph(
                new Wp.ParagraphProperties(new Wp.SpacingBetweenLines { After = "240" }),
                new Wp.Run(new Wp.Text("spaced"))),
            Table("T3"),
            Table("T4"));

        Assert.Equal(".001pt", TopMargin(html, "T1")); // first block: nothing before it
        Assert.Equal("7.5pt", TopMargin(html, "T2"));  // after a paragraph with no space after
        Assert.Equal(".001pt", TopMargin(html, "T3")); // after a paragraph that already spaces itself
        Assert.Equal("7.5pt", TopMargin(html, "T4"));  // after another table
    }

    private static Wp.Paragraph SeparatorParagraph(string text) => new(
        new Wp.ParagraphProperties(new Wp.ParagraphMarkRunProperties(new Wp.SpecVanish())),
        new Wp.Run(new Wp.Text(text)));

    private static Wp.Table Table(string text) => new(
        new Wp.TableProperties(new Wp.TableWidth { Width = "5000", Type = Wp.TableWidthUnitValues.Pct }),
        new Wp.TableGrid(new Wp.GridColumn { Width = "4000" }),
        new Wp.TableRow(new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text(text))))));

    private static string? TopMargin(XElement html, string cellText)
    {
        var table = html.Descendants().Single(e => e.Name.LocalName == "table" && e.Value.Contains(cellText));
        var style = (string?)table.Attribute("style") ?? string.Empty;
        return style.Split(';').Select(d => d.Split(':')).Where(kv => kv.Length == 2 && kv[0].Trim() == "margin-top")
            .Select(kv => kv[1].Trim()).SingleOrDefault();
    }

    private static XElement Convert(params OpenXmlElement[] blocks)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document = new Wp.Document(new Wp.Body(blocks));
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("blocks.docx", ms.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
    }
}
