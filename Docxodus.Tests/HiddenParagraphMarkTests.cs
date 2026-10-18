using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class HiddenParagraphMarkTests
{
    private static byte[] Fixture(bool hiddenMark, bool hiddenNumber = false, bool hiddenRun = false, bool tableBetween = false, bool hiddenDefaults = false, bool hiddenStyle = false, bool sectionEnd = false, bool bookmarks = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var defaultRunProperties = new Wp.RunPropertiesBaseStyle();
            if (hiddenDefaults) defaultRunProperties.Append(new Wp.Vanish());
            var headingRunProperties = new Wp.StyleRunProperties(new Wp.Bold());
            if (hiddenStyle) headingRunProperties.Append(new Wp.Vanish());
            var numberRunProperties = new Wp.NumberingSymbolRunProperties();
            if (hiddenNumber) numberRunProperties.Append(new Wp.Vanish());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(new Wp.RunPropertiesDefault(defaultRunProperties)),
                new Wp.Style(new Wp.StyleName { Val = "Normal" }) { Type = Wp.StyleValues.Paragraph, StyleId = "Normal", Default = true },
                new Wp.Style(new Wp.StyleName { Val = "heading 2" }, new Wp.BasedOn { Val = "Normal" },
                    new Wp.StyleParagraphProperties(new Wp.OutlineLevel { Val = 1 }),
                    headingRunProperties) { Type = Wp.StyleValues.Paragraph, StyleId = "Heading2" });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Wp.Numbering(
                new Wp.AbstractNum(new Wp.MultiLevelType { Val = Wp.MultiLevelValues.SingleLevel },
                    new Wp.Level(new Wp.StartNumberingValue { Val = 1 },
                        new Wp.NumberingFormat { Val = Wp.NumberFormatValues.Decimal },
                        new Wp.LevelSuffix { Val = Wp.LevelSuffixValues.Space }, new Wp.LevelText { Val = "%1." },
                        numberRunProperties) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                new Wp.NumberingInstance(new Wp.AbstractNumId { Val = 1 }) { NumberID = 1 });
            var heading = new Wp.Paragraph(new Wp.ParagraphProperties(
                    new Wp.ParagraphStyleId { Val = "Heading2" },
                    new Wp.NumberingProperties(new Wp.NumberingLevelReference { Val = 0 }, new Wp.NumberingId { Val = 1 }),
                    new Wp.ParagraphMarkRunProperties(new Wp.Vanish { Val = hiddenMark })),
                new Wp.Run(new Wp.Text("Cedar.")));
            if (sectionEnd) heading.ParagraphProperties!.Append(new Wp.SectionProperties());
            if (hiddenRun) heading.Append(new Wp.Run(new Wp.RunProperties(new Wp.Vanish()), new Wp.Text("Secret")));
            var body = new Wp.Body(heading);
            if (bookmarks) body.Append(new Wp.BookmarkStart { Id = "1", Name = "CedarBoundary" }, new Wp.BookmarkEnd { Id = "1" });
            if (tableBetween) body.Append(new Wp.Table(new Wp.TableProperties(), new Wp.TableGrid(new Wp.GridColumn { Width = "2000" }),
                new Wp.TableRow(new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("Boundary")))))));
            body.Append(new Wp.Paragraph(new Wp.Run(new Wp.Text(" Return the cedar folder before dusk.") { Space = SpaceProcessingModeValues.Preserve })),
                new Wp.SectionProperties());
            main.Document = new Wp.Document(body);
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    private static XElement Convert(bool hiddenMark, bool hiddenNumber = false, bool hiddenRun = false, bool tableBetween = false, bool hiddenDefaults = false, bool hiddenStyle = false, bool sectionEnd = false, bool bookmarks = false) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-hidden-mark.docx", Fixture(hiddenMark, hiddenNumber, hiddenRun, tableBetween, hiddenDefaults, hiddenStyle, sectionEnd, bookmarks)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    private static bool Hidden(XElement element) => element.AncestorsAndSelf().Any(e =>
        ((string?)e.Attribute("style") ?? "").Split(';').Any(s => s.Trim().Replace(" ", "") == "display:none"));

    [Fact]
    public void HiddenParagraphMark_FoldsTheFollowingBodyIntoTheHeading()
    {
        var html = Convert(true);
        var heading = html.Descendants().Single(e => e.Name.LocalName == "h2");
        Assert.Contains("Return the cedar folder before dusk.", heading.Value);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.LocalName == "p" && e.Value.Contains("Return the cedar folder"));
    }

    [Fact]
    public void HiddenParagraphMark_DoesNotHideTheNumber()
    {
        var html = Convert(true);
        var number = html.Descendants().Single(e => e.Nodes().OfType<XText>().Any(t => t.Value.Trim() == "1."));
        Assert.False(Hidden(number));
    }

    [Fact]
    public void ExplicitlyVisibleParagraphMark_KeepsSeparateBlocks()
    {
        var html = Convert(false);
        Assert.DoesNotContain("Return the cedar folder", html.Descendants().Single(e => e.Name.LocalName == "h2").Value);
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "p" && e.Value.Contains("Return the cedar folder"));
    }

    [Fact]
    public void NumberingLevelAndRunVisibility_ArePreserved()
    {
        var html = Convert(true, hiddenNumber: true, hiddenRun: true);
        Assert.True(Hidden(html.Descendants().Single(e => e.Nodes().OfType<XText>().Any(t => t.Value.Trim() == "1."))));
        Assert.True(Hidden(html.Descendants().Single(e => e.Nodes().OfType<XText>().Any(t => t.Value == "Secret"))));
    }

    [Fact]
    public void JoinedParagraphs_PreserveEachFormattingRevision()
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture(true)); stream.Position = 0;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Wp.Paragraph>().ToArray();
            for (var i = 0; i < paragraphs.Length; i++)
            {
                if (paragraphs[i].ParagraphProperties == null) paragraphs[i].PrependChild(new Wp.ParagraphProperties());
                paragraphs[i].ParagraphProperties!.Append(new Wp.ParagraphPropertiesChange(new Wp.ParagraphPropertiesExtended())
                    { Id = (i + 1).ToString(), Author = "Cedar reviewer " + i });
            }
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-revisions.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = true });
        var revisions = html.Descendants().Where(e => ((string?)e.Attribute("class") ?? "").Split(' ').Contains("rev-para-format-change")).ToArray();
        Assert.Equal(new[] { "Cedar reviewer 0", "Cedar reviewer 1" }, revisions.Select(e => (string?)e.Attribute("data-author")));
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void InheritedRunVisibility_StillHidesTheNumber(bool hiddenDefaults, bool hiddenStyle)
    {
        var html = Convert(true, hiddenDefaults: hiddenDefaults, hiddenStyle: hiddenStyle);
        Assert.True(Hidden(html.Descendants().Single(e => e.Nodes().OfType<XText>().Any(t => t.Value.Trim() == "1."))));
    }

    [Fact]
    public void HiddenMark_DoesNotFoldAcrossASectionBoundary()
    {
        var html = Convert(true, sectionEnd: true);
        Assert.DoesNotContain("Return the cedar folder", html.Descendants().Single(e => e.Name.LocalName == "h2").Value);
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "p" && e.Value.Contains("Return the cedar folder"));
    }

    [Fact]
    public void BookmarkRangeMarkers_DoNotCreateAPhysicalParagraphBoundary()
    {
        var html = Convert(true, bookmarks: true);
        Assert.Contains("Return the cedar folder", html.Descendants().Single(e => e.Name.LocalName == "h2").Value);
    }

    [Fact]
    public void HiddenMark_DoesNotFoldAcrossATable()
    {
        var html = Convert(true, tableBetween: true);
        Assert.DoesNotContain("Return the cedar folder", html.Descendants().Single(e => e.Name.LocalName == "h2").Value);
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "p" && e.Value.Contains("Return the cedar folder"));
    }
}
