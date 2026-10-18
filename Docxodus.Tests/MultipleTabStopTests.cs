using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class MultipleTabStopTests
{
    private static byte[] Fixture(string label, bool hyperlink, string? tabFormat = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(
                new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Cedar Unavailable Serif", HighAnsi = "Cedar Unavailable Serif" },
                    new Wp.FontSize { Val = "24" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph(new Wp.ParagraphProperties(
                new Wp.Tabs(new Wp.TabStop { Val = Wp.TabStopValues.Left, Position = 1800 },
                    new Wp.TabStop { Val = Wp.TabStopValues.Right, Position = 9000, Leader = Wp.TabStopLeaderCharValues.Dot }),
                new Wp.Indentation { Left = "1800", Hanging = "720" }),
                new Wp.Run(new Wp.Text(label)), new Wp.Run(new Wp.TabChar()),
                new Wp.Run(new Wp.Text("Cedar archive inventory")), new Wp.Run(new Wp.TabChar()),
                new Wp.Run(new Wp.Text("27"))),
                new Wp.SectionProperties(new Wp.PageSize { Width = 12240, Height = 15840 },
                    new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 })));
            if (tabFormat != null)
            {
                var tabRun = main.Document.Descendants<Wp.Run>().Where(r => r.GetFirstChild<Wp.TabChar>() != null).Last();
                var properties = new Wp.RunProperties();
                if (tabFormat == "hidden") properties.Append(new Wp.Vanish());
                if (tabFormat == "color" || tabFormat == "revision") properties.Append(new Wp.Color { Val = "0000FF" });
                if (tabFormat == "large") properties.Append(new Wp.FontSize { Val = "48" });
                if (tabFormat == "revision") properties.Append(new Wp.RunPropertiesChange(new Wp.PreviousRunProperties())
                    { Id = "1", Author = "Cedar tab reviewer" });
                tabRun.PrependChild(properties);
            }
            if (hyperlink)
            {
                var paragraph = main.Document.Body!.Elements<Wp.Paragraph>().Single();
                var runs = paragraph.Elements<Wp.Run>().ToArray();
                foreach (var run in runs) run.Remove();
                var relationship = main.AddHyperlinkRelationship(new System.Uri("https://example.invalid/cedar"), true);
                paragraph.Append(new Wp.Hyperlink(runs) { Id = relationship.Id });
            }
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("3.2.", false)]
    [InlineData("3.12.", false)]
    [InlineData("3.13.", false)]
    [InlineData("3.2.", true)]
    [InlineData("3.12.", true)]
    [InlineData("3.13.", true)]
    public void EstimatedWidths_KeepTheLeftTitleStopBeforeTheRightLeaderStop(string label, bool hyperlink)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-tabs.docx", Fixture(label, hyperlink)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var tabs = html.Descendants().Where(e => e.Attribute("data-docx-tab") != null).ToArray();
        Assert.Equal(new[] { "left", "right" }, tabs.Select(e => (string?)e.Attribute("data-docx-tab")));
        Assert.Null(tabs[0].Attribute("data-docx-tab-leader"));
        Assert.Equal("dot", (string?)tabs[1].Attribute("data-docx-tab-leader"));
        Assert.Contains(label, html.Value);
        var segment = tabs[1].Ancestors().First(e => ((string?)e.Attribute("style") ?? "").Contains("display: inline-flex"));
        Assert.Contains("Cedar archive inventory", segment.Value);
        Assert.Contains("27", segment.Value);
        var width = ((string?)segment.Attribute("style") ?? "").Split(';')
            .Select(d => d.Split(':', 2)).Where(d => d.Length == 2 && d[0].Trim() == "width")
            .Select(d => d[1].Trim()).Single();
        Assert.Equal("360pt", width);
    }
    [Theory]
    [InlineData("color")]
    [InlineData("hidden")]
    [InlineData("revision")]
    [InlineData("large")]
    public void LaterTabs_PreserveTheirRunFormattingAndRevisionMetadata(string format)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-styled-tab.docx", Fixture("3.2.", false, format)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = format == "revision" });
        var tab = html.Descendants().Single(e => (string?)e.Attribute("data-docx-tab-leader") == "dot");
        var styles = string.Join(";", tab.AncestorsAndSelf().Attributes("style").Select(a => a.Value)).ToUpperInvariant();
        if (format == "color") Assert.Contains("COLOR: #0000FF", styles);
        if (format == "hidden") Assert.Contains("DISPLAY:NONE", styles.Replace(" ", ""));
        if (format == "large") Assert.Contains("FONT-SIZE: 24PT", styles);
        if (format == "revision") Assert.Contains(tab.AncestorsAndSelf(), e => (string?)e.Attribute("data-author") == "Cedar tab reviewer");
    }

}
