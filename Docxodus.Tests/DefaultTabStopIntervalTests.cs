using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

/// <summary>
/// A paragraph whose tabs fall on automatic stops, under a w:defaultTabStop interval so small
/// (12 twips) or zero that the text before the first tab already runs past the stops the converter
/// generated up front. Conversion threw "Sequence contains no matching element" (issue #1269).
/// </summary>
public class DefaultTabStopIntervalTests
{
    private static readonly string[] Texts = { "Observatory inventory", "North", "East", "South", "Copper lens" };

    private static byte[] Fixture(int interval, int? customStop = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(
                new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Cedar Unavailable Serif", HighAnsi = "Cedar Unavailable Serif" },
                    new Wp.FontSize { Val = "24" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings(new Wp.DefaultTabStop { Val = (short)interval });
            var paragraph = new Wp.Paragraph();
            if (customStop is { } stop)
                paragraph.Append(new Wp.ParagraphProperties(new Wp.Tabs(new Wp.TabStop { Val = Wp.TabStopValues.Left, Position = stop })));
            // Text, tab, text, tab, text, tab, text: the issue's paragraph shape.
            for (var i = 0; i < 4; i++)
            {
                if (i > 0) paragraph.Append(new Wp.Run(new Wp.TabChar()));
                paragraph.Append(new Wp.Run(new Wp.Text(Texts[i])));
            }
            paragraph.Append(new Wp.Run(new Wp.Text(Texts[4])));
            main.Document = new Wp.Document(new Wp.Body(paragraph,
                new Wp.SectionProperties(new Wp.PageSize { Width = 12240, Height = 15840 },
                    new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0 })));
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    private static XElement Convert(byte[] docx) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("tabs.docx", docx),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    /// <summary>
    /// Each tab's stop, in twips from the line start. A tab renders inside an inline-flex segment
    /// whose width is the advance from the previous stop (the preceding text plus the tab), so the
    /// running total of those widths is where each tab lands.
    /// </summary>
    private static int[] TabStopsInTwips(XElement html)
    {
        var tabs = html.Descendants().Where(e => e.Attribute("data-docx-tab") != null).ToArray();
        var stop = 0;
        return tabs.Select(tab =>
        {
            var segment = tab.Ancestors().First(e => ((string?)e.Attribute("style") ?? "").Contains("display: inline-flex"));
            var width = ((string)segment.Attribute("style")!).Split(';')
                .Select(d => d.Split(':', 2)).Where(d => d.Length == 2 && d[0].Trim() == "width")
                .Select(d => d[1].Trim()).Single();
            Assert.EndsWith("pt", width);
            stop += (int)(decimal.Parse(width[..^2], CultureInfo.InvariantCulture) * 20m);
            return stop;
        }).ToArray();
    }

    /// <summary>Each tab's own advance, in inches, as the converter reports it.</summary>
    private static decimal[] TabAdvancesInInches(XElement html) =>
        html.Descendants().Where(e => e.Attribute("data-docx-tab") != null)
            .Select(e => decimal.Parse((string)e.Attribute("data-docx-tab-width")!, CultureInfo.InvariantCulture))
            .ToArray();

    private static void AssertTextKeptInOrder(XElement html)
    {
        var text = html.Element(html.Name.Namespace + "body")!.Value;
        var at = 0;
        foreach (var expected in Texts)
        {
            var found = text.IndexOf(expected, at, System.StringComparison.Ordinal);
            Assert.True(found >= 0, $"'{expected}' is missing after offset {at} of: {text}");
            at = found + expected.Length;
        }
    }

    [Theory]
    [InlineData(12, 12)]
    [InlineData(0, 720)]
    [InlineData(720, 720)]
    public void EveryTabLandsOnAnAutomaticStopAndNoTextIsLost(int interval, int effectiveInterval)
    {
        var html = Convert(Fixture(interval));

        AssertTextKeptInOrder(html);
        var stops = TabStopsInTwips(html);
        Assert.Equal(3, stops.Length);
        Assert.All(stops, s => Assert.Equal(0, s % effectiveInterval));
        Assert.True(stops[0] > 1200, $"the first stop {stops[0]} should lie past the 100th 12-twip stop");
        Assert.True(stops[0] < stops[1] && stops[1] < stops[2], string.Join(", ", stops));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(360)]
    public void TinyIntervalAdvancesEachTabToTheNextStopPastThePen(int? customStop)
    {
        // The first tab follows ~3,000 twips of text, past the 100 stops (1,200 twips) generated
        // up front. Each tab still advances only to the next 12-twip multiple, under 0.01in, not
        // the half-inch the converter used to substitute.
        var advances = TabAdvancesInInches(Convert(Fixture(12, customStop)));

        Assert.Equal(3, advances.Length);
        Assert.All(advances, a => Assert.InRange(a, 0.001m, 0.009m));
    }

    [Fact]
    public void ZeroIntervalFallsBackToTheHalfInchDefault()
    {
        // A zero multiplier generates no automatic stops; it is read as if the element were
        // absent, the spec's 720-twip default (LibreOffice likewise falls back to its own default).
        Assert.Equal(TabStopsInTwips(Convert(Fixture(720))), TabStopsInTwips(Convert(Fixture(0))));
    }
}
