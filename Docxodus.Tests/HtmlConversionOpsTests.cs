#nullable enable
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class HtmlConversionOpsTests
{
    private readonly Xunit.Abstractions.ITestOutputHelper _output;
    public HtmlConversionOpsTests(Xunit.Abstractions.ITestOutputHelper output) => _output = output;

    private static byte[] TourPlanBytes() =>
        File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC001-5DayTourPlanTemplate.docx"));

    private const string TransitionalMain =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string TransitionalRels =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictRels =
        "http://purl.oclc.org/ooxml/officeDocument/relationships";

    private static byte[] TabStopDocxBytes(
        Wp.TabStopValues alignment,
        string before,
        string after,
        Wp.TabStopLeaderCharValues? leader = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var tabStop = new Wp.TabStop { Val = alignment, Position = 5760 };
            if (leader != null)
                tabStop.Leader = leader.Value;

            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.ParagraphProperties(new Wp.Tabs(tabStop)),
                        new Wp.Run(new Wp.Text(before)),
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text(after))),
                    new Wp.SectionProperties(
                        new Wp.PageSize { Width = 12240, Height = 15840 },
                        new Wp.PageMargin
                        {
                            Top = 1440,
                            Right = 1440,
                            Bottom = 1440,
                            Left = 1440,
                        })));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(
                            new Wp.RunFonts
                            {
                                Ascii = "MissingTabGeometryFont",
                                HighAnsi = "MissingTabGeometryFont",
                            },
                            new Wp.FontSize { Val = "24" }))));
            main.Document.Save();
        }

        return stream.ToArray();
    }

    private static byte[] StrictDocumentOnlyDocxBytes(string text)
    {
        var bytes = DocumentOnlyDocxBytes(text);
        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                    !entry.FullName.EndsWith(".rels", System.StringComparison.OrdinalIgnoreCase))
                    continue;

                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    xml = reader.ReadToEnd();
                var strict = xml
                    .Replace(TransitionalMain, StrictMain, System.StringComparison.Ordinal)
                    .Replace(TransitionalRels, StrictRels, System.StringComparison.Ordinal);
                if (strict == xml)
                    continue;
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(strict);
            }
        }
        return ms.ToArray();
    }

    [Fact]
    public void HCO001_ConvertBytes_ProducesHtmlWithPrefix()
    {
        var options = new HtmlConversionOptions { CssClassPrefix = "zz-" };

        string html = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), options);

        Assert.Contains("<html", html);
        Assert.Contains("zz-", html);
    }

    [Fact]
    public void HCO003_PaginatedHtml_LeavesTheCaptureHostBodyFlush()
    {
        // Paginated HTML is injected into the React viewer's capture host. Its fixed-size page boxes
        // own geometry, so a converter-level body margin must not shrink/overflow that host. Standalone
        // conversion retains the readable 20px margin for existing consumers.
        string paginated = HtmlConversionOps.ConvertToHtml(TourPlanBytes(),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });
        string standalone = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), new HtmlConversionOptions());

        Assert.Contains("body { font-family: Arial, sans-serif; margin: 0; }", paginated);
        Assert.Contains("body { font-family: Arial, sans-serif; margin: 20px; }", standalone);
    }

    [Fact]
    public void HCO076_PaginatedHtml_UsesDocumentPageSizeWithoutOuterPrintMargin()
    {
        // The paginator has already applied the Word margins within each page box. Its capture
        // path must advertise the paper size to Chromium without applying those margins again.
        string paginated = HtmlConversionOps.ConvertToHtml(
            PageSizedDocxBytes(width: 11906, height: 16838),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });
        string standalone = HtmlConversionOps.ConvertToHtml(
            PageSizedDocxBytes(width: 11906, height: 16838), new HtmlConversionOptions());

        Assert.Contains("@page", paginated);
        Assert.Contains("size: 8.27in 11.69in;", paginated);
        Assert.Contains("margin: 0;", paginated);
        Assert.DoesNotContain("@page docxodus-section-", paginated);
        Assert.DoesNotContain("@page", standalone);
    }

    [Fact]
    public void HCO077_MixedPaginatedSections_UseNamedPrintPages()
    {
        // The staging and final paginator page boxes retain their data-section-index. Named
        // pages let Chromium print each section at its own physical size without relying on a
        // caller to customize page.pdf options.
        string html = HtmlConversionOps.ConvertToHtml(MixedPageSizedDocxBytes(),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });

        Assert.Contains("@page docxodus-section-0", html);
        Assert.Contains("size: 8.27in 11.69in;", html);
        Assert.Contains("@page docxodus-section-1", html);
        Assert.Contains("size: 11.00in 8.50in;", html);
        Assert.Contains(".page-box[data-section-index=\"0\"]", html);
        Assert.Contains("page: docxodus-section-0;", html);
        Assert.Contains(".page-box[data-section-index=\"1\"]", html);
        Assert.Contains("page: docxodus-section-1;", html);
        Assert.DoesNotContain("@page {", html);
        Assert.Contains("data-section-index=\"0\"", html);
        Assert.Contains("data-page-width=\"595.3\"", html);
        Assert.Contains("data-page-height=\"841.9\"", html);
        Assert.Contains("data-section-index=\"1\"", html);
        Assert.Contains("data-page-width=\"792.0\"", html);
        Assert.Contains("data-page-height=\"612.0\"", html);
    }

    [Fact]
    public void HCO078_PaginatedHtml_WithoutSectionProperties_UsesLetterPageSize()
    {
        string html = HtmlConversionOps.ConvertToHtml(DocumentOnlyDocxBytes("No section properties"),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });

        Assert.Contains("size: 8.50in 11.00in;", html);
        Assert.Contains("margin: 0;", html);
    }

    [Fact]
    public void HCO079_WebHiddenRun_IsVisibleOnlyInPaginatedPrintLayout()
    {
        // Word uses w:webHidden on cached TOC page numbers: they are hidden in Web layout but
        // remain visible in Print layout. Build the document independently so this regression
        // does not depend on an external fixture or snapshot.
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.Text("Visible TOC title")),
                        new Wp.Run(
                            new Wp.RunProperties(new Wp.WebHidden()),
                            new Wp.TabChar(),
                            new Wp.Text("7"))),
                    new Wp.SectionProperties()));
            main.Document.Save();
        }

        string paginated = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions
            {
                PaginationMode = (int)PaginationMode.Paginated,
                FabricateCssClasses = false,
            });
        string web = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("Visible TOC title", paginated);
        Assert.Contains(">7<", paginated);
        Assert.Contains("Visible TOC title", web);
        Assert.DoesNotContain(">7<", web);
    }

    [Theory]
    [InlineData("right", "W", "3.500", "width: 3.900in")]
    [InlineData("center", "WWWW", "3.400", "width: 3.800in")]
    [InlineData("decimal", "12.34", "3.400", "width: 3.800in")]
    public void HCO088_AlignedTabWidth_MeasuresOnlyFollowingText(
        string alignment, string after, string expectedTabWidth, string expectedPrecedingWidth)
    {
        // A tab's following segment is measured exactly as authored. The old unconditional
        // trailing blank shifted right, center, and decimal alignment before the HTML renderer
        // saw the run. Use an unavailable font to make the deterministic fallback explicit.
        string html = HtmlConversionOps.ConvertToHtml(
            TabStopDocxBytes(
                alignment switch
                {
                    "right" => Wp.TabStopValues.Right,
                    "center" => Wp.TabStopValues.Center,
                    "decimal" => Wp.TabStopValues.Decimal,
                    _ => throw new System.ArgumentOutOfRangeException(nameof(alignment)),
                },
                "iiii",
                after),
            new HtmlConversionOptions { FabricateCssClasses = false });
        var root = XElement.Parse(html);
        var tab = root.Descendants()
            .Single(element => (string?)element.Attribute("data-docx-tab") == alignment);

        Assert.Equal(alignment, (string?)tab.Attribute("data-docx-tab"));
        Assert.Equal(expectedTabWidth, (string?)tab.Attribute("data-docx-tab-width"));
        Assert.Contains(expectedPrecedingWidth, (string?)tab.Parent!.Attribute("style"));
        Assert.Contains($"iiii{after}", tab.Parent!.Parent!.Value);
        Assert.Empty(tab.Nodes());
    }

    [Fact]
    public void HCO089_RightTabWidth_TracksCurrentPositionAndFollowingRunWidth()
    {
        static (decimal TabWidth, decimal PrecedingWidth) Geometry(string before, string after)
        {
            string html = HtmlConversionOps.ConvertToHtml(
                TabStopDocxBytes(Wp.TabStopValues.Right, before, after),
                new HtmlConversionOptions { FabricateCssClasses = false });
            var root = XElement.Parse(html);
            var tab = root.Descendants()
                .Single(element => (string?)element.Attribute("data-docx-tab") == "right");
            var tabStyle = (string)tab.Attribute("style")!;
            var precedingStyle = (string)tab.Parent!.Attribute("style")!;
            return (
                decimal.Parse(
                    tabStyle.Split("width: ")[1].Split("in")[0],
                    System.Globalization.CultureInfo.InvariantCulture),
                decimal.Parse(
                    precedingStyle.Split("width: ")[1].Split("in")[0],
                    System.Globalization.CultureInfo.InvariantCulture));
        }

        var narrowCurrent = Geometry("iiii", "W");
        var wideCurrent = Geometry("iiiiiiii", "W");
        var wideFollowing = Geometry("iiii", "WWWWW");

        Assert.Equal(3.50m, narrowCurrent.TabWidth);
        Assert.Equal(3.10m, wideCurrent.TabWidth);
        Assert.Equal(narrowCurrent.PrecedingWidth, wideCurrent.PrecedingWidth);
        Assert.Equal(3.900m, narrowCurrent.PrecedingWidth);
        Assert.Equal(3.500m, wideFollowing.PrecedingWidth);
    }

    [Fact]
    public void HCO090_DotLeader_FillsItsCalculatedAdvance()
    {
        string html = HtmlConversionOps.ConvertToHtml(
            TabStopDocxBytes(
                Wp.TabStopValues.Right,
                "iiii",
                "W",
                Wp.TabStopLeaderCharValues.Dot),
            new HtmlConversionOptions { FabricateCssClasses = false });
        var root = XElement.Parse(html);
        var leader = root.Descendants()
            .Single(element => (string?)element.Attribute("data-docx-tab-leader") == "dot");

        Assert.Equal("right", (string?)leader.Attribute("data-docx-tab"));
        Assert.Empty(leader.Value);
        Assert.Contains("width: 3.50in", (string?)leader.Attribute("style"));
        Assert.Contains("border-bottom: 1px dotted currentColor", (string?)leader.Attribute("style"));
    }

    [Fact]
    public void HCO094_ListMarkerSuffixTab_AdvancesToTextIndentNotNextDefaultStop()
    {
        // Issue #415: on a hanging-indent numbered paragraph (marker at left − hanging, text at
        // left), the marker's suffix tab must advance to the paragraph's text indent when the
        // number ends before it — not to the next w:defaultTabStop multiple. The flat general
        // 0.6 em/char width estimate nearly doubles "(a)"/"(iii)", overshooting the text-indent
        // stop; marker runs therefore measure through the character-class estimate. The marker
        // wrapper's width equals (chosen stop − marker start), so the correct stop makes every
        // wrapper exactly hanging-width wide (360 twips = 0.25") regardless of the estimated
        // marker width.
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            static Wp.Paragraph ListParagraph(int ilvl, string text) => new(
                new Wp.ParagraphProperties(
                    new Wp.NumberingProperties(
                        new Wp.NumberingLevelReference { Val = ilvl },
                        new Wp.NumberingId { Val = 1 })),
                new Wp.Run(new Wp.Text(text)));

            main.Document = new Wp.Document(
                new Wp.Body(
                    ListParagraph(0, "Confidential Information definition."),
                    ListParagraph(1, "advisory and analysis services;"),
                    ListParagraph(1, "implementation and configuration services; and"),
                    ListParagraph(1, "training and knowledge-transfer services."),
                    new Wp.SectionProperties(
                        new Wp.PageSize { Width = 12240, Height = 15840 },
                        new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));

            static Wp.Level NumberingLevel(
                int ilvl, Wp.NumberFormatValues format, string text, int left) => new(
                new Wp.StartNumberingValue { Val = 1 },
                new Wp.NumberingFormat { Val = format },
                new Wp.LevelText { Val = text },
                new Wp.LevelJustification { Val = Wp.LevelJustificationValues.Left },
                new Wp.PreviousParagraphProperties(
                    new Wp.Indentation { Left = left.ToString(), Hanging = "360" }))
            {
                LevelIndex = ilvl,
            };

            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Wp.Numbering(
                new Wp.AbstractNum(
                    NumberingLevel(0, Wp.NumberFormatValues.LowerLetter, "(%1)", 1080),
                    NumberingLevel(1, Wp.NumberFormatValues.LowerRoman, "(%2)", 1800))
                {
                    AbstractNumberId = 0,
                },
                new Wp.NumberingInstance(
                    new Wp.AbstractNumId { Val = 0 })
                {
                    NumberID = 1,
                });

            // An unavailable font forces the deterministic character-class width estimate.
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(
                            new Wp.RunFonts
                            {
                                Ascii = "MissingTabGeometryFont",
                                HighAnsi = "MissingTabGeometryFont",
                            },
                            new Wp.FontSize { Val = "24" }))));
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(stream.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });
        var root = XElement.Parse(html);

        // The outer marker wrapper (number + suffix tab) carries the width; markers themselves
        // render as "(a)", "(i)", "(ii)", "(iii)".
        var wrappers = root.Descendants()
            .Where(element => (string?)element.Attribute("data-list-marker") == "true" &&
                ((string?)element.Attribute("style") ?? "").Contains("width:"))
            .ToList();

        Assert.Equal(4, wrappers.Count);
        Assert.Equal(new[] { "(a)", "(i)", "(ii)", "(iii)" }, wrappers.Select(w => w.Value.Trim()));
        Assert.All(wrappers, wrapper =>
            Assert.Contains("width: 0.250in", (string?)wrapper.Attribute("style")));
    }

    [Fact]
    public void HCO086_PaginatedFootnoteSeparator_UsesWordDefaultTwoInchWidth()
    {
        // The CSS is emitted only when note rendering is enabled. Use an independently generated
        // package so the regression has no fixture or snapshot licensing dependency.
        string html = HtmlConversionOps.ConvertToHtml(DocumentOnlyDocxBytes("Body with note CSS"),
            new HtmlConversionOptions
            {
                PaginationMode = (int)PaginationMode.Paginated,
                RenderFootnotesAndEndnotes = true,
                FabricateCssClasses = false,
            });

        Assert.Contains("width: 2in;", html);
        Assert.Contains("max-width: 100%;", html);
        Assert.DoesNotContain("width: 33%;", html);
    }

    // Chart namespace for the generated chart-package tests below (HCO087/HCO095-HCO098).
    private static readonly XNamespace ChartC =
        "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private static XElement GeneratedChartSeries(int index, string name, double[] values)
    {
        var c = ChartC;
        var categories = new[] { "Alpha", "Beta", "Gamma" };
        return new XElement(c + "ser",
            new XElement(c + "idx", new XAttribute("val", index)),
            new XElement(c + "order", new XAttribute("val", index)),
            new XElement(c + "tx",
                new XElement(c + "strRef",
                    new XElement(c + "strCache",
                        new XElement(c + "ptCount", new XAttribute("val", 1)),
                        new XElement(c + "pt", new XAttribute("idx", 0),
                            new XElement(c + "v", name))))),
            new XElement(c + "cat",
                new XElement(c + "strRef",
                    new XElement(c + "strCache",
                        new XElement(c + "ptCount", new XAttribute("val", categories.Length)),
                        categories.Select((value, point) =>
                            new XElement(c + "pt", new XAttribute("idx", point),
                                new XElement(c + "v", value)))))),
            new XElement(c + "val",
                new XElement(c + "numRef",
                    new XElement(c + "numCache",
                        new XElement(c + "formatCode", "General"),
                        new XElement(c + "ptCount", new XAttribute("val", values.Length)),
                        values.Select((value, point) =>
                            new XElement(c + "pt", new XAttribute("idx", point),
                                new XElement(c + "v",
                                    value.ToString("R", System.Globalization.CultureInfo.InvariantCulture))))))));
    }

    // Build the package independently: chart caches are the portable OOXML display source,
    // while the embedded workbook is optional and must not be needed by the browser renderer.
    private static byte[] GeneratedChartDocxBytes(params XElement[] plotAreaChildren)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var c = ChartC;
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var chartPart = main.AddNewPart<ChartPart>();
            var chartRelationshipId = main.GetIdOfPart(chartPart);
            chartPart.PutXDocument(new XDocument(
                new XElement(c + "chartSpace",
                    new XElement(c + "chart",
                        new XElement(c + "title"),
                        new XElement(c + "plotArea", plotAreaChildren),
                        new XElement(c + "legend")))));
            main.PutXDocument(new XDocument(
                new XElement(w + "document",
                    new XAttribute(XNamespace.Xmlns + "w", w),
                    new XAttribute(XNamespace.Xmlns + "wp", wp),
                    new XAttribute(XNamespace.Xmlns + "a", a),
                    new XAttribute(XNamespace.Xmlns + "c", c),
                    new XAttribute(XNamespace.Xmlns + "r", r),
                    new XElement(w + "body",
                        new XElement(w + "p",
                            new XElement(w + "r",
                                new XElement(w + "drawing",
                                    new XElement(wp + "inline",
                                        new XElement(wp + "extent",
                                            new XAttribute("cx", 5486400),
                                            new XAttribute("cy", 3200400)),
                                        new XElement(wp + "docPr",
                                            new XAttribute("id", 1),
                                            new XAttribute("name", "Generated chart")),
                                        new XElement(a + "graphic",
                                            new XElement(a + "graphicData",
                                                new XAttribute("uri", c.NamespaceName),
                                                new XElement(c + "chart",
                                                    new XAttribute(r + "id", chartRelationshipId)))))))),
                        new XElement(w + "sectPr")))));
        }
        return stream.ToArray();
    }

    private static XElement ConvertChartSvg(byte[] bytes)
    {
        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });
        var root = XElement.Parse(html);
        return root.Descendants().Single(element => element.Name.LocalName == "svg");
    }

    [Theory]
    [InlineData("col", "column")]
    [InlineData("bar", "bar")]
    public void HCO087_CachedClusteredChart_RendersAccessibleSvgAtStoredExtent(
        string barDirection, string expectedChartType)
    {
        var c = ChartC;
        var svg = ConvertChartSvg(GeneratedChartDocxBytes(
            new XElement(c + "barChart",
                new XElement(c + "barDir", new XAttribute("val", barDirection)),
                new XElement(c + "grouping", new XAttribute("val", "clustered")),
                GeneratedChartSeries(0, "North", new[] { 2.0, 4.0, 3.0 }),
                GeneratedChartSeries(1, "South", new[] { 3.0, 1.0, 5.0 }),
                new XElement(c + "gapWidth", new XAttribute("val", 150))),
            new XElement(c + "valAx",
                new XElement(c + "scaling"))));
        var bars = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-bar")
            .ToList();

        Assert.Equal(expectedChartType, (string?)svg.Attribute("data-chart-type"));
        Assert.Equal("Generated chart", (string?)svg.Attribute("aria-label"));
        Assert.Contains("width: 432pt", (string?)svg.Attribute("style"));
        Assert.Contains("height: 252pt", (string?)svg.Attribute("style"));
        Assert.Equal(6, bars.Count);
        Assert.Contains("Chart Title", svg.Value);
        Assert.Contains("Alpha", svg.Value);
        Assert.Contains("North", svg.Value);
        Assert.Contains(bars, bar => (string?)bar.Attribute("fill") == "#5B9BD5");
        Assert.Contains(bars, bar => (string?)bar.Attribute("fill") == "#ED7D31");
    }

    private static XElement ConvertFixtureChartSvg(string fixture) =>
        ConvertChartSvg(File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            Path.Combine(fixture.Split('/')))));

    [Fact]
    public void HCO088_CachedStackedColumnChart_RendersOneStackPerCategory()
    {
        var svg = ConvertFixtureChartSvg("VP/VP001-Chart-Stacked-Column.docx");
        var bars = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-bar")
            .ToList();

        Assert.Equal("column-stacked", (string?)svg.Attribute("data-chart-type"));
        // Three series across four categories, cached fully: 12 stacked segments.
        Assert.Equal(12, bars.Count);
        // Segments of one category stack on a single x, one on top of another.
        var firstCategory = bars
            .Where(bar => (string?)bar.Attribute("data-chart-category") == "0")
            .ToList();
        Assert.Equal(3, firstCategory.Count);
        Assert.Single(firstCategory.Select(bar => (string?)bar.Attribute("x")).Distinct());
        Assert.Equal(3, firstCategory.Select(bar => (string?)bar.Attribute("y")).Distinct().Count());
        Assert.Contains("Series 1", svg.Value);
    }

    [Fact]
    public void HCO089_CachedPie3DChart_RendersSlicesWithPerPointColors()
    {
        var svg = ConvertFixtureChartSvg("CU002-Chart-Cached-Data-02.docx");
        var slices = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-slice")
            .ToList();

        Assert.Equal("pie", (string?)svg.Attribute("data-chart-type"));
        Assert.Equal(4, slices.Count);
        // Each data point carries its own accent color, so all slices differ.
        Assert.Equal(4, slices.Select(slice => (string?)slice.Attribute("fill")).Distinct().Count());
        Assert.Equal("320", (string?)slices[0].Attribute("data-chart-value"));
        // The pie legend lists categories, not the single series.
        Assert.Contains("Cars", svg.Value);
        Assert.Contains("Boats", svg.Value);
    }

    [Fact]
    public void HCO090_CachedLineChart_RendersOnePolylinePerSeriesWithDateCategories()
    {
        var svg = ConvertFixtureChartSvg("CU004-Chart-Cached-Data-04.docx");
        var lines = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-line")
            .ToList();

        Assert.Equal("line", (string?)svg.Attribute("data-chart-type"));
        Assert.Equal(3, lines.Count);
        Assert.All(lines, line => Assert.Equal(20,
            ((string?)line.Attribute("points"))!.Split(' ').Length));
        Assert.Equal(3, lines.Select(line => (string?)line.Attribute("stroke")).Distinct().Count());
        // The date axis caches serial day numbers (41518 = 9/1/2013); labels render as dates.
        Assert.Contains("9/1/2013", svg.Value);
        Assert.Contains("Car", svg.Value);
    }

    [Fact]
    public void HCO095_CachedPercentStackedColumnChart_NormalizesEveryCategoryTo100()
    {
        var c = ChartC;
        var svg = ConvertChartSvg(GeneratedChartDocxBytes(
            new XElement(c + "barChart",
                new XElement(c + "barDir", new XAttribute("val", "col")),
                new XElement(c + "grouping", new XAttribute("val", "percentStacked")),
                GeneratedChartSeries(0, "North", new[] { 1.0, 3.0, 2.0 }),
                GeneratedChartSeries(1, "South", new[] { 3.0, 1.0, 6.0 })),
            new XElement(c + "valAx",
                new XElement(c + "scaling"))));
        var bars = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-bar")
            .ToList();

        Assert.Equal("column-percent-stacked", (string?)svg.Attribute("data-chart-type"));
        Assert.Equal(6, bars.Count);
        // Unequal raw totals per category all normalize to a full-height 100% stack.
        foreach (var category in bars.GroupBy(bar => (string?)bar.Attribute("data-chart-category")))
            Assert.Equal(100.0, category.Sum(bar => double.Parse(
                (string)bar.Attribute("data-chart-value")!,
                System.Globalization.CultureInfo.InvariantCulture)), 6);
        // The value axis pins at 100% with %-suffixed ticks instead of adding tick headroom.
        Assert.Contains(svg.Descendants().Where(element => element.Name.LocalName == "text"),
            text => text.Value == "100%");
    }

    [Fact]
    public void HCO096_CachedDoughnutChart_RendersRingSlicesWithHole()
    {
        var c = ChartC;
        var svg = ConvertChartSvg(GeneratedChartDocxBytes(
            new XElement(c + "doughnutChart",
                GeneratedChartSeries(0, "Share", new[] { 5.0, 3.0, 2.0 }),
                new XElement(c + "holeSize", new XAttribute("val", 50)))));
        var slices = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-slice")
            .ToList();

        Assert.Equal("doughnut", (string?)svg.Attribute("data-chart-type"));
        Assert.Equal(3, slices.Count);
        // A doughnut slice is a ring segment (outer arc + inner hole arc), never a wedge
        // reaching the center like a pie slice.
        Assert.All(slices, slice =>
            Assert.Equal(2, ((string)slice.Attribute("d")!).Split('A').Length - 1));
        // The pie-family legend lists categories, not the single series.
        Assert.Contains("Alpha", svg.Value);
        Assert.Contains("Gamma", svg.Value);
    }

    [Fact]
    public void HCO097_CachedStackedAreaChart_StacksSecondBandOnFirst()
    {
        var c = ChartC;
        var svg = ConvertChartSvg(GeneratedChartDocxBytes(
            new XElement(c + "areaChart",
                new XElement(c + "grouping", new XAttribute("val", "stacked")),
                GeneratedChartSeries(0, "North", new[] { 2.0, 4.0, 3.0 }),
                GeneratedChartSeries(1, "South", new[] { 3.0, 1.0, 5.0 })),
            new XElement(c + "valAx",
                new XElement(c + "scaling"))));
        var areas = svg.Descendants()
            .Where(element => (string?)element.Attribute("class") == "docx-chart-area")
            .ToList();

        Assert.Equal("area-stacked", (string?)svg.Attribute("data-chart-type"));
        Assert.Equal(2, areas.Count);
        static string[] Points(XElement area) => ((string)area.Attribute("points")!).Split(' ');
        Assert.All(areas, area => Assert.Equal(6, Points(area).Length));
        // The second band's bottom edge is the first band's top edge: stacked, not overlaid.
        Assert.Equal(Points(areas[0]).Take(3),
            Points(areas[1]).Skip(3).Reverse());
    }

    [Fact]
    public void HCO098_UnsupportedChartFamily_DegradesToBlankExtentWithoutCrash()
    {
        // Scatter has no cached-values projection; the drawing must degrade to the established
        // no-chart output (blank extent) rather than fail the whole conversion.
        var c = ChartC;
        string html = HtmlConversionOps.ConvertToHtml(GeneratedChartDocxBytes(
                new XElement(c + "scatterChart",
                    new XElement(c + "scatterStyle", new XAttribute("val", "lineMarker")),
                    GeneratedChartSeries(0, "North", new[] { 2.0, 4.0, 3.0 }))),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.DoesNotContain(XElement.Parse(html).Descendants(),
            element => element.Name.LocalName == "svg");
    }

    [Fact]
    public void HCO020_BulletListMarker_RendersUnicodeBullet()
    {
        // A bullet list item carries the Symbol-font glyph U+F0B7, which renders as a blank box in a
        // browser without the proprietary font installed. The converter should map list-marker
        // symbol glyphs to their Unicode equivalents (U+F0B7 -> U+2022 "•").
        var bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "Blank-wml.docx"));
        using var session = new DocxSession(bytes);
        var anchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind is "p" or "h" or "li").Anchor.Id;

        var edit = session.ReplaceText(anchor, "First bullet item");
        Assert.True(edit.Success, edit.Error?.Message);
        var li = session.ApplyListFormat(edit.Modified[0].Id, ListFormat.Bullet);
        Assert.True(li.Success, li.Error?.Message);

        string html = HtmlConversionOps.ConvertToHtml(session.Save(), new HtmlConversionOptions());

        Assert.Contains("•", html);       // • rendered for the bullet marker
        Assert.DoesNotContain("", html); // the raw Symbol private-use glyph is gone
    }

    [Fact]
    public void HCO002_ConvertSession_ReflectsEdit()
    {
        using var session = new DocxSession(TourPlanBytes());
        var projection = session.Project();

        // First body paragraph/heading/list-item anchor, in document order.
        // C# AnchorTarget nests the anchor: record struct Anchor(Id, Kind, Scope, Unid).
        string FirstAnchor()
        {
            string? best = null;
            int bestPos = int.MaxValue;
            foreach (var target in projection.AnchorIndex.Values)
            {
                if (target.Anchor.Scope != "body") continue;
                if (target.Anchor.Kind is not ("p" or "h" or "li")) continue;
                int pos = projection.Markdown.IndexOf("{#" + target.Anchor.Id + "}", System.StringComparison.Ordinal);
                if (pos >= 0 && pos < bestPos) { bestPos = pos; best = target.Anchor.Id; }
            }
            Assert.NotNull(best);
            return best!;
        }

        var edit = session.ReplaceText(FirstAnchor(), "HCO002UNIQUEMARKER edited body.");
        Assert.True(edit.Success, edit.Error?.Message);

        string html = HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions());

        Assert.Contains("HCO002UNIQUEMARKER", html);
    }

    // THE FEASIBILITY GATE (spec docs/architecture/ir_editor_feasibility.md §5/§6.1):
    // The full-document render is ground truth. RenderBlockHtml(anchor) is "faithful"
    // iff its output matches the data-anchor-stamped element from the full render —
    // same tag and same visible text. Proves single-block render out of whole-doc
    // context. (List-continuation + inline-image blocks are known PoC limits, skipped.)
    [Theory]
    [InlineData("HC006-Test-01.docx")]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void HCO050_RenderBlockHtml_MatchesFullRenderPerAnchor(string fileName)
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", fileName));

        // Full render = oracle; StampAnchors assigns the same deterministic Unids.
        var full = System.Xml.Linq.XElement.Parse(
            HtmlConversionOps.ConvertToHtml(bytes,
                new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false }));

        var fullByAnchor = full.Descendants()
            .Where(e => (string?)e.Attribute("data-anchor") != null)
            .GroupBy(e => (string)e.Attribute("data-anchor")!)
            .ToDictionary(g => g.Key, g => g.First());

        // Stamping must work at all (this is the editor's actual render path).
        Assert.NotEmpty(fullByAnchor);

        static string Norm(string s) =>
            System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
        static bool HasImg(System.Xml.Linq.XElement e) =>
            e.Descendants().Any(d => d.Name.LocalName == "img");

        var targets = fullByAnchor
            .Where(kv => (kv.Value.Name.LocalName is "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                         && !HasImg(kv.Value) && Norm(kv.Value.Value).Length > 0)
            .Take(12).ToList();
        Assert.NotEmpty(targets);

        int verified = 0;
        foreach (var kv in targets)
        {
            // data-anchor carries the bare unid; RenderBlockHtml accepts a bare unid
            // OR a full kind:scope:unid (it keys on the unid tail). This is exactly
            // what the editor passes back from a DOM block's data-anchor.
            string html = HtmlConversionOps.RenderBlockHtml(bytes, kv.Key,
                new HtmlConversionOptions { FabricateCssClasses = false });
            var blockEl = System.Xml.Linq.XElement.Parse(html);
            Assert.Equal(kv.Value.Name.LocalName, blockEl.Name.LocalName);
            Assert.Equal(Norm(kv.Value.Value), Norm(blockEl.Value));
            verified++;
        }

        Assert.True(verified > 0, "no blocks verified");
    }

    // Proves (a) the session-attached render resolves the SAME anchors the full render
    // stamps (one Unid scheme across convertDocxToHtml ↔ DocxSession ↔ RenderBlock) and
    // produces equivalent output, and (b) it avoids the per-call byte re-open + whole-doc
    // Unid pass, so it is no slower than the stateless path. Logs per-block latency.
    [Fact]
    public void HCO052_SessionAttachedRender_EquivalentAndNotSlower()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        var opts = new HtmlConversionOptions { FabricateCssClasses = false };

        var full = System.Xml.Linq.XElement.Parse(
            HtmlConversionOps.ConvertToHtml(bytes,
                new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false }));
        var anchors = full.Descendants()
            .Where(e => (e.Name.LocalName is "p" or "h1" or "h2" or "h3" or "h4")
                        && (string?)e.Attribute("data-anchor") != null
                        && e.Descendants().All(d => d.Name.LocalName != "img"))
            .Select(e => (string)e.Attribute("data-anchor")!)
            .Where(u => u.Length == 32)
            .Distinct().Take(20).ToList();
        Assert.NotEmpty(anchors);

        static string Text(string html) => System.Text.RegularExpressions.Regex.Replace(
            System.Xml.Linq.XElement.Parse(html).Value, "\\s+", " ").Trim();

        using var session = new DocxSession(bytes);

        // (a) Equivalence: session-attached resolves the full-render anchor (same scheme)
        // and yields the same text as the stateless path. This is the editor's invariant:
        // a DOM block's data-anchor is a valid DocxSession/RenderBlock anchor.
        foreach (var a in anchors.Take(6))
        {
            string viaBytes = HtmlConversionOps.RenderBlockHtml(bytes, a, opts);
            string viaSession = HtmlConversionOps.RenderBlockHtml(session, a, opts);
            Assert.Equal(Text(viaBytes), Text(viaSession));
        }

        // Warmup (JIT + first projection on the session path).
        HtmlConversionOps.RenderBlockHtml(bytes, anchors[0], opts);
        HtmlConversionOps.RenderBlockHtml(session, anchors[0], opts);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        foreach (var a in anchors) HtmlConversionOps.RenderBlockHtml(bytes, a, opts);
        double statelessMs = sw.Elapsed.TotalMilliseconds / anchors.Count;

        sw.Restart();
        foreach (var a in anchors) HtmlConversionOps.RenderBlockHtml(session, a, opts);
        double sessionMs = sw.Elapsed.TotalMilliseconds / anchors.Count;

        _output.WriteLine($"PROFILE HC031 n={anchors.Count}: stateless={statelessMs:F2}ms/block " +
                          $"session-attached={sessionMs:F2}ms/block speedup={statelessMs / sessionMs:F2}x");

        // Session-attached must not be materially slower (it skips re-open + whole-doc
        // Unid assignment). Generous margin keeps the assertion robust to CI noise.
        Assert.True(sessionMs <= statelessMs * 1.25,
            $"session-attached slower than stateless: stateless={statelessMs:F2} session={sessionMs:F2}");
    }

    // The single-block render path sets SkipFormattingPartsSimplification=true to avoid re-walking
    // the (potentially huge) style gallery on every keystroke commit. That pass only strips
    // rendering-irrelevant rsids from the style parts, so it MUST be byte-for-byte rendering-neutral.
    // Prove it directly: a full-document convert with the flag on vs off produces identical HTML
    // (covers CSS classes + theme fonts + list markers, not just tag+text like HCO050).
    [Theory]
    [InlineData("HC031-Complicated-Document.docx", false)]
    [InlineData("HC001-5DayTourPlanTemplate.docx", false)]
    [InlineData("HC031-Complicated-Document.docx", true)]
    public void HCO053_SkipFormattingPartsSimplification_IsRenderingNeutral(string fileName, bool paginated)
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", fileName));
        string Render(bool skip)
        {
            using var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            ms.Position = 0;
            using var doc = WordprocessingDocument.Open(ms, true);
            var settings = new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                StampAnchors = true,
                RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None,
                SkipFormattingPartsSimplification = skip,
            };
            return WmlToHtmlConverter.ConvertToHtml(doc, settings)
                .ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }

        Assert.Equal(Render(false), Render(true));
    }

    // The session-attached render path reuses a cached formatting "shell" across calls. Prove it is
    // (a) consistent across calls (cache reuse doesn't drift) and (b) byte-identical to the stateless
    // path (which HCO050 already ties to the full-render oracle).
    [Fact]
    public void HCO054_SessionShellRender_ConsistentAndMatchesStateless()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var opts = new HtmlConversionOptions { FabricateCssClasses = false, CssClassPrefix = "pt-" };
        var anchors = session.Project().AnchorIndex.Keys
            .Where(k => k.StartsWith("p:") || k.StartsWith("h:") || k.StartsWith("li:"))
            .Take(12).ToList();
        Assert.NotEmpty(anchors);

        int verified = 0;
        foreach (var a in anchors)
        {
            string first = HtmlConversionOps.RenderBlockHtml(session, a, opts);   // builds the shell
            string second = HtmlConversionOps.RenderBlockHtml(session, a, opts);  // reuses the shell
            Assert.Equal(first, second);
            string stateless = HtmlConversionOps.RenderBlockHtml(bytes, a, opts); // independent path
            Assert.Equal(stateless, first);
            verified++;
        }
        Assert.True(verified > 0);
    }

    // A mid-session format op (ApplyListFormat) mutates the numbering part, so the cached shell MUST
    // be rebuilt (signature change) — otherwise the freshly-list-ified paragraph would render WITHOUT
    // its marker against a stale (numbering-less) shell. Also covers the no-list -> list transition.
    [Fact]
    public void HCO055_SessionShellRender_RebuildsAfterFormattingMutation()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var opts = new HtmlConversionOptions { FabricateCssClasses = false, CssClassPrefix = "pt-" };

        var plain = session.Project().AnchorIndex
            .First(kv => kv.Key.StartsWith("p:") && kv.Value.TextPreview.Trim().Length > 3);

        // Prime the shell (no marker yet).
        string before = HtmlConversionOps.RenderBlockHtml(session, plain.Key, opts);
        Assert.DoesNotContain("data-list-marker", before);

        // Mutate the numbering part; the next render must rebuild the shell and show the marker.
        var r = session.ApplyListFormat(plain.Key, ListFormat.Bullet);
        Assert.True(r.Success, r.Error?.Message);
        string after = HtmlConversionOps.RenderBlockHtml(session, r.Modified[0].Id, opts);
        Assert.Contains("data-list-marker", after);
    }

    // A borderless layout table (w:tblBorders all w:val="none", with NO w:sz) — the standard way real
    // S-1 covers lay out multi-column rows — used to CRASH the whole conversion: both
    // FormattingAssembler.ResolveInsideBorder and WmlToHtmlConverter.ResolveCellBorder cast the
    // absent w:sz to a value type (only "nil" was special-cased; "none" fell through). It must render.
    [Fact]
    public void HCO056_BorderlessTable_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var noneBorders = new Wp.TableBorders(
                new Wp.TopBorder { Val = Wp.BorderValues.None },
                new Wp.LeftBorder { Val = Wp.BorderValues.None },
                new Wp.BottomBorder { Val = Wp.BorderValues.None },
                new Wp.RightBorder { Val = Wp.BorderValues.None },
                new Wp.InsideHorizontalBorder { Val = Wp.BorderValues.None },
                new Wp.InsideVerticalBorder { Val = Wp.BorderValues.None });
            main.Document.Body!.Append(new Wp.Table(
                new Wp.TableProperties(noneBorders),
                new Wp.TableRow(
                    new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("LeftCellText")))),
                    new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("RightCellText")))))));
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());
        Assert.Contains("LeftCellText", html);
        Assert.Contains("RightCellText", html);
    }

    // Minimal OOXML packages (document.xml + styles.xml only — no word/settings.xml) are legal:
    // ECMA-376 does not require DocumentSettingsPart, and Word opens them without repair.
    // CalculateSpanWidthForTabs used to call DocumentSettingsPart.GetXDocument() unconditionally,
    // which threw ArgumentNullException("part") and aborted conversion. Default tab stop is 720 twips.
    [Fact]
    public void HCO057_MissingDocumentSettingsPart_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.Text("Hello no-settings package")))));
            // Styles are required by FormattingAssembler; settings intentionally omitted.
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(
                            new Wp.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                            new Wp.FontSize { Val = "24" }))));
            main.Document.Save();
        }

        // Prove the part is absent (not just that we forgot to assert the repro shape).
        using (var reopen = WordprocessingDocument.Open(ms, false))
        {
            Assert.Null(reopen.MainDocumentPart!.DocumentSettingsPart);
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());
        Assert.Contains("Hello no-settings package", html);
    }

    // CalculateSpanWidthForTabs (WmlToHtmlConverter.cs) computes a tab's rendered width from
    // w:defaultTabStop. This pins the actual numeric fallback (720 twips == 0.5in) that
    // HCO057 only proved didn't crash — i.e. the missing-settings path doesn't just avoid
    // throwing, it produces the SAME width Word itself defaults to for an unset tab stop.
    [Fact]
    public void HCO058_MissingDocumentSettingsPart_TabWidthDefaultsTo720Twips()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text("AfterTab")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            // DocumentSettingsPart intentionally omitted.
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("AfterTab", html);
        // 720 twips (Word's implicit default tab stop) == 0.5in from position 0.
        Assert.Contains("width: 0.50in", html);
    }

    // Same computation, but with an explicit DocumentSettingsPart that overrides
    // w:defaultTabStop — proves the "settingsPart != null" branch introduced by the same
    // refactor still reads the configured value correctly (not just the null-guard path).
    [Fact]
    public void HCO059_DocumentSettingsPartWithCustomDefaultTabStop_TabWidthUsesConfiguredValue()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text("AfterTab")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings =
                new Wp.Settings(new Wp.DefaultTabStop { Val = 1440 }); // 1 inch
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("AfterTab", html);
        Assert.Contains("width: 1.00in", html);
        Assert.DoesNotContain("width: 0.50in", html);
    }

    // DocumentSettingsPart present but with no w:defaultTabStop element at all (legal — the
    // element is optional within w:settings). Must fall back to the same 720-twip default as
    // when the whole part is absent, not throw and not silently use 0.
    [Fact]
    public void HCO060_DocumentSettingsPartWithoutDefaultTabStopElement_FallsBackTo720Twips()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text("AfterTab")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings(); // no DefaultTabStop child
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("AfterTab", html);
        Assert.Contains("width: 0.50in", html);
    }

    // AddFormattingParts copies formatting parts into the RenderBlockHtml throwaway doc but no
    // longer invents a dummy DocumentSettingsPart. Regression: a source with no settings part
    // must still round-trip through RenderBlockHtml without crashing (converter defaults tab stop).
    [Fact]
    public void HCO061_RenderBlockHtml_SourceMissingDocumentSettingsPart_DoesNotCrash()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.Text("HCO061 block text")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            // DocumentSettingsPart intentionally omitted from the source document.
            main.Document.Save();
        }
        byte[] bytes = ms.ToArray();

        using (var reopenStream = new MemoryStream(bytes))
        using (var reopen = WordprocessingDocument.Open(reopenStream, false))
        {
            Assert.Null(reopen.MainDocumentPart!.DocumentSettingsPart);
        }

        var opts = new HtmlConversionOptions { FabricateCssClasses = false };
        string full = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false });
        var anchorEl = System.Xml.Linq.XElement.Parse(full).Descendants()
            .First(e => (string?)e.Attribute("data-anchor") != null);
        string anchorId = (string)anchorEl.Attribute("data-anchor")!;

        string block = HtmlConversionOps.RenderBlockHtml(bytes, anchorId, opts);

        Assert.Contains("HCO061 block text", block);
    }

    // Builds a document-only package (word/document.xml only — no styles, no settings) and
    // proves the repro shape: StyleDefinitionsPart really is absent after reopen.
    private static byte[] DocumentOnlyDocxBytes(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.Text(text)))));
            main.Document.Save();
        }

        using (var reopen = WordprocessingDocument.Open(ms, false))
        {
            Assert.Null(reopen.MainDocumentPart!.StyleDefinitionsPart);
        }
        return ms.ToArray();
    }

    private static byte[] PageSizedDocxBytes(uint width, uint height)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(new Wp.Run(new Wp.Text("Page-sized test content"))),
                    new Wp.SectionProperties(
                        new Wp.PageSize { Width = width, Height = height },
                        new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] MixedPageSizedDocxBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.ParagraphProperties(
                            new Wp.SectionProperties(
                                new Wp.PageSize { Width = 11906, Height = 16838 },
                                new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })),
                        new Wp.Run(new Wp.Text("A4 section"))),
                    new Wp.Paragraph(new Wp.Run(new Wp.Text("Landscape section"))),
                    new Wp.SectionProperties(
                        new Wp.PageSize { Width = 15840, Height = 12240 },
                        new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    // Issue #265 — sibling of the missing-settings crash fixed in #264. word/styles.xml
    // (StyleDefinitionsPart) is also optional in OOXML: Word opens a document-only package
    // without repair, but FormattingAssembler.AssembleFormatting dereferenced
    // StyleDefinitionsPart unconditionally (many sites), throwing ArgumentNullException("part")
    // at WmlToHtmlConverter.cs's AssembleFormatting call — before any HTML was produced.
    [Fact]
    public void HCO062_MissingStyleDefinitionsPart_DoesNotCrashConverter()
    {
        byte[] bytes = DocumentOnlyDocxBytes("Hello no-styles package");

        string html = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions());

        Assert.Contains("Hello no-styles package", html);
    }

    // Same crash, real-world shape: the RPR fixtures contain ONLY [Content_Types].xml,
    // _rels/.rels, and word/document.xml (no styles, no settings) and crashed on conversion.
    [Fact]
    public void HCO063_DocumentOnlyPackage_RprFixture_ConvertsToHtml()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "RPR-FivePageTestDoc.docx"));

        string html = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions());

        Assert.Contains("Page 1 paragraph 1", html);
        Assert.Contains("Page 5 paragraph 1", html);
    }

    // RenderBlockHtml's throwaway doc copies the source's formatting parts; with no styles
    // part to copy, the single-block path must survive a styles-less source end to end.
    [Fact]
    public void HCO064_RenderBlockHtml_SourceMissingStyleDefinitionsPart_DoesNotCrash()
    {
        byte[] bytes = DocumentOnlyDocxBytes("HCO064 block text");

        var opts = new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false };
        string full = HtmlConversionOps.ConvertToHtml(bytes, opts);
        var anchorEl = System.Xml.Linq.XElement.Parse(full).Descendants()
            .First(e => (string?)e.Attribute("data-anchor") != null);
        string anchorId = (string)anchorEl.Attribute("data-anchor")!;

        string block = HtmlConversionOps.RenderBlockHtml(bytes, anchorId, opts);

        Assert.Contains("HCO064 block text", block);
    }

    // Some producer packages use a VML <v:imagedata> relationship id for a non-image part. The
    // unsupported VML image is safely omitted, but it must not cast CustomXmlPart to ImagePart and
    // take down the document conversion (the complex_style_attr benchmark fixtures have this shape).
    [Fact]
    public void HCO065_VmlImageDataReferencingCustomXmlPart_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var customXml = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using (var customWriter = new StreamWriter(customXml.GetStream(FileMode.Create, FileAccess.Write)))
                customWriter.Write("<payload/>");
            string relationshipId = main.GetIdOfPart(customXml);

            using (var documentWriter = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                documentWriter.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:v=\"urn:schemas-microsoft-com:vml\"><w:body><w:p><w:r><w:t>" +
                    "HCO065 retained text</w:t></w:r><w:r><w:pict><v:shape style=\"width:10pt;height:10pt\">" +
                    $"<v:imagedata r:id=\"{relationshipId}\"/></v:shape></w:pict></w:r></w:p>" +
                    "<w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("HCO065 retained text", html);
    }

    // Word tolerates malformed pPr payloads where lineRule is present but the required line value
    // is absent (including documents with duplicate pPr elements). Treat it as the implicit browser
    // line-height instead of casting the missing attribute and aborting the complete conversion.
    [Fact]
    public void HCO066_AutoLineRuleWithoutLineValue_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var documentWriter = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                documentWriter.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:spacing w:lineRule=\"auto\"/></w:pPr><w:pPr/>" +
                    "<w:r><w:t>HCO066 retained text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("HCO066 retained text", html);
        // No line spacing is DERIVED for the paragraph: neither the automatic-spacing multiplier
        // nor a computed height. The document-layout stylesheet's own `sup, sub { line-height: 0 }`
        // is always present and is not a paragraph declaration, so assert on what the paragraph
        // produced rather than on the whole document containing the string at all.
        Assert.DoesNotContain("--docx-auto-line-spacing", html);
        Assert.DoesNotContain("calc(1lh", html);
        Assert.DoesNotContain("line-height: 1", html);
    }

    // Strict/compatibility producers can express paragraph spacing as fractional point measures
    // rather than raw twips. It must use the same measure parser for before and line spacing so
    // one style default cannot abort every paragraph in the document.
    [Fact]
    public void HCO073_PointSuffixedAutoLineSpacing_NormalizesToTwipsBeforeDerivingCss()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:spacing w:before=\"2pt\" w:line=\"12.95pt\" w:lineRule=\"auto\"/>" +
                    "</w:pPr><w:r><w:t>HCO073 retained text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO073 retained text", html);
        // 12.95pt is 259 twips, so the derived multiple is 259/240. What the test pins is the
        // NORMALIZATION (a point-suffixed w:line is parsed at all); the CSS shape is the auto
        // line-spacing model, which expresses the multiple against the font's own line box
        // rather than as a percentage of its em square (issues #396/#397).
        Assert.Contains("--docx-auto-line-spacing: 1.079", html);
        Assert.Contains("line-height: normal", html);
    }

    // Table indentation and preceding paragraph spacing can use point measures too. A table-cell
    // fill with no explicit shading pattern is likewise a common Word-compatible clear shading
    // form. Normalize both shapes without throwing while probing the shade mapper.
    [Fact]
    public void HCO074_CellFillWithoutShadeValue_RendersClearFill()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:spacing w:after=\"8pt\"/></w:pPr><w:r><w:t>HCO074 preceding text</w:t></w:r></w:p>" +
                    "<w:tbl><w:tblPr><w:tblInd w:w=\"0pt\" w:type=\"dxa\"/></w:tblPr>" +
                    "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc>" +
                    "<w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/><w:shd w:fill=\"D9EAF7\"/></w:tcPr>" +
                    "<w:p><w:r><w:t>HCO074 retained text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                    "<w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO074 retained text", html);
        Assert.Contains("background: #D9EAF7", html);
    }

    // The viewer's byte-based HTML bridge must open Strict OOXML packages just as DocxDiff does.
    // Exercise both full-document and anchor-addressed block rendering; before normalization the
    // converter sees no transitional w:body and throws on these packages.
    [Fact]
    public void HCO067_StrictOoxml_NormalizesBeforeFullAndBlockRender()
    {
        byte[] strict = StrictDocumentOnlyDocxBytes("HCO067 strict retained text");
        var options = new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false };

        string full = HtmlConversionOps.ConvertToHtml(strict, options);
        Assert.Contains("HCO067 strict retained text", full);
        var anchor = System.Xml.Linq.XElement.Parse(full).Descendants()
            .First(e => (string?)e.Attribute("data-anchor") != null)
            .Attribute("data-anchor")!.Value;

        string block = HtmlConversionOps.RenderBlockHtml(strict, anchor, options);
        Assert.Contains("HCO067 strict retained text", block);
    }

    // Word writes one text box twice inside mc:AlternateContent: a modern DrawingML/wps branch
    // and a VML fallback. The renderer must select the supported modern branch exactly once,
    // retain the visible text, and not double the logical box in HTML.
    [Fact]
    public void HCO068_ModernDrawingMlTextBox_RendersChoiceWithoutVmlDuplicate()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<mc:AlternateContent>" +
            "<mc:Choice Requires=\"wps\"><w:drawing><wp:inline><wp:extent cx=\"1524000\" cy=\"762000\"/>" +
            "<a:graphic><a:graphicData><wps:wsp><wps:spPr><a:solidFill><a:srgbClr val=\"FFFFFF\"/>" +
            "</a:solidFill><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/>" +
            "</a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>" +
            "HCO068 modern text box</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
            "<wps:bodyPr lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"><a:spAutoFit/>" +
            "</wps:bodyPr>" +
            "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape style=\"width:120pt;height:60pt\"><v:textbox>" +
            "<w:txbxContent><w:p><w:r><w:t>HCO068 fallback text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
            "</mc:AlternateContent>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO068 modern text box", html);
        Assert.DoesNotContain("HCO068 fallback text box", html);
        Assert.Contains("width: 120pt", html);
        Assert.DoesNotContain("height: 60pt", html);
        Assert.Contains("margin-bottom: 0", html);
        Assert.DoesNotContain("data-docx-drawing-anchor", html);
        Assert.DoesNotContain("position: absolute", html);
    }

    // Floating drawings cannot be positioned until pagination supplies the concrete page and
    // anchor-paragraph boxes. Preserve each OOXML input independently: offsets locate the object,
    // wrap distances describe clearance from surrounding text, bodyPr insets locate text inside it,
    // the stored extent remains the fixed-size fallback when no relative size overrides it.
    [Fact]
    public void HCO091_AnchoredDrawingMlTextBox_PreservesGeometryForPagination()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<w:drawing><wp:anchor distT=\"45720\" distR=\"114300\" distB=\"45720\" distL=\"114300\" " +
            "simplePos=\"0\" relativeHeight=\"12\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">" +
            "<wp:simplePos x=\"0\" y=\"0\"/>" +
            "<wp:positionH relativeFrom=\"page\"><wp:posOffset>457200</wp:posOffset></wp:positionH>" +
            "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>228600</wp:posOffset></wp:positionV>" +
            "<wp:extent cx=\"1524000\" cy=\"762000\"/><wp:wrapSquare wrapText=\"bothSides\"/>" +
            "<wp:docPr id=\"1\" name=\"Generated anchor\"/><wp:cNvGraphicFramePr/>" +
            "<a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>" +
            "<w:p><w:r><w:t>HCO091 anchored text box</w:t></w:r></w:p>" +
            "</w:txbxContent></wps:txbx><wps:bodyPr lIns=\"91440\" tIns=\"45720\" " +
            "rIns=\"91440\" bIns=\"45720\"/></wps:wsp></a:graphicData></a:graphic>" +
            "<wp14:sizeRelH relativeFrom=\"margin\"><wp14:pctWidth>40000</wp14:pctWidth></wp14:sizeRelH>" +
            "</wp:anchor></w:drawing>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions
            {
                FabricateCssClasses = false,
                PaginationMode = (int)PaginationMode.Paginated,
            });

        Assert.Contains("data-docx-drawing-anchor=\"true\"", html);
        Assert.Contains("data-docx-anchor-extent-width=\"120\"", html);
        Assert.Contains("data-docx-anchor-extent-height=\"60\"", html);
        Assert.Contains("data-docx-anchor-h-relative=\"page\"", html);
        Assert.Contains("data-docx-anchor-h-offset=\"36\"", html);
        Assert.Contains("data-docx-anchor-v-relative=\"paragraph\"", html);
        Assert.Contains("data-docx-anchor-v-offset=\"18\"", html);
        Assert.Contains("data-docx-anchor-wrap-top=\"3.6\"", html);
        Assert.Contains("data-docx-anchor-wrap-right=\"9\"", html);
        Assert.Contains("data-docx-anchor-width-relative=\"margin\"", html);
        Assert.Contains("data-docx-anchor-width-percent=\"40\"", html);
        Assert.Contains("width: 120pt", html);
        Assert.Contains("height: 60pt", html);
        Assert.Contains("padding-left: 7.2pt", html);
        Assert.Contains("padding-top: 3.6pt", html);
        Assert.Contains("position: absolute", html);
    }

    // A floating picture anchored with wp:wrapSquare must exclude body text from its rect
    // (issue #412): Word wraps the paragraph's lines beside the picture, while an inline
    // <img> pushes every following line below it. The picture in DB007 is offset-placed
    // against the column with its center past the column midpoint, so it floats right; the
    // anchor's 114300 EMU distL/distR become 9pt clearance margins.
    [Fact]
    public void HCO092_AnchoredPictureWithSquareWrap_FloatsRight()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "DB007-WhitePaper.docx"));

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("float: right", html);
        Assert.Contains("margin-left: 9pt", html);
        Assert.Contains("margin-right: 9pt", html);
    }

    // wrapTight's polygon degrades to its bounding box. This picture sits at column offset 0,
    // so its center is left of the column midpoint and it floats left.
    [Fact]
    public void HCO093_AnchoredPictureWithTightWrap_FloatsLeft()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "VP", "VP002-Image-Wrap-Tight.docx"));

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("float: left", html);
        Assert.DoesNotContain("float: right", html);
    }

    // Some legacy Word documents keep the modern DrawingML text-box body in a related XML
    // part. The synthetic package deliberately uses a distinct VML fallback so this verifies
    // that the supported choice gains its external body without rendering both copies.
    [Fact]
    public void HCO075_ExternalDrawingMlTextBox_RendersChoiceWithoutVmlDuplicate()
    {
        byte[] bytes = ExternalTextBoxDocxBytes();

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO075 external textbox text", html);
        Assert.DoesNotContain("HCO075 fallback text box", html);
        Assert.Contains("width: 120pt", html);
    }

    // Old Office 2008 wps markup is not a namespace this renderer understands. Markup
    // Compatibility requires selecting its portable VML fallback rather than dropping the
    // entire logical text box.
    [Fact]
    public void HCO069_LegacyDrawingMlTextBox_RendersVmlFallback()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<mc:AlternateContent>" +
            "<mc:Choice Requires=\"legacywps\"><w:drawing><wp:inline><wp:extent cx=\"1524000\" cy=\"762000\"/>" +
            "<a:graphic><a:graphicData><legacywps:wsp/></a:graphicData></a:graphic>" +
            "</wp:inline></w:drawing></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox>" +
            "<w:txbxContent><w:p><w:r><w:t>HCO069 legacy fallback text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
            "</mc:AlternateContent>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO069 legacy fallback text box", html);
        Assert.Contains("width: 100pt", html);
        Assert.Contains("height: 40pt", html);
    }

    // A direct VML text box is not an AlternateContent compatibility fallback and remains a
    // supported, standalone shape. Preserve its content and size in the HTML projection.
    [Fact]
    public void HCO070_DirectVmlTextBox_RendersTextAndDimensions()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>HCO070 direct VML text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO070 direct VML text box", html);
        Assert.Contains("width: 100pt", html);
        Assert.Contains("height: 40pt", html);
    }

    // VML theme colour values can append Word palette metadata. Keep the colour itself, but do
    // not feed the suffix (or arbitrary CSS declarations) through to the generated style string.
    [Fact]
    public void HCO071_VmlThemeColors_NormalizeAndRejectStyleInjection()
    {
        byte[] themed = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\" fillcolor=\"#156082 [3204]\" " +
            "strokecolor=\"white [3212]\"><v:textbox><w:txbxContent><w:p><w:r><w:t>" +
            "HCO071 themed VML text box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>" +
            "</w:pict>");
        byte[] unsafeColor = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\" fillcolor=\"red; color: blue\"><v:textbox>" +
            "<w:txbxContent><w:p><w:r><w:t>HCO071 unsafe VML text box</w:t></w:r></w:p></w:txbxContent>" +
            "</v:textbox></v:shape></w:pict>");

        string themedHtml = HtmlConversionOps.ConvertToHtml(themed,
            new HtmlConversionOptions { FabricateCssClasses = false });
        string unsafeHtml = HtmlConversionOps.ConvertToHtml(unsafeColor,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("background-color: #156082", themedHtml);
        Assert.Contains("border: 1pt solid white", themedHtml);
        Assert.DoesNotContain("3204", themedHtml);
        Assert.DoesNotContain("color: blue", unsafeHtml);
    }

    [Fact]
    public void HCO072_DirectAutoFitVmlTextBox_DropsStoredHeightAndTrailingSpacing()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox style=\"mso-fit-shape-to-text:t\">" +
            "<w:txbxContent><w:p><w:r><w:t>HCO072 auto-fit VML text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO072 auto-fit VML text box", html);
        Assert.DoesNotContain("height: 40pt", html);
        Assert.Contains("margin-bottom: 0", html);
    }

    // The clean-view fast path must recognize every revision family the accepter handles. A cell
    // deletion has no w:ins/w:del wrapper, so the former body-only detector skipped acceptance and
    // leaked its text into the supposedly accepted HTML.
    [Fact]
    public void HCO073_CleanView_AcceptsCellDeletionRevision()
    {
        string html = HtmlConversionOps.ConvertToHtml(CellDeletionTableDocxBytes(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO073 retained cell", html);
        Assert.DoesNotContain("HCO073 deleted cell", html);
    }

    private static byte[] CellDeletionTableDocxBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/>" +
                    "</w:tblGrid><w:tr>" +
                    "<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>" +
                    "HCO073 retained cell</w:t></w:r></w:p></w:tc>" +
                    "<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/><w:cellDel w:id=\"1\" " +
                    "w:author=\"Test\" w:date=\"2026-01-01T00:00:00Z\"/></w:tcPr><w:p><w:r><w:t>" +
                    "HCO073 deleted cell</w:t></w:r></w:p></w:tc>" +
                    "</w:tr></w:tbl><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }
        return ms.ToArray();
    }

    private static byte[] TextBoxDocxBytes(string runContent)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                    "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
                    "xmlns:legacywps=\"http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingShape\">" +
                    "<w:body><w:p><w:r>" + runContent + "</w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }
        return ms.ToArray();
    }

    private static byte[] ExternalTextBoxDocxBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var externalTextBox = main.AddExtendedPart(
                "http://schemas.microsoft.com/office/2006/relationships/txbx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.txbx+xml",
                ".xml",
                "rIdExternal");
            using (var writer = new StreamWriter(externalTextBox.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w14:txbx xmlns:w14=\"http://schemas.microsoft.com/office/word/2008/9/12/wordml\" " +
                    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:p><w:r><w:t>HCO075 external textbox text</w:t></w:r></w:p></w14:txbx>");
            }

            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                    "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
                    "<w:body><w:p><w:r><mc:AlternateContent>" +
                    "<mc:Choice Requires=\"wps\"><w:drawing><wp:inline><wp:extent cx=\"1524000\" cy=\"762000\"/>" +
                    "<a:graphic><a:graphicData><wps:wsp><wps:txbx r:txbx=\"rIdExternal\"/>" +
                    "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
                    "<mc:Fallback><w:pict><v:shape style=\"width:120pt;height:60pt\"><v:textbox>" +
                    "<w:txbxContent><w:p><w:r><w:t>HCO075 fallback text box</w:t></w:r></w:p>" +
                    "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
                    "</mc:AlternateContent></w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }
        return ms.ToArray();
    }

    // Unids are CONTENT-ADDRESSED, so identical content in DIFFERENT parts shares one unid.
    // HC031 carries default/first/even footer stories, all empty, which collide. Resolving a
    // block by bare unid then renders whichever part the scan reaches first, so an editor's
    // header/footer band would DISPLAY one story while editing another. The session-attached
    // path must resolve through the anchor index, which knows the owning part.
    [Fact]
    public void HCO080_RenderBlockHtml_ResolvesTheAnchorsOwnPart_WhenUnidsCollide()
    {
        var bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var body = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

        var refs = session.GetSectionInfo(body)!.FooterRefs;
        var markers = new System.Collections.Generic.Dictionary<HeaderFooterKind, string>
        {
            [HeaderFooterKind.Default] = "DDD-DEFAULT",
            [HeaderFooterKind.First] = "FFF-FIRST",
            [HeaderFooterKind.Even] = "EEE-EVEN",
        };

        // Precondition: the stories really do collide, or this test covers nothing.
        var unids = refs
            .Select(r => session.Project().AnchorIndex.Values.First(t => t.PartUri == r.PartUri).Unid)
            .ToList();
        Assert.True(unids.Distinct().Count() < unids.Count,
            "precondition: this fixture's empty footer stories should share a content-addressed unid");

        var anchorOf = refs.ToDictionary(
            r => r.Kind,
            r => session.Project().AnchorIndex.Values
                     .First(t => t.PartUri == r.PartUri && t.Anchor.Kind == "p").Anchor.Id);

        foreach (var (kind, text) in markers)
            Assert.True(session.ReplaceText(anchorOf[kind], text).Success);

        var opts = new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false };
        foreach (var (kind, text) in markers)
        {
            var html = HtmlConversionOps.RenderBlockHtml(session, anchorOf[kind], opts);
            Assert.Contains(text, html);
            foreach (var (otherKind, otherText) in markers)
                if (otherKind != kind) Assert.DoesNotContain(otherText, html);
        }
    }

    // THE BATCH GATE: RenderBlocksHtml output must be ELEMENT-IDENTICAL to the
    // corresponding data-anchor element of a full render — including list-item
    // markers deep in a list (numbering continuation, the M9 gap the single-block
    // path had) and contextualSpacing-dependent margins (neighbor context). This is
    // deliberately stronger than HCO050's tag+text check.
    [Fact]
    public void HCO081_RenderBlocksHtml_MatchesFullRenderFragments()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true };

        var full = System.Xml.Linq.XElement.Parse(HtmlConversionOps.ConvertToHtml(session, options));
        var fullByAnchor = full.Descendants()
            .Where(e => (string?)e.Attribute("data-anchor") != null)
            .GroupBy(e => (string)e.Attribute("data-anchor")!)
            .ToDictionary(g => g.Key, g => g.First());

        static bool HasImg(System.Xml.Linq.XElement e) =>
            e.Descendants().Any(d => d.Name.LocalName == "img");

        var plan = session.ListBlocks();
        var ids = plan.Body.Where(u => u.Kind == "li").Take(10)
            .Concat(plan.Body.Where(u => u.Kind == "p").Take(4))
            .Concat(plan.Body.Where(u => u.Kind == "h").Take(2))
            .Where(u => fullByAnchor.TryGetValue(u.Id.Substring(u.Id.LastIndexOf(':') + 1), out var el) && !HasImg(el))
            .Select(u => u.Id)
            .ToList();
        Assert.True(ids.Count >= 8, $"fixture too thin: only {ids.Count} usable units");

        var json = HtmlConversionOps.RenderBlocksHtml(session, ids, options);
        using var map = System.Text.Json.JsonDocument.Parse(json);

        foreach (var id in ids)
        {
            var unid = id.Substring(id.LastIndexOf(':') + 1);
            var expected = fullByAnchor[unid].ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            var actual = map.RootElement.GetProperty(id).GetString();
            Assert.NotNull(actual);
            // One extra Parse round-trip on the actual normalizes serializer escaping
            // (&#x00a0; vs the raw NBSP char) — the equality is structural + textual.
            Assert.Equal(expected,
                System.Xml.Linq.XElement.Parse(actual!).ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
        }
    }

    // Regression: rendering a block AFTER a structural edit added a paragraph used to
    // throw "should never set ilvl more than once" — re-initializing ListItemRetriever
    // over a partially annotated live document was not idempotent, and the editor's
    // Enter-split silently dropped the split when the subsequent block render errored.
    [Fact]
    public void HCO083_RenderBlockHtml_AfterSplit_OnListBearingDocument()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true };

        var target = session.ListBlocks().Body.First(u => u.Kind == "p");
        var res = session.SplitParagraph(target.Id, 6);
        Assert.True(res.Success, res.Error?.Message);

        var first = HtmlConversionOps.RenderBlockHtml(session, res.Modified[0].Id, options);
        var second = HtmlConversionOps.RenderBlockHtml(session, res.Created[0].Id, options);
        Assert.StartsWith("<p", first);
        Assert.StartsWith("<p", second);
    }

    // A default-formatted run containing only w:br must not contribute its own font-size
    // strut when the paragraph declares exact line spacing. Chromium otherwise expands a
    // 10pt arcade row to ~15.3px because the generated 11pt span wraps the break.
    [Fact]
    public void HCO084_ExactLineHeight_RendersBreakOnlyRunsWithoutStyledSpan()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(
                    new Wp.ParagraphProperties(
                        new Wp.SpacingBetweenLines
                        {
                            Line = "200",
                            LineRule = Wp.LineSpacingRuleValues.Exact,
                        }),
                    new Wp.Run(
                        new Wp.RunProperties(new Wp.FontSize { Val = "16" }),
                        new Wp.Text("ROW ONE")),
                    new Wp.Run(new Wp.Break()),
                    new Wp.Run(
                        new Wp.RunProperties(new Wp.FontSize { Val = "16" }),
                        new Wp.Text("ROW TWO")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(new Wp.FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document.Save();
        }

        var html = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });
        var root = System.Xml.Linq.XElement.Parse(html);
        var paragraph = root.Descendants().Single(e => e.Name.LocalName == "p");
        var lineBreak = paragraph.Descendants().Single(e => e.Name.LocalName == "br");

        Assert.Contains("line-height: 10.0pt", html);
        Assert.Contains("vertical-align: top", html);
        Assert.Equal("p", lineBreak.Parent!.Name.LocalName);
        Assert.Contains("font-size: 0", (string?)lineBreak.Attribute("style"));
        Assert.Contains("line-height: 0", (string?)lineBreak.Attribute("style"));
        Assert.DoesNotContain(paragraph.Nodes().OfType<System.Xml.Linq.XText>(),
            text => text.Value.Contains('\u200e'));
        Assert.DoesNotContain(paragraph.Descendants(), e =>
            e.Name.LocalName == "span" && e.Descendants().Any(d => d.Name.LocalName == "br"));
    }

    // An empty Word paragraph is represented only by its paragraph mark, whose line box uses the
    // font's native single-line height. Treating an auto value such as 259 as 107.9% of font-size
    // makes repeated blank lines materially too short. The generated HTML keeps `normal` as that
    // native base and applies the OOXML multiplier to the synthesized inline placeholder.
    [Fact]
    public void HCO085_EmptyParagraphAutoLineSpacing_MultipliesNativeLineHeight()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph()));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(new Wp.FontSize { Val = "22" })),
                    new Wp.ParagraphPropertiesDefault(
                        new Wp.ParagraphPropertiesBaseStyle(
                            new Wp.SpacingBetweenLines
                            {
                                Line = "259",
                                LineRule = Wp.LineSpacingRuleValues.Auto,
                            }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document.Save();
        }

        var html = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });
        var root = System.Xml.Linq.XElement.Parse(html);
        var paragraph = root.Descendants().Single(e => e.Name.LocalName == "p");
        var placeholder = paragraph.Elements().Single(e => e.Name.LocalName == "span");

        Assert.Contains("--docx-auto-line-spacing: 1.079", (string?)paragraph.Attribute("style"));
        Assert.Contains("line-height: normal", (string?)paragraph.Attribute("style"));
        Assert.Contains("line-height: calc(1lh * var(--docx-auto-line-spacing))",
            (string?)placeholder.Attribute("style"));
    }

    // Error contract: an unresolvable anchor maps to JSON null; good anchors in the
    // same call still render. (The reconciler falls back to a full remount per null.)
    [Fact]
    public void HCO082_RenderBlocksHtml_UnresolvableAnchorMapsToNull()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var options = new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true };
        var goodId = session.ListBlocks().Body[0].Id;
        var json = HtmlConversionOps.RenderBlocksHtml(
            session, new[] { goodId, "p:body:00000000000000000000000000000000" }, options);
        using var map = System.Text.Json.JsonDocument.Parse(json);
        Assert.False(string.IsNullOrEmpty(map.RootElement.GetProperty(goodId).GetString()));
        Assert.Equal(System.Text.Json.JsonValueKind.Null,
            map.RootElement.GetProperty("p:body:00000000000000000000000000000000").ValueKind);
    }
}
