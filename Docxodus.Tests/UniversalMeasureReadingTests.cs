using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A Strict save writes every twips measure as a universal measure with a unit (<c>w:pgSz
/// w:w="595.30pt"</c>, <c>w:tcW w:w="120pt"</c>, a style's <c>w:ind w:start="36pt"</c>), and the
/// transitional measure types admit the same form (issue #1764). One document, an A4 page with 2 cm
/// margins, is written twice: as a Strict package with units, and as a transitional package with the
/// same measures in twips. Every reader must see the same document in both.
/// </summary>
public class UniversalMeasureReadingTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // (Strict value, transitional value) for each measure.
    private static readonly (string Strict, string Transitional) PageWidth = ("595.30pt", "11906"), PageHeight = ("841.90pt", "16838"),
        Margin = ("56.70pt", "1134"), Distance = ("35.40pt", "708"), Indent = ("36pt", "720"), Before = ("18pt", "360"),
        Hanging = ("18pt", "360"), Line = ("15pt", "300"), Table = ("240pt", "4800"), Cell = ("120pt", "2400"),
        TableIndent = ("5.40pt", "108"), Tab = ("42.55pt", "851");

    private static byte[] MeasuresPackage(bool strict)
    {
        string V((string Strict, string Transitional) m) => strict ? m.Strict : m.Transitional;
        var w = strict ? "http://purl.oclc.org/ooxml/wordprocessingml/main" : W.NamespaceName;
        var rel = strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string Cell(string text) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"{V(UniversalMeasureReadingTests.Cell)}\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        var document =
            $"<w:document xmlns:w=\"{w}\"{(strict ? " w:conformance=\"strict\"" : "")}><w:body>" +
            $"<w:p><w:pPr><w:spacing w:before=\"{V(Before)}\"/><w:ind w:start=\"{V(Indent)}\"/></w:pPr><w:r><w:t>Indented paragraph.</w:t></w:r></w:p>" +
            $"<w:p><w:pPr><w:spacing w:line=\"{V(Line)}\" w:lineRule=\"exact\"/></w:pPr><w:r><w:t>Exactly spaced line.</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Load the kiln.</w:t></w:r></w:p>" +
            $"<w:tbl><w:tblPr><w:tblStyle w:val=\"FiringGrid\"/><w:tblW w:w=\"{V(Table)}\" w:type=\"dxa\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
            $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"{V(Table)}\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>Firing log</w:t></w:r></w:p></w:tc></w:tr>" +
            $"<w:tr>{Cell("Kiln")}{Cell("Cone 6")}</w:tr></w:tbl>" +
            "<w:p><w:r><w:t>Vent every flue before the night shift.</w:t></w:r></w:p>" +
            $"<w:sectPr><w:pgSz w:w=\"{V(PageWidth)}\" w:h=\"{V(PageHeight)}\"/><w:pgMar w:top=\"{V(Margin)}\" w:right=\"{V(Margin)}\" " +
            $"w:bottom=\"{V(Margin)}\" w:left=\"{V(Margin)}\" w:header=\"{V(Distance)}\" w:footer=\"{V(Distance)}\" w:gutter=\"0{(strict ? "pt" : "")}\"/></w:sectPr>" +
            "</w:body></w:document>";
        var styles =
            $"<w:styles xmlns:w=\"{w}\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
            $"<w:style w:type=\"paragraph\" w:styleId=\"Note\"><w:name w:val=\"Note\"/><w:pPr><w:spacing w:before=\"{V(Before)}\"/><w:ind w:start=\"{V(Indent)}\"/></w:pPr></w:style>" +
            $"<w:style w:type=\"table\" w:styleId=\"FiringGrid\"><w:name w:val=\"Firing Grid\"/><w:tblPr><w:tblW w:w=\"{V(Table)}\" w:type=\"dxa\"/>" +
            $"<w:tblInd w:w=\"{V(TableIndent)}\" w:type=\"dxa\"/></w:tblPr></w:style></w:styles>";
        var numbering =
            $"<w:numbering xmlns:w=\"{w}\"><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>" +
            $"<w:lvlText w:val=\"%1.\"/><w:pPr><w:ind w:start=\"{V(Indent)}\" w:hanging=\"{V(Hanging)}\"/></w:pPr></w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>";
        return Package(rel, ("word/document.xml", document), ("word/styles.xml", styles), ("word/numbering.xml", numbering),
            ("word/settings.xml", $"<w:settings xmlns:w=\"{w}\"><w:defaultTabStop w:val=\"{V(Tab)}\"/></w:settings>"));
    }

    private static byte[] Package(string rel, params (string Name, string Xml)[] parts)
    {
        string ContentType(string name) => name switch
        {
            "word/document.xml" => "wordprocessingml.document.main+xml",
            "word/styles.xml" => "wordprocessingml.styles+xml",
            "word/numbering.xml" => "wordprocessingml.numbering+xml",
            "word/settings.xml" => "wordprocessingml.settings+xml",
            _ => "drawingml.chart+xml",
        };
        string RelType(string name) => name switch
        {
            "word/styles.xml" => "styles",
            "word/numbering.xml" => "numbering",
            "word/settings.xml" => "settings",
            _ => "chart",
        };
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + content);
            }
            Add("[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                string.Concat(parts.Select(p => $"<Override PartName=\"/{p.Name}\" ContentType=\"application/vnd.openxmlformats-officedocument.{ContentType(p.Name)}\"/>")) +
                "</Types>");
            Add("_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"rId1\" Type=\"{rel}/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Add("word/_rels/document.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                string.Concat(parts.Skip(1).Select((p, i) => $"<Relationship Id=\"rIdP{i}\" Type=\"{rel}/{RelType(p.Name)}\" Target=\"{p.Name["word/".Length..]}\"/>")) +
                "</Relationships>");
            foreach (var (name, xml) in parts) Add(name, xml);
        }
        return ms.ToArray();
    }

    private static string Anchor(DocxSession session, string kind, int index = 0) =>
        session.Project().AnchorIndex.Values.Select(v => v.Anchor).Where(a => a.Kind == kind).Select(a => a.Id).ElementAt(index);

    private static XElement SavedBody(DocxSession session)
    {
        using var zip = new ZipArchive(new MemoryStream(session.Save()), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("word/document.xml")!.Open();
        return XDocument.Load(stream).Root!.Element(W + "body")!;
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Section_info_reads_the_page_in_either_form(bool strict)
    {
        using var session = new DocxSession(MeasuresPackage(strict));
        var info = session.GetSectionInfo(Anchor(session, "p"))!;

        Assert.Equal((11906, 16838), (info.PageWidthTwips, info.PageHeightTwips));
        Assert.Equal((1134, 1134, 1134, 1134), (info.MarginTopTwips, info.MarginRightTwips, info.MarginBottomTwips, info.MarginLeftTwips));
        Assert.Equal((708, 708), (info.HeaderDistanceTwips, info.FooterDistanceTwips));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Page_setup_changes_only_what_it_names(bool strict)
    {
        using var margins = new DocxSession(MeasuresPackage(strict));
        Assert.True(margins.SetPageSetup(Anchor(margins, "p"), new PageSetupOp { MarginTopTwips = 720 }).Success);
        var pgMar = SavedBody(margins).Element(W + "sectPr")!.Element(W + "pgMar")!;
        Assert.Equal(new[] { "720", "1134", "1134", "1134", "708", "708" },
            new[] { "top", "right", "bottom", "left", "header", "footer" }.Select(a => (string?)pgMar.Attribute(W + a)));

        using var rotate = new DocxSession(MeasuresPackage(strict));
        Assert.True(rotate.SetPageSetup(Anchor(rotate, "p"), new PageSetupOp { Landscape = true }).Success);
        var pgSz = SavedBody(rotate).Element(W + "sectPr")!.Element(W + "pgSz")!;
        Assert.Equal(("16838", "11906", "landscape"), ((string?)pgSz.Attribute(W + "w"), (string?)pgSz.Attribute(W + "h"), (string?)pgSz.Attribute(W + "orient")));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Column_edits_resize_a_spanning_cell_written_with_units(bool strict)
    {
        using var insert = new DocxSession(MeasuresPackage(strict));
        var inserted = insert.InsertTableColumn(Anchor(insert, "tc", 1), Position.After);
        Assert.True(inserted.Success, inserted.Error?.Message);
        Assert.Equal("7200", SpanningCellWidth(insert));

        using var delete = new DocxSession(MeasuresPackage(strict));
        var deleted = delete.DeleteTableColumn(Anchor(delete, "tc", 1));
        Assert.True(deleted.Success, deleted.Error?.Message);
        Assert.Equal("2400", SpanningCellWidth(delete));

        static string? SpanningCellWidth(DocxSession session) =>
            (string?)SavedBody(session).Descendants(W + "tc").First().Element(W + "tcPr")!.Element(W + "tcW")!.Attribute(W + "w");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Formatting_and_list_readers_read_indents_and_spacing_with_units(bool strict)
    {
        using var session = new DocxSession(MeasuresPackage(strict));

        var direct = session.GetFormatting(Anchor(session, "p"))!.DirectParagraph;
        Assert.Equal((720, 360), (direct.LeftIndentTwips, direct.SpacingBeforeTwips));
        Assert.Equal(300, session.GetFormatting(Anchor(session, "p", 1))!.DirectParagraph.LineSpacing);

        var styles = session.ListStyles();
        var note = styles.Single(s => s.Id == "Note").ResolvedParagraph!;
        Assert.Equal((720, 360), (note.LeftIndentTwips, note.SpacingBeforeTwips));
        var grid = styles.Single(s => s.Id == "FiringGrid").ResolvedTable!;
        Assert.Equal((4800, 108), (grid.WidthTwips, grid.IndentTwips));

        var list = session.GetListMembership(Anchor(session, "li"))!;
        Assert.Equal((720, 360), (list.LeftIndentTwips, list.HangingIndentTwips));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Html_keeps_the_page_size_and_the_table_and_cell_widths(bool strict)
    {
        var bytes = MeasuresPackage(strict);
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("measures.docx", bytes),
            new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated }).ToString();

        Assert.Contains("data-page-width=\"595.3\" data-page-height=\"841.9\"", html);
        var flow = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("measures.docx", bytes),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var table = flow.Descendants().First(e => e.Name.LocalName == "table");
        Assert.Contains("width: 240pt", (string?)table.Attribute("style"));
        Assert.Equal(new[] { true, true }, table.Descendants().Where(e => e.Name.LocalName == "td").Skip(1)
            .Select(td => ((string?)td.Attribute("style") ?? "").Contains("width: 120pt")));
    }

    // A transitional package whose chart part alone is written in the Strict namespaces (the #1702
    // shape), against the same package with a transitional chart part.
    private static byte[] ChartPackage(bool strictChart)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string tc = "http://schemas.openxmlformats.org/drawingml/2006/chart", ta = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var (c, a) = strictChart ? ("http://purl.oclc.org/ooxml/drawingml/chart", "http://purl.oclc.org/ooxml/drawingml/main") : (tc, ta);
        static string Points(params object[] values) =>
            $"<c:ptCount val=\"{values.Length}\"/>" + string.Concat(values.Select((v, i) => $"<c:pt idx=\"{i}\"><c:v>{v}</c:v></c:pt>"));
        static string Axis(string kind, int id, string position, int cross) =>
            $"<c:{kind}><c:axId val=\"{id}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"{position}\"/><c:crossAx val=\"{cross}\"/></c:{kind}>";
        var chart =
            $"<c:chartSpace xmlns:c=\"{c}\" xmlns:a=\"{a}\"><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/><c:barChart><c:barDir val=\"col\"/>" +
            "<c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
            $"<c:spPr><a:solidFill><a:srgbClr val=\"5B9BD5\"><a:lumMod val=\"{(strictChart ? "75%" : "75000")}\"/></a:srgbClr></a:solidFill></c:spPr>" +
            $"<c:cat><c:strLit>{Points("Kiln", "Flue", "Damper")}</c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode>{Points(41, 36, 29)}</c:numLit></c:val>" +
            "</c:ser><c:axId val=\"501\"/><c:axId val=\"502\"/></c:barChart>" + Axis("catAx", 501, "b", 502) + Axis("valAx", 502, "l", 501) +
            "</c:plotArea><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>";
        var document =
            $"<w:document xmlns:w=\"{w}\" xmlns:r=\"{r}\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"{ta}\" xmlns:c=\"{tc}\"><w:body>" +
            "<w:p><w:r><w:t>Firings per week</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
            "<wp:extent cx=\"4572000\" cy=\"2743200\"/><wp:docPr id=\"1\" name=\"Firing chart\"/><a:graphic>" +
            $"<a:graphicData uri=\"{tc}\"><c:chart r:id=\"rIdP0\"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>" +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";
        return Package(r, ("word/document.xml", document), ("word/charts/chart1.xml", chart));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_strict_chart_part_in_a_transitional_package_is_drawn(bool strictChart)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("chart.docx", ChartPackage(strictChart)),
            new WmlToHtmlConverterSettings()).ToString();

        Assert.Equal(3, Regex.Matches(html, "<rect[^>]*fill=\"#4" + "474A0\"", RegexOptions.IgnoreCase).Count);
        foreach (var label in new[] { "Kiln", "Flue", "Damper" })
            Assert.Contains($">{label}</text>", html);
    }
}
