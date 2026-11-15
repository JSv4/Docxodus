using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Strict and transitional write some values differently, so converting a Strict package means more than
/// renaming its namespaces (issue #1758). Strict writes the DrawingML percentage types as "50%" where
/// transitional writes thousandths of a percent ("50000"), and gives <c>w:cnfStyle</c> only its twelve
/// flags where transitional also requires <c>w:val</c>.
/// </summary>
public class StrictOoxmlNormalizerValueTests
{
    private const string StrictW = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string StrictA = "http://purl.oclc.org/ooxml/drawingml/main";
    private const string StrictWp = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
    private const string StrictC = "http://purl.oclc.org/ooxml/drawingml/chart";
    private const string Wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    private const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly string[] Flags =
    {
        "firstRow", "lastRow", "firstColumn", "lastColumn", "oddVBand", "evenVBand", "oddHBand", "evenHBand",
        "firstRowFirstColumn", "firstRowLastColumn", "lastRowFirstColumn", "lastRowLastColumn",
    };

    /// <summary>A Strict <c>w:cnfStyle</c>: the twelve flags, the named ones on, and no <c>w:val</c>.</summary>
    private static string Cnf(params string[] on) =>
        "<w:cnfStyle " + string.Join(" ", Flags.Select(f => $"w:{f}=\"{(on.Contains(f) ? 1 : 0)}\"")) + "/>";

    // A Strict save folds a shape's wps elements into the wordprocessingDrawing namespace.
    private static string Document(string body, bool strict = true) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        $"<w:document xmlns:w=\"{(strict ? StrictW : W.NamespaceName)}\"" +
        $" xmlns:a=\"{(strict ? StrictA : A.NamespaceName)}\"" +
        $" xmlns:wp=\"{(strict ? StrictWp : "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")}\"" +
        $" xmlns:wps=\"{(strict ? StrictWp : Wps)}\"" +
        $" xmlns:wp14=\"{Wp14}\" xmlns:a14=\"{A14}\"{(strict ? " w:conformance=\"strict\"" : "")}>" +
        $"<w:body>{body}<w:sectPr/></w:body></w:document>";

    /// <summary>A package holding the given parts beside the main document, in either conformance class.</summary>
    private static byte[] Package(string documentXml, bool strict = true, params (string Name, string Xml)[] extra)
    {
        var rel = strict
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string content)
            {
                using var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                w.Write(content);
            }
            Add("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>");
            Add("_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                $"<Relationship Id=\"rId1\" Type=\"{rel}\" Target=\"word/document.xml\"/></Relationships>");
            Add("word/document.xml", documentXml);
            foreach (var (name, xml) in extra) Add(name, xml);
        }
        return ms.ToArray();
    }

    private static XDocument Normalized(byte[] package, string entry = "word/document.xml")
    {
        var bytes = StrictOoxmlNormalizer.NormalizeToTransitional(package);
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(entry)!.Open());
        return XDocument.Parse(reader.ReadToEnd());
    }

    private static string Shape(string fill) =>
        "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"Badge\"/>" +
        $"<a:graphic><a:graphicData uri=\"{Wps}\"><wps:wsp><wps:cNvSpPr/><wps:spPr>" +
        "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{fill}</wps:spPr><wps:bodyPr/></wps:wsp>" +
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    [Fact]
    public void Drawing_percentages_become_thousandths()
    {
        var body = Shape(
            "<a:gradFill><a:gsLst>" +
            "<a:gs pos=\"0%\"><a:schemeClr val=\"accent1\"><a:lumMod val=\"50%\"/><a:satOff val=\"-12.5%\"/></a:schemeClr></a:gs>" +
            "<a:gs pos=\"33.3335%\"><a:scrgbClr r=\"100%\" g=\"0%\" b=\"7.25%\"/></a:gs>" +
            "</a:gsLst><a:tileRect l=\"-5%\" t=\"0%\" r=\"10%\" b=\"0%\"/></a:gradFill>") +
            "<w:p><w:r><w:drawing><wp:anchor><wp14:sizeRelH relativeFrom=\"margin\"><wp14:pctWidth>40%</wp14:pctWidth></wp14:sizeRelH>" +
            "<wp14:pctPosHOffset>7.5%</wp14:pctPosHOffset>" +
            "<a:graphic><a:graphicData uri=\"http://purl.oclc.org/ooxml/drawingml/picture\"><a:blip>" +
            $"<a:extLst><a:ext uri=\"{{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}}\"><a14:imgProps><a14:imgLayer><a14:imgEffect>" +
            "<a14:sharpenSoften amount=\"-100%\"/></a14:imgEffect><a14:imgEffect><a14:artisticCement trans=\"20%\" crackSpacing=\"35\"/>" +
            "</a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext></a:extLst>" +
            "<a:lum bright=\"24%\" contrast=\"-38%\"/></a:blip></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";

        var doc = Normalized(Package(Document(body)));

        XNamespace wp14 = Wp14, a14 = A14;
        Assert.Equal(new[] { "0", "33334" }, doc.Descendants(A + "gs").Select(e => (string?)e.Attribute("pos")));
        Assert.Equal("50000", (string?)doc.Descendants(A + "lumMod").Single().Attribute("val"));
        Assert.Equal("-12500", (string?)doc.Descendants(A + "satOff").Single().Attribute("val"));
        var scrgb = doc.Descendants(A + "scrgbClr").Single();
        Assert.Equal(new[] { "100000", "0", "7250" }, new[] { "r", "g", "b" }.Select(n => (string?)scrgb.Attribute(n)));
        var tileRect = doc.Descendants(A + "tileRect").Single();
        Assert.Equal(new[] { "-5000", "0", "10000", "0" }, new[] { "l", "t", "r", "b" }.Select(n => (string?)tileRect.Attribute(n)));
        Assert.Equal("40000", doc.Descendants(wp14 + "pctWidth").Single().Value);
        Assert.Equal("7500", doc.Descendants(wp14 + "pctPosHOffset").Single().Value);
        Assert.Equal("-100000", (string?)doc.Descendants(a14 + "sharpenSoften").Single().Attribute("amount"));
        var cement = doc.Descendants(a14 + "artisticCement").Single();
        Assert.Equal("20000", (string?)cement.Attribute("trans"));
        Assert.Equal("35", (string?)cement.Attribute("crackSpacing"));
        var lum = doc.Descendants(A + "lum").Single();
        Assert.Equal(new[] { "24000", "-38000" }, new[] { "bright", "contrast" }.Select(n => (string?)lum.Attribute(n)));
    }

    [Fact]
    public void Theme_and_chart_parts_are_converted_too()
    {
        var theme =
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"{StrictA}\" name=\"Office\">" +
            "<a:themeElements><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:gradFill><a:gsLst>" +
            "<a:gs pos=\"100%\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103%\"/><a:tint val=\"94%\"/></a:schemeClr></a:gs>" +
            "</a:gsLst></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\"><a:miter lim=\"800%\"/></a:ln></a:lnStyleLst>" +
            "</a:fmtScheme></a:themeElements></a:theme>";
        var chart =
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"{StrictC}\" xmlns:a=\"{StrictA}\">" +
            "<c:chart><c:plotArea><c:valAx><c:numFmt formatCode=\"0%\" sourceLinked=\"0\"/><c:txPr><a:bodyPr/><a:p><a:pPr>" +
            "<a:defRPr baseline=\"0%\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65%\"/><a:lumOff val=\"35%\"/>" +
            "</a:schemeClr></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:valAx></c:plotArea></c:chart></c:chartSpace>";
        var package = Package(Document("<w:p/>"), true, ("word/theme/theme1.xml", theme), ("word/charts/chart1.xml", chart));

        var themeDoc = Normalized(package, "word/theme/theme1.xml");
        Assert.Equal("100000", (string?)themeDoc.Descendants(A + "gs").Single().Attribute("pos"));
        Assert.Equal("103000", (string?)themeDoc.Descendants(A + "satMod").Single().Attribute("val"));
        Assert.Equal("94000", (string?)themeDoc.Descendants(A + "tint").Single().Attribute("val"));
        Assert.Equal("800000", (string?)themeDoc.Descendants(A + "miter").Single().Attribute("lim"));

        var chartDoc = Normalized(package, "word/charts/chart1.xml");
        Assert.Equal("0", (string?)chartDoc.Descendants(A + "defRPr").Single().Attribute("baseline"));
        Assert.Equal(new[] { "65000", "35000" },
            chartDoc.Descendants().Where(e => e.Name == A + "lumMod" || e.Name == A + "lumOff").Select(e => (string?)e.Attribute("val")));
        // A number format is text, not a percentage, even where it reads like one.
        Assert.Equal("0%", (string?)chartDoc.Descendants().Single(e => e.Name.LocalName == "numFmt").Attribute("formatCode"));
    }

    [Fact]
    public void A_strict_part_in_a_transitional_package_is_converted_alone()
    {
        var chart =
            $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"{StrictC}\" xmlns:a=\"{StrictA}\">" +
            "<c:chart><c:plotArea><c:barChart><c:ser><c:spPr><a:solidFill><a:srgbClr val=\"5B9BD5\"><a:lumMod val=\"75%\"/></a:srgbClr>" +
            "</a:solidFill></c:spPr></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";
        // The transitional document names the Strict chart namespace in its text; only a part whose own
        // root is Strict converts, so the document is left exactly as it was.
        var document = Document("<w:p><w:r><w:t>Charts here use " + StrictC + ".</w:t></w:r></w:p>", strict: false);
        var package = Package(document, false, ("word/charts/chart1.xml", chart));

        var chartDoc = Normalized(package, "word/charts/chart1.xml");
        Assert.Equal(XName.Get("chartSpace", "http://schemas.openxmlformats.org/drawingml/2006/chart"), chartDoc.Root!.Name);
        Assert.Equal("75000", (string?)chartDoc.Descendants(A + "lumMod").Single().Attribute("val"));
        using var zip = new ZipArchive(new MemoryStream(StrictOoxmlNormalizer.NormalizeToTransitional(package)), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
        Assert.Equal(document, reader.ReadToEnd());
    }

    [Fact]
    public void Each_part_is_read_no_further_than_a_root_start_tag_runs()
    {
        // Every part of every package is sniffed before the depth door, so the sniff is bounded: a
        // root start tag padded past any real one is not read to its end, and the part is left as is.
        var padded = $"<c:chartSpace xmlns:c=\"{StrictC}\" pad=\"{new string('x', 1 << 21)}\"/>";
        var package = Package(Document("<w:p/>", strict: false), false, ("word/charts/chart1.xml", padded));
        Assert.Same(package, StrictOoxmlNormalizer.NormalizeToTransitional(package));
    }

    [Fact]
    public void Text_that_looks_like_a_percentage_is_left_alone()
    {
        var body =
            "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"1\" cy=\"1\"/><wp:docPr id=\"2\" name=\"50%\" descr=\"50%\"/>" +
            "<a:graphic><a:graphicData uri=\"http://purl.oclc.org/ooxml/drawingml/picture\">" +
            "<a:hlinkClick tooltip=\"75%\"/><a:tile tx=\"0\" ty=\"0\" sx=\"50%\" sy=\"50%\"/>" +
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t>50%</w:t></w:r></w:p>";

        var doc = Normalized(Package(Document(body)));

        Assert.Equal("50%", (string?)doc.Descendants().Single(e => e.Name.LocalName == "docPr").Attribute("descr"));
        Assert.Equal("75%", (string?)doc.Descendants(A + "hlinkClick").Single().Attribute("tooltip"));
        Assert.Equal("50%", doc.Descendants(W + "t").Single().Value);
        // The control: on the same tile, sx is a percentage and does change.
        Assert.Equal("50000", (string?)doc.Descendants(A + "tile").Single().Attribute("sx"));
    }

    [Fact]
    public void Values_already_in_thousandths_or_out_of_range_are_left_alone()
    {
        var body = Shape("<a:solidFill><a:schemeClr val=\"accent1\"><a:lumMod val=\"50000\"/>" +
                         "<a:lumOff val=\"79228162514264337593543950335%\"/></a:schemeClr></a:solidFill>") +
                   "<w:tbl><w:tr><w:trPr><w:cnfStyle w:val=\"010000000000\" w:lastRow=\"1\"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>";

        var doc = Normalized(Package(Document(body)));

        Assert.Equal("50000", (string?)doc.Descendants(A + "lumMod").Single().Attribute("val"));
        // No integer holds it in thousandths, so it is left as it was rather than overflowing.
        Assert.Equal("79228162514264337593543950335%", (string?)doc.Descendants(A + "lumOff").Single().Attribute("val"));
        var cnf = doc.Descendants(W + "cnfStyle").Single();
        Assert.Equal("010000000000", (string?)cnf.Attribute(W + "val"));
        Assert.Equal(2, cnf.Attributes().Count());
    }

    [Theory]
    [InlineData("firstRow", "100000000000")]
    [InlineData("oddHBand", "000000100000")]
    // The corner digits run NE, NW, SE, SW: the reference conversion writes a top-left cell as 000000000100.
    [InlineData("firstRowLastColumn", "000000001000")]
    [InlineData("firstRowFirstColumn", "000000000100")]
    [InlineData("lastRowLastColumn", "000000000010")]
    [InlineData("lastRowFirstColumn", "000000000001")]
    public void CnfStyle_gains_the_val_its_flags_spell(string flag, string val)
    {
        var body = $"<w:tbl><w:tr><w:trPr>{Cnf(flag)}</w:trPr><w:tc><w:tcPr>{Cnf(flag)}</w:tcPr><w:p><w:pPr>{Cnf(flag)}</w:pPr></w:p></w:tc></w:tr></w:tbl>";

        var doc = Normalized(Package(Document(body)));

        var cnfs = doc.Descendants(W + "cnfStyle").ToList();
        Assert.Equal(3, cnfs.Count);
        Assert.All(cnfs, cnf =>
        {
            Assert.Equal(val, (string?)cnf.Attribute(W + "val"));
            Assert.Equal("1", (string?)cnf.Attribute(W + flag));
            Assert.Equal(13, cnf.Attributes().Count());
        });
    }

    /// <summary>The issue's pair: a table whose header row carries <c>w:cnfStyle</c> and a shape darkened by
    /// <c>a:lumMod</c>, with one word of the closing paragraph changed.</summary>
    private static WmlDocument IssuePair(bool strict, string closing)
    {
        var cnf = strict ? Cnf("firstRow") : Cnf("firstRow").Replace("<w:cnfStyle ", "<w:cnfStyle w:val=\"100000000000\" ");
        string cell(string text, bool header) =>
            $"<w:tc><w:tcPr>{(header ? cnf : "")}<w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr>" +
            $"<w:p>{(header ? $"<w:pPr>{cnf}</w:pPr>" : "")}<w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        var body =
            "<w:p><w:r><w:t>Loading list</w:t></w:r></w:p>" +
            "<w:tbl><w:tblPr><w:tblW w:w=\"4800\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
            $"<w:tr><w:trPr>{cnf}</w:trPr>{cell("Crate", true)}{cell("Weight", true)}</w:tr>" +
            $"<w:tr>{cell("Ropes", false)}{cell("12 kg", false)}</w:tr></w:tbl>" +
            Shape($"<a:solidFill><a:schemeClr val=\"accent1\"><a:lumMod val=\"{(strict ? "50%" : "50000")}\"/></a:schemeClr></a:solidFill>") +
            $"<w:p><w:r><w:t xml:space=\"preserve\">{closing}</w:t></w:r></w:p>";
        return new WmlDocument("pair.docx", Package(Document(body, strict), strict));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)] // the control: the same pair written as transitional
    public void The_redline_of_a_strict_pair_validates_like_its_transitional_control(bool strict)
    {
        var original = IssuePair(strict, "Check every rope before the van leaves.");
        var revised = IssuePair(strict, "Check every strap before the van leaves.");

        var redline = DocxCompare.Compare(original, revised);

        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
        Assert.Empty(errors);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Contains("strap", body.InnerText);
        Assert.NotEmpty(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.InsertedRun>());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)] // the control: the same shape written as transitional
    public void A_strict_shape_renders_with_its_darkened_fill(bool strict)
    {
        // 4472C4 at 50% luminance is 223962.
        var shape = Shape($"<a:solidFill><a:srgbClr val=\"4472C4\"><a:lumMod val=\"{(strict ? "50%" : "50000")}\"/></a:srgbClr></a:solidFill>");

        var html = HtmlConversionOps.ConvertToHtml(Package(Document(shape, strict), strict),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("223962", html, System.StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("4472C4", html, System.StringComparison.OrdinalIgnoreCase);
    }
}
