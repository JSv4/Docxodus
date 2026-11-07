// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// DrawingML shapes and groups that are not text boxes render as inline SVG of their extent, and an
/// inline drawing the converter cannot draw still occupies its extent (issue #1270). Before, such a
/// drawing produced no HTML while the empty-paragraph pass still counted it as content, so its
/// paragraph had no line box and PDF export failed PageMap verification.
/// </summary>
public class HtmlDrawingShapeTests
{
    private const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " +
        "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\"";

    private const string ShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private const string GroupUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    private const string Teal = "<a:solidFill><a:srgbClr val=\"19647E\"/></a:solidFill>";

    /// <summary>A 2 × 1 inch shape, as the issue's reproduction generator writes it.</summary>
    private static string Shape(string geometry = "rect", string fill = Teal, string xfrm = "",
        string extra = "", long cx = 1828800, long cy = 914400) =>
        $"<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm{xfrm}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"{geometry}\"><a:avLst/></a:prstGeom>{fill}{extra}</wps:spPr><wps:bodyPr/></wps:wsp>";

    private static string CustomShape(string paths) => Shape().Replace(
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>",
        "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"r\" b=\"b\"/>" +
        "<a:pathLst>" + paths + "</a:pathLst></a:custGeom>");

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LiteralCustomPaths_PreserveIndependentCoordinatesAndPaintFlags(bool grouped, bool paginated)
    {
        var shape = CustomShape("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"100\"/></a:moveTo>" +
            "<a:lnTo><a:pt x=\"50\" y=\"0\"/></a:lnTo><a:quadBezTo><a:pt x=\"75\" y=\"25\"/><a:pt x=\"100\" y=\"100\"/></a:quadBezTo>" +
            "<a:cubicBezTo><a:pt x=\"75\" y=\"75\"/><a:pt x=\"25\" y=\"75\"/><a:pt x=\"0\" y=\"100\"/></a:cubicBezTo><a:close/></a:path>" +
            "<a:path w=\"200\" h=\"200\" fill=\"none\" stroke=\"false\"><a:moveTo><a:pt x=\"50\" y=\"50\"/></a:moveTo>" +
            "<a:lnTo><a:pt x=\"150\" y=\"150\"/></a:lnTo></a:path>")
            .Replace("</wps:spPr>", "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"883311\"/></a:solidFill></a:ln></wps:spPr>");
        var input = Docx(InlineDrawing(grouped ? GroupUri : ShapeUri, grouped ? Group(shape) : shape));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, paginated ? PaginationMode.Paginated : PaginationMode.None, placeholders: true);
        var paths = Drawn(OnlySvg(paragraph), "path").ToArray();
        Assert.Equal(2, paths.Length);
        Assert.Equal("M 0 914400 L 914400 0 Q 1371600 228600 1828800 914400 C 1371600 685800 457200 685800 0 914400 Z",
            (string?)paths[0].Attribute("d"));
        Assert.Equal("M 457200 228600 L 1371600 685800", (string?)paths[1].Attribute("d"));
        Assert.Equal("none", (string?)paths[1].Attribute("fill"));
        Assert.Equal("none", (string?)paths[1].Attribute("stroke"));
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
    }

    [Theory]
    [InlineData("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"100\" y=\"50\"/></a:moveTo><a:arcTo wR=\"50\" hR=\"50\" stAng=\"0\" swAng=\"10800000\"/></a:path>")]
    [InlineData("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"wd2\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo></a:path>")]
    [InlineData("<a:path w=\"100\" h=\"100\" fill=\"darken\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo></a:path>")]
    public void UnsupportedCustomPath_RefusesTheWholeGeometry(string unsupported)
    {
        var input = Docx(InlineDrawing(ShapeUri, CustomShape(
            "<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>" +
            "<a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo></a:path>" + unsupported)));
        AssertSchemaValid(input);
        AssertUnrenderedBox(MiddleParagraph(input));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CustomTextbox_PreservesTextAndReportsUnrepresentedGeometry(bool grouped)
    {
        var shape = CustomShape("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>" +
            "<a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo></a:path>")
            .Replace("<wps:bodyPr/>", "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Custom caption</w:t></w:r></w:p>" +
                "</w:txbxContent></wps:txbx><wps:bodyPr><a:noAutofit/></wps:bodyPr>");
        var input = Docx(InlineDrawing(grouped ? GroupUri : ShapeUri, grouped ? Group(shape) : shape));
        AssertSchemaValid(input);
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(input, new HtmlConversionOptions
        {
            RenderUnsupportedContentPlaceholders = true,
        }));
        Assert.Contains("Custom caption", html.Value);
        Assert.Contains(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        Assert.DoesNotContain(html.Descendants(), e => e.Attribute("data-docx-textbox-source-clip") != null);
    }

    [Theory]
    [InlineData("<a:lnTo><a:pt x=\"10\" y=\"10\"/></a:lnTo>")]
    [InlineData("<a:close/>")]
    [InlineData("<a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:cubicBezTo><a:pt x=\"10\" y=\"10\"/></a:cubicBezTo>")]
    [InlineData("<a:moveTo><a:pt x=\"9223372036854775808\" y=\"0\"/></a:moveTo>")]
    public void MalformedCustomPath_KeepsTheExtentInsteadOfPaintingPartialCommands(string commands)
    {
        var input = Docx(InlineDrawing(ShapeUri, CustomShape("<a:path w=\"100\" h=\"100\">" + commands + "</a:path>")));
        AssertUnrenderedBox(MiddleParagraph(input));
    }

    [Theory]
    [InlineData("<a:effectLst><a:glow rad=\"12700\"><a:srgbClr val=\"FF0000\"/></a:glow></a:effectLst>")]
    [InlineData("<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln>")]
    [InlineData("<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:headEnd type=\"triangle\"/></a:ln>")]
    public void CustomPath_WithUnrepresentedPaint_RemainsPartial(string paint)
    {
        var shape = CustomShape("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>" +
            "<a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo></a:path>").Replace("</wps:spPr>", paint + "</wps:spPr>");
        var input = Docx(InlineDrawing(ShapeUri, shape));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, placeholders: true);
        Assert.Single(Drawn(OnlySvg(paragraph), "path"));
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
    }

    [Theory]
    [InlineData("fill", "alphaMod", "0", true)]
    [InlineData("line", "alphaOff", "-100000", true)]
    [InlineData("group", "alphaMod", "0", true)]
    [InlineData("fill", "alpha", "40000", false)]
    [InlineData("line", "alpha", "40000", false)]
    [InlineData("group", "alpha", "40000", false)]
    public void CustomPath_ColorTransformsRemainDiagnosed(string scope, string modifier, string value, bool partial)
    {
        var shape = CustomShape("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo>" +
            "<a:lnTo><a:pt x=\"100\" y=\"100\"/></a:lnTo></a:path>");
        var fill = $"<a:solidFill><a:srgbClr val=\"19647E\"><a:{modifier} val=\"{value}\"/></a:srgbClr></a:solidFill>";
        shape = scope == "line" ? shape.Replace("</wps:spPr>", "<a:ln w=\"25400\">" + fill + "</a:ln></wps:spPr>")
            : shape.Replace(Teal, scope == "group" ? "<a:grpFill/>" : fill);
        var group = Group(shape).Replace("</wpg:grpSpPr>", fill + "</wpg:grpSpPr>");
        var input = Docx(InlineDrawing(scope == "group" ? GroupUri : ShapeUri, scope == "group" ? group : shape));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, placeholders: true);
        Assert.Single(Drawn(OnlySvg(paragraph), "path"));
        Assert.Equal(partial, paragraph.Descendants().Any(e => (string?)e.Attribute("data-content-type") == "Other"));
        if (!partial) Assert.Contains("\"0.4\"", paragraph.ToString());
    }

    /// <summary>The issue's group: the same rectangle under an identity child coordinate space.</summary>
    private static string Group(string children, string childOffset = "0", long childCx = 1828800,
        long childCy = 914400) =>
        "<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1828800\" cy=\"914400\"/>" +
        $"<a:chOff x=\"{childOffset}\" y=\"{childOffset}\"/><a:chExt cx=\"{childCx}\" cy=\"{childCy}\"/></a:xfrm></wpg:grpSpPr>" +
        children + "</wpg:wgp>";

    private static string InlineDrawing(string uri, string data) =>
        "<w:r><w:drawing><wp:inline><wp:extent cx=\"1828800\" cy=\"914400\"/><wp:docPr id=\"7\" name=\"Teal inspection panel\"/>" +
        $"<a:graphic><a:graphicData uri=\"{uri}\">{data}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    private static string FloatingDrawing(string uri, string data) =>
        "<w:r><w:drawing><wp:anchor distT=\"0\" distR=\"114300\" distB=\"0\" distL=\"114300\" simplePos=\"0\" " +
        "relativeHeight=\"10\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/>" +
        "<wp:positionH relativeFrom=\"page\"><wp:posOffset>4572000</wp:posOffset></wp:positionH>" +
        "<wp:positionV relativeFrom=\"page\"><wp:posOffset>4572000</wp:posOffset></wp:positionV>" +
        "<wp:extent cx=\"1828800\" cy=\"914400\"/><wp:wrapNone/><wp:docPr id=\"7\" name=\"Teal inspection panel\"/>" +
        $"<a:graphic><a:graphicData uri=\"{uri}\">{data}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

    private static string Deleted(string runs) =>
        $"<w:del w:id=\"31\" w:author=\"Review Example\" w:date=\"2026-09-12T10:00:00Z\">{runs}</w:del>";

    private static byte[] Docx(string middleParagraphContent)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write(
                    $"<w:document {Namespaces}><w:body>" +
                    "<w:p><w:r><w:t>Optical panel</w:t></w:r></w:p>" +
                    $"<w:p>{middleParagraphContent}</w:p>" +
                    "<w:p><w:r><w:t>End of inspection</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            if (middleParagraphContent.Contains("<w:footnoteReference"))
            {
                using var writer = new StreamWriter(main.AddNewPart<FootnotesPart>().GetStream(FileMode.Create, FileAccess.Write));
                writer.Write("<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/><w:t>Line note</w:t></w:r></w:p></w:footnote></w:footnotes>");
            }
            doc.Save();
        }
        return stream.ToArray();
    }

    private static XElement MiddleParagraph(byte[] docx, PaginationMode mode = PaginationMode.None,
        bool placeholders = false)
    {
        var html = HtmlConversionOps.ConvertToHtml(docx, new HtmlConversionOptions
        {
            FabricateCssClasses = false,
            PaginationMode = (int)mode,
            RenderTrackedChanges = true,
            RenderUnsupportedContentPlaceholders = placeholders,
        });
        return XElement.Parse(html).Descendants().Where(e => e.Name.LocalName == "p")
            .Single(p => !p.Value.Contains("Optical panel") && !p.Value.Contains("End of inspection"));
    }

    private static XElement OnlySvg(XElement paragraph) =>
        paragraph.Descendants().Single(e => e.Name.LocalName == "svg");

    private static XElement[] Drawn(XElement svg, string localName) =>
        svg.Descendants().Where(e => e.Name.LocalName == localName).ToArray();

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void IssueShapeOrGroup_RendersTheTealRectangleAtItsExtent(bool grouped, bool deleted)
    {
        var drawing = grouped ? InlineDrawing(GroupUri, Group(Shape())) : InlineDrawing(ShapeUri, Shape());
        var paragraph = MiddleParagraph(Docx(deleted ? Deleted(drawing) : drawing));

        var svg = OnlySvg(paragraph);
        var box = (string?)svg.Parent!.Attribute("style");
        Assert.Contains("width: 144pt", box);
        Assert.Contains("height: 72pt", box);
        Assert.Equal(("192", "96"), ((string?)svg.Attribute("width"), (string?)svg.Attribute("height")));
        var rect = Assert.Single(Drawn(svg, "rect"));
        Assert.Equal("#19647E", (string?)rect.Attribute("fill"));
        Assert.Equal(deleted, svg.Ancestors().Any(a => a.Name.LocalName == "del"));
    }

    [Fact]
    public void GroupChildren_AreMappedThroughTheChildCoordinateSpace()
    {
        // A child space half the group's size, offset by 1000: the child's frame doubles on render,
        // and its 1pt outline stays 1pt, as Word and LibreOffice draw it.
        var child = "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"1000\" y=\"1000\"/><a:ext cx=\"457200\" cy=\"457200\"/></a:xfrm>" +
                    $"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{Teal}" +
                    "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>";
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing(GroupUri, Group(child, "1000", 914400, 457200)))));

        // The group maps child (1000, 1000)–(458200, 458200) onto (0, 0)–(914400, 914400).
        var rect = Assert.Single(Drawn(svg, "rect"));
        Assert.Equal(("0", "0"), ((string?)rect.Attribute("x"), (string?)rect.Attribute("y")));
        Assert.Equal(("914400", "914400"), ((string?)rect.Attribute("width"), (string?)rect.Attribute("height")));
        Assert.Equal("12700", (string?)rect.Attribute("stroke-width"));
        Assert.All(Drawn(svg, "g"), g => Assert.StartsWith("scale(", (string?)g.Attribute("transform")));
    }

    [Fact]
    public void CanvasChildren_ArePlacedInTheDrawingsOwnCoordinates()
    {
        var child = Shape(cx: 457200, cy: 457200).Replace("<a:off x=\"0\" y=\"0\"/>", "<a:off x=\"914400\" y=\"0\"/>");
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing(
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", $"<wpc:wpc><wpc:bg/><wpc:whole/>{child}</wpc:wpc>"))));

        var rect = Assert.Single(Drawn(svg, "rect"));
        Assert.Equal("914400", (string?)rect.Attribute("x"));
        Assert.Equal("457200", (string?)rect.Attribute("width"));
    }

    [Fact]
    public void Ellipse_RendersAsAnEllipseFillingItsFrame()
    {
        var ellipse = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("ellipse"))))), "ellipse"));
        Assert.Equal("914400", (string?)ellipse.Attribute("rx"));
        Assert.Equal("457200", (string?)ellipse.Attribute("ry"));
        Assert.Equal("#19647E", (string?)ellipse.Attribute("fill"));
    }

    [Fact]
    public void RoundedRectangle_TakesItsCornerRadiusFromTheShorterSide()
    {
        var rect = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("roundRect"))))), "rect"));
        // Default adj 16667 of the 914400 EMU short side.
        Assert.Equal("152403.048", (string?)rect.Attribute("rx"));
    }

    [Theory]
    [InlineData("triangle", 3, false)]
    [InlineData("triangle", 3, true)]
    [InlineData("star5", 10, false)]
    [InlineData("star5", 10, true)]
    public void PresetPolygon_IsDrawnAloneAndInsideAGroup(string preset, int vertices, bool grouped)
    {
        var shape = Shape(preset);
        var drawing = grouped ? InlineDrawing(GroupUri, Group(shape)) : InlineDrawing(ShapeUri, shape);
        var paragraph = MiddleParagraph(Docx(drawing), placeholders: true);
        var polygon = Assert.Single(Drawn(OnlySvg(paragraph), "polygon"));
        Assert.Equal(vertices, ((string)polygon.Attribute("points")!).Split(' ').Length);
        Assert.Equal("#19647E", (string?)polygon.Attribute("fill"));
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
        Assert.DoesNotContain("[SHAPE]", paragraph.Value);
    }

    [Theory]
    [InlineData(-10000, 0)]
    [InlineData(25000, 457200)]
    [InlineData(75000, 1371600)]
    [InlineData(150000, 1828800)]
    public void TriangleApex_FollowsTheClampedAdjustment(int adjustment, int apex)
    {
        var shape = Shape("triangle").Replace("<a:avLst/>", $"<a:avLst><a:gd name=\"adj\" fmla=\"val {adjustment}\"/></a:avLst>");
        var polygon = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, shape)))), "polygon"));
        Assert.Equal($"0,914400 {apex},0 1828800,914400", (string?)polygon.Attribute("points"));
    }

    [Fact]
    public void StarWithZeroInnerRadius_ReturnsToItsDeclaredCenterBetweenTips()
    {
        var shape = Shape("star5").Replace("<a:avLst/>", "<a:avLst><a:gd name=\"adj\" fmla=\"val 0\"/>"
            + "<a:gd name=\"hf\" fmla=\"val 100000\"/><a:gd name=\"vf\" fmla=\"val 100000\"/></a:avLst>");
        var polygon = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, shape)))), "polygon"));
        var points = ((string)polygon.Attribute("points")!).Split(' ');
        Assert.Equal(10, points.Length);
        Assert.Equal("914400,0", points[2]);
        Assert.All(new[] { 1, 3, 5, 7, 9 }, index => Assert.Equal("914400,457200", points[index]));
        // Native preset definitions place the left upper outer tip at 18 degrees.
        var left = points[0].Split(',').Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(44753.921, left[0], 3);
        Assert.Equal(315917.43, left[1], 3);
    }

    [Theory]
    [InlineData("triangle", "adj", "*/ w 1 2")]
    [InlineData("star5", "adj", "val NaN")]
    [InlineData("star5", "hf", "val 1e309")]
    [InlineData("star5", "vf", "val 1e308")]
    public void UnsupportedPresetAdjustments_KeepThePlaceholder(string preset, string guide, string formula)
    {
        var shape = Shape(preset).Replace("<a:avLst/>", $"<a:avLst><a:gd name=\"{guide}\" fmla=\"{formula}\"/></a:avLst>");
        AssertUnrenderedBox(MiddleParagraph(Docx(InlineDrawing(ShapeUri, shape))));
    }

    [Fact]
    public void DefaultStar_FollowsThePublishedPresetPath()
    {
        var polygon = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("star5"))))), "polygon"));
        // Independently evaluated from the published avLst/gdLst/pathLst, at a 1828800×914400 extent.
        var expected = new[] { 1.934, 349268.833, 698543.295, 349271.291, 914400, 0,
            1130256.705, 349271.291, 1828798.066, 349268.833, 1263663.485, 565127.905,
            1479529.084, 914397.677, 914400, 698534.628, 349270.916, 914397.677, 565136.515, 565127.905 };
        var actual = ((string)polygon.Attribute("points")!).Split(' ', ',')
            .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(expected.Length, actual.Length);
        for (var index = 0; index < actual.Length; index++) Assert.Equal(expected[index], actual[index], 3);
    }

    [Fact]
    public void AdjustedTriangle_UsesTheMappedGroupFrameForItsFlipAndRotation()
    {
        var shape = Shape("triangle", xfrm: " flipH=\"1\" rot=\"5400000\"", cx: 457200, cy: 228600)
            .Replace("<a:off x=\"0\" y=\"0\"/>", "<a:off x=\"1000\" y=\"1000\"/>")
            .Replace("<a:avLst/>", "<a:avLst><a:gd name=\"adj\" fmla=\"val 25000\"/></a:avLst>");
        var group = Group(shape, "1000", 914400, 457200);
        var polygon = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(GroupUri, group)))), "polygon"));
        Assert.Equal("0,457200 228600,0 914400,457200", (string?)polygon.Attribute("points"));
        Assert.Equal("rotate(90 457200 228600) matrix(-1 0 0 1 914400 0)", (string?)polygon.Attribute("transform"));
    }

    [Theory]
    [InlineData("child", true, false)]
    [InlineData("child", true, true)]
    [InlineData("nested", true, false)]
    [InlineData("cancelled", false, false)]
    [InlineData("root", false, false)]
    [InlineData("reflection", false, false)]
    [InlineData("half-turn", false, false)]
    public void UnequallyScaledGroupRotation_IsReportedWithoutDeletingPaint(string kind, bool partial, bool text)
    {
        var transform = kind == "reflection" ? " flipH=\"1\"" : kind == "half-turn" ? " rot=\"10800000\""
            : kind is "root" or "nested" ? "" : " rot=\"5400000\"";
        var child = Shape("star5", xfrm: transform);
        if (text) child = child.Replace("<wps:bodyPr/>", "<wps:txbx><w:txbxContent><w:p><w:r>"
            + "<w:t>Group label</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>");
        if (kind is "nested" or "cancelled")
        {
            child = Group(child, childCx: kind == "cancelled" ? 3657600 : 1828800)
                .Replace("wpg:wgp", "wpg:grpSp");
            if (kind == "nested") child = child.Replace("<a:xfrm>", "<a:xfrm rot=\"5400000\">");
        }
        var group = Group(child, childCx: 914400);
        if (kind == "root") group = group.Replace("<wpg:grpSpPr><a:xfrm>", "<wpg:grpSpPr><a:xfrm rot=\"5400000\">");
        var bytes = Docx(InlineDrawing(GroupUri, group));
        using (var stream = new MemoryStream(bytes))
        using (var document = WordprocessingDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document));
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions
        {
            FabricateCssClasses = false, RenderUnsupportedContentPlaceholders = true,
        }));
        Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "polygon"));
        Assert.Equal(partial, html.Descendants().Any(e => (string?)e.Attribute("data-docx-shape") == "partial"));
        if (partial) Assert.Contains(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        if (text) Assert.Contains("Group label", html.Value);
    }

    [Fact]
    public void FlippedLine_RunsCornerToCornerWithItsOutline()
    {
        var outline = "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>";
        var line = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri,
            Shape("line", fill: "", xfrm: " flipH=\"1\"", extra: outline))))), "line"));
        Assert.Equal("none", (string?)line.Attribute("fill"));
        Assert.Equal("#FF0000", (string?)line.Attribute("stroke"));
        Assert.Equal("25400", (string?)line.Attribute("stroke-width"));
        Assert.Equal("matrix(-1 0 0 1 1828800 0)", (string?)line.Attribute("transform"));
    }

    [Fact]
    public void HorizontalLine_WithAZeroHeightExtent_IsStillDrawn()
    {
        var outline = "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>";
        var drawing = InlineDrawing(ShapeUri, Shape("line", fill: "", extra: outline, cy: 0)).Replace("cy=\"914400\"", "cy=\"0\"");
        var svg = OnlySvg(MiddleParagraph(Docx(drawing)));

        Assert.Contains("height: 0pt", (string?)svg.Parent!.Attribute("style"));
        Assert.Null(svg.Attribute("viewBox"));
        // A degenerate viewBox would disable rendering; the export needs a positive size to lay it out.
        Assert.Equal("1", (string?)svg.Attribute("height"));
        var line = Assert.Single(Drawn(svg, "line"));
        Assert.Equal("0", (string?)line.Attribute("y2"));
        Assert.Equal("1828800", (string?)line.Attribute("x2"));
    }

    [Fact]
    public void StyleReferences_ResolveThemeFillAndOutline()
    {
        // Word's default shape style: accent1 fill, and an outline of accent1 at 50% shade.
        var styled = Shape(fill: "").Replace("<wps:bodyPr/>",
            "<wps:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef>" +
            "<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:effectRef>" +
            "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></wps:style><wps:bodyPr/>");
        var bytes = Docx(InlineDrawing(ShapeUri, styled));
        using (var stream = new MemoryStream())
        {
            stream.Write(bytes);
            using (var doc = WordprocessingDocument.Open(stream, true))
            {
                var theme = doc.MainDocumentPart!.AddNewPart<ThemePart>();
                using var writer = new StreamWriter(theme.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements>" +
                             "<a:clrScheme name=\"T\"><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1></a:clrScheme>" +
                             "<a:fmtScheme name=\"T\"><a:lnStyleLst><a:ln w=\"6350\"/><a:ln w=\"12700\"/><a:ln w=\"19050\"/></a:lnStyleLst></a:fmtScheme>" +
                             "</a:themeElements></a:theme>");
            }
            bytes = stream.ToArray();
        }

        var rect = Assert.Single(Drawn(OnlySvg(MiddleParagraph(bytes)), "rect"));
        Assert.Equal("#4472C4", (string?)rect.Attribute("fill"));
        Assert.Equal("#2F528F", (string?)rect.Attribute("stroke"));
        Assert.Equal("12700", (string?)rect.Attribute("stroke-width"));
    }

    /// <summary>The box an undrawable drawing leaves: a blank SVG of its extent, marked and hidden.</summary>
    private static void AssertUnrenderedBox(XElement paragraph)
    {
        var svg = OnlySvg(paragraph);
        Assert.Empty(svg.Elements());
        var box = svg.Parent!;
        Assert.Equal("unrendered", (string?)box.Attribute("data-docx-shape"));
        Assert.Equal("true", (string?)box.Attribute("aria-hidden"));
        Assert.Contains("width: 144pt", (string?)box.Attribute("style"));
        Assert.Contains("height: 72pt", (string?)box.Attribute("style"));
    }

    [Fact]
    public void UndrawnGeometry_ReservesItsExtentAsAnUnrenderedBox() =>
        AssertUnrenderedBox(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("cloud")))));

    [Fact]
    public void InlineDrawingWithUnrecognizedContent_ReservesItsExtentAsAnUnrenderedBox() =>
        AssertUnrenderedBox(MiddleParagraph(Docx(InlineDrawing("urn:example:unknown", "<unknown xmlns=\"urn:example:unknown\"/>"))));

    [Theory]
    [InlineData("cloud", "[SHAPE]")]
    [InlineData("missing-picture", "[UNSUPPORTED IMAGE]")]
    public void WithPlaceholdersOn_AnUndrawableDrawingIsTheStandardPlaceholderAtItsExtent(string kind, string label)
    {
        var drawing = kind == "cloud" ? InlineDrawing(ShapeUri, Shape("cloud")) : MissingPicture;
        var placeholder = Assert.Single(MiddleParagraph(Docx(drawing), placeholders: true).Descendants(),
            e => (string?)e.Attribute("data-docx-shape") == "unrendered");

        // data-content-type is what the PDF export inventories as unsupported content.
        Assert.Equal("Other", (string?)placeholder.Attribute("data-content-type"));
        Assert.Contains("unsupported-placeholder", (string?)placeholder.Attribute("class"));
        Assert.Equal(label, placeholder.Value);
        Assert.Contains("width: 144pt", (string?)placeholder.Attribute("style"));
        Assert.Contains("box-sizing: border-box", (string?)placeholder.Attribute("style"));
    }

    [Fact]
    public void GroupWithAnUndrawableChild_DrawsTheRestAndIsMarkedPartial()
    {
        var group = Group(Shape() + Shape("cloud"));
        var plain = OnlySvg(MiddleParagraph(Docx(InlineDrawing(GroupUri, group)))).Parent!;
        var flagged = OnlySvg(MiddleParagraph(Docx(InlineDrawing(GroupUri, group)), placeholders: true)).Parent!;

        Assert.Single(Drawn(plain, "rect"));
        Assert.Equal("partial", (string?)plain.Attribute("data-docx-shape"));
        Assert.Null(plain.Attribute("data-content-type"));
        Assert.Equal("Other", (string?)flagged.Attribute("data-content-type"));
    }

    [Theory]
    [InlineData("<a:solidFill><a:prstClr val=\"red\"/></a:solidFill>", true)]
    [InlineData("<a:blipFill><a:stretch><a:fillRect/></a:stretch></a:blipFill>", true)]
    [InlineData(Teal, false)]
    public void UndrawableShapePaint_IsReportedAsPartial(string fill, bool partial)
    {
        var source = Docx(InlineDrawing(ShapeUri, Shape(fill: fill)));
        using (var stream = new MemoryStream(source))
        using (var document = WordprocessingDocument.Open(stream, false))
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(document));
        var wrapper = OnlySvg(MiddleParagraph(source, placeholders: true)).Parent!;

        Assert.Equal(partial ? "partial" : null, (string?)wrapper.Attribute("data-docx-shape"));
        Assert.Equal(partial ? "Other" : null, (string?)wrapper.Attribute("data-content-type"));
    }

    [Fact]
    public void TranslucentAndGroupFills_AreDrawn()
    {
        var translucent = "<a:solidFill><a:srgbClr val=\"19647E\"><a:alpha val=\"40000\"/></a:srgbClr></a:solidFill>";
        var group = Group(Shape(fill: translucent) + Shape(fill: "<a:grpFill/>"))
            .Replace("</a:xfrm></wpg:grpSpPr>", "</a:xfrm><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></wpg:grpSpPr>");
        var rects = Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(GroupUri, group)))), "rect");

        Assert.Equal(("#19647E", "0.4"), ((string?)rects[0].Attribute("fill"), (string?)rects[0].Attribute("fill-opacity")));
        Assert.Equal("#FF0000", (string?)rects[1].Attribute("fill"));
    }

    [Theory]
    [InlineData("<a:scrgbClr r=\"100000\" g=\"21586\" b=\"0\"/>", "#FF8000")]
    [InlineData("<a:hslClr hue=\"1800000\" sat=\"100000\" lum=\"50000\"/>", "#FF8000")]
    public void LinearRgbAndHslColors_Resolve(string color, string expected)
    {
        var rect = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri,
            Shape(fill: $"<a:solidFill>{color}</a:solidFill>"))))), "rect"));
        Assert.Equal(expected, (string?)rect.Attribute("fill"));
    }

    [Fact]
    public void DrawingWithNoDescriptionOrName_IsHiddenFromAssistiveTechnology()
    {
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape()).Replace(" name=\"Teal inspection panel\"", ""))));
        Assert.Equal("true", (string?)svg.Attribute("aria-hidden"));
        Assert.Null(svg.Attribute("role"));
    }

    /// <summary>An inline picture whose image relationship does not exist.</summary>
    private static readonly string MissingPicture = InlineDrawing("http://schemas.openxmlformats.org/drawingml/2006/picture",
        "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"p\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        "<pic:blipFill><a:blip r:embed=\"rIdMissing\"/></pic:blipFill><pic:spPr/></pic:pic>");

    [Fact]
    public void PictureWhoseImageIsMissing_ReservesItsExtentAsAnUnrenderedBox() =>
        AssertUnrenderedBox(MiddleParagraph(Docx(MissingPicture)));

    [Fact]
    public void DrawingInAnUnselectedChoice_IsNotParagraphContent()
    {
        // The Choice needs a namespace the converter does not render, so the (empty) VML Fallback
        // is selected; the Choice's drawing must not count as content, or the paragraph has no line.
        var alternate = "<w:r><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                        "xmlns:x=\"urn:example:unsupported\"><mc:Choice Requires=\"x\">" +
                        InlineDrawing(ShapeUri, Shape()).Replace("<w:r>", "").Replace("</w:r>", "") +
                        "</mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>";
        var paragraph = MiddleParagraph(Docx(alternate));

        Assert.DoesNotContain(paragraph.Descendants(), e => e.Name.LocalName == "svg");
        Assert.NotEmpty(paragraph.Value);
        Assert.True(string.IsNullOrWhiteSpace(paragraph.Value));
    }

    [Fact]
    public void PictureWithNoImageHandler_StaysOmitted()
    {
        // The caller passing no image handler has asked for pictures to be left out.
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("p.docx", Docx(MissingPicture)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        Assert.DoesNotContain(html.Descendants(), e => e.Name.LocalName == "svg");
    }

    [Theory]
    [InlineData(PaginationMode.None, false, false)]
    [InlineData(PaginationMode.None, false, true)]
    [InlineData(PaginationMode.None, true, false)]
    [InlineData(PaginationMode.None, true, true)]
    [InlineData(PaginationMode.Paginated, false, false)]
    [InlineData(PaginationMode.Paginated, false, true)]
    [InlineData(PaginationMode.Paginated, true, false)]
    [InlineData(PaginationMode.Paginated, true, true)]
    public void Shape_and_group_keep_ink_and_anchor_geometry(PaginationMode mode, bool grouped, bool floating)
    {
        var content = grouped ? Group(Shape()) : Shape();
        var uri = grouped ? GroupUri : ShapeUri;
        var input = Docx(floating ? FloatingDrawing(uri, content) : InlineDrawing(uri, content));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, mode);
        var svg = OnlySvg(paragraph);

        Assert.Equal("#19647E", (string?)Assert.Single(Drawn(svg, "rect")).Attribute("fill"));
        Assert.Equal("192", (string?)svg.Attribute("width"));
        Assert.Equal("96", (string?)svg.Attribute("height"));
        if (floating)
        {
            var anchor = Assert.Single(svg.Ancestors().Where(e => (string?)e.Attribute("data-docx-drawing-anchor") == "true"));
            Assert.Equal("144", (string?)anchor.Attribute("data-docx-anchor-extent-width"));
            Assert.Equal("72", (string?)anchor.Attribute("data-docx-anchor-extent-height"));
            Assert.Equal("page", (string?)anchor.Attribute("data-docx-anchor-h-relative"));
            Assert.Equal("360", (string?)anchor.Attribute("data-docx-anchor-h-offset"));
            Assert.Equal("page", (string?)anchor.Attribute("data-docx-anchor-v-relative"));
            Assert.Equal("360", (string?)anchor.Attribute("data-docx-anchor-v-offset"));
        }
        else Assert.DoesNotContain(svg.Ancestors(), e => e.Attribute("data-docx-drawing-anchor") != null);
    }

    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void An_unsupported_floating_shape_is_reported(PaginationMode mode)
    {
        var input = Docx(FloatingDrawing(ShapeUri, Shape("pentagon")));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, mode, placeholders: true);

        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-docx-drawing-anchor") == "true");
    }

    [Theory]
    [InlineData(PaginationMode.None)]
    [InlineData(PaginationMode.Paginated)]
    public void A_plain_Vml_shape_is_drawn_or_reported(PaginationMode mode)
    {
        var input = Docx("<w:r><w:pict><v:rect xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
            "style=\"width:144pt;height:72pt\" fillcolor=\"#19647E\"/></w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, mode, placeholders: true);

        Assert.Contains(paragraph.Descendants(), e => e.Name.LocalName == "svg" || e.Attribute("data-content-type") != null);
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(doc)
            .Select(e => e.Description));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_floating_hexagon_draws_its_six_preset_vertices(bool grouped)
    {
        var input = Docx(FloatingDrawing(grouped ? GroupUri : ShapeUri,
            grouped ? Group(Shape("hexagon")) : Shape("hexagon")));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var points = PolygonPoints(OnlySvg(paragraph));
        Assert.Equal(6, points.Length);
        Assert.Equal(new[] { 0d, 18, 126, 144, 126, 18 }, points.Select(p => p.X).ToArray());
        var expectedY = new[] { 36d, 0, 0, 36, 72, 72 };
        for (var i = 0; i < points.Length; i++) Assert.Equal(expectedY[i], points[i].Y, 4);
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
    }

    [Fact]
    public void A_hexagon_keeps_its_source_adjustment_under_nonuniform_group_scaling()
    {
        var input = Docx(InlineDrawing(GroupUri, Group(Shape("hexagon", cx: 914400, cy: 914400),
            childCx: 914400, childCy: 914400)));
        AssertSchemaValid(input);
        // A square child has an 18pt inset; the group doubles its horizontal coordinates only.
        var points = PolygonPoints(OnlySvg(MiddleParagraph(input)));
        Assert.Equal(36, points[1].X);
        Assert.Equal(108, points[2].X);
    }

    [Fact]
    public void A_hexagon_honors_its_adjustment_and_vertical_factor()
    {
        var shape = Shape("hexagon").Replace("<a:avLst/>",
            "<a:avLst><a:gd name=\"adj\" fmla=\"val 50000\"/><a:gd name=\"vf\" fmla=\"val 100000\"/></a:avLst>");
        var input = Docx(InlineDrawing(ShapeUri, shape));
        AssertSchemaValid(input);
        var points = PolygonPoints(OnlySvg(MiddleParagraph(input)));
        Assert.Equal(36, points[1].X);
        Assert.Equal(36 - 36 * System.Math.Sin(System.Math.PI / 3), points[1].Y, 4);
    }

    private static (double X, double Y)[] PolygonPoints(XElement svg) =>
        Assert.Single(Drawn(svg, "polygon")).Attribute("points")!.Value.Split(' ')
            .Select(pair => pair.Split(','))
            .Select(point => (double.Parse(point[0], System.Globalization.CultureInfo.InvariantCulture) / 12700,
                double.Parse(point[1], System.Globalization.CultureInfo.InvariantCulture) / 12700)).ToArray();

    [Fact]
    public void A_hexagon_with_an_unsupported_adjustment_formula_is_reported()
    {
        var shape = Shape("hexagon").Replace("<a:avLst/>",
            "<a:avLst><a:gd name=\"adj\" fmla=\"*/ w 1 4\"/></a:avLst>");
        var input = Docx(InlineDrawing(ShapeUri, shape));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, placeholders: true);
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
    }

    [Theory]
    [InlineData("", "#ffffff", "#000000", "9525")]
    [InlineData("fillcolor=\"#19647E\"", "#19647E", "#000000", "9525")]
    [InlineData("filled=\"f\" stroked=\"f\"", "none", "none", null)]
    [InlineData("fillcolor=\"red\" strokecolor=\"blue\" strokeweight=\"2pt\"", "red", "blue", "25400")]
    public void A_plain_Vml_rectangle_draws_its_declared_paint(string attributes, string fill, string stroke, string? width)
    {
        var input = Docx("<w:r><w:pict><v:rect xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
            "style=\"width:144pt;height:72pt\" " + attributes + "/></w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var rect = Assert.Single(Drawn(OnlySvg(paragraph), "rect"));
        Assert.Equal(fill, (string?)rect.Attribute("fill"));
        Assert.Equal(stroke, (string?)rect.Attribute("stroke"));
        Assert.Equal(width, (string?)rect.Attribute("stroke-width"));
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
    }

    [Theory]
    [InlineData("0,0", "100pt,0", 1270000, 0, null, "9525")]
    [InlineData("0,0", "0,60pt", 0, 762000, "2pt", "25400")]
    [InlineData("10pt,10pt", "80pt,35pt", 889000, 317500, "2pt", "25400")]
    [InlineData("-10pt,-5pt", "60pt,20pt", 889000, 317500, "2pt", "25400")]
    [InlineData("0,0", "96,48", 914400, 457200, "2pt", "25400")]
    [InlineData("0,0", "1in,6pc", 914400, 914400, "25400", "25400")]
    [InlineData("0,0", "2.54cm,25.4mm", 914400, 914400, "2px", "19050")]
    public void InlineVmlLine_DrawsItsCoordinateSpanAndStroke(string from, string to, long width, long height,
        string? strokeWeight, string expectedStroke)
    {
        var weight = strokeWeight == null ? "" : $" strokeweight=\"{strokeWeight}\"";
        var input = Docx($"<w:r><w:pict><v:line xmlns:v=\"urn:schemas-microsoft-com:vml\" from=\"{from}\" "
            + $"to=\"{to}\" strokecolor=\"#16777A\"{weight}/></w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var svg = OnlySvg(paragraph);
        var line = Assert.Single(Drawn(svg, "line"));
        Assert.Equal("0", (string?)line.Attribute("x1"));
        Assert.Equal("0", (string?)line.Attribute("y1"));
        Assert.Equal(width.ToString(CultureInfo.InvariantCulture), (string?)line.Attribute("x2"));
        Assert.Equal(height.ToString(CultureInfo.InvariantCulture), (string?)line.Attribute("y2"));
        Assert.Equal("none", (string?)line.Attribute("fill"));
        Assert.Equal("#16777A", (string?)line.Attribute("stroke"));
        Assert.Equal(expectedStroke, (string?)line.Attribute("stroke-width"));
        Assert.Null(svg.Attribute("viewBox"));
        var style = (string?)svg.Parent!.Attribute("style");
        Assert.Contains("margin-left: 9pt", style);
        Assert.Contains("margin-right: 9pt", style);
        Assert.Contains("width: " + (width / 12700d).ToString("0.###", CultureInfo.InvariantCulture) + "pt", style);
        Assert.Contains("height: " + (height / 12700d).ToString("0.###", CultureInfo.InvariantCulture) + "pt", style);
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
    }

    [Theory]
    [InlineData("", "0", "0", "635000", "635000", "#000000")]
    [InlineData("from=\"10pt,20pt\"", "0", "0", "508000", "381000", "#000000")]
    [InlineData("from=\"80pt,35pt\" to=\"10pt,10pt\"", "889000", "317500", "0", "0", "#000000")]
    [InlineData("from=\"0,25pt\" to=\"70pt,0\"", "0", "317500", "889000", "0", "#000000")]
    [InlineData("to=\"70pt,25pt\" stroked=\"f\"", "0", "0", "889000", "317500", "none")]
    [InlineData("to=\"70pt,25pt\" style=\"width:1pt;height:2pt\" filled=\"t\" fillcolor=\"red\"", "0", "0", "889000", "317500", "#000000")]
    public void InlineVmlLine_PreservesDefaultsDirectionAndHiddenPaint(string attributes,
        string x1, string y1, string x2, string y2, string stroke)
    {
        var input = Docx("<w:r><w:pict><v:line xmlns:v=\"urn:schemas-microsoft-com:vml\" " + attributes + "/></w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var svg = OnlySvg(paragraph);
        var line = Assert.Single(Drawn(svg, "line"));
        Assert.Equal(new[] { x1, y1, x2, y2 }, new[] { "x1", "y1", "x2", "y2" }.Select(a => (string?)line.Attribute(a)));
        Assert.Equal("none", (string?)line.Attribute("fill"));
        Assert.Equal(stroke, (string?)line.Attribute("stroke"));
        Assert.Equal(stroke == "none" ? null : "9525", (string?)line.Attribute("stroke-width"));
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
    }

    [Theory]
    [InlineData("<v:line from=\"0,0\" to=\"0,0\"/>")]
    [InlineData("<v:line to=\"100pt,0\" strokeweight=\"0\"/>")]
    [InlineData("<v:line from=\"0,0\" to=\"100pt,0\"><v:stroke dashstyle=\"dash\"/></v:line>")]
    [InlineData("<v:line from=\"0,0\" to=\"100pt,0\" style=\"position:absolute;left:20pt;top:30pt\"/>")]
    [InlineData("<v:group style=\"width:100pt;height:50pt\" coordsize=\"100,50\"><v:line from=\"0,0\" to=\"100,50\"/></v:group>")]
    [InlineData("<v:line to=\"100pt,0\"/><v:line to=\"0,50pt\"/>")]
    public void VmlLine_UnsupportedPlacementOrPaintRetainsItsDiagnostic(string source)
    {
        var input = Docx("<w:r><w:pict xmlns:v=\"urn:schemas-microsoft-com:vml\">" + source + "</w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Name.LocalName == "line");
    }

    [Theory]
    [InlineData("to=\"NaN,0\"")]
    [InlineData("to=\"1e300pt,0\"")]
    [InlineData("to=\"1%,0\"")]
    [InlineData("to=\"1,2,3\"")]
    [InlineData("to=\"100pt,0\" strokeweight=\"-1pt\"")]
    public void VmlLine_InvalidGeometryOrStrokeRetainsItsDiagnostic(string attributes)
    {
        var input = Docx("<w:r><w:pict><v:line xmlns:v=\"urn:schemas-microsoft-com:vml\" " + attributes + "/></w:pict></w:r>");
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Name.LocalName == "line");
    }

    [Theory]
    [InlineData("pict", "", true)]
    [InlineData("object", "", true)]
    [InlineData("pict", "<w:r><w:t>Adjacent</w:t></w:r>", false)]
    [InlineData("pict", "<w:r><w:br/></w:r>", false)]
    [InlineData("pict", "<w:r><w:cr/></w:r>", false)]
    [InlineData("pict", "<w:r><w:footnoteReference w:id=\"1\"/></w:r>", false)]
    [InlineData("pict", "<w:r><w:t>\u00a0</w:t></w:r>", false)]
    [InlineData("pict", "<w:r><w:t/></w:r>", true)]
    [InlineData("pict", "<w:r><w:rPr><w:b/></w:rPr><w:t/></w:r>", true)]
    [InlineData("pict", "<w:hyperlink w:anchor=\"Line\"/>", true)]
    [InlineData("pict", "<w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t></w:r>", true)]
    [InlineData("pict", "<w:bookmarkStart w:id=\"3\" w:name=\"Line\"/><w:bookmarkEnd w:id=\"3\"/>", true)]
    [InlineData("pict", "<w:r><w:pict><v:rect xmlns:v=\"urn:schemas-microsoft-com:vml\" style=\"width:10pt;height:10pt\"/></w:pict></w:r>", false)]
    [InlineData("pict", "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" style=\"width:20pt;height:30pt\" fillcolor=\"red\"><v:textbox><w:txbxContent><w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>", false)]
    public void VmlLine_OnlyASoleVisibleObjectAlignsToTheParagraphTop(string host, string adjacent, bool sole)
    {
        var input = Docx(adjacent + $"<w:r><w:{host}><v:line xmlns:v=\"urn:schemas-microsoft-com:vml\" "
            + $"to=\"0,60pt\"/></w:{host}></w:r>");
        AssertSchemaValid(input);
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(input, new HtmlConversionOptions
        {
            FabricateCssClasses = false, PaginationMode = (int)PaginationMode.Paginated,
            RenderUnsupportedContentPlaceholders = true, RenderFootnotesAndEndnotes = true,
        }));
        var line = Assert.Single(html.Descendants(), e => e.Name.LocalName == "line");
        var style = (string?)line.Ancestors().First(e => e.Name.LocalName == "svg").Parent!.Attribute("style");
        if (sole) Assert.Contains("vertical-align: top", style);
        else Assert.DoesNotContain("vertical-align: top", style);
    }

    [Fact]
    public void Vml_lengths_keep_the_Css_decimal_point_under_a_comma_decimal_culture()
    {
        var previous = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("fr-FR");
            var input = Docx("<w:r><w:pict><v:rect xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "style=\"width:144.5pt;height:72pt\"/></w:pict></w:r>");
            AssertSchemaValid(input);
            var rect = Assert.Single(Drawn(OnlySvg(MiddleParagraph(input)), "rect"));
            Assert.Equal("1835150", (string?)rect.Attribute("width"));
        }
        finally { System.Globalization.CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void Positioned_Vml_is_reported_instead_of_becoming_an_inline_rectangle()
    {
        var input = Docx("<w:r><w:pict><v:rect xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
            "style=\"position:absolute;left:36pt;top:24pt;width:144pt;height:72pt\"/></w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, placeholders: true);
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Name.LocalName == "svg");
    }

    [Fact]
    public void A_simple_position_uses_its_page_coordinates_instead_of_the_unused_axes()
    {
        var input = Docx(FloatingDrawing(ShapeUri, Shape()).Replace("simplePos=\"0\"", "simplePos=\"1\"")
            .Replace("<wp:simplePos x=\"0\" y=\"0\"/>", "<wp:simplePos x=\"914400\" y=\"1828800\"/>"));
        AssertSchemaValid(input);
        var anchor = OnlySvg(MiddleParagraph(input, PaginationMode.Paginated)).Parent!;
        Assert.Equal("72", (string?)anchor.Attribute("data-docx-anchor-h-offset"));
        Assert.Equal("144", (string?)anchor.Attribute("data-docx-anchor-v-offset"));
        Assert.Equal("page", (string?)anchor.Attribute("data-docx-anchor-h-relative"));
    }

    private static string ParagraphShape(string wrap = "<wp:wrapTopAndBottom/>") =>
        FloatingDrawing(ShapeUri, Shape())
            .Replace("<wp:positionH relativeFrom=\"page\">", "<wp:positionH relativeFrom=\"column\">")
            .Replace("<wp:positionV relativeFrom=\"page\">", "<wp:positionV relativeFrom=\"paragraph\">")
            .Replace("4572000", "0").Replace("<wp:wrapNone/>", wrap);

    [Theory]
    [InlineData(0, "72")]
    [InlineData(24, "96")]
    public void Overlapping_top_bottom_shapes_reserve_their_union(int secondOffset, string expectedHeight)
    {
        var second = ParagraphShape().Replace("id=\"7\"", "id=\"8\"")
            .Replace("<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset>",
                $"<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>{secondOffset * 12700}</wp:posOffset>");
        var input = Docx(ParagraphShape() + second);
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var clearance = Assert.Single(paragraph.Descendants().Where(e => e.Attribute("data-docx-shape-clearance") != null));
        Assert.Contains("height: " + expectedHeight + "pt", (string?)clearance.Attribute("style"));
        Assert.DoesNotContain(paragraph.Descendants(), e => e.Attribute("data-content-type") != null);
    }

    [Fact]
    public void Interior_wrapping_is_reported_while_the_shape_keeps_its_anchored_geometry()
    {
        var input = Docx(ParagraphShape("<wp:wrapSquare wrapText=\"bothSides\"/>")
            .Replace("<wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset>",
                "<wp:positionH relativeFrom=\"column\"><wp:posOffset>1270000</wp:posOffset>"));
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var anchor = OnlySvg(paragraph).Parent!;
        Assert.Equal("Other", (string?)anchor.Attribute("data-content-type"));
        Assert.Equal("100", (string?)anchor.Attribute("data-docx-anchor-h-offset"));
    }

    [Fact]
    public void A_floating_group_preserves_its_shape_and_textbox_geometry()
    {
        var textBox = Shape().Replace("<wps:bodyPr/>",
            "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Group caption</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>");
        var input = Docx(FloatingDrawing(GroupUri, Group(Shape() + textBox)));
        AssertSchemaValid(input);
        var html = HtmlConversionOps.ConvertToHtml(input, new HtmlConversionOptions
        {
            PaginationMode = (int)PaginationMode.Paginated,
            RenderUnsupportedContentPlaceholders = true,
        });
        var document = XElement.Parse(html);
        Assert.Contains("Group caption", document.Value);
        Assert.Equal(2, document.Descendants().Count(e => e.Name.LocalName == "rect"));
        Assert.DoesNotContain(document.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
    }

    [Fact]
    public void A_Vml_fallback_does_not_invent_ink_from_disabled_fill_or_outline()
    {
        var input = Docx("<w:r><w:pict><v:oval xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
            "style=\"width:144pt;height:72pt\" fillcolor=\"#19647E\" strokecolor=\"#FF9900\" filled=\"f\" stroked=\"f\"/></w:pict></w:r>");
        AssertSchemaValid(input);
        var paragraph = MiddleParagraph(input, placeholders: true);
        var placeholder = Assert.Single(paragraph.Descendants().Where(e => e.Attribute("data-content-type") != null));
        Assert.Contains("width: 144pt", (string?)placeholder.Attribute("style"));
        Assert.Contains("height: 72pt", (string?)placeholder.Attribute("style"));
        Assert.DoesNotContain("#19647E", placeholder.ToString());
        Assert.DoesNotContain("#FF9900", placeholder.ToString());
    }

    [Theory]
    [InlineData(PaginationMode.Paginated)]
    [InlineData(PaginationMode.None)]
    public void FloatingShape_RetainsItsParagraphMarkAlongsideItsGraphic(PaginationMode mode)
    {
        // A floating drawing is not paragraph content (issue #880), so nothing may take its line.
        var paragraph = MiddleParagraph(Docx(FloatingDrawing(ShapeUri, Shape())), mode);

        Assert.Single(paragraph.Descendants().Where(e => e.Name.LocalName == "svg"));
        Assert.NotEmpty(paragraph.Value);
        Assert.True(string.IsNullOrWhiteSpace(paragraph.Value));
    }
    [Theory]
    [InlineData(true, "oval")]
    [InlineData(true, "rect")]
    [InlineData(false, "oval")]
    public void LegacyShapeFallbackUsesSupportedPageCoordinates(bool positioned, string kind)
    {
        var placement = positioned ? "position:absolute;margin-left:8pt;margin-top:12pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;" : "";
        var input = Docx("<w:r><w:pict><v:" + kind + " xmlns:v=\"urn:schemas-microsoft-com:vml\" style=\"" +
            placement + "width:90pt;height:72pt;z-index:-1\"><v:fill color=\"#ddeedd\"/></v:" + kind + "></w:pict></w:r>");
        var paragraph = MiddleParagraph(input, PaginationMode.Paginated, placeholders: true);
        var shape = paragraph.Descendants().Single(e => (string?)e.Attribute("data-docx-shape") == "unrendered");
        Assert.Equal(positioned ? "true" : null, (string?)shape.Attribute("data-docx-drawing-anchor"));
        Assert.Contains("[SHAPE]", shape.Value);
        if (positioned)
        {
            Assert.Equal("8", (string?)shape.Attribute("data-docx-anchor-h-offset"));
            Assert.Equal("12", (string?)shape.Attribute("data-docx-anchor-v-offset"));
            Assert.Equal("true", (string?)shape.Attribute("data-docx-anchor-behind"));
        }
    }

}
