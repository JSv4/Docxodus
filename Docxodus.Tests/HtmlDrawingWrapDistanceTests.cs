using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlDrawingWrapDistanceTests
{
    private static byte[] Drawing(long rawDistance, decimal vmlDistance, bool overrideWrap)
    {
        // The two independently authored alternatives describe the same 72pt horizontal rule.
        // A signed -0.2pt clearance is -2540 EMU, or 4294964756 when stored as uint32.
        var fallbackDistance = vmlDistance.ToString(CultureInfo.InvariantCulture);
        var xml = $"""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
              xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
              xmlns:v="urn:schemas-microsoft-com:vml"><w:body>
              <w:p><w:r><w:t>Willow rule inventory.</w:t></w:r></w:p>
              <w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
                <wp:anchor distT="{(overrideWrap ? 127000 : rawDistance)}" distB="{(overrideWrap ? 127000 : rawDistance)}"
                  distL="25400" distR="38100" simplePos="0" relativeHeight="4" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                  <wp:simplePos x="0" y="0"/>
                  <wp:positionH relativeFrom="column"><wp:posOffset>25400</wp:posOffset></wp:positionH>
                  <wp:positionV relativeFrom="paragraph"><wp:posOffset>127000</wp:posOffset></wp:positionV>
                  <wp:extent cx="914400" cy="0"/>
                  <wp:wrapTopAndBottom {(overrideWrap ? $"distT=\"{rawDistance}\" distB=\"{rawDistance}\"" : "")}/>
                  <wp:docPr id="4" name="Willow horizontal rule"/>
                  <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm>
                    <a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="315E42"/></a:solidFill></a:ln>
                    </wps:spPr><wps:bodyPr/></wps:wsp>
                  </a:graphicData></a:graphic>
                </wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict>
                  <v:line from="2pt,10pt" to="74pt,10pt" style="position:absolute;mso-wrap-distance-top:{fallbackDistance}pt;mso-wrap-distance-bottom:{fallbackDistance}pt"/>
                </w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
              <w:p><w:r><w:t>Willow rule complete.</w:t></w:r></w:p><w:sectPr/>
            </w:body></w:document>
            """;
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write))) writer.Write(xml);
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            package.Save();
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(4294964756L, -0.2, -0.2, false)]
    [InlineData(4294964756L, -0.2, -0.2, true)]
    [InlineData(-2540L, -0.2, -0.2, false)]
    [InlineData(-2540L, -0.2, -0.2, true)]
    [InlineData(25400L, 2.0, 2.0, false)]
    [InlineData(25400L, 2.0, 2.0, true)]
    [InlineData(4294967294L, -0.0001574803, -0.0002, false)]
    [InlineData(-2L, -0.0001574803, -0.0002, false)]
    [InlineData(0L, 0.0, 0.0, false)]
    public void DrawingClearanceMatchesItsSignedVmlAlternative(long rawDistance, double vmlDistance,
        double expectedDistance, bool overrideWrap)
    {
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(Drawing(rawDistance, (decimal)vmlDistance, overrideWrap),
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = (int)PaginationMode.Paginated }));
        var anchor = html.Descendants().Single(e => (string?)e.Attribute("data-docx-shape-anchor") == "true");
        Assert.Equal(expectedDistance, double.Parse(anchor.Attribute("data-docx-anchor-wrap-top")!.Value, CultureInfo.InvariantCulture));
        Assert.Equal(expectedDistance, double.Parse(anchor.Attribute("data-docx-anchor-wrap-bottom")!.Value, CultureInfo.InvariantCulture));
        Assert.Equal("2", anchor.Attribute("data-docx-anchor-wrap-left")!.Value);
        Assert.Equal("3", anchor.Attribute("data-docx-anchor-wrap-right")!.Value);
        Assert.Equal("2", anchor.Attribute("data-docx-anchor-h-offset")!.Value);
        Assert.Equal("10", anchor.Attribute("data-docx-anchor-v-offset")!.Value);
        Assert.Equal("72", anchor.Attribute("data-docx-anchor-extent-width")!.Value);
        Assert.Equal("0", anchor.Attribute("data-docx-anchor-extent-height")!.Value);
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "svg");
        var clearance = html.Descendants().Single(e => (string?)e.Attribute("data-docx-shape-clearance") == "true");
        Assert.Contains($"height: {(10 + expectedDistance).ToString(CultureInfo.InvariantCulture)}pt", clearance.Attribute("style")!.Value);
        Assert.Contains("Willow rule inventory.", html.Value);
        Assert.Contains("Willow rule complete.", html.Value);
    }
    [Theory]
    [InlineData("wrapSquare", 4294964756L, "-0.2")]
    [InlineData("wrapSquare", 25400L, "2")]
    [InlineData("wrapTight", 4294964756L, "-0.2")]
    [InlineData("wrapTight", 25400L, "2")]
    [InlineData("wrapThrough", 4294964756L, "-0.2")]
    [InlineData("wrapThrough", 25400L, "2")]
    public void SideClearanceOverridesUseTheSameSignedDistance(string wrapKind, long rawDistance, string expected)
    {
        using var stream = new MemoryStream();
        stream.Write(Drawing(rawDistance, 0m, false));
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var xml = package.MainDocumentPart!.GetXDocument();
            var anchor = xml.Descendants(WP.anchor).Single();
            anchor.Element(WP.positionH)!.ReplaceNodes(new XElement(WP.align, "left"));
            anchor.Element(WP.positionV)!.ReplaceNodes(new XElement(WP.posOffset, 0));
            anchor.Element(WP.wrapTopAndBottom)!.ReplaceWith(new XElement(XName.Get(wrapKind, WP.anchor.NamespaceName),
                new XAttribute("distL", rawDistance), new XAttribute("distR", rawDistance),
                new XAttribute("wrapText", "right")));
            package.MainDocumentPart.PutXDocument();
        }
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = (int)PaginationMode.Paginated }));
        var anchorHtml = html.Descendants().Single(e => (string?)e.Attribute("data-docx-shape-anchor") == "true");
        Assert.Equal(expected, anchorHtml.Attribute("data-docx-anchor-wrap-left")!.Value);
        Assert.Equal(expected, anchorHtml.Attribute("data-docx-anchor-wrap-right")!.Value);
        var exclusion = html.Descendants().Single(e => (string?)e.Attribute("data-docx-shape-exclusion") == wrapKind);
        Assert.Contains($"margin-right: {expected}pt", exclusion.Attribute("style")!.Value);
        Assert.Contains($"margin-top: {expected}pt", exclusion.Attribute("style")!.Value);
        Assert.Contains($"margin-bottom: {expected}pt", exclusion.Attribute("style")!.Value);
        Assert.Contains("margin-left: 0pt", exclusion.Attribute("style")!.Value);
    }

    [Theory]
    [InlineData(2147483647L, 2147483647L)]
    [InlineData(2147483648L, -2147483648L)]
    [InlineData(4294967295L, -1L)]
    [InlineData(4294967296L, 4294967296L)]
    public void OnlyTheUnsigned32BitNegativeRangeIsDecoded(long stored, long signed)
    {
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(Drawing(stored, 0m, false),
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = (int)PaginationMode.Paginated }));
        var anchor = html.Descendants().Single(e => (string?)e.Attribute("data-docx-shape-anchor") == "true");
        var expected = (signed / 12700m).ToString("0.####", CultureInfo.InvariantCulture);
        Assert.Equal(expected, anchor.Attribute("data-docx-anchor-wrap-top")!.Value);
        Assert.Equal(expected, anchor.Attribute("data-docx-anchor-wrap-bottom")!.Value);
        Assert.Equal("2", anchor.Attribute("data-docx-anchor-h-offset")!.Value);
        Assert.Equal("72", anchor.Attribute("data-docx-anchor-extent-width")!.Value);
    }

}
