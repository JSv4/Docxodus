using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlGroupedVmlTextBoxTests
{
    private const string Placement = "position:absolute;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;";

    private static string Box(string id, string text, string style) => $"""
        <v:shape id="{id}" style="{style}" stroked="f"><v:textbox inset="0,0,0,0">
          <w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent>
        </v:textbox></v:shape>
        """;

    private static string Drawing(bool grouped, bool nested = false)
    {
        var first = Box("alder", "ALDER BOX NORTH", grouped
            ? "position:absolute;left:100;top:200;width:240;height:56"
            : Placement + "margin-left:72pt;margin-top:100pt;width:120pt;height:28pt");
        var second = Box("birch", "BIRCH BOX SOUTH", grouped
            ? nested ? "position:absolute;left:10;top:20;width:240;height:56"
                : "position:absolute;left:400;top:300;width:240;height:56"
            : Placement + "margin-left:222pt;margin-top:150pt;width:120pt;height:28pt");
        if (nested)
            second = $"<v:group style=\"position:absolute;left:400;top:300;width:240;height:100\" coordorigin=\"10,20\" coordsize=\"240,100\">{second}</v:group>";
        return grouped
            ? $"<w:pict><v:group id=\"grove\" style=\"{Placement}margin-left:72pt;margin-top:100pt;width:300pt;height:100pt\" coordorigin=\"100,200\" coordsize=\"600,200\">{first}{second}</v:group></w:pict>"
            : $"<w:pict>{first}</w:pict><w:pict>{second}</w:pict>";
    }

    private static byte[] Document(string drawing, bool header)
    {
        var namespaces = $"xmlns:w=\"{W.w}\" xmlns:v=\"{VML.vml}\" xmlns:r=\"{R.r}\" xmlns:mc=\"{MC.mc}\"";
        var carrier = $"<w:p><w:r>{drawing}</w:r></w:p>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w=\"{W.w}\"/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
            if (header)
                main.AddNewPart<HeaderPart>("header").PutXDocument(XDocument.Parse($"<w:hdr {namespaces}>{carrier}</w:hdr>"));
            main.PutXDocument(XDocument.Parse($"""
                <w:document {namespaces}><w:body>
                  <w:p><w:r><w:t>Grove register control</w:t></w:r></w:p>{(header ? "" : carrier)}
                  <w:sectPr>{(header ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "")}
                    <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440" w:header="360" w:footer="360"/>
                  </w:sectPr>
                </w:body></w:document>
                """));
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document));
        }
        return stream.ToArray();
    }

    private static XElement Render(string drawing, bool header = false) => XElement.Parse(
        HtmlConversionOps.ConvertToHtml(Document(drawing, header), new HtmlConversionOptions
        {
            FabricateCssClasses = false, PaginationMode = 1, RenderHeadersAndFooters = true,
            RenderUnsupportedContentPlaceholders = true,
        }));

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void BodyAndHeaderKeepBothIndependentStories(bool header, bool grouped)
    {
        var boxes = Render(Drawing(grouped), header).Descendants()
            .Where(e => (string?)e.Attribute("data-docx-textbox") == "true").ToArray();
        Assert.Equal(new[] { "ALDER BOX NORTH", "BIRCH BOX SOUTH" }, boxes.Select(box => box.Value.Trim()));
        Assert.All(boxes, box => Assert.Contains("width: 120pt", (string?)box.Attribute("style")));
        Assert.All(boxes, box => Assert.Contains("height: 28pt", (string?)box.Attribute("style")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GroupCoordinatesPreserveSiblingOffsetsAndNestedScale(bool nested)
    {
        var html = Render(Drawing(true, nested));
        var boxes = html.Descendants().Where(e => (string?)e.Attribute("data-docx-textbox") == "true").ToArray();
        Assert.Equal(2, boxes.Length);
        Assert.All(boxes, box => Assert.Contains("width: 120pt", (string?)box.Attribute("style")));
        Assert.Contains("left: 0pt", (string?)boxes[0].Attribute("style"));
        Assert.Contains("top: 0pt", (string?)boxes[0].Attribute("style"));
        var secondFrame = nested ? boxes[1].Parent! : boxes[1];
        Assert.Contains("left: 150pt", (string?)secondFrame.Attribute("style"));
        Assert.Contains("top: 50pt", (string?)secondFrame.Attribute("style"));
        var anchor = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-drawing-anchor") != null);
        Assert.Equal("72", (string?)anchor.Attribute("data-docx-anchor-h-offset"));
        Assert.Equal("100", (string?)anchor.Attribute("data-docx-anchor-v-offset"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CompatibilitySelectionKeepsOneGroupAndNoFallbackDuplicates(bool supported)
    {
        var drawing = $"<mc:AlternateContent xmlns:choice=\"{(supported ? VML.vml.NamespaceName : "urn:unsupported")}\"><mc:Choice Requires=\"choice\">{Drawing(true)}</mc:Choice><mc:Fallback>{Drawing(true).Replace("ALDER BOX NORTH", "FALLBACK NORTH").Replace("BIRCH BOX SOUTH", "FALLBACK SOUTH")}</mc:Fallback></mc:AlternateContent>";
        var boxes = Render(drawing).Descendants().Where(e => e.Attribute("data-docx-textbox") != null).ToArray();
        Assert.Equal(supported ? new[] { "ALDER BOX NORTH", "BIRCH BOX SOUTH" } : new[] { "FALLBACK NORTH", "FALLBACK SOUTH" },
            boxes.Select(box => box.Value.Trim()));
    }
}
