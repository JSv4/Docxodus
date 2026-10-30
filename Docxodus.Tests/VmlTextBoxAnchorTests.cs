using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class VmlTextBoxAnchorTests
{
    private const string Placement = "position:absolute;width:143.5pt;height:78.25pt;z-index:-19;" +
        "margin-left:36.5pt;margin-top:91.25pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:margin";

    private static XElement Render(string style, bool paginated = true, bool autoFit = false,
        bool precedingShape = false, bool wrapsText = false, bool precedingAutoFit = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}" xmlns:v="urn:schemas-microsoft-com:vml"
                  xmlns:w10="urn:schemas-microsoft-com:office:word"><w:body><w:p><w:r><w:pict>
                  {(precedingShape ? "<v:shape style=\"width:7pt;height:6pt\">" +
                    (precedingAutoFit ? "<v:textbox style=\"mso-fit-shape-to-text:true\"/>" : "") + "</v:shape>" : "")}
                  <v:shape style="{style}" fillcolor="#9dcac6">
                    {(wrapsText ? "<w10:wrap type=\"square\"/>" : "")}
                    <v:textbox inset="0,0,0,0" style="{(autoFit ? "mso-fit-shape-to-text:true" : "")}">
                      <w:txbxContent><w:p><w:r><w:t>Plover copper panel</w:t></w:r></w:p></w:txbxContent>
                    </v:textbox>
                  </v:shape>
                </w:pict></w:r></w:p><w:p><w:r><w:t>Plover body control</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w=\"{W.w}\"/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            FabricateCssClasses = false,
            PaginationMode = paginated ? 1 : 0,
        }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AbsoluteTextBoxUsesItsOwningShapesPageCoordinates(bool precedingShape)
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            var html = Render(Placement, precedingShape: precedingShape);
            var box = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-drawing-anchor") != null);
            Assert.Equal("Plover copper panel", box.Value);
            Assert.Equal("true", (string?)box.Attribute("data-docx-anchor-behind"));
            Assert.Equal("page", (string?)box.Attribute("data-docx-anchor-h-relative"));
            Assert.Equal("margin", (string?)box.Attribute("data-docx-anchor-v-relative"));
            Assert.Equal("36.5", (string?)box.Attribute("data-docx-anchor-h-offset"));
            Assert.Equal("91.25", (string?)box.Attribute("data-docx-anchor-v-offset"));
            Assert.Equal("143.5", (string?)box.Attribute("data-docx-anchor-extent-width"));
            Assert.Equal("78.25", (string?)box.Attribute("data-docx-anchor-extent-height"));
            Assert.Contains("position: absolute", (string?)box.Attribute("style"));
            Assert.Contains("background-color: #9dcac6", (string?)box.Attribute("style"));
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void AnAnchoredAutoFitBoxStillUsesItsContentHeight()
    {
        var box = Assert.Single(Render(Placement, autoFit: true).Descendants(),
            e => e.Attribute("data-docx-drawing-anchor") != null);
        Assert.Equal("true", (string?)box.Attribute("data-docx-anchor-autofit"));
        Assert.DoesNotContain(((string?)box.Attribute("style") ?? "").Split(';'),
            declaration => declaration.TrimStart().StartsWith("height:", System.StringComparison.Ordinal));
    }

    [Fact]
    public void AnUnrelatedSiblingsAutoFitDoesNotChangeTheFixedBox()
    {
        var box = Assert.Single(Render(Placement, precedingShape: true, precedingAutoFit: true).Descendants(),
            e => e.Attribute("data-docx-drawing-anchor") != null);
        Assert.Null(box.Attribute("data-docx-anchor-autofit"));
        Assert.Contains("height: 78.25pt", (string?)box.Attribute("style"));
    }

    [Theory]
    [InlineData("width:143.5pt;height:78.25pt")]
    [InlineData("position:absolute;width:143.5pt;height:78.25pt")]
    [InlineData("position:absolute;width:143.5pt;height:78.25pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:page")]
    [InlineData("position:absolute;width:143.5pt;height:78.25pt;margin-left:36px;mso-position-horizontal-relative:page;mso-position-vertical-relative:page")]
    public void InlineOrUnsupportedLegacyCoordinatesKeepTheirExistingFlow(string style)
    {
        var html = Render(style);
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
        Assert.Contains("Plover copper panel", html.Value);
    }

    [Fact]
    public void ForegroundWrappingKeepsItsExistingFallback()
    {
        var html = Render(Placement.Replace("z-index:-19", "z-index:19"), wrapsText: true);
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
        Assert.Contains("Plover copper panel", html.Value);
    }

    [Fact]
    public void ContinuousConversionKeepsItsExistingFlow()
    {
        var html = Render(Placement, paginated: false);
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
        Assert.Contains("Plover copper panel", html.Value);
    }
}
