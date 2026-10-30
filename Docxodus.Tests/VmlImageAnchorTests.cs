using System.Globalization;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using SkiaSharp;
using Xunit;

namespace Docxodus.Tests;

public class VmlImageAnchorTests
{
    private const string BackgroundStyle = "position:absolute;width:640pt;height:820pt;z-index:-37;" +
        "margin-left:0;margin-top:0;mso-position-horizontal:center;mso-position-horizontal-relative:margin;" +
        "mso-position-vertical:center;mso-position-vertical-relative:margin";

    private static XElement Render(string style, bool paginated = true, bool precedingShape = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>("header");
            var imagePart = header.AddImagePart(ImagePartType.Png, "picture");
            using var bitmap = new SKBitmap(3, 2);
            bitmap.Erase(new SKColor(41, 117, 137));
            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            using var imageBytes = new MemoryStream(data.ToArray());
            imagePart.FeedData(imageBytes);
            header.PutXDocument(XDocument.Parse($"""
                <w:hdr xmlns:w="{W.w}" xmlns:r="{R.r}" xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:p><w:r><w:t>Oriole background ledger.</w:t></w:r><w:r><w:pict>
                    {(precedingShape ? "<v:shape id=\"EmptyShape\" style=\"width:11pt;height:9pt\"/>" : "")}
                    <v:shape id="OrioleBackdrop" style="{style}"><v:imagedata r:id="picture"/></v:shape>
                  </w:pict></w:r></w:p>
                </w:hdr>
                """));
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}" xmlns:r="{R.r}"><w:body>
                  <w:p><w:r><w:t>Oriole body record.</w:t></w:r></w:p>
                  <w:sectPr><w:headerReference w:type="default" r:id="header"/>
                    <w:pgSz w:w="12800" w:h="16400"/>
                    <w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200" w:header="600" w:footer="600"/>
                  </w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false, RenderHeadersAndFooters = true,
                PaginationMode = paginated ? 1 : 0 }));
    }

    [Fact]
    public void BackgroundImageKeepsItsPageAnchorWithoutConsumingHeaderFlow()
    {
        var html = Render(BackgroundStyle);
        var frame = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-drawing-anchor") == "true");
        Assert.Equal("true", (string?)frame.Attribute("data-docx-anchor-behind"));
        Assert.Equal("margin", (string?)frame.Attribute("data-docx-anchor-h-relative"));
        Assert.Equal("margin", (string?)frame.Attribute("data-docx-anchor-v-relative"));
        Assert.Equal("center", (string?)frame.Attribute("data-docx-anchor-h-align"));
        Assert.Equal("center", (string?)frame.Attribute("data-docx-anchor-v-align"));
        Assert.Contains("position: absolute", (string?)frame.Attribute("style"));
        Assert.Contains("width: 640pt", (string?)frame.Attribute("style"));
        Assert.Contains("height: 820pt", (string?)frame.Attribute("style"));
        Assert.Single(frame.Descendants(), e => e.Name.LocalName == "img");
        Assert.Contains("Oriole body record.", html.Value);
        Assert.DoesNotContain("mso-position", html.ToString());
    }

    [Fact]
    public void SupportedOffsetsUseTheSharedAnchorCoordinatesInEveryCulture()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            var html = Render("position:absolute;width:137.5pt;height:83.25pt;z-index:53;" +
                "margin-left:17.25pt;margin-top:23.5pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page");
            var frame = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-drawing-anchor") == "true");
            Assert.Equal("false", (string?)frame.Attribute("data-docx-anchor-behind"));
            Assert.Equal("17.25", (string?)frame.Attribute("data-docx-anchor-h-offset"));
            Assert.Equal("23.5", (string?)frame.Attribute("data-docx-anchor-v-offset"));
            Assert.Equal("137.5", (string?)frame.Attribute("data-docx-anchor-extent-width"));
            Assert.Equal("83.25", (string?)frame.Attribute("data-docx-anchor-extent-height"));
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void ContinuousConversionKeepsItsExistingImageFallback()
    {
        var html = Render(BackgroundStyle, paginated: false);
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "img");
    }

    [Fact]
    public void PlacementComesFromTheShapeOwningTheImageRelationship()
    {
        var html = Render(BackgroundStyle, precedingShape: true);
        var frame = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-drawing-anchor") == "true");
        Assert.Equal("640", (string?)frame.Attribute("data-docx-anchor-extent-width"));
        Assert.Equal("820", (string?)frame.Attribute("data-docx-anchor-extent-height"));
    }

    [Theory]
    [InlineData("position:relative;width:137pt;height:83pt")]
    [InlineData("position:absolute;width:137pt;height:83pt;mso-position-horizontal-relative:unknown;mso-position-vertical-relative:page")]
    [InlineData("position:absolute;width:137pt;height:83pt;margin-left:17px;mso-position-horizontal-relative:page;mso-position-vertical-relative:page")]
    [InlineData("position:absolute;width:NaNpt;height:83pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page")]
    [InlineData("position:absolute;width:999999999999999999pt;height:83pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page")]
    [InlineData("position:absolute;width:137pt;height:83pt;z-index:2147483648;mso-position-horizontal-relative:page;mso-position-vertical-relative:page")]
    public void UnsupportedPlacementDoesNotBecomeArbitraryCss(string style)
    {
        var html = Render(style);
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
        Assert.Single(html.Descendants(), e => e.Name.LocalName == "img");
    }
}
