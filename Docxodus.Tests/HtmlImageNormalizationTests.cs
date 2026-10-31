using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using SkiaSharp;
using Xunit;

namespace Docxodus.Tests;

public class HtmlImageNormalizationTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void TiffNormalizesToPngInBodyAndHeaderPictures(bool header, bool vml)
    {
        var image = Picture(Render(Tiff(), "image/tiff", header, vml));
        Assert.StartsWith("data:image/png;base64,", (string?)image.Attribute("src"));
        using var bitmap = SKBitmap.Decode(Bytes(image));
        Assert.NotNull(bitmap);
        Assert.Equal((2, 3), (bitmap.Width, bitmap.Height));
        Assert.Equal(new SKColor(255, 0, 0), bitmap.GetPixel(0, 0));
        Assert.Equal(new SKColor(0, 255, 255), bitmap.GetPixel(1, 2));
        Assert.Contains(vml ? "width: 72pt" : "width: 1in", (string?)image.Attribute("style"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void WmfOutlineShapesBecomeAnEmbeddedStaticImage(bool header, bool vml)
    {
        var html = Render(Wmf(), "image/x-wmf", header, vml);
        var image = Picture(html);
        Assert.StartsWith("data:image/svg+xml;base64,", (string?)image.Attribute("src"));
        var svg = XElement.Parse(Encoding.UTF8.GetString(Bytes(image)));
        Assert.Equal("svg", svg.Name.LocalName);
        Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "rect"));
        Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "ellipse"));
        Assert.DoesNotContain("[WMF IMAGE]", html.Value);
        Assert.Contains(vml ? "height: 48pt" : "height: 0.666", (string?)image.Attribute("style"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void ExistingPngBytesAndPictureExtentAreUnchanged(bool header, bool vml)
    {
        using var bitmap = new SKBitmap(2, 3);
        bitmap.Erase(new SKColor(48, 80, 128));
        using var png = bitmap.Encode(SKEncodedImageFormat.Png, 100);
        var bytes = png.ToArray();
        var image = Picture(Render(bytes, "image/png", header, vml));
        Assert.Equal(bytes, Bytes(image));
        Assert.Contains(vml ? "width: 72pt" : "width: 1in", (string?)image.Attribute("style"));
    }

    [Theory]
    [InlineData(1, 0, 0)]
    [InlineData(2, 1, 0)]
    [InlineData(3, 1, 2)]
    [InlineData(4, 0, 2)]
    [InlineData(5, 0, 0)]
    [InlineData(6, 2, 0)]
    [InlineData(7, 2, 1)]
    [InlineData(8, 0, 1)]
    public void TiffOrientationMovesPixelsAndSwapsIntrinsicDimensions(int orientation, int redX, int redY)
    {
        using var bitmap = SKBitmap.Decode(Bytes(Picture(Render(Tiff(orientation: orientation), "image/tiff"))));
        Assert.NotNull(bitmap);
        Assert.Equal(orientation < 5 ? (2, 3) : (3, 2), (bitmap.Width, bitmap.Height));
        Assert.Equal(new SKColor(255, 0, 0), bitmap.GetPixel(redX, redY));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void TiffPackBitsAndByteOrderPreserveUnassociatedAlpha(bool bigEndian, bool packBits)
    {
        var html = Render(Tiff(bigEndian: bigEndian, packBits: packBits, alpha: true), "image/tiff");
        Assert.StartsWith("data:image/png;base64,", (string?)Picture(html).Attribute("src"));
        using var bitmap = SKBitmap.Decode(Bytes(Picture(html)));
        Assert.NotNull(bitmap);
        Assert.Equal(new SKColor(255, 0, 0, 128), bitmap.GetPixel(0, 0));
    }

    [Theory]
    [InlineData("truncated")]
    [InlineData("huge-dimensions")]
    [InlineData("unsupported-compression")]
    [InlineData("truncated-packbits")]
    public void UnusableTiffHasAnExplicitPlaceholderInsteadOfUndecodableBrowserBytes(string defect)
    {
        var bytes = Tiff(packBits: defect == "truncated-packbits");
        if (defect is "truncated" or "truncated-packbits") bytes = bytes[..^4];
        if (defect == "huge-dimensions") Array.Copy(BitConverter.GetBytes(uint.MaxValue), 0, bytes, 18, 4);
        if (defect == "unsupported-compression") Array.Copy(BitConverter.GetBytes((ushort)5), 0, bytes, 54, 2);
        var html = Render(bytes, "image/tiff");
        Assert.Empty(html.Descendants().Where(e => e.Name.LocalName == "img"));
        Assert.Contains(html.Descendants(), e => e.Attribute("data-content-type") != null);
        Assert.Contains("CypressAfter", html.Value);
    }

    [Fact]
    public void UnknownWmfRecordDoesNotProducePartiallyRenderedArtwork()
    {
        var bytes = Wmf(unsupported: true);
        var html = Render(bytes, "image/x-wmf");
        Assert.Empty(html.Descendants().Where(e => e.Name.LocalName == "img"));
        Assert.Contains(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "WmfImage");
    }

    private static XElement Picture(XElement html) => Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "img"));
    private static byte[] Bytes(XElement image) => Convert.FromBase64String(((string)image.Attribute("src")!).Split(',')[1]);

    private static XElement Render(byte[] bytes, string contentType, bool header = false, bool vml = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            OpenXmlPart owner = header ? main.AddNewPart<HeaderPart>("header") : main;
            var image = owner.AddNewPart<ImagePart>(contentType, "picture");
            using var data = new MemoryStream(bytes);
            image.FeedData(data);
            var picture = vml
                ? "<w:pict><v:shape id=\"Cypress\" style=\"width:72pt;height:48pt\"><v:imagedata r:id=\"picture\"/></v:shape></w:pict>"
                : $"""
                  <w:drawing><wp:inline><wp:extent cx="914400" cy="609600"/>
                    <wp:docPr id="1" name="CypressGeometry" descr="Cypress rectangle and ellipse"/>
                    <a:graphic><a:graphicData uri="{Pic.pic}"><pic:pic>
                      <pic:nvPicPr><pic:cNvPr id="1" name="CypressGeometry"/><pic:cNvPicPr/></pic:nvPicPr>
                      <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="609600"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                    </pic:pic></a:graphicData></a:graphic>
                  </wp:inline></w:drawing>
                  """;
            var namespaces = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\" xmlns:a=\"{A.a}\" xmlns:wp=\"{WP.wp}\" xmlns:pic=\"{Pic.pic}\" xmlns:v=\"{VML.vml}\"";
            var paragraph = $"<w:p><w:r>{picture}</w:r></w:p>";
            if (header) owner.PutXDocument(XDocument.Parse($"<w:hdr {namespaces}>{paragraph}</w:hdr>"));
            main.PutXDocument(XDocument.Parse($"""
                <w:document {namespaces}><w:body>{(header ? "" : paragraph)}
                  <w:p><w:r><w:t>CypressAfter</w:t></w:r></w:p><w:sectPr>
                  {(header ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "")}
                  <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360"/>
                </w:sectPr></w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w=\"{W.w}\"/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions {
            FabricateCssClasses = false, PaginationMode = 1, RenderHeadersAndFooters = true,
            RenderUnsupportedContentPlaceholders = true }));
    }

    private static byte[] Tiff(int orientation = 1, bool bigEndian = false, bool packBits = false, bool alpha = false)
    {
        using var stream = new MemoryStream();
        void U16(int v) { var b = BitConverter.GetBytes((ushort)v); if (bigEndian) Array.Reverse(b); stream.Write(b); }
        void U32(int v) { var b = BitConverter.GetBytes(v); if (bigEndian) Array.Reverse(b); stream.Write(b); }
        int samples = alpha ? 4 : 3, entries = alpha ? 12 : 11, bits = 8 + 2 + entries * 12 + 4;
        var pixels = new byte[] { 255,0,0, 0,255,0, 0,0,255, 255,255,0, 255,0,255, 0,255,255 };
        if (alpha) pixels = pixels.Chunk(3).SelectMany(p => p.Concat(new byte[] { 128 })).ToArray();
        if (packBits) pixels = new byte[] { (byte)(pixels.Length - 1) }.Concat(pixels).ToArray();
        stream.Write(bigEndian ? new byte[] { 77, 77 } : new byte[] { 73, 73 }); U16(42); U32(8); U16(entries);
        void Field(int tag, int type, int count, int value) { U16(tag); U16(type); U32(count); if (type == 3 && count == 1) { U16(value); U16(0); } else U32(value); }
        Field(256,4,1,2); Field(257,4,1,3); Field(258,3,samples,bits); Field(259,3,1,packBits ? 32773 : 1);
        Field(262,3,1,2); Field(273,4,1,bits + samples * 2); Field(274,3,1,orientation); Field(277,3,1,samples);
        Field(278,4,1,3); Field(279,4,1,pixels.Length); Field(284,3,1,1); if (alpha) Field(338,3,1,2);
        U32(0); for (int i = 0; i < samples; i++) U16(8); stream.Write(pixels); return stream.ToArray();
    }

    private static byte[] Wmf(bool unsupported = false)
    {
        using var records = new MemoryStream(); using var writer = new BinaryWriter(records);
        void Record(ushort function, params short[] words) { writer.Write(3 + words.Length); writer.Write(function); foreach (var word in words) writer.Write(word); }
        Record(0x0103,8); Record(0x020b,0,0); Record(0x020c,64,96); Record(0x02fa,0,2,0,0x5030,0x0080);
        Record(0x012d,0); Record(0x02fc,1,0,0,0); Record(0x012d,1); Record(0x041b,55,43,8,8); Record(0x0418,48,88,12,52);
        if (unsupported) Record(0x061d,0,0,0,0,0,0); Record(0);
        using var output = new MemoryStream(); using var header = new BinaryWriter(output);
        header.Write(0x9ac6cdd7U); header.Write((ushort)0); header.Write((short)0); header.Write((short)0);
        header.Write((short)96); header.Write((short)64); header.Write((ushort)96); header.Write(0);
        var prefix = output.ToArray(); ushort checksum = 0;
        for (int i = 0; i < prefix.Length; i += 2) checksum ^= BitConverter.ToUInt16(prefix,i);
        header.Write(checksum); header.Write((ushort)1); header.Write((ushort)9); header.Write((ushort)0x300);
        header.Write((uint)(9 + records.Length / 2)); header.Write((ushort)2); header.Write((uint)8); header.Write((ushort)0);
        header.Write(records.ToArray()); return output.ToArray();
    }
}
