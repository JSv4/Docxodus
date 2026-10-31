using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
        Assert.Equal("none", (string?)svg.Attribute("preserveAspectRatio"));
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
    [InlineData(0)]
    [InlineData(128)]
    [InlineData(255)]
    public void AssociatedAlphaIsUnpremultipliedBeforePngEncoding(byte alpha)
    {
        using var bitmap = SKBitmap.Decode(Bytes(Picture(Render(Tiff(alpha: true, associated: true, alphaValue: alpha), "image/tiff"))));
        Assert.NotNull(bitmap);
        Assert.Equal(new SKColor(alpha == 0 ? (byte)0 : (byte)255, 0, 0, alpha), bitmap.GetPixel(0, 0));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MultipleStripsRetainEveryRow(bool packBits)
    {
        using var bitmap = SKBitmap.Decode(Bytes(Picture(Render(Tiff(packBits: packBits, bigEndian: true, rowsPerStrip: 1), "image/tiff"))));
        Assert.NotNull(bitmap);
        Assert.Equal(new SKColor(255, 0, 0), bitmap.GetPixel(0, 0));
        Assert.Equal(new SKColor(0, 0, 255), bitmap.GetPixel(0, 1));
        Assert.Equal(new SKColor(0, 255, 255), bitmap.GetPixel(1, 2));
    }

    [Theory]
    [InlineData("truncated")]
    [InlineData("huge-dimensions")]
    [InlineData("unsupported-compression")]
    [InlineData("truncated-packbits")]
    [InlineData("expanded-packbits")]
    [InlineData("multiple-pages")]
    [InlineData("tiled")]
    [InlineData("custom-color")]
    public void UnusableTiffHasAnExplicitPlaceholderInsteadOfUndecodableBrowserBytes(string defect)
    {
        var bytes = Tiff(packBits: defect is "truncated-packbits" or "expanded-packbits");
        if (defect is "truncated" or "truncated-packbits") bytes = bytes[..^4];
        if (defect == "huge-dimensions") Array.Copy(BitConverter.GetBytes(uint.MaxValue), 0, bytes, 18, 4);
        if (defect == "unsupported-compression") Array.Copy(BitConverter.GetBytes((ushort)5), 0, bytes, 54, 2);
        if (defect == "expanded-packbits") bytes[^19] = 129; // A 128-byte run cannot fit the 18-byte image.
        if (defect == "multiple-pages") Array.Copy(BitConverter.GetBytes(8), 0, bytes, 142, 4);
        if (defect == "tiled") Array.Copy(BitConverter.GetBytes((ushort)322), 0, bytes, 130, 2);
        if (defect == "custom-color")
        {
            int colorOffset = bytes.Length;
            bytes = bytes.Concat(new uint[] { 16, 1, 235, 1, 16, 1, 235, 1, 16, 1, 235, 1 }.SelectMany(BitConverter.GetBytes)).ToArray();
            Array.Copy(BitConverter.GetBytes((ushort)532), 0, bytes, 130, 2); // ReferenceBlackWhite
            Array.Copy(BitConverter.GetBytes((ushort)5), 0, bytes, 132, 2);
            Array.Copy(BitConverter.GetBytes(6), 0, bytes, 134, 4);
            Array.Copy(BitConverter.GetBytes(colorOffset), 0, bytes, 138, 4);
        }
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

    [Fact]
    public void WmfAnisotropicWindowKeepsAUniformPenWidth()
    {
        var bytes = Wmf();
        Array.Copy(BitConverter.GetBytes((short)48), 0, bytes, 66, 2);
        Array.Copy(BitConverter.GetBytes((short)4), 0, bytes, 76, 2);
        var svg = XElement.Parse(Encoding.UTF8.GetString(Bytes(Picture(Render(bytes, "image/x-wmf")))));
        var rect = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "rect"));
        Assert.Equal("16", (string?)rect.Attribute("x"));
        Assert.Equal("8.333333333333334vw", (string?)rect.Attribute("stroke-width"));
        Assert.Equal("non-scaling-stroke", (string?)rect.Attribute("vector-effect"));
        Assert.Null(rect.Attribute("transform")); // Geometry is mapped before applying the uniform pen.
    }

    [Theory]
    [InlineData(2, 1, false, false)]
    [InlineData(1, 2, false, false)]
    [InlineData(2, 1, false, true)]
    [InlineData(1, 2, false, true)]
    [InlineData(2, 1, true, false)]
    [InlineData(1, 2, true, false)]
    [InlineData(2, 1, true, true)]
    [InlineData(1, 2, true, true)]
    public void WmfAuthoredExtentMapsGeometryWithoutStretchingThePen(int widthScale, int heightScale, bool header, bool vml)
    {
        var image = Picture(Render(Wmf(), "image/x-wmf", header, vml, 914400 * widthScale, 609600 * heightScale));
        var svg = XElement.Parse(Encoding.UTF8.GetString(Bytes(image)));
        Assert.Equal("0 0 96 64", (string?)svg.Attribute("viewBox"));
        Assert.Contains(vml ? $"width: {72 * widthScale}pt" : $"width: {widthScale}in", (string?)image.Attribute("style"));
        var rect = Assert.Single(svg.Descendants().Where(e => e.Name.LocalName == "rect"));
        Assert.Equal("2.0833333333333335vw", (string?)rect.Attribute("stroke-width"));
        Assert.Equal("non-scaling-stroke", (string?)rect.Attribute("vector-effect"));
    }

    [Fact]
    public void SubpixelWmfIntrinsicSizeStaysDecodableAtItsAuthoredExtent()
    {
        var bytes = Wmf();
        Array.Copy(BitConverter.GetBytes((short)1), 0, bytes, 10, 2);
        Array.Copy(BitConverter.GetBytes((short)1), 0, bytes, 12, 2);
        Array.Copy(BitConverter.GetBytes(ushort.MaxValue), 0, bytes, 14, 2);
        ushort checksum = 0; for (int i = 0; i < 20; i += 2) checksum ^= BitConverter.ToUInt16(bytes, i);
        Array.Copy(BitConverter.GetBytes(checksum), 0, bytes, 20, 2);
        var image = Picture(Render(bytes, "image/x-wmf"));
        var svg = XElement.Parse(Encoding.UTF8.GetString(Bytes(image)));
        Assert.True((double)svg.Attribute("width")! >= 1);
        Assert.True((double)svg.Attribute("height")! >= 1);
        Assert.Contains("width: 1in", (string?)image.Attribute("style"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Docx2HtmlKeepsNormalizedResourcesInEmbeddedAndExternalModes(bool tiff, bool embed)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        var configuration = directory.Parent!.Name;
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "Docxodus.Tests", "Docxodus.Tests.csproj"))) directory = directory.Parent;
        var tool = Path.Combine(directory!.FullName, "tools", "docx2html", "bin", configuration, "net10.0", "docx2html.dll");
        var temp = Directory.CreateTempSubdirectory("image-normalization-").FullName;
        try
        {
            var input = Path.Combine(temp, "input.docx"); var output = Path.Combine(temp, "output.html");
            File.WriteAllBytes(input, Package(tiff ? Tiff() : Wmf(), tiff ? "image/tiff" : "image/x-wmf"));
            var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet") {
                RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            foreach (var arg in new[] { tool, input, output }) start.ArgumentList.Add(arg);
            if (!embed) start.ArgumentList.Add("--extract-images");
            using var process = Process.Start(start)!;
            var stdout = process.StandardOutput.ReadToEndAsync(); var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            try { await process.WaitForExitAsync(timeout.Token); }
            catch { if (!process.HasExited) process.Kill(entireProcessTree: true); throw; }
            Assert.True(process.ExitCode == 0, await stdout + await stderr);
            var image = Picture(XElement.Load(output));
            var bytes = embed ? Bytes(image) : File.ReadAllBytes(Path.Combine(temp, (string)image.Attribute("src")!));
            if (tiff) { using var bitmap = SKBitmap.Decode(bytes); Assert.NotNull(bitmap); Assert.Equal((2, 3), (bitmap.Width, bitmap.Height)); }
            else Assert.Equal("svg", XElement.Parse(Encoding.UTF8.GetString(bytes)).Name.LocalName);
        }
        finally { Directory.Delete(temp, recursive: true); }
    }

    private static XElement Picture(XElement html) => Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "img"));
    private static byte[] Bytes(XElement image) => Convert.FromBase64String(((string)image.Attribute("src")!).Split(',')[1]);

    private static XElement Render(byte[] bytes, string contentType, bool header = false, bool vml = false,
        int widthEmus = 914400, int heightEmus = 609600)
        => XElement.Parse(HtmlConversionOps.ConvertToHtml(Package(bytes, contentType, header, vml, widthEmus, heightEmus), new HtmlConversionOptions {
            FabricateCssClasses = false, PaginationMode = 1, RenderHeadersAndFooters = true,
            RenderUnsupportedContentPlaceholders = true }));

    private static byte[] Package(byte[] bytes, string contentType, bool header = false, bool vml = false,
        int widthEmus = 914400, int heightEmus = 609600)
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
                ? FormattableString.Invariant($"<w:pict><v:shape id=\"Cypress\" style=\"width:{widthEmus / 12700.0}pt;height:{heightEmus / 12700.0}pt\"><v:imagedata r:id=\"picture\"/></v:shape></w:pict>")
                : $"""
                  <w:drawing><wp:inline><wp:extent cx="{widthEmus}" cy="{heightEmus}"/>
                    <wp:docPr id="1" name="CypressGeometry" descr="Cypress rectangle and ellipse"/>
                    <a:graphic><a:graphicData uri="{Pic.pic}"><pic:pic>
                      <pic:nvPicPr><pic:cNvPr id="1" name="CypressGeometry"/><pic:cNvPicPr/></pic:nvPicPr>
                      <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{widthEmus}" cy="{heightEmus}"/></a:xfrm>
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
        return stream.ToArray();
    }

    private static byte[] Tiff(int orientation = 1, bool bigEndian = false, bool packBits = false, bool alpha = false,
        bool associated = false, byte alphaValue = 128, int rowsPerStrip = 3)
    {
        using var stream = new MemoryStream();
        void U16(int v) { var b = BitConverter.GetBytes((ushort)v); if (bigEndian) Array.Reverse(b); stream.Write(b); }
        void U32(int v) { var b = BitConverter.GetBytes(v); if (bigEndian) Array.Reverse(b); stream.Write(b); }
        int samples = alpha ? 4 : 3, entries = alpha ? 12 : 11, bits = 8 + 2 + entries * 12 + 4;
        var pixels = new byte[] { 255,0,0, 0,255,0, 0,0,255, 255,255,0, 255,0,255, 0,255,255 };
        if (alpha) pixels = pixels.Chunk(3).SelectMany(p => p.Select(v => associated ? (byte)(v * alphaValue / 255) : v).Concat(new byte[] { alphaValue })).ToArray();
        var strips = pixels.Chunk(2 * rowsPerStrip * samples).Select(p => packBits ? new byte[] { (byte)(p.Length - 1) }.Concat(p).ToArray() : p).ToArray();
        int offsetArray = bits + samples * 2, countArray = offsetArray + strips.Length * 4;
        int pixelStart = strips.Length == 1 ? offsetArray : countArray + strips.Length * 4;
        stream.Write(bigEndian ? new byte[] { 77, 77 } : new byte[] { 73, 73 }); U16(42); U32(8); U16(entries);
        void Field(int tag, int type, int count, int value) { U16(tag); U16(type); U32(count); if (type == 3 && count == 1) { U16(value); U16(0); } else U32(value); }
        Field(256,4,1,2); Field(257,4,1,3); Field(258,3,samples,bits); Field(259,3,1,packBits ? 32773 : 1);
        Field(262,3,1,2); Field(273,4,strips.Length,strips.Length == 1 ? pixelStart : offsetArray); Field(274,3,1,orientation); Field(277,3,1,samples);
        Field(278,4,1,rowsPerStrip); Field(279,4,strips.Length,strips.Length == 1 ? strips[0].Length : countArray);
        Field(284,3,1,1); if (alpha) Field(338,3,1,associated ? 1 : 2);
        U32(0); for (int i = 0; i < samples; i++) U16(8);
        if (strips.Length > 1) {
            int offset = pixelStart; foreach (var strip in strips) { U32(offset); offset += strip.Length; }
            foreach (var strip in strips) U32(strip.Length);
        }
        foreach (var strip in strips) stream.Write(strip);
        return stream.ToArray();
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
        header.Write((uint)(9 + records.Length / 2)); header.Write((ushort)2); header.Write((uint)(unsupported ? 9 : 8)); header.Write((ushort)0);
        header.Write(records.ToArray()); return output.ToArray();
    }
}
