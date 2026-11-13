using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using SkiaSharp;
using Xunit;

namespace Docxodus.Tests;

/// <summary>EMF and WMF playback beyond rectangles and ellipses (#1706): each record family is built
/// byte by byte and its SVG checked; the realistic files go through the public converter.</summary>
public class MetafileRendererTests
{
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    // ---- Geometry ------------------------------------------------------------------------------

    [Fact]
    public void ConsecutiveLinesWithOnePenBecomeOneStrokedPath()
    {
        var path = Single(Emf(R(27, 10, 10), R(54, 50, 10), R(54, 50, 40), R(27, 0, 0), R(54, 5, 5)), "path");
        Assert.Equal("M10 10L50 10L50 40M0 0L5 5", (string?)path.Attribute("d"));
        Assert.Equal("none", (string?)path.Attribute("fill"));
        Assert.Equal("non-scaling-stroke", (string?)path.Attribute("vector-effect"));
    }

    [Fact]
    public void PolygonFamiliesIncludingSixteenBitFormsFillWithThePolyFillMode()
    {
        var shapes = Elements(Emf(
            R(86, 0, 0, 0, 0, 3, Pack16(0, 0, 10, 0, 10, 10)), // POLYGON16
            R(19, 2), // SETPOLYFILLMODE WINDING
            R(8, 0, 0, 0, 0, 2, 6, 3, 3, 0, 0, 4, 0, 4, 4, 9, 9, 12, 9, 12, 12), // POLYPOLYGON
            R(2, 0, 0, 0, 0, 4, 0, 0, 1, 2, 3, 2, 4, 0), // POLYBEZIER
            R(87, 0, 0, 0, 0, 2, Pack16(1, 1, 7, 7))), "path"); // POLYLINE16
        // Open figures drawn with one pen share a path.
        Assert.Equal(new[] { "M0 0L10 0L10 10Z", "M0 0L4 0L4 4ZM9 9L12 9L12 12Z", "M0 0C1 2 3 2 4 0M1 1L7 7" },
            shapes.Select(e => (string?)e.Attribute("d")));
        Assert.Equal(new[] { "evenodd", "nonzero", null }, shapes.Select(e => (string?)e.Attribute("fill-rule")));
        Assert.Equal("#ffffff", (string?)shapes[1].Attribute("fill"));
        Assert.Equal("none", (string?)shapes[2].Attribute("fill"));
    }

    [Theory]
    [InlineData(45, false, "M100 50A50 50 0 0 0 50 0")]          // ARC: anticlockwise on the device
    [InlineData(46, true, "M100 50A50 50 0 0 0 50 0Z")]          // CHORD
    [InlineData(47, true, "M50 50L100 50A50 50 0 0 0 50 0Z")]    // PIE
    public void ArcsSweepAnticlockwiseByDefault(uint kind, bool filled, string d)
    {
        var path = Single(Emf(R(kind, 0, 0, 100, 100, 100, 50, 50, 0)), "path");
        Assert.Equal(d, (string?)path.Attribute("d"));
        Assert.Equal(filled ? "#ffffff" : "none", (string?)path.Attribute("fill"));
    }

    [Fact]
    public void ClockwiseArcDirectionTakesTheOtherSide()
        => Assert.Equal("M100 50A50 50 0 1 1 50 0", (string?)Single(Emf(R(57, 2), R(45, 0, 0, 100, 100, 100, 50, 50, 0)), "path").Attribute("d"));

    [Fact]
    public void RoundRectKeepsItsCornerRadii()
    {
        var rect = Single(Emf(R(44, 0, 0, 40, 20, 10, 6)), "rect");
        Assert.Equal(("5", "3"), ((string?)rect.Attribute("rx"), (string?)rect.Attribute("ry")));
    }

    [Fact]
    public void PathBracketsFillAndStrokeTheirFigures()
    {
        var shapes = Elements(Emf(R(59), R(27, 0, 0), R(54, 30, 0), R(54, 30, 30), R(61), R(43, 40, 40, 50, 50), R(60), R(63),
            R(59), R(27, 0, 0), R(54, 9, 9), R(60), R(64), R(59), R(27, 0, 0), R(54, 9, 0), R(54, 9, 9), R(60), R(62)), "path");
        Assert.Equal(new[] { "M0 0L30 0L30 30ZM40 40L50 40L50 50L40 50Z", "M0 0L9 9", "M0 0L9 0L9 9Z" },
            shapes.Select(e => (string?)e.Attribute("d")));
        Assert.Equal(new[] { "#ffffff", "none", "#ffffff" }, shapes.Select(e => (string?)e.Attribute("fill")));
        Assert.Equal(new[] { "#000000", "#000000", "none" }, shapes.Select(e => (string?)e.Attribute("stroke")));
        Assert.Empty(Elements(Emf(R(59), R(27, 0, 0), R(54, 9, 9), R(60), R(68), R(62)), "path")); // ABORTPATH
    }

    // ---- Device context ------------------------------------------------------------------------

    [Fact]
    public void MappingModesAndSaveRestoreDriveEachShapeTransform()
    {
        var rects = Elements(Emf(
            R(33), R(17, 8), R(9, 200, 100), R(11, 100, 100), R(10, 10, 0), R(12, 5, 5), R(43, 10, 0, 20, 10), // anisotropic
            R(17, 7), R(43, 10, 0, 20, 10),                                                                   // isotropic
            R(34, -1), R(43, 0, 0, 10, 10),                                                                   // restored MM_TEXT
            R(17, 3), R(43, 0, 0, 1000, 1000)), "rect");                                                       // HIMETRIC: y up
        Assert.Equal(new[] { "matrix(0.5 0 0 1 0 5)", "matrix(0.5 0 0 0.5 0 5)", null, "scale(0.032 -0.032)" },
            rects.Select(e => (string?)e.Attribute("transform")));
    }

    [Fact]
    public void WorldTransformsComposeWithTheirModifyMode()
    {
        var rects = Elements(Emf(
            R(35, F(0), F(1), F(-1), F(0), F(10), F(0)), R(43, 0, 0, 5, 5),   // SETWORLDTRANSFORM: rotate 90°, shift
            R(36, F(2), F(0), F(0), F(2), F(0), F(0), 2), R(43, 0, 0, 5, 5),  // left-multiply: scale first
            R(36, F(1), F(0), F(0), F(1), F(0), F(0), 1), R(43, 0, 0, 5, 5)), "rect"); // identity
        Assert.Equal(new[] { "matrix(0 1 -1 0 10 0)", "matrix(0 2 -2 0 10 0)", null }, rects.Select(e => (string?)e.Attribute("transform")));
    }

    [Fact]
    public void DegenerateInputsBehaveAsGdiDoes()
    {
        // An empty POLYBEZIER draws nothing, a zero window extent is refused (the old one stays), and a
        // tiny world scale is kept rather than rounded to nothing.
        var rects = Elements(Emf(R(2, 0, 0, 0, 0, 0), R(17, 8), R(9, 0, 5), R(43, 0, 0, 5, 5),
            R(17, 1), R(35, F(1e-5f), F(0), F(0), F(1e-5f), F(0), F(0)), R(43, 0, 0, 500000, 500000)), "rect");
        Assert.Equal(new[] { null, "scale(1E-05 1E-05)" }, rects.Select(e => (string?)e.Attribute("transform")));
    }

    [Fact]
    public void RestoringBelowTheSavedDepthRejectsThePicture()
        => Assert.Throws<InvalidDataException>(() => EmfImage.Read(Emf(R(33), R(34, -2))));

    // ---- Pens and brushes ----------------------------------------------------------------------

    [Fact]
    public void PenStylesCarryDashesCapsAndJoins()
    {
        // PS_GEOMETRIC | PS_DASH | PS_ENDCAP_FLAT | PS_JOIN_BEVEL, width 4; then a cosmetic PS_DOT.
        var paths = Elements(Emf(R(95, 1, 0, 0, 0, 0, 0x10000 | 0x1 | 0x200 | 0x1000, 4, 0, 0x0000ff, 0, 0), R(37, 1), R(27, 0, 0), R(54, 9, 9),
            R(38, 2, 2, 0, 0, 0x00ff00), R(37, 2), R(54, 0, 9)), "path");
        Assert.Equal(("#ff0000", "4", "72 24", "bevel"), ((string?)paths[0].Attribute("stroke"), (string?)paths[0].Attribute("stroke-width"),
            (string?)paths[0].Attribute("stroke-dasharray"), (string?)paths[0].Attribute("stroke-linejoin")));
        Assert.Null(paths[0].Attribute("stroke-linecap"));
        Assert.Null(paths[0].Attribute("vector-effect"));
        Assert.Equal(("#00ff00", "3 3", "non-scaling-stroke"), ((string?)paths[1].Attribute("stroke"),
            (string?)paths[1].Attribute("stroke-dasharray"), (string?)paths[1].Attribute("vector-effect")));
    }

    [Fact]
    public void HatchedBrushesFillWithADevicePixelPattern()
    {
        var svg = EmfImage.Read(Emf(R(39, 1, 2, 0x0000ff, 4), R(37, 1), R(18, 1), R(43, 0, 0, 10, 10))).Drawing;
        var pattern = Assert.Single(svg.Descendants(Svg + "pattern"));
        Assert.Equal($"url(#{(string?)pattern.Attribute("id")})", (string?)Assert.Single(svg.Descendants(Svg + "rect"), e => e.Parent?.Name.LocalName != "pattern").Attribute("fill"));
        Assert.Equal("M0 4.5H8M4.5 0V8", (string?)pattern.Element(Svg + "path")!.Attribute("d"));
        Assert.Null(pattern.Element(Svg + "rect")); // TRANSPARENT background mode leaves the gaps unpainted
    }

    [Fact]
    public void FillingWithAPatternBrushRejectsThePicture()
        => Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(94, 1, 0, 0, 0, 0, 0), R(37, 1), R(43, 0, 0, 10, 10))));

    // ---- Clipping ------------------------------------------------------------------------------

    [Fact]
    public void ClipRectanglesIntersectExcludeAndRestore()
    {
        var svg = EmfImage.Read(Emf(R(33), R(30, 0, 0, 50, 50), R(29, 10, 10, 20, 20), R(43, 0, 0, 60, 60),
            R(34, -1), R(43, 1, 1, 2, 2))).Drawing;
        var clips = svg.Descendants(Svg + "clipPath").ToList();
        Assert.Equal(2, clips.Count);
        Assert.Equal("M0 0L50 0L50 50L0 50Z", (string?)clips[0].Element(Svg + "path")!.Attribute("d"));
        Assert.Equal($"url(#{(string?)clips[0].Attribute("id")})", (string?)clips[1].Attribute("clip-path"));
        Assert.Equal("evenodd", (string?)clips[1].Element(Svg + "path")!.Attribute("clip-rule"));
        var rects = svg.Descendants(Svg + "rect").ToList();
        Assert.Equal($"url(#{(string?)clips[1].Attribute("id")})", (string?)rects[0].Parent!.Attribute("clip-path"));
        Assert.Null(rects[1].Parent!.Attribute("clip-path")); // RESTOREDC dropped the clip
    }

    [Fact]
    public void ClipPathsAndRegionsReplaceOrIntersect()
    {
        var svg = EmfImage.Read(Emf(R(59), R(27, 0, 0), R(54, 9, 0), R(54, 9, 9), R(60), R(67, 5), R(43, 0, 0, 5, 5),
            R(75, 0, 5), R(43, 0, 0, 6, 6), R(75, 48, 1, 32, 1, 1, 16, 0, 0, 4, 4, 0, 0, 4, 4), R(43, 0, 0, 7, 7))).Drawing;
        var rects = svg.Descendants(Svg + "rect").ToList();
        var clips = svg.Descendants(Svg + "clipPath").ToList();
        Assert.Equal("M0 0L9 0L9 9Z", (string?)clips[0].Element(Svg + "path")!.Attribute("d"));
        Assert.Equal($"url(#{(string?)clips[0].Attribute("id")})", (string?)rects[0].Parent!.Attribute("clip-path"));
        Assert.Null(rects[1].Parent!.Attribute("clip-path")); // EXTSELECTCLIPRGN RGN_COPY without data resets
        Assert.Equal("M0 0H4V4H0Z", (string?)clips[1].Element(Svg + "path")!.Attribute("d"));
        Assert.Equal($"url(#{(string?)clips[1].Attribute("id")})", (string?)rects[2].Parent!.Attribute("clip-path"));
    }

    // ---- Text ----------------------------------------------------------------------------------

    [Fact]
    public void ExtTextOutDrawsSelectableTextInItsFontColourAndAlignment()
    {
        var svg = EmfImage.Read(Emf(Font(1, "Arial", -20, 0, 700, italic: true), R(37, 1), R(24, 0x804020), R(22, 6 | 24),
            TextW(30, 40, "Hi!", new[] { 10, 6, 8 }))).Drawing;
        var text = Assert.Single(svg.Descendants(Svg + "text"));
        Assert.Equal("Hi!", text.Value);
        Assert.Equal(("'Arial', sans-serif", "20", "700", "italic", "#204080"), ((string?)text.Attribute("font-family"),
            (string?)text.Attribute("font-size"), (string?)text.Attribute("font-weight"), (string?)text.Attribute("font-style"), (string?)text.Attribute("fill")));
        Assert.Equal("-12 -2 4", (string?)text.Attribute("x")); // TA_CENTER over the 24-unit advance
        Assert.Null(text.Attribute("dominant-baseline"));       // TA_BASELINE
        Assert.Equal("matrix(1 0 0 1 30 40)", (string?)text.Attribute("transform"));
    }

    [Fact]
    public void EscapementRotatesTextAndFlippedMappingKeepsItUpright()
    {
        var text = Assert.Single(EmfImage.Read(Emf(R(17, 3), Font(1, "Arial", -500, 900, 400), R(37, 1), TextW(0, 0, "Up", null))).Drawing
            .Descendants(Svg + "text"));
        // MM_HIMETRIC flips y; text still reads upward on the page.
        Assert.Equal("matrix(0 -0.032 0.032 0 0 0)", (string?)text.Attribute("transform"));
        Assert.Equal("text-before-edge", (string?)text.Attribute("dominant-baseline"));
    }

    [Fact]
    public void GlyphIndexTextRejectsThePicture()
        => Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(TextW(0, 0, "x", null, options: 0x10))));

    // ---- Bitmaps -------------------------------------------------------------------------------

    [Fact]
    public void StretchDibitsEmbedsTheBitmapAsAPngInItsDestination()
    {
        var image = Assert.Single(EmfImage.Read(Emf(StretchDib(10, 20, 40, 30))).Drawing.Descendants(Svg + "image"));
        Assert.Equal("matrix(40 0 0 30 10 20)", (string?)image.Attribute("transform"));
        var href = (string)image.Attribute("href")!;
        Assert.StartsWith("data:image/png;base64,", href);
        using var bitmap = SKBitmap.Decode(Convert.FromBase64String(href.Split(',')[1]));
        // Bottom-up rows: the first stored row is the bottom of the picture.
        Assert.Equal(new SKColor(255, 0, 0), bitmap.GetPixel(0, 1));
        Assert.Equal(new SKColor(0, 255, 0), bitmap.GetPixel(1, 1));
        Assert.Equal(new SKColor(0, 0, 255), bitmap.GetPixel(0, 0));
        Assert.Equal(new SKColor(255, 255, 255), bitmap.GetPixel(1, 0));
    }

    [Theory]
    [InlineData(0x00660046u)] // SRCINVERT combines with the destination
    [InlineData(0x00CC0020u)] // SRCCOPY, but with palette colours (DIB_PAL_COLORS)
    public void BitmapsThatNeedTheDestinationOrAPaletteRejectThePicture(uint rop)
        => Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(StretchDib(0, 0, 2, 2, rop, usage: rop == 0x00CC0020u ? 1 : 0))));

    [Fact]
    public void AlphaBlendBakesPremultipliedAndConstantAlphaIntoThePng()
    {
        // One 32-bit pixel, premultiplied: red 64 at alpha 128; blend function AC_SRC_ALPHA with constant alpha 128.
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 0, 0, 0, 0, 5, 5, 10, 10, 0x01800000, 0, 0, F(1), F(0), F(0), F(1), F(0), F(0), 0, 0, 108, 40, 148, 4, 1, 1 }) w.Write(value);
        foreach (int value in new[] { 40, 1, 1, 1 | 32 << 16, 0, 4, 0, 0, 0, 0 }) w.Write(value);
        w.Write(new byte[] { 0, 0, 64, 128 });
        var image = Assert.Single(EmfImage.Read(Emf(Raw(114, stream.ToArray()))).Drawing.Descendants(Svg + "image"));
        using var bitmap = SKBitmap.Decode(Convert.FromBase64String(((string)image.Attribute("href")!).Split(',')[1]));
        var pixel = bitmap.GetPixel(0, 0);
        Assert.Equal(64, pixel.Alpha);
        Assert.InRange(pixel.Red, 126, 128); // 64 · 255 / 128, through PNG and Skia's premultiplied decode
    }

    // ---- EMF+ and unknown records --------------------------------------------------------------

    [Fact]
    public void DualEmfPlusPlaysItsGdiRecordsAndEmfPlusOnlyIsUnsupported()
    {
        Assert.Single(EmfImage.Read(Emf(EmfPlusHeader(dual: true), R(43, 0, 0, 10, 10))).Drawing.Descendants(Svg + "rect"));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(EmfPlusHeader(dual: false), R(43, 0, 0, 10, 10))));
    }

    [Fact]
    public void HarmlessStateRecordsAreIgnoredButUnknownRecordsStillReject()
    {
        Assert.Single(Elements(Emf(R(18, 1), R(20, 13), R(21, 4), R(98, 1), R(13, 0, 0), R(16, 0), R(58, F(10)), R(43, 0, 0, 5, 5)), "rect"));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(43, 0, 0, 5, 5), R(118))));  // GRADIENTFILL
    }

    [Fact]
    public void PointCountsBeyondTheRecordAreInvalid()
        => Assert.Throws<InvalidDataException>(() => EmfImage.Read(Emf(R(4, 0, 0, 0, 0, int.MaxValue, 1, 2))));

    // ---- WMF -----------------------------------------------------------------------------------

    [Fact]
    public void WmfPlaysLinesPolygonsArcsTextClippingAndBitmaps()
    {
        var bitmap = Dib24();
        var svg = XElement.Parse(Encoding.UTF8.GetString(WmfImage.ToSvg(Wmf(
            W(0x02fa, 0, 2, 0, 0x00ff, 0), W(0x012d, 0),                      // red 2-unit pen
            W(0x0416, 60, 90, 0, 0),                                           // INTERSECTCLIPRECT (0,0)-(90,60)
            W(0x0214, 5, 5), W(0x0213, 5, 50),                                 // MOVETO / LINETO
            W(0x0324, 3, 10, 10, 30, 10, 20, 30),                              // POLYGON
            W(0x0817, 0, 50, 25, 75, 50, 75, 0, 25),                           // ARC
            W(0x0209, 0x00ff, 0), W(0x02fb, -12, 0, 0, 0, 400, 0, 0, 0, 0, 0x7254, 0x6c69), W(0x012d, 1), // text colour, font "Tril"
            W(0x0521, 2, 0x6948, 40, 10),                                      // TEXTOUT "Hi" at (10,40)
            W(0x0f43, Words(new[] { 0x0020, 0x00cc, 0, 2, 2, 0, 0, 20, 20, 30, 60 }, bitmap)))))); // STRETCHDIB
        Assert.Contains(svg.Descendants(Svg + "path"), e => (string?)e.Attribute("d") == "M5 5L50 5");
        Assert.Contains(svg.Descendants(Svg + "path"), e => (string?)e.Attribute("d") == "M10 10L30 10L20 30Z");
        Assert.Contains(svg.Descendants(Svg + "path"), e => ((string?)e.Attribute("d"))!.Contains('A'));
        var text = Assert.Single(svg.Descendants(Svg + "text"));
        Assert.Equal(("Hi", "#ff0000", "'Tril', sans-serif"), (text.Value, (string?)text.Attribute("fill"), (string?)text.Attribute("font-family")));
        Assert.StartsWith("data:image/png;base64,", (string?)Assert.Single(svg.Descendants(Svg + "image")).Attribute("href"));
        Assert.All(svg.Descendants().Where(e => e.Name.LocalName is "path" or "text" or "image" && e.Parent?.Name.LocalName == "g"),
            e => Assert.NotNull(e.Parent!.Attribute("clip-path")));
    }

    [Fact]
    public void WmfObjectTableKeepsSlotsForPalettesAndRegions()
    {
        // CREATEPALETTE takes slot 0, so the pen lands in slot 1; selecting slot 0 would select the palette.
        var svg = XElement.Parse(Encoding.UTF8.GetString(WmfImage.ToSvg(Wmf(W(0x00f7, 0x300, 0), W(0x02fa, 0, 3, 0, 0x00ff, 0), W(0x012d, 1),
            W(0x0214, 0, 0), W(0x0213, 9, 9)))));
        Assert.Equal("3", (string?)Assert.Single(svg.Descendants(Svg + "path")).Attribute("stroke-width"));
        Assert.Throws<NotSupportedException>(() => WmfImage.ToSvg(Wmf(W(0x00f7, 0x300, 0), W(0x012d, 0))));
    }

    // ---- Bounds and the unsupported-keeps-the-placeholder contract ------------------------------

    [Fact]
    public void WmfPolyPolygonCountsThatSumPastAnIntAreInvalid()
    {
        // 16,384 polygons of 65,535 points and one of 16,384: the counts sum to 2^30, and four bytes a
        // point is 2^32, which a 32-bit product wraps to zero. No points follow.
        var counts = Enumerable.Repeat(65535, 16384).Append(16384);
        Assert.Throws<InvalidDataException>(() => WmfImage.ToSvg(Wmf(W(0x0538, counts.Prepend(16385).ToArray()))));
        // 2^29 points: the 32-bit product is negative, which also passed "not more than the record holds".
        Assert.Throws<InvalidDataException>(() => WmfImage.ToSvg(Wmf(W(0x0538, Enumerable.Repeat(65535, 8192).Append(8192).Prepend(8193).ToArray()))));
    }

    [Fact]
    public void APictureMayNotDrawMoreThanItsPointBudget()
    {
        // The budget is for the whole picture: two polylines of 125,001 points are one too many.
        byte[] Polyline(int count) => R(4, 0, 0, 0, 0, count, new int[2 * count]);
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Polyline(125_001), Polyline(125_001))));
        Assert.Single(Elements(Emf(Polyline(125_000), Polyline(125_000)), "path"));
    }

    [Fact]
    public void APictureMayNotWriteMoreSvgThanItsOutputBudget()
    {
        // Noise does not compress: a 2400×2000 bitmap is about 19 million base64 characters of PNG.
        var noise = new byte[2400 * 2000 * 3];
        new Random(1706).NextBytes(noise);
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(StretchDib24(2400, 2000, noise))));
        Assert.Single(Elements(Emf(StretchDib24(240, 200, noise[..(240 * 200 * 3)])), "image"));
    }

    [Fact]
    public void LargeCoordinatesKeepEveryDigit()
    {
        // A world transform brings logical coordinates of a hundred million onto the page.
        var path = Single(Emf(R(35, F(1e-6f), 0, 0, F(1e-6f), 0, 0), R(4, 0, 0, 0, 0, 2, 100000001, 0, 100000003, 0)), "path");
        Assert.Equal("M100000001 0L100000003 0", (string?)path.Attribute("d"));
    }

    [Fact]
    public void ClipRegionsNestedBeyondTheSupportedDepthRejectThePicture()
    {
        byte[][] Clips(int n) => Enumerable.Range(0, n).Select(_ => R(30, 0, 0, 50, 50)).Append(R(43, 0, 0, 5, 5)).ToArray();
        Assert.Single(Elements(Emf(Clips(MetafileCanvas.MaxClipDepth)), "rect"));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Clips(MetafileCanvas.MaxClipDepth + 1))));
    }

    [Theory]
    [InlineData(4)] // BI_JPEG
    [InlineData(5)] // BI_PNG
    public void BitmapsThatWrapAJpegOrPngRejectThePicture(int compression)
        => Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(StretchDib24(1, 1, new byte[64], compression))));

    [Fact]
    public void AnsiTextInANonWesternCharacterSetRejectsThePicture()
    {
        byte[] Picture(byte charset, params byte[] text) => Emf(Font(1, "MS Gothic", -20, 0, 400, charset: charset), R(37, 1), TextA(0, 0, text));
        // SHIFTJIS_CHARSET (128): 0x83 0x41 is one katakana, not the two Windows-1252 characters.
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Picture(128, 0x83, 0x41)));
        Assert.Equal("Ab", Single(Picture(128, 0x41, 0x62), "text").Value);  // ASCII reads the same in every set
        Assert.Equal("é", Single(Picture(0, 0xe9), "text").Value);          // ANSI_CHARSET is Windows-1252
    }

    [Fact]
    public void SymbolFontAndPrivateUseTextRejectThePicture()
    {
        // SYMBOL_CHARSET (2): "a" in Symbol is an alpha, and in Wingdings a sign of the zodiac.
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Font(1, "Symbol", -20, 0, 400, charset: 2), R(37, 1), TextW(0, 0, "a", null))));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(TextW(0, 0, "", null))));
        Assert.Throws<NotSupportedException>(() => WmfImage.ToSvg(Wmf(
            W(0x02fb, -12, 0, 0, 0, 400, 0, 0x0200, 0, 0, 0x7953, 0x626d), W(0x012d, 0), W(0x0521, 2, 0x6948, 40, 10))));
    }

    [Fact]
    public void TextThatAdvancesTheCurrentPositionByItsOwnWidthRejectsThePicture()
    {
        // TA_UPDATECP without a spacing array: the next run starts where this one's glyphs end.
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(22, 1), TextW(0, 0, "Hi", null))));
        Assert.Single(Elements(Emf(R(22, 1), TextW(0, 0, "Hi", new[] { 7, 5 })), "text"));
    }

    [Fact]
    public void DrawingModesAndSpacingThatChangeTheInkRejectThePicture()
    {
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(20, 7), R(43, 0, 0, 5, 5))));            // SETROP2 R2_XORPEN
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(120, 3, 1), TextW(0, 0, "a b", null)))); // SETTEXTJUSTIFICATION
        Assert.Throws<NotSupportedException>(() => WmfImage.ToSvg(Wmf(W(0x0104, 6))));                           // SETROP2 R2_NOT
        Assert.Throws<NotSupportedException>(() => WmfImage.ToSvg(Wmf(W(0x0108, 2))));                           // SETTEXTCHAREXTRA
        Assert.Throws<NotSupportedException>(() => WmfImage.ToSvg(Wmf(W(0x020a, 1, 3))));                        // SETTEXTJUSTIFICATION
        WmfImage.ToSvg(Wmf(W(0x0104, 13), W(0x0108, 0), W(0x020a, 0, 0)));                                       // the defaults draw as before
    }

    // ---- Realistic files through the public converter ------------------------------------------

    [Fact]
    public void LibreOfficeEmfRendersItsGeometryTextAndBitmapInline()
    {
        // TestFiles/Metafiles/libreoffice-chart.emf: LibreOffice 26.8 export of libreoffice-chart.source.svg
        // (`soffice --headless --convert-to emf`), a dual EMF+ file with paths, text, a clip and STRETCHDIBITS.
        var html = HtmlImageNormalizationTests.Render(Fixture("libreoffice-chart.emf"), "image/x-emf");
        Assert.DoesNotContain("[EMF IMAGE]", html.Value);
        var svg = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-vector-frame") == "emf");
        Assert.Contains("Quarterly 42", svg.Value);
        Assert.True(svg.Descendants().Count(e => e.Name.LocalName == "path") >= 5);
        Assert.Single(svg.Descendants(), e => e.Name.LocalName == "image");
        Assert.Contains("CypressAfter", html.Value);
    }

    [Fact]
    public void LibreOfficeWmfRendersAsAnEmbeddedSvgImage()
    {
        // TestFiles/Metafiles/libreoffice-chart.wmf: the same drawing exported with `--convert-to wmf`.
        var html = HtmlImageNormalizationTests.Render(Fixture("libreoffice-chart.wmf"), "image/x-wmf");
        Assert.DoesNotContain("[WMF IMAGE]", html.Value);
        var svg = XElement.Parse(Encoding.UTF8.GetString(HtmlImageNormalizationTests.Bytes(HtmlImageNormalizationTests.Picture(html))));
        Assert.Contains("Quarterly 42", svg.Value);
        Assert.Single(svg.Descendants(Svg + "image"));
    }

    [Fact]
    public void WordWrittenDualEmfPlusThumbnailPlaysItsGdiText()
    {
        // Word wrote this thumbnail (EMF+ dual, text and lines) into an existing corpus fixture.
        var svg = EmfImage.Read(PackagePart(Path.Combine("WC", "WC-BodyBookmarks-Before.docx"), "docProps/thumbnail.emf")).Drawing;
        Assert.NotEmpty(svg.Descendants(Svg + "text"));
        Assert.NotEmpty(svg.Descendants(Svg + "path"));
    }

    [Fact]
    public void WordEmbeddedWorkbookIconRendersThroughThePublicConverter()
    {
        // HC044's OLE preview is a Word-written EMF: an ALPHABLEND icon above its caption, among GDI comments.
        // The converter shows the OLE object itself as a placeholder, so the preview is placed as a picture.
        var html = HtmlImageNormalizationTests.Render(PackagePart("HC044-Embedded-Workbook.docx", "word/media/image1.emf"), "image/x-emf");
        Assert.DoesNotContain("[EMF IMAGE]", html.Value);
        var svg = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-vector-frame") == "emf");
        Assert.Equal("Test.xlsx", Assert.Single(svg.Descendants(), e => e.Name.LocalName == "text").Value);
        Assert.Single(svg.Descendants(), e => e.Name.LocalName == "image");
    }

    [Theory]
    [InlineData("libreoffice-chart.emf")]
    [InlineData("libreoffice-chart.wmf")]
    [InlineData("hc044")]
    public void CorruptedMetafilesFailOnlyWithTheExceptionsTheConverterCatches(string name)
    {
        // The converter turns InvalidData/NotSupported/Overflow into the placeholder; anything else would
        // abort the whole document. Overwrite bytes at seeded random places and play each result. Every
        // other trial instead overwrites one whole field (a 16-bit WMF word, a 32-bit EMF field) with a
        // value at the edge of its range, which is what reaches a count or size that overflows.
        var original = name == "hc044" ? PackagePart("HC044-Embedded-Workbook.docx", "word/media/image1.emf") : Fixture(name);
        int field = name.EndsWith(".wmf", StringComparison.Ordinal) ? 2 : 4;
        uint[] edges = { 0, 1, 0x7fff, 0x8000, 0xffff, 0x10000, 0x20000000, 0x40000000, 0x7fffffff, 0x80000000, 0xffffffff };
        var random = new Random(1706);
        for (int trial = 0; trial < 1600; trial++)
        {
            var bytes = (byte[])original.Clone();
            if (trial % 2 == 0)
                for (int i = random.Next(1, 4); i > 0; i--) bytes[random.Next(bytes.Length)] = (byte)random.Next(256);
            else
                BitConverter.GetBytes(edges[random.Next(edges.Length)]).AsSpan(0, field)
                    .CopyTo(bytes.AsSpan(random.Next(bytes.Length / field) * field));
            try
            {
                if (name.EndsWith(".wmf", StringComparison.Ordinal)) WmfImage.ToSvg(bytes);
                else EmfImage.Read(bytes).Drawing.ToXmlString(SaveOptions.DisableFormatting);
            }
            catch (Exception e) when (e is InvalidDataException or NotSupportedException or OverflowException) { }
        }
    }

    [Fact]
    public void EmfPlusOnlyPictureKeepsThePlaceholderThroughTheConverter()
    {
        var html = HtmlImageNormalizationTests.Render(Emf(EmfPlusHeader(dual: false)), "image/x-emf");
        Assert.Contains("[EMF IMAGE]", html.Value);
        Assert.DoesNotContain(html.Descendants(), e => e.Attribute("data-docx-vector-frame") != null);
    }

    // ---- Builders ------------------------------------------------------------------------------

    private static string FixtureRoot => Path.GetFullPath("../../../../TestFiles/");
    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(FixtureRoot, "Metafiles", name));

    private static byte[] PackagePart(string docx, string part)
    {
        using var package = ZipFile.OpenRead(Path.Combine(FixtureRoot, docx));
        using var stream = new MemoryStream();
        using (var entry = package.GetEntry(part)!.Open()) entry.CopyTo(stream);
        return stream.ToArray();
    }

    private static XElement Single(byte[] emf, string name) => Assert.Single(Elements(emf, name));
    private static List<XElement> Elements(byte[] emf, string name)
        => EmfImage.Read(emf).Drawing.Descendants(Svg + name).Where(e => e.Ancestors().All(a => a.Name.LocalName is not ("pattern" or "clipPath"))).ToList();

    private static int F(float value) => BitConverter.SingleToInt32Bits(value);
    private static int[] Pack16(params short[] values)
        => Enumerable.Range(0, (values.Length + 1) / 2).Select(i => (ushort)values[i * 2] | (i * 2 + 1 < values.Length ? values[i * 2 + 1] << 16 : 0)).ToArray();

    /// <summary>An EMF record of 32-bit fields; nested arrays are flattened in place.</summary>
    private static byte[] R(uint type, params object[] fields)
    {
        var values = fields.SelectMany(f => f is int[] list ? list.Cast<object>() : new[] { f }).Select(Convert.ToInt32).ToArray();
        var record = new byte[8 + values.Length * 4];
        BitConverter.GetBytes(type).CopyTo(record, 0); BitConverter.GetBytes(record.Length).CopyTo(record, 4);
        for (int i = 0; i < values.Length; i++) BitConverter.GetBytes(values[i]).CopyTo(record, 8 + i * 4);
        return record;
    }

    private static byte[] Raw(uint type, byte[] payload)
    {
        var record = new byte[8 + (payload.Length + 3) / 4 * 4];
        BitConverter.GetBytes(type).CopyTo(record, 0); BitConverter.GetBytes(record.Length).CopyTo(record, 4);
        payload.CopyTo(record, 8);
        return record;
    }

    /// <summary>Issue #1554's header: Frame (0,0)-(2646,2646), a 1024×768 device of 320×240mm.</summary>
    private static byte[] Emf(params byte[][] records)
    {
        var body = records.Concat(new[] { R(14, 0, 16, 20) }).SelectMany(r => r).ToArray();
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 1, 88, 0, 0, 100, 100, 0, 0, 2646, 2646, 0x464d4520, 0x10000, 88 + body.Length, records.Length + 2 }) w.Write(value);
        w.Write((ushort)8); w.Write((ushort)0);
        foreach (int value in new[] { 0, 0, 0, 1024, 768, 320, 240 }) w.Write(value);
        w.Write(body);
        return stream.ToArray();
    }

    private static byte[] Font(int slot, string face, int height, int escapement, int weight, bool italic = false, byte charset = 0)
    {
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { slot, height, 0, escapement, escapement, weight }) w.Write(value);
        w.Write(new byte[] { (byte)(italic ? 1 : 0), 0, 0, charset, 0, 0, 0, 0 });
        var name = new byte[64]; Encoding.Unicode.GetBytes(face).CopyTo(name, 0); w.Write(name);
        return Raw(82, stream.ToArray());
    }

    private static byte[] TextW(int x, int y, string text, int[]? dx, int options = 0)
    {
        int stringAt = 76, dxAt = dx == null ? 0 : stringAt + (text.Length * 2 + 3) / 4 * 4;
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 0, 0, 0, 0, 1, F(1), F(1), x, y, text.Length, stringAt, options, 0, 0, 0, 0, dxAt }) w.Write(value);
        var chars = new byte[(text.Length * 2 + 3) / 4 * 4]; Encoding.Unicode.GetBytes(text).CopyTo(chars, 0); w.Write(chars);
        foreach (int advance in dx ?? Array.Empty<int>()) w.Write(advance);
        return Raw(84, stream.ToArray());
    }

    /// <summary>EXTTEXTOUTA: the same record as <see cref="TextW"/> with one byte per character.</summary>
    private static byte[] TextA(int x, int y, byte[] text)
    {
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 0, 0, 0, 0, 1, F(1), F(1), x, y, text.Length, 76, 0, 0, 0, 0, 0, 0 }) w.Write(value);
        w.Write(text);
        return Raw(83, stream.ToArray());
    }

    /// <summary>STRETCHDIBITS of a whole <paramref name="width"/>×<paramref name="height"/> 24-bit DIB.</summary>
    private static byte[] StretchDib24(int width, int height, byte[] pixels, int compression = 0)
    {
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 0, 0, 0, 0, 0, 0, 0, 0, width, height, 80, 40, 120, pixels.Length, 0, 0x00CC0020, width, height }) w.Write(value);
        foreach (int value in new[] { 40, width, height, 1 | 24 << 16, compression, pixels.Length, 0, 0, 0, 0 }) w.Write(value);
        w.Write(pixels);
        return Raw(81, stream.ToArray());
    }

    /// <summary>A 2×2 bottom-up 24-bit DIB: red, green on the bottom row; blue, white on the top.</summary>
    private static byte[] Dib24()
    {
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 40, 2, 2, 1 | 24 << 16, 0, 16, 0, 0, 0, 0 }) w.Write(value);
        w.Write(new byte[] { 0, 0, 255, 0, 255, 0, 0, 0, 255, 0, 0, 255, 255, 255, 0, 0 });
        return stream.ToArray();
    }

    private static byte[] StretchDib(int x, int y, int cx, int cy, uint rop = 0x00CC0020, int usage = 0)
    {
        var dib = Dib24();
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 0, 0, 0, 0, x, y, 0, 0, 2, 2, 80, 40, 120, 16, usage, unchecked((int)rop), cx, cy }) w.Write(value);
        w.Write(dib);
        return Raw(81, stream.ToArray());
    }

    private static byte[] EmfPlusHeader(bool dual)
    {
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        w.Write(32); w.Write(0x2B464D45);
        w.Write((ushort)0x4001); w.Write((ushort)(dual ? 1 : 0)); w.Write(28); w.Write(16);
        foreach (int value in new[] { unchecked((int)0xDBC01002), 0, 96, 96 }) w.Write(value);
        return Raw(70, stream.ToArray());
    }

    /// <summary>A WMF record of 16-bit parameter words.</summary>
    private static byte[] W(ushort function, params int[] words)
    {
        var record = new byte[6 + words.Length * 2];
        BitConverter.GetBytes(3 + words.Length).CopyTo(record, 0); BitConverter.GetBytes(function).CopyTo(record, 4);
        for (int i = 0; i < words.Length; i++) BitConverter.GetBytes((ushort)words[i]).CopyTo(record, 6 + i * 2);
        return record;
    }

    private static int[] Words(int[] leading, byte[] tail)
        => leading.Concat(Enumerable.Range(0, (tail.Length + 1) / 2).Select(i => tail[i * 2] | (i * 2 + 1 < tail.Length ? tail[i * 2 + 1] << 8 : 0))).ToArray();

    /// <summary>A placeable WMF: bbox (0,0)-(96,64) at 96 units per inch, MM_ANISOTROPIC with a matching window.</summary>
    private static byte[] Wmf(params byte[][] records)
    {
        var all = new[] { W(0x0103, 8), W(0x020b, 0, 0), W(0x020c, 64, 96) }.Concat(records).Concat(new[] { W(0) }).ToArray();
        var body = all.SelectMany(r => r).ToArray();
        using var output = new MemoryStream(); using var header = new BinaryWriter(output);
        header.Write(0x9ac6cdd7U); header.Write((ushort)0); header.Write((short)0); header.Write((short)0);
        header.Write((short)96); header.Write((short)64); header.Write((ushort)96); header.Write(0);
        var prefix = output.ToArray(); ushort checksum = 0;
        for (int i = 0; i < prefix.Length; i += 2) checksum ^= BitConverter.ToUInt16(prefix, i);
        header.Write(checksum); header.Write((ushort)1); header.Write((ushort)9); header.Write((ushort)0x300);
        header.Write((uint)(9 + body.Length / 2)); header.Write((ushort)4); header.Write((uint)all.Max(r => r.Length / 2)); header.Write((ushort)0);
        header.Write(body);
        return output.ToArray();
    }
}
