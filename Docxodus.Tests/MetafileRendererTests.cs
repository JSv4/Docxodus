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

    [Fact]
    public void APathBracketTakesTheTransformOfItsFirstFigure()
    {
        // Office scales the world by 1/8 inside BEGINPATH and back before ENDPATH, and GDI turns each point into
        // device space as it is added. The pen, though, is realised at the transform in effect when stroking.
        byte[][] Bracket(uint draw) => new[] { R(59), R(36, F(0.125f), 0, 0, F(0.125f), 0, 0, 2),
            R(86, 0, 0, 0, 0, 3, Pack16(0, 0, 80, 0, 80, 80)), R(36, F(8), 0, 0, F(8), 0, 0, 2), R(60), R(draw) };
        var path = Single(Emf(Bracket(62)), "path"); // FILLPATH
        Assert.Equal(("M0 0L80 0L80 80Z", "scale(0.125 0.125)"), ((string?)path.Attribute("d"), (string?)path.Attribute("transform")));
        Assert.Single(Elements(Emf(Bracket(64)), "path")); // STROKEPATH with the default one-pixel pen
        var geometric = R(95, 1, 0, 0, 0, 0, 0x10000, 8, 0, 0, 0, 0);
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(new[] { geometric, R(37, 1) }.Concat(Bracket(64)).ToArray())));
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
    public void HatchedBrushesFillWithAnEightPixelPatternInCssPixels()
    {
        var svg = EmfImage.Read(Emf(R(39, 1, 2, 0x0000ff, 4), R(37, 1), R(18, 1), R(43, 0, 0, 10, 10))).Drawing;
        var pattern = Assert.Single(svg.Descendants(Svg + "pattern"));
        Assert.Equal($"url(#{(string?)pattern.Attribute("id")})", (string?)Assert.Single(svg.Descendants(Svg + "rect"), e => e.Parent?.Name.LocalName != "pattern").Attribute("fill"));
        Assert.Equal("M0 4.5H8M4.5 0V8", (string?)pattern.Element(Svg + "path")!.Attribute("d"));
        Assert.Null(pattern.Element(Svg + "rect")); // TRANSPARENT background mode leaves the gaps unpainted
        // GDI realises a brush in the pixels of the device it plays onto. The reference device has 3.2 pixels
        // a millimetre, so one CSS pixel (0.2646 mm) is 0.8467 of them.
        Assert.Equal("scale(0.8466666667 0.8466666667)", (string?)pattern.Attribute("patternTransform"));
    }

    [Fact]
    public void DibPatternBrushesTileTheirBitmap()
    {
        // CREATEDIBPATTERNBRUSHPT and WMF DIBCREATEPATTERNBRUSH in RGB colours: the 2×2 bitmap is the tile.
        var dib = Dib24();
        var patterns = EmfImage.Read(Emf(Raw(94, Le(1, 0, 32, 40, 72, dib.Length - 40).Concat(dib).ToArray()),
            R(37, 1), R(43, 0, 0, 10, 10), R(43, 20, 0, 30, 10))).Drawing.Descendants(Svg + "pattern").ToList();
        Assert.Equal(("2", "2", "scale(0.8466666667 0.8466666667)"), ((string?)patterns[0].Attribute("width"),
            (string?)patterns[0].Attribute("height"), (string?)patterns[0].Attribute("patternTransform")));
        var tiles = patterns.Select(p => (string)p.Element(Svg + "image")!.Attribute("href")!).ToList();
        Assert.StartsWith("data:image/png;base64,", tiles[0]);
        Assert.Same(tiles[0], tiles[1]); // the tile is encoded once, not again for every fill
        var wmf = Play(Wmf(W(0x0142, Words(new[] { 5, 0 }, dib)), W(0x012d, 0), W(0x041b, 20, 20, 0, 0)));
        Assert.Equal("2", (string?)Assert.Single(wmf.Descendants(Svg + "pattern")).Attribute("width"));
    }

    [Fact]
    public void MonochromePatternBrushesPaintTheTextAndBackgroundColoursOfTheFill()
    {
        // GDI paints a monochrome brush's 0 bits in the text colour and its 1 bits in the background colour current
        // at the fill. GDI writes CREATEMONOBRUSH without a colour table, and a WMF BS_PATTERN brush with one it ignores.
        var bits = OneBit(0x40, 0x80); // top row: 0 then 1; bottom row: 1 then 0
        var svg = EmfImage.Read(Emf(R(24, 0x0000ff), R(25, 0x00ff00), Raw(93, Le(1, 2, 32, 40, 72, 8).Concat(bits).ToArray()),
            R(37, 1), R(43, 0, 0, 10, 10), R(24, 0xff0000), R(43, 20, 0, 30, 10), R(43, 40, 0, 50, 10))).Drawing;
        var tiles = svg.Descendants(Svg + "pattern").Select(p => p.Element(Svg + "image")!).ToList();
        Assert.Equal(new[] { SKColors.Red, SKColors.Lime }, new[] { Png(tiles[0]).GetPixel(0, 0), Png(tiles[0]).GetPixel(1, 0) });
        Assert.Equal(SKColors.Blue, Png(tiles[1]).GetPixel(0, 0)); // the next fill's text colour
        Assert.Same((string?)tiles[1].Attribute("href"), (string?)tiles[2].Attribute("href")); // encoded once per colour pair
        var wmf = Play(Wmf(W(0x0209, 0x00ff, 0), W(0x0201, 0xff00, 0),
            W(0x0142, Words(new[] { 3, 0 }, OneBit(0x40, 0x80, 0x0000ff, 0xff0000))), W(0x012d, 0), W(0x041b, 20, 20, 0, 0)));
        var tile = Png(wmf.Descendants(Svg + "pattern").Single().Element(Svg + "image")!);
        Assert.Equal(new[] { SKColors.Red, SKColors.Lime }, new[] { tile.GetPixel(0, 0), tile.GetPixel(1, 0) });
    }

    [Fact]
    public void PaletteIndexPatternBrushesTakeTheDefaultPalette()
    {
        // DIB_PAL_COLORS: the colour table holds indexes into the palette, the default one unless the metafile
        // selects its own. GDI writes them packed in an EMF and in four-byte slots in a WMF.
        byte[] Indexed(int tableBytes) => OneBit(0x40, 0x80).Take(40).Concat(new byte[] { 1, 0, 4, 0 }).Concat(new byte[tableBytes - 4])
            .Concat(OneBit(0x40, 0x80).Skip(40)).ToArray();
        byte[] Brush(byte[]? palette = null) => Raw(94, Le(1, 1, 32, 44, 76, 8).Concat(palette ?? Indexed(4)).ToArray());
        var tile = Png(EmfImage.Read(Emf(Brush(), R(37, 1), R(43, 0, 0, 10, 10))).Drawing.Descendants(Svg + "image").Single());
        Assert.Equal(new[] { new SKColor(0x80, 0, 0), new SKColor(0, 0, 0x80) }, new[] { tile.GetPixel(0, 0), tile.GetPixel(1, 0) });
        var wmf = Png(Play(Wmf(W(0x0142, Words(new[] { 5, 1 }, Indexed(8))), W(0x012d, 0), W(0x041b, 20, 20, 0, 0))).Descendants(Svg + "image").Single());
        Assert.Equal(new SKColor(0, 0, 0x80), wmf.GetPixel(1, 0));
        EmfImage.Read(Emf(R(48, unchecked((int)0x8000000F)), Brush(), R(37, 1), R(43, 0, 0, 10, 10))); // the stock default palette
        // Controls: the metafile's own palette, an index past the default palette, and DIB_PAL_INDICES.
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(49, 2, 0x300 | 1 << 16, 0), R(48, 2), Brush(), R(37, 1), R(43, 0, 0, 10, 10))));
        var past = Indexed(4); past[40] = 20;
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Brush(past), R(37, 1), R(43, 0, 0, 10, 10))));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Raw(94, Le(1, 2, 32, 40, 72, 8).Concat(OneBit(0x40, 0x80)).ToArray()),
            R(37, 1), R(43, 0, 0, 10, 10))));
    }

    [Fact]
    public void AZeroWidthWmfPenDrawsOneCssPixelLikeAnEmfCosmeticPen()
    {
        // PowerPoint outlines arrowheads with zero-width pens.
        var path = Assert.Single(Play(Wmf(W(0x02fa, 0, 0, 0, 0x00ff, 0), W(0x012d, 0), W(0x0214, 0, 0), W(0x0213, 9, 9))).Descendants(Svg + "path"));
        Assert.Equal(("#ff0000", "1", "non-scaling-stroke"), ((string?)path.Attribute("stroke"), (string?)path.Attribute("stroke-width"),
            (string?)path.Attribute("vector-effect")));
    }

    [Fact]
    public void TheNopPenModeDrawsNothingWithThePenOrBrush()
    {
        // R2_NOP leaves the destination as it is; PowerPoint brackets the clip paths it writes for PostScript
        // printers with it. PATBLT is a raster operation of its own and still paints.
        var svg = Play(Wmf(W(0x0104, 11), W(0x0324, 3, 0, 0, 9, 0, 9, 9), W(0x0214, 0, 0), W(0x0213, 9, 9),
            W(0x061d, 0x0021, 0x00f0, 4, 4, 0, 0), W(0x0104, 13), W(0x041b, 30, 30, 20, 20)));
        Assert.Empty(svg.Descendants(Svg + "path"));
        Assert.Equal(2, svg.Descendants(Svg + "rect").Count());
    }

    [Fact]
    public void FillsThatShareAnEdgeWithTheFillBeforeThemAreDrawnCrisp()
    {
        // Office draws a gradient as abutting strips, which GDI fills without antialiasing; antialiased, the
        // background shows along each shared edge.
        byte[] Strip(short top) => R(86, 0, 0, 0, 0, 4, Pack16(0, top, 10, top, 10, (short)(top + 5), 0, (short)(top + 5)));
        var noPen = R(37, unchecked((int)0x80000008));
        Assert.Equal(new[] { "crispEdges", "crispEdges", "crispEdges" },
            Elements(Emf(noPen, Strip(0), Strip(5), Strip(10)), "path").Select(e => (string?)e.Attribute("shape-rendering")));
        // Controls: a fill alone, fills apart, outlined fills and curved ones sharing a radius keep smooth edges.
        Assert.Null(Single(Emf(noPen, Strip(0)), "path").Attribute("shape-rendering"));
        foreach (var emf in new[] { Emf(noPen, Strip(0), Strip(6)), Emf(Strip(0), Strip(5)),
            Emf(noPen, R(47, 0, 0, 10, 10, 10, 5, 5, 0), R(47, 0, 0, 10, 10, 5, 0, 0, 5)) })
            Assert.All(Elements(emf, "path"), e => Assert.Null(e.Attribute("shape-rendering")));
    }

    // ---- Inversions ----------------------------------------------------------------------------

    /// <summary>Office's semi-transparent fill: PATINVERT a solid brush, fill the shape with a black and white
    /// pattern under R2_MASKPEN, PATINVERT again. <paramref name="between"/> replaces the shape's fill.</summary>
    private static byte[] Inversion(byte[]? between = null, int repeatColor = 0x00d8d2cc) => Emf(
        R(39, 1, 0, 0x00d8d2cc, 0), R(37, 1), PatBlt(0, 0, 20, 10, 0x005A0049), R(20, 9),
        Raw(94, Le(2, 0, 32, 48, 80, 8).Concat(Mono(0x00, 0x40)).ToArray()), R(37, 2), R(37, unchecked((int)0x80000008)),
        between ?? R(86, 0, 0, 0, 0, 3, Pack16(0, 0, 10, 0, 10, 10)),
        R(20, 13), R(39, 3, 0, repeatColor, 0), R(37, 3), PatBlt(0, 0, 20, 10, 0x005A0049));

    [Fact]
    public void AnInversionDrawsItsContentThroughTheBlackOfItsMask()
    {
        // Three of the pattern's four pixels are black: where the shape masks, the brush shows at 75%, and the two
        // PATINVERTs cancel everywhere else, whatever lay beneath.
        var svg = EmfImage.Read(Inversion()).Drawing;
        var mask = Assert.Single(svg.Descendants(Svg + "mask"));
        var shape = Assert.Single(mask.Elements(Svg + "path"));
        Assert.Equal(("M0 0L10 0L10 10Z", "#ffffff", "0.75"),
            ((string?)shape.Attribute("d"), (string?)shape.Attribute("fill"), (string?)shape.Attribute("fill-opacity")));
        // The region is the drawing's bounds and two CSS pixels: Chromium draws nothing through a far larger one.
        Assert.Equal(new[] { -1.693, -1.693, 23.387, 13.387 }, new[] { "x", "y", "width", "height" }
            .Select(name => Math.Round(double.Parse((string)mask.Attribute(name)!, System.Globalization.CultureInfo.InvariantCulture), 3)));
        var content = Assert.Single(svg.Descendants(Svg + "g"), g => g.Attribute("mask") != null);
        Assert.Equal($"url(#{(string?)mask.Attribute("id")})", (string?)content.Attribute("mask"));
        var rect = Assert.Single(content.Elements());
        Assert.Equal(("rect", "#ccd2d8", "20"), (rect.Name.LocalName, (string?)rect.Attribute("fill"), (string?)rect.Attribute("width")));
        Assert.Single(svg.Descendants(Svg + "rect")); // drawn once, not once per pass
    }

    [Fact]
    public void XorPenPassesAroundABlackShapeClipTheirStripsToIt()
    {
        // PowerPoint's WMF clips a gradient to an area this way: the strips XORed, the area filled black, the
        // strips XORed again.
        byte[][] Strips() => new[] { W(0x0104, 7), W(0x012d, 1), W(0x041b, 10, 40, 0, 0), W(0x012d, 2), W(0x041b, 20, 40, 10, 0) };
        var svg = Play(Wmf(new[] { W(0x02fa, 5, 0, 0, 0, 0), W(0x012d, 0), W(0x02fc, 0, 0x00ff, 0, 0), W(0x02fc, 0, 0, 0x00ff, 0),
            W(0x02fc, 0, 0, 0, 0) }.Concat(Strips()).Concat(new[] { W(0x0104, 13), W(0x012d, 3), W(0x0324, 3, 0, 0, 40, 0, 0, 20) })
            .Concat(Strips()).Append(W(0x0104, 13)).ToArray()));
        var shape = Assert.Single(Assert.Single(svg.Descendants(Svg + "mask")).Elements());
        Assert.Equal(("#ffffff", null), ((string?)shape.Attribute("fill"), (string?)shape.Attribute("fill-opacity")));
        var strips = Assert.Single(svg.Descendants(Svg + "g"), g => g.Attribute("mask") != null).Elements(Svg + "rect").ToList();
        Assert.Equal(new[] { "#ff0000", "#0000ff" }, strips.Select(e => (string?)e.Attribute("fill")));
    }

    [Fact]
    public void InversionsThatWouldReadTheDestinationRejectThePicture()
    {
        EmfImage.Read(Inversion()); // the control draws
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Inversion(repeatColor: 0x00ff0000)));            // the second pass differs
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Inversion(new[] { R(39, 4, 2, 0, 0), R(37, 4),
            R(43, 0, 0, 5, 5) }.SelectMany(r => r).ToArray())));                                                    // a hatched brush paints the mask
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Inversion(new[] { R(20, 13), R(39, 4, 0, 0x0000ff, 0), R(37, 4),
            R(43, 0, 0, 5, 5) }.SelectMany(r => r).ToArray())));                                                    // the mask is painted red
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Inversion(TextW(0, 0, "a", null))));             // text between the passes
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(20, 9), R(43, 0, 0, 5, 5))));              // R2_MASKPEN outside one
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(PatBlt(0, 0, 9, 9, 0x005A0049))));           // never repeated
    }

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

    [Fact]
    public void ConsecutiveClipExclusionsCutHolesInOneClipPath()
    {
        // PowerPoint carves a shape's clip from hundreds of one-unit strips, far more than clip paths may nest.
        var strips = Enumerable.Range(0, 100).Select(i => R(29, 0, i, 10, i + 1)).Append(R(43, 0, 0, 20, 200)).ToArray();
        var clip = Assert.Single(EmfImage.Read(Emf(strips)).Drawing.Descendants(Svg + "clipPath"));
        Assert.Equal(101, ((string)clip.Element(Svg + "path")!.Attribute("d")!).Count(c => c == 'Z'));
        // Even-odd filling would fill an overlap of two holes again, and a clip that something already refers
        // to (a drawn shape, a saved state) must not change: the second exclusion nests in each case.
        foreach (var records in new[] {
            new[] { R(29, 0, 0, 10, 10), R(29, 5, 5, 15, 15), R(43, 0, 0, 20, 20) },
            new[] { R(29, 0, 0, 5, 5), R(43, 0, 0, 20, 20), R(29, 10, 10, 15, 15), R(43, 0, 0, 20, 20) },
            new[] { R(29, 0, 0, 5, 5), R(33), R(29, 10, 10, 15, 15), R(34, -1), R(43, 0, 0, 20, 20) } })
        {
            var clips = EmfImage.Read(Emf(records)).Drawing.Descendants(Svg + "clipPath").ToList();
            Assert.Equal(2, clips.Count);
            Assert.Equal(2, ((string)clips[0].Element(Svg + "path")!.Attribute("d")!).Count(c => c == 'Z'));
        }
    }

    [Fact]
    public void WmfSelectClipRegionZeroResetsTheClip()
    {
        // Index 0 is the null region, which GDI's EMF-to-WMF conversion and PowerPoint write for SelectClipRgn(NULL).
        var svg = Play(Wmf(W(0x0416, 10, 10, 0, 0), W(0x012c, 0), W(0x041b, 20, 20, 0, 0)));
        Assert.Null(Assert.Single(svg.Descendants(Svg + "rect")).Parent!.Attribute("clip-path"));
        Assert.Throws<NotSupportedException>(() => WmfImage.ToSvg(Wmf(W(0x012c, 1)))); // a real region
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
    public void GlyphIndexTextRejectsThePictureUnlessItHasNoGlyphs()
    {
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(TextW(0, 0, "x", null, options: 0x10))));
        // Office's embedded-object icons write empty glyph-index runs, which draw nothing.
        Assert.Empty(Elements(Emf(TextW(0, 0, "", null, options: 0x10)), "text"));
        Assert.Empty(Play(Wmf(W(0x0a32, 0, 0, 0, 0x10))).Descendants(Svg + "text"));
    }

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
    public void AMaskThenItsImageDrawOneTransparentBitmap()
    {
        // An icon is drawn SRCAND with a monochrome mask (black opaque), then SRCINVERT with the image (black where
        // transparent); Office draws SRCPAINT with the mask (white opaque), then SRCAND with the image (white where
        // transparent). Either pair merges into one PNG; the same raster operations play from WMF.
        foreach (var (records, color) in new[] {
            (new[] { StretchDib(0, 0, 20, 20, 0x008800C6, Mono(0x40, 0xc0)), StretchDib(0, 0, 20, 20, 0x00660046, Rgb24(0xff0000, 0, 0, 0)) }, new SKColor(255, 0, 0)),
            (new[] { StretchDib(0, 0, 20, 20, 0x00EE0086, Mono(0x80, 0x00)), StretchDib(0, 0, 20, 20, 0x008800C6, Rgb24(0x00ff00, 0xffffff, 0xffffff, 0xffffff)) }, new SKColor(0, 255, 0)) })
        {
            using var png = Png(Single(Emf(records), "image"));
            Assert.Equal(color, png.GetPixel(0, 0));
            Assert.All(new[] { (1, 0), (0, 1), (1, 1) }, p => Assert.Equal(0, png.GetPixel(p.Item1, p.Item2).Alpha));
        }
        byte[] WmfBlit(uint rop, byte[] dib) => W(0x0f43, Words(new[] { (int)(rop & 0xffff), (int)(rop >> 16), 0, 2, 2, 0, 0, 20, 20, 0, 0 }, dib));
        var image = Assert.Single(Play(Wmf(WmfBlit(0x008800C6, Mono(0x40, 0xc0)), WmfBlit(0x00660046, Rgb24(0xff0000, 0, 0, 0)))).Descendants(Svg + "image"));
        using var wmf = Png(image);
        Assert.Equal((new SKColor(255, 0, 0), (byte)0), (wmf.GetPixel(0, 0), wmf.GetPixel(1, 1).Alpha));
    }

    [Fact]
    public void ABitmapThatReadsAnUnknownDestinationRejectsThePicture()
    {
        // SRCINVERT's coloured pixels depend on what lies beneath: unknown where the mask left the destination
        // as it was, and everywhere when the mask went to another rectangle.
        var mask = StretchDib(0, 0, 20, 20, 0x008800C6, Mono(0x40, 0xc0));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(mask, StretchDib(0, 0, 20, 20, 0x00660046, Rgb24(0xff0000, 0, 0, 0x0000ff)))));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(mask, StretchDib(5, 0, 20, 20, 0x00660046, Rgb24(0xff0000, 0, 0, 0)))));
    }

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
    public void AWmfWithAnEmbeddedEmfPlaysTheEmf()
    {
        // GDI's EMF-to-WMF conversion keeps the source EMF in META_ESCAPE comments, and GDI plays that EMF in place
        // of the WMF records. Here the EMF draws a rectangle and the WMF records an ellipse.
        var ellipse = W(0x0418, 48, 88, 12, 52);
        var emf = Emf(R(43, 0, 0, 10, 10));
        Assert.Single(Play(Wmf(Embedded(emf, emf.Length).Append(ellipse).ToArray())).Descendants(Svg + "rect"));
        // Chunks that do not add up to the declared size, or an EMF this player rejects, leave the records to draw.
        Assert.Single(Play(Wmf(Embedded(emf, emf.Length + 4).Append(ellipse).ToArray())).Descendants(Svg + "ellipse"));
        var gradient = Emf(R(43, 0, 0, 10, 10), R(118));
        Assert.Single(Play(Wmf(Embedded(gradient, gradient.Length).Append(ellipse).ToArray())).Descendants(Svg + "ellipse"));
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
        // The budget of 250,000 is for the whole picture, and a stroked point is charged when it is read and
        // again when it is drawn: two polylines of 62,500 points spend exactly all of it.
        byte[] Polyline(int count) => R(4, 0, 0, 0, 0, count, new int[2 * count]);
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Polyline(62_501), Polyline(62_501))));
        Assert.Single(Elements(Emf(Polyline(62_500), Polyline(62_500)), "path"));
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
    public void SymbolFontTextIsMappedToUnicode()
    {
        // Symbol's "a" is an alpha, which Office stores as the font's own code or moved to U+F061; Wingdings'
        // byte 0xFC is a check mark. Each advance stays with its character.
        var text = Single(Emf(Font(1, "Symbol", -20, 0, 400, charset: 2), R(37, 1), TextW(0, 0, "\uF061a", new[] { 5, 6 })), "text");
        Assert.Equal(("αα", "sans-serif", "0 5"), (text.Value, (string?)text.Attribute("font-family"), (string?)text.Attribute("x")));
        var wmf = Play(Wmf(W(0x02fb, -12, 0, 0, 0, 400, 0, 0x0200, 0, 0, 0x6957, 0x676e, 0x6964, 0x676e, 0x0073), W(0x012d, 0),
            W(0x0521, 1, 0x00fc, 40, 10)));
        Assert.Equal("\u2713", Assert.Single(wmf.Descendants(Svg + "text")).Value);
    }

    [Fact]
    public void SymbolTextWithoutAUnicodeEquivalentAndPrivateUseTextRejectThePicture()
    {
        // Symbol's byte 0xA0 has no glyph in the font Windows ships, and another symbol-charset font has no mapping.
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Font(1, "Symbol", -20, 0, 400, charset: 2), R(37, 1), TextW(0, 0, "\uF0A0", null))));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(Font(1, "Marlett", -20, 0, 400, charset: 2), R(37, 1), TextW(0, 0, "a", null))));
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(TextW(0, 0, "\uF0B7", null)))); // private use in an ordinary font
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
        Assert.Throws<NotSupportedException>(() => EmfImage.Read(Emf(R(20, 7), R(43, 0, 0, 5, 5))));            // an XOR pass never repeated
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
    [InlineData("office-scatter-markers.emf", "Sales scatter-markers", "image")] // markers: a mask, then the image
    [InlineData("office-area-gradient.emf", "Sales area-gradient", "clipPath")] // a clip path drawn at a 1/8 world scale
    [InlineData("office-symbols.emf", "Symbol font: αβ≥∑Wingdings: ☺✓", "text")]
    [InlineData("office-symbols-converted.wmf", "Unicode: αβ ≤ ± →", "text")]   // its embedded EMF keeps what ANSI lost
    [InlineData("office-pattern-star.emf", null, "pattern")]                     // a DIB pattern brush
    [InlineData("office-pattern-star.wmf", null, "pattern")]
    [InlineData("office-pie-chart.wmf", "1st Qtr", "clipPath")]                  // SELECTCLIPREGION 0 between clip rectangles
    [InlineData("office-dashed-arrow.wmf", null, "path")]                        // a zero-width pen outlines the arrowhead
    [InlineData("office-smartart-transparent.emf", null, "mask")]                // semi-transparent fills by PATINVERT and R2_MASKPEN
    [InlineData("office-area-gradient.wmf", "Sales area-gradient", "mask")]      // a gradient clipped by XOR passes
    public void OfficeMetafilesRenderThroughThePublicConverter(string name, string? text, string element)
    {
        // TestFiles/Metafiles/README.md records how Office wrote each one.
        bool emf = name.EndsWith(".emf", StringComparison.Ordinal);
        var html = HtmlImageNormalizationTests.Render(Fixture(name), emf ? "image/x-emf" : "image/x-wmf");
        Assert.DoesNotContain(emf ? "[EMF IMAGE]" : "[WMF IMAGE]", html.Value);
        var svg = emf ? Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-vector-frame") == "emf")
            : XElement.Parse(Encoding.UTF8.GetString(HtmlImageNormalizationTests.Bytes(HtmlImageNormalizationTests.Picture(html))));
        if (text != null) Assert.Contains(text, svg.Value);
        Assert.Contains(svg.Descendants(), e => e.Name.LocalName == element);
    }

    [Theory]
    [InlineData("libreoffice-chart.emf")]
    [InlineData("libreoffice-chart.wmf")]
    [InlineData("hc044")]
    [InlineData("office-scatter-markers.emf")]
    [InlineData("office-area-gradient.emf")]
    [InlineData("office-symbols.emf")]
    [InlineData("office-symbols-converted.wmf")]
    [InlineData("office-pattern-star.emf")]
    [InlineData("office-pattern-star.wmf")]
    [InlineData("office-pie-chart.wmf")]
    [InlineData("office-dashed-arrow.wmf")]
    [InlineData("office-smartart-transparent.emf")]
    [InlineData("office-area-gradient.wmf")]
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

    /// <summary>STRETCHDIBITS of a whole 2×2 DIB (<see cref="Dib24"/> unless given), whose header and colour table
    /// are all but its last eight or sixteen bytes.</summary>
    private static byte[] StretchDib(int x, int y, int cx, int cy, uint rop = 0x00CC0020, byte[]? dib = null, int usage = 0)
    {
        dib ??= Dib24();
        int bits = BitConverter.ToUInt16(dib, 14) == 1 ? 8 : 16;
        using var stream = new MemoryStream(); using var w = new BinaryWriter(stream);
        foreach (int value in new[] { 0, 0, 0, 0, x, y, 0, 0, 2, 2, 80, dib.Length - bits, 80 + dib.Length - bits, bits, usage, unchecked((int)rop), cx, cy }) w.Write(value);
        w.Write(dib);
        return Raw(81, stream.ToArray());
    }

    /// <summary>A 2×2 one-bit DIB, black for 0 and white for 1; each row's bits start at the left pixel.</summary>
    private static byte[] Mono(byte top, byte bottom)
        => Le(40, 2, 2, 1 | 1 << 16, 0, 8, 0, 0, 0, 0, 0, 0xffffff).Concat(new byte[] { bottom, 0, 0, 0, top, 0, 0, 0 }).ToArray();

    /// <summary>A 2×2 bottom-up one-bit DIB with the given colour table (none by default); each row's bits start at
    /// the left pixel.</summary>
    private static byte[] OneBit(byte top, byte bottom, params int[] table)
        => Le(40, 2, 2, 1 | 1 << 16, 0, 8, 0, 0, 0, 0).Concat(Le(table)).Concat(new byte[] { bottom, 0, 0, 0, top, 0, 0, 0 }).ToArray();

    /// <summary>EMF BITBLT without a source bitmap: a pattern blit.</summary>
    private static byte[] PatBlt(int x, int y, int cx, int cy, uint rop)
        => R(76, 0, 0, 0, 0, x, y, cx, cy, unchecked((int)rop), 0, 0, F(1), 0, 0, F(1), 0, 0, 0, 0, 0, 0, 0, 0);

    /// <summary>A 2×2 bottom-up 24-bit DIB of RGB colours.</summary>
    private static byte[] Rgb24(int topLeft, int topRight, int bottomLeft, int bottomRight)
    {
        byte[] Bgr(int c) => new[] { (byte)c, (byte)(c >> 8), (byte)(c >> 16) };
        return Le(40, 2, 2, 1 | 24 << 16, 0, 16, 0, 0, 0, 0)
            .Concat(Bgr(bottomLeft)).Concat(Bgr(bottomRight)).Concat(new byte[2]).Concat(Bgr(topLeft)).Concat(Bgr(topRight)).Concat(new byte[2]).ToArray();
    }

    private static byte[] Le(params int[] values) => values.SelectMany(BitConverter.GetBytes).ToArray();

    private static SKBitmap Png(XElement image) => SKBitmap.Decode(Convert.FromBase64String(((string)image.Attribute("href")!).Split(',')[1]));

    private static XElement Play(byte[] wmf) => XElement.Parse(Encoding.UTF8.GetString(WmfImage.ToSvg(wmf)));

    /// <summary>An EMF split across two META_ESCAPE_ENHANCED_METAFILE records that declare <paramref name="total"/> bytes.</summary>
    private static byte[][] Embedded(byte[] emf, int total)
    {
        int half = emf.Length / 4 * 2;
        byte[] Chunk(int from, int to) => W(0x0626, Words(new[] { 0x000f, 34 + to - from, 0x4d57, 0x4346, 1, 0, 0, 1, 0, 0, 0, 2, 0,
            (to - from) & 0xffff, (to - from) >> 16, (total - to) & 0xffff, (total - to) >> 16, total & 0xffff, total >> 16 }, emf[from..to]));
        return new[] { Chunk(0, half), Chunk(half, emf.Length) };
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
