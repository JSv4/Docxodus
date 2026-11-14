using System.Buffers.Binary;
using System.Text;
using System.Xml.Linq;
using static Docxodus.Internal.MetafileCanvas;

namespace Docxodus.Internal;

/// <summary>Bounded EMF playback to inline SVG through <see cref="MetafileCanvas"/>. Unsupported records
/// reject the whole image. EMF+ records are skipped when the file also carries its GDI drawing.</summary>
internal sealed record EmfImage(XElement Drawing)
{
    internal static EmfImage Read(byte[] bytes)
    {
        if (bytes.Length < 88 || U32(bytes, 0) != 1 || U32(bytes, 40) != 0x464d4520 || U32(bytes, 44) != 0x10000)
            throw new NotSupportedException("EMF requires a version-one header.");
        uint headerSize = U32(bytes, 4), records = U32(bytes, 52);
        int handles = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(56, 2));
        if (headerSize < 88 || headerSize % 4 != 0 || headerSize > bytes.Length || U32(bytes, 48) != bytes.Length
            || records < 2 || records > MaxRecords || handles < 1
            || I32(bytes, 16) < I32(bytes, 8) || I32(bytes, 20) < I32(bytes, 12))
            throw new InvalidDataException("Invalid or oversized EMF header.");
        double frameX = I32(bytes, 24), frameY = I32(bytes, 28);
        double frameWidth = I32(bytes, 32) - frameX, frameHeight = I32(bytes, 36) - frameY;
        if (frameWidth <= 0 || frameHeight <= 0 || I32(bytes, 72) <= 0 || I32(bytes, 76) <= 0
            || I32(bytes, 80) <= 0 || I32(bytes, 84) <= 0)
            throw new InvalidDataException("EMF requires positive frame and reference device dimensions.");
        // Device coordinates are reference-device pixels; Frame is in hundredths of a millimeter.
        double sx = 100.0 * I32(bytes, 80) / I32(bytes, 72), sy = 100.0 * I32(bytes, 84) / I32(bytes, 76);
        var drawing = new XElement(Svg + "g", new XAttribute("transform", $"scale({N(sx)} {N(sy)})"));
        var canvas = new MetafileCanvas(drawing, bytes,
            pixelsPerMm: ((double)I32(bytes, 72) / I32(bytes, 80), (double)I32(bytes, 76) / I32(bytes, 84)));
        var objects = new GdiObject?[handles];
        int count = 1;
        for (int offset = (int)headerSize; offset < bytes.Length;)
        {
            // LibreOffice counts the records after the header, one fewer than MS-EMF's nRecords.
            if (++count > records + 1 || bytes.Length - offset < 8)
                throw new InvalidDataException("EMF record count or truncated record.");
            uint kind = U32(bytes, offset), size = U32(bytes, offset + 4);
            if (size < 8 || size % 4 != 0 || size > bytes.Length - offset)
                throw new InvalidDataException("Invalid EMF record length.");
            int length = (int)size, at = offset;
            void Need(long expected) { if (length < expected) throw new InvalidDataException($"Invalid EMF record {kind} parameters."); }
            int I(int field) => I32(bytes, at + field);
            uint U(int field) => U32(bytes, at + field);
            switch (kind)
            {
                case 14: // EOF
                    Need(20);
                    if (offset + length != bytes.Length || count < records) throw new InvalidDataException("Invalid EMF end record.");
                    return Finish();
                case 2 or 3 or 4 or 5 or 6 or 85 or 86 or 87 or 88 or 89: // POLYBEZIER … POLYLINETO, and their 16-bit forms
                {
                    Need(28);
                    bool small = kind >= 85;
                    var pts = Points(bytes, at, length, 28, U(24), small);
                    canvas.Polygonal((kind - (small ? 85u : 2u)) switch {
                        0 => Poly.Bezier, 1 => Poly.Gon, 2 => Poly.Line, 3 => Poly.BezierTo, _ => Poly.LineTo }, pts);
                    break;
                }
                case 7 or 8 or 90 or 91: // POLYPOLYLINE, POLYPOLYGON, and their 16-bit forms
                {
                    Need(32);
                    uint polygons = U(24), total = U(28);
                    if (polygons > total + 1 || polygons > MaxRecords) throw new InvalidDataException("Invalid EMF poly-polygon counts.");
                    Need(32 + 4L * polygons);
                    var counts = new int[polygons];
                    long sum = 0;
                    for (int i = 0; i < counts.Length; i++) { counts[i] = (int)Math.Min(U(32 + i * 4), int.MaxValue); sum += counts[i]; }
                    if (sum != total) throw new InvalidDataException("EMF poly-polygon counts do not match its points.");
                    var pts = Points(bytes, at, length, 32 + 4 * (int)polygons, total, kind >= 90);
                    canvas.PolyPolygonal(kind is 8 or 91, pts, counts);
                    break;
                }
                case 56 or 92: // POLYDRAW, POLYDRAW16
                {
                    Need(28);
                    uint n = U(24);
                    bool small = kind == 92;
                    var pts = Points(bytes, at, length, 28, n, small);
                    int typesAt = 28 + (int)n * (small ? 4 : 8);
                    Need(typesAt + (long)n);
                    canvas.PolyDraw(pts, bytes.AsSpan(at + typesAt, (int)n));
                    break;
                }
                case 9 or 10 or 11 or 12: // SETWINDOWEXTEX, SETWINDOWORGEX, SETVIEWPORTEXTEX, SETVIEWPORTORGEX
                    Need(16);
                    if (kind == 9) canvas.SetWindowExtent(I(8), I(12));
                    else if (kind == 10) canvas.SetWindowOrigin(I(8), I(12));
                    else if (kind == 11) canvas.SetViewportExtent(I(8), I(12));
                    else canvas.SetViewportOrigin(I(8), I(12));
                    break;
                case 31 or 32: // SCALEVIEWPORTEXTEX, SCALEWINDOWEXTEX
                {
                    Need(24);
                    if (I(12) == 0 || I(20) == 0) throw new InvalidDataException("EMF extent scale divides by zero.");
                    var (x, y) = kind == 31 ? canvas.ViewportExtent : canvas.WindowExtent;
                    x = x * I(8) / I(12); y = y * I(16) / I(20);
                    if (kind == 31) canvas.SetViewportExtent(x, y); else canvas.SetWindowExtent(x, y);
                    break;
                }
                case 17: Need(12); canvas.SetMapMode(I(8)); break; // SETMAPMODE
                case 18: Need(12); canvas.SetBackgroundMode(I(8)); break; // SETBKMODE
                case 19: Need(12); canvas.SetPolyFillMode(I(8)); break; // SETPOLYFILLMODE
                case 22: Need(12); canvas.SetTextAlign(I(8)); break; // SETTEXTALIGN
                case 24: Need(12); canvas.SetTextColor(Color(U(8))); break; // SETTEXTCOLOR
                case 25: Need(12); canvas.SetBackgroundColor(Color(U(8))); break; // SETBKCOLOR
                case 57: Need(12); canvas.SetArcDirection(I(8)); break; // SETARCDIRECTION
                case 26: // OFFSETCLIPRGN
                    if (canvas.HasClip) throw new NotSupportedException("Offsetting an EMF clip region is not supported.");
                    break;
                case 27: Need(16); canvas.MoveTo(I(8), I(12)); break; // MOVETOEX
                case 54: Need(16); canvas.LineTo(I(8), I(12)); break; // LINETO
                case 29 or 30: // EXCLUDECLIPRECT, INTERSECTCLIPRECT
                    Need(24);
                    if (kind == 29) canvas.ExcludeClipRect(I(8), I(12), I(16), I(20));
                    else canvas.IntersectClipRect(I(8), I(12), I(16), I(20));
                    break;
                case 33: canvas.Save(); break; // SAVEDC
                case 34: Need(12); canvas.Restore(I(8)); break; // RESTOREDC
                case 35: Need(32); canvas.SetWorldTransform(Xform(bytes, at + 8)); break; // SETWORLDTRANSFORM
                case 36: Need(36); canvas.ModifyWorldTransform(Xform(bytes, at + 8), U(32)); break; // MODIFYWORLDTRANSFORM
                case 37 or 40: // SELECTOBJECT, DELETEOBJECT
                {
                    Need(12);
                    uint index = U(8);
                    if (kind == 37 && (index & 0x80000000) != 0)
                    {
                        GdiObject? stock = (index & 0x7fffffff) switch {
                            0 or 18 => WhiteBrush, 1 => new Brush("#c0c0c0"), 2 => new Brush("#808080"), 3 => new Brush("#404040"),
                            4 => new Brush("#000000"), 5 => NullBrush,
                            6 => new Pen("#ffffff", 1, true, null, "round", "round"), 7 or 19 => BlackPen, 8 => NullPen,
                            10 or 11 or 12 or 13 or 14 or 16 or 17 => DefaultFont, 15 => null,
                            _ => throw new NotSupportedException("Unsupported EMF stock object.") };
                        if (stock != null) canvas.Select(stock);
                        break;
                    }
                    if (index == 0 || index >= handles || objects[index] is not { } selected)
                        throw new InvalidDataException("Invalid EMF object selection.");
                    if (kind == 40) objects[index] = null;
                    else canvas.Select(selected);
                    break;
                }
                case 38: // CREATEPEN
                {
                    Need(28);
                    uint style = U(12);
                    int width = I(16);
                    // Width zero, or one without PS_GEOMETRIC, is a one-device-pixel cosmetic pen.
                    bool cosmetic = width <= 0 || (width == 1 && (style & 0x10000) == 0);
                    Store(objects, U(8), canvas.CreatePen(style & 0xffff, width, Color(U(24)), cosmetic));
                    break;
                }
                case 95: // EXTCREATEPEN
                {
                    Need(52);
                    uint style = U(28), brushStyle = U(36), entries = U(48);
                    if (entries > 16) throw new InvalidDataException("Invalid EMF pen style entry count.");
                    Need(52 + 4L * entries);
                    var dashes = new double[entries];
                    for (int i = 0; i < dashes.Length; i++) dashes[i] = U(52 + i * 4);
                    string color = brushStyle switch { 0 or 2 => Color(U(40)), 1 => "none",
                        _ => throw new NotSupportedException("EMF pattern pens are not supported.") };
                    double width = U(32);
                    bool cosmetic = (style & 0x10000) == 0 || width == 0;
                    var pen = color == "none" ? NullPen : canvas.CreatePen(style, width, color, cosmetic, dashes);
                    Store(objects, U(8), pen);
                    break;
                }
                case 39: // CREATEBRUSHINDIRECT
                    Need(24);
                    Store(objects, U(8), CreateBrush(U(12), U(12) == 1 ? "none" : Color(U(16)), U(20)));
                    break;
                case 93: Need(12); Store(objects, U(8), new Brush("none", Pattern: true)); break; // CREATEMONOBRUSH
                case 94: // CREATEDIBPATTERNBRUSHPT: only RGB colours (DIB_RGB_COLORS) can be drawn
                    Need(32);
                    Store(objects, U(8), U(12) != 0 ? new Brush("none", Pattern: true)
                        : canvas.PatternBrush(Part(bytes, at, length, U(16), U(20)), Part(bytes, at, length, U(24), U(28))));
                    break;
                case 49 or 99 or 122: Need(12); Store(objects, U(8), new OtherObject()); break; // CREATEPALETTE, CREATECOLORSPACE(W)
                case 101: // DELETECOLORSPACE
                    Need(12);
                    if (U(8) < handles) objects[U(8)] = null;
                    break;
                case 82: // EXTCREATEFONTINDIRECTW
                {
                    Need(104);
                    var face = Encoding.Unicode.GetString(bytes, at + 40, 64);
                    Store(objects, U(8), CreateFont(face, I(12), I(20), I(28), bytes[at + 32] != 0, bytes[at + 33] != 0, bytes[at + 34] != 0)
                        with { Charset = bytes[at + 35] });
                    break;
                }
                case 83 or 84: // EXTTEXTOUTA, EXTTEXTOUTW
                {
                    Need(76);
                    uint chars = U(44), stringAt = U(48), options = U(52), dxAt = U(72);
                    if ((options & 0x10) != 0 && chars > 0) throw new NotSupportedException("EMF glyph-index text is not supported.");
                    bool wide = kind == 84;
                    if (chars > MaxRecords || stringAt > length || stringAt + chars * (wide ? 2L : 1L) > length)
                        throw new InvalidDataException("Invalid EMF text record.");
                    string text = wide ? Encoding.Unicode.GetString(bytes, at + (int)stringAt, (int)chars * 2)
                        : canvas.Ansi(bytes.AsSpan(at + (int)stringAt, (int)chars));
                    double[]? advances = null;
                    if (dxAt != 0 && chars > 0)
                    {
                        int stride = (options & 0x2000) != 0 ? 8 : 4; // ETO_PDY stores x,y pairs
                        if (dxAt > length || dxAt + (long)chars * stride > length) throw new InvalidDataException("Invalid EMF text spacing.");
                        advances = new double[chars];
                        for (int i = 0; i < advances.Length; i++) advances[i] = I((int)dxAt + i * stride);
                    }
                    (double, double, double, double)? opaque = (options & 0x2) != 0 ? (I(56), I(60), I(64), I(68)) : null;
                    canvas.Text(I(36), I(40), text, advances, opaque);
                    break;
                }
                case 41: // ANGLEARC
                    Need(28);
                    canvas.AngleArc(I(8), I(12), U(16), F32(bytes, at + 20), F32(bytes, at + 24));
                    break;
                case 42: Need(24); canvas.Ellipse(I(8), I(12), I(16), I(20)); break; // ELLIPSE
                case 43: Need(24); canvas.Rectangle(I(8), I(12), I(16), I(20)); break; // RECTANGLE
                case 44: Need(32); canvas.Rectangle(I(8), I(12), I(16), I(20), I(24), I(28)); break; // ROUNDRECT
                case 45 or 46 or 47 or 55: // ARC, CHORD, PIE, ARCTO
                    Need(40);
                    canvas.Arc(kind switch { 45 => ArcKind.Arc, 46 => ArcKind.Chord, 47 => ArcKind.Pie, _ => ArcKind.ArcTo },
                        I(8), I(12), I(16), I(20), I(24), I(28), I(32), I(36));
                    break;
                case 59: canvas.BeginPath(); break; // BEGINPATH
                case 60: canvas.EndPath(); break; // ENDPATH
                case 61: canvas.CloseFigure(); break; // CLOSEFIGURE
                case 62: canvas.DrawPath(true, false); break; // FILLPATH
                case 63: canvas.DrawPath(true, true); break; // STROKEANDFILLPATH
                case 64: canvas.DrawPath(false, true); break; // STROKEPATH
                case 67: Need(12); canvas.SelectClipPath(I(8)); break; // SELECTCLIPPATH
                case 68: canvas.AbortPath(); break; // ABORTPATH
                case 70: // COMMENT
                    Need(12);
                    Comment(bytes, at, length, U(8));
                    break;
                case 75: // EXTSELECTCLIPRGN
                {
                    Need(16);
                    uint dataSize = U(8);
                    int mode = I(12);
                    if (dataSize == 0)
                    {
                        if (mode != 5) throw new NotSupportedException("Unsupported EMF clip-region combination.");
                        canvas.ResetClip();
                        break;
                    }
                    Need(16 + (long)dataSize);
                    if (dataSize < 32 || U(16) != 32 || U(20) != 1) throw new InvalidDataException("Invalid EMF clip region.");
                    uint rects = U(24);
                    if (32 + 16L * rects > dataSize) throw new InvalidDataException("Invalid EMF clip region.");
                    var list = new (double, double, double, double)[rects];
                    for (int i = 0; i < list.Length; i++)
                    {
                        int r = 48 + i * 16;
                        list[i] = (I(r), I(r + 4), I(r + 8), I(r + 12));
                    }
                    canvas.ClipToDeviceRects(list, mode);
                    break;
                }
                case 76 or 77: // BITBLT, STRETCHBLT
                {
                    Need(kind == 76 ? 100 : 108);
                    int cx = I(32), cy = I(36);
                    int cxSrc = kind == 77 ? I(100) : cx, cySrc = kind == 77 ? I(104) : cy;
                    Blit(canvas, bytes, at, length, I(24), I(28), cx, cy, U(40), I(44), I(48), cxSrc, cySrc,
                        U(80), U(84), U(88), U(92), U(96), fromBottom: false);
                    break;
                }
                case 114: // ALPHABLEND: STRETCHBLT's layout with a blend function in place of the raster operation
                {
                    Need(108);
                    uint blend = U(40);
                    // BlendFlags (the second byte) should be zero; Word writes 0x80 there, and it changes nothing.
                    if ((blend & 0xff) != 0 || (blend >> 24) > 1) throw new NotSupportedException("Unsupported EMF blend function.");
                    if (U(88) == 0) throw new InvalidDataException("EMF ALPHABLEND requires a source bitmap.");
                    Blit(canvas, bytes, at, length, I(24), I(28), I(32), I(36), null, I(44), I(48), I(100), I(104),
                        U(80), U(84), U(88), U(92), U(96), fromBottom: false, (byte)(blend >> 16), (blend >> 24) == 1);
                    break;
                }
                case 81: // STRETCHDIBITS
                    Need(80);
                    Blit(canvas, bytes, at, length, I(24), I(28), I(72), I(76), U(68), I(32), I(36), I(40), I(44),
                        U(64), U(48), U(52), U(56), U(60), fromBottom: true);
                    break;
                case 115: // SETLAYOUT
                    Need(12);
                    if (U(8) != 0) throw new NotSupportedException("Mirrored EMF layout is not supported.");
                    break;
                case 20: Need(12); canvas.SetBinaryRasterOperation(U(8)); break; // SETROP2
                case 120: // SETTEXTJUSTIFICATION: break extra, then break count
                    Need(16);
                    if (I(8) != 0) throw new NotSupportedException("EMF text spacing adjustments are not supported.");
                    break;
                case 13 or 16 or 21 or 23 or 28 or 48 or 50 or 51 or 52 or 58 or 65 or 98 or 100 or 112 or 113 or 121:
                    // Brush origin, mapper flags, stretch mode, colour adjustment, meta region, palettes,
                    // miter limit, FLATTENPATH and ICM do not change the SVG we draw.
                    break;
                default:
                    throw new NotSupportedException($"Unsupported EMF record {kind}.");
            }
            offset += length;
        }
        throw new InvalidDataException("EMF is missing EOF.");

        EmfImage Finish()
        {
            canvas.Finish();
            double minX = frameX, minY = frameY, maxX = frameX + frameWidth, maxY = frameY + frameHeight;
            if (canvas.Bounds is { } b)
            {
                minX = Math.Min(minX, b.MinX * sx); minY = Math.Min(minY, b.MinY * sy);
                maxX = Math.Max(maxX, b.MaxX * sx); maxY = Math.Max(maxY, b.MaxY * sy);
            }
            double width = maxX - minX, height = maxY - minY, pxWidth = width * 96 / 2540 + 1, pxHeight = height * 96 / 2540 + 1;
            if (pxWidth > 32768 || pxHeight > 32768 || pxWidth * pxHeight > 16 * 1024 * 1024)
                throw new InvalidDataException("EMF intrinsic size exceeds the 16-megapixel limit.");
            // Frame establishes mapping, not a clip. Inline SVG keeps overflow visible and
            // cosmetic strokes one CSS pixel wide, including when the picture is resized.
            return new EmfImage(new XElement(Svg + "svg", new XAttribute("width", N(frameWidth * 96 / 2540)),
                new XAttribute("height", N(frameHeight * 96 / 2540)), new XAttribute("data-docx-vector-frame", "emf"),
                new XAttribute("viewBox", $"{N(frameX)} {N(frameY)} {N(frameWidth)} {N(frameHeight)}"),
                new XAttribute("preserveAspectRatio", "none"), new XAttribute("overflow", "visible"), drawing));
        }
    }

    /// <summary>GDI comments carry no drawing, except EMF+: an EMF+-only file has no GDI fallback to play.</summary>
    private static void Comment(byte[] bytes, int at, int length, uint dataSize)
    {
        if (dataSize > length - 12) throw new InvalidDataException("Invalid EMF comment size.");
        if (dataSize < 4 || U32(bytes, at + 12) != 0x2B464D45) return; // "EMF+"
        int end = at + 12 + (int)dataSize;
        for (int record = at + 16; end - record >= 12;)
        {
            ushort type = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(record, 2));
            ushort flags = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(record + 2, 2));
            uint size = U32(bytes, record + 4);
            if (size < 12 || size > end - record) throw new InvalidDataException("Invalid EMF+ record.");
            if (type == 0x4001 && (flags & 1) == 0) // EmfPlusHeader without the dual flag
                throw new NotSupportedException("EMF+-only metafiles are not supported.");
            record += (int)size;
        }
    }

    /// <summary>BITBLT, STRETCHBLT, STRETCHDIBITS and ALPHABLEND (<paramref name="rop"/> null). Without a source
    /// bitmap only the pattern and constant raster operations draw.</summary>
    private static void Blit(MetafileCanvas canvas, byte[] bytes, int at, int length, int x, int y, int cx, int cy, uint? rop,
        int xSrc, int ySrc, int cxSrc, int cySrc, uint usage, uint infoAt, uint infoSize, uint bitsAt, uint bitsSize, bool fromBottom,
        byte constantAlpha = 255, bool sourceAlpha = false)
    {
        if (infoSize == 0)
        {
            string? color = rop switch { 0x00F00021 => null, 0x00000042 => "#000000", 0x00FF0062 => "#ffffff",
                0x00AA0029 => "", _ => throw new NotSupportedException("Unsupported EMF raster operation.") };
            if (color != "") canvas.FillRectangle(x, y, (double)x + cx, (double)y + cy, color);
            return;
        }
        if (usage != 0) throw new NotSupportedException("EMF bitmaps with palette colours are not supported.");
        var info = Part(bytes, at, length, infoAt, infoSize);
        var bits = Part(bytes, at, length, bitsAt, bitsSize);
        canvas.ReserveBitmap(Math.Abs((long)cxSrc * cySrc));
        if (DibImage.Decode(info, bits, xSrc, ySrc, cxSrc, cySrc, fromBottom, constantAlpha, sourceAlpha) is { } image)
            canvas.Bitmap(x, y, cx, cy, image, rop);
    }

    /// <summary>The bytes a record locates by offset and size from its own start.</summary>
    private static ReadOnlySpan<byte> Part(byte[] bytes, int at, int length, uint offset, uint size)
    {
        if (offset > length || size > length - offset) throw new InvalidDataException("EMF bitmap lies outside its record.");
        return bytes.AsSpan(at + (int)offset, (int)size);
    }

    private static (double X, double Y)[] Points(byte[] bytes, int at, int length, int first, uint count, bool small)
    {
        int size = small ? 4 : 8;
        if (count > MaxPoints || first + (long)count * size > length) throw new InvalidDataException("EMF point list exceeds its record.");
        var points = new (double, double)[count];
        for (int i = 0; i < points.Length; i++)
        {
            int p = at + first + i * size;
            points[i] = small ? (BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(p, 2)), BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(p + 2, 2)))
                : (I32(bytes, p), I32(bytes, p + 4));
        }
        return points;
    }

    private static void Store(GdiObject?[] objects, uint slot, GdiObject value)
    {
        if (slot == 0 || slot >= objects.Length) throw new InvalidDataException("Invalid EMF object slot.");
        objects[slot] = value;
    }

    private static Matrix Xform(byte[] bytes, int at)
        => new(F32(bytes, at), F32(bytes, at + 4), F32(bytes, at + 8), F32(bytes, at + 12), F32(bytes, at + 16), F32(bytes, at + 20));

    /// <summary>A COLORREF. Palette-relative RGB (0x02) draws as RGB; palette indexes are unsupported.</summary>
    internal static string Color(uint value)
    {
        if ((value >> 24) is not (0 or 2)) throw new NotSupportedException("Metafile palette colors are not supported.");
        return $"#{value & 255:x2}{(value >> 8) & 255:x2}{(value >> 16) & 255:x2}";
    }

    /// <summary>Decodes ANSI metafile text as Windows-1252, which .NET does not ship by default.</summary>
    internal static string Windows1252(ReadOnlySpan<byte> bytes)
    {
        const string high = "€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008dŽ\u008f"
            + "\u0090‘’“”•–—˜™š›œ\u009džŸ";
        var text = new StringBuilder(bytes.Length);
        foreach (byte b in bytes) text.Append(b is >= 0x80 and < 0xa0 ? high[b - 0x80] : (char)b);
        return text.ToString();
    }

    private static uint U32(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
    private static int I32(byte[] bytes, int offset) => BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset, 4));
    private static double F32(byte[] bytes, int offset)
    {
        float value = BinaryPrimitives.ReadSingleLittleEndian(bytes.AsSpan(offset, 4));
        if (!float.IsFinite(value)) throw new InvalidDataException("Invalid EMF floating-point value.");
        // The double nearest the float's shortest decimal form: 1e-5f is 1E-05, not 9.999999747E-06.
        return double.Parse(value.ToString(System.Globalization.CultureInfo.InvariantCulture), System.Globalization.CultureInfo.InvariantCulture);
    }
}
