using System.Buffers.Binary;
using System.Text;
using System.Xml.Linq;
using static Docxodus.Internal.MetafileCanvas;

namespace Docxodus.Internal;

/// <summary>Bounded placeable-WMF playback to a static SVG through <see cref="MetafileCanvas"/>. Unknown
/// records reject the whole image, never partial art.</summary>
internal static class WmfImage
{
    internal static byte[] ToSvg(byte[] bytes)
    {
        if (bytes.Length < 46 || U32(bytes, 0) != 0x9ac6cdd7 || U16(bytes, 4) != 0 || U32(bytes, 16) != 0)
            throw new NotSupportedException("WMF requires a valid placeable header.");
        ushort checksum = 0;
        for (int i = 0; i < 20; i += 2) checksum ^= U16(bytes, i);
        if (checksum != U16(bytes, 20)) throw new InvalidDataException("Invalid WMF placeable checksum.");
        int left = I16(bytes, 6), top = I16(bytes, 8), width = I16(bytes, 10) - left, height = I16(bytes, 12) - top;
        int inch = U16(bytes, 14), objects = U16(bytes, 32);
        if (width <= 0 || height <= 0 || inch == 0 || objects > 1024
            || U16(bytes, 22) is not (1 or 2) || U16(bytes, 24) != 9 || U16(bytes, 26) is not (0x100 or 0x300)
            || (ulong)U32(bytes, 28) * 2 != (ulong)bytes.Length - 22 || U16(bytes, 38) != 0)
            throw new InvalidDataException("Invalid or oversized WMF header.");
        double viewportWidth = width * 96.0 / inch, viewportHeight = height * 96.0 / inch;
        double pixelWidth = Math.Ceiling(viewportWidth), pixelHeight = Math.Ceiling(viewportHeight);
        if (pixelWidth > 32768 || pixelHeight > 32768 || pixelWidth * pixelHeight > 16 * 1024 * 1024)
            throw new InvalidDataException("WMF intrinsic size exceeds the 16-megapixel limit.");
        var svg = new XElement(Svg + "svg", new XAttribute("width", N(pixelWidth)), new XAttribute("height", N(pixelHeight)),
            new XAttribute("viewBox", $"0 0 {N(viewportWidth)} {N(viewportHeight)}"), new XAttribute("preserveAspectRatio", "none"));
        // The placeable header fixes the device box; the window always maps onto it. Word draws visible
        // zero-width pens as printer hairlines, which have no resolution-independent width.
        var canvas = new MetafileCanvas(svg, bytes, fixedDevice: (viewportWidth, viewportHeight), rejectCosmeticPens: true, defaultJoin: "miter");
        canvas.SetWindowOrigin(left, top);
        canvas.SetWindowExtent(width, height);
        var table = new GdiObject?[objects];
        uint maxRecord = U32(bytes, 34);
        int recordCount = 0;
        for (int offset = 40; offset < bytes.Length;)
        {
            if (++recordCount > MaxRecords || bytes.Length - offset < 6) throw new InvalidDataException("WMF record limit or truncated record.");
            uint words = U32(bytes, offset);
            if (words < 3 || words > maxRecord || (ulong)words * 2 > (uint)(bytes.Length - offset))
                throw new InvalidDataException("Invalid WMF record length.");
            int length = checked((int)words * 2), function = U16(bytes, offset + 4), start = offset + 6, parameters = length - 6;
            void Need(int count) { if (parameters < count * 2) throw new InvalidDataException($"Invalid WMF record 0x{function:X4} parameters."); }
            int P(int index) => I16(bytes, start + index * 2);
            ushort W(int index) => U16(bytes, start + index * 2);
            switch (function)
            {
                case 0: // EOF
                    if (offset + length != bytes.Length) throw new InvalidDataException("WMF data follows EOF.");
                    canvas.Finish();
                    return Encoding.UTF8.GetBytes(svg.ToXmlString(SaveOptions.DisableFormatting));
                case 0x0103: Need(1); canvas.SetMapMode(P(0)); break; // SETMAPMODE: the placeable box still fixes the mapping
                case 0x020b: Need(2); canvas.SetWindowOrigin(P(1), P(0)); break; // SETWINDOWORG
                case 0x020c: Need(2); canvas.SetWindowExtent(P(1), P(0)); break; // SETWINDOWEXT
                case 0x020f: // OFFSETWINDOWORG
                {
                    Need(2);
                    var (x, y) = canvas.WindowOrigin;
                    canvas.SetWindowOrigin(x + P(1), y + P(0));
                    break;
                }
                case 0x0410: // SCALEWINDOWEXT
                {
                    Need(4);
                    if (P(0) == 0 || P(2) == 0) throw new InvalidDataException("WMF extent scale divides by zero.");
                    var (x, y) = canvas.WindowExtent;
                    canvas.SetWindowExtent(x * P(3) / P(2), y * P(1) / P(0));
                    break;
                }
                case 0x0102: Need(1); canvas.SetBackgroundMode(W(0)); break; // SETBKMODE
                case 0x0201: Need(2); canvas.SetBackgroundColor(EmfImage.Color(U32(bytes, start))); break; // SETBKCOLOR
                case 0x0209: Need(2); canvas.SetTextColor(EmfImage.Color(U32(bytes, start))); break; // SETTEXTCOLOR
                case 0x012e: Need(1); canvas.SetTextAlign(W(0)); break; // SETTEXTALIGN
                case 0x0106: Need(1); canvas.SetPolyFillMode(W(0)); break; // SETPOLYFILLMODE
                case 0x001e: canvas.Save(); break; // SAVEDC
                case 0x0127: Need(1); canvas.Restore(P(0)); break; // RESTOREDC
                case 0x02fa: // CREATEPENINDIRECT
                {
                    Need(5);
                    int penWidth = Math.Abs(P(1));
                    Store(table, canvas.CreatePen(W(0), penWidth, EmfImage.Color(U32(bytes, start + 6)), cosmetic: penWidth == 0));
                    break;
                }
                case 0x02fc: // CREATEBRUSHINDIRECT
                    Need(4);
                    Store(table, CreateBrush(W(0), W(0) == 1 ? "none" : EmfImage.Color(U32(bytes, start + 2)), W(3)));
                    break;
                case 0x02fb: // CREATEFONTINDIRECT
                {
                    Need(9);
                    int faceLength = Math.Min(32, parameters - 18);
                    var face = EmfImage.Windows1252(bytes.AsSpan(start + 18, Math.Max(0, faceLength)));
                    Store(table, CreateFont(face, P(0), P(2), P(4), bytes[start + 10] != 0, bytes[start + 11] != 0, bytes[start + 12] != 0)
                        with { Charset = bytes[start + 13] });
                    break;
                }
                case 0x01f9 or 0x0142: Store(table, new Brush("none", Pattern: true)); break; // CREATEPATTERNBRUSH, DIBCREATEPATTERNBRUSH
                case 0x00f7 or 0x06ff: Store(table, new OtherObject()); break; // CREATEPALETTE, CREATEREGION
                case 0x012d or 0x01f0: // SELECTOBJECT, DELETEOBJECT
                {
                    Need(1);
                    int index = W(0);
                    if (index >= table.Length || table[index] is not { } value) throw new InvalidDataException("Invalid WMF object selection.");
                    if (function == 0x01f0) table[index] = null;
                    else canvas.Select(value);
                    break;
                }
                case 0x0214: Need(2); canvas.MoveTo(P(1), P(0)); break; // MOVETO
                case 0x0213: Need(2); canvas.LineTo(P(1), P(0)); break; // LINETO
                case 0x0324 or 0x0325: // POLYGON, POLYLINE
                {
                    Need(1);
                    var pts = Points(bytes, start + 2, W(0), parameters - 2);
                    canvas.Polygonal(function == 0x0324 ? Poly.Gon : Poly.Line, pts);
                    break;
                }
                case 0x0538: // POLYPOLYGON
                {
                    Need(1);
                    int polygons = W(0);
                    Need(1 + polygons);
                    var counts = new int[polygons];
                    long total = 0; // 65,535 counts of 65,535 points do not fit an int
                    for (int i = 0; i < polygons; i++) total += counts[i] = W(1 + i);
                    var pts = Points(bytes, start + 2 + polygons * 2, total, parameters - 2 - polygons * 2);
                    canvas.PolyPolygonal(true, pts, counts);
                    break;
                }
                case 0x041b or 0x0418 or 0x061c: // RECTANGLE, ELLIPSE, ROUNDRECT
                {
                    bool round = function == 0x061c;
                    int o = round ? 2 : 0;
                    Need(4 + o);
                    int bottom = P(o), right = P(o + 1), y = P(o + 2), x = P(o + 3);
                    if (right <= x || bottom <= y) throw new NotSupportedException("WMF requires nonempty shape bounds.");
                    var m = canvas.Transform;
                    // Word excludes one device pixel after truncating its logical-coordinate size.
                    double shapeWidth = right - x - Math.Truncate(1 / Math.Abs(m.A)), shapeHeight = bottom - y - Math.Truncate(1 / Math.Abs(m.D));
                    if (shapeWidth <= 0 || shapeHeight <= 0) throw new NotSupportedException("WMF collapsed shape bounds are not supported.");
                    if (function == 0x0418) canvas.Ellipse(x, y, x + shapeWidth, y + shapeHeight);
                    else canvas.Rectangle(x, y, x + shapeWidth, y + shapeHeight, round ? P(1) : 0, round ? P(0) : 0);
                    break;
                }
                case 0x0817 or 0x0830 or 0x081a: // ARC, CHORD, PIE
                    Need(8);
                    canvas.Arc(function switch { 0x0817 => ArcKind.Arc, 0x0830 => ArcKind.Chord, _ => ArcKind.Pie },
                        P(7), P(6), P(5), P(4), P(3), P(2), P(1), P(0));
                    break;
                case 0x0416 or 0x0415: // INTERSECTCLIPRECT, EXCLUDECLIPRECT
                    Need(4);
                    if (function == 0x0416) canvas.IntersectClipRect(P(3), P(2), P(1), P(0));
                    else canvas.ExcludeClipRect(P(3), P(2), P(1), P(0));
                    break;
                case 0x0521: // TEXTOUT
                {
                    Need(1);
                    int count = W(0), padded = (count + 1) & ~1;
                    Need(1 + padded / 2 + 2);
                    canvas.Text(P(2 + padded / 2), P(1 + padded / 2), canvas.Ansi(bytes.AsSpan(start + 2, count)), null, null);
                    break;
                }
                case 0x0a32: // EXTTEXTOUT
                {
                    Need(4);
                    int count = W(2), options = W(3), at = 8;
                    (double, double, double, double)? opaque = null;
                    if ((options & 0x6) != 0)
                    {
                        Need(8);
                        if ((options & 0x2) != 0) opaque = (P(4), P(5), P(6), P(7));
                        at += 8;
                    }
                    if ((options & 0x10) != 0) throw new NotSupportedException("WMF glyph-index text is not supported.");
                    int padded = (count + 1) & ~1;
                    if (parameters < at + padded) throw new InvalidDataException("Invalid WMF text record.");
                    double[]? advances = null;
                    if (parameters >= at + padded + count * 2 && count > 0)
                    {
                        advances = new double[count];
                        for (int i = 0; i < count; i++) advances[i] = I16(bytes, start + at + padded + i * 2);
                    }
                    canvas.Text(P(1), P(0), canvas.Ansi(bytes.AsSpan(start + at, count)), advances, opaque);
                    break;
                }
                case 0x061d: // PATBLT
                    Need(6);
                    Fill(canvas, U32(bytes, start), P(5), P(4), P(3), P(2));
                    break;
                case 0x0f43: // STRETCHDIB
                    Need(11);
                    Blit(canvas, U32(bytes, start), W(2), P(10), P(9), P(8), P(7), P(6), P(5), P(4), P(3),
                        bytes.AsSpan(start + 22, parameters - 22), fromBottom: true);
                    break;
                case 0x0b41 or 0x0940: // DIBSTRETCHBLT, DIBBITBLT
                {
                    bool stretch = function == 0x0b41;
                    if (words == (function >> 8) + 3) // no source bitmap
                    {
                        Need(stretch ? 11 : 9);
                        int o = stretch ? 7 : 5;
                        Fill(canvas, U32(bytes, start), P(o + 3), P(o + 2), P(o + 1), P(o));
                        break;
                    }
                    int d = stretch ? 10 : 8;
                    Need(d);
                    int destHeight = P(d - 4), destWidth = P(d - 3);
                    int srcWidth = stretch ? P(3) : destWidth, srcHeight = stretch ? P(2) : destHeight;
                    int ySrc = stretch ? P(4) : P(2), xSrc = stretch ? P(5) : P(3);
                    Blit(canvas, U32(bytes, start), 0, P(d - 1), P(d - 2), destWidth, destHeight, xSrc, ySrc, srcWidth, srcHeight,
                        bytes.AsSpan(start + d * 2, parameters - d * 2), fromBottom: false);
                    break;
                }
                case 0x0149: // SETLAYOUT
                    Need(1);
                    if (W(0) != 0) throw new NotSupportedException("Mirrored WMF layout is not supported.");
                    break;
                case 0x0104: // SETROP2: only R2_COPYPEN draws the pen and brush as they are
                    Need(1);
                    if (W(0) != 13) throw new NotSupportedException("Unsupported WMF binary raster operation.");
                    break;
                case 0x0108 or 0x020a: // SETTEXTCHAREXTRA; SETTEXTJUSTIFICATION (break count, then break extra)
                    Need(function == 0x0108 ? 1 : 2);
                    if (P(function == 0x0108 ? 0 : 1) != 0) throw new NotSupportedException("WMF text spacing adjustments are not supported.");
                    break;
                case 0x020d or 0x020e or 0x0211 or 0x0412:
                    // Viewport records: the placeable header already fixes the device box.
                case 0x0105 or 0x0107 or 0x0231 or 0x0234 or 0x0035 or 0x0037 or 0x0139 or 0x0436 or 0x0626:
                    // Relabs, stretch mode, mapper flags, palettes and escapes do not change the SVG we draw.
                    break;
                default:
                    throw new NotSupportedException($"Unsupported WMF record 0x{function:X4}.");
            }
            offset += length;
        }
        throw new InvalidDataException("WMF is missing EOF.");
    }

    private static void Fill(MetafileCanvas canvas, uint rop, int x, int y, int width, int height)
    {
        string? color = rop switch { 0x00F00021 => null, 0x00000042 => "#000000", 0x00FF0062 => "#ffffff",
            0x00AA0029 => "", _ => throw new NotSupportedException("Unsupported WMF raster operation.") };
        if (color != "") canvas.FillRectangle(x, y, (double)x + width, (double)y + height, color);
    }

    private static void Blit(MetafileCanvas canvas, uint rop, int usage, int x, int y, int width, int height,
        int xSrc, int ySrc, int srcWidth, int srcHeight, ReadOnlySpan<byte> dib, bool fromBottom)
    {
        if (rop != 0x00CC0020 || usage != 0) throw new NotSupportedException("WMF bitmaps support SRCCOPY with RGB colours only.");
        int info = DibImage.InfoSize(dib);
        if (info > dib.Length) throw new InvalidDataException("Truncated WMF bitmap.");
        canvas.ReserveBitmap(Math.Abs((long)srcWidth * srcHeight));
        if (DibImage.Decode(dib[..info], dib[info..], xSrc, ySrc, srcWidth, srcHeight, fromBottom) is { } image)
            canvas.Image(x, y, width, height, image.Href, image.Pixels);
    }

    private static (double X, double Y)[] Points(byte[] bytes, int at, long count, int available)
    {
        if (count > MaxPoints || count * 4 > available) throw new InvalidDataException("WMF point list exceeds its record.");
        var points = new (double, double)[count];
        for (int i = 0; i < count; i++) points[i] = (I16(bytes, at + i * 4), I16(bytes, at + i * 4 + 2));
        return points;
    }

    private static void Store(GdiObject?[] table, GdiObject value)
    {
        int slot = Array.IndexOf(table, null);
        if (slot < 0) throw new InvalidDataException("WMF object table is full.");
        table[slot] = value;
    }

    private static ushort U16(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));
    private static short I16(byte[] bytes, int offset) => BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(offset, 2));
    private static uint U32(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
}
