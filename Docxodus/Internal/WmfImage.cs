using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>Static placeable WMF subset. Unknown records reject the whole image, never partial art.</summary>
internal static class WmfImage
{
    private sealed record Paint(bool IsPen, string Color, int Width = 0);

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
            || U16(bytes, 22) is not (1 or 2) || U16(bytes, 24) != 9 || U16(bytes, 26) != 0x300
            || (ulong)U32(bytes, 28) * 2 != (ulong)bytes.Length - 22 || U16(bytes, 38) != 0)
            throw new InvalidDataException("Invalid or oversized WMF header.");
        double viewportWidth = width * 96.0 / inch, viewportHeight = height * 96.0 / inch;
        double pixelWidth = Math.Ceiling(viewportWidth), pixelHeight = Math.Ceiling(viewportHeight);
        if (pixelWidth > 32768 || pixelHeight > 32768 || pixelWidth * pixelHeight > 16 * 1024 * 1024)
            throw new InvalidDataException("WMF intrinsic size exceeds the 16-megapixel limit.");
        XNamespace ns = "http://www.w3.org/2000/svg";
        var svg = new XElement(ns + "svg", new XAttribute("width", N(pixelWidth)), new XAttribute("height", N(pixelHeight)),
            new XAttribute("viewBox", $"0 0 {N(viewportWidth)} {N(viewportHeight)}"), new XAttribute("preserveAspectRatio", "none"));
        var table = new Paint?[objects];
        Paint pen = new(true, "#000000"), brush = new(false, "#ffffff");
        int originX = left, originY = top, windowWidth = width, windowHeight = height, recordCount = 0;
        for (int offset = 40; offset < bytes.Length;)
        {
            if (++recordCount > 10000 || bytes.Length - offset < 6) throw new InvalidDataException("WMF record limit or truncated record.");
            uint words = U32(bytes, offset);
            if (words < 3 || words > U32(bytes, 34) || (ulong)words * 2 > (uint)(bytes.Length - offset))
                throw new InvalidDataException("Invalid WMF record length.");
            int length = checked((int)words * 2), function = U16(bytes, offset + 4), start = offset + 6;
            void Size(int expected) { if (length != expected) throw new InvalidDataException("Invalid WMF record parameters."); }
            switch (function)
            {
                case 0: // EOF
                    Size(6);
                    if (offset + length != bytes.Length) throw new InvalidDataException("WMF data follows EOF.");
                    return Encoding.UTF8.GetBytes(svg.ToString(SaveOptions.DisableFormatting));
                case 0x0103: // SETMAPMODE
                    Size(8);
                    if (I16(bytes, start) != 8) throw new NotSupportedException("WMF requires anisotropic mapping.");
                    break;
                case 0x020b: // SETWINDOWORG
                    Size(10); originY = I16(bytes, start); originX = I16(bytes, start + 2); break;
                case 0x020c: // SETWINDOWEXT
                    Size(10); windowHeight = I16(bytes, start); windowWidth = I16(bytes, start + 2);
                    if (windowWidth <= 0 || windowHeight <= 0) throw new NotSupportedException("WMF requires positive window extents.");
                    break;
                case 0x02fa: // CREATEPENINDIRECT
                case 0x02fc: // CREATEBRUSHINDIRECT
                    bool isPen = function == 0x02fa;
                    Size(isPen ? 16 : 14);
                    int style = U16(bytes, start), slot = Array.FindIndex(table, value => value == null);
                    if (slot < 0) throw new InvalidDataException("WMF object table is full.");
                    if (isPen ? style is not (0 or 5) : style is not (0 or 1))
                        throw new NotSupportedException("WMF supports solid or null pens and brushes.");
                    string color = Color(U32(bytes, start + (isPen ? 6 : 2)));
                    table[slot] = new Paint(isPen, style == 0 ? color : "none", isPen ? Math.Abs((int)I16(bytes, start + 2)) : 0);
                    break;
                case 0x012d: // SELECTOBJECT
                case 0x01f0: // DELETEOBJECT
                    Size(8);
                    int index = U16(bytes, start);
                    if (index >= table.Length || table[index] is not { } paint) throw new InvalidDataException("Invalid WMF object selection.");
                    if (function == 0x01f0) table[index] = null;
                    else if (paint.IsPen) pen = paint;
                    else brush = paint;
                    break;
                case 0x041b: // RECTANGLE
                case 0x0418: // ELLIPSE
                    Size(14);
                    int bottom = I16(bytes, start), right = I16(bytes, start + 2), y = I16(bytes, start + 4), x = I16(bytes, start + 6);
                    if (right <= x || bottom <= y) throw new NotSupportedException("WMF requires nonempty shape bounds.");
                    double sx = viewportWidth / windowWidth, sy = viewportHeight / windowHeight;
                    // Word excludes one device pixel after truncating its logical-coordinate size.
                    double shapeWidth = right - x - Math.Truncate(1 / sx), shapeHeight = bottom - y - Math.Truncate(1 / sy);
                    if (shapeWidth <= 0 || shapeHeight <= 0 || (pen.Color != "none" && pen.Width == 0))
                        throw new NotSupportedException("WMF device hairlines and collapsed shape bounds are not supported.");
                    var shape = function == 0x041b
                        ? new XElement(ns + "rect", new XAttribute("x", x - originX), new XAttribute("y", y - originY),
                            new XAttribute("width", N(shapeWidth)), new XAttribute("height", N(shapeHeight)))
                        : new XElement(ns + "ellipse", new XAttribute("cx", N(x - originX + shapeWidth / 2)), new XAttribute("cy", N(y - originY + shapeHeight / 2)),
                            new XAttribute("rx", N(shapeWidth / 2)), new XAttribute("ry", N(shapeHeight / 2)));
                    // Both WMF mapping and later OOXML sizing stretch the complete painted graphic.
                    shape.Add(new XAttribute("fill", brush.Color), new XAttribute("stroke", pen.Color),
                        new XAttribute("stroke-width", pen.Width), new XAttribute("stroke-linejoin", "miter"),
                        new XAttribute("transform", $"scale({N(sx)} {N(sy)})"));
                    svg.Add(shape);
                    break;
                default:
                    throw new NotSupportedException($"Unsupported WMF record 0x{function:X4}.");
            }
            offset += length;
        }
        throw new InvalidDataException("WMF is missing EOF.");
    }

    private static string Color(uint value)
    {
        if ((value & 0xff000000) != 0) throw new NotSupportedException("WMF palette colors are not supported.");
        return $"#{value & 255:x2}{(value >> 8) & 255:x2}{(value >> 16) & 255:x2}";
    }
    private static string N(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    private static ushort U16(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));
    private static short I16(byte[] bytes, int offset) => BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(offset, 2));
    private static uint U32(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
}
