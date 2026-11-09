using System.Buffers.Binary;
using System.Globalization;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>Bounded, static MM_TEXT EMF playback. Unsupported records reject the whole image.</summary>
internal sealed record EmfImage(XElement Drawing)
{
    private sealed record Paint(bool IsPen, string Color, int Width = 1, bool Cosmetic = true);

    internal static EmfImage Read(byte[] bytes)
    {
        if (bytes.Length < 108 || U32(bytes, 0) != 1 || U32(bytes, 4) != 88
            || U32(bytes, 40) != 0x464d4520 || U32(bytes, 44) != 0x10000)
            throw new NotSupportedException("EMF requires a version-one base header.");
        int handles = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(56, 2));
        uint records = U32(bytes, 52);
        if (U32(bytes, 48) != bytes.Length || records < 2 || records > 10000 || handles is < 1 or > 1024
            || I32(bytes, 16) < I32(bytes, 8) || I32(bytes, 20) < I32(bytes, 12))
            throw new InvalidDataException("Invalid or oversized EMF header.");
        if (U32(bytes, 60) != 0 || U32(bytes, 64) != 0 || U32(bytes, 68) != 0)
            throw new NotSupportedException("EMF descriptions and palettes are not supported.");
        double frameX = I32(bytes, 24), frameY = I32(bytes, 28);
        double frameWidth = I32(bytes, 32) - frameX, frameHeight = I32(bytes, 36) - frameY;
        if (frameWidth <= 0 || frameHeight <= 0 || I32(bytes, 72) <= 0 || I32(bytes, 76) <= 0
            || I32(bytes, 80) <= 0 || I32(bytes, 84) <= 0)
            throw new InvalidDataException("EMF requires positive frame and reference device dimensions.");
        // MM_TEXT coordinates are reference-device pixels; Frame is in hundredths of a millimeter.
        double sx = 100.0 * I32(bytes, 80) / I32(bytes, 72), sy = 100.0 * I32(bytes, 84) / I32(bytes, 76);
        double minX = frameX, minY = frameY, maxX = frameX + frameWidth, maxY = frameY + frameHeight;
        var objects = new Paint?[handles];
        Paint defaultPen = new(true, "#000000"), defaultBrush = new(false, "#ffffff");
        Paint pen = defaultPen, brush = defaultBrush;
        XNamespace ns = "http://www.w3.org/2000/svg";
        var drawing = new XElement(ns + "g", new XAttribute("transform", $"scale({N(sx)} {N(sy)})"));
        int count = 1;
        for (int offset = 88; offset < bytes.Length;)
        {
            if (++count > records || bytes.Length - offset < 8)
                throw new InvalidDataException("EMF record count or truncated record.");
            uint kind = U32(bytes, offset), size = U32(bytes, offset + 4);
            if (size < 8 || size % 4 != 0 || size > bytes.Length - offset)
                throw new InvalidDataException("Invalid EMF record length.");
            int length = (int)size, start = offset + 8;
            void Size(int expected) { if (length != expected) throw new InvalidDataException("Invalid EMF record parameters."); }
            switch (kind)
            {
                case 14: // EOF
                    Size(20);
                    if (offset + length != bytes.Length || count != records || U32(bytes, start) != 0
                        || U32(bytes, start + 8) != size)
                        throw new InvalidDataException("Invalid EMF end record.");
                    return Finish();
                case 17: // SETMAPMODE
                    Size(12);
                    if (I32(bytes, start) != 1) throw new NotSupportedException("EMF supports MM_TEXT mapping only.");
                    break;
                case 38: // CREATEPEN
                case 39: // CREATEBRUSHINDIRECT
                    bool isPen = kind == 38;
                    Size(isPen ? 28 : 24);
                    uint slot = U32(bytes, start), style = U32(bytes, start + 4);
                    if (slot == 0 || slot >= handles || objects[slot] != null)
                        throw new InvalidDataException("Invalid EMF object slot.");
                    bool cosmetic = (style & 0x10000) == 0;
                    int width = isPen ? I32(bytes, start + 8) : 0;
                    if (isPen ? style is not (0 or 5 or 0x10000 or 0x10005) || (cosmetic ? width != 1 : width <= 0)
                        : style is not (0 or 1))
                        throw new NotSupportedException("EMF supports solid/null brushes and solid/null cosmetic or geometric pens.");
                    string color = Color(U32(bytes, start + (isPen ? 16 : 8)));
                    objects[slot] = new Paint(isPen, (style & 0xffff) == 0 ? color : "none", width, cosmetic);
                    break;
                case 37: // SELECTOBJECT
                case 40: // DELETEOBJECT
                    Size(12);
                    uint index = U32(bytes, start);
                    Paint paint;
                    if (kind == 37 && (index & 0x80000000) != 0)
                        paint = index switch {
                            0x80000000 => defaultBrush, 0x80000004 => new(false, "#000000"),
                            0x80000005 => new(false, "none"), 0x80000006 => new(true, "#ffffff"),
                            0x80000007 => defaultPen, 0x80000008 => new(true, "none"),
                            _ => throw new NotSupportedException("Unsupported EMF stock object.") };
                    else
                    {
                        if (index == 0 || index >= handles || objects[index] is not { } selected)
                            throw new InvalidDataException("Invalid EMF object selection.");
                        paint = selected;
                    }
                    if (kind == 40)
                    {
                        objects[index] = null;
                        if (ReferenceEquals(pen, paint)) pen = defaultPen;
                        if (ReferenceEquals(brush, paint)) brush = defaultBrush;
                    }
                    else if (paint.IsPen) pen = paint;
                    else brush = paint;
                    break;
                case 42: // ELLIPSE
                case 43: // RECTANGLE
                    Size(24);
                    double x = I32(bytes, start), y = I32(bytes, start + 4);
                    double widthShape = I32(bytes, start + 8) - x, heightShape = I32(bytes, start + 12) - y;
                    if (widthShape <= 0 || heightShape <= 0 || pen.Color == "none")
                        throw new NotSupportedException("EMF requires positive shape bounds and a visible outline.");
                    var shape = kind == 43
                        ? new XElement(ns + "rect", new XAttribute("x", N(x)), new XAttribute("y", N(y)),
                            new XAttribute("width", N(widthShape)), new XAttribute("height", N(heightShape)))
                        : new XElement(ns + "ellipse", new XAttribute("cx", N(x + widthShape / 2)), new XAttribute("cy", N(y + heightShape / 2)),
                            new XAttribute("rx", N(widthShape / 2)), new XAttribute("ry", N(heightShape / 2)));
                    shape.Add(new XAttribute("fill", brush.Color), new XAttribute("stroke", pen.Color),
                        new XAttribute("stroke-width", pen.Width), new XAttribute("stroke-linejoin", "round"));
                    if (pen.Cosmetic) shape.Add(new XAttribute("vector-effect", "non-scaling-stroke"));
                    drawing.Add(shape);
                    double pad = pen.Cosmetic ? 0 : pen.Width / 2.0;
                    minX = Math.Min(minX, (x - pad) * sx); minY = Math.Min(minY, (y - pad) * sy);
                    maxX = Math.Max(maxX, (x + widthShape + pad) * sx); maxY = Math.Max(maxY, (y + heightShape + pad) * sy);
                    break;
                default:
                    throw new NotSupportedException($"Unsupported EMF record {kind}.");
            }
            offset += length;
        }
        throw new InvalidDataException("EMF is missing EOF.");

        EmfImage Finish()
        {
            double width = maxX - minX, height = maxY - minY, pxWidth = width * 96 / 2540 + 1, pxHeight = height * 96 / 2540 + 1;
            if (pxWidth > 32768 || pxHeight > 32768 || pxWidth * pxHeight > 16 * 1024 * 1024)
                throw new InvalidDataException("EMF intrinsic size exceeds the 16-megapixel limit.");
            // Frame establishes mapping, not a clip. Inline SVG keeps overflow visible and
            // cosmetic strokes one CSS pixel wide, including when the picture is resized.
            return new EmfImage(new XElement(ns + "svg", new XAttribute("width", N(frameWidth * 96 / 2540)),
                new XAttribute("height", N(frameHeight * 96 / 2540)), new XAttribute("data-docx-vector-frame", "emf"),
                new XAttribute("viewBox", $"{N(frameX)} {N(frameY)} {N(frameWidth)} {N(frameHeight)}"),
                new XAttribute("preserveAspectRatio", "none"), new XAttribute("overflow", "visible"), drawing));
        }
    }

    private static string Color(uint value)
    {
        if ((value & 0xff000000) != 0) throw new NotSupportedException("EMF palette colors are not supported.");
        return $"#{value & 255:x2}{(value >> 8) & 255:x2}{(value >> 16) & 255:x2}";
    }
    private static string N(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    private static uint U32(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
    private static int I32(byte[] bytes, int offset) => BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset, 4));
}
