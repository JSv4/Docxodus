using System.Buffers.Binary;

namespace Docxodus.Internal;

/// <summary>Bounded decoding of the device-independent bitmaps that EMF and WMF records embed, to RGBA
/// pixels. Unsupported encodings reject the whole picture: RLE, and the JPEG or PNG a DIB can wrap for a
/// printer (BI_JPEG, BI_PNG), whose real size its header does not bound.</summary>
internal static class DibImage
{
    private const int MaxSide = 32768, MaxPixels = 16 * 1024 * 1024;

    /// <summary>A decoded source rectangle: top-down, unpremultiplied RGBA.</summary>
    internal sealed record Bitmap(int Width, int Height, byte[] Rgba)
    {
        private string? href;

        /// <summary>The PNG data URI, encoded once: a pattern brush's tile is written for every fill.</summary>
        internal string Href => href ??= "data:image/png;base64," + Convert.ToBase64String(TiffImage.Png(Width, Height, Rgba));
    }

    /// <summary>Bytes taken by the header, colour masks and colour table of a packed DIB.</summary>
    internal static int InfoSize(ReadOnlySpan<byte> packed)
    {
        var info = Header(packed);
        return info.HeaderSize + info.MaskBytes + info.PaletteEntries * info.EntrySize;
    }

    /// <summary>Decodes the source rectangle (top-down pixel coordinates) of a DIB. Returns null when the
    /// rectangle misses the bitmap. <paramref name="constantAlpha"/> and <paramref name="sourceAlpha"/>
    /// (premultiplied 32-bit alpha) are ALPHABLEND's blend function, baked into the PNG.</summary>
    internal static Bitmap? Decode(ReadOnlySpan<byte> info, ReadOnlySpan<byte> bits, int srcX, int srcY, int srcWidth, int srcHeight,
        bool srcYFromBottom, byte constantAlpha = 255, bool sourceAlpha = false)
    {
        var h = Header(info);
        if (info.Length < h.HeaderSize + h.MaskBytes + h.PaletteEntries * h.EntrySize)
            throw new InvalidDataException("Truncated metafile bitmap colour table.");
        long stride = ((long)h.Width * h.BitCount + 31) / 32 * 4;
        if (stride * h.Height > bits.Length) throw new InvalidDataException("Truncated metafile bitmap.");
        // StretchDIBits measures a bottom-up bitmap's source rectangle from its lower-left corner.
        if (srcYFromBottom && !h.TopDown) srcY = h.Height - srcY - srcHeight;
        if (srcWidth < 0) { srcX += srcWidth; srcWidth = -srcWidth; }
        if (srcHeight < 0) { srcY += srcHeight; srcHeight = -srcHeight; }
        int left = Math.Max(0, srcX), top = Math.Max(0, srcY);
        int right = (int)Math.Min(h.Width, (long)srcX + srcWidth), bottom = (int)Math.Min(h.Height, (long)srcY + srcHeight);
        if (right <= left || bottom <= top) return null;
        if (left != srcX || top != srcY || right - left != srcWidth || bottom - top != srcHeight)
            throw new NotSupportedException("Metafile bitmap source rectangle extends past the bitmap.");
        int width = right - left, height = bottom - top;

        var palette = new uint[h.PaletteEntries];
        int paletteStart = h.HeaderSize + h.MaskBytes;
        for (int i = 0; i < palette.Length; i++)
        {
            var entry = info.Slice(paletteStart + i * h.EntrySize, 3);
            palette[i] = entry[0] | (uint)entry[1] << 8 | (uint)entry[2] << 16;
        }
        var masks = h.Masks;
        var rgba = new byte[(long)width * height * 4];
        for (int row = 0; row < height; row++)
        {
            int sourceRow = top + row, stored = h.TopDown ? sourceRow : h.Height - 1 - sourceRow;
            var line = bits.Slice((int)(stored * stride), (int)stride);
            for (int column = 0; column < width; column++)
            {
                int x = left + column, at = (row * width + column) * 4;
                uint bgr, alpha = 255;
                switch (h.BitCount)
                {
                    case 1: case 4: case 8:
                        int index = h.BitCount == 8 ? line[x] : h.BitCount == 4 ? (line[x >> 1] >> ((x & 1) == 0 ? 4 : 0)) & 15
                            : (line[x >> 3] >> (7 - (x & 7))) & 1;
                        bgr = index < palette.Length ? palette[index] : 0;
                        break;
                    case 24:
                        bgr = line[x * 3] | (uint)line[x * 3 + 1] << 8 | (uint)line[x * 3 + 2] << 16;
                        break;
                    default:
                        uint value = h.BitCount == 16 ? BinaryPrimitives.ReadUInt16LittleEndian(line.Slice(x * 2, 2))
                            : BinaryPrimitives.ReadUInt32LittleEndian(line.Slice(x * 4, 4));
                        bgr = Channel(value, masks.Blue) | Channel(value, masks.Green) << 8 | Channel(value, masks.Red) << 16;
                        if (sourceAlpha && h.BitCount == 32) alpha = value >> 24;
                        break;
                }
                uint r = (bgr >> 16) & 255, g = (bgr >> 8) & 255, b = bgr & 255;
                if (alpha != 255)
                {
                    r = alpha == 0 ? 0 : Math.Min(255, r * 255 / alpha);
                    g = alpha == 0 ? 0 : Math.Min(255, g * 255 / alpha);
                    b = alpha == 0 ? 0 : Math.Min(255, b * 255 / alpha);
                }
                rgba[at] = (byte)r; rgba[at + 1] = (byte)g; rgba[at + 2] = (byte)b; rgba[at + 3] = (byte)(alpha * constantAlpha / 255);
            }
        }
        return new Bitmap(width, height, rgba);
    }

    /// <summary>The pixel size a DIB header declares.</summary>
    internal static (int Width, int Height) Size(ReadOnlySpan<byte> info)
    {
        var h = Header(info);
        return (h.Width, h.Height);
    }

    /// <summary>Decodes a whole DIB, as a pattern brush carries it.</summary>
    internal static Bitmap Decode(ReadOnlySpan<byte> info, ReadOnlySpan<byte> bits)
    {
        var (width, height) = Size(info);
        return Decode(info, bits, 0, 0, width, height, false)!;
    }

    private static uint Channel(uint value, uint mask)
    {
        if (mask == 0) return 0;
        int shift = System.Numerics.BitOperations.TrailingZeroCount(mask), bits = System.Numerics.BitOperations.PopCount(mask);
        uint raw = (value & mask) >> shift, max = bits >= 32 ? uint.MaxValue : (1u << bits) - 1;
        return (uint)((ulong)raw * 255 / max);
    }

    private readonly record struct Info(int HeaderSize, int Width, int Height, bool TopDown, int BitCount, uint Compression,
        int ImageSize, int PaletteEntries, int EntrySize, int MaskBytes, (uint Red, uint Green, uint Blue) Masks);

    private static Info Header(ReadOnlySpan<byte> info)
    {
        if (info.Length < 12) throw new InvalidDataException("Truncated metafile bitmap header.");
        uint size = BinaryPrimitives.ReadUInt32LittleEndian(info);
        int width, height, bitCount, palette = 0, imageSize = 0, maskBytes = 0;
        uint compression = 0;
        bool topDown = false;
        (uint, uint, uint) masks = (0x7c00, 0x3e0, 0x1f);
        if (size == 12)
        {
            width = BinaryPrimitives.ReadUInt16LittleEndian(info[4..]);
            height = BinaryPrimitives.ReadUInt16LittleEndian(info[6..]);
            bitCount = BinaryPrimitives.ReadUInt16LittleEndian(info[10..]);
        }
        else if (size is 40 or 52 or 56 or 108 or 124 && info.Length >= size)
        {
            width = BinaryPrimitives.ReadInt32LittleEndian(info[4..]);
            int rawHeight = BinaryPrimitives.ReadInt32LittleEndian(info[8..]);
            if (rawHeight == int.MinValue) throw new InvalidDataException("Invalid metafile bitmap height.");
            topDown = rawHeight < 0; height = Math.Abs(rawHeight);
            bitCount = BinaryPrimitives.ReadUInt16LittleEndian(info[14..]);
            compression = BinaryPrimitives.ReadUInt32LittleEndian(info[16..]);
            imageSize = BinaryPrimitives.ReadInt32LittleEndian(info[20..]);
            uint used = BinaryPrimitives.ReadUInt32LittleEndian(info[32..]);
            if (used > 256) throw new InvalidDataException("Invalid metafile bitmap colour count.");
            palette = (int)used;
            if (compression == 3)
            {
                if (bitCount is not (16 or 32)) throw new NotSupportedException("Unsupported metafile bitfield bitmap.");
                if (size == 40) maskBytes = 12;
                if (info.Length < 40 + 12) throw new InvalidDataException("Truncated metafile bitmap masks.");
                masks = (BinaryPrimitives.ReadUInt32LittleEndian(info[40..]), BinaryPrimitives.ReadUInt32LittleEndian(info[44..]),
                    BinaryPrimitives.ReadUInt32LittleEndian(info[48..]));
            }
            else if (bitCount == 32) masks = (0xff0000, 0xff00, 0xff);
        }
        else throw new NotSupportedException("Unsupported metafile bitmap header.");
        if (compression is not (0 or 3)) throw new NotSupportedException("Compressed metafile bitmaps are not supported.");
        if (bitCount is not (1 or 4 or 8 or 16 or 24 or 32)) throw new NotSupportedException("Unsupported metafile bitmap depth.");
        if (bitCount is > 0 and <= 8 && palette == 0) palette = 1 << bitCount;
        if (bitCount > 8 && size == 12) palette = 0;
        if (width <= 0 || height <= 0 || width > MaxSide || height > MaxSide || (long)width * height > MaxPixels || imageSize < 0)
            throw new InvalidDataException("Metafile bitmap dimensions exceed the supported bound.");
        return new Info((int)size, width, height, topDown, bitCount, compression, imageSize, bitCount <= 8 || size != 12 ? palette : 0,
            size == 12 ? 3 : 4, maskBytes, masks);
    }
}
