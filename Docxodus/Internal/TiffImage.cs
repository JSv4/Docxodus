using System.Buffers.Binary;
using System.IO.Compression;

namespace Docxodus.Internal;

/// <summary>Bounded, platform-independent conversion of single-page, chunky 8-bit RGB(A) TIFF.</summary>
internal sealed class TiffImage(byte[] bytes)
{
    private const int MaxPixels = 16 * 1024 * 1024;
    private readonly bool littleEndian = bytes.Length >= 2 && bytes[0] == 73 && bytes[1] == 73;
    private readonly Dictionary<int, uint[]> fields = new();

    internal byte[] ToPng()
    {
        if (bytes.Length < 8 || (!littleEndian && (bytes[0] != 77 || bytes[1] != 77)) || U16(2) != 42)
            throw new InvalidDataException("Invalid classic TIFF header.");
        int directory = Offset(U32(4), 2), entries = U16(directory);
        if (directory < 8 || entries > 4096) throw new InvalidDataException("TIFF directory exceeds the supported bound.");
        Offset((uint)directory, 2 + entries * 12 + 4);
        if (U32(directory + 2 + entries * 12) != 0)
            throw new NotSupportedException("Multi-page TIFF images are not supported.");
        for (int i = 0; i < entries; i++)
        {
            int entry = directory + 2 + i * 12, tag = U16(entry);
            if (tag is 301 or 318 or 319 or 320 or 322 or 323 or 324 or 325 or 342 or 532 or 34675)
                throw new NotSupportedException("TIFF tiles, palettes and custom color profiles are not supported.");
            if (tag is not (256 or 257 or 258 or 259 or 262 or 266 or 273 or 274 or 277 or 278 or 279 or 284 or 317 or 338 or 339)) continue;
            int type = U16(entry + 2), size = type == 3 ? 2 : type == 4 ? 4 : 0;
            uint count = U32(entry + 4);
            if (size == 0 || count is 0 or > 65536 || fields.ContainsKey(tag))
                throw new InvalidDataException("Invalid TIFF sample or strip field.");
            int length = checked((int)count * size), start = length <= 4 ? entry + 8 : Offset(U32(entry + 8), length);
            var values = new uint[count];
            for (int j = 0; j < values.Length; j++) values[j] = size == 2 ? U16(start + j * size) : U32(start + j * size);
            fields.Add(tag, values);
        }
        int width = checked((int)One(256)), height = checked((int)One(257));
        if (width <= 0 || height <= 0 || width > 32768 || height > 32768 || (long)width * height > MaxPixels)
            throw new InvalidDataException("TIFF decoded dimensions exceed the 16-megapixel limit.");
        int samples = checked((int)One(277, 1)), orientation = checked((int)One(274, 1));
        var bits = Values(258);
        if (One(262) != 2 || samples is not (3 or 4) || bits.Length != samples || bits.Any(b => b != 8)
            || One(284, 1) != 1 || One(266, 1) != 1 || One(317, 1) != 1
            || (fields.TryGetValue(339, out var formats) && (formats.Length != samples || formats.Any(f => f != 1))))
            throw new NotSupportedException("TIFF requires chunky unsigned 8-bit RGB or RGBA samples.");
        int alpha = samples == 4 ? checked((int)One(338)) : 0;
        if (orientation is < 1 or > 8 || (samples == 4 ? alpha is not (1 or 2) : fields.ContainsKey(338)))
            throw new NotSupportedException("Unsupported TIFF orientation or extra samples.");
        uint compression = One(259, 1), rows = One(278, uint.MaxValue);
        if (compression is not (1 or 32773)) throw new NotSupportedException("TIFF supports uncompressed and PackBits strips.");
        if (rows == 0) throw new InvalidDataException("TIFF strip row count must be positive.");
        int rowsPerStrip = (int)Math.Min(rows, (uint)height), strips = (height + rowsPerStrip - 1) / rowsPerStrip;
        var offsets = Values(273); var sizes = Values(279);
        if (offsets.Length != strips || sizes.Length != strips) throw new InvalidDataException("TIFF strip inventory is incomplete.");
        var pixels = new byte[checked(width * height * samples)];
        for (int strip = 0; strip < strips; strip++)
        {
            int length = checked(width * Math.Min(rowsPerStrip, height - strip * rowsPerStrip) * samples);
            var source = bytes.AsSpan(Offset(offsets[strip], checked((int)sizes[strip])), checked((int)sizes[strip]));
            var target = pixels.AsSpan(checked(strip * rowsPerStrip * width * samples), length);
            if (compression == 1)
            {
                if (source.Length != target.Length) throw new InvalidDataException("Invalid TIFF strip length.");
                source.CopyTo(target);
            }
            else Unpack(source, target);
        }
        int outputWidth = orientation < 5 ? width : height, outputHeight = orientation < 5 ? height : width;
        var rgba = new byte[checked(width * height * 4)];
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++)
        {
            var (ox, oy) = orientation switch {
                2 => (width - 1 - x, y), 3 => (width - 1 - x, height - 1 - y), 4 => (x, height - 1 - y),
                5 => (y, x), 6 => (height - 1 - y, x), 7 => (height - 1 - y, width - 1 - x),
                8 => (y, width - 1 - x), _ => (x, y) };
            int input = (y * width + x) * samples, output = (oy * outputWidth + ox) * 4;
            int a = samples == 4 ? pixels[input + 3] : 255;
            for (int channel = 0; channel < 3; channel++)
            {
                int value = pixels[input + channel];
                if (alpha == 1 && value > a) throw new InvalidDataException("TIFF associated color exceeds its alpha.");
                rgba[output + channel] = (byte)(alpha == 1 ? a == 0 ? 0 : (value * 255 + a / 2) / a : value);
            }
            rgba[output + 3] = (byte)a;
        }
        return Png(outputWidth, outputHeight, rgba);
    }

    private uint[] Values(int tag) => fields.TryGetValue(tag, out var values) ? values
        : throw new InvalidDataException($"TIFF is missing field {tag}.");
    private uint One(int tag, uint? fallback = null)
    {
        if (!fields.TryGetValue(tag, out var values) && fallback.HasValue) return fallback.Value;
        values ??= Values(tag);
        return values.Length == 1 ? values[0] : throw new InvalidDataException($"TIFF field {tag} must be scalar.");
    }
    private int Offset(uint offset, int length)
    {
        if (length < 0 || offset > bytes.Length || (ulong)offset + (uint)length > (uint)bytes.Length)
            throw new InvalidDataException("TIFF field points outside the image.");
        return (int)offset;
    }
    private ushort U16(int offset) => littleEndian ? BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2))
        : BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));
    private uint U32(int offset) => littleEndian ? BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4))
        : BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4));

    private static void Unpack(ReadOnlySpan<byte> source, Span<byte> target)
    {
        int input = 0, output = 0;
        while (input < source.Length)
        {
            int control = unchecked((sbyte)source[input++]);
            if (control == -128) continue;
            int count = control >= 0 ? control + 1 : 1 - control;
            if (count > target.Length - output || (control >= 0 ? count : 1) > source.Length - input)
                throw new InvalidDataException("TIFF PackBits run exceeds the strip bounds.");
            if (control >= 0) { source.Slice(input, count).CopyTo(target[output..]); input += count; }
            else target.Slice(output, count).Fill(source[input++]);
            output += count;
        }
        if (output != target.Length) throw new InvalidDataException("Incomplete TIFF PackBits strip.");
    }

    private static byte[] Png(int width, int height, byte[] rgba)
    {
        using var output = new MemoryStream();
        output.Write(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
        void Chunk(ReadOnlySpan<byte> type, ReadOnlySpan<byte> data)
        {
            Span<byte> value = stackalloc byte[4];
            BinaryPrimitives.WriteInt32BigEndian(value, data.Length); output.Write(value); output.Write(type); output.Write(data);
            BinaryPrimitives.WriteUInt32BigEndian(value, ~Crc32.Append(Crc32.Append(uint.MaxValue, type), data)); output.Write(value);
        }
        var header = new byte[13]; BinaryPrimitives.WriteInt32BigEndian(header, width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height); header[8] = 8; header[9] = 6;
        Chunk("IHDR"u8, header);
        using var encoded = new MemoryStream();
        using (var compressor = new ZLibStream(encoded, CompressionLevel.Fastest, leaveOpen: true))
            for (int row = 0; row < height; row++) { compressor.WriteByte(0); compressor.Write(rgba.AsSpan(row * width * 4, width * 4)); }
        Chunk("IDAT"u8, encoded.ToArray()); Chunk("IEND"u8, ReadOnlySpan<byte>.Empty);
        return output.ToArray();
    }
}
