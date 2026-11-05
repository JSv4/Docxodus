namespace Docxodus.Internal;

internal static class Crc32
{
    private static readonly uint[] Table = CreateTable();

    internal static uint Compute(ReadOnlySpan<byte> bytes) => ~Append(uint.MaxValue, bytes);

    internal static uint Append(uint state, ReadOnlySpan<byte> bytes)
    {
        foreach (var value in bytes) state = Table[(state ^ value) & 0xff] ^ (state >> 8);
        return state;
    }

    private static uint[] CreateTable()
    {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++)
        {
            var value = index;
            for (var bit = 0; bit < 8; bit++)
                value = (value & 1) != 0 ? 0xedb88320U ^ (value >> 1) : value >> 1;
            table[index] = value;
        }
        return table;
    }
}
