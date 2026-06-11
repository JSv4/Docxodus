#nullable enable

using System;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;

namespace Docxodus.Ir;

/// <summary>
/// A 32-byte SHA-256 digest stored inline as four <see cref="ulong"/> fields (no heap
/// allocation). Used by the Document IR to give every node a stable, value-equal content
/// hash. Equality is full structural equality over the digest bytes.
/// </summary>
internal readonly struct IrHash : IEquatable<IrHash>
{
    private readonly ulong _a;
    private readonly ulong _b;
    private readonly ulong _c;
    private readonly ulong _d;

    private IrHash(ulong a, ulong b, ulong c, ulong d)
    {
        _a = a;
        _b = b;
        _c = c;
        _d = d;
    }

    /// <summary>Compute the SHA-256 digest of <paramref name="data"/>.</summary>
    public static IrHash Compute(ReadOnlySpan<byte> data)
    {
        Span<byte> digest = stackalloc byte[32];
        SHA256.HashData(data, digest);
        return new IrHash(
            BinaryPrimitives.ReadUInt64BigEndian(digest.Slice(0, 8)),
            BinaryPrimitives.ReadUInt64BigEndian(digest.Slice(8, 8)),
            BinaryPrimitives.ReadUInt64BigEndian(digest.Slice(16, 8)),
            BinaryPrimitives.ReadUInt64BigEndian(digest.Slice(24, 8)));
    }

    /// <summary>Compute the SHA-256 digest of the UTF-8 encoding of <paramref name="text"/>.</summary>
    public static IrHash Compute(string text)
    {
        return Compute(Encoding.UTF8.GetBytes(text));
    }

    /// <summary>Render the digest as 64 lowercase hex characters.</summary>
    public string ToHex()
    {
        Span<byte> digest = stackalloc byte[32];
        BinaryPrimitives.WriteUInt64BigEndian(digest.Slice(0, 8), _a);
        BinaryPrimitives.WriteUInt64BigEndian(digest.Slice(8, 8), _b);
        BinaryPrimitives.WriteUInt64BigEndian(digest.Slice(16, 8), _c);
        BinaryPrimitives.WriteUInt64BigEndian(digest.Slice(24, 8), _d);
        return Convert.ToHexString(digest).ToLowerInvariant();
    }

    public bool Equals(IrHash other) => _a == other._a && _b == other._b && _c == other._c && _d == other._d;

    public override bool Equals(object? obj) => obj is IrHash other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(_a, _b, _c, _d);

    public static bool operator ==(IrHash left, IrHash right) => left.Equals(right);

    public static bool operator !=(IrHash left, IrHash right) => !left.Equals(right);

    public override string ToString() => ToHex();
}
