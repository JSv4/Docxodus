#nullable enable

using System;
using System.Buffers;
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

    /// <summary>The largest buffer <see cref="ArrayPool{T}.Shared"/> actually pools (1 MiB). A rent
    /// above it allocates a fresh array and a return discards it, so the pooled-buffer trade stops
    /// paying and an exact size becomes the cheaper one.</summary>
    private const int PooledCeiling = 1024 * 1024;

    /// <summary>Compute the SHA-256 digest of the UTF-8 encoding of <paramref name="text"/>.</summary>
    public static IrHash Compute(string text)
    {
        return ComputeUtf8(text);
    }

    /// <summary>
    /// The digest of <paramref name="text"/>'s UTF-8 encoding, computed without materializing that
    /// encoding: short strings encode into a stack buffer, longer ones into a pooled array. Identical
    /// bytes in, identical digest out — this is purely the allocation-free spelling of
    /// <see cref="Compute(string)"/>, for the hot canonical-XML hashing paths.
    /// </summary>
    public static IrHash ComputeUtf8(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        Span<byte> inline = stackalloc byte[1024];

        // GetMaxByteCount is 3n+3, and ArrayPool.Shared stops pooling above 1 MB: renting the
        // worst case for a large canonical subtree (a block content control, an opaque table) would
        // allocate three times the string's actual UTF-8 length and then drop it, which is WORSE than
        // the single exact-sized array this method exists to avoid. Past that threshold, pay one
        // counting pass and rent the exact size instead.
        int maxBytes = Encoding.UTF8.GetMaxByteCount(text.Length);
        if (maxBytes > PooledCeiling)
            maxBytes = Encoding.UTF8.GetByteCount(text);

        byte[]? rented = maxBytes <= inline.Length ? null : ArrayPool<byte>.Shared.Rent(maxBytes);
        Span<byte> buffer = rented is null ? inline : rented;

        try
        {
            int written = Encoding.UTF8.GetBytes(text, buffer);
            return Compute(buffer[..written]);
        }
        finally
        {
            if (rented is not null)
                ArrayPool<byte>.Shared.Return(rented);
        }
    }

    /// <summary>
    /// Write the 32 raw digest bytes into <paramref name="destination"/> in the same
    /// big-endian order as <see cref="ToHex"/>. The span must be at least 32 bytes long.
    /// </summary>
    public void CopyTo(Span<byte> destination)
    {
        BinaryPrimitives.WriteUInt64BigEndian(destination.Slice(0, 8), _a);
        BinaryPrimitives.WriteUInt64BigEndian(destination.Slice(8, 8), _b);
        BinaryPrimitives.WriteUInt64BigEndian(destination.Slice(16, 8), _c);
        BinaryPrimitives.WriteUInt64BigEndian(destination.Slice(24, 8), _d);
    }

    /// <summary>The 32 raw digest bytes, big-endian (matching <see cref="ToHex"/>).</summary>
    public byte[] ToBytes()
    {
        var bytes = new byte[32];
        CopyTo(bytes);
        return bytes;
    }

    /// <summary>Render the digest as 64 lowercase hex characters.</summary>
    public string ToHex()
    {
        Span<byte> digest = stackalloc byte[32];
        CopyTo(digest);
        return Convert.ToHexString(digest).ToLowerInvariant();
    }

    public bool Equals(IrHash other) => _a == other._a && _b == other._b && _c == other._c && _d == other._d;

    public override bool Equals(object? obj) => obj is IrHash other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(_a, _b, _c, _d);

    public static bool operator ==(IrHash left, IrHash right) => left.Equals(right);

    public static bool operator !=(IrHash left, IrHash right) => !left.Equals(right);

    public override string ToString() => ToHex();
}
