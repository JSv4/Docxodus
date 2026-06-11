#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;

namespace Docxodus.Ir;

/// <summary>
/// An immutable, value-equal list of IR child nodes.
/// </summary>
/// <remarks>
/// C# records compute equality member-by-member, but a record property typed as a bare
/// <see cref="IReadOnlyList{T}"/> (e.g. an <c>T[]</c>) is compared by <em>reference</em>, which
/// would break the IR's "two reads of the same bytes produce node-for-node value-equal trees"
/// guarantee (§8): two structurally identical paragraphs whose <c>Inlines</c> arrays were built
/// separately would compare unequal. To fix this, every IR node holds its children as an
/// <see cref="IrNodeList{T}"/> — whose <see cref="Equals(IrNodeList{T})"/>/<see cref="GetHashCode"/>
/// are sequence-based — so record equality composes correctly down the tree.
/// <para/>
/// Construct via <see cref="IrNodeList.From{T}(IEnumerable{T})"/> or
/// <see cref="IrNodeList.Empty{T}"/>; the wrapper copies into a private array and never mutates.
/// </remarks>
internal sealed class IrNodeList<T> : IReadOnlyList<T>, IEquatable<IrNodeList<T>>
{
    private readonly T[] _items;

    internal IrNodeList(T[] items) => _items = items;

    public T this[int index] => _items[index];

    public int Count => _items.Length;

    public IEnumerator<T> GetEnumerator() => ((IEnumerable<T>)_items).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => _items.GetEnumerator();

    public bool Equals(IrNodeList<T>? other)
    {
        if (other is null)
            return false;
        if (ReferenceEquals(this, other))
            return true;
        if (_items.Length != other._items.Length)
            return false;

        var comparer = EqualityComparer<T>.Default;
        for (int i = 0; i < _items.Length; i++)
        {
            if (!comparer.Equals(_items[i], other._items[i]))
                return false;
        }

        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as IrNodeList<T>);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        foreach (var item in _items)
            hash.Add(item);
        return hash.ToHashCode();
    }

    public static bool operator ==(IrNodeList<T>? left, IrNodeList<T>? right) =>
        left is null ? right is null : left.Equals(right);

    public static bool operator !=(IrNodeList<T>? left, IrNodeList<T>? right) => !(left == right);
}

/// <summary>Factory helpers for <see cref="IrNodeList{T}"/>.</summary>
internal static class IrNodeList
{
    /// <summary>Wrap a sequence into a value-equal <see cref="IrNodeList{T}"/> (copies the elements).</summary>
    public static IrNodeList<T> From<T>(IEnumerable<T> items)
    {
        if (items is null)
            throw new ArgumentNullException(nameof(items));
        return new IrNodeList<T>(System.Linq.Enumerable.ToArray(items));
    }

    /// <summary>The empty <see cref="IrNodeList{T}"/>.</summary>
    public static IrNodeList<T> Empty<T>() => new(Array.Empty<T>());
}
