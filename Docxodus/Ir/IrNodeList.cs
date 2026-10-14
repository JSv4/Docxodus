#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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
/// <see cref="IrNodeList.Empty{T}"/>; the wrapper takes ownership of a private array and never
/// mutates it, so callers must hand in an array they will not retain or mutate.
/// </remarks>
internal sealed class IrNodeList<T> : IReadOnlyList<T>, IEquatable<IrNodeList<T>>
{
    private readonly T[] _items;

    // Lazily-computed hash cache. The list is immutable so the sequence hash never changes;
    // null means "not yet computed". A 0-sentinel would be wrong because 0 is a legitimate hash.
    // Thread-safety: a benign race is acceptable — concurrent callers may each compute the hash,
    // but the computation is idempotent so they all store the same value.
    private int? _hashCode;

    // Construction is funnelled through the IrNodeList.From/Empty factories so the no-copy
    // immutability guarantee is structural: nothing outside this assembly can hand us an array
    // it still holds a reference to.
    private IrNodeList(T[] items) => _items = items;

    /// <summary>
    /// Wrap an array directly without copying. The caller transfers ownership and must not
    /// retain or mutate <paramref name="items"/> afterwards.
    /// </summary>
    internal static IrNodeList<T> WrapNoCopy(T[] items) => new(items);

    public T this[int index] => _items[index];

    public int Count => _items.Length;

    /// <summary>Allocation-free struct enumerator so <c>foreach</c> over an <see cref="IrNodeList{T}"/> boxes nothing.</summary>
    public Enumerator GetEnumerator() => new(_items);

    IEnumerator<T> IEnumerable<T>.GetEnumerator() => ((IEnumerable<T>)_items).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => _items.GetEnumerator();

    /// <summary>A value-type enumerator over the backing array; avoids the boxing allocation of the interface enumerator.</summary>
    public struct Enumerator : IEnumerator<T>
    {
        private readonly T[] _items;
        private int _index;

        internal Enumerator(T[] items)
        {
            _items = items;
            _index = -1;
        }

        public T Current => _items[_index];

        object? IEnumerator.Current => Current;

        public bool MoveNext() => ++_index < _items.Length;

        public void Reset() => _index = -1;

        public void Dispose() { }
    }

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
        // See _hashCode remarks: cache lazily; the benign race is acceptable.
        if (_hashCode is int cached)
            return cached;

        var hash = new HashCode();
        foreach (var item in _items)
            hash.Add(item);
        int computed = hash.ToHashCode();
        _hashCode = computed;
        return computed;
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
        return IrNodeList<T>.WrapNoCopy(items.ToArray());
    }

    /// <summary>The empty <see cref="IrNodeList{T}"/>.</summary>
    public static IrNodeList<T> Empty<T>() => IrNodeList<T>.WrapNoCopy(Array.Empty<T>());
}
