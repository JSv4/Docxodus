using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// Source contributions of the leaves committed to one paragraph's IR. This retained-XML annotation
/// participates in neither value equality nor serialization. Repeated sources represent multiple leaves
/// in an indivisible carrier; omitted source elements contribute no text or atomic position.
/// </summary>
internal sealed record IrSourceContent(IReadOnlyList<IrSourceContent.Entry> Entries)
{
    // Atom ordinals are independent of word tokenization and contain no retained source references.
    internal readonly record struct CarrierSpan(int Start, int End);

    internal IrNodeList<CarrierSpan> SharedCarriers()
    {
        var carriers = new Dictionary<XElement, (int Start, int End)>();
        int atom = 0;
        foreach (var entry in Entries)
        {
            if (!entry.Atomic) continue;
            carriers[entry.Source] = carriers.TryGetValue(entry.Source, out var span)
                ? (span.Start, atom + 1) : (atom, atom + 1);
            atom++;
        }
        return IrNodeList.From(carriers.Values.Where(span => span.End - span.Start > 1)
            .OrderBy(span => span.Start).Select(span => new CarrierSpan(span.Start, span.End)).ToList());
    }

    internal readonly record struct Entry(XElement Source, int TextLength, bool Atomic);

    internal static IEnumerable<IrInline> Leaves(IEnumerable<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            if (inline is IrHyperlink link)
            {
                foreach (var leaf in Leaves(link.Inlines)) yield return leaf;
            }
            else if (inline is IrFieldRun field)
            {
                foreach (var leaf in Leaves(field.CachedResult)) yield return leaf;
            }
            else yield return inline;
        }
    }

    /// <summary>Temporary provenance across reader normalization; coalescing transfers the original sources.</summary>
    internal sealed class Builder
    {
        private readonly Dictionary<IrInline, Contributions> _sources = new(ReferenceEqualityComparer.Instance);
        private readonly record struct Contributions(Entry First, List<Entry>? Rest = null);

        public void Register(IrInline inline, XElement source)
        {
            if (inline is IrHyperlink or IrFieldRun) return;
            _sources[inline] = new(new(source, inline is IrTextRun text ? text.Text.Length : 0, inline is not IrTextRun));
        }

        public void Transfer(IrInline original, IrInline replacement)
        {
            if (_sources.Remove(original, out var source)) _sources[replacement] = source;
        }

        public void Merge(IrInline left, IrInline right, IrInline merged)
        {
            var a = _sources[left];
            var b = _sources[right];
            var rest = a.Rest ?? new List<Entry>();
            rest.Add(b.First);
            if (b.Rest is not null) rest.AddRange(b.Rest);
            _sources.Remove(left);
            _sources.Remove(right);
            _sources[merged] = new(a.First, rest);
        }

        public IrSourceContent Build(IEnumerable<IrInline> inlines, bool includeText)
        {
            var result = new List<Entry>();
            foreach (var inline in Leaves(inlines))
            {
                if (!includeText && inline is IrTextRun) continue;
                var source = _sources[inline];
                result.Add(source.First);
                if (source.Rest is not null) result.AddRange(source.Rest);
            }
            return new(result.ToArray());
        }
    }
}
