#nullable enable

using System;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// Back-reference from an IR node to the OOXML it was read from (the source
/// <see cref="XElement"/> and the originating part <see cref="PartUri"/>).
/// </summary>
/// <remarks>
/// IR snapshots must be value-equal node-for-node with provenance <em>excluded</em>: two IR
/// trees built from different physical documents that have identical structure/content should
/// compare equal even though their provenance differs. C# records, however, include every
/// property in their generated equality — so an <c>IrProvenance Source</c> property on a record
/// would leak the source element/part into the comparison.
/// <para/>
/// The trick: this type's <see cref="Equals(object?)"/> returns <c>true</c> for <em>any</em>
/// other <see cref="IrProvenance"/> instance and <see cref="GetHashCode"/> always returns
/// <c>0</c>. Records that embed an <c>IrProvenance</c> therefore compare equal regardless of
/// the provenance they carry, while still exposing the source for diagnostics/round-tripping.
/// </remarks>
internal sealed class IrProvenance
{
    /// <summary>The source OOXML element this IR node was read from, if known.</summary>
    public XElement? Element { get; init; }

    /// <summary>The URI of the part the source element lived in, if known.</summary>
    public Uri? PartUri { get; init; }

    /// <summary>Always equal to any other <see cref="IrProvenance"/> so provenance is excluded from record equality.</summary>
    public override bool Equals(object? obj) => obj is IrProvenance;

    /// <summary>Always <c>0</c> so provenance contributes nothing to a containing record's hash.</summary>
    public override int GetHashCode() => 0;
}
