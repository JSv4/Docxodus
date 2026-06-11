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
    /// <summary>
    /// A single shared instance carrying no element/part — used for every node when the reader runs
    /// with <see cref="IrReaderOptions.RetainSources"/> = <c>false</c>. Sharing one instance means
    /// retention-off reads pay zero per-node provenance allocation, and because provenance is
    /// equality-neutral (see the type remarks) reusing it never perturbs any node's value/hash.
    /// </summary>
    public static readonly IrProvenance Empty = new();

    /// <summary>The source OOXML element this IR node was read from, if known.</summary>
    public XElement? Element { get; init; }

    /// <summary>The URI of the part the source element lived in, if known.</summary>
    public Uri? PartUri { get; init; }

    /// <summary>
    /// True when the block carrying this provenance was delivered by a block-level <c>w:sdt</c>
    /// (content control) the reader unwrapped, rather than being a direct child of its scope's
    /// body/cell. Lives on provenance precisely because it is equality-neutral: SDT unwrap is a
    /// content-transparent normalization (a paragraph reads the same with or without its SDT wrapper —
    /// see the <c>Read_BlockSdt_*</c> tests), so this positional fact must NOT perturb block value
    /// equality or the determinism/diff guarantees.
    /// <para>
    /// The markdown emitter reads it to mirror the ORACLE's block walk: <c>EmitBlocks</c> dispatches
    /// only direct <c>w:p</c>/<c>w:tbl</c>/<c>w:sectPr</c> children and silently skips a <c>w:sdt</c>
    /// wrapper, so the oracle never RENDERS a block an SDT delivers — though it still INDEXES it (via
    /// <c>Descendants</c>), which the IR matches because the block is present in its scope.
    /// </para>
    /// </summary>
    public bool FromBlockSdt { get; init; }

    /// <summary>Always equal to any other <see cref="IrProvenance"/> so provenance is excluded from record equality.</summary>
    public override bool Equals(object? obj) => obj is IrProvenance;

    /// <summary>Always <c>0</c> so provenance contributes nothing to a containing record's hash.</summary>
    public override int GetHashCode() => 0;
}
