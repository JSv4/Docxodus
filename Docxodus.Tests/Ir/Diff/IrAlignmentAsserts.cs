#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Shared invariant checks for <see cref="IrBlockAligner"/> output, extracted so the Task 2 unit
/// tests, the Task 3 WC-corpus smoke, and the adversarial/scale fixtures all assert the SAME
/// totality + per-kind hash invariants the plan pins.
/// </summary>
internal static class IrAlignmentAsserts
{
    /// <summary>
    /// The aligner invariants the plan pins — run against EVERY case's output:
    /// <list type="bullet">
    /// <item>Inserted ⇒ Left null, Right non-null; Deleted ⇒ Left non-null, Right null.</item>
    /// <item>Unchanged ⇒ both present, ContentHash equal AND format-equal under the policy
    /// (<see cref="IrFormatComparison"/>): exact FormatFingerprint under Full, boundary-normalized
    /// modeled-only block signature under ModeledOnly — so a content-equal pair differing ONLY in
    /// unmodeled rPr noise (lang/bCs/iCs/…) is legitimately Unchanged under the default policy even
    /// though its stored FormatFingerprint differs.</item>
    /// <item>FormatOnly ⇒ both present, ContentHash equal, format DIFFERS under the policy.</item>
    /// <item>Moved ⇒ both present, ContentHash equal. Structural blocks may retain an unprojected
    /// format delta while the renderer has no reversible in-move representation for it.</item>
    /// <item>Modified ⇒ both present (no hash constraint).</item>
    /// <item>MovedModified ⇒ both present (no hash constraint). It covers fuzzy moved+edited pairs
    /// and content-equal paragraph moves with a modeled format delta, whose token diff contains
    /// FormatChanged spans.</item>
    /// <item>Split (M2.6) ⇒ Left non-null, Right null, MultiBlocks ≥2 right blocks; the N right
    /// members count toward the right totality multiset.</item>
    /// <item>Merge (M2.6) ⇒ Left null, Right non-null, MultiBlocks ≥2 left blocks; the N left
    /// members count toward the left totality multiset.</item>
    /// <item>Every left/right body block appears in exactly one entry (totality + no duplication),
    /// by reference identity to the input lists.</item>
    /// </list>
    /// </summary>
    public static void AssertInvariants(
        IrDocument left, IrDocument right, IrBlockAlignment a, IrDiffSettings? settings = null)
    {
        settings ??= new IrDiffSettings();
        var leftSeen = new List<IrBlock>();
        var rightSeen = new List<IrBlock>();

        foreach (var e in a.Entries)
        {
            switch (e.Kind)
            {
                case IrAlignmentKind.Inserted:
                    Assert.Null(e.Left);
                    Assert.NotNull(e.Right);
                    break;
                case IrAlignmentKind.Deleted:
                    Assert.NotNull(e.Left);
                    Assert.Null(e.Right);
                    break;
                case IrAlignmentKind.Unchanged:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    Assert.True(FormatEqual(e.Left!, e.Right!, settings),
                        "Unchanged entry must be format-equal under the policy.");
                    break;
                case IrAlignmentKind.FormatOnly:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    Assert.False(FormatEqual(e.Left!, e.Right!, settings),
                        "FormatOnly entry must differ in format under the policy.");
                    break;
                case IrAlignmentKind.Moved:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    break;
                case IrAlignmentKind.Modified:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    break;
                case IrAlignmentKind.MovedModified:
                    // Both sides are present. Content may differ for a fuzzy moved+edit, or it may be equal
                    // for a moved paragraph whose format-only delta needs a reversible token projection.
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    break;
                case IrAlignmentKind.Split:
                    Assert.NotNull(e.Left);
                    Assert.Null(e.Right);
                    Assert.NotNull(e.MultiBlocks);
                    Assert.True(e.MultiBlocks!.Count >= 2, "Split entry needs ≥2 right members.");
                    break;
                case IrAlignmentKind.Merge:
                    Assert.Null(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.NotNull(e.MultiBlocks);
                    Assert.True(e.MultiBlocks!.Count >= 2, "Merge entry needs ≥2 left members.");
                    break;
            }

            // Only Split/Merge entries may carry the plural side (mirrors the op-level
            // AssertSplitMergePairing null rule for non-split/merge ops).
            if (e.Kind is not (IrAlignmentKind.Split or IrAlignmentKind.Merge))
                Assert.True(e.MultiBlocks is null, $"{e.Kind} entry must not carry MultiBlocks.");

            if (e.Left is not null)
                leftSeen.Add(e.Left);
            if (e.Right is not null)
                rightSeen.Add(e.Right);
            if (e.MultiBlocks is { } multi)
            {
                if (e.Kind == IrAlignmentKind.Split)
                    rightSeen.AddRange(multi);
                else if (e.Kind == IrAlignmentKind.Merge)
                    leftSeen.AddRange(multi);
            }
        }

        // Every left/right body block appears in exactly one entry (totality + no duplication).
        AssertSameMultiset(left.Body.Blocks, leftSeen, "left");
        AssertSameMultiset(right.Body.Blocks, rightSeen, "right");

        AssertLeftOrderReconstructible(left, a);
    }

    /// <summary>
    /// ORDER invariant (issue #288). Rejecting a redline materializes each entry's LEFT content in ENTRY
    /// order, so <c>reject ≡ left</c> holds EXACTLY — not merely as a word multiset — only when the left
    /// blocks the alignment owns in place appear in ascending LEFT-document order.
    /// <para><see cref="IrAlignmentKind.Moved"/>/<see cref="IrAlignmentKind.MovedModified"/> entries are
    /// exempt: a move's single entry sits at its DESTINATION, and the edit script re-emits a separate
    /// source op back at the left position. Every other left-owning entry must ascend —
    /// Unchanged/FormatOnly/Modified/Deleted, a <see cref="IrAlignmentKind.Split"/>'s singular left, and a
    /// <see cref="IrAlignmentKind.Merge"/>'s members in member order.</para>
    /// <para>A violation is exactly the #288 failure mode: content survives both directions, but reject
    /// reconstructs the paragraphs permuted.</para>
    /// </summary>
    public static void AssertLeftOrderReconstructible(IrDocument left, IrBlockAlignment a)
    {
        // Reference identity, not record equality: duplicate-content blocks are value-EQUAL records, and
        // telling the two occurrences apart is the whole point of this invariant.
        var indexOf = new Dictionary<IrBlock, int>(BlockReferenceComparer.Instance);
        for (int i = 0; i < left.Body.Blocks.Count; i++)
            indexOf[left.Body.Blocks[i]] = i;

        int previous = -1;
        string previousLabel = "<start>";
        foreach (var e in a.Entries)
        {
            if (e.Kind is IrAlignmentKind.Moved or IrAlignmentKind.MovedModified)
                continue;

            var owned = new List<IrBlock>();
            if (e.Kind == IrAlignmentKind.Merge)
                owned.AddRange(e.MultiBlocks!);
            else if (e.Left is { } single)
                owned.Add(single);

            foreach (var block in owned)
            {
                if (!indexOf.TryGetValue(block, out int index))
                    continue; // a nested-scope block cannot appear here; totality above already proved that
                Assert.True(index > previous,
                    $"Alignment emits left block {index} ({e.Kind}) after left block {previous} " +
                    $"({previousLabel}) — reject would reconstruct the left document out of order.");
                previous = index;
                previousLabel = e.Kind.ToString();
            }
        }
    }

    private sealed class BlockReferenceComparer : IEqualityComparer<IrBlock>
    {
        public static readonly BlockReferenceComparer Instance = new();
        public bool Equals(IrBlock? x, IrBlock? y) => ReferenceEquals(x, y);
        public int GetHashCode(IrBlock obj) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }

    private static void AssertSameMultiset(IReadOnlyList<IrBlock> expected, List<IrBlock> seen, string side)
    {
        Assert.Equal(expected.Count, seen.Count);
        // Reference identity: the aligner must return the very block instances from the input lists.
        var pool = new List<IrBlock>(expected);
        foreach (var b in seen)
        {
            int idx = pool.FindIndex(x => ReferenceEquals(x, b));
            Assert.True(idx >= 0, $"{side} block appeared that was not in the input (or appeared twice).");
            pool.RemoveAt(idx);
        }
        Assert.Empty(pool);
    }

    /// <summary>
    /// Diff-time format equality under the policy, mirroring the aligner's private rule: modeled-only
    /// block signature for a paragraph pair under ModeledOnly, the stored FormatFingerprint otherwise.
    /// </summary>
    private static bool FormatEqual(IrBlock left, IrBlock right, IrDiffSettings settings)
    {
        if (settings.FormatComparison == IrFormatComparison.ModeledOnly
            && left is IrParagraph lp && right is IrParagraph rp)
            return IrModeledFormat.BlockSignature(lp, settings) == IrModeledFormat.BlockSignature(rp, settings);
        return left.FormatFingerprint.Equals(right.FormatFingerprint);
    }

    /// <summary>Count entries of a given kind.</summary>
    public static int Count(IrBlockAlignment a, IrAlignmentKind k) => a.Entries.Count(e => e.Kind == k);

    /// <summary>
    /// A deterministic per-kind histogram string (every kind, in enum order) for ITestOutputHelper logging.
    /// </summary>
    public static string Histogram(IrBlockAlignment a)
    {
        var order = new[]
        {
            IrAlignmentKind.Unchanged, IrAlignmentKind.FormatOnly, IrAlignmentKind.Modified,
            IrAlignmentKind.Moved, IrAlignmentKind.MovedModified, IrAlignmentKind.Inserted,
            IrAlignmentKind.Deleted, IrAlignmentKind.Split, IrAlignmentKind.Merge,
        };
        return string.Join(" ", order.Select(k => $"{k}={Count(a, k)}"));
    }
}
