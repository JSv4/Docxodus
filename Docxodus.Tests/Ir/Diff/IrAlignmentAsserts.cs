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
    /// <item>Unchanged ⇒ both present, ContentHash AND FormatFingerprint equal.</item>
    /// <item>FormatOnly ⇒ both present, ContentHash equal, FormatFingerprint differs.</item>
    /// <item>Moved ⇒ both present, ContentHash equal (format may differ).</item>
    /// <item>Modified ⇒ both present (no hash constraint).</item>
    /// <item>MovedModified ⇒ both present (no hash constraint — M2.2 fuzzy moved+edited;
    /// ContentHash equality is NOT required and would mean it should have been plain Moved).</item>
    /// <item>Every left/right body block appears in exactly one entry (totality + no duplication),
    /// by reference identity to the input lists.</item>
    /// </list>
    /// </summary>
    public static void AssertInvariants(IrDocument left, IrDocument right, IrBlockAlignment a)
    {
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
                    Assert.Equal(e.Left!.FormatFingerprint, e.Right!.FormatFingerprint);
                    break;
                case IrAlignmentKind.FormatOnly:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    Assert.NotEqual(e.Left!.FormatFingerprint, e.Right!.FormatFingerprint);
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
                    // M2.2 Task 3: fuzzy moved+edited. Both present; ContentHash NOT required equal
                    // (equal ContentHash would mean it should have classified as plain Moved instead).
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    break;
            }

            if (e.Left is not null)
                leftSeen.Add(e.Left);
            if (e.Right is not null)
                rightSeen.Add(e.Right);
        }

        // Every left/right body block appears in exactly one entry (totality + no duplication).
        AssertSameMultiset(left.Body.Blocks, leftSeen, "left");
        AssertSameMultiset(right.Body.Blocks, rightSeen, "right");
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
            IrAlignmentKind.Deleted,
        };
        return string.Join(" ", order.Select(k => $"{k}={Count(a, k)}"));
    }
}
