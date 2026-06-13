# 1:N Paragraph Split / N:1 Merge — Implementation Plan (M2.6 follow-on)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give the IR diff engine first-class 1:N paragraph-split and N:1 paragraph-merge semantics so WC-1450 and WC-1830 convert from DEVIATION to genuine PASS (scoreboard 177→179), per the resolved spec `docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md` and its appended adversarial review's MUST-FIX gate (F1.1–F4.3).

**Architecture:** Two new additive `IrEditOpKind` members (`SplitBlock`, `MergeBlock`) on the existing `IrEditOp` record (trailing nullable `SplitMergeAnchors` + `SegmentDiffs` fields — the `TableDiff`/`TextboxDiffs` precedent); two new `IrAlignmentKind` members (`Split`, `Merge`) + a trailing `MultiBlocks` field on `IrAlignedBlock` (the spec leaves the alignment-layer representation implicit — this plan pins it: singular side in the existing `Left`/`Right` field, plural side in `MultiBlocks`, mirroring the op model exactly); one new detection pass in `IrBlockAligner.FillOneGap` (after the table-residue rule, before the 1×1 rule) that handles BOTH entry states (fully-free, and prefix/suffix-already-similarity-paired) via a unified scan that may dissolve a same-gap Modified pairing; segment slicing + per-segment Myers re-diff in a new `IrSplitSegmenter`, giving F3.3's partition invariant for free (each segment diff tiles its slice, so boundaries are implicit in the diff ops); verifier/JSON/revision-renderer/markup-renderer cases purely additive. Everything is gated behind a new diff-time setting `DetectSplitMerge` (default **false** until Task 8 flips it, so every intermediate task leaves the whole suite green).

**Tech stack:** .NET 8 / xUnit. All work in `Docxodus/Ir/Diff/` + `Docxodus.Tests/Ir/Diff/`. No WASM/npm ripple (internal types only — `DocxDiff`'s public surface is untouched; the edit-script JSON is the internal diagnostic wire and gains only optional fields).

**Build/test commands used throughout:**

```bash
dotnet build Docxodus.sln                                          # debug build
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "<F>"    # targeted
dotnet test Docxodus.Tests/Docxodus.Tests.csproj                   # full suite (final gates)
```

**MUST-FIX traceability (the review gate — every item lands in a named task):**

| Finding | Where closed |
|---|---|
| F1.1 (N:M only assert-rejected, not type-rejected) | Task 2 (`AssertSplitMergePairing` asserts `RightAnchor is null` for Split / `LeftAnchor is null` for Merge) + Task 10 (spec §1.1/§1.5 restated) |
| F1.2 (anchor-walker enumeration) | Task 2 step 7 (the walker audit table, each walker extended-or-proven-anchor-free) |
| F1.3 (`IrSegmentDiff` scalar wrapper) | Adopted: NO wrapper record — `SegmentDiffs` is `IrNodeList<IrTokenDiff>?` directly (Task 1) |
| F2.1 (merge is confidence, not deviation-closure) | Merge is implemented alongside split in every layer but FRAMED as apply-path confidence + fuzzer support; constructed fixture only (Task 4/5/9); Task 10 re-words spec §1.4 |
| F2.2 (overlapping runs → N:M drift) | Task 4 (consumed-rights set; deterministic scan order) + Task 2 (no-shared-anchor assert) + Task 9 (two-adjacent-splits fuzz case) |
| F3.1 (body `Verify` needs a real new case + builder-ordering contract) | Task 5 (new `case`; the N rights must be right-contiguous at the op position — proven by the existing count/order/ReferenceEquals loop) |
| F3.2 (cell path has no identity proof; fixtures are cell-scope) | Task 5 (`ReconstructBlocks` extended; cell reconstruction additionally asserts the produced anchor SEQUENCE equals the right cell's block anchor sequence — a strengthened order/identity check on the path the fixtures take) |
| F3.3 (slice boundaries not serialized) | Task 1/3 (partition invariant: segment diffs' left spans tile the slice; slice lengths derived by summation; asserted in Task 5) |
| F4.1 (thresholds are hypotheses; sweep is a gate) | Task 8 (sweep diagnostic reports margin-to-nearest-flip; pinned values asserted; blocker procedure if no plateau) |
| F4.2 (WC022 identity-reservation interaction) | Task 4 (Unchanged/FormatOnly pairs are NOT promotion candidates — content-equal ⇒ zero unmatched tail, documented + regression-tested; WC022 both-direction round-trip re-asserted in Task 7) |
| F4.3 (empty-mark prune scope in cells) | VERIFIED during planning: cell paragraphs in body tables carry `p:body:…` anchors (IrReader assigns scope `"body"`; only `p:fn:`/`p:en:` are excluded from the prune at `IrRevisionRenderer.cs:325-327`), so the prune DOES fire in cell scope. Task 6 adds an explicit test. |

---

## File structure (locked decomposition)

| File | Change |
|---|---|
| `Docxodus/Ir/Diff/IrEditScript.cs` | +2 enum members, +2 trailing record fields, doc-table rows |
| `Docxodus/Ir/Diff/IrBlockAlignment.cs` | +2 alignment kinds, +`MultiBlocks` trailing field |
| `Docxodus/Ir/Diff/IrSplitSegmenter.cs` | **NEW** — LCS scoring (coverage/slack), run trimming, segment boundary assignment, per-segment diffs |
| `Docxodus/Ir/Diff/IrDiffSettings.cs` | +`DetectSplitMerge`, `SplitCoverageThreshold`, `SplitForeignSlack`, `SplitMaxRunLength` |
| `Docxodus/Ir/Diff/IrBlockAligner.cs` | detection pass in `FillOneGap` + group threading + `EmitEntries` grouping |
| `Docxodus/Ir/Diff/IrEditScriptBuilder.cs` | `ProjectAlignment` Split/Merge cases; move/deletion interleave bookkeeping for groups |
| `Docxodus/Ir/Diff/IrEditScriptJson.cs` | optional `splitMergeAnchors` / `segmentDiffs` write+read |
| `Docxodus/Ir/Diff/IrRevisionRenderer.cs` | Split/Merge dispatch (Fine + compat) |
| `Docxodus/Ir/Diff/IrMarkupRenderer.cs` | `RenderSplitBlock`/`RenderMergeBlock` (anchored-split shape; reuses `MarkParagraphMark` with `RevKind.Ins`/`RevKind.Del`) |
| `Docxodus.Tests/Ir/Diff/IrEditScriptVerifier.cs` | body `case`, `ReconstructBlocks` cases, `AssertSplitMergePairing`, anchor-resolve walk |
| `Docxodus.Tests/Ir/Diff/IrAlignmentAsserts.cs` | Split/Merge invariants + totality over `MultiBlocks` + histogram |
| `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs` | **NEW** — unit tests for segmenter, detection, projection, verifier, JSON golden |
| `Docxodus.Tests/Ir/Diff/IrSplitThresholdSweepTests.cs` | **NEW** — sweep diagnostic + pinned-threshold assert |
| `Docxodus.Tests/Ir/Diff/DiffFuzzer.cs` + `IrDiffFuzzTests.cs` | +`SplitParagraph`/`MergeParagraphs` mutations |
| `Docxodus.Tests/Ir/Diff/IrParityScoreboardTests.cs` | remove WC-1450/WC-1830 deviations; `GenuinePassFloor` 177→179 |
| `Docxodus.Tests/Ir/Diff/IrMarkupRendererTests.cs` | accept/reject round-trip tests for both fixtures + synthetic split/merge |
| `CHANGELOG.md`, `docs/architecture/ir_diff_engine.md`, the spec | Task 10 |

Existing helpers to REUSE, not duplicate: `IrTokenDiffer.Diff` (per-segment re-diff), `IrTokenDiffAsserts.AssertInvariants` (per-segment verification), `MarkParagraphMark` (`IrMarkupRenderer.cs:1071`), `SourceRunModel` + the span-walk inside `RenderModifiedParagraph` (`IrMarkupRenderer.cs:1151` — refactored into a shared helper in Task 7), `RenderTokenOps`/`RenderTokenOpsCompatible` (`IrRevisionRenderer.cs:541/603`).

**Convention pinned for the whole plan (F3.3):** `SegmentDiffs[i]` is a COMPLETE `IrTokenDiff` of the i-th SLICE of the singular side's token list vs the i-th multi-side block's full token list. Spans on the slice side are **slice-local** (0-based within the slice). The slice lengths are NOT stored: slice i's left length = Σ left-span lengths of `SegmentDiffs[i]`'s ops, and the slices must tile the singular side's token list exactly, in order (the partition invariant). For `SplitBlock` the singular side is LEFT (slices of `tokens(L)`, multi-side = the N right blocks in `SplitMergeAnchors`); for `MergeBlock` the singular side is RIGHT (slices of `tokens(R)`, multi-side = the N left blocks).

---

### Task 1: Op model + JSON wire

**Files:**
- Modify: `Docxodus/Ir/Diff/IrEditScript.cs` (enum ~line 72, record ~line 97, remarks table ~line 80)
- Modify: `Docxodus/Ir/Diff/IrEditScriptJson.cs` (`WriteOp` ~line 60, `ReadOp` ~line 227)
- Test: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs` (new file)

- [ ] **Step 1: Write the failing tests** — create `IrSplitMergeTests.cs`:

```csharp
#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>M2.6 split/merge — op model, JSON, segmenter, detection, projection unit tests.</summary>
public class IrSplitMergeTests
{
    private static IrTokenDiff Diff(params IrTokenOp[] ops) => new(IrNodeList.From(ops.ToList()));

    private static IrEditOp SplitOp() => new(
        IrEditOpKind.SplitBlock,
        LeftAnchor: "p:body:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
        RightAnchor: null, TokenDiff: null, MoveGroupId: null, IsMoveSource: null,
        SplitMergeAnchors: IrNodeList.From(new List<string>
        {
            "p:body:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
            "p:body:cccccccccccccccccccccccccccccccc",
        }),
        SegmentDiffs: IrNodeList.From(new List<IrTokenDiff>
        {
            Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 3, 0, 3)),
            Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 2, 0, 2),
                 new IrTokenOp(IrTokenOpKind.Insert, 2, 2, 2, 4)),
        }));

    [Fact]
    public void Split_op_json_round_trips_and_is_deterministic()
    {
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { SplitOp() }));
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));
    }

    [Fact]
    public void Split_op_json_golden_shape()
    {
        var json = IrEditScriptJson.Write(new IrEditScript(IrNodeList.From(new List<IrEditOp> { SplitOp() })));
        Assert.Contains("\"kind\": \"SplitBlock\"", json);
        Assert.Contains("\"splitMergeAnchors\"", json);
        Assert.Contains("\"segmentDiffs\"", json);
        // The singular side rides the EXISTING field; no rightAnchor on a split op.
        Assert.Contains("\"leftAnchor\"", json);
        Assert.DoesNotContain("\"rightAnchor\"", json);
    }

    [Fact]
    public void Scripts_without_splits_serialize_without_new_fields()
    {
        var op = new IrEditOp(IrEditOpKind.InsertBlock, null, "p:body:dddddddddddddddddddddddddddddddd",
            null, null, null);
        var json = IrEditScriptJson.Write(new IrEditScript(IrNodeList.From(new List<IrEditOp> { op })));
        Assert.DoesNotContain("splitMergeAnchors", json);
        Assert.DoesNotContain("segmentDiffs", json);
    }
}
```

- [ ] **Step 2: Run them to confirm they fail to COMPILE** (no `SplitBlock`, no new ctor params):

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests"`
Expected: build error `'IrEditOpKind' does not contain a definition for 'SplitBlock'`.

- [ ] **Step 3: Extend the op model** in `IrEditScript.cs`. Append to `IrEditOpKind` (after `MoveModifyBlock`):

```csharp
    /// <summary>
    /// One LEFT paragraph whose content migrated, in order, across N≥2 RIGHT paragraphs (a paragraph
    /// SPLIT — M2.6). The singular side rides <see cref="IrEditOp.LeftAnchor"/>; the ordered right
    /// anchors ride <see cref="IrEditOp.SplitMergeAnchors"/> with one complete per-segment token diff
    /// each in <see cref="IrEditOp.SegmentDiffs"/> (slice-local left spans; the slices tile the left
    /// token stream exactly — the partition invariant the apply-verifier enforces).
    /// <para><b>N:M is rejected by <c>AssertSplitMergePairing</c> + never emitted by the builder; the
    /// field set physically permits it (nullable fields), so the pairing assert is load-bearing —
    /// a SplitBlock must carry a null <see cref="IrEditOp.RightAnchor"/>.</b></para>
    /// </summary>
    SplitBlock,

    /// <summary>
    /// N≥2 adjacent LEFT paragraphs fused into one RIGHT paragraph (a paragraph MERGE — the byte-mirror
    /// of <see cref="SplitBlock"/>; M2.6). Singular side rides <see cref="IrEditOp.RightAnchor"/>;
    /// the ordered left anchors ride <see cref="IrEditOp.SplitMergeAnchors"/>; <see cref="IrEditOp.SegmentDiffs"/>
    /// holds one diff per left block against the corresponding slice of the right token stream.
    /// <see cref="IrEditOp.LeftAnchor"/> must be null (pairing-assert-enforced; see SplitBlock note).
    /// Shipped alongside split as apply-path CONFIDENCE for the N↔1 reconstruction machinery + fuzzer
    /// coverage — no corpus deviation demands it (the two retained deviations are both splits).
    /// </summary>
    MergeBlock,
```

Extend the record (two trailing defaulted params — every existing positional call site compiles unchanged):

```csharp
internal sealed record IrEditOp(
    IrEditOpKind Kind,
    string? LeftAnchor,
    string? RightAnchor,
    IrTokenDiff? TokenDiff,
    int? MoveGroupId,
    bool? IsMoveSource,
    IrTableDiff? TableDiff = null,
    IrNodeList<IrTextboxDiff>? TextboxDiffs = null,
    IrNodeList<string>? SplitMergeAnchors = null,
    IrNodeList<IrTokenDiff>? SegmentDiffs = null);
```

(No `IrSegmentDiff` wrapper record — review finding F1.3 adopted.) Add two rows to the field-presence `<remarks>` list on the record:

```
SplitBlock: LeftAnchor set, RightAnchor null; SplitMergeAnchors (N≥2, right-doc order) and
SegmentDiffs (same count, slice-local left spans, partition invariant) set; all move fields null.
MergeBlock: RightAnchor set, LeftAnchor null; SplitMergeAnchors = the N left anchors; mirror otherwise.
```

- [ ] **Step 4: JSON write** — in `IrEditScriptJson.WriteOp`, after the `textboxDiffs` block (line ~91), add:

```csharp
        if (op.SplitMergeAnchors is { } smAnchors)
        {
            writer.WriteStartArray("splitMergeAnchors");
            foreach (var a in smAnchors)
                writer.WriteStringValue(a);
            writer.WriteEndArray();
        }
        if (op.SegmentDiffs is { } segDiffs)
        {
            writer.WriteStartArray("segmentDiffs");
            foreach (var d in segDiffs)
                WriteTokenDiff(writer, d);
            writer.WriteEndArray();
        }
```

- [ ] **Step 5: JSON read** — in `ReadOp`, before the constructor call, add:

```csharp
        IrNodeList<string>? splitMergeAnchors = null;
        if (element.TryGetProperty("splitMergeAnchors", out var sma))
        {
            var list = new List<string>();
            foreach (var a in sma.EnumerateArray())
                list.Add(a.GetString()!);
            splitMergeAnchors = IrNodeList.From(list);
        }
        IrNodeList<IrTokenDiff>? segmentDiffs = null;
        if (element.TryGetProperty("segmentDiffs", out var sd))
        {
            var list = new List<IrTokenDiff>();
            foreach (var d in sd.EnumerateArray())
                list.Add(ReadTokenDiff(d));
            segmentDiffs = IrNodeList.From(list);
        }
        return new IrEditOp(kind, leftAnchor, rightAnchor, tokenDiff, moveGroupId, isMoveSource,
            tableDiff, textboxDiffs, splitMergeAnchors, segmentDiffs);
```

- [ ] **Step 6: Run the new tests + the existing JSON/edit-script tests:**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests|FullyQualifiedName~IrEditScriptTests"`
Expected: ALL PASS.

- [ ] **Step 7: Commit**

```bash
git add Docxodus/Ir/Diff/IrEditScript.cs Docxodus/Ir/Diff/IrEditScriptJson.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): SplitBlock/MergeBlock op kinds + JSON wire (M2.6 1:N, additive fields only)"
```

---

### Task 2: Alignment model + pairing assert + anchor-walker audit

**Files:**
- Modify: `Docxodus/Ir/Diff/IrBlockAlignment.cs`
- Modify: `Docxodus.Tests/Ir/Diff/IrAlignmentAsserts.cs`
- Modify: `Docxodus.Tests/Ir/Diff/IrEditScriptVerifier.cs` (`AssertAnchorsResolve` + new `AssertSplitMergePairing`)
- Test: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs`

- [ ] **Step 1: Write the failing tests** (append to `IrSplitMergeTests`):

```csharp
    [Fact]
    public void Pairing_assert_rejects_a_split_op_that_also_sets_RightAnchor() // F1.1: the assert is load-bearing
    {
        var bad = SplitOp() with { RightAnchor = "p:body:eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee" };
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { bad }));
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(script));
    }

    [Fact]
    public void Pairing_assert_rejects_anchor_shared_between_two_split_ops() // F2.2 overlap ceiling
    {
        var a = SplitOp();
        var b = SplitOp() with { LeftAnchor = "p:body:ffffffffffffffffffffffffffffffff" };
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { a, b }));
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(script));
    }

    [Fact]
    public void Pairing_assert_rejects_count_mismatch_and_short_anchor_lists()
    {
        var oneAnchor = SplitOp() with
        {
            SplitMergeAnchors = IrNodeList.From(new List<string> { "p:body:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb" }),
        };
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(
            new IrEditScript(IrNodeList.From(new List<IrEditOp> { oneAnchor }))));
    }

    [Fact]
    public void Pairing_assert_accepts_a_well_formed_split_and_merge()
    {
        var merge = new IrEditOp(IrEditOpKind.MergeBlock, null, "p:body:99999999999999999999999999999999",
            null, null, null,
            SplitMergeAnchors: IrNodeList.From(new List<string>
            {
                "p:body:11111111111111111111111111111111", "p:body:22222222222222222222222222222222",
            }),
            SegmentDiffs: IrNodeList.From(new List<IrTokenDiff>
            {
                Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 2, 0, 2)),
                Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 2, 2, 4)),
            }));
        IrEditScriptVerifier.AssertSplitMergePairing(
            new IrEditScript(IrNodeList.From(new List<IrEditOp> { SplitOp(), merge })));
    }
```

- [ ] **Step 2: Run** `dotnet test ... --filter "FullyQualifiedName~IrSplitMergeTests"` — Expected: FAIL (no `AssertSplitMergePairing`).

- [ ] **Step 3: Alignment model** — in `IrBlockAlignment.cs` append to `IrAlignmentKind`:

```csharp
    /// <summary>One left paragraph split across N≥2 adjacent right paragraphs (M2.6). <c>Left</c> set,
    /// <c>Right</c> null, <see cref="IrAlignedBlock.MultiBlocks"/> = the N right blocks in right order.
    /// Emitted at the FIRST member right block's position; the other members get no entry of their own.</summary>
    Split,

    /// <summary>N≥2 adjacent left paragraphs merged into one right paragraph (M2.6). <c>Right</c> set,
    /// <c>Left</c> null, <see cref="IrAlignedBlock.MultiBlocks"/> = the N left blocks in left order.</summary>
    Merge,
```

and extend the entry record (trailing nullable, all call sites compile):

```csharp
internal sealed record IrAlignedBlock(
    IrAlignmentKind Kind, IrBlock? Left, IrBlock? Right,
    IrNodeList<IrBlock>? MultiBlocks = null);
```

- [ ] **Step 4: `IrAlignmentAsserts.AssertInvariants`** — add cases (and count `MultiBlocks` members toward the totality multiset):

```csharp
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
```

and after the existing `if (e.Right is not null) rightSeen.Add(e.Right);` add:

```csharp
            if (e.MultiBlocks is { } multi)
            {
                if (e.Kind == IrAlignmentKind.Split)
                    rightSeen.AddRange(multi);
                else if (e.Kind == IrAlignmentKind.Merge)
                    leftSeen.AddRange(multi);
            }
```

Add `IrAlignmentKind.Split, IrAlignmentKind.Merge` to the `Histogram` order array.

- [ ] **Step 5: `AssertSplitMergePairing`** — add to `IrEditScriptVerifier` (make it `public static` so tests call it directly; `Verify` calls it in Task 5):

```csharp
    /// <summary>
    /// Shape invariants for SplitBlock/MergeBlock ops (M2.6, review findings F1.1/F2.2/F3.3):
    /// a SplitBlock has a non-null LeftAnchor, a NULL RightAnchor (N:M is physically representable by
    /// the nullable fields, so this assert is the load-bearing scope ceiling), SplitMergeAnchors.Count ≥ 2,
    /// SegmentDiffs non-null with the same count, and no move fields. MergeBlock mirrors (RightAnchor
    /// set, LeftAnchor null). No anchor may appear in two ops' SplitMergeAnchors, and a split's right
    /// anchors / merge's left anchors may not collide with any other op's same-side anchor.
    /// Non-split/merge ops must carry null SplitMergeAnchors/SegmentDiffs.
    /// </summary>
    public static void AssertSplitMergePairing(IrEditScript script)
    {
        var multiAnchorsSeen = new HashSet<string>(System.StringComparer.Ordinal);
        foreach (var op in AllOps(script))
        {
            if (op.Kind is not (IrEditOpKind.SplitBlock or IrEditOpKind.MergeBlock))
            {
                Assert.Null(op.SplitMergeAnchors);
                Assert.Null(op.SegmentDiffs);
                continue;
            }

            Assert.Null(op.MoveGroupId);
            Assert.Null(op.IsMoveSource);
            Assert.Null(op.TokenDiff);
            Assert.Null(op.TableDiff);
            Assert.NotNull(op.SplitMergeAnchors);
            Assert.NotNull(op.SegmentDiffs);
            Assert.True(op.SplitMergeAnchors!.Count >= 2,
                $"{op.Kind} must carry ≥2 SplitMergeAnchors (got {op.SplitMergeAnchors.Count}).");
            Assert.Equal(op.SplitMergeAnchors.Count, op.SegmentDiffs!.Count);

            if (op.Kind == IrEditOpKind.SplitBlock)
            {
                Assert.NotNull(op.LeftAnchor);
                Assert.Null(op.RightAnchor); // F1.1: N:M physically possible; rejected HERE.
            }
            else
            {
                Assert.NotNull(op.RightAnchor);
                Assert.Null(op.LeftAnchor);
            }

            foreach (var a in op.SplitMergeAnchors)
                Assert.True(multiAnchorsSeen.Add(a),
                    $"anchor '{a}' appears in two split/merge ops' SplitMergeAnchors (F2.2 overlap)."); 
        }
    }

    /// <summary>Every op in the script: body, note, textbox-nested, and table-cell-nested.</summary>
    private static IEnumerable<IrEditOp> AllOps(IrEditScript script)
    {
        IEnumerable<IrEditOp> Expand(IrEditOp op)
        {
            yield return op;
            if (op.TextboxDiffs is { } tbx)
                foreach (var d in tbx)
                    foreach (var inner in d.Ops)
                        foreach (var e in Expand(inner))
                            yield return e;
            if (op.TableDiff is { } td)
                foreach (var row in td.RowOps)
                    if (row.CellOps is { } cells)
                        foreach (var cell in cells)
                            if (cell.BlockOps is { } blocks)
                                foreach (var inner in blocks)
                                    foreach (var e in Expand(inner))
                                        yield return e;
        }

        foreach (var op in script.Operations)
            foreach (var e in Expand(op))
                yield return e;
        if (script.NoteOps is { } notes)
            foreach (var n in notes)
                foreach (var op in n.Ops)
                    foreach (var e in Expand(op))
                        yield return e;
    }
```

- [ ] **Step 6: Extend `AssertAnchorsResolve`** (`IrEditScriptVerifier.cs:447`) to walk the plural side — replace the loop body with:

```csharp
        foreach (var op in script.Operations)
        {
            if (op.LeftAnchor is { } la)
                Assert.True(left.AnchorIndex.ContainsKey(la), $"LeftAnchor '{la}' does not resolve in left.AnchorIndex.");
            if (op.RightAnchor is { } ra)
                Assert.True(right.AnchorIndex.ContainsKey(ra), $"RightAnchor '{ra}' does not resolve in right.AnchorIndex.");
            if (op.SplitMergeAnchors is { } multi)
            {
                // Split: plural side = RIGHT anchors; Merge: plural side = LEFT anchors (F1.2).
                var doc = op.Kind == IrEditOpKind.SplitBlock ? right : left;
                string side = op.Kind == IrEditOpKind.SplitBlock ? "right" : "left";
                foreach (var a in multi)
                    Assert.True(doc.AnchorIndex.ContainsKey(a), $"SplitMergeAnchor '{a}' does not resolve in {side}.AnchorIndex.");
            }
        }
```

- [ ] **Step 7: Anchor-walker audit (F1.2)** — grep every reader of `op.LeftAnchor`/`op.RightAnchor` and record the disposition as a comment block above `AssertSplitMergePairing`:

Run: `grep -rn "\.LeftAnchor\|\.RightAnchor" Docxodus/Ir/Diff/ Docxodus.Tests/Ir/Diff/ --include="*.cs" -l`

Expected walkers and their dispositions (verify each; extend in its own task if listed):
| Walker | Disposition |
|---|---|
| `IrEditScriptVerifier.AssertAnchorsResolve` | EXTENDED (this task, step 6) |
| `IrEditScriptVerifier.Verify` / `ReconstructBlocks` | EXTENDED (Task 5 — new cases read `SplitMergeAnchors`) |
| `IrRevisionRenderer.RenderBlockOp` + run-segmentation in `RenderInsDelRun` | EXTENDED (Task 6) |
| `IrMarkupRenderer.RenderBlockOp` / `IsSectionBreakOp` | EXTENDED (Task 7); `IsSectionBreakOp` is anchor-free for split ops (splits are paragraph-only by construction — detection gate) |
| `IrRevisionRenderer.Render` move pre-pass (`IsMoveSource == true`) | ANCHOR-FREE for split/merge (move fields asserted null) |
| `IrEditScriptJson` | EXTENDED (Task 1) |

If the grep surfaces a walker not in this table, STOP and add a step extending it before proceeding.

- [ ] **Step 8: Run + commit**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests|FullyQualifiedName~IrBlockAlignerTests|FullyQualifiedName~IrEditScriptTests"`
Expected: ALL PASS.

```bash
git add Docxodus/Ir/Diff/IrBlockAlignment.cs Docxodus.Tests/Ir/Diff/IrAlignmentAsserts.cs Docxodus.Tests/Ir/Diff/IrEditScriptVerifier.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): Split/Merge alignment kinds + pairing assert + anchor-walker audit (F1.1/F1.2/F2.2)"
```

---

### Task 3: `IrSplitSegmenter` — LCS scoring, run trimming, segment diffs

**Files:**
- Create: `Docxodus/Ir/Diff/IrSplitSegmenter.cs`
- Modify: `Docxodus/Ir/Diff/IrDiffSettings.cs`
- Test: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs`

- [ ] **Step 1: Settings** — add to `IrDiffSettings` (after `MoveMinimumTokenCount`):

```csharp
    /// <summary>
    /// DIFF-TIME setting (M2.6). When true, the aligner's gap fill runs the 1:N paragraph split / N:1
    /// merge containment scan (after similarity pairing, before the 1×1-residue rule), emitting
    /// <see cref="IrEditOpKind.SplitBlock"/>/<see cref="IrEditOpKind.MergeBlock"/> ops. Default FALSE
    /// during the M2.6 build-out; flipped to true when the full pipeline (apply/revisions/markup) lands.
    /// </summary>
    public bool DetectSplitMerge { get; init; } = false;

    /// <summary>
    /// DIFF-TIME setting (M2.6). Minimum in-order LCS coverage of the singular-side paragraph's content
    /// tokens by the candidate run for a split/merge to fire. STARTING HYPOTHESIS 0.90 (spec §2.2) —
    /// the corpus sweep (IrSplitThresholdSweepTests) is the gate that pins the shipped value (F4.1).
    /// </summary>
    public double SplitCoverageThreshold { get; init; } = 0.90;

    /// <summary>
    /// DIFF-TIME setting (M2.6). Maximum fraction of the candidate run's content tokens NOT matched by
    /// the LCS (net-new content, e.g. WC-1830's inserted math paragraph). Starting hypothesis 0.34; swept.
    /// </summary>
    public double SplitForeignSlack { get; init; } = 0.34;

    /// <summary>DIFF-TIME setting (M2.6). Hard cap on a split/merge candidate run's block count
    /// (bounds the per-gap O(G²) candidate scan on pathological gaps).</summary>
    public int SplitMaxRunLength { get; init; } = 8;
```

- [ ] **Step 2: Write the failing segmenter tests** (append to `IrSplitMergeTests`; build token lists via the real tokenizer over `IrTestDocuments`-style paragraphs — reuse the same helper pattern `IrTokenDifferTests` uses to construct `IrParagraph`s, e.g. via `IrTestDocuments.FromBodyXml` + `IrReader.Read`):

```csharp
    // -------- segmenter --------

    private static (IrDocument doc, List<IrParagraph> paras) ReadParas(params string[] texts)
    {
        var xml = string.Concat(texts.Select(t =>
            $"<w:p><w:r><w:t xml:space=\"preserve\">{t}</w:t></w:r></w:p>"));
        var doc = IrReader.Read(IrTestDocuments.FromBodyXml(xml),
            new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept });
        return (doc, doc.Body.Blocks.OfType<IrParagraph>().ToList());
    }

    private static readonly IrDiffSettings S = new() { DetectSplitMerge = true };

    [Fact]
    public void Segmenter_scores_a_clean_split_at_full_coverage_zero_slack()
    {
        var (_, lp) = ReadParas("alpha bravo charlie. delta echo foxtrot.");
        var (_, rp) = ReadParas("alpha bravo charlie. ", "delta echo foxtrot.");
        var score = IrSplitSegmenter.Score(lp[0], new List<IrParagraph> { rp[0], rp[1] }, S);
        Assert.True(score.Coverage >= 0.99, $"coverage {score.Coverage}");
        Assert.True(score.ForeignSlack <= 0.01, $"slack {score.ForeignSlack}");
    }

    [Fact]
    public void Segmenter_scores_keyword_coincidence_below_threshold()
    {
        var (_, lp) = ReadParas("the contract terminates on delivery of the goods.");
        var (_, rp) = ReadParas("the parties agree on many things.", "delivery of pizza is unrelated to the goods here.");
        var score = IrSplitSegmenter.Score(lp[0], new List<IrParagraph> { rp[0], rp[1] }, S);
        Assert.True(score.Coverage < S.SplitCoverageThreshold || score.ForeignSlack > S.SplitForeignSlack,
            $"coincidence must not qualify (cov={score.Coverage}, slack={score.ForeignSlack})");
    }

    [Fact]
    public void Segmenter_segment_diffs_tile_the_left_token_stream_exactly() // F3.3 partition invariant
    {
        var (_, lp) = ReadParas("alpha bravo charlie. delta echo foxtrot.");
        var (_, rp) = ReadParas("alpha bravo charlie. ", "NEW WORDS HERE", "delta echo foxtrot.");
        var rights = new List<IrParagraph> { rp[0], rp[1], rp[2] };
        var diffs = IrSplitSegmenter.ComputeSegmentDiffs(lp[0], rights, S);
        Assert.Equal(3, diffs.Count);
        // F3.3: the segment slices tile the left token stream exactly — slice i's length is the sum
        // of its non-Insert left-span lengths, and the slice lengths sum to the full left token count.
        int leftTotal = IrDiffTokenizer.Tokenize(lp[0], S).Count;
        Assert.Equal(leftTotal, diffs.Sum(d => d.Ops.Where(o => o.Kind != IrTokenOpKind.Insert).Sum(o => o.LeftLength)));
        // And each segment diff right-tiles its right block (IrTokenDiffer invariant, re-checked):
        for (int i = 0; i < 3; i++)
        {
            int rightCount = IrDiffTokenizer.Tokenize(rights[i], S).Count;
            Assert.Equal(rightCount, diffs[i].Ops.Where(o => o.Kind != IrTokenOpKind.Delete).Sum(o => o.RightLength));
        }
    }
```

(If `IrTestDocuments.FromBodyXml` lives in `Docxodus.Tests/Ir/` — it does, the fuzzer uses it — reuse it directly.)

- [ ] **Step 3: Run** — Expected: FAIL (`IrSplitSegmenter` missing).

- [ ] **Step 4: Implement `IrSplitSegmenter.cs`:**

```csharp
#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// M2.6 split/merge segmentation: scores a (singular paragraph, candidate run of paragraphs) pair via an
/// in-order LCS over content-token MatchKeys (coverage + foreign slack, spec §2.2), and computes the
/// per-segment token diffs by slicing the singular side's token stream at the LCS assignment boundaries
/// and re-running <see cref="IrTokenDiffer.Diff"/> per slice — which guarantees every segment diff carries
/// the full IrTokenDiff invariants over (slice, member block), making slice boundaries IMPLICIT in the
/// diff ops (review F3.3: slice i's left length = Σ non-Insert left-span lengths; the slices tile the
/// singular token stream exactly, in order).
/// </summary>
/// <remarks>
/// <para><b>Determinism.</b> Standard O(n·m) LCS DP with the fixed back-walk tie-break (prefer the
/// left-advance on equal subproblem values — the same discipline as
/// <c>IrEditScriptBuilder.LongestCommonSubsequence</c>). Unmatched singular-side tokens attach to the
/// segment of the nearest PRECEDING matched token (leading unmatched → segment 0) — a total, documented
/// rule. No dictionary enumeration feeds output.</para>
/// <para><b>Content tokens.</b> Coverage and slack count only non-Separator, non-Textbox tokens
/// (separators are connective; a masked textbox placeholder is not content) — the same rule
/// <c>IrRevisionRenderer.CountContent</c> applies. The LCS itself runs over ALL tokens so boundary
/// assignment has separator context, but only content-token matches score.</para>
/// </remarks>
internal static class IrSplitSegmenter
{
    internal readonly record struct SplitScore(double Coverage, double ForeignSlack, int MatchedContent);

    /// <summary>Score one candidate: the singular paragraph vs the concatenated run, per spec §2.2.</summary>
    public static SplitScore Score(IrParagraph singular, IReadOnlyList<IrParagraph> run, IrDiffSettings settings)
    {
        var single = IrDiffTokenizer.Tokenize(singular, settings);
        var runTokens = new List<IrDiffToken>();
        foreach (var p in run)
            runTokens.AddRange(IrDiffTokenizer.Tokenize(p, settings));

        var matchedSingle = LcsMatch(single, runTokens, out var matchedRun);

        int singleContent = CountContent(single);
        int runContent = CountContent(runTokens);
        int matchedContent = 0;
        for (int i = 0; i < single.Count; i++)
            if (matchedSingle[i] >= 0 && IsContent(single[i]))
                matchedContent++;
        int runMatchedContent = 0;
        for (int j = 0; j < runTokens.Count; j++)
            if (matchedRun[j] && IsContent(runTokens[j]))
                runMatchedContent++;

        double coverage = singleContent == 0 ? 0.0 : (double)matchedContent / singleContent;
        double slack = runContent == 0 ? 0.0 : (double)(runContent - runMatchedContent) / runContent;
        return new SplitScore(coverage, slack, matchedContent);
    }

    /// <summary>
    /// Per-segment diffs for a confirmed split (singular = LEFT) — or, with the arguments mirrored by the
    /// caller, a merge (singular = RIGHT; the caller swaps each returned diff's sides via
    /// <see cref="MirrorDiff"/>). Returns one COMPLETE IrTokenDiff per run member: slice-local
    /// singular-side spans, full member-side spans.
    /// </summary>
    public static IrNodeList<IrTokenDiff> ComputeSegmentDiffs(
        IrParagraph singular, IReadOnlyList<IrParagraph> run, IrDiffSettings settings)
    {
        var single = IrDiffTokenizer.Tokenize(singular, settings);
        var memberTokens = run.Select(p => IrDiffTokenizer.Tokenize(p, settings)).ToList();
        var flat = new List<IrDiffToken>();
        var memberOfFlat = new List<int>();
        for (int m = 0; m < memberTokens.Count; m++)
            foreach (var t in memberTokens[m])
            {
                flat.Add(t);
                memberOfFlat.Add(m);
            }

        var matchedSingle = LcsMatch(single, flat, out _, out var matchPartner);

        // Assign every singular token to a member segment: a matched token goes to its partner's member;
        // an unmatched token goes to the segment of the nearest preceding matched token (leading → 0).
        var segmentOf = new int[single.Count];
        int current = 0;
        for (int i = 0; i < single.Count; i++)
        {
            if (matchedSingle[i] >= 0)
                current = memberOfFlat[matchPartner[i]];
            segmentOf[i] = current;
        }
        // Segments must be non-decreasing (LCS is in-order); enforce monotonicity defensively.
        for (int i = 1; i < single.Count; i++)
            if (segmentOf[i] < segmentOf[i - 1])
                segmentOf[i] = segmentOf[i - 1];

        var diffs = new List<IrTokenDiff>(run.Count);
        int cursor = 0;
        for (int m = 0; m < run.Count; m++)
        {
            int start = cursor;
            while (cursor < single.Count && segmentOf[cursor] == m)
                cursor++;
            // tokens with segmentOf < m were consumed; segmentOf > m wait for later members. Tokens
            // assigned to earlier members but appearing after (impossible post-monotonicity) cannot occur.
            var slice = Slice(single, start, cursor);
            diffs.Add(IrTokenDiffer.Diff(slice, memberTokens[m], settings));
        }
        // Trailing unmatched singular tokens beyond the last member's assignment: monotonicity puts them
        // in the LAST segment already (current carries forward), so cursor == single.Count here.
        System.Diagnostics.Debug.Assert(cursor == single.Count);
        return IrNodeList.From(diffs);
    }

    /// <summary>Swap a token diff's left/right sides (Insert↔Delete, spans mirrored) — turns a
    /// split-shaped (slice vs member) diff into the merge-shaped (member vs slice) orientation.</summary>
    public static IrTokenDiff MirrorDiff(IrTokenDiff diff)
    {
        var ops = diff.Ops.Select(o => new IrTokenOp(
            o.Kind switch
            {
                IrTokenOpKind.Insert => IrTokenOpKind.Delete,
                IrTokenOpKind.Delete => IrTokenOpKind.Insert,
                var k => k,
            },
            o.RightStart, o.RightEnd, o.LeftStart, o.LeftEnd)).ToList();
        return new IrTokenDiff(IrNodeList.From(ops));
    }

    // ------------------------------------------------------------------ internals

    private static IReadOnlyList<IrDiffToken> Slice(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        var list = new List<IrDiffToken>(end - start);
        for (int i = start; i < end; i++)
            list.Add(tokens[i]);
        return list;
    }

    private static bool IsContent(IrDiffToken t) =>
        t.Kind is not (IrDiffTokenKind.Separator or IrDiffTokenKind.Textbox);

    private static int CountContent(IReadOnlyList<IrDiffToken> tokens)
    {
        int n = 0;
        foreach (var t in tokens)
            if (IsContent(t))
                n++;
        return n;
    }

    private static int[] LcsMatch(
        IReadOnlyList<IrDiffToken> a, IReadOnlyList<IrDiffToken> b, out bool[] bMatched)
        => LcsMatch(a, b, out bMatched, out _);

    /// <summary>Standard LCS DP over MatchKeys. Returns per-a-index matched flag (partner b index or -1);
    /// <paramref name="bMatched"/> marks consumed b indices; <paramref name="partner"/> = a→b mapping.</summary>
    private static int[] LcsMatch(
        IReadOnlyList<IrDiffToken> a, IReadOnlyList<IrDiffToken> b,
        out bool[] bMatched, out int[] partner)
    {
        int n = a.Count, m = b.Count;
        var dp = new int[n + 1, m + 1];
        for (int i = n - 1; i >= 0; i--)
            for (int j = m - 1; j >= 0; j--)
                dp[i, j] = a[i].MatchKey == b[j].MatchKey
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);

        partner = new int[n];
        Array.Fill(partner, -1);
        bMatched = new bool[m];
        for (int i = 0, j = 0; i < n && j < m;)
        {
            if (a[i].MatchKey == b[j].MatchKey) { partner[i] = j; bMatched[j] = true; i++; j++; }
            else if (dp[i + 1, j] >= dp[i, j + 1]) i++;
            else j++;
        }
        return partner;
    }
}
```

NOTE on cost: the DP is O(|L|·|run|) per candidate. Detection (Task 4) calls `Score` per candidate run and `ComputeSegmentDiffs` ONCE per fired group (in the builder), both gap-bounded and capped by `SplitMaxRunLength` — consistent with the aligner's documented G² class.

- [ ] **Step 5: Run + fix until green; then commit**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests"`
Expected: ALL PASS.

```bash
git add Docxodus/Ir/Diff/IrSplitSegmenter.cs Docxodus/Ir/Diff/IrDiffSettings.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): IrSplitSegmenter (LCS coverage/slack scoring + partition-invariant segment diffs) + split settings"
```

---

### Task 4: Detection in `IrBlockAligner.FillOneGap` (states a + b, merge mirror)

**Files:**
- Modify: `Docxodus/Ir/Diff/IrBlockAligner.cs`
- Test: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs`

**Pre-step (evidence before design-freeze): write a THROWAWAY diagnostic** dumping the two fixtures' cell-gap states so the reconcile handles the states the corpus ACTUALLY produces. Add to `IrSplitMergeTests` (keep it as a permanent `[Fact]` with `ITestOutputHelper` — it is the WC-1450/1830 alignment-state record):

- [ ] **Step 1: Diagnostic** —

```csharp
    [Fact]
    public void Diagnostic_fixture_cell_alignment_states() // evidence for the reconcile design (spec §2.3)
    {
        foreach (var (name, l, r) in new[]
        {
            ("WC-1830", "WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx"),
            ("WC-1450", "WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx"),
        })
        {
            var left = IrReader.Read(new WmlDocument(System.IO.Path.Combine("../../../../TestFiles/", l)), WcCorpus.ReadOpts);
            var right = IrReader.Read(new WmlDocument(System.IO.Path.Combine("../../../../TestFiles/", r)), WcCorpus.ReadOpts);
            var script = IrEditScriptBuilder.Build(left, right, new IrDiffSettings()); // detection OFF — baseline states
            _out.WriteLine($"== {name} ==");
            foreach (var op in script.Operations)
                DumpOp(op, "", left, right);
        }
    }
    // DumpOp: recurse TableDiff cell BlockOps printing Kind + anchors + first 40 chars of each side's text.
```

(Write `DumpOp` accordingly; constructor gains `ITestOutputHelper`.) Run it, paste the two cells' op states into the test's comment. This pins which entry states detection must convert (expected from the deviation catalog: WC-1830 — `Modified(L0↔R0) + Insert(math) + Insert/1×1 tail`; WC-1450 — a Modified pairing to the SUFFIX half in one cell, plus the cleaner `Second`-cell shape).

- [ ] **Step 2: Write the failing detection unit tests** (synthetic, via `AlignBlocks` over `ReadParas`-built docs; `S` = settings with `DetectSplitMerge = true`):

```csharp
    private static IrBlockAlignment Align(IrDocument l, IrDocument r, IrDiffSettings s) =>
        IrBlockAligner.Align(l, r, s);

    [Fact]
    public void Detection_fires_for_a_fully_free_clean_split() // state (a)
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.", "unrelated anchor paragraph one two three.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.", "unrelated anchor paragraph one two three.");
        var a = Align(l, r, S);
        IrAlignmentAsserts.AssertInvariants(l, r, a, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(2, split.MultiBlocks!.Count);
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Deleted));
    }

    [Fact]
    public void Detection_absorbs_an_interior_net_new_block() // the WC-1830 math-paragraph shape
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd eee fff. ggg hhh iii jjj kkk lll.");
        var (r, _) = ReadParas("aaa bbb ccc ddd eee fff. ", "zzz", "ggg hhh iii jjj kkk lll.");
        var a = Align(l, r, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(3, split.MultiBlocks!.Count); // interior net-new member absorbed
    }

    [Fact]
    public void Detection_promotes_a_similarity_paired_prefix_with_trailing_tail_inserts() // state (b)
    {
        // Prefix dominant enough that SimilarityPair pairs L with R0 (Jaccard > 0.5), tail falls out free.
        var (l, _) = ReadParas("aaa bbb ccc ddd eee fff ggg hhh iii jjj. kkk lll.");
        var (r, _) = ReadParas("aaa bbb ccc ddd eee fff ggg hhh iii jjj. ", "kkk lll.");
        var a = Align(l, r, S);
        Assert.Single(a.Entries.Where(e => e.Kind == IrAlignmentKind.Split));
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Modified));
    }

    [Fact]
    public void Detection_does_not_fire_on_keyword_coincidence()
    {
        var (l, _) = ReadParas("the contract terminates on delivery of the goods.");
        var (r, _) = ReadParas("the parties agree on many things today.", "delivery of pizza is unrelated to goods.");
        var a = Align(l, r, S);
        Assert.Empty(a.Entries.Where(e => e.Kind is IrAlignmentKind.Split or IrAlignmentKind.Merge));
    }

    [Fact]
    public void Detection_excludes_an_unrelated_edge_insert_from_the_run() // R2 guard
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.",
            "totally unrelated new paragraph words.", "anchor one two three four five.");
        var a = Align(l, r, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(2, split.MultiBlocks!.Count); // edge net-new EXCLUDED → stays a plain Insert
        Assert.Equal(1, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
    }

    [Fact]
    public void Detection_never_promotes_an_identity_reserved_unchanged_pair() // F4.2 / WC022 guard
    {
        // Content-equal prefix (same text) + a following insert: an Unchanged pair has NO unmatched
        // tail (ContentHash-equal ⇒ all tokens matched), so the insert is genuinely new — no split.
        var (l, _) = ReadParas("same text here one two three.");
        var (r, _) = ReadParas("same text here one two three.", "a new paragraph appended after.");
        var a = Align(l, r, S);
        Assert.Empty(a.Entries.Where(e => e.Kind == IrAlignmentKind.Split));
        Assert.Equal(1, IrAlignmentAsserts.Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
    }

    [Fact]
    public void Detection_merge_mirror_fires_for_a_clean_merge()
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four.");
        var a = Align(l, r, S);
        var merge = a.Entries.Single(e => e.Kind == IrAlignmentKind.Merge);
        Assert.Equal(2, merge.MultiBlocks!.Count);
    }

    [Fact]
    public void Detection_two_adjacent_splits_never_share_a_right_block() // F2.2
    {
        var (l, _) = ReadParas("aaa bbb ccc. ddd eee fff.", "ggg hhh iii. jjj kkk lll.");
        var (r, _) = ReadParas("aaa bbb ccc. ", "ddd eee fff.", "ggg hhh iii. ", "jjj kkk lll.");
        var a = Align(l, r, S);
        var splits = a.Entries.Where(e => e.Kind == IrAlignmentKind.Split).ToList();
        Assert.Equal(2, splits.Count);
        var members = splits.SelectMany(e => e.MultiBlocks!).ToList();
        Assert.Equal(members.Count, members.Distinct(ReferenceEqualityComparer.Instance).Count());
    }

    [Fact]
    public void Detection_off_by_default_changes_nothing()
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.");
        var a = Align(l, r, new IrDiffSettings()); // DetectSplitMerge false
        Assert.Empty(a.Entries.Where(e => e.Kind is IrAlignmentKind.Split or IrAlignmentKind.Merge));
    }
```

- [ ] **Step 3: Run** — Expected: FAIL (no Split entries produced).

- [ ] **Step 4: Implement detection.** Changes to `IrBlockAligner`:

(4a) **Group bookkeeping.** Add at the top of `AlignBlocks` (after the kind/match arrays):

```csharp
        // M2.6 split/merge groups: each consumes ONE singular-side block + N≥2 multi-side blocks.
        // Group members are marked in leftKind/rightKind with the Split/Merge kind and leftMatch/
        // rightMatch pointing at the singular partner; EmitEntries collapses each group to ONE entry.
        var splitGroups = new List<(int LeftIndex, List<int> RightIndexes)>();
        var mergeGroups = new List<(int RightIndex, List<int> LeftIndexes)>();
```

Thread both lists through `FillGaps` → `FillOneGap` (new parameters).

(4b) **The scan**, inserted in `FillOneGap` AFTER the unambiguous-table-residue block (line ~373) and BEFORE the 1×1 rule (line ~383):

```csharp
        // M2.6: 1:N split / N:1 merge containment scan (spec §2.2/§2.3) — runs on what similarity
        // declined, plus may DISSOLVE a same-gap Modified pairing whose partner turns out to be one
        // segment of a split (entry state (b): similarity paired the dominant half; the other half
        // was stranded as a free Insert — the exact +1 the WC-1450/1830 deviations document).
        // Unchanged/FormatOnly pairs are NEVER candidates: a content-equal pair has zero unmatched
        // tail by construction, so promoting one could only manufacture a false split — and leaving
        // them untouched preserves the WC022 identity-reservation reject-order invariant (F4.2).
        if (settings.DetectSplitMerge)
        {
            DetectSplits(leftBlocks, rightBlocks, leftFrom, leftTo, rightFrom, rightTo,
                leftKind, rightKind, leftMatch, rightMatch, leftoverLeft, leftoverRight,
                splitGroups, similarity, settings);
            DetectMerges(rightBlocks, leftBlocks, rightFrom, rightTo, leftFrom, leftTo,
                rightKind, leftKind, rightMatch, leftMatch, leftoverRight, leftoverLeft,
                mergeGroups, similarity, settings);
        }
```

(4c) **`DetectSplits`** (write `DetectMerges` as the strict mirror — sides swapped; or implement one generic side-parameterized worker and two thin wrappers, preferred):

```csharp
    /// <summary>
    /// One direction of the M2.6 containment scan. Candidates: each gap paragraph L that is FREE or
    /// Modified-paired (by this gap's SimilarityPair) to a right paragraph. For each, take the maximal
    /// run of ADJACENT right paragraphs that are free (or L's own partner), length-capped by
    /// SplitMaxRunLength; TRIM the run to [first..last] member holding an LCS match with L (edge
    /// net-new members are excluded — the R2 false-positive guard; INTERIOR net-new members, e.g. an
    /// inserted math paragraph between the two halves, are absorbed; empty edge members are excluded
    /// too — they remain plain Inserts the empty-mark prune suppresses). Fire iff the trimmed run has
    /// ≥2 non-empty members, contains L's partner if L was paired, and Score(L, run) clears
    /// SplitCoverageThreshold / SplitForeignSlack. On fire: dissolve the Modified pairing if any,
    /// mark every member consumed, and record the group. Scan order: ascending L index; first
    /// qualifying candidate wins; consumed members never reused (F2.2).
    /// </summary>
```

Concrete algorithm inside (single-side worker; for the merge call the roles of "left/right" swap and the group list/kind constants differ):

```csharp
    private static void DetectSplits(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        int leftFrom, int leftTo, int rightFrom, int rightTo,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        List<int> leftoverLeft, List<int> leftoverRight,
        List<(int LeftIndex, List<int> RightIndexes)> groups,
        IrBlockSimilarity similarity, IrDiffSettings settings)
    {
        for (int li = leftFrom; li < leftTo; li++)
        {
            if (leftBlocks[li] is not IrParagraph lp)
                continue;
            bool free = leftMatch[li] == -1;
            bool pairedModified = !free && leftKind[li] == IrAlignmentKind.Modified
                && leftMatch[li] >= rightFrom && leftMatch[li] < rightTo
                && rightBlocks[leftMatch[li]] is IrParagraph;
            if (!free && !pairedModified)
                continue; // Unchanged/FormatOnly/Moved/etc. are never candidates (F4.2)

            int partner = pairedModified ? leftMatch[li] : -1;

            // Maximal adjacent run of right indices that are free paragraphs (or the partner),
            // anchored around the partner when paired, else scanned over every free window start.
            // Deterministic: try the SMALLEST (a,b) window first that passes trimming+thresholds.
            var candidate = FindQualifyingRun(lp, partner, rightBlocks, rightFrom, rightTo,
                rightMatch, settings);
            if (candidate is not { } run)
                continue;

            // Fire: dissolve the prior pairing (if any), consume everything, record the group.
            if (pairedModified)
            {
                rightKind[partner] = null; // re-stamped Split below
            }
            leftKind[li] = IrAlignmentKind.Split;
            leftMatch[li] = run[0];
            foreach (int rj in run)
            {
                rightKind[rj] = IrAlignmentKind.Split;
                rightMatch[rj] = li;
            }
            groups.Add((li, run));
            leftoverLeft.Remove(li);
            leftoverRight.RemoveAll(run.Contains);
        }
    }
```

`FindQualifyingRun` implementation notes (write it fully):
- Collect the contiguous window of right indices in `[rightFrom, rightTo)` around each start where every index is either free (`rightMatch == -1`) and an `IrParagraph`, or equals `partner`. Cap window length at `settings.SplitMaxRunLength`.
- For each window (smallest `(a,b)` first, expanding b before advancing a), tokenize via `IrSplitSegmenter.Score(lp, windowParagraphs, settings)`; ALSO compute the per-member matched flags (expose a `ScoreDetailed` overload from `IrSplitSegmenter` returning per-member matched-content counts) to trim edges: drop leading/trailing members with zero matched content tokens (this both excludes unrelated edge inserts AND edge empty carriers); re-check: trimmed length ≥2 non-empty members, partner (if any) inside, thresholds pass on the TRIMMED run.
- Return the trimmed run's indices or null.
- If `partner >= 0` require the trimmed run to contain ≥1 free member besides the partner (otherwise nothing to absorb).

(4d) **`EmitEntries` grouping.** Pass the two group lists in. Build lookup sets:

```csharp
        var splitFirstRight = new Dictionary<int, (int LeftIndex, List<int> Rights)>();
        foreach (var g in splitGroups)
            splitFirstRight[g.RightIndexes[0]] = g;
        var mergeByRight = mergeGroups.ToDictionary(g => g.RightIndex);
```

In the right-walk loop, BEFORE the generic emit, handle the new kinds:

```csharp
            if (rightKind[j] == IrAlignmentKind.Split)
            {
                if (splitFirstRight.TryGetValue(j, out var g))
                {
                    entries.Add(new IrAlignedBlock(IrAlignmentKind.Split, leftBlocks[g.LeftIndex], null,
                        IrNodeList.From(g.Rights.Select(rj => rightBlocks[rj]).ToList())));
                    EmitDeletions(deletionsAfterLeft, g.LeftIndex, leftBlocks, entries);
                }
                continue; // non-first members: consumed by the group entry
            }
            if (rightKind[j] == IrAlignmentKind.Merge)
            {
                var g = mergeByRight[j];
                entries.Add(new IrAlignedBlock(IrAlignmentKind.Merge, null, rightBlocks[j],
                    IrNodeList.From(g.LeftIndexes.Select(li2 => leftBlocks[li2]).ToList())));
                foreach (int li2 in g.LeftIndexes)
                    EmitDeletions(deletionsAfterLeft, li2, leftBlocks, entries);
                continue;
            }
```

The deletion-bucketing walk before it needs no change: split/merge members have `leftMatch != -1` so they correctly act as `lastPairedLeft` anchors, and the explicit `EmitDeletions` calls above flush every bucket they anchor.

- [ ] **Step 5: Run the detection tests until green.** Iterate on `FindQualifyingRun` details (the synthetic fixtures in step 2 encode the required behavior). Also re-run the aligner regression suites:

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests|FullyQualifiedName~IrBlockAlignerTests|FullyQualifiedName~IrAlignerAdversarialTests|FullyQualifiedName~IrAlignerCorpusTests"`
Expected: ALL PASS (corpus unaffected — flag is off by default).

- [ ] **Step 6: Commit**

```bash
git add Docxodus/Ir/Diff/IrBlockAligner.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): split/merge containment scan in FillOneGap (states a+b, edge-trim R2 guard, WC022-safe)"
```

---

### Task 5: Builder projection + verifier extension

**Files:**
- Modify: `Docxodus/Ir/Diff/IrEditScriptBuilder.cs` (`ProjectAlignment` switch ~line 504; `IsPairedInPlace`; `BuildSourceInterleave`)
- Modify: `Docxodus.Tests/Ir/Diff/IrEditScriptVerifier.cs` (body switch ~line 67; `ReconstructBlocks` ~line 390; call `AssertSplitMergePairing` from `Verify`)
- Test: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs`

- [ ] **Step 1: Failing end-to-end tests:**

```csharp
    private static (IrDocument L, IrDocument R, IrEditScript S) BuildScript(string[] left, string[] right)
    {
        var (l, _) = ReadParas(left);
        var (r, _) = ReadParas(right);
        var script = IrEditScriptBuilder.Build(l, r, S);
        return (l, r, script);
    }

    [Fact]
    public void Split_script_carries_one_SplitBlock_and_apply_verifies()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five." },
            new[] { "aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four five." });
        var split = Assert.Single(script.Operations.Where(o => o.Kind == IrEditOpKind.SplitBlock));
        Assert.Equal(2, split.SplitMergeAnchors!.Count);
        IrEditScriptVerifier.Verify(l, r, script, S); // count/order/ReferenceEquals proves apply (F3.1)
    }

    [Fact]
    public void Merge_script_carries_one_MergeBlock_and_apply_verifies()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four five." },
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five." });
        Assert.Single(script.Operations.Where(o => o.Kind == IrEditOpKind.MergeBlock));
        IrEditScriptVerifier.Verify(l, r, script, S);
    }

    [Fact]
    public void Split_with_interior_insert_and_prefix_edit_apply_verifies()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd eee fff. ggg hhh iii jjj kkk lll.", "anchor one two three." },
            new[] { "PRE aaa bbb ccc ddd eee fff. ", "zzz", "ggg hhh iii jjj kkk lll.", "anchor one two three." });
        Assert.Single(script.Operations.Where(o => o.Kind == IrEditOpKind.SplitBlock));
        IrEditScriptVerifier.Verify(l, r, script, S);
    }

    [Fact]
    public void Split_script_json_round_trips()
    {
        var (_, _, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh." },
            new[] { "aaa bbb ccc ddd. ", "eee fff ggg hhh." });
        var json = IrEditScriptJson.Write(script);
        Assert.Equal(script, IrEditScriptJson.Read(json));
        Assert.Equal(json, IrEditScriptJson.Write(IrEditScriptJson.Read(json)));
    }
```

- [ ] **Step 2: Run** — Expected: FAIL (builder throws nothing but the alignment Split entries fall through the `switch` un-projected → verifier count mismatch).

- [ ] **Step 3: Builder projection.** In `ProjectAlignment`'s switch add:

```csharp
                case IrAlignmentKind.Split:
                {
                    var lp = (IrParagraph)entry.Left!;
                    var members = entry.MultiBlocks!.Cast<IrParagraph>().ToList();
                    ops.Add(new IrEditOp(IrEditOpKind.SplitBlock,
                        lp.Anchor.ToString(), null, null, null, null, null, null,
                        IrNodeList.From(members.Select(m => m.Anchor.ToString()).ToList()),
                        IrSplitSegmenter.ComputeSegmentDiffs(lp, members, settings)));
                    break;
                }

                case IrAlignmentKind.Merge:
                {
                    var rp = (IrParagraph)entry.Right!;
                    var members = entry.MultiBlocks!.Cast<IrParagraph>().ToList();
                    // Segment diffs are computed singular-vs-members (rp sliced against each left
                    // member) then MIRRORED so each diff reads left-member → right-slice, keeping
                    // the universal "left side = left document" orientation for every consumer.
                    var sliced = IrSplitSegmenter.ComputeSegmentDiffs(rp, members, settings);
                    ops.Add(new IrEditOp(IrEditOpKind.MergeBlock,
                        null, rp.Anchor.ToString(), null, null, null, null, null,
                        IrNodeList.From(members.Select(m => m.Anchor.ToString()).ToList()),
                        IrNodeList.From(sliced.Select(IrSplitSegmenter.MirrorDiff).ToList())));
                    break;
                }
```

NOTE (detection gate guarantees the casts): the scan only forms groups over `IrParagraph` members.

Update `IsPairedInPlace` to include `IrAlignmentKind.Split` (its `entry.Left` is a paired-in-place left), and in `BuildSourceInterleave` add merge members to the `pairedInPlace` set:

```csharp
        foreach (var entry in alignment.Entries)
        {
            if (entry.Left is not null && IsPairedInPlace(entry.Kind))
                pairedInPlace.Add(leftIndex[entry.Left]);
            if (entry.Kind == IrAlignmentKind.Merge && entry.MultiBlocks is { } lefts)
                foreach (var lb in lefts)
                    pairedInPlace.Add(leftIndex[lb]);
        }
```

and after emitting a Merge entry in the projection loop, flush move-sources for each member left (mirror the deletion flush):

```csharp
            if (entry.Kind == IrAlignmentKind.Merge && entry.MultiBlocks is { } mergeLefts)
                foreach (var lb in mergeLefts)
                    EmitSources(sourcesAfterLeft, leftIndex[lb], moves, ops);
```

(The existing `if (entry.Left is not null && IsPairedInPlace(...))` line already handles the Split entry.)

- [ ] **Step 4: Verifier.** (a) Call the pairing assert at the top of `Verify`: after `AssertMovePairing(script);` add `AssertSplitMergePairing(script);`. (b) Add a shared segment-apply helper + the body cases:

```csharp
    /// <summary>Apply one segment diff (slice-local left spans) and return the reconstructed member
    /// token keys; also re-asserts the full token-diff invariant battery over (slice, member) and
    /// returns the slice length consumed (the partition-invariant accumulator, F3.3).</summary>
    private static (IReadOnlyList<string> Tokens, int SliceLen) ApplySegment(
        IReadOnlyList<IrDiffToken> singularTokens, int offset,
        IReadOnlyList<IrDiffToken> memberTokens, IrTokenDiff diff, IrDiffSettings settings)
    {
        int sliceLen = diff.Ops.Where(o => o.Kind != IrTokenOpKind.Insert).Sum(o => o.LeftLength);
        Assert.True(offset + sliceLen <= singularTokens.Count,
            $"segment slice [{offset},{offset + sliceLen}) overruns the singular token stream ({singularTokens.Count}).");
        var slice = new List<IrDiffToken>(sliceLen);
        for (int k = offset; k < offset + sliceLen; k++)
            slice.Add(singularTokens[k]);

        IrTokenDiffAsserts.AssertInvariants(slice, memberTokens, diff, settings);

        var result = new List<string>();
        foreach (var op in diff.Ops)
        {
            switch (op.Kind)
            {
                case IrTokenOpKind.Equal:
                case IrTokenOpKind.FormatChanged:
                    for (int k = op.LeftStart; k < op.LeftEnd; k++)
                        result.Add(slice[k].MatchKey);
                    break;
                case IrTokenOpKind.Insert:
                    for (int k = op.RightStart; k < op.RightEnd; k++)
                        result.Add(memberTokens[k].MatchKey);
                    break;
                case IrTokenOpKind.Delete:
                    break;
            }
        }
        return (result, sliceLen);
    }
```

Body `Verify` switch — new cases (before `case IrEditOpKind.DeleteBlock`):

```csharp
                case IrEditOpKind.SplitBlock:
                {
                    var leftBlock = (IrParagraph)ResolveLeft(left, op.LeftAnchor!);
                    var leftTokens = MaskTextboxKeys(IrDiffTokenizer.Tokenize(leftBlock, settings));
                    int offset = 0;
                    for (int s = 0; s < op.SplitMergeAnchors!.Count; s++)
                    {
                        var rightBlock = ResolveRight(right, op.SplitMergeAnchors[s]);
                        var memberTokens = MaskTextboxKeys(IrDiffTokenizer.Tokenize((IrParagraph)rightBlock, settings));
                        var (tokens, sliceLen) = ApplySegment(leftTokens, offset, memberTokens, op.SegmentDiffs![s], settings);
                        offset += sliceLen;
                        // Each segment pushes ONE reconstructed tuple → the count/order/ReferenceEquals
                        // loop below proves the N right blocks are right-contiguous at the op's position
                        // (the F3.1 builder-ordering obligation, enforced by the EXISTING assertions).
                        reconstructed.Add((rightBlock, tokens, rightBlock));
                    }
                    Assert.Equal(leftTokens.Count, offset); // F3.3 partition invariant
                    break;
                }

                case IrEditOpKind.MergeBlock:
                {
                    var rightBlock = (IrParagraph)ResolveRight(right, op.RightAnchor!);
                    var rightTokens = MaskTextboxKeys(IrDiffTokenizer.Tokenize(rightBlock, settings));
                    var combined = new List<string>();
                    int offset = 0;
                    for (int s = 0; s < op.SplitMergeAnchors!.Count; s++)
                    {
                        var leftMember = (IrParagraph)ResolveLeft(left, op.SplitMergeAnchors[s]);
                        var memberTokens = MaskTextboxKeys(IrDiffTokenizer.Tokenize(leftMember, settings));
                        // Merge diffs read member→slice; mirror to slice→member for ApplySegment's
                        // singular-side orientation, then apply over the RIGHT stream.
                        var mirrored = MirrorForVerify(op.SegmentDiffs![s]);
                        var (tokens, sliceLen) = ApplySegment(rightTokens, offset, memberTokens, mirrored, settings);
                        offset += sliceLen;
                        _ = tokens; // a merge produces ONE right tuple; per-segment application only
                                    // proves slicing; the reconstructed text is the right stream itself.
                        combined.AddRange(ReconstructMergeSegment(memberTokens, op.SegmentDiffs[s]));
                    }
                    Assert.Equal(rightTokens.Count, offset);
                    reconstructed.Add((rightBlock, combined, rightBlock));
                    break;
                }
```

with two small helpers: `MirrorForVerify` = the test-side copy of `IrSplitSegmenter.MirrorDiff` (or make that helper `internal` and reuse — preferred: reuse), and `ReconstructMergeSegment(memberTokens, diff)` = apply the member→slice diff forward (Equal/FormatChanged copy member tokens, Insert takes right-slice tokens — i.e. the merge segment's reconstruction is the slice; simplest correct implementation: return the slice keys recovered from the mirrored apply's `tokens`). Implement carefully; the assertion that matters is `combined` text-equals the right paragraph (the existing loop at `:154-159` checks it).

(c) **Cell path** — `ReconstructBlocks` (`:390`): add the same two cases operating on `leftByAnchor`/`rightByAnchor` (plural anchors resolve against the CELL/note block dictionaries), each split segment `result.Add(...)` one string per member; **strengthened cell identity check (F3.2):** at the end of the two new cases nothing extra is needed for text equality, but ADD — once, at the top of `ReconstructBlocks` — an anchor-sequence assertion:

```csharp
        // F3.2: the cell/note path proves text equality only; strengthen with an anchor-order proof —
        // the right-producing ops must name the right blocks in right-document order.
        var producedRightAnchors = new List<string>();
        // (append inside each case that produces right content: EqualBlock/FormatOnly → the matched
        //  right anchor is op.RightAnchor; Insert → op.RightAnchor; Modify → op.RightAnchor;
        //  SplitBlock → each SplitMergeAnchors[i]; MergeBlock → op.RightAnchor; Move dest → op.RightAnchor)
        // ... after the loop:
        Assert.Equal(rightBlocks.Select(b => b.Anchor.ToString()).ToList(), producedRightAnchors);
```

CAREFUL: `EqualBlock`/`FormatOnlyBlock` in `ReconstructBlocks` currently read only `op.LeftAnchor`; for the anchor-order list use `op.RightAnchor!` (both are set for those kinds). Run the full corpus after this — if any existing corpus case trips the new anchor-order assert, STOP and investigate before weakening (it should hold; the body path already proves it by ReferenceEquals).

- [ ] **Step 5: Run until green** (the Task 5 step-1 tests + the regression suites):

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests|FullyQualifiedName~IrEditScriptTests|FullyQualifiedName~IrEditScriptCorpusTests|FullyQualifiedName~IrTableDifferTests"`
Expected: ALL PASS.

- [ ] **Step 6: Commit**

```bash
git add Docxodus/Ir/Diff/IrEditScriptBuilder.cs Docxodus.Tests/Ir/Diff/IrEditScriptVerifier.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): SplitBlock/MergeBlock projection + apply-verifier extension (F3.1/F3.2/F3.3 closed)"
```

---

### Task 6: Revision renderer (Fine + compat)

**Files:**
- Modify: `Docxodus/Ir/Diff/IrRevisionRenderer.cs` (`RenderBlockOp` ~line 259)
- Test: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs`

- [ ] **Step 1: Failing tests** — fixture-count tests with the flag ON (these are the deviation-closure drivers) + a Fine-mode shape test + the F4.3 cell-prune check:

```csharp
    private static List<IrRevision> FixtureRevisions(string l, string r)
    {
        var settings = IrWmlComparerAdapter.MapSettings(new WmlComparerSettings()) with { DetectSplitMerge = true };
        var left = IrReader.Read(new WmlDocument(System.IO.Path.Combine("../../../../TestFiles/", l)), WcCorpus.ReadOpts);
        var right = IrReader.Read(new WmlDocument(System.IO.Path.Combine("../../../../TestFiles/", r)), WcCorpus.ReadOpts);
        var script = IrEditScriptBuilder.Build(left, right, settings);
        return IrRevisionRenderer.Render(script, left, right, settings).ToList();
    }

    [Fact]
    public void WC1830_compat_revisions_match_oracle_count()
    {
        var revs = FixtureRevisions("WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx");
        Assert.Equal(2, revs.Count); // oracle: Deleted "When you click…add." + Inserted "\n"
    }

    [Fact]
    public void WC1450_compat_revisions_match_oracle_count()
    {
        var revs = FixtureRevisions("WC/WC023-Table-4-Row-Image-Before.docx",
            "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx");
        Assert.Equal(7, revs.Count);
    }

    [Fact]
    public void Fine_mode_split_reports_per_segment_revisions_only()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh." },
            new[] { "aaa bbb ccc ddd. ", "NEW eee fff ggg hhh." });
        var revs = IrRevisionRenderer.Render(script, l, r, S); // Fine
        // Engine truth: the only content change is the inserted "NEW " inside segment 1.
        Assert.All(revs, rv => Assert.NotEqual(IrRevisionType.Deleted, rv.Type));
        Assert.Contains(revs, rv => rv.Type == IrRevisionType.Inserted && rv.Text.Contains("NEW"));
    }

    [Fact]
    public void Cell_scope_empty_mark_prune_fires() // F4.3 verification, pinned as a test
    {
        // A body-table cell paragraph anchors as p:body:… (IrReader assigns scope "body" throughout the
        // body, including table cells; only p:fn:/p:en: are excluded from the prune at
        // IrRevisionRenderer.IsZeroWidthBlock), so the compat empty-mark prune applies in cells.
        // RIGHT's cell gains one EMPTY paragraph: compat mode must report NO revision for it.
        const string cellL =
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
            "<w:p><w:r><w:t>cell text here</w:t></w:r></w:p>" +
            "</w:tc></w:tr></w:tbl>";
        const string cellR =
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
            "<w:p><w:r><w:t>cell text here</w:t></w:r></w:p><w:p/>" +
            "</w:tc></w:tr></w:tbl>";
        var ro = new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept };
        var l = IrReader.Read(IrTestDocuments.FromBodyXml(cellL), ro);
        var r = IrReader.Read(IrTestDocuments.FromBodyXml(cellR), ro);
        var compat = new IrDiffSettings
        {
            DetectSplitMerge = true,
            RevisionGranularity = RevisionGranularity.WmlComparerCompatible,
        };
        var script = IrEditScriptBuilder.Build(l, r, compat);
        var revs = IrRevisionRenderer.Render(script, l, r, compat);
        Assert.Empty(revs); // the empty-mark insert in a CELL is pruned (body-scope anchor)
    }
```

- [ ] **Step 2: Run** — Expected: the fixture tests FAIL with today's deviation counts (3 vs 2, 8 vs 7); the Fine test FAILS (split op unhandled → no revisions at all for the pair).

- [ ] **Step 3: Implement.** In `RenderBlockOp` add:

```csharp
            case IrEditOpKind.SplitBlock:
                RenderSplitBlock(op, ctx, sink);
                break;

            case IrEditOpKind.MergeBlock:
                RenderMergeBlock(op, ctx, sink);
                break;
```

`RenderSplitBlock` (per spec §4.2): resolve the left paragraph's tokens once; walk segments accumulating the slice offset (slice length = Σ non-Insert left lengths, the same derivation the verifier uses); per segment build the SLICE token list and call the existing `RenderTokenOps(segmentDiff, sliceTokens, memberTokens, op.LeftAnchor, memberAnchor, ctx, sink)` — Fine and compat both ride the existing machinery (compat gets region coalescing + affix trim per segment automatically). After the per-segment token revisions, in COMPAT mode only, emit the paragraph-mark account: one `Inserted "\n"` revision per NEW mark — i.e. per segment after the first whose member is reached by pressing Enter (N−1 marks), MINUS marks the oracle's empty-mark prune logic would suppress. Start with the spec's mapping (one `Inserted` with text `"\n"` per split-off boundary, anchored at the member's anchor) and **iterate against the two fixture counts** — this is the one deliberately empirical sub-step; the constraints are: (1) only compat-mode code changes, (2) Fine output = per-segment token revisions only (no synthetic mark revisions — engine truth), (3) no other scoreboard row may regress (Task 8 gate).

`RenderMergeBlock` mirrors: per-member segment diffs rendered via `RenderTokenOps(diff, memberTokens, sliceTokens, memberAnchor, op.RightAnchor, ...)` (diffs are already member→slice oriented from Task 5); compat adds `Deleted ""`/`"\n"` paragraph-mark revisions per removed mark, same iteration rule.

Also extend the compat run-segmentation guard: `RenderBlockOpList` treats `SplitBlock`/`MergeBlock` like any non-ins/del op (they fall to `RenderBlockOp` directly) — verify no change needed (the `if (kind is InsertBlock or DeleteBlock)` test already excludes them).

- [ ] **Step 4: Iterate until the four tests pass.** Use the diagnostic from Task 4 step 1 to inspect the fixture revision lists when counts mismatch:

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrSplitMergeTests"`
Expected: ALL PASS.

- [ ] **Step 5: Commit**

```bash
git add Docxodus/Ir/Diff/IrRevisionRenderer.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): split/merge revision rendering — Fine per-segment truth, compat oracle counts (WC-1450/1830 green under flag)"
```

---

### Task 7: Markup renderer + accept/reject round-trip

**Files:**
- Modify: `Docxodus/Ir/Diff/IrMarkupRenderer.cs` (`RenderBlockOp` ~line 224; refactor shared span-walk out of `RenderModifiedParagraph` ~line 1151)
- Test: `Docxodus.Tests/Ir/Diff/IrMarkupRendererTests.cs`

- [ ] **Step 1: Failing tests** (in `IrMarkupRendererTests`, following its existing accept/reject helper pattern — it already has `AcceptedBodyText`/`RejectedBodyText`-style helpers; reuse them):

```csharp
    // M2.6: split markup round-trip — synthetic body split.
    [Fact]
    public void Split_markup_accept_yields_right_reject_yields_left()
    {
        var left  = TestDoc("aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five.");
        var right = TestDoc("aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four five.");
        var settings = new IrDiffSettings { DetectSplitMerge = true };
        var produced = RenderMarkup(left, right, settings);
        AssertSchemaValid(produced);
        AssertAcceptEquals(produced, right);   // accept ≡ RIGHT (three paragraphs)
        AssertRejectEquals(produced, left);    // reject ≡ LEFT  (re-merged paragraph)
        // The split-off mark is a paragraph-mark revision (empty w:ins in pPr/rPr).
        var body = BodyOf(produced);
        Assert.Contains(body.Descendants(W.pPr).Elements(W.rPr).Elements(W.ins), _ => true);
    }

    [Fact]
    public void Merge_markup_accept_yields_right_reject_yields_left()
    {
        var left  = TestDoc("aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four five.");
        var right = TestDoc("aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five.");
        var settings = new IrDiffSettings { DetectSplitMerge = true };
        var produced = RenderMarkup(left, right, settings);
        AssertSchemaValid(produced);
        AssertAcceptEquals(produced, right);
        AssertRejectEquals(produced, left);
        Assert.Contains(BodyOf(produced).Descendants(W.pPr).Elements(W.rPr).Elements(W.del), _ => true);
    }

    // The two CORPUS fixtures (cell-scope splits): accept→RIGHT, reject→LEFT, schema-valid.
    [Theory]
    [InlineData("WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx")]
    [InlineData("WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx")]
    public void Fixture_split_markup_round_trips(string l, string r)
    {
        var left = new WmlDocument(Path.Combine(SourceDir.FullName, l));
        var right = new WmlDocument(Path.Combine(SourceDir.FullName, r));
        var produced = RenderMarkup(left, right, new IrDiffSettings { DetectSplitMerge = true });
        AssertSchemaValid(produced);
        AssertAcceptEquals(produced, right);
        AssertRejectEquals(produced, left);
    }

    // F4.2 regression: the WC022 adjacent-empty-paragraph fixture stays reject-order-stable BOTH
    // directions with detection ON. This sits next to the existing
    // WC022_adjacent_empty_paragraphs_round_trip_both_directions (IrMarkupRendererTests.cs:801,
    // commit 697611e) and reuses its AssertRoundTrip helper, which accepts an IrDiffSettings.
    [Fact]
    public void WC022_reject_order_invariant_holds_with_detection_on()
    {
        var before = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC022-Image-Math-Para-Before.docx"));
        var after = new WmlDocument(Path.Combine(WcCorpus.WcDir.FullName, "WC022-Image-Math-Para-After.docx"));
        var s = new IrDiffSettings { DetectSplitMerge = true };
        AssertRoundTrip(before, after, s, label: "WC022-split-on-fwd");
        AssertRoundTrip(after, before, s, label: "WC022-split-on-rev");
    }
```

(Use the file's existing helper names — `IrMarkupRendererTests` already has `AssertRoundTrip(left, right, settings?, label:)` plus schema-validation helpers; read the file's helper region first and adapt `TestDoc`/`AssertAcceptEquals`/`AssertRejectEquals`/`AssertSchemaValid` in the synthetic tests above to whatever it actually provides. Do NOT invent parallel helpers if equivalents exist.)

- [ ] **Step 2: Run** — Expected: FAIL (split ops fall through `RenderBlockOp`'s switch → blocks dropped from the body → accept text mismatch).

- [ ] **Step 3: Refactor the span-walk.** Extract the per-token-op content assembly inside `RenderModifiedParagraph` (the `foreach (var tokenOp in tokenDiff.Ops)` block, lines ~1179-1257) into:

```csharp
    /// <summary>Build the run-level content for one token diff over explicit token lists/run models —
    /// shared by RenderModifiedParagraph (whole-paragraph diff) and the M2.6 split/merge segment
    /// rendering (slice diffs). Token spans index the GIVEN lists; char spans resolve through the
    /// tokens' own absolute StartChar/EndChar (slice tokens retain their source-paragraph positions,
    /// so SourceRunModel slicing works unchanged).</summary>
    private static List<XElement> BuildTokenOpContent(
        IrTokenDiff tokenDiff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        SourceRunModel leftRuns, SourceRunModel rightRuns, RenderState state)
```

`RenderModifiedParagraph` becomes a thin caller. Run the FULL markup test suite after this refactor alone (zero behavior change):

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrMarkupRendererTests|FullyQualifiedName~IrMarkupParityScoreboardTests"`
Expected: ALL PASS. Commit the refactor separately:

```bash
git add Docxodus/Ir/Diff/IrMarkupRenderer.cs
git commit -m "refactor(diff): extract BuildTokenOpContent from RenderModifiedParagraph (no behavior change)"
```

- [ ] **Step 4: Implement `RenderSplitBlock`** — the anchored-split shape (spec §3.3), with the mark placement that makes reject re-merge (REJECT of an inserted mark on paragraph k merges k's content into k+1, so the inserted marks go on paragraphs 0..N−2 and the LAST paragraph keeps the original — unmarked — pilcrow):

```csharp
    /// <summary>
    /// M2.6 split markup (anchored-split, spec §3.3): emit N paragraphs. Paragraph i's content is the
    /// i-th segment diff rendered over (left slice, right member) via the shared span walk; its pPr is
    /// the right member's (accepted-state properties). Paragraphs 0..N-2 get an INSERTED paragraph
    /// mark (MarkParagraphMark, RevKind.Ins — the new pilcrows the split introduced); the LAST
    /// paragraph's mark is the original left pilcrow and stays unmarked. ACCEPT keeps the marks → the
    /// N right paragraphs. REJECT removes each inserted mark, merging every paragraph into the next,
    /// and rejects the per-segment ins/del — reconstructing the single LEFT paragraph.
    /// Falls back to conservative whole-block del(left)+ins(members) when a source element is missing.
    /// </summary>
    private static void RenderSplitBlock(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var leftPara = SourceElement(op.LeftAnchor, state.Left);
        var leftTokens = ParagraphTokens(op.LeftAnchor, state.Left, state.Settings);
        if (leftPara == null || op.SplitMergeAnchors is null || op.SegmentDiffs is null)
        {
            EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
            if (op.SplitMergeAnchors is { } anchors)
                foreach (var a in anchors)
                    EmitWholeBlock(a, state.Right, state, sink, RevKind.Ins, fromRight: true);
            return;
        }

        var leftRuns = new SourceRunModel(leftPara);
        int offset = 0;
        int n = op.SplitMergeAnchors.Count;
        for (int i = 0; i < n; i++)
        {
            var memberAnchor = op.SplitMergeAnchors[i];
            var memberPara = SourceElement(memberAnchor, state.Right);
            var memberTokens = ParagraphTokens(memberAnchor, state.Right, state.Settings);
            var diff = op.SegmentDiffs[i];
            int sliceLen = SliceLength(diff);
            var sliceTokens = SliceTokens(leftTokens, offset, sliceLen);
            offset += sliceLen;
            if (memberPara == null)
            {
                EmitWholeBlock(memberAnchor, state.Right, state, sink, RevKind.Ins, fromRight: true);
                continue;
            }

            var rightRuns = new SourceRunModel(memberPara);
            var newPara = new XElement(W.p);
            var rightPPr = memberPara.Element(W.pPr);
            if (rightPPr != null)
                newPara.Add(StripUnids(new XElement(rightPPr)));
            newPara.Add(BuildTokenOpContent(diff, sliceTokens, memberTokens, leftRuns, rightRuns, state));
            if (i < n - 1)
                MarkParagraphMark(newPara, RevKind.Ins, state); // the new pilcrow (RevKind.Ins — F-nit in §3.3)
            sink.Add(newPara);
        }
    }
```

with helpers `SliceLength(IrTokenDiff)` (Σ non-Insert `LeftLength`) and `SliceTokens(list, offset, len)`. NOTE: `SourceRunModel` slicing uses absolute char spans — slice tokens keep their original `StartChar/EndChar`, and `BuildTokenOpContent` resolves spans via `tokens[op.LeftStart].StartChar`, so slice-local INDEX + absolute CHAR positions compose correctly (this is why slices are sub-lists of the original token list, never re-tokenized).

`RenderMergeBlock` is the mirror: N paragraphs; paragraph i's content = segment i's diff (already member→slice oriented) rendered over (member tokens, right-slice tokens) with pPr from the LEFT member for i<N−1 (those paragraphs vanish on accept and must restore left properties on reject) and from the RIGHT paragraph for the last; paragraphs 0..N−2 get `MarkParagraphMark(p, RevKind.Del, state)` (accept of a deleted mark merges into the next paragraph → the single RIGHT paragraph; reject restores the marks → the N LEFT paragraphs).

Add the dispatch arms in `RenderBlockOp`:

```csharp
            case IrEditOpKind.SplitBlock:
                RenderSplitBlock(op, state, sink);
                break;
            case IrEditOpKind.MergeBlock:
                RenderMergeBlock(op, state, sink);
                break;
```

(The table cell path reaches these automatically — `RenderModifyRow` dispatches cell BlockOps through the same `RenderBlockOp`.)

- [ ] **Step 5: Run until green:**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~IrMarkupRendererTests|FullyQualifiedName~IrMarkupParityScoreboardTests|FullyQualifiedName~IrSplitMergeTests"`
Expected: ALL PASS (round-trips green on both fixtures; the WC022 both-direction test green).

- [ ] **Step 6: Commit**

```bash
git add Docxodus/Ir/Diff/IrMarkupRenderer.cs Docxodus.Tests/Ir/Diff/IrMarkupRendererTests.cs
git commit -m "feat(diff): split/merge native markup — anchored-split shape, MarkParagraphMark reuse, accept/reject round-trip (R4)"
```

---

### Task 8: Default-on flip + threshold sweep + scoreboard ratchet

**Files:**
- Modify: `Docxodus/Ir/Diff/IrDiffSettings.cs` (`DetectSplitMerge` default → `true`)
- Modify: `Docxodus.Tests/Ir/Diff/IrParityScoreboardTests.cs` (remove the two deviations; floors)
- Create: `Docxodus.Tests/Ir/Diff/IrSplitThresholdSweepTests.cs`
- Modify: `Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs` (drop the explicit `DetectSplitMerge = true` overrides where now redundant — keep `S` but note default)

- [ ] **Step 1: Flip the default** to `true` (update the doc comment: "Default true; set false for strict 1:1 op semantics").

- [ ] **Step 2: Run the FULL test suite.** This is the no-silent-regression gate (apply-verifier over the corpus both directions, both scoreboards, fuzz, markup round-trips):

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj`
Expected outcome to handle:
- `IrParityScoreboardTests` FAILS with **STALE DEVIATION** for WC-1450 and WC-1830 (they now pass) — that is the success signal.
- Anything ELSE failing = a real regression from default-on; fix it (likely a threshold blast-radius row — see step 4) before touching the scoreboard.

- [ ] **Step 3: Scoreboard update** — in `IrParityScoreboardTests`:
  1. DELETE the `["WC-1450"]` and `["WC-1830"]` entries; replace with a dated CLOSED comment in the catalog's style: `// ---- WC-1450/WC-1830 (1:N sub-paragraph split): CLOSED in M2.6 — engine-level SplitBlock/MergeBlock semantics (spec 2026-06-12-subparagraph-split-merge-design.md). The aligner's containment scan pairs the before-paragraph with BOTH after-halves; the compat renderer projects the oracle's del/ins+mark account. No catalog entry — the rows PASS.`
  2. `GenuinePassFloor` 177 → **179**; update its trailing comment.
  3. Update the long narration comment (lines ~100-111): the three retained deviations drop to one class — WC-1920 only... **CHECK**: WC-1920 was CLOSED in M2.5 Task 3 per the comments, so after this the catalog should be EMPTY; `ParityFloor` stays 179 (= the full runnable set, now all genuine PASS). Verify `board.Deviation == 0` in the report output and say so in the comment.

- [ ] **Step 4: Threshold sweep (F4.1 — the gate, not a formality).** Create `IrSplitThresholdSweepTests.cs`:

```csharp
#nullable enable
// M2.6 threshold sweep (review F4.1): the 0.90/0.34 defaults are hypotheses until this sweep
// proves a stable plateau over the corpus. Pattern: the M2.4b WS-B 0.67/≥8 sweep.

[Trait("Category", "Parity")]
public class IrSplitThresholdSweepTests
{
    [Fact]
    public void Sweep_split_thresholds_over_the_scoreboard_corpus()
    {
        // For coverage in {0.80,0.85,0.88,0.90,0.92,0.95} × slack in {0.20,0.27,0.34,0.40,0.50}:
        //   run the WC003 row set (reuse WC003_Compare_Rows() — make it internal/static-accessible
        //   or duplicate the row list) through IrWmlComparerAdapter with the candidate thresholds,
        //   count exact-match rows. Emit the full grid to test output.
        // ASSERT: the shipped (0.90, 0.34) cell attains the grid MAXIMUM pass count, AND every
        //   neighboring cell (±1 step each axis) attains the same count (plateau — margin ≥ 1 step
        //   to the nearest flip, the review's required margin report).
        // If NO plateau exists (the maximum is a single isolated cell), the milestone is BLOCKED per
        // F4.1: do not ship default-on; file the grid as evidence and stop for adjudication.
    }

    [Fact]
    public void Shipped_thresholds_are_pinned()
    {
        var s = new IrDiffSettings();
        Assert.True(s.DetectSplitMerge);
        Assert.Equal(0.90, s.SplitCoverageThreshold);
        Assert.Equal(0.34, s.SplitForeignSlack);
        Assert.Equal(8, s.SplitMaxRunLength);
    }
}
```

Implement the sweep body fully (grid loop, per-cell scoreboard run, output table, plateau assert). If the sweep moves the optimum off (0.90, 0.34), update the defaults AND the pinned test AND the spec — the swept value is the shipped value.

- [ ] **Step 5: Full suite again, including fuzz at elevated seeds:**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj` then `DOCXODUS_FUZZ_SEEDS=500 dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "Category=Fuzz"`
Expected: ALL PASS; fuzz own-oracle green over 500 seeds with detection default-on.

- [ ] **Step 6: Release-config build (warnings-as-errors gate):**

Run: `dotnet build -c Release Docxodus.sln`
Expected: 0 errors.

- [ ] **Step 7: Commit**

```bash
git add Docxodus/Ir/Diff/IrDiffSettings.cs Docxodus.Tests/Ir/Diff/IrParityScoreboardTests.cs Docxodus.Tests/Ir/Diff/IrSplitThresholdSweepTests.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "feat(diff): DetectSplitMerge default-on — scoreboard 179/179 genuine, thresholds sweep-pinned (F4.1)"
```

---

### Task 9: Fuzzer mutation classes

**Files:**
- Modify: `Docxodus.Tests/Ir/Diff/DiffFuzzer.cs` (`MutationKind` ~line 36, `Describe` ~line 77, `PickMutation` ~line 241, `Apply` ~line 278)
- Modify: `Docxodus.Tests/Ir/Diff/IrDiffFuzzTests.cs` (fixed-seed regression Theory only if a new seed is interesting)
- Test: existing fuzz harness (no new wiring needed — `RunOwnOracle` covers apply/JSON/determinism automatically)

- [ ] **Step 1: Add the mutations:**

```csharp
        /// <summary>Split one body paragraph at a word boundary into two paragraphs. (Comparable class —
        /// both engines account a clean split as equal-content + an inserted paragraph mark.)</summary>
        SplitParagraph,

        /// <summary>Merge two adjacent body paragraphs into one (space-joined). (Comparable class.)</summary>
        MergeParagraphs,
```

`Describe` cases:

```csharp
            MutationKind.SplitParagraph => $"SplitParagraph(para={Index}, atWord={Target})",
            MutationKind.MergeParagraphs => $"MergeParagraphs(first={Index})",
```

`PickMutation` pool additions (always available):

```csharp
            MutationKind.SplitParagraph,
            MutationKind.MergeParagraphs,
```

`Apply` cases:

```csharp
            case MutationKind.SplitParagraph:
            {
                if (model.Paragraphs.Count == 0)
                    return false;
                int pi = m.Index % model.Paragraphs.Count;
                var words = model.Paragraphs[pi].WordRuns;
                if (words.Count < 4)
                    return false; // need ≥2 words per half for a detectable split
                int at = 1 + (m.Target % (words.Count - 1)); // split AFTER word index at-1
                var first = Para.Words(words.Take(at).Select(r => r.Text));
                var second = Para.Words(words.Skip(at).Select(r => r.Text));
                model.Paragraphs[pi] = first;
                model.Paragraphs.Insert(pi + 1, second);
                return true;
            }

            case MutationKind.MergeParagraphs:
            {
                if (model.Paragraphs.Count < 2)
                    return false;
                int pi = m.Index % (model.Paragraphs.Count - 1);
                var a = model.Paragraphs[pi];
                var b = model.Paragraphs[pi + 1];
                if (a.WordRuns.Count == 0 || b.WordRuns.Count == 0)
                    return false;
                var merged = Para.Words(a.WordRuns.Select(r => r.Text).Concat(b.WordRuns.Select(r => r.Text)));
                model.Paragraphs[pi] = merged;
                model.Paragraphs.RemoveAt(pi + 1);
                return true;
            }
```

**Comparability decision:** START with both kinds in the comparable class (no `IsComparableClass` exclusion). Run the fuzz battery; the differential check hard-fails ONLY on new-empty regressions, and mismatches are characterized to artifacts. If the artifacts show a systematic framing difference (engines disagree on the split account beyond the char bag), exclude the kinds in `IsComparableClass` with a comment documenting the observed difference — the `RelocateParagraph` precedent. Record the decision either way.

- [ ] **Step 2: The F2.2 adjacent-splits case** — find a seed exercising two adjacent splits OR add a deterministic direct case to `IrSplitMergeTests` (already done in Task 4: `Detection_two_adjacent_splits_never_share_a_right_block`) — additionally add one full-pipeline variant: Build + Verify + JSON round-trip over the two-adjacent-splits doc pair (4 lines, reuse `BuildScript`).

- [ ] **Step 3: Run:**

Run: `DOCXODUS_FUZZ_SEEDS=500 dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "Category=Fuzz"`
Expected: own-oracle green on all seeds; zero new-empty regressions; review the mismatch artifacts for the comparability decision (step 1).

- [ ] **Step 4: Commit**

```bash
git add Docxodus.Tests/Ir/Diff/DiffFuzzer.cs Docxodus.Tests/Ir/Diff/IrDiffFuzzTests.cs Docxodus.Tests/Ir/Diff/IrSplitMergeTests.cs
git commit -m "test(diff): SplitParagraph/MergeParagraphs fuzzer mutations + adjacent-splits pipeline case (F2.2)"
```

---

### Task 10: Docs — CHANGELOG, architecture doc, spec MUST-FIX closures

**Files:**
- Modify: `CHANGELOG.md` (`[Unreleased]` → `### Added`)
- Modify: `docs/architecture/ir_diff_engine.md`
- Modify: `docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md`
- Modify: `docs/superpowers/plans/2026-06-11-ir-diff-layout-program-plan.md` (decision-log entry)

- [ ] **Step 1: CHANGELOG** — under `[Unreleased]` / `### Added`:

```markdown
- IR diff engine (`DocxDiff`): first-class 1:N paragraph-split and N:1 paragraph-merge semantics
  (`SplitBlock`/`MergeBlock` edit-script ops). A paragraph split mid-text (Enter pressed) or two
  paragraphs fused now report as the oracle does — a paragraph-mark revision plus fine-grained
  per-segment edits — instead of an inflated whole-paragraph delete+insert pair. Closes the last
  two GetRevisions parity deviations (WC-1450, WC-1830): the scoreboard is now 179/179 genuine.
  New `IrDiffSettings`: `DetectSplitMerge` (default true), `SplitCoverageThreshold`,
  `SplitForeignSlack`, `SplitMaxRunLength` (sweep-pinned). Edit-script JSON gains optional
  `splitMergeAnchors`/`segmentDiffs` arrays on split/merge ops only — existing scripts serialize
  byte-identically.
```

- [ ] **Step 2: `ir_diff_engine.md`** — add a "1:N split / N:1 merge (M2.6)" section: the op shapes, the field-presence rules, the partition invariant, the detection placement (after similarity, before 1×1), the two entry states, the anchored-split markup shape with the mark-placement rule (inserted marks on paragraphs 0..N−2, last keeps the original pilcrow; the inverse for merge), and the parity status update (179/179).

- [ ] **Step 3: Spec MUST-FIX edits** (the review's documentation items):
  1. §1.1/§1.5 (F1.1): restate as "N:M is rejected by `AssertSplitMergePairing` + never emitted; the field set physically permits it, so the pairing assert is load-bearing."
  2. §1.2 (F1.2): paste the Task 2 walker audit table.
  3. §1.3 (F1.3): note `IrSegmentDiff` was dropped — `SegmentDiffs` is `IrNodeList<IrTokenDiff>` directly.
  4. §1.4 (F2.1): reframe merge as "apply-path confidence + fuzzer coverage; no corpus deviation demanded it."
  5. §3.3 (reviewer nit): correct "already invoked" — the move path used `RevKind.MoveTo`; split is the first `RevKind.Ins` caller of `MarkParagraphMark`. Also record the IMPLEMENTED mark-placement rule (marks on 0..N−2; the spec's §3.2 oracle excerpt shows the oracle marking the SECOND paragraph — note that our placement differs from that excerpt but satisfies the same accept≡right/reject≡left contract the spec adjudicates on).
  6. §4.1 (F3.1/F3.2/F3.3): record the builder-ordering contract, the cell-path anchor-order strengthening, and the partition invariant as SHIPPED.
  7. §5.1: add **R8** ("reconcile promotes an identity-reserved (WC022) pair → reject-order regression — CLOSED: Unchanged/FormatOnly pairs are never candidates; regression-tested both directions").
  8. Header: `Status: DESIGN-RESOLVED — implementation deferred` → `Status: IMPLEMENTED (M2.6, <commit range>)`. Record the swept threshold values.
  9. Program plan decision log: one dated entry (M2.6 split/merge shipped; scoreboard 179/179; merge framing per F2.1).

- [ ] **Step 4: Final full-suite + Release gate, then commit:**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj && dotnet build -c Release Docxodus.sln`
Expected: ALL PASS, 0 warnings-as-errors.

```bash
git add CHANGELOG.md docs/architecture/ir_diff_engine.md docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md docs/superpowers/plans/2026-06-11-ir-diff-layout-program-plan.md
git commit -m "docs(diff): M2.6 1:N split/merge — CHANGELOG, arch doc, spec MUST-FIX closures + IMPLEMENTED status"
```

---

## Self-review (spec coverage)

- §1 op model → Task 1 (kinds/fields), Task 2 (pairing assert), F1.3 adopted (no wrapper).
- §2.1 placement → Task 4 step 4b (after table residue, before 1×1).
- §2.2 criterion → Task 3 (`Score`: in-order LCS coverage + slack, ≥2 non-empty members, order via LCS), thresholds in settings.
- §2.3 entry states (a)+(b) + F4.2 → Task 4 (unified scan with pairing dissolution; Unchanged/FormatOnly excluded with proof + test) + Task 7 WC022 round-trip.
- §2.4 interactions → similarity runs first (scan consumes residue only); moves run after gap fill over leftovers (split members are consumed, not leftovers); low-coverage coarsening untouched (split segments render through the same compat path).
- §2.5 determinism/cost → integer-indexed scans, smallest-(a,b) first, `SplitMaxRunLength` cap, LCS DP gap-bounded.
- §3 markup → Task 7 (anchored-split; `MarkParagraphMark` with `RevKind.Ins`/`Del`; mark placement derived from RevisionProcessor accept/reject semantics — deviation from the §3.2 excerpt documented in Task 10 step 3.5; parity judged on round-trip + counts per §3.3).
- §4.1 apply → Task 5; §4.2 revisions → Task 6; §4.3 JSON → Task 1; §4.4 compat/Fine divergence → Task 6 (Fine = per-segment truth, compat = oracle counts).
- §5.2 fuzzer → Task 9; §5.3 test plan items 1-7 → Tasks 6 (1), 7 (2), 5+8 (3), 1+5 (4), 4 (5), 9 (6), 4+5+7 (7 — constructed merge fixtures); §5.4 sweep → Task 8.
- Known deliberate deviations from the spec text: no `IrSegmentDiff` record (F1.3); alignment-layer representation pinned here (spec was silent); split-mark placement on 0..N−2 rather than byte-matching either oracle strategy (allowed by §3.3/§5.5).

## Execution risk notes for the implementer

1. **Task 6 step 3 is the empirical core** — the compat paragraph-mark account must land exactly on oracle counts 2 (WC-1830) and 7 (WC-1450). Iterate there, never in Fine mode, never in the engine. The Task 4 diagnostic prints what the script actually contains when stuck.
2. **Task 4's `FindQualifyingRun` trimming rule** (drop zero-match edge members) is the R2 guard — if a synthetic test wants an edge empty carrier absorbed instead, the empty-mark prune already keeps counts right without absorbing it; do not weaken the trim.
3. **If the Task 8 sweep finds no plateau, STOP** — that is the F4.1 blocker condition; file the grid and re-adjudicate rather than shipping a knife-edge threshold.
4. Cell-scope splits flow through `IrTableDiffer` → `ProjectAlignment` automatically (shared machinery) — no table-differ changes are needed or wanted.
