// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

public sealed partial class DocxSession
{
    /// <summary>
    /// Compose a high-signal snapshot of the session's edit-state — total anchors,
    /// remaining bracketed placeholders, bare underscore runs, and footnote/comment
    /// counts. Pure composition of existing primitives (<see cref="Project"/>,
    /// <see cref="FindPlaceholders"/>, <see cref="Grep"/>) with no new logic, so
    /// every count is exactly what the caller would compute by hand. Designed as
    /// the canonical "what's left to fill in?" check after a mutation batch.
    /// </summary>
    /// <remarks>
    /// The bare-underscore regex <c>(?&lt;![\[_])_{3,}(?![\]_])</c> uses lookarounds
    /// that exclude both a bracket and an adjacent underscore, so they guard the
    /// boundaries of the maximal underscore run (not just the regex match) and
    /// avoid false positives inside <c>[_____]</c>. Bracketed underscore runs are
    /// surfaced via <see cref="EditSummary.RemainingPlaceholders"/>, so the two
    /// collections are disjoint by construction. Both queries run against
    /// <see cref="ProjectionScopes.All"/> so headers/footers/footnotes/endnotes/comments
    /// are counted symmetrically.
    /// </remarks>
    public EditSummary GetEditSummary()
    {
        ThrowIfDisposed();

        var projection = Project();
        var placeholders = FindPlaceholders(PlaceholderKinds.All, ProjectionScopes.All);
        var underscoreRuns = Grep(@"(?<![\[_])_{3,}(?![\]_])", scope: ProjectionScopes.All);

        int footnoteCount = 0;
        int commentCount = 0;
        foreach (var t in projection.AnchorIndex.Values)
        {
            if (t.Anchor.Kind == "fn" && t.Anchor.Scope == "fn") footnoteCount++;
            else if (t.Anchor.Kind == "cmt" && t.Anchor.Scope == "cmt") commentCount++;
        }

        var main = _doc!.MainDocumentPart;
        int inlineFnRefs = 0;
        if (main is not null)
            inlineFnRefs = main.GetXDocument().Root!.Descendants(W.footnoteReference).Count();

        return new EditSummary
        {
            TotalAnchors = projection.AnchorIndex.Count,
            RemainingPlaceholders = placeholders,
            BareUnderscoreRuns = underscoreRuns,
            FootnoteCount = footnoteCount,
            InlineFootnoteRefCount = inlineFnRefs,
            CommentCount = commentCount,
        };
    }

    /// <summary>
    /// Diffs the projection captured at session construction against the current projection
    /// and returns an anchor-keyed change list. Keyed by <c>(scope, Unid)</c> — the Unid
    /// is stable across mutations and kind flips (a paragraph promoted to a heading keeps
    /// its Unid while its anchor kind goes from "p" to "h"), and the scope qualifier guards
    /// against cross-part Unid collisions (the deterministic Unid scheme seeds each scope's
    /// root with the root element's local name, so two header parts whose first paragraph
    /// has identical structure end up with the same raw Unid in different scopes — see
    /// issue #187). Requires <see cref="DocxSessionSettings.CaptureInitialProjection"/>
    /// to have been <c>true</c> at construction time.
    /// </summary>
    /// <param name="format">Output shape. <see cref="DiffFormat.Json"/> (default) returns
    /// an anchor-keyed JSON array; <see cref="DiffFormat.Unified"/> returns a
    /// <c>patch(1)</c>-compatible unified diff over the markdown projections;
    /// <see cref="DiffFormat.SideBySide"/> returns a two-column human-review diff.</param>
    /// <returns>For <see cref="DiffFormat.Json"/>, a JSON array of <see cref="DiffEntry"/>
    /// records. Entries are grouped by op (all deletes first, then modifies, then inserts);
    /// within each group, by anchor-index iteration order (which is document order in
    /// practice, since the projector builds the index via a depth-first descendant walk).
    /// Returns <c>"[]"</c> when the document has not been mutated since construction.
    /// For <see cref="DiffFormat.Unified"/>, a standard unified diff with <c>--- initial</c>
    /// / <c>+++ current</c> headers and 3 lines of context; empty string when nothing changed.
    /// For <see cref="DiffFormat.SideBySide"/>, a two-column rendering with the initial
    /// projection padded to 72 chars on the left, a single marker character, then the
    /// current projection.</returns>
    /// <exception cref="InvalidOperationException">Thrown when
    /// <see cref="DocxSessionSettings.CaptureInitialProjection"/> was <c>false</c>.</exception>
    /// <exception cref="NotSupportedException">Thrown for <paramref name="format"/> values
    /// outside the defined <see cref="DiffFormat"/> range.</exception>
    public string GetDiff(DiffFormat format = DiffFormat.Json)
    {
        ThrowIfDisposed();
        if (_initialProjection is null)
            throw new InvalidOperationException(
                "GetDiff requires CaptureInitialProjection = true in DocxSessionSettings.");

        var current = Project();

        return format switch
        {
            DiffFormat.Json => SerializeDiff(ComputeDiff(_initialProjection, current)),
            DiffFormat.Unified => SerializeUnifiedDiff(_initialProjection.Markdown, current.Markdown),
            DiffFormat.SideBySide => SerializeSideBySideDiff(_initialProjection.Markdown, current.Markdown),
            _ => throw new NotSupportedException(
                $"DiffFormat.{format} is not a recognized value."),
        };
    }

    /// <summary>
    /// Compare the exact package supplied at session construction with the session's current logical
    /// package and return the stable semantic-change schema. The current checkpoint includes dirty
    /// in-memory part caches without mutating or saving the live session.
    /// </summary>
    /// <remarks>
    /// Requires <see cref="DocxSessionSettings.CaptureInitialProjection"/> to have been enabled at
    /// construction time. The same switch owns both initial projection and initial package capture so
    /// existing callers can disable all baseline memory/cost with one setting.
    /// </remarks>
    public Verification.SemanticChangeSet GetSemanticChanges(
        Verification.SemanticDiffOptions? options = null)
    {
        lock (_mutationGate)
        {
            ThrowIfDisposed();
            if (_initialPackageBytes is null)
                throw new InvalidOperationException(
                    "GetSemanticChanges requires CaptureInitialProjection = true in DocxSessionSettings.");

            // The baseline flows through the SAME checkpoint serialization as the current side
            // (lazily, cached). Comparing the raw opening bytes against an SDK-cloned checkpoint
            // reported clone normalization itself as document changes — an orphan part the clone
            // drops became a spurious Delete, and a stray content-type-less entry the session
            // opened fine failed only the right side's preflight.
            _initialCheckpointBytes ??= NormalizeOpeningPackage(_initialPackageBytes);
            var currentPackageBytes = SerializePackageCheckpoint();
            return Verification.SemanticDiff.Compare(
                _initialCheckpointBytes,
                currentPackageBytes,
                options);
        }
    }

    /// <summary>Compact canonical JSON counterpart of <see cref="GetSemanticChanges"/>.</summary>
    public string GetSemanticChangesJson(Verification.SemanticDiffOptions? options = null) =>
        GetSemanticChanges(options).ToCanonicalJson();

    /// <summary>
    /// Verify the session's normal clean-save bytes. When initial-package capture is enabled, the
    /// exact opening bytes are also used as the baseline for finding disposition and delta policy.
    /// Use <see cref="PrepareDeliverable"/> when the verified bytes will be written or transmitted.
    /// </summary>
    public Verification.DeliverableVerificationResult VerifyDeliverable(
        Verification.DeliverableVerificationOptions? options = null,
        Verification.SemanticChangeSet? expectedSemanticChanges = null,
        IReadOnlyList<Verification.DeliverablePackageChangeExpectation>? expectedPackageChanges = null,
        IReadOnlyList<Verification.DeliverableCompanionArtifactInput>? companionArtifacts = null)
    {
        return PrepareDeliverable(options, expectedSemanticChanges, expectedPackageChanges,
            companionArtifacts).Report;
    }

    /// <summary>
    /// Produce the exact normal clean-save package and verify that same immutable byte snapshot.
    /// Returning the bytes with the report prevents a later serialization from drifting from the
    /// package identity recorded by verification.
    /// </summary>
    public Verification.VerifiedDeliverable PrepareDeliverable(
        Verification.DeliverableVerificationOptions? options = null,
        Verification.SemanticChangeSet? expectedSemanticChanges = null,
        IReadOnlyList<Verification.DeliverablePackageChangeExpectation>? expectedPackageChanges = null,
        IReadOnlyList<Verification.DeliverableCompanionArtifactInput>? companionArtifacts = null)
    {
        lock (_mutationGate)
        {
            ThrowIfDisposed();
            var deliverableBytes = Save(persistAnchorIds: false);
            var report = Verification.DeliverableVerifier.VerifyDeliverable(
                new Verification.DeliverableVerificationRequest
                {
                    DeliverableBytes = deliverableBytes,
                    BaselineBytes = _initialPackageBytes?.ToArray(),
                    ExpectedSemanticChanges = expectedSemanticChanges,
                    ExpectedPackageChanges = expectedPackageChanges
                        ?? Array.Empty<Verification.DeliverablePackageChangeExpectation>(),
                    CompanionArtifacts = companionArtifacts
                        ?? Array.Empty<Verification.DeliverableCompanionArtifactInput>(),
                }, options);
            return new Verification.VerifiedDeliverable
            {
                DeliverableBytes = deliverableBytes,
                Report = report,
            };
        }
    }

    /// <summary>Compact canonical JSON counterpart of <see cref="VerifyDeliverable"/>.</summary>
    public string VerifyDeliverableJson(
        Verification.DeliverableVerificationOptions? options = null) =>
        VerifyDeliverable(options).ToCanonicalJson();

    private static List<DiffEntry> ComputeDiff(MarkdownProjection initial, MarkdownProjection current)
    {
        // Key by (scope, Unid). Two reasons we cannot use Unid alone:
        //   1. AnchorIndex is dual-keyed under non-FullUnid rendering (the same
        //      AnchorTarget is reachable via its full Unid key and its rendered
        //      alias key), so AnchorIndex.Values enumerates the same target twice.
        //   2. The deterministic Unid scheme seeds each scope's root with the root
        //      element's local name ("hdr" for every header part, "ftr" for every
        //      footer part), so two header parts whose first paragraph has the
        //      same content + position end up with identical raw Unids in
        //      different scopes (reproduced on the NVCA Model COI — issue #187).
        // DistinctBy collapses duplicates from case (1); the composite key
        // separates legitimately distinct targets from case (2).
        var initialByKey = initial.AnchorIndex.Values
            .DistinctBy(t => (t.Anchor.Scope, t.Unid))
            .ToDictionary(t => (t.Anchor.Scope, t.Unid));
        var currentByKey = current.AnchorIndex.Values
            .DistinctBy(t => (t.Anchor.Scope, t.Unid))
            .ToDictionary(t => (t.Anchor.Scope, t.Unid));

        var entries = new List<DiffEntry>();

        // Deletes: in initial, missing from current.
        foreach (var (key, target) in initialByKey)
        {
            if (currentByKey.ContainsKey(key)) continue;
            entries.Add(new DiffEntry
            {
                Op = "delete",
                AnchorId = target.Anchor.Id,
                Before = target.TextPreview,
            });
        }

        // Modifies: present in both, text preview OR kind differs.
        // Kind can flip without a text change (e.g., SetParagraphStyle promoting
        // a paragraph to a heading flips Anchor.Kind from "p" to "h" while
        // preserving the Unid and TextPreview).
        foreach (var (key, initialTarget) in initialByKey)
        {
            if (!currentByKey.TryGetValue(key, out var currentTarget)) continue;
            if (initialTarget.TextPreview == currentTarget.TextPreview
                && initialTarget.Anchor.Kind == currentTarget.Anchor.Kind) continue;
            entries.Add(new DiffEntry
            {
                Op = "modify",
                AnchorId = currentTarget.Anchor.Id,
                Before = initialTarget.TextPreview,
                After = currentTarget.TextPreview,
            });
        }

        // Inserts: in current, missing from initial.
        foreach (var (key, target) in currentByKey)
        {
            if (initialByKey.ContainsKey(key)) continue;
            entries.Add(new DiffEntry
            {
                Op = "insert",
                AnchorId = target.Anchor.Id,
                After = target.TextPreview,
            });
        }

        return entries;
    }

    private static string SerializeDiff(List<DiffEntry> entries)
    {
        // Hand-rolled JSON so SerializeDiff stays trim/AOT-safe; the WASM build
        // ships with reflection-based serialization disabled, so
        // `System.Text.Json.JsonSerializer.Serialize(...)` throws
        // `JsonSerializerIsReflectionDisabled` at runtime in the browser.
        if (entries.Count == 0) return "[]";
        var sb = new System.Text.StringBuilder(entries.Count * 100 + 2);
        sb.Append('[');
        for (int i = 0; i < entries.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var e = entries[i];
            sb.Append("{\"op\":\"").Append(e.Op).Append("\"")
              .Append(",\"anchorId\":");
            AppendJsonString(sb, e.AnchorId);
            if (e.Before is not null)
            {
                sb.Append(",\"before\":");
                AppendJsonString(sb, e.Before);
            }
            if (e.After is not null)
            {
                sb.Append(",\"after\":");
                AppendJsonString(sb, e.After);
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static void AppendJsonString(System.Text.StringBuilder sb, string s)
    {
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("X4"));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
    }

    // ─── Line-based LCS for DiffFormat.Unified / SideBySide ────────────────
    //
    // Hand-rolled O(n*m) LCS over arrays of lines. We deliberately avoid pulling
    // in DiffPlex / DiffMatchPatch — the WASM build disables reflection-based
    // serialization and we want this path to stay AOT-friendly without a NuGet
    // edge case. The unified path is parseable by patch(1); the side-by-side
    // path mirrors `diff -y` markers.

    private enum LineDiffKind { Equal, Delete, Insert }

    private readonly record struct LineDiffOp(LineDiffKind Kind, int AIdx, int BIdx);

    private static List<LineDiffOp> ComputeLineDiff(string[] a, string[] b)
    {
        int n = a.Length, m = b.Length;
        // dp[i, j] = length of LCS of a[..i] and b[..j].
        var dp = new int[n + 1, m + 1];
        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                dp[i, j] = a[i - 1] == b[j - 1]
                    ? dp[i - 1, j - 1] + 1
                    : Math.Max(dp[i - 1, j], dp[i, j - 1]);
            }
        }

        var ops = new List<LineDiffOp>(n + m);
        int x = n, y = m;
        while (x > 0 && y > 0)
        {
            if (a[x - 1] == b[y - 1])
            {
                ops.Add(new LineDiffOp(LineDiffKind.Equal, x - 1, y - 1));
                x--; y--;
            }
            else if (dp[x - 1, y] > dp[x, y - 1])
            {
                ops.Add(new LineDiffOp(LineDiffKind.Delete, x - 1, -1));
                x--;
            }
            else
            {
                // Ties (dp[x-1,y] == dp[x,y-1]) go to Insert during backward traversal
                // so that after List.Reverse() the forward order shows Delete before
                // Insert — the conventional ordering for unified diffs and the
                // precondition for `SerializeSideBySideDiff`'s Delete+Insert →
                // "modify" pairing.
                ops.Add(new LineDiffOp(LineDiffKind.Insert, -1, y - 1));
                y--;
            }
        }
        while (x > 0) { ops.Add(new LineDiffOp(LineDiffKind.Delete, x - 1, -1)); x--; }
        while (y > 0) { ops.Add(new LineDiffOp(LineDiffKind.Insert, -1, y - 1)); y--; }

        ops.Reverse();
        return ops;
    }

    private static string SerializeUnifiedDiff(string initial, string current)
    {
        // Split on '\n' only — the markdown projector emits LF line terminators.
        // Trailing '\n' produces a trailing empty element; that round-trips
        // correctly through patch(1) provided we don't add a phantom newline.
        var a = initial.Split('\n');
        var b = current.Split('\n');
        var ops = ComputeLineDiff(a, b);

        // No changes → empty string. Lets `if (string.IsNullOrEmpty(diff))` be the
        // "did anything change?" check on the call site.
        bool anyChange = false;
        for (int i = 0; i < ops.Count; i++)
        {
            if (ops[i].Kind != LineDiffKind.Equal) { anyChange = true; break; }
        }
        if (!anyChange) return string.Empty;

        var sb = new System.Text.StringBuilder();
        sb.Append("--- initial\n");
        sb.Append("+++ current\n");

        int idx = 0;
        while (idx < ops.Count)
        {
            // Skip leading Equal ops between hunks.
            while (idx < ops.Count && ops[idx].Kind == LineDiffKind.Equal) idx++;
            if (idx >= ops.Count) break;

            int hunkStart = Math.Max(0, idx - UnifiedContextLines);
            int lastChange = idx;
            int scan = idx;
            while (scan < ops.Count)
            {
                if (ops[scan].Kind != LineDiffKind.Equal)
                {
                    lastChange = scan;
                    scan++;
                    continue;
                }

                // Break when we'd have more than 2 * contextLines equal ops between
                // the last change and the next one — that's where one hunk ends and
                // the next begins.
                int gap = 0;
                while (scan < ops.Count && ops[scan].Kind == LineDiffKind.Equal)
                {
                    gap++;
                    if (gap > 2 * UnifiedContextLines) break;
                    scan++;
                }
                if (gap > 2 * UnifiedContextLines) break;
            }
            int hunkEnd = Math.Min(ops.Count, lastChange + UnifiedContextLines + 1);

            // Compute 1-based line numbers and counts for the hunk header.
            int aStart = 0, bStart = 0;
            for (int k = 0; k < hunkStart; k++)
            {
                if (ops[k].Kind != LineDiffKind.Insert) aStart++;
                if (ops[k].Kind != LineDiffKind.Delete) bStart++;
            }
            int aLines = 0, bLines = 0;
            for (int k = hunkStart; k < hunkEnd; k++)
            {
                if (ops[k].Kind != LineDiffKind.Insert) aLines++;
                if (ops[k].Kind != LineDiffKind.Delete) bLines++;
            }

            // Unified-diff convention: when count is 0, the start position is the
            // line *before* the change (so a pure-insert hunk reads "@@ -0,0 +1,N @@").
            // When count is >0, we emit "start+1" to convert from 0-based to 1-based.
            int aHeaderStart = aLines == 0 ? aStart : aStart + 1;
            int bHeaderStart = bLines == 0 ? bStart : bStart + 1;

            sb.Append("@@ -").Append(aHeaderStart).Append(',').Append(aLines)
              .Append(" +").Append(bHeaderStart).Append(',').Append(bLines)
              .Append(" @@\n");

            for (int k = hunkStart; k < hunkEnd; k++)
            {
                var op = ops[k];
                switch (op.Kind)
                {
                    case LineDiffKind.Equal:
                        sb.Append(' ').Append(a[op.AIdx]).Append('\n');
                        break;
                    case LineDiffKind.Delete:
                        sb.Append('-').Append(a[op.AIdx]).Append('\n');
                        break;
                    case LineDiffKind.Insert:
                        sb.Append('+').Append(b[op.BIdx]).Append('\n');
                        break;
                }
            }

            idx = hunkEnd;
        }

        return sb.ToString();
    }

    private static string SerializeSideBySideDiff(string initial, string current)
    {
        var a = initial.Split('\n');
        var b = current.Split('\n');
        var ops = ComputeLineDiff(a, b);
        if (ops.Count == 0) return string.Empty;

        var sb = new System.Text.StringBuilder();
        int i = 0;
        while (i < ops.Count)
        {
            var op = ops[i];

            // Pair an adjacent Delete + Insert into a single "modified" row marked
            // '|' — matches `diff -y`'s presentation and keeps the row count tight
            // when text on a line is rewritten in place.
            if (op.Kind == LineDiffKind.Delete
                && i + 1 < ops.Count
                && ops[i + 1].Kind == LineDiffKind.Insert)
            {
                AppendSideBySideRow(sb, a[op.AIdx], b[ops[i + 1].BIdx], '|');
                i += 2;
                continue;
            }

            switch (op.Kind)
            {
                case LineDiffKind.Equal:
                    AppendSideBySideRow(sb, a[op.AIdx], b[op.BIdx], ' ');
                    break;
                case LineDiffKind.Delete:
                    AppendSideBySideRow(sb, a[op.AIdx], string.Empty, '<');
                    break;
                case LineDiffKind.Insert:
                    AppendSideBySideRow(sb, string.Empty, b[op.BIdx], '>');
                    break;
            }
            i++;
        }

        return sb.ToString();
    }

    private static void AppendSideBySideRow(System.Text.StringBuilder sb, string left, string right, char marker)
    {
        // Truncate (with U+2026 tail) anything past the column width so the marker
        // column stays aligned. The right column is allowed to run to end-of-line —
        // a terminal will wrap it; a viewer that hard-wraps can post-process.
        string leftDisp = left.Length > SideBySideColumnWidth
            ? string.Concat(left.AsSpan(0, SideBySideColumnWidth - 1), "…")
            : left.PadRight(SideBySideColumnWidth);
        sb.Append(leftDisp).Append(' ').Append(marker).Append(' ').Append(right).Append('\n');
    }
}
