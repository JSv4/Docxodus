#nullable enable

using System;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Hand-written, deterministic JSON serializer for <see cref="IrCompositeScript"/> — the consolidate-side
/// counterpart to <see cref="IrEditScriptJson"/>. It mirrors that serializer's per-op shape
/// (<c>kind</c>/anchors/<c>tokenDiff</c>/<c>tableDiff</c>/<c>textboxDiffs</c>/<c>splitMergeAnchors</c>/
/// <c>segmentDiffs</c>) and ADDITIVELY emits, per op, the composite attribution
/// (<c>author</c>, <c>sourceReviewer</c>, optional <c>conflictId</c>, and — for a composed multi-reviewer
/// Modify — <c>authoredTokens</c> and <c>sourceRightAnchors</c>), plus a top-level <c>conflicts</c> array.
/// </summary>
/// <remarks>
/// <para><b>Shape.</b> <c>{"operations":[ … ],"conflicts":[ … ]}</c>. Each operation is the embedded edit op's
/// object (same field order as <see cref="IrEditScriptJson"/>) extended with <c>author</c> then
/// <c>sourceReviewer</c>, then (when present) <c>conflictId</c>, <c>authoredTokens</c>, and
/// <c>sourceRightAnchors</c>. A token op inside <c>tokenDiff</c>/<c>segmentDiffs</c> is the same compact
/// 5-element array <c>[kindCode, ls, le, rs, re]</c> the two-way serializer uses. An <c>authoredTokens</c>
/// entry is <c>{"op":[…],"author":"…","sourceReviewer":N}</c>; a <c>sourceRightAnchors</c> entry is
/// <c>{"reviewer":N,"anchor":"…"}</c>. Each conflict is
/// <c>{"id","baseAnchor","tokenStart","tokenEnd","policy","competitors":[{"author","resultText"}]}</c>.</para>
/// <para><b>Determinism.</b> Field order is fixed in code; numbers are written via
/// <see cref="Utf8JsonWriter"/> (invariant). Two <see cref="Write"/> calls on equal scripts produce
/// byte-identical JSON.</para>
/// </remarks>
internal static class IrCompositeScriptJson
{
    private static readonly JsonWriterOptions WriteOptions = new() { Indented = true };

    public static string Write(IrCompositeScript script)
    {
        ArgumentNullException.ThrowIfNull(script);

        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, WriteOptions))
        {
            writer.WriteStartObject();

            writer.WriteStartArray("operations");
            foreach (var op in script.Operations)
                WriteCompositeOp(writer, op);
            writer.WriteEndArray();

            writer.WriteStartArray("conflicts");
            foreach (var conflict in script.Conflicts)
                WriteConflict(writer, conflict);
            writer.WriteEndArray();

            if (script.NoteOps is { } noteOps && noteOps.Count > 0)
            {
                writer.WriteStartArray("noteOps");
                foreach (var noteOp in noteOps)
                    WriteNoteDiff(writer, noteOp);
                writer.WriteEndArray();
            }

            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    // ------------------------------------------------------------------ composite op

    private static void WriteCompositeOp(Utf8JsonWriter writer, IrCompositeOp compositeOp)
    {
        writer.WriteStartObject();
        WriteEditOpBody(writer, compositeOp.Op);

        // Additive composite attribution.
        writer.WriteString("author", compositeOp.Author);
        writer.WriteNumber("sourceReviewer", compositeOp.SourceReviewer);
        if (compositeOp.ConflictId is { } cid)
            writer.WriteNumber("conflictId", cid);

        if (compositeOp.AuthoredTokens is { } authored)
        {
            writer.WriteStartArray("authoredTokens");
            foreach (var a in authored)
            {
                writer.WriteStartObject();
                writer.WritePropertyName("op");
                WriteTokenOpArray(writer, a.Op);
                writer.WriteString("author", a.Author);
                writer.WriteNumber("sourceReviewer", a.SourceReviewer);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }

        if (compositeOp.SourceRightAnchors is { } sourceRightAnchors)
        {
            writer.WriteStartArray("sourceRightAnchors");
            foreach (var sra in sourceRightAnchors)
            {
                writer.WriteStartObject();
                writer.WriteNumber("reviewer", sra.Reviewer);
                writer.WriteString("anchor", sra.Anchor);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }

        // Additive composed-table attribution (FOLLOW-ON B): the merged tableDiff above is the apply/JSON
        // truth; authoredRows is the renderer/revision attribution view. Absent (null) for all non-table ops,
        // so existing JSON tests are byte-unaffected.
        if (compositeOp.AuthoredRows is { } authoredRows)
        {
            writer.WriteStartArray("authoredRows");
            foreach (var rowOp in authoredRows)
                WriteAuthoredRowOp(writer, rowOp);
            writer.WriteEndArray();
        }

        writer.WriteEndObject();
    }

    /// <summary>Write one composed-table authored row: kind/baseRowAnchor/author/sourceReviewer + (for an
    /// InsertRow) rightRowAnchor + (for a ModifyRow) nested composedCells → composedBlockOps (recursive
    /// composite-op writer).</summary>
    private static void WriteAuthoredRowOp(Utf8JsonWriter writer, IrAuthoredRowOp rowOp)
    {
        writer.WriteStartObject();
        writer.WriteString("kind", rowOp.Kind.ToString());
        if (rowOp.BaseRowAnchor is { } bra) writer.WriteString("baseRowAnchor", bra);
        if (rowOp.RightRowAnchor is { } rra) writer.WriteString("rightRowAnchor", rra);
        writer.WriteString("author", rowOp.Author);
        writer.WriteNumber("sourceReviewer", rowOp.SourceReviewer);
        if (rowOp.ComposedCells is { } cells)
        {
            writer.WriteStartArray("composedCells");
            foreach (var cell in cells)
            {
                writer.WriteStartObject();
                if (cell.BaseCellAnchor is { } bca) writer.WriteString("baseCellAnchor", bca);
                // Additive shell attribution: present only when the merger sourced this cell's shell
                // (w:tcPr) from a reviewer, so pre-existing JSON outputs are byte-unaffected.
                if (cell.ShellSourceReviewer >= 0)
                {
                    writer.WriteNumber("shellSourceReviewer", cell.ShellSourceReviewer);
                    if (cell.ShellRightCellAnchor is { } sra) writer.WriteString("shellRightCellAnchor", sra);
                }
                // Additive cell kind + attribution: present only for a reviewer-inserted/-deleted cell
                // (column add/remove), so pre-existing JSON outputs are byte-unaffected.
                if (cell.Kind != IrAuthoredCellKind.Content)
                {
                    writer.WriteString("cellKind", cell.Kind.ToString());
                    writer.WriteString("cellAuthor", cell.Author);
                }
                if (cell.ComposedBlockOps is { } blockOps)
                {
                    writer.WriteStartArray("composedBlockOps");
                    foreach (var blockOp in blockOps)
                        WriteCompositeOp(writer, blockOp);
                    writer.WriteEndArray();
                }
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ embedded edit op (mirrors IrEditScriptJson)

    /// <summary>Write the embedded <see cref="IrEditOp"/>'s fields INTO an already-open object (no
    /// StartObject/EndObject), so the composite op can append its attribution after them.</summary>
    private static void WriteEditOpBody(Utf8JsonWriter writer, IrEditOp op)
    {
        writer.WriteString("kind", op.Kind.ToString());
        if (op.LeftAnchor is { } left) writer.WriteString("leftAnchor", left);
        if (op.RightAnchor is { } right) writer.WriteString("rightAnchor", right);
        // Move fields belong ONLY to MoveBlock/MoveModifyBlock per the IrEditOp field-presence contract
        // (IrEditScript.cs). The composite merger lowers reviewer moves to Insert/Delete and RETAINS the
        // MoveGroupId/IsMoveSource marker on a lowered move-source DeleteBlock for its internal
        // contested-relocation detection (IrCompositeMerger.LowerStructuralOps); that marker is stripped
        // before emission (IrCompositeMerger.EmitOp), so a Delete/Insert reaching here should already be
        // clean. This kind gate is the belt-and-suspenders second line of defence: only Move* kinds ever
        // serialize move fields, so no lowering can leak them into the public edit-script JSON.
        bool isMoveKind = op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock;
        if (isMoveKind && op.MoveGroupId is { } group) writer.WriteNumber("moveGroupId", group);
        if (isMoveKind && op.IsMoveSource is { } source) writer.WriteBoolean("isMoveSource", source);
        if (op.TokenDiff is { } diff)
        {
            writer.WritePropertyName("tokenDiff");
            WriteTokenDiff(writer, diff);
        }
        if (op.TableDiff is { } tableDiff)
        {
            writer.WritePropertyName("tableDiff");
            WriteTableDiff(writer, tableDiff);
        }
        if (op.TextboxDiffs is { } textboxDiffs)
        {
            writer.WriteStartArray("textboxDiffs");
            foreach (var tbxDiff in textboxDiffs)
            {
                writer.WriteStartObject();
                writer.WriteStartArray("ops");
                foreach (var blockOp in tbxDiff.Ops)
                    WriteEmbeddedOp(writer, blockOp);
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
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
    }

    /// <summary>Write a plain (non-composite) embedded <see cref="IrEditOp"/> as its own object — used for the
    /// nested ops inside table cell/textbox diffs, which carry no composite attribution.</summary>
    private static void WriteEmbeddedOp(Utf8JsonWriter writer, IrEditOp op)
    {
        writer.WriteStartObject();
        WriteEditOpBody(writer, op);
        writer.WriteEndObject();
    }

    private static void WriteTableDiff(Utf8JsonWriter writer, IrTableDiff diff)
    {
        writer.WriteStartObject();
        writer.WriteStartArray("rowOps");
        foreach (var rowOp in diff.RowOps)
            WriteRowOp(writer, rowOp);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private static void WriteRowOp(Utf8JsonWriter writer, IrRowOp op)
    {
        writer.WriteStartObject();
        writer.WriteString("kind", op.Kind.ToString());
        if (op.LeftRowAnchor is { } l) writer.WriteString("leftRowAnchor", l);
        if (op.RightRowAnchor is { } r) writer.WriteString("rightRowAnchor", r);
        if (op.MoveGroupId is { } g) writer.WriteNumber("moveGroupId", g);
        if (op.IsMoveSource is { } s) writer.WriteBoolean("isMoveSource", s);
        if (op.CellOps is { } cellOps)
        {
            writer.WriteStartArray("cellOps");
            foreach (var cellOp in cellOps)
                WriteCellOp(writer, cellOp);
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
    }

    private static void WriteCellOp(Utf8JsonWriter writer, IrCellOp op)
    {
        writer.WriteStartObject();
        if (op.LeftCellAnchor is { } l) writer.WriteString("leftCellAnchor", l);
        if (op.RightCellAnchor is { } r) writer.WriteString("rightCellAnchor", r);
        if (op.BlockOps is { } blockOps)
        {
            writer.WriteStartArray("blockOps");
            foreach (var blockOp in blockOps)
                WriteEmbeddedOp(writer, blockOp);
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
    }

    private static void WriteTokenDiff(Utf8JsonWriter writer, IrTokenDiff diff)
    {
        writer.WriteStartObject();
        writer.WriteStartArray("ops");
        foreach (var tokenOp in diff.Ops)
            WriteTokenOpArray(writer, tokenOp);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    /// <summary>The compact 5-element token-op array <c>[kindCode, leftStart, leftEnd, rightStart, rightEnd]</c>,
    /// identical in encoding to <see cref="IrEditScriptJson"/>.</summary>
    private static void WriteTokenOpArray(Utf8JsonWriter writer, IrTokenOp tokenOp)
    {
        writer.WriteStartArray();
        writer.WriteNumberValue(TokenKindCode(tokenOp.Kind));
        writer.WriteNumberValue(tokenOp.LeftStart);
        writer.WriteNumberValue(tokenOp.LeftEnd);
        writer.WriteNumberValue(tokenOp.RightStart);
        writer.WriteNumberValue(tokenOp.RightEnd);
        writer.WriteEndArray();
    }

    /// <summary>One N-way composed note diff: <c>kind</c> + <c>baseNoteId</c> (base-matched) or
    /// <c>sourceReviewer</c>/<c>reviewerNoteId</c> (reviewer-inserted), with the composed ops written
    /// through the SAME composite-op writer the body uses (per-op attribution included).</summary>
    private static void WriteNoteDiff(Utf8JsonWriter writer, IrCompositeNoteDiff diff)
    {
        writer.WriteStartObject();
        writer.WriteString("kind", diff.Kind.ToString());
        if (diff.BaseNoteId is { } baseId) writer.WriteString("baseNoteId", baseId);
        if (diff.SourceReviewer >= 0) writer.WriteNumber("sourceReviewer", diff.SourceReviewer);
        if (diff.ReviewerNoteId is { } revId) writer.WriteString("reviewerNoteId", revId);
        writer.WriteStartArray("ops");
        foreach (var op in diff.Ops)
            WriteCompositeOp(writer, op);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ conflicts

    private static void WriteConflict(Utf8JsonWriter writer, IrConflict conflict)
    {
        writer.WriteStartObject();
        writer.WriteNumber("id", conflict.Id);
        writer.WriteString("baseAnchor", conflict.BaseAnchor);
        writer.WriteNumber("tokenStart", conflict.TokenStart);
        writer.WriteNumber("tokenEnd", conflict.TokenEnd);
        writer.WriteString("policy", conflict.AppliedPolicy.ToString());
        writer.WriteStartArray("competitors");
        foreach (var competitor in conflict.Competitors)
        {
            writer.WriteStartObject();
            writer.WriteString("author", competitor.Author);
            writer.WriteString("resultText", competitor.ResultText);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ token-kind codes (match IrEditScriptJson)

    private static int TokenKindCode(IrTokenOpKind kind) => kind switch
    {
        IrTokenOpKind.Equal => 0,
        IrTokenOpKind.Insert => 1,
        IrTokenOpKind.Delete => 2,
        IrTokenOpKind.FormatChanged => 3,
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown token op kind."),
    };
}
