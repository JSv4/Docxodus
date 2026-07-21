#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Hand-written, deterministic JSON serializer for <see cref="IrEditScript"/> (M2.2 Task 2). Mirrors
/// the <c>IrDiagnosticJson</c> style: one method per node type, fixed field order, no reflection, no
/// timestamps/paths. <see cref="Write"/> and <see cref="Read"/> round-trip exactly:
/// <c>Read(Write(s))</c> is record-equal to <c>s</c>.
/// </summary>
/// <remarks>
/// <para><b>Shape.</b> <c>{"operations":[ … ]}</c>. Each op is an object with a fixed field order:
/// <c>kind</c> (enum name), then any of <c>leftAnchor</c>/<c>rightAnchor</c>/<c>moveGroupId</c>/
/// <c>isMoveSource</c>/<c>bodyFullRewriteGroupId</c>/<c>tokenDiff</c> that are present (absent fields are simply omitted, matching the
/// record's nullability). A <c>tokenDiff</c> is <c>{"ops":[ [kind,ls,le,rs,re], … ]}</c> — each token
/// op a COMPACT 5-element array: an integer kind code (0=Equal,1=Insert,2=Delete,3=FormatChanged) plus
/// the four half-open span bounds. The compact array keeps a large corpus script terse while staying
/// fully self-describing for the reader. For <c>SplitBlock</c>/<c>MergeBlock</c> ops, an optional
/// <c>splitMergeAnchors</c> string array and <c>segmentDiffs</c> array (one tokenDiff object per anchor)
/// are additionally present; both are omitted on every other kind.</para>
/// <para><b>Determinism.</b> Field order is fixed in code; numbers are written via
/// <see cref="Utf8JsonWriter"/> (invariant). Two <see cref="Write"/> calls on equal scripts produce
/// byte-identical JSON.</para>
/// </remarks>
internal static class IrEditScriptJson
{
    private static readonly JsonWriterOptions WriteOptions = new() { Indented = true };

    // ------------------------------------------------------------------ write

    public static string Write(IrEditScript script)
    {
        ArgumentNullException.ThrowIfNull(script);

        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, WriteOptions))
        {
            writer.WriteStartObject();
            writer.WriteStartArray("operations");
            foreach (var op in script.Operations)
                WriteOp(writer, op);
            writer.WriteEndArray();
            if (script.NoteOps is { } noteOps)
            {
                writer.WriteStartArray("noteOps");
                foreach (var noteOp in noteOps)
                    WriteNoteDiff(writer, noteOp);
                writer.WriteEndArray();
            }
            if (script.HeaderFooterOps is { } headerFooterOps)
            {
                writer.WriteStartArray("headerFooterOps");
                foreach (var hfOp in headerFooterOps)
                    WriteHeaderFooterDiff(writer, hfOp);
                writer.WriteEndArray();
            }
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    private static void WriteOp(Utf8JsonWriter writer, IrEditOp op)
    {
        writer.WriteStartObject();
        writer.WriteString("kind", op.Kind.ToString());
        if (op.LeftAnchor is { } left) writer.WriteString("leftAnchor", left);
        if (op.RightAnchor is { } right) writer.WriteString("rightAnchor", right);
        if (op.MoveGroupId is { } group) writer.WriteNumber("moveGroupId", group);
        if (op.IsMoveSource is { } source) writer.WriteBoolean("isMoveSource", source);
        if (op.BodyFullRewriteGroupId is { } rewriteGroup)
            writer.WriteNumber("bodyFullRewriteGroupId", rewriteGroup);
        if (op.RequiresWholeParagraphReplace)
            writer.WriteBoolean("requiresWholeParagraphReplace", true);
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
                    WriteOp(writer, blockOp);
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
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ note diff

    private static void WriteNoteDiff(Utf8JsonWriter writer, IrNoteDiff diff)
    {
        writer.WriteStartObject();
        writer.WriteString("kind", diff.Kind.ToString());
        writer.WriteString("noteId", diff.NoteId);
        if (diff.LeftNoteId is { } leftId) writer.WriteString("leftNoteId", leftId);
        writer.WriteStartArray("ops");
        foreach (var op in diff.Ops)
            WriteOp(writer, op);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ header/footer diff

    private static void WriteHeaderFooterDiff(Utf8JsonWriter writer, IrHeaderFooterDiff diff)
    {
        writer.WriteStartObject();
        writer.WriteBoolean("isHeader", diff.IsHeader);
        writer.WriteString("kind", diff.Kind.ToString());
        writer.WriteNumber("sectionIndex", diff.SectionIndex);
        writer.WriteString("scope", diff.ScopeName);
        if (diff.LeftScopeName is { } leftScope) writer.WriteString("leftScope", leftScope);
        // Part URIs serialize as strings (package-relative, e.g. "/word/header1.xml") so the diff-as-data
        // consumer can identify the story's part and Read(Write(s)) stays exactly record-equal.
        if (diff.LeftPartUri is { } leftUri) writer.WriteString("leftPartUri", leftUri.ToString());
        if (diff.RightPartUri is { } rightUri) writer.WriteString("rightPartUri", rightUri.ToString());
        writer.WriteStartArray("ops");
        foreach (var op in diff.Ops)
            WriteOp(writer, op);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ table diff

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
                WriteOp(writer, blockOp);
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
    }

    private static void WriteTokenDiff(Utf8JsonWriter writer, IrTokenDiff diff)
    {
        writer.WriteStartObject();
        writer.WriteStartArray("ops");
        foreach (var tokenOp in diff.Ops)
        {
            // Compact 5-element array: [kindCode, leftStart, leftEnd, rightStart, rightEnd].
            writer.WriteStartArray();
            writer.WriteNumberValue(TokenKindCode(tokenOp.Kind));
            writer.WriteNumberValue(tokenOp.LeftStart);
            writer.WriteNumberValue(tokenOp.LeftEnd);
            writer.WriteNumberValue(tokenOp.RightStart);
            writer.WriteNumberValue(tokenOp.RightEnd);
            writer.WriteEndArray();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    // ------------------------------------------------------------------ read

    /// <summary>
    /// Parse JSON produced by <see cref="Write"/> back into an <see cref="IrEditScript"/>.
    /// </summary>
    /// <remarks>
    /// <para><b>Crash-on-garbage contract (by design).</b> This is an INTERNAL diagnostic format, not a
    /// public/untrusted wire protocol. <see cref="Read"/> assumes well-formed input emitted by
    /// <see cref="Write"/> and performs no tolerant/defensive parsing: malformed input THROWS rather than
    /// returning a partial or "best-effort" script. Specifically — non-JSON throws
    /// <see cref="JsonException"/>; a missing <c>operations</c> array, a missing <c>kind</c>, or a
    /// missing token-op array element throws <see cref="KeyNotFoundException"/>/
    /// <see cref="System.IndexOutOfRangeException"/>; an unrecognized <c>kind</c> enum name throws
    /// <see cref="ArgumentException"/>; an unrecognized token-op kind code throws
    /// <see cref="ArgumentOutOfRangeException"/> (see <see cref="TokenKindFromCode"/>); a wrong JSON value
    /// type (e.g. string where a number is expected) throws <see cref="InvalidOperationException"/>. We
    /// surface these loudly so a corrupt diagnostic artifact fails fast at the read site instead of
    /// silently degrading downstream. Callers that must tolerate arbitrary input should validate/guard
    /// upstream; do not add silent fallbacks here.</para>
    /// </remarks>
    public static IrEditScript Read(string json)
    {
        ArgumentNullException.ThrowIfNull(json);

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var ops = new List<IrEditOp>();
        foreach (var opElement in root.GetProperty("operations").EnumerateArray())
            ops.Add(ReadOp(opElement));

        IrNodeList<IrNoteDiff>? noteOps = null;
        if (root.TryGetProperty("noteOps", out var noteOpsElement))
        {
            var list = new List<IrNoteDiff>();
            foreach (var noteElement in noteOpsElement.EnumerateArray())
                list.Add(ReadNoteDiff(noteElement));
            noteOps = IrNodeList.From(list);
        }

        IrNodeList<IrHeaderFooterDiff>? headerFooterOps = null;
        if (root.TryGetProperty("headerFooterOps", out var hfOpsElement))
        {
            var list = new List<IrHeaderFooterDiff>();
            foreach (var hfElement in hfOpsElement.EnumerateArray())
                list.Add(ReadHeaderFooterDiff(hfElement));
            headerFooterOps = IrNodeList.From(list);
        }

        return new IrEditScript(IrNodeList.From(ops), noteOps, headerFooterOps);
    }

    private static IrHeaderFooterDiff ReadHeaderFooterDiff(JsonElement element)
    {
        bool isHeader = element.GetProperty("isHeader").GetBoolean();
        var kind = Enum.Parse<Docxodus.Ir.IrHeaderFooterKind>(element.GetProperty("kind").GetString()!);
        int sectionIndex = element.GetProperty("sectionIndex").GetInt32();
        string scope = element.GetProperty("scope").GetString()!;
        string? leftScope = element.TryGetProperty("leftScope", out var ls) ? ls.GetString() : null;
        Uri? leftPartUri = element.TryGetProperty("leftPartUri", out var lu)
            ? new Uri(lu.GetString()!, UriKind.RelativeOrAbsolute) : null;
        Uri? rightPartUri = element.TryGetProperty("rightPartUri", out var ru)
            ? new Uri(ru.GetString()!, UriKind.RelativeOrAbsolute) : null;
        var ops = new List<IrEditOp>();
        foreach (var opElement in element.GetProperty("ops").EnumerateArray())
            ops.Add(ReadOp(opElement));
        return new IrHeaderFooterDiff(isHeader, kind, sectionIndex, scope, leftScope,
            leftPartUri, rightPartUri, IrNodeList.From(ops));
    }

    private static IrNoteDiff ReadNoteDiff(JsonElement element)
    {
        var kind = Enum.Parse<IrNoteKind>(element.GetProperty("kind").GetString()!);
        string noteId = element.GetProperty("noteId").GetString()!;
        string? leftNoteId = element.TryGetProperty("leftNoteId", out var ln) ? ln.GetString() : null;
        var ops = new List<IrEditOp>();
        foreach (var opElement in element.GetProperty("ops").EnumerateArray())
            ops.Add(ReadOp(opElement));
        return new IrNoteDiff(kind, noteId, IrNodeList.From(ops), leftNoteId);
    }

    private static IrEditOp ReadOp(JsonElement element)
    {
        var kind = Enum.Parse<IrEditOpKind>(element.GetProperty("kind").GetString()!);
        string? leftAnchor = element.TryGetProperty("leftAnchor", out var l) ? l.GetString() : null;
        string? rightAnchor = element.TryGetProperty("rightAnchor", out var r) ? r.GetString() : null;
        int? moveGroupId = element.TryGetProperty("moveGroupId", out var g) ? g.GetInt32() : null;
        bool? isMoveSource = element.TryGetProperty("isMoveSource", out var s) ? s.GetBoolean() : null;
        int? bodyFullRewriteGroupId = element.TryGetProperty("bodyFullRewriteGroupId", out var brg)
            ? brg.GetInt32() : null;
        bool requiresWholeParagraphReplace = element.TryGetProperty("requiresWholeParagraphReplace", out var whp)
            && whp.GetBoolean();
        IrTokenDiff? tokenDiff = element.TryGetProperty("tokenDiff", out var t) ? ReadTokenDiff(t) : null;
        IrTableDiff? tableDiff = element.TryGetProperty("tableDiff", out var td) ? ReadTableDiff(td) : null;
        IrNodeList<IrTextboxDiff>? textboxDiffs = null;
        if (element.TryGetProperty("textboxDiffs", out var tbx))
        {
            var list = new List<IrTextboxDiff>();
            foreach (var tbxElement in tbx.EnumerateArray())
            {
                var inner = new List<IrEditOp>();
                foreach (var blockElement in tbxElement.GetProperty("ops").EnumerateArray())
                    inner.Add(ReadOp(blockElement));
                list.Add(new IrTextboxDiff(IrNodeList.From(inner)));
            }
            textboxDiffs = IrNodeList.From(list);
        }
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
            tableDiff, textboxDiffs, splitMergeAnchors, segmentDiffs, bodyFullRewriteGroupId,
            requiresWholeParagraphReplace);
    }

    private static IrTableDiff ReadTableDiff(JsonElement element)
    {
        var rowOps = new List<IrRowOp>();
        foreach (var rowElement in element.GetProperty("rowOps").EnumerateArray())
            rowOps.Add(ReadRowOp(rowElement));
        return new IrTableDiff(IrNodeList.From(rowOps));
    }

    private static IrRowOp ReadRowOp(JsonElement element)
    {
        var kind = Enum.Parse<IrRowOpKind>(element.GetProperty("kind").GetString()!);
        string? leftRowAnchor = element.TryGetProperty("leftRowAnchor", out var l) ? l.GetString() : null;
        string? rightRowAnchor = element.TryGetProperty("rightRowAnchor", out var r) ? r.GetString() : null;
        int? moveGroupId = element.TryGetProperty("moveGroupId", out var g) ? g.GetInt32() : null;
        bool? isMoveSource = element.TryGetProperty("isMoveSource", out var s) ? s.GetBoolean() : null;
        IrNodeList<IrCellOp>? cellOps = null;
        if (element.TryGetProperty("cellOps", out var c))
        {
            var list = new List<IrCellOp>();
            foreach (var cellElement in c.EnumerateArray())
                list.Add(ReadCellOp(cellElement));
            cellOps = IrNodeList.From(list);
        }
        return new IrRowOp(kind, leftRowAnchor, rightRowAnchor, cellOps, moveGroupId, isMoveSource);
    }

    private static IrCellOp ReadCellOp(JsonElement element)
    {
        string? leftCellAnchor = element.TryGetProperty("leftCellAnchor", out var l) ? l.GetString() : null;
        string? rightCellAnchor = element.TryGetProperty("rightCellAnchor", out var r) ? r.GetString() : null;
        IrNodeList<IrEditOp>? blockOps = null;
        if (element.TryGetProperty("blockOps", out var b))
        {
            var list = new List<IrEditOp>();
            foreach (var blockElement in b.EnumerateArray())
                list.Add(ReadOp(blockElement));
            blockOps = IrNodeList.From(list);
        }
        return new IrCellOp(leftCellAnchor, rightCellAnchor, blockOps);
    }

    private static IrTokenDiff ReadTokenDiff(JsonElement element)
    {
        var tokenOps = new List<IrTokenOp>();
        foreach (var arr in element.GetProperty("ops").EnumerateArray())
        {
            var kind = TokenKindFromCode(arr[0].GetInt32());
            tokenOps.Add(new IrTokenOp(kind, arr[1].GetInt32(), arr[2].GetInt32(), arr[3].GetInt32(), arr[4].GetInt32()));
        }
        return new IrTokenDiff(IrNodeList.From(tokenOps));
    }

    // ------------------------------------------------------------------ token-kind codes

    private static int TokenKindCode(IrTokenOpKind kind) => kind switch
    {
        IrTokenOpKind.Equal => 0,
        IrTokenOpKind.Insert => 1,
        IrTokenOpKind.Delete => 2,
        IrTokenOpKind.FormatChanged => 3,
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unknown token op kind."),
    };

    private static IrTokenOpKind TokenKindFromCode(int code) => code switch
    {
        0 => IrTokenOpKind.Equal,
        1 => IrTokenOpKind.Insert,
        2 => IrTokenOpKind.Delete,
        3 => IrTokenOpKind.FormatChanged,
        _ => throw new ArgumentOutOfRangeException(nameof(code), code, "Unknown token op kind code."),
    };
}
