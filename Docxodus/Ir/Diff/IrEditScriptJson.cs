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
/// <c>isMoveSource</c>/<c>tokenDiff</c> that are present (absent fields are simply omitted, matching the
/// record's nullability). A <c>tokenDiff</c> is <c>{"ops":[ [kind,ls,le,rs,re], … ]}</c> — each token
/// op a COMPACT 5-element array: an integer kind code (0=Equal,1=Insert,2=Delete,3=FormatChanged) plus
/// the four half-open span bounds. The compact array keeps a large corpus script terse while staying
/// fully self-describing for the reader.</para>
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
        if (op.TokenDiff is { } diff)
        {
            writer.WritePropertyName("tokenDiff");
            WriteTokenDiff(writer, diff);
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
        return new IrEditScript(IrNodeList.From(ops));
    }

    private static IrEditOp ReadOp(JsonElement element)
    {
        var kind = Enum.Parse<IrEditOpKind>(element.GetProperty("kind").GetString()!);
        string? leftAnchor = element.TryGetProperty("leftAnchor", out var l) ? l.GetString() : null;
        string? rightAnchor = element.TryGetProperty("rightAnchor", out var r) ? r.GetString() : null;
        int? moveGroupId = element.TryGetProperty("moveGroupId", out var g) ? g.GetInt32() : null;
        bool? isMoveSource = element.TryGetProperty("isMoveSource", out var s) ? s.GetBoolean() : null;
        IrTokenDiff? tokenDiff = element.TryGetProperty("tokenDiff", out var t) ? ReadTokenDiff(t) : null;
        return new IrEditOp(kind, leftAnchor, rightAnchor, tokenDiff, moveGroupId, isMoveSource);
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
