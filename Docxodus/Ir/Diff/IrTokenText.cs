using System;
using System.Collections.Generic;

namespace Docxodus.Ir.Diff;

/// <summary>Character coordinates and exact source text for aligned token spans.</summary>
internal static class IrTokenText
{
    /// <summary>An Equal or FormatChanged span whose raw text is identical on both sides: it emits the revision's
    /// runs unwrapped.</summary>
    internal static bool KeepsTextUnchanged(
        IrTokenOp op, IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens) =>
        op.Kind is IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged &&
        string.Equals(RawSpanText(leftTokens, op.LeftStart, op.LeftEnd),
            RawSpanText(rightTokens, op.RightStart, op.RightEnd), StringComparison.Ordinal);

    /// <summary>Concatenate the RAW token text over a half-open token-index span (empty span ⇒ "").</summary>
    internal static string RawSpanText(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        if (start >= end)
            return string.Empty;
        var sb = new System.Text.StringBuilder();
        for (int i = start; i < end; i++)
            sb.Append(tokens[i].Text);
        return sb.ToString();
    }

    /// <summary>Right char span of a token op: empty (zero-width at the right anchor) for a Delete op.</summary>
    internal static (int Start, int End) RightSpanPositions(IReadOnlyList<IrDiffToken> tokens, IrTokenOp op)
    {
        if (op.RightStart >= op.RightEnd)
        {
            // Empty right span: position is the start-char of the right anchor token (or end-of-paragraph).
            int at = op.RightStart < tokens.Count ? tokens[op.RightStart].StartChar
                   : (tokens.Count > 0 ? tokens[^1].EndChar : 0);
            return (at, at);
        }
        return (tokens[op.RightStart].StartChar, tokens[op.RightEnd - 1].EndChar);
    }

    /// <summary>Left char span of a token op: empty (zero-width) for an Insert op.</summary>
    internal static (int Start, int End) LeftSpanPositions(IReadOnlyList<IrDiffToken> tokens, IrTokenOp op)
    {
        if (op.LeftStart >= op.LeftEnd)
        {
            int at = op.LeftStart < tokens.Count ? tokens[op.LeftStart].StartChar
                   : (tokens.Count > 0 ? tokens[^1].EndChar : 0);
            return (at, at);
        }
        return (tokens[op.LeftStart].StartChar, tokens[op.LeftEnd - 1].EndChar);
    }

}
