#nullable enable

using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Tokenizes an <see cref="IrParagraph"/> into the diff engine's word/separator/atomic token stream
/// (M2.1). The walk mirrors the §6.1 content-hash byte stream so token equality at a given kind
/// corresponds to content-hash equality at the same granularity, and it mirrors the reader's
/// comment-target char counter (Docxodus/Ir/IrReader.cs <c>CommentTracker</c>, documented on
/// <c>IrCommentTarget</c>) so token <c>StartChar</c>/<c>EndChar</c> live in the same coordinate space
/// as comment targets and <c>DocxSession.ApplyFormat</c>.
/// </summary>
/// <remarks>
/// <para><b>Shared coordinate-space contract.</b> The char offset advances by the length of every
/// emitted <c>IrTextRun</c>'s text — INCLUDING text inside a field's cached result (the reader emits
/// those as ordinary <c>IrTextRun</c>s and advances by their length; we recurse into
/// <c>CachedResult</c> and do the same). Tabs, breaks, note refs, images, opaque inlines, and the
/// textbox placeholder each advance the counter by 0. This is exactly the rule <c>IrCommentTarget</c>
/// documents, so a comment range and a token computed over the same paragraph agree on offsets.</para>
/// <para><b>The tokenizer needs no provenance.</b> It reads only the built IR node tree
/// (<see cref="IrParagraph.Inlines"/> and nested inline lists), never <c>Source</c>, so it works on
/// an IR read with <c>RetainSources=false</c>.</para>
/// </remarks>
internal static class IrDiffTokenizer
{
    // Atomic-kind MatchKeys are prefixed with U+0001. The non-collision guarantee rests on XML 1.0:
    // U+0001 is an illegal character in XML text content, so no normalized word/separator key —
    // always derived from w:t text — can ever begin with it. A literal word "tab" yields MatchKey
    // "tab"; an IrTab yields U+0001 + "tab". (Same justification as the content-hash stream's
    // sentinel framing, spec §6.1.)
    private const char AtomicSentinel = '\u0001';

    public static IReadOnlyList<IrDiffToken> Tokenize(IrParagraph paragraph, IrDiffSettings settings)
    {
        var tokens = new List<IrDiffToken>();
        int charOffset = 0;
        WalkInlines(paragraph.Inlines, settings, linkSuffix: null, tokens, ref charOffset);
        return tokens;
    }

    /// <summary>
    /// Walk an inline list in document order, appending tokens and advancing <paramref name="charOffset"/>.
    /// <paramref name="linkSuffix"/> is the accumulated hyperlink-target suffix to append to every
    /// word/separator MatchKey produced within this scope (composed in document order — an outer link
    /// applied before an inner link).
    /// </summary>
    private static void WalkInlines(
        IReadOnlyList<IrInline> inlines, IrDiffSettings settings, string? linkSuffix,
        List<IrDiffToken> tokens, ref int charOffset)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun run:
                    EmitTextRun(run.Text, run.Format, settings, linkSuffix, tokens, ref charOffset);
                    break;

                case IrFieldRun field:
                    // §6.1 / N9: the cached result is tokenized transparently — its IrTextRuns are
                    // indistinguishable from literal text, and (like the reader) their chars advance
                    // the offset. The instruction is never tokenized.
                    WalkInlines(field.CachedResult, settings, linkSuffix, tokens, ref charOffset);
                    break;

                case IrHyperlink link:
                    // §6.1 framed target: recurse transparently, but every produced token's MatchKey
                    // gets a "lnk:<target>" suffix so linked text ≠ plain text and a target change is
                    // a content change. Suffixes compose in document order (outer applied first).
                    var target = link.Target ?? link.InternalTarget?.ToString() ?? "";
                    var composed = linkSuffix is null ? LinkSuffix(target) : linkSuffix + LinkSuffix(target);
                    WalkInlines(link.Inlines, settings, composed, tokens, ref charOffset);
                    break;

                case IrTab tab:
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.Tab, "", AtomicKey("tab"), charOffset, charOffset, tab.Format));
                    break;

                case IrBreak brk:
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.Break, "", AtomicKey("brk:" + brk.Kind), charOffset, charOffset, null));
                    break;

                case IrNoteRef note:
                    // Id-less (kind only), consistent with §6.1 (renumbering must not flip equality).
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.NoteRef, "", AtomicKey(note.Kind == IrNoteKind.Footnote ? "fn" : "en"),
                        charOffset, charOffset, null));
                    break;

                case IrInlineImage image:
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.Image, "", AtomicKey("img:" + image.ImageBytesHash.ToHex()),
                        charOffset, charOffset, null));
                    break;

                case IrOpaqueInline opaque:
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.Opaque, "", AtomicKey("opq:" + opaque.CanonicalHash.ToHex()),
                        charOffset, charOffset, null));
                    break;

                case IrTextbox textbox:
                    // ONE placeholder token; its inner blocks are aligned as blocks separately. The key
                    // rolls the inner-block ContentHashes in document order (mirrors §6.1's textbox
                    // sentinel framing), so two textboxes with identical inner text share a key.
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.Textbox, "", AtomicKey("tbx:" + TextboxRollKey(textbox)),
                        charOffset, charOffset, null));
                    break;

                default:
                    // No other inline kinds exist; future kinds default to a zero-width opaque-style
                    // token rather than being silently dropped.
                    tokens.Add(new IrDiffToken(
                        IrDiffTokenKind.Opaque, "", AtomicKey("unk:" + inline.GetType().Name),
                        charOffset, charOffset, null));
                    break;
            }
        }
    }

    /// <summary>
    /// Split a text run on <see cref="IrDiffSettings.WordSeparators"/> into alternating Word and
    /// Separator tokens (one Separator token per separator char). Advances <paramref name="charOffset"/>
    /// by the run's raw length.
    /// </summary>
    private static void EmitTextRun(
        string text, IrRunFormat format, IrDiffSettings settings, string? linkSuffix,
        List<IrDiffToken> tokens, ref int charOffset)
    {
        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (settings.WordSeparators.Contains(c))
            {
                int start = charOffset + i;
                string raw = c.ToString();
                tokens.Add(new IrDiffToken(
                    IrDiffTokenKind.Separator, raw, ApplyLink(NormalizeWord(raw, settings), linkSuffix),
                    start, start + 1, format));
                i++;
            }
            else
            {
                int wordStart = i;
                while (i < text.Length && !settings.WordSeparators.Contains(text[i]))
                    i++;
                string raw = text.Substring(wordStart, i - wordStart);
                int start = charOffset + wordStart;
                tokens.Add(new IrDiffToken(
                    IrDiffTokenKind.Word, raw, ApplyLink(NormalizeWord(raw, settings), linkSuffix),
                    start, start + raw.Length, format));
            }
        }
        charOffset += text.Length;
    }

    /// <summary>
    /// Normalize a word/separator's text into its match key: case fold (per
    /// <see cref="IrDiffSettings.CaseInsensitive"/> + culture, ordinal/invariant when culture is null),
    /// and fold NBSP (U+00A0) → space when conflating. U+2011 (non-breaking hyphen) is left distinct.
    /// </summary>
    private static string NormalizeWord(string raw, IrDiffSettings settings)
    {
        string s = raw;
        if (settings.ConflateBreakingAndNonbreakingSpaces && s.IndexOf('\u00A0') >= 0)
            s = s.Replace('\u00A0', ' ');
        if (settings.CaseInsensitive)
            s = settings.Culture is { } culture ? s.ToLower(culture) : s.ToLowerInvariant();
        return s;
    }

    private static string ApplyLink(string key, string? linkSuffix) =>
        linkSuffix is null ? key : key + linkSuffix;

    private static string LinkSuffix(string target) => "lnk:" + target;

    private static string AtomicKey(string body) => AtomicSentinel + body;

    /// <summary>Roll a textbox's inner-block ContentHashes (document order) into one hex key.</summary>
    private static string TextboxRollKey(IrTextbox textbox)
    {
        var sb = new StringBuilder();
        foreach (var block in textbox.Blocks)
        {
            sb.Append(block.ContentHash.ToHex());
            sb.Append('.');
        }
        return sb.ToString();
    }
}
