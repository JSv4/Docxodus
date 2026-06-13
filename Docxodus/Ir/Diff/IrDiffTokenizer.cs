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
/// <para><b>Intra-word atom interruption (§6.1-adjacent, M2.5 Task 1).</b> A word's diff identity is its
/// atom STRUCTURE, not merely its visible characters: a word split by a zero-width content atom (a note
/// ref / image / opaque / textbox) with NO separator on either side is a different word from its
/// contiguous form (a note reference relocated INTO the middle of <c>Video</c> — <c>Vi</c>⟨ref⟩<c>deo</c>
/// — is a real edit even though the letters still spell <c>Video</c>). <see cref="InterruptionPostPass"/>
/// frames the two flanking words' match keys with the interrupting atoms' keys to express this, while a
/// ref BETWEEN words (separator-adjacent) and the ref token's OWN key are left exactly as the §6.1 stream
/// has them — so the change is confined to the rare intra-word case and never disturbs the common one.</para>
/// </remarks>
internal static class IrDiffTokenizer
{
    // Atomic-kind MatchKeys are prefixed with U+0001. The non-collision guarantee rests on XML 1.0:
    // U+0001 is an illegal character in XML text content, so no normalized word/separator key —
    // always derived from w:t text — can ever begin with it. A literal word "tab" yields MatchKey
    // "tab"; an IrTab yields U+0001 + "tab". (Same justification as the content-hash stream's
    // sentinel framing, spec §6.1.)
    private const char AtomicSentinel = '\u0001';

    // Word tokens flanking an intra-word interruption (a zero-width CONTENT atom that splits a word
    // with no separator on either side) get this sentinel-framed marker appended to their MatchKey, so
    // `Vi⟨ref⟩deo` is NOT word-equal to a contiguous `Video`. The marker carries the interrupting atoms'
    // OWN keys (in document order) so the equality break is specific to what interrupted the word: a ref
    // moving inside `Video` reads as a word change, but an image interrupting it reads differently again.
    // Sentinel-framed (U+0001) for the same XML-illegality non-collision guarantee as atomic keys, and
    // distinct body (`iw:`) so it can never alias an atomic key. See InterruptionPostPass.
    private const string InterruptionMarkerPrefix = "iw:";

    public static IReadOnlyList<IrDiffToken> Tokenize(IrParagraph paragraph, IrDiffSettings settings)
    {
        var tokens = new List<IrDiffToken>();
        int charOffset = 0;
        WalkInlines(paragraph.Inlines, settings, linkSuffix: null, tokens, ref charOffset);
        InterruptionPostPass(tokens);
        return tokens;
    }

    /// <summary>
    /// Rewrite the MatchKeys of word tokens that flank an <b>intra-word interruption</b>: a maximal run
    /// of one or more zero-width CONTENT atoms (NoteRef / Image / Opaque / Textbox) sitting between two
    /// Word tokens with NO separator on either side (the words touch the atoms at their char offsets).
    /// In that configuration the word's atom structure genuinely changed — a note reference relocated INTO
    /// the middle of <c>Video</c> (<c>Vi</c>[ref]<c>deo</c>) is a real structural edit the contiguous
    /// <c>Video</c> does not share — so the two flanking words must NOT be word-equal to their contiguous
    /// form. Each flanking word gets <see cref="InterruptionMarkerPrefix"/> + the interrupting atoms' keys
    /// appended to its MatchKey.
    /// </summary>
    /// <remarks>
    /// <para><b>Zero corpus blast radius outside intra-word cases.</b> The OVERWHELMINGLY common placement
    /// — a ref BETWEEN words (separator-adjacent: <c>Video </c>[ref] or [ref]<c> provides</c>) — is NOT an
    /// interruption: a Separator token sits on at least one side, so neither flanking word is rewritten and
    /// the keys are byte-identical to the pre-pass output. Only the rare mid-word placement is affected.
    /// The interrupting atom tokens themselves are NOT rewritten (their keys stay position-independent, so
    /// a between-word ref that did NOT move still matches across the pair).</para>
    /// <para><b>Engine truth, not render policy.</b> This runs unconditionally in <see cref="Tokenize"/>,
    /// so it is identical under every <see cref="IrDiffSettings"/> — Fine and WmlComparerCompatible see the
    /// same token stream. It mirrors the §6.1 content-hash stream's word-structure granularity: a word
    /// split by an inline atom is a different word.</para>
    /// <para><b>Determinism.</b> A single forward linear scan keyed on token kinds and char offsets; the
    /// marker body is the interrupting atoms' keys in document order. Pure function of the token list.</para>
    /// </remarks>
    private static void InterruptionPostPass(List<IrDiffToken> tokens)
    {
        for (int i = 0; i < tokens.Count; i++)
        {
            // Anchor on a Word immediately followed by a tightly-adjacent zero-width content atom.
            if (tokens[i].Kind != IrDiffTokenKind.Word)
                continue;
            int leftWord = i;

            int j = i + 1;
            int atomCount = 0;
            while (j < tokens.Count &&
                   IsZeroWidthContentAtom(tokens[j].Kind) &&
                   tokens[j].StartChar == tokens[j - 1].EndChar)
            {
                atomCount++;
                j++;
            }

            // Need ≥1 interrupting atom AND a tightly-adjacent Word resuming after it (no separator).
            if (atomCount == 0 ||
                j >= tokens.Count ||
                tokens[j].Kind != IrDiffTokenKind.Word ||
                tokens[j].StartChar != tokens[j - 1].EndChar)
            {
                continue;
            }
            int rightWord = j;

            // Build the interruption marker from the interrupting atoms' keys (document order). Lead with
            // the AtomicSentinel so the word↔marker boundary is unambiguous: a normalized word — always
            // derived from w:t text — can never contain U+0001, so `Vi` + marker can never alias a literal
            // word that merely happened to end in the marker body.
            var sb = new StringBuilder();
            sb.Append(AtomicSentinel).Append(InterruptionMarkerPrefix);
            for (int a = leftWord + 1; a < rightWord; a++)
                sb.Append(tokens[a].MatchKey);
            string marker = sb.ToString();

            tokens[leftWord] = AppendMarker(tokens[leftWord], marker);
            tokens[rightWord] = AppendMarker(tokens[rightWord], marker);

            // Continue scanning from the resuming word: it may itself be the left side of a FURTHER
            // interruption (`Vi[ref]de[ref]o`). Re-anchoring on it (i = rightWord) handles the chain.
            i = rightWord - 1;
        }
    }

    private static bool IsZeroWidthContentAtom(IrDiffTokenKind kind) => kind is
        IrDiffTokenKind.NoteRef or IrDiffTokenKind.Image or
        IrDiffTokenKind.Opaque or IrDiffTokenKind.Textbox;

    private static IrDiffToken AppendMarker(IrDiffToken token, string marker) =>
        token with { MatchKey = token.MatchKey + marker };

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
    /// <remarks>
    /// <para><b>NBSP conflation happens at SPLIT time, not just in the match key.</b> When
    /// <see cref="IrDiffSettings.ConflateBreakingAndNonbreakingSpaces"/> is true, U+00A0 (non-breaking
    /// space) is treated as a word separator EQUIVALENT to an ordinary space — even though it is not a
    /// member of <see cref="IrDiffSettings.WordSeparators"/>. This matches WmlComparer, which conflates
    /// NBSP→space BEFORE its word split: in <c>l'article&#160;1</c> (NBSP) vs <c>l'article 1</c> (space)
    /// the only change is the space character, so both must tokenize to the SAME word/separator boundaries
    /// — <c>{l'article}{sep}{1}</c> — and the resulting Separator token's match key normalizes to <c>" "</c>.
    /// Folding only in the post-split match key (as a NBSP-inside-a-word) would leave the two sides with
    /// DIFFERENT token boundaries (one word <c>l'article&#160;1</c> vs three tokens), producing a spurious
    /// diff. The Separator token's raw <see cref="IrDiffToken.Text"/> preserves the original U+00A0 so the
    /// rendered output is byte-faithful. When the setting is false, U+00A0 is an ordinary word character
    /// (its previous behavior).</para>
    /// </remarks>
    private static void EmitTextRun(
        string text, IrRunFormat format, IrDiffSettings settings, string? linkSuffix,
        List<IrDiffToken> tokens, ref int charOffset)
    {
        bool nbspIsSeparator = settings.ConflateBreakingAndNonbreakingSpaces;

        bool IsSeparator(char c) =>
            settings.WordSeparators.Contains(c) || (nbspIsSeparator && c == '\u00A0');

        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (IsSeparator(c))
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
                while (i < text.Length && !IsSeparator(text[i]))
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
