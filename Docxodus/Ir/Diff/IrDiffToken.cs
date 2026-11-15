#nullable enable

namespace Docxodus.Ir.Diff;

/// <summary>
/// The kind of a diff token produced by <see cref="IrDiffTokenizer"/>. The kinds mirror the §6.1
/// content-hash stream's granularity so token equality at a given kind corresponds to content-hash
/// equality at the same granularity.
/// </summary>
/// <remarks>
/// <para>There is deliberately no <c>FieldResultBoundary</c> kind: a field's cached result is
/// tokenized <b>transparently</b> (§6.1 / N9) — a PAGE field whose result reads "5" produces the
/// same tokens as a literal "5". Fields leave no marker in the token stream.</para>
/// <para>There is deliberately no <c>HyperlinkBoundary</c> kind: a hyperlink's child inlines are
/// tokenized transparently too, but each produced token's <see cref="IrDiffToken.MatchKey"/> carries
/// a <c>"lnk:&lt;target&gt;"</c> suffix (§6.1 framed-target hashing — linked text is never
/// content-equal to identical plain text, and a target change is a content change). The link affects
/// keys, not kinds.</para>
/// </remarks>
internal enum IrDiffTokenKind
{
    /// <summary>A word of non-separator text, or one Han, kana, or Hangul text element.</summary>
    Word,

    /// <summary>A single separator character (one token per separator char — §6.1 atom granularity).</summary>
    Separator,

    /// <summary>An <c>IrTab</c>.</summary>
    Tab,

    /// <summary>An <c>IrBreak</c> (line/page/column — distinguished by <see cref="IrDiffToken.MatchKey"/>).</summary>
    Break,

    /// <summary>An <c>IrNoteRef</c> (footnote/endnote). Id-less, consistent with §6.1.</summary>
    NoteRef,

    /// <summary>An <c>IrInlineImage</c> (keyed by image bytes hash).</summary>
    Image,

    /// <summary>An <c>IrOpaqueInline</c> (keyed by its canonical hash).</summary>
    Opaque,

    /// <summary>An <c>IrTextbox</c> — one placeholder token keyed by its rolled inner-block hashes.</summary>
    Textbox,
}

/// <summary>
/// One diff-time token over an IR paragraph. <see cref="Text"/> is the raw source text (empty for
/// non-text kinds); <see cref="MatchKey"/> is the normalized equality key. <see cref="StartChar"/>/
/// <see cref="EndChar"/> are half-open char offsets in the same coordinate space as
/// <c>IrCommentTarget</c> and <c>DocxSession.ApplyFormat</c> — counting only emitted <c>IrTextRun</c>
/// characters (including a field's cached-result text), with tab/break/note-ref/image/opaque/textbox
/// each contributing 0. <see cref="Format"/> is the governing run format (null for non-run kinds).
/// Deviation from the M2.1 plan sketch (deliberate): the sketched <c>IrHash? AtomHash</c> field is
/// omitted — atomic content identity is already encoded in <see cref="MatchKey"/> (hash hex inside
/// the key). Add it back additively if M2.2 wants structured hashes alongside keys.
/// </summary>
internal sealed record IrDiffToken(
    IrDiffTokenKind Kind,
    string Text,
    string MatchKey,
    int StartChar,
    int EndChar,
    IrRunFormat? Format)
{
    /// <summary>True for a token that stands for an object rather than for text: a note reference, image, opaque
    /// inline or text box. Every other kind (a word, separator, tab or break) is a character of the paragraph's
    /// text. An object's key is coarse (every footnote reference keys alike), so it is paired in full context,
    /// and one that splits a word without a separator makes a different word.</summary>
    public bool IsObject => Kind is IrDiffTokenKind.NoteRef or IrDiffTokenKind.Image
        or IrDiffTokenKind.Opaque or IrDiffTokenKind.Textbox;

    /// <summary>True for whitespace between words: a whitespace separator, a tab or a break of any kind. A
    /// comparison never anchors on one, and keeps one that both sides share at the edge of a changed span; the
    /// reference compare treats a tab or break exactly as it treats a space (issue #1759).</summary>
    public bool IsConnective => Kind is IrDiffTokenKind.Tab or IrDiffTokenKind.Break
        || (Kind == IrDiffTokenKind.Separator && string.IsNullOrWhiteSpace(Text));
}
