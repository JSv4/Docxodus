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
    /// <summary>A maximal run of non-separator characters from an <c>IrTextRun</c>.</summary>
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
    IrRunFormat? Format);
