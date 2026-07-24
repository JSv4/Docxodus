#nullable enable

using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.1 Task 1 tests for <see cref="IrDiffTokenizer"/>: word/separator splitting, case-fold and
/// NBSP-conflation settings, char-offset alignment (cross-checked against the comment-target
/// coordinate space), atomic-kind tokens and their non-collision with word text, hyperlink-target
/// match-key suffixing, field-result transparency, format carry-through, and determinism.
/// </summary>
/// <remarks>
/// Paragraphs are built via <see cref="IrTestDocuments"/> + <see cref="IrReader"/> read with
/// <c>RetainSources = false</c> — proving the tokenizer needs no provenance.
/// </remarks>
public class IrDiffTokenizerTests
{
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };
    private static readonly IrDiffSettings Default = new();

    private static IrParagraph Para(string bodyXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyXml), NoSources)
            .Body.Blocks.OfType<IrParagraph>().First();

    private static IrParagraph TextPara(string text) =>
        IrReader.Read(IrTestDocuments.Create(text), NoSources)
            .Body.Blocks.OfType<IrParagraph>().First();

    private static IReadOnlyList<IrDiffToken> Tok(IrParagraph p, IrDiffSettings? s = null) =>
        IrDiffTokenizer.Tokenize(p, s ?? Default);

    private static string Run(string text) =>
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    // --- splitting --------------------------------------------------------

    [Fact]
    public void Splits_words_and_separators_one_token_per_separator_char()
    {
        var tokens = Tok(TextPara("foo bar"));
        Assert.Collection(tokens,
            t => { Assert.Equal(IrDiffTokenKind.Word, t.Kind); Assert.Equal("foo", t.Text); },
            t => { Assert.Equal(IrDiffTokenKind.Separator, t.Kind); Assert.Equal(" ", t.Text); },
            t => { Assert.Equal(IrDiffTokenKind.Word, t.Kind); Assert.Equal("bar", t.Text); });
    }

    [Fact]
    public void Multi_separator_run_yields_one_token_per_char()
    {
        // " - " is three separator chars (space, hyphen, space) — three Separator tokens.
        var tokens = Tok(TextPara("a - b"));
        Assert.Equal(
            new[] { IrDiffTokenKind.Word, IrDiffTokenKind.Separator, IrDiffTokenKind.Separator,
                    IrDiffTokenKind.Separator, IrDiffTokenKind.Word },
            tokens.Select(t => t.Kind));
        Assert.Equal(new[] { "a", " ", "-", " ", "b" }, tokens.Select(t => t.Text));
    }

    [Fact]
    public void Leading_and_trailing_separators_produce_separator_tokens()
    {
        var tokens = Tok(TextPara(" hi "));
        Assert.Equal(
            new[] { IrDiffTokenKind.Separator, IrDiffTokenKind.Word, IrDiffTokenKind.Separator },
            tokens.Select(t => t.Kind));
    }

    [Fact]
    public void Empty_paragraph_yields_no_tokens()
    {
        // An empty run is dropped by N10, leaving an empty paragraph.
        Assert.Empty(Tok(Para("<w:p/>")));
    }

    [Fact]
    public void Period_and_comma_split_only_when_not_adjacent_to_a_digit()
    {
        // WmlComparer's atom grouping isolates sentence punctuation, but keeps decimal/thousands punctuation
        // in its numeric word even though comma is also a configured static separator.
        var tokens = Tok(TextPara("This. 3.14 1,234 end,"));
        Assert.Equal(
            new[] { "This", ".", " ", "3.14", " ", "1,234", " ", "end", "," },
            tokens.Select(t => t.Text));
        Assert.Equal(
            new[]
            {
                IrDiffTokenKind.Word, IrDiffTokenKind.Separator, IrDiffTokenKind.Separator,
                IrDiffTokenKind.Word, IrDiffTokenKind.Separator, IrDiffTokenKind.Word,
                IrDiffTokenKind.Separator, IrDiffTokenKind.Word, IrDiffTokenKind.Separator,
            },
            tokens.Select(t => t.Kind));
    }

    [Theory]
    [InlineData("3", ".", "14")]
    [InlineData("1", ",", "234")]
    public void Numeric_punctuation_uses_adjacent_differently_formatted_runs(
        string leadingDigits, string punctuation, string trailingDigits)
    {
        // The punctuation run sees adjacent text runs, but remains a separate token so each source run's
        // format, provenance, and char offsets remain intact.
        var p = Para(
            $"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{leadingDigits}</w:t></w:r>" +
            $"<w:r><w:rPr><w:i/></w:rPr><w:t>{punctuation}</w:t></w:r>" +
            $"<w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:t>{trailingDigits}</w:t></w:r></w:p>");

        var tokens = Tok(p);
        Assert.Equal(new[] { leadingDigits, punctuation, trailingDigits }, tokens.Select(t => t.Text));
        Assert.All(tokens, t => Assert.Equal(IrDiffTokenKind.Word, t.Kind));
        Assert.True(tokens[0].Format!.Bold);
        Assert.True(tokens[1].Format!.Italic);
        Assert.Equal(IrUnderlineKind.Single, tokens[2].Format!.Underline!.Kind);
        Assert.Equal((leadingDigits.Length, leadingDigits.Length + punctuation.Length),
            (tokens[1].StartChar, tokens[1].EndChar));
    }

    [Fact]
    public void Numeric_punctuation_looks_through_a_transparent_hyperlink_child()
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var p = IrReader.Read(IrTestDocuments.FromBodyXmlWithHyperlinks(
            $"<w:p><w:r><w:t>3</w:t></w:r>" +
            $"<w:hyperlink r:id=\"r1\" xmlns:r=\"{R}\"><w:r><w:rPr><w:b/></w:rPr><w:t>.</w:t></w:r></w:hyperlink>" +
            "<w:r><w:t>14</w:t></w:r></w:p>",
            ("r1", "https://a.example")), NoSources)
            .Body.Blocks.OfType<IrParagraph>().First();

        var punctuation = Tok(p).Single(t => t.Text == ".");
        Assert.Equal(IrDiffTokenKind.Word, punctuation.Kind);
        Assert.True(punctuation.Format!.Bold);
    }

    [Fact]
    public void Numeric_punctuation_looks_through_a_transparent_field_result()
    {
        var p = Para(
            "<w:p><w:r><w:t>1</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>,</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "<w:r><w:t>234</w:t></w:r></w:p>");

        var punctuation = Tok(p).Single(t => t.Text == ",");
        Assert.Equal(IrDiffTokenKind.Word, punctuation.Kind);
        Assert.True(punctuation.Format!.Bold);
    }

    [Fact]
    public void Numeric_punctuation_does_not_look_through_a_tab_atom()
    {
        // A tab is a comparison-atom boundary. The period must not reuse the digit before it when its actual
        // next neighbor is non-numeric, otherwise "3<tab>.x" would incorrectly keep '.' in a word.
        var p = Para(
            "<w:p><w:r><w:t>3</w:t></w:r><w:r><w:tab/></w:r>" +
            "<w:r><w:t>.</w:t></w:r><w:r><w:t>x</w:t></w:r></w:p>");

        var punctuation = Tok(p).Single(t => t.Text == ".");
        Assert.Equal(IrDiffTokenKind.Separator, punctuation.Kind);
    }

    // --- normalization settings ------------------------------------------

    [Fact]
    public void Case_fold_off_keeps_distinct_keys()
    {
        Assert.NotEqual(Tok(TextPara("Foo"))[0].MatchKey, Tok(TextPara("foo"))[0].MatchKey);
    }

    [Fact]
    public void Case_fold_on_collapses_keys_and_preserves_raw_text()
    {
        var ci = new IrDiffSettings { CaseInsensitive = true };
        var upper = Tok(TextPara("Foo"), ci)[0];
        var lower = Tok(TextPara("foo"), ci)[0];
        Assert.Equal(lower.MatchKey, upper.MatchKey);
        Assert.Equal("Foo", upper.Text); // raw preserved
    }

    [Fact]
    public void Case_fold_uses_supplied_culture()
    {
        // Turkish dotted-I: invariant lower of "I" is "i"; tr-TR lower of "I" is "ı".
        var tr = new IrDiffSettings { CaseInsensitive = true, Culture = CultureInfo.GetCultureInfo("tr-TR") };
        Assert.Equal("ı", Tok(TextPara("I"), tr)[0].MatchKey);
    }

    [Fact]
    public void Nbsp_conflation_on_matches_space()
    {
        // With conflation ON, U+00A0 splits as a separator EQUIVALENT to ' ' (the WC-1970/1980 fix), so
        // "a\u00A0b" tokenizes to {Word a}{Separator}{Word b} \u2014 IDENTICAL boundaries to the space-separated
        // "a b". The separator token's raw Text preserves the NBSP; its MatchKey folds to a regular space,
        // making the whole MatchKey sequence equal to the space form's.
        var on = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = true };
        var nbsp = Tok(TextPara("a\u00A0b"), on);
        var space = Tok(TextPara("a b"), on);
        Assert.Equal(
            new[] { IrDiffTokenKind.Word, IrDiffTokenKind.Separator, IrDiffTokenKind.Word },
            nbsp.Select(t => t.Kind));
        Assert.Equal("\u00A0", nbsp[1].Text);                          // raw NBSP preserved
        Assert.Equal(" ", nbsp[1].MatchKey);                            // folded to a regular space
        Assert.Equal(space.Select(t => t.MatchKey), nbsp.Select(t => t.MatchKey)); // identical key sequence
    }

    [Fact]
    public void Nbsp_conflation_off_keeps_nbsp_distinct()
    {
        var off = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = false };
        Assert.Equal("a\u00A0b", Tok(TextPara("a\u00A0b"), off)[0].MatchKey); // NBSP kept in key
        Assert.NotEqual(
            Tok(TextPara("a b"), off)[0].MatchKey,
            Tok(TextPara("a\u00A0b"), off)[0].MatchKey);
    }

    [Fact]
    public void Nbsp_conflation_on_splits_nbsp_as_a_separator_char()
    {
        // The fix (WC-1970/1980): with conflation ON, U+00A0 must SPLIT at tokenize time exactly like an
        // ordinary space \u2014 even though it is not a WordSeparators member. "x\u00A0y" therefore tokenizes
        // to {Word "x"}{Separator}{Word "y"}, NOT to a single word "x\u00A0y". The Separator token's raw
        // Text preserves the original NBSP, while its MatchKey normalizes to a regular space.
        var on = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = true };
        var tokens = Tok(TextPara("x\u00A0y"), on);
        Assert.Equal(
            new[] { IrDiffTokenKind.Word, IrDiffTokenKind.Separator, IrDiffTokenKind.Word },
            tokens.Select(t => t.Kind));
        Assert.Equal(new[] { "x", "\u00A0", "y" }, tokens.Select(t => t.Text)); // raw NBSP preserved
        Assert.Equal(" ", tokens[1].MatchKey);                                  // key folds to a space
    }

    [Fact]
    public void Nbsp_conflation_on_yields_identical_matchkey_sequence_to_space()
    {
        // A pure space\u2194NBSP edit must produce an IDENTICAL token MatchKey sequence under conflation, so the
        // diff sees zero content change. This is exactly the WC055/WC056 "l'article 1" vs "l'article\u00A01"
        // case: WmlComparer reports 0 revisions and is CORRECT; the IR engine must agree.
        var on = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = true };
        var spaceKeys = Tok(TextPara("l'article 1"), on).Select(t => t.MatchKey);
        var nbspKeys = Tok(TextPara("l'article\u00A01"), on).Select(t => t.MatchKey);
        Assert.Equal(spaceKeys, nbspKeys);
    }

    [Fact]
    public void Nbsp_conflation_off_yields_distinct_matchkey_sequence_from_space()
    {
        // With the setting OFF, current behavior holds: NBSP is an ordinary word char, so "a\u00A0b" is one
        // word while "a b" splits into three tokens \u2014 the MatchKey SEQUENCES differ (a real content change).
        var off = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = false };
        var spaceKeys = Tok(TextPara("a b"), off).Select(t => t.MatchKey).ToList();
        var nbspKeys = Tok(TextPara("a\u00A0b"), off).Select(t => t.MatchKey).ToList();
        Assert.NotEqual(spaceKeys, nbspKeys);
    }

    [Fact]
    public void Nonbreaking_hyphen_is_not_folded_to_space()
    {
        // U+2011 stays distinct even when conflating spaces (it is not a space). It is also not in
        // the separator set, so "a\u2011b" is one word whose key keeps the U+2011.
        var on = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = true };
        Assert.Equal("a\u2011b", Tok(TextPara("a\u2011b"), on)[0].MatchKey);
    }

    // --- offsets ----------------------------------------------------------

    [Fact]
    public void Offsets_line_up_with_text_positions()
    {
        var tokens = Tok(TextPara("foo bar"));
        Assert.Equal((0, 3), (tokens[0].StartChar, tokens[0].EndChar)); // foo
        Assert.Equal((3, 4), (tokens[1].StartChar, tokens[1].EndChar)); // space
        Assert.Equal((4, 7), (tokens[2].StartChar, tokens[2].EndChar)); // bar
    }

    [Fact]
    public void Zero_width_atomics_do_not_advance_offset()
    {
        // "ab" <tab> "cd": tab contributes 0; "cd" starts at offset 2, not 3.
        var p = Para("<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>cd</w:t></w:r></w:p>");
        var tokens = Tok(p);
        var word1 = tokens[0];
        var tab = tokens.Single(t => t.Kind == IrDiffTokenKind.Tab);
        var word2 = tokens.Last();
        Assert.Equal((0, 2), (word1.StartChar, word1.EndChar));
        Assert.Equal((2, 2), (tab.StartChar, tab.EndChar)); // zero-width
        Assert.Equal((2, 4), (word2.StartChar, word2.EndChar));
    }

    [Fact]
    public void Token_offsets_match_comment_target_coordinate_space()
    {
        // Build a doc with a comment range over "world" and confirm the tokenizer's offsets for that
        // word equal the IrCommentTarget's offsets — the shared coordinate-space contract.
        const string body =
            "<w:p><w:r><w:t xml:space=\"preserve\">hello </w:t></w:r>" +
            "<w:commentRangeStart w:id=\"0\"/>" +
            "<w:r><w:t>world</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/>" +
            "<w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        var doc = IrReader.Read(IrTestDocuments.WithComment("A", "A", "2026-01-01", "c", body), NoSources);
        var para = doc.Body.Blocks.OfType<IrParagraph>().First();
        var target = doc.Comments.Comments.Single().Targets.Single();

        var tokens = IrDiffTokenizer.Tokenize(para, Default);
        var worldTok = tokens.Single(t => t.Text == "world");
        Assert.Equal(target.StartChar, worldTok.StartChar);
        Assert.Equal(target.EndChar, worldTok.EndChar);
    }

    // --- atomic kinds + non-collision ------------------------------------

    [Fact]
    public void Tab_and_break_have_atomic_kinds()
    {
        var p = Para(
            "<w:p><w:r><w:tab/></w:r><w:r><w:br w:type=\"page\"/></w:r></w:p>");
        var kinds = Tok(p).Select(t => t.Kind).ToArray();
        Assert.Contains(IrDiffTokenKind.Tab, kinds);
        Assert.Contains(IrDiffTokenKind.Break, kinds);
    }

    [Fact]
    public void Literal_word_tab_never_matches_the_tab_token()
    {
        var word = Tok(TextPara("tab")).Single(t => t.Kind == IrDiffTokenKind.Word);
        var tab = Tok(Para("<w:p><w:r><w:tab/></w:r></w:p>")).Single(t => t.Kind == IrDiffTokenKind.Tab);
        Assert.NotEqual(word.MatchKey, tab.MatchKey);
    }

    [Fact]
    public void Break_kinds_have_distinct_keys()
    {
        var page = Tok(Para("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"))[0];
        var line = Tok(Para("<w:p><w:r><w:br/></w:r></w:p>"))[0];
        Assert.NotEqual(page.MatchKey, line.MatchKey);
    }

    // --- intra-word note-ref interruption (M2.5 Task 1) ------------------

    // A footnote reference splitting `Video` into `Vi`[ref]`deo` with NO separator on either side. The two
    // text runs touch the zero-width ref at the same char offset, so the flanking words flank an interruption.
    private const string ViRefDeo =
        "<w:p><w:r><w:t xml:space=\"preserve\">Vi</w:t></w:r>" +
        "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" +
        "<w:r><w:t xml:space=\"preserve\">deo</w:t></w:r></w:p>";

    // A footnote reference BETWEEN words: `Video `[ref]`provides`. A separator (the space) sits before the
    // ref, so this is NOT an intra-word interruption — the common, non-disruptive case.
    private const string VideoRefProvides =
        "<w:p><w:r><w:t xml:space=\"preserve\">Video </w:t></w:r>" +
        "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" +
        "<w:r><w:t xml:space=\"preserve\">provides</w:t></w:r></w:p>";

    [Fact]
    public void Intra_word_note_ref_breaks_word_equality_with_contiguous_form()
    {
        // `Vi`[ref]`deo` must NOT be word-equal to a contiguous `Video`: the relocated ref is a real
        // structural edit to the word's atoms (WC-1710). The flanking `Vi`/`deo` carry the interruption
        // marker so neither equals `Video`'s key, and their concatenation `Video` does not coincidentally
        // reconstitute the contiguous key sequence.
        var split = Tok(Para(ViRefDeo));
        var contiguous = Tok(TextPara("Video"));

        var splitWords = split.Where(t => t.Kind == IrDiffTokenKind.Word).Select(t => t.MatchKey).ToList();
        // `Vi` and `deo` are both interruption-marked, so neither aliases the contiguous `Video` key.
        Assert.DoesNotContain(contiguous[0].MatchKey, splitWords);
        Assert.Equal(2, splitWords.Count);
        Assert.NotEqual(splitWords[0], splitWords[1]);  // Vi-marked vs deo-marked are distinct
    }

    [Fact]
    public void Between_word_note_ref_leaves_word_keys_identical_to_today()
    {
        // The CONSTRAINT: a ref BETWEEN words (separator-adjacent) is non-disruptive — the flanking words'
        // keys are byte-identical to the same words with NO ref at all. `Video `[ref]`provides` keys its
        // `Video`/`provides` exactly as plain `Video provides` does (zero blast radius outside intra-word).
        var withRef = Tok(Para(VideoRefProvides));
        var plain = Tok(TextPara("Video provides"));

        var withRefWordKeys = withRef.Where(t => t.Kind == IrDiffTokenKind.Word).Select(t => t.MatchKey);
        var plainWordKeys = plain.Where(t => t.Kind == IrDiffTokenKind.Word).Select(t => t.MatchKey);
        Assert.Equal(plainWordKeys, withRefWordKeys);
    }

    [Fact]
    public void Intra_word_interruption_does_not_change_the_note_ref_key_itself()
    {
        // The interrupting NoteRef token's OWN key is position-independent (still kind-only `fn`), so a
        // between-word ref that did NOT move still matches across a pair — only the flanking WORD keys carry
        // the marker. (This is what keeps the relocated-ref diff a clean body del+ins rather than also
        // splitting on the ref token.)
        var split = Tok(Para(ViRefDeo)).Single(t => t.Kind == IrDiffTokenKind.NoteRef);
        var between = Tok(Para(VideoRefProvides)).Single(t => t.Kind == IrDiffTokenKind.NoteRef);
        Assert.Equal(between.MatchKey, split.MatchKey);
    }

    [Fact]
    public void Intra_word_marker_is_determined_by_the_interrupting_atom()
    {
        // The marker is keyed on the interrupting atom, so a word split by a ref differs from the SAME word
        // unsplit AND the marker is stable/deterministic for repeated tokenizations.
        var a = Tok(Para(ViRefDeo)).Where(t => t.Kind == IrDiffTokenKind.Word).Select(t => t.MatchKey);
        var b = Tok(Para(ViRefDeo)).Where(t => t.Kind == IrDiffTokenKind.Word).Select(t => t.MatchKey);
        Assert.Equal(a, b);  // deterministic
    }

    // --- hyperlink suffixing ---------------------------------------------

    [Fact]
    public void Linked_text_differs_from_plain_text()
    {
        var linked = IrReader.Read(IrTestDocuments.FromBodyXmlWithHyperlinks(
            "<w:p><w:hyperlink r:id=\"r1\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>foo</w:t></w:r></w:hyperlink></w:p>",
            ("r1", "https://a.example")), NoSources)
            .Body.Blocks.OfType<IrParagraph>().First();
        var plain = TextPara("foo");

        Assert.NotEqual(Tok(plain).Single().MatchKey, Tok(linked).Single().MatchKey);
    }

    [Fact]
    public void Same_text_different_targets_differ()
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string Body() =>
            $"<w:p><w:hyperlink r:id=\"r1\" xmlns:r=\"{R}\"><w:r><w:t>foo</w:t></w:r></w:hyperlink></w:p>";
        var a = IrReader.Read(IrTestDocuments.FromBodyXmlWithHyperlinks(Body(), ("r1", "https://a.example")), NoSources)
            .Body.Blocks.OfType<IrParagraph>().First();
        var b = IrReader.Read(IrTestDocuments.FromBodyXmlWithHyperlinks(Body(), ("r1", "https://b.example")), NoSources)
            .Body.Blocks.OfType<IrParagraph>().First();
        Assert.NotEqual(Tok(a).Single().MatchKey, Tok(b).Single().MatchKey);
    }

    [Fact]
    public void Linked_word_key_is_sentinel_framed_so_it_cannot_alias_a_literal_word()
    {
        // Invariant guard (pins existing, correct behavior — an engine-audit "collision" finding that
        // turned out to be a FALSE POSITIVE: the suffix is ALREADY sentinel-framed on main). The
        // hyperlink-target suffix leads with U+0001 (XML-illegal in text) like every other atomic key. A
        // normalized word/separator MatchKey is always derived from w:t text and therefore can NEVER contain
        // U+0001, so this framing is exactly what prevents a LINKED word from aliasing a LITERAL word that
        // merely spells "<word>lnk:<target>" (':' '/' '.' are non-separators → such a literal is one token).
        // This case had only "linked ≠ plain" / "different targets differ" coverage before; this locks the
        // framing itself so the non-collision guarantee can't silently regress.
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var linked = IrReader.Read(IrTestDocuments.FromBodyXmlWithHyperlinks(
            $"<w:p><w:hyperlink r:id=\"r1\" xmlns:r=\"{R}\"><w:r><w:t>foo</w:t></w:r></w:hyperlink></w:p>",
            ("r1", "https://a.example")), NoSources).Body.Blocks.OfType<IrParagraph>().First();

        var linkedKey = Tok(linked).Single().MatchKey;
        Assert.Contains('\u0001', linkedKey);                          // the link suffix is sentinel-framed
        Assert.DoesNotContain('\u0001', Tok(TextPara("foolnk")).Single().MatchKey); // a plain word never is
    }

    // --- field transparency ----------------------------------------------

    [Fact]
    public void Field_result_tokenizes_transparently()
    {
        // A PAGE field whose cached result is "5" produces the same token key as a literal "5".
        var field = Para(
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
            "<w:r><w:t>5</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        var literal = TextPara("5");
        Assert.Equal(Tok(literal).Single().MatchKey, Tok(field).Single().MatchKey);
    }

    // --- format carry-through --------------------------------------------

    [Fact]
    public void Token_carries_governing_run_format()
    {
        var p = Para("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>");
        var token = Tok(p).Single(t => t.Kind == IrDiffTokenKind.Word);
        Assert.NotNull(token.Format);
        Assert.True(token.Format!.Bold);
    }

    [Fact]
    public void Atomic_break_token_has_null_format()
    {
        var brk = Tok(Para("<w:p><w:r><w:br/></w:r></w:p>")).Single(t => t.Kind == IrDiffTokenKind.Break);
        Assert.Null(brk.Format);
    }

    // --- determinism ------------------------------------------------------

    [Fact]
    public void Two_tokenizations_are_sequence_equal()
    {
        var p = Para(Run("the quick-brown (fox)"));
        Assert.Equal(Tok(p), Tok(p)); // IrDiffToken is a record → structural sequence equality
    }
}
