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
        // U+00A0 is NOT in the separator set, so "a\u00A0b" is one word; conflation folds the NBSP
        // in the key to a regular space so it equals the space-separated word's key.
        // (Space IS a separator, so "a b" splits into Word/Sep/Word; NBSP is NOT, so "a\u00A0b" is one
        // word. The behavior under test is the key fold: the NBSP word's key uses a regular space.)
        var on = new IrDiffSettings { ConflateBreakingAndNonbreakingSpaces = true };
        var nbsp = Tok(TextPara("a\u00A0b"), on)[0];
        Assert.Equal(IrDiffTokenKind.Word, nbsp.Kind);
        Assert.Equal("a\u00A0b", nbsp.Text);     // raw preserved
        Assert.Equal("a b", nbsp.MatchKey);       // folded to a regular space
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
    public void Tab_break_noteref_have_atomic_kinds()
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
