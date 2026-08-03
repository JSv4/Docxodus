#nullable enable

using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace OxPt;

/// <summary>
/// Word Compare's "Show changes at" radio pair (<see cref="DocxDiffSettings.ChangeGranularity"/>):
/// <see cref="DocxDiffChangeGranularity.Word"/> marks a changed word whole,
/// <see cref="DocxDiffChangeGranularity.Character"/> narrows the revision to the characters that differ.
/// </summary>
public class DocxDiffChangeGranularityTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static readonly DocxDiffSettings Chars =
        new() { ChangeGranularity = DocxDiffChangeGranularity.Character };

    private static WmlDocument Doc(string bodyXml)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>{bodyXml}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("granularity.docx", ms.ToArray());
    }

    /// <summary>One sentence between two unique anchoring paragraphs, so the pair aligns as Modified.</summary>
    private static WmlDocument Sentence(string text) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    /// <summary>The changed word carries its own run so a side can format it differently.</summary>
    private static WmlDocument FormattedWord(string word, bool bold) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t xml:space=\"preserve\">The </w:t></w:r>" +
        $"<w:r>{(bold ? "<w:rPr><w:b/></w:rPr>" : "")}<w:t>{word}</w:t></w:r>" +
        "<w:r><w:t xml:space=\"preserve\"> is bright.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    private static WmlDocument TableSentence(string text) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>" +
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    private static WmlDocument FooterSentence(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var footer = main.AddNewPart<FooterPart>("rFtr1");
            using (var fs = footer.GetStream(FileMode.Create, FileAccess.Write))
            using (var fw = new StreamWriter(fs, new UTF8Encoding(false)))
                fw.Write($"<w:ftr xmlns:w=\"{W}\"><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:ftr>");

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:footerReference w:type=\"default\" r:id=\"rFtr1\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("footer-granularity.docx", ms.ToArray());
    }

    /// <summary>Same display text on both sides; only the hyperlink relationship target differs.</summary>
    private static WmlDocument HyperlinkDoc(string url)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.AddHyperlinkRelationship(new System.Uri(url), true, "rL1");

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Please </w:t></w:r>" +
                "<w:hyperlink r:id=\"rL1\"><w:r><w:t>clickhere</w:t></w:r></w:hyperlink>" +
                "<w:r><w:t xml:space=\"preserve\"> now.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("hyperlink-granularity.docx", ms.ToArray());
    }

    /// <summary>The changed sentence lives in a footnote definition; the body is identical on both sides.</summary>
    private static WmlDocument FootnoteSentence(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var notes = main.AddNewPart<FootnotesPart>("rFn1");
            using (var fs = notes.GetStream(FileMode.Create, FileAccess.Write))
            using (var fw = new StreamWriter(fs, new UTF8Encoding(false)))
            {
                fw.Write($"<w:footnotes xmlns:w=\"{W}\">" +
                    "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                    "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\">" +
                    "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                    $"<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>" +
                    "</w:footnote></w:footnotes>");
            }

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Cited here.</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("footnote-granularity.docx", ms.ToArray());
    }

    private static string NotesXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.FootnotesPart?.Footnotes.OuterXml ?? string.Empty;
    }

    private static string BodyXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.Document.Body!.OuterXml;
    }

    private static string FooterXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return string.Concat(wDoc.MainDocumentPart!.FooterParts.Select(p => p.Footer.OuterXml));
    }

    private static void AssertValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(wDoc).Select(e => e.Description));
    }

    /// <summary>Paragraph texts of a materialized package, so a round trip is compared at the text level.</summary>
    private static string[] Paragraphs(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>().Select(p => p.InnerText).ToArray();
    }

    private static string[] Accepted(WmlDocument compared) =>
        Paragraphs(Docxodus.Internal.DocxDiffOps.AcceptRevisions(compared.DocumentByteArray));

    private static string[] Rejected(WmlDocument compared) =>
        Paragraphs(Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray));

    /// <summary>The revision list rendered as "Type=text" so a grain change is asserted exactly.</summary>
    private static string[] RevisionSummary(WmlDocument left, WmlDocument right, DocxDiffSettings? settings) =>
        (settings is null ? DocxDiff.GetRevisions(left, right) : DocxDiff.GetRevisions(left, right, settings))
            .Select(r => $"{r.Type}={r.Text}").ToArray();

    // --- default and the word-level regression guard -------------------------------------------------

    [Fact]
    public void ChangeGranularity_defaults_to_word()
    {
        Assert.Equal(DocxDiffChangeGranularity.Word, new DocxDiffSettings().ChangeGranularity);
    }

    [Fact]
    public void Word_level_marks_the_whole_changed_word()
    {
        var body = BodyXml(DocxDiff.Compare(
            Sentence("The colour is bright."), Sentence("The color is bright.")));

        Assert.Contains("<w:delText>colour</w:delText>", body);
        Assert.Contains("<w:t>color</w:t>", body);
    }

    // --- character level: markup --------------------------------------------------------------------

    [Fact]
    public void Character_level_marks_only_a_deleted_character_inside_the_word()
    {
        // colour -> color: shared "colo" + "r", differing middle is the left-only "u". There is nothing
        // inserted, so no w:ins survives at all.
        var compared = DocxDiff.Compare(
            Sentence("The colour is bright."), Sentence("The color is bright."), Chars);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("<w:delText xml:space=\"preserve\">u</w:delText>", body);
        Assert.Contains("<w:t xml:space=\"preserve\">colo</w:t>", body);
        Assert.Contains("<w:t xml:space=\"preserve\">r</w:t>", body);
        Assert.DoesNotContain("colour", body);
        Assert.DoesNotContain("<w:ins ", body);
    }

    [Fact]
    public void Character_level_marks_only_a_changed_suffix()
    {
        // walked -> walking: shared "walk", no shared suffix, so both sides keep a differing middle.
        var compared = DocxDiff.Compare(
            Sentence("They walked home."), Sentence("They walking home."), Chars);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("<w:t xml:space=\"preserve\">walk</w:t>", body);
        Assert.Contains("<w:delText xml:space=\"preserve\">ed</w:delText>", body);
        Assert.Contains("<w:t xml:space=\"preserve\">ing</w:t>", body);
    }

    [Fact]
    public void Character_level_marks_only_an_inserted_character_inside_the_word()
    {
        // cat -> cart: shared "ca" + "t", so only the right-only "r" is marked and no w:del survives.
        var compared = DocxDiff.Compare(
            Sentence("The cat is here."), Sentence("The cart is here."), Chars);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("<w:t xml:space=\"preserve\">r</w:t>", body);
        Assert.DoesNotContain("<w:del ", body);
        Assert.DoesNotContain("<w:delText", body);
    }

    [Fact]
    public void Character_level_leaves_a_wholly_different_word_alone()
    {
        // cat/dog share no leading or trailing characters, so there is nothing to narrow — the output must
        // be identical to word level, not merely "still correct".
        var left = Sentence("The cat is bright.");
        var right = Sentence("The dog is bright.");

        Assert.Equal(BodyXml(DocxDiff.Compare(left, right)), BodyXml(DocxDiff.Compare(left, right, Chars)));
    }

    [Fact]
    public void Character_level_narrows_every_changed_word_in_a_paragraph()
    {
        var compared = DocxDiff.Compare(
            Sentence("The colour and the flavour differ."),
            Sentence("The color and the flavor differ."),
            Chars);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("<w:t xml:space=\"preserve\">colo</w:t>", body);
        Assert.Contains("<w:t xml:space=\"preserve\">flavo</w:t>", body);
        Assert.Equal(2, body.Split("<w:delText xml:space=\"preserve\">u</w:delText>").Length - 1);
    }

    // --- the round-trip contract, which the refinement must not touch --------------------------------

    [Theory]
    [InlineData("The colour is bright.", "The color is bright.")]
    [InlineData("They walked home.", "They walking home.")]
    [InlineData("The cat is here.", "The cart is here.")]
    [InlineData("The cat is bright.", "The dog is bright.")]
    [InlineData("Section 12345 applies.", "Section 12945 applies.")]
    public void Character_level_round_trips(string leftText, string rightText)
    {
        // A character shared by the deleted and the inserted text survives accept AND reject either way, so
        // lifting it out of the wrappers cannot change either side. Asserted against the WORD-level result
        // rather than a literal, so the two modes are proven equivalent end to end.
        var left = Sentence(leftText);
        var right = Sentence(rightText);
        var word = DocxDiff.Compare(left, right);
        var chars = DocxDiff.Compare(left, right, Chars);

        AssertValid(chars);
        Assert.Equal(Accepted(word), Accepted(chars));
        Assert.Equal(Rejected(word), Rejected(chars));
        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(chars));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(chars));
    }

    // --- character level: the revision list ---------------------------------------------------------

    [Fact]
    public void Character_level_narrows_the_revision_text()
    {
        var left = Sentence("The colour is bright.");
        var right = Sentence("The color is bright.");

        Assert.Equal(new[] { "Deleted=colour", "Inserted=color" }, RevisionSummary(left, right, null));
        Assert.Equal(new[] { "Deleted=u" }, RevisionSummary(left, right, Chars));
    }

    [Fact]
    public void Character_level_keeps_both_revisions_when_both_sides_differ()
    {
        Assert.Equal(
            new[] { "Deleted=ed", "Inserted=ing" },
            RevisionSummary(Sentence("They walked home."), Sentence("They walking home."), Chars));
    }

    [Fact]
    public void Character_level_leaves_an_unrelated_word_pair_untouched_in_revisions()
    {
        var left = Sentence("The cat is bright.");
        var right = Sentence("The dog is bright.");

        Assert.Equal(RevisionSummary(left, right, null), RevisionSummary(left, right, Chars));
    }

    // --- scope reach and the format guard -----------------------------------------------------------

    [Fact]
    public void Character_level_reaches_a_table_cell()
    {
        var compared = DocxDiff.Compare(
            TableSentence("The colour is bright."), TableSentence("The color is bright."), Chars);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("<w:delText xml:space=\"preserve\">u</w:delText>", body);
        Assert.DoesNotContain("colour", body);
    }

    [Fact]
    public void Character_level_reaches_a_footer_story()
    {
        // The post-pass iterates every rendered story part, not just the main document.
        var compared = DocxDiff.Compare(
            FooterSentence("The colour is bright."), FooterSentence("The color is bright."), Chars);

        AssertValid(compared);
        var footer = FooterXml(compared);
        Assert.Contains("<w:delText xml:space=\"preserve\">u</w:delText>", footer);
        Assert.DoesNotContain("colour", footer);
    }

    [Fact]
    public void Character_level_leaves_a_pair_whose_formatting_also_changed_at_word_level()
    {
        // The shared characters are NOT unchanged when the word's format changed too — lifting them into one
        // plain run would have to pick a side and would drop the format change. Such a pair stays whole.
        var left = FormattedWord("colour", bold: false);
        var right = FormattedWord("color", bold: true);

        var compared = DocxDiff.Compare(left, right, Chars);
        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("colour", body);
        Assert.DoesNotContain("<w:t xml:space=\"preserve\">colo</w:t>", body);

        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(compared));
    }

    // --- the engine's truth is untouched -------------------------------------------------------------

    [Fact]
    public void ChangeGranularity_does_not_affect_the_edit_script()
    {
        // Alignment stays word-grained under both values; this is a rendering refinement only.
        var left = Sentence("The colour is bright.");
        var right = Sentence("The color is bright.");

        Assert.Equal(
            DocxDiff.GetEditScriptJson(left, right),
            DocxDiff.GetEditScriptJson(left, right, Chars));
    }

    [Fact]
    public void Character_level_is_a_no_op_on_identical_documents()
    {
        var left = Sentence("The colour is bright.");
        var right = Sentence("The colour is bright.");

        Assert.Empty(DocxDiff.GetRevisions(left, right, Chars));
        Assert.Equal(BodyXml(DocxDiff.Compare(left, right)), BodyXml(DocxDiff.Compare(left, right, Chars)));
    }

    // --- text that a naive char-by-char cut would corrupt --------------------------------------------

    [Theory]
    [InlineData("\U0001F600", "\U0001F601")]   // emoji: same high surrogate, differing low surrogate
    [InlineData("\U00020000", "\U00020001")]   // CJK Ext-B ideographs: likewise
    public void Character_level_never_splits_a_surrogate_pair(string leftChar, string rightChar)
    {
        // Comparing UTF-16 code units alone stops the common prefix BETWEEN the surrogates, leaving lone
        // surrogates that cannot be written as XML at all (the serializer throws). Cuts must land on a
        // grapheme-cluster boundary.
        var left = Sentence($"The {leftChar}x is bright.");
        var right = Sentence($"The {rightChar}x is bright.");

        var compared = DocxDiff.Compare(left, right, Chars);
        AssertValid(compared);
        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(compared));

        foreach (var revision in DocxDiff.GetRevisions(left, right, Chars))
            Assert.DoesNotContain(revision.Text, c => char.IsSurrogate(c) && revision.Text.Length == 1);
    }

    [Fact]
    public void Character_level_keeps_a_combining_mark_with_its_base()
    {
        // café -> cafe differs only in the combining acute. Cutting between "e" and U+0301 would mark a bare
        // floating diacritic, so the whole cluster moves together.
        var left = Sentence("The café is open.");
        var right = Sentence("The cafe is open.");

        var compared = DocxDiff.Compare(left, right, Chars);
        AssertValid(compared);
        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(compared));

        // The prefix backs off to before the whole "e"+U+0301 cluster, so the marked middles are the complete
        // clusters "é" and "e" — never a bare U+0301, which a code-unit cut would have produced.
        Assert.Equal(
            new[] { "Deleted=é", "Inserted=e" },
            RevisionSummary(left, right, Chars));
    }

    // --- pairs the refinement must decline ----------------------------------------------------------

    [Fact]
    public void Character_level_keeps_an_equal_text_pair_whole()
    {
        // A retargeted hyperlink keeps its display text, so the del/ins texts are EQUAL and the change lives
        // in the relationship. Narrowing would empty BOTH wrappers and erase the revision — losing the author,
        // the date and the fact that anything changed — so the pair must stay whole and the two surfaces agree.
        var left = HyperlinkDoc("http://a.example/");
        var right = HyperlinkDoc("http://b.example/");

        var atWord = RevisionSummary(left, right, null);
        Assert.NotEmpty(atWord);
        Assert.Equal(atWord, RevisionSummary(left, right, Chars));
        Assert.Equal(BodyXml(DocxDiff.Compare(left, right)), BodyXml(DocxDiff.Compare(left, right, Chars)));
    }

    [Fact]
    public void Character_level_declines_an_equal_text_sibling_pair_in_the_markup()
    {
        // The markup twin of the hyperlink case above, where the two wrappers ARE siblings (in the hyperlink
        // shape each sits inside its own w:hyperlink, so the pass never reaches this guard). Narrowing here
        // would empty both wrappers and collapse a tracked revision into one untracked run.
        var root = System.Xml.Linq.XElement.Parse(
            $"<w:p xmlns:w=\"{W}\">" +
            "<w:ins w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\" w:id=\"1\"><w:r><w:t>same</w:t></w:r></w:ins>" +
            "<w:del w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\" w:id=\"2\"><w:r><w:delText>same</w:delText></w:r></w:del>" +
            "</w:p>");

        Assert.False(Docxodus.Ir.Diff.IrCharacterGranularity.RefineInRoot(root));
        Assert.Contains("<w:ins", root.ToString());
        Assert.Contains("<w:del", root.ToString());
    }

    [Fact]
    public void Character_level_declines_a_pair_whose_wrappers_have_different_authors()
    {
        // Only pairs THIS engine produced are refinable — they all carry one author/date. A pair assembled
        // from two authors' wrappers must stay whole: lifting a shared character into an untracked run would
        // drop one author's claim to it. Exercised through the post-pass directly, since a two-way compare
        // stamps a single author by construction.
        var root = System.Xml.Linq.XElement.Parse(
            $"<w:p xmlns:w=\"{W}\">" +
            "<w:ins w:author=\"Ann\" w:date=\"2020-01-01T00:00:00Z\" w:id=\"1\"><w:r><w:t>color</w:t></w:r></w:ins>" +
            "<w:del w:author=\"Bob\" w:date=\"2020-01-01T00:00:00Z\" w:id=\"2\"><w:r><w:delText>colour</w:delText></w:r></w:del>" +
            "</w:p>");

        Assert.False(Docxodus.Ir.Diff.IrCharacterGranularity.RefineInRoot(root));
        Assert.Contains("colour", root.ToString());
    }

    [Fact]
    public void Character_level_declines_a_pair_separated_by_a_bookmark()
    {
        // The markup pass pairs by sibling adjacency; anything between the wrappers means they are not a pair
        // it can reason about, so it must degrade to a no-op rather than reach across.
        var root = System.Xml.Linq.XElement.Parse(
            $"<w:p xmlns:w=\"{W}\">" +
            "<w:ins w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\" w:id=\"1\"><w:r><w:t>color</w:t></w:r></w:ins>" +
            "<w:bookmarkStart w:id=\"9\" w:name=\"mark\"/>" +
            "<w:del w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\" w:id=\"2\"><w:r><w:delText>colour</w:delText></w:r></w:del>" +
            "</w:p>");

        Assert.False(Docxodus.Ir.Diff.IrCharacterGranularity.RefineInRoot(root));
    }

    [Fact]
    public void Consolidate_is_word_level_v1_ceiling()
    {
        // Forced off in BOTH the merger and the composite revision renderer, which receive settings by
        // different routes. A multi-author composite can place two reviewers' wrappers side by side, and the
        // composite markup renderer has no refinement pass — so a pass-through would narrow the revisions
        // while the markup stayed whole.
        var baseDoc = Sentence("The colour is bright.");
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = Sentence("The color is bright."), Author = "Reviewer One" },
        };
        var settings = new DocxDiffConsolidateSettings { Diff = Chars };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers, settings);
        AssertValid(merged);
        Assert.Contains("<w:delText>colour</w:delText>", BodyXml(merged));

        Assert.Contains("colour",
            DocxDiff.GetConsolidatedRevisions(baseDoc, reviewers, settings).Select(r => r.Text));
    }

    [Fact]
    public void Character_level_reaches_a_footnote()
    {
        var compared = DocxDiff.Compare(
            FootnoteSentence("The colour is bright."), FootnoteSentence("The color is bright."), Chars);

        AssertValid(compared);
        var notes = NotesXml(compared);
        Assert.Contains("<w:delText xml:space=\"preserve\">u</w:delText>", notes);
        Assert.DoesNotContain("colour", notes);
    }

    // --- wire ---------------------------------------------------------------------------------------

    [Fact]
    public void ChangeGranularity_parses_from_the_settings_wire()
    {
        var left = Sentence("The colour is bright.").DocumentByteArray;
        var right = Sentence("The color is bright.").DocumentByteArray;

        var atWord = BodyXml(new WmlDocument("w.docx",
            Docxodus.Internal.DocxDiffOps.Compare(left, right, null)));
        var atChar = BodyXml(new WmlDocument("c.docx",
            Docxodus.Internal.DocxDiffOps.Compare(left, right, "{\"changeGranularity\":1}")));

        Assert.Contains("colour", atWord);
        Assert.DoesNotContain("colour", atChar);
        Assert.Contains("<w:delText xml:space=\"preserve\">u</w:delText>", atChar);
    }
}
