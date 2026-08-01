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
/// Word Compare's "Fields" comparison option (<see cref="DocxDiffSettings.CompareFields"/>): whether a
/// paragraph's FIELD CODE state — instruction text, simple-vs-complex form, inline position, fldChar
/// scaffolding — participates in the comparison. Field RESULTS are ordinary text either way.
/// </summary>
public class DocxDiffFieldsOptionTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static readonly DocxDiffSettings Off = new() { CompareFields = false };

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
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyXml}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("fields.docx", ms.ToArray());
    }

    /// <summary>A complex (w:fldChar) field surrounded by unique prose, so the field paragraph anchors on
    /// content rather than falling out of a degenerate single-paragraph alignment.</summary>
    private static WmlDocument ComplexFieldDoc(string instruction, string result = "5", string tail = "")
        => Doc(
            "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            $"<w:r><w:instrText xml:space=\"preserve\">{instruction}</w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
            $"<w:r><w:t>{result}</w:t></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            $"<w:r><w:t xml:space=\"preserve\"> of many.{tail}</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    private static WmlDocument SimpleFieldDoc(string instruction) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
        $"<w:p><w:r><w:t xml:space=\"preserve\">Page </w:t></w:r><w:fldSimple w:instr=\"{instruction}\">" +
        "<w:r><w:t>5</w:t></w:r></w:fldSimple></w:p>" +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    /// <summary>The same complex field, inside a single-cell table.</summary>
    private static WmlDocument TableFieldDoc(string instruction) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>" +
        "<w:p><w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        $"<w:r><w:instrText xml:space=\"preserve\">{instruction}</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>5</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>" +
        "</w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    /// <summary>The complex field as a run sequence, for reuse outside the body scope.</summary>
    private static string FieldRuns(string instruction) =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        $"<w:r><w:instrText xml:space=\"preserve\">{instruction}</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>5</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    /// <summary>The field lives in a FOOTER story; the body is identical on both sides.</summary>
    private static WmlDocument FooterFieldDoc(string instruction)
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
            {
                fw.Write($"<w:ftr xmlns:w=\"{W}\"><w:p>" +
                    $"<w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>{FieldRuns(instruction)}</w:p></w:ftr>");
            }

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:footerReference w:type=\"default\" r:id=\"rFtr1\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("footer-field.docx", ms.ToArray());
    }

    /// <summary>The field lives in a FOOTNOTE definition; the body is identical on both sides.</summary>
    private static WmlDocument FootnoteFieldDoc(string instruction)
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
                    "<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">Note page </w:t></w:r>" +
                    $"{FieldRuns(instruction)}</w:p></w:footnote></w:footnotes>");
            }

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Cited here.</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("footnote-field.docx", ms.ToArray());
    }

    /// <summary>The field paragraph sits either after the intro or at the end — a relocation.</summary>
    private static WmlDocument MovableFieldDoc(bool atEnd, string instruction)
    {
        var fieldPara = "<w:p><w:r><w:t xml:space=\"preserve\">Movable page </w:t></w:r>" +
            FieldRuns(instruction) + "<w:r><w:t xml:space=\"preserve\"> counted here today.</w:t></w:r></w:p>";
        var intro = "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>";
        var middle = "<w:p><w:r><w:t>Beta unique middle paragraph.</w:t></w:r></w:p>";
        var closing = "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>";
        return Doc(atEnd
            ? intro + middle + closing + fieldPara
            : intro + fieldPara + middle + closing);
    }

    /// <summary>A complex HYPERLINK field — the reader canonicalizes it away from the field envelope.</summary>
    private static WmlDocument HyperlinkFieldDoc(string url) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        $"<w:r><w:instrText xml:space=\"preserve\"> HYPERLINK \"{url}\" </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>the site</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>" +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    private static string BodyXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.Document.Body!.OuterXml;
    }

    private static void AssertValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(wDoc).Select(e => e.Description));
    }

    private static WmlDocument AcceptedDoc(WmlDocument compared) => new("a.docx",
        Docxodus.Internal.DocxDiffOps.AcceptRevisions(compared.DocumentByteArray));

    private static WmlDocument RejectedDoc(WmlDocument compared) => new("r.docx",
        Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray));

    private static string Accepted(WmlDocument compared) => BodyXml(AcceptedDoc(compared));

    private static string Rejected(WmlDocument compared) => BodyXml(RejectedDoc(compared));

    /// <summary>Concatenated XML of every header/footer part (the story scope's output).</summary>
    private static string StoryXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var main = wDoc.MainDocumentPart!;
        return string.Concat(main.HeaderParts.Select(p => p.Header.OuterXml))
             + string.Concat(main.FooterParts.Select(p => p.Footer.OuterXml));
    }

    private static string NotesXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.FootnotesPart?.Footnotes.OuterXml ?? string.Empty;
    }

    [Fact]
    public void CompareFields_defaults_to_true()
    {
        Assert.True(new DocxDiffSettings().CompareFields);
    }

    // --- default (on): a field-code change is tracked ---------------------------------------------

    [Fact]
    public void Default_reports_an_instruction_only_change()
    {
        var revisions = DocxDiff.GetRevisions(ComplexFieldDoc(" PAGE "), ComplexFieldDoc(" NUMPAGES "));

        Assert.NotEmpty(revisions);
    }

    [Fact]
    public void Default_tracks_the_old_instruction_as_delInstrText()
    {
        // The field plumbing cannot be sliced into a partial revision, so the paragraph is rendered as one
        // reversible del/ins pair — the deleted instruction as w:delInstrText, the new one as w:instrText.
        var compared = DocxDiff.Compare(ComplexFieldDoc(" PAGE "), ComplexFieldDoc(" NUMPAGES "));

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("<w:delInstrText", body);
        Assert.Contains(" PAGE ", body);
        Assert.Contains(" NUMPAGES ", body);
    }

    [Fact]
    public void Default_round_trips_an_instruction_change()
    {
        var compared = DocxDiff.Compare(ComplexFieldDoc(" PAGE "), ComplexFieldDoc(" NUMPAGES "));

        var accepted = Accepted(compared);
        Assert.Contains(" NUMPAGES ", accepted);
        Assert.DoesNotContain(" PAGE ", accepted);

        var rejected = Rejected(compared);
        Assert.Contains(" PAGE ", rejected);
        Assert.DoesNotContain(" NUMPAGES ", rejected);
    }

    [Fact]
    public void Default_reports_a_changed_REF_target_with_an_identical_result()
    {
        // Same displayed text, different bookmark: only the code moved. This is the case a
        // result-only comparison (WmlComparer) cannot see at all.
        var left = ComplexFieldDoc(" REF _RefA \\h ");
        var right = ComplexFieldDoc(" REF _RefB \\h ");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));

        var compared = DocxDiff.Compare(left, right);
        AssertValid(compared);
        Assert.Contains("<w:delInstrText", BodyXml(compared));
        Assert.Contains("_RefB", Accepted(compared));
        Assert.Contains("_RefA", Rejected(compared));
    }

    [Fact]
    public void Default_reports_an_added_format_switch()
    {
        var left = ComplexFieldDoc(" DATE ");
        var right = ComplexFieldDoc(" DATE \\@ \"yyyy-MM-dd\" ");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));

        var compared = DocxDiff.Compare(left, right);
        AssertValid(compared);
        Assert.Contains("<w:delInstrText", BodyXml(compared));
        Assert.Contains("yyyy-MM-dd", Accepted(compared));
        Assert.DoesNotContain("yyyy-MM-dd", Rejected(compared));
    }

    [Fact]
    public void Default_reports_a_simple_field_instruction_change()
    {
        var revisions = DocxDiff.GetRevisions(SimpleFieldDoc(" PAGE "), SimpleFieldDoc(" NUMPAGES "));

        Assert.NotEmpty(revisions);
    }

    // --- off: the code stops being compared -------------------------------------------------------

    [Fact]
    public void CompareFields_false_silences_an_instruction_only_change()
    {
        Assert.Empty(DocxDiff.GetRevisions(ComplexFieldDoc(" PAGE "), ComplexFieldDoc(" NUMPAGES "), Off));
    }

    [Fact]
    public void CompareFields_false_silences_a_simple_field_instruction_change()
    {
        Assert.Empty(DocxDiff.GetRevisions(SimpleFieldDoc(" PAGE "), SimpleFieldDoc(" NUMPAGES "), Off));
    }

    [Fact]
    public void CompareFields_false_emits_no_revision_markup_at_all()
    {
        var compared = DocxDiff.Compare(ComplexFieldDoc(" PAGE "), ComplexFieldDoc(" NUMPAGES "), Off);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.DoesNotContain("<w:ins ", body);
        Assert.DoesNotContain("<w:del ", body);
        Assert.DoesNotContain("<w:delInstrText", body);
    }

    [Fact]
    public void CompareFields_false_carries_the_right_field_code_through_untracked()
    {
        // The documented one-sidedness: an uncompared difference cannot be reversible. The output is
        // right-shaped, so accept is still exact and reject keeps the right code.
        var compared = DocxDiff.Compare(ComplexFieldDoc(" PAGE "), ComplexFieldDoc(" NUMPAGES "), Off);

        Assert.Contains(" NUMPAGES ", BodyXml(compared));
        Assert.DoesNotContain(" PAGE ", BodyXml(compared));
        Assert.Contains(" NUMPAGES ", Accepted(compared));
        Assert.Contains(" NUMPAGES ", Rejected(compared));
    }

    [Fact]
    public void CompareFields_false_still_reports_a_field_RESULT_change()
    {
        // Results are transparent to the tokenizer — plain text on both settings. The instruction is
        // identical here, so this cannot pass for a field-code reason.
        var revisions = DocxDiff.GetRevisions(
            ComplexFieldDoc(" PAGE ", result: "5"), ComplexFieldDoc(" PAGE ", result: "7"), Off);

        Assert.NotEmpty(revisions);
    }

    [Fact]
    public void CompareFields_false_still_reports_surrounding_prose_changes()
    {
        var compared = DocxDiff.Compare(
            ComplexFieldDoc(" PAGE ", tail: " Sigma."),
            ComplexFieldDoc(" NUMPAGES ", tail: " Kappa."),
            Off);

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("Kappa", body);
        Assert.Contains("Sigma", body);
        Assert.Contains("<w:ins ", body);
        Assert.Contains("<w:del ", body);
        // The prose change is tracked at word granularity, NOT lowered to a whole-paragraph replacement
        // by the (now uncompared) field code.
        Assert.DoesNotContain("<w:delInstrText", body);
    }

    [Fact]
    public void CompareFields_false_round_trips_the_prose_change_beside_an_uncompared_field()
    {
        var compared = DocxDiff.Compare(
            ComplexFieldDoc(" PAGE ", tail: " Sigma."),
            ComplexFieldDoc(" NUMPAGES ", tail: " Kappa."),
            Off);

        // This is the one shape that fine-slices PAST a differing field carrier, so validate the
        // materialized packages, not just their text.
        AssertValid(AcceptedDoc(compared));
        AssertValid(RejectedDoc(compared));

        var accepted = Accepted(compared);
        Assert.Contains("Kappa", accepted);
        Assert.DoesNotContain("Sigma", accepted);

        var rejected = Rejected(compared);
        Assert.Contains("Sigma", rejected);
        Assert.DoesNotContain("Kappa", rejected);
    }

    [Fact]
    public void CompareFields_false_is_a_no_op_when_the_field_is_unchanged()
    {
        var left = ComplexFieldDoc(" PAGE ");
        var right = ComplexFieldDoc(" PAGE ");

        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
    }

    [Fact]
    public void CompareFields_reaches_inside_a_table()
    {
        // A cell's identity folds its paragraphs' field-carrier digests at READ time, so the row aligns
        // unequal either way — but the per-paragraph op is still gated, so the setting decides.
        var left = TableFieldDoc(" PAGE ");
        var right = TableFieldDoc(" NUMPAGES ");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.DoesNotContain("<w:delInstrText", body);
        Assert.Contains(" NUMPAGES ", body);
    }

    // --- per-scope one-sidedness when off -----------------------------------------------------------

    [Fact]
    public void CompareFields_false_keeps_the_LEFT_code_in_a_footer_story()
    {
        // An uncompared difference is not reversible, and which side survives follows the scope's output
        // source. A story that produces no ops keeps its part carried over verbatim from the LEFT package
        // (the renderer only rebuilds parts that produced ops), so here it is reject — not accept — that
        // is exact. The body-scope test above pins the opposite direction.
        var left = FooterFieldDoc(" PAGE ");
        var right = FooterFieldDoc(" NUMPAGES ");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Contains(" PAGE ", StoryXml(compared));
        Assert.DoesNotContain(" NUMPAGES ", StoryXml(compared));
        Assert.Contains(" PAGE ", StoryXml(AcceptedDoc(compared)));
        Assert.Contains(" PAGE ", StoryXml(RejectedDoc(compared)));
    }

    [Fact]
    public void CompareFields_false_keeps_the_LEFT_code_in_a_footnote()
    {
        // Same carry-over rule as the footer story: a note definition producing no ops is not rebuilt.
        var left = FootnoteFieldDoc(" PAGE ");
        var right = FootnoteFieldDoc(" NUMPAGES ");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Contains(" PAGE ", NotesXml(compared));
        Assert.DoesNotContain(" NUMPAGES ", NotesXml(compared));
        Assert.Contains(" PAGE ", NotesXml(AcceptedDoc(compared)));
    }

    [Fact]
    public void Default_tracks_a_footnote_field_code_change_reversibly()
    {
        var compared = DocxDiff.Compare(FootnoteFieldDoc(" PAGE "), FootnoteFieldDoc(" NUMPAGES "));

        AssertValid(compared);
        Assert.Contains(" NUMPAGES ", NotesXml(AcceptedDoc(compared)));
        Assert.DoesNotContain(" NUMPAGES ", NotesXml(RejectedDoc(compared)));
        Assert.Contains(" PAGE ", NotesXml(RejectedDoc(compared)));
    }

    // --- shape and carve-outs -----------------------------------------------------------------------

    [Fact]
    public void CompareFields_decides_whether_a_moved_paragraph_keeps_native_move_markup()
    {
        // The Moved call site. On, the field-code difference lowers the pair to a plain del/ins pair and
        // the native move markup is lost; off, the relocation stays a w:moveFrom/w:moveTo pair. Both
        // round-trip (a move materializes both copies).
        var left = MovableFieldDoc(atEnd: false, " PAGE ");
        var right = MovableFieldDoc(atEnd: true, " NUMPAGES ");

        var tracked = DocxDiff.Compare(left, right);
        AssertValid(tracked);
        Assert.DoesNotContain("<w:moveFrom ", BodyXml(tracked));
        Assert.Contains("<w:delInstrText", BodyXml(tracked));

        var untracked = DocxDiff.Compare(left, right, Off);
        AssertValid(untracked);
        Assert.Contains("<w:moveFrom ", BodyXml(untracked));
    }

    [Fact]
    public void CompareFields_does_not_gate_a_canonicalized_HYPERLINK_target()
    {
        // Deliberate carve-out: the reader folds a clean HYPERLINK field into an IrHyperlink whose target
        // is ordinary CONTENT identity, not field-envelope state, so the target never reaches the gate.
        var left = HyperlinkFieldDoc("http://a.example");
        var right = HyperlinkFieldDoc("http://b.example");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.NotEmpty(DocxDiff.GetRevisions(left, right, Off));
    }

    [Fact]
    public void Consolidate_honors_CompareFields_per_reviewer()
    {
        // Consolidate passes the setting through rather than forcing it (unlike CompareHeadersFooters).
        // Off, a reviewer's field-code-only edit is not merged and the base code survives.
        var baseDoc = ComplexFieldDoc(" PAGE ");
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = ComplexFieldDoc(" NUMPAGES "), Author = "Reviewer One" },
        };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers);
        AssertValid(merged);
        Assert.Contains(" NUMPAGES ", Accepted(merged));

        var mergedOff = DocxDiff.Consolidate(baseDoc, reviewers,
            new DocxDiffConsolidateSettings { Diff = Off });
        AssertValid(mergedOff);
        Assert.Contains(" PAGE ", Accepted(mergedOff));
        Assert.DoesNotContain(" NUMPAGES ", Accepted(mergedOff));
    }

    // --- wire ---------------------------------------------------------------------------------------

    [Fact]
    public void CompareFields_parses_from_the_settings_wire()
    {
        var left = ComplexFieldDoc(" PAGE ").DocumentByteArray;
        var right = ComplexFieldDoc(" NUMPAGES ").DocumentByteArray;

        var tracked = BodyXml(new WmlDocument("d.docx",
            Docxodus.Internal.DocxDiffOps.Compare(left, right, null)));
        var untracked = BodyXml(new WmlDocument("d.docx",
            Docxodus.Internal.DocxDiffOps.Compare(left, right, "{\"compareFields\":false}")));

        Assert.Contains("delInstrText", tracked);
        Assert.DoesNotContain("delInstrText", untracked);
    }
}
