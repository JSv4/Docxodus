#nullable enable
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// CROSS-KIND note nesting through the N-way CONSOLIDATE engine — the last untested corner of the note-scope
/// merge (item M-A #4). A footnote body that cites an endnote (and an endnote body that cites a footnote) must
/// keep its nested reference resolvable after the composite renderer's body-order renumber, exactly as the
/// two-way engine does (<see cref="Docxodus.Tests.DocxDiffFootnoteRobustnessTests.CrossKindNestedNoteReference_ToRenumberedNote_StaysResolvable"/>).
/// The existing <c>IrCompositeNoteTests</c> fixture is footnote-only with body-order ids, so it never exercised
/// this: cross-kind nesting, gapped ids that force a real shift, and the reviewer-INSERTED-note sub-case are all
/// new here.
/// </summary>
public class IrCompositeCrossKindNoteTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Wns = W;

    private const string FootnoteReserved =
        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";
    private const string EndnoteReserved =
        "<w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>" +
        "<w:endnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>";

    // ------------------------------------------------------------------ fixture builder

    /// <summary>Build a doc from raw fragments: <paramref name="bodyInner"/> is the body paragraphs (sectPr is
    /// appended), <paramref name="footnoteDefs"/>/<paramref name="endnoteDefs"/> are the REAL note definitions
    /// (reserved boilerplate is prepended). An endnotes part is added iff <paramref name="endnoteDefs"/> is set.</summary>
    private static WmlDocument Build(string bodyInner, string footnoteDefs, string? endnoteDefs)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            WritePartXml(main.AddNewPart<FootnotesPart>(),
                $"<w:footnotes xmlns:w=\"{W}\">{FootnoteReserved}{footnoteDefs}</w:footnotes>");
            if (endnoteDefs != null)
                WritePartXml(main.AddNewPart<EndnotesPart>(),
                    $"<w:endnotes xmlns:w=\"{W}\">{EndnoteReserved}{endnoteDefs}</w:endnotes>");

            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInner}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("doc.docx", ms.ToArray());
    }

    private static string Para(string text, string? refXml = null)
        => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>{refXml ?? ""}</w:p>";

    private static string FnRef(int id) => $"<w:r><w:footnoteReference w:id=\"{id}\"/></w:r>";
    private static string EnRef(int id) => $"<w:r><w:endnoteReference w:id=\"{id}\"/></w:r>";

    private static string Footnote(int id, string text, string? nestedRefXml = null)
        => $"<w:footnote w:id=\"{id}\"><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>{nestedRefXml ?? ""}</w:p></w:footnote>";
    private static string Endnote(int id, string text, string? nestedRefXml = null)
        => $"<w:endnote w:id=\"{id}\"><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>{nestedRefXml ?? ""}</w:p></w:endnote>";

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    private static WmlDocument Consolidate(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.Consolidate(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    // ------------------------------------------------------------------ oracles

    /// <summary>Every note reference of <paramref name="refLocal"/> kind ANYWHERE (document body, footnotes part,
    /// endnotes part) that does NOT resolve to exactly one non-reserved definition of <paramref name="defLocal"/>
    /// kind in its part. The cross-kind resolvability oracle (mirrors the two-way robustness test).</summary>
    private static List<string> UnresolvedRefs(WmlDocument doc, string refLocal, string defLocal, bool defsInFootnotesPart)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var fnRoot = main.FootnotesPart?.GetXDocument().Root;
        var enRoot = main.EndnotesPart?.GetXDocument().Root;
        var defRoot = defsInFootnotesPart ? fnRoot : enRoot;
        var defCounts = (defRoot?.Elements(Wns + defLocal)
                            .Where(e => e.Attribute(Wns + "type") == null)
                            .Select(e => (string?)e.Attribute(Wns + "id"))
                            .Where(x => x != null).Select(x => x!) ?? Enumerable.Empty<string>())
                        .GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var refs = new List<XElement>();
        var body = main.GetXDocument().Root?.Element(Wns + "body");
        if (body != null) refs.AddRange(body.Descendants(Wns + refLocal));
        if (fnRoot != null) refs.AddRange(fnRoot.Descendants(Wns + refLocal));
        if (enRoot != null) refs.AddRange(enRoot.Descendants(Wns + refLocal));
        return refs.Select(r => (string?)r.Attribute(Wns + "id"))
                   .Where(id => id == null || !defCounts.TryGetValue(id, out var n) || n != 1)
                   .Select(id => $"{refLocal}:{id ?? "(null)"}").ToList();
    }

    /// <summary>Assert BOTH cross-kind reference directions resolve everywhere in <paramref name="doc"/>.</summary>
    private static void AssertAllNoteRefsResolve(WmlDocument doc, string because)
    {
        Assert.True(UnresolvedRefs(doc, "footnoteReference", "footnote", defsInFootnotesPart: true).Count == 0,
            $"{because}: dangling footnote refs -> {string.Join(",", UnresolvedRefs(doc, "footnoteReference", "footnote", true))}");
        Assert.True(UnresolvedRefs(doc, "endnoteReference", "endnote", defsInFootnotesPart: false).Count == 0,
            $"{because}: dangling endnote refs -> {string.Join(",", UnresolvedRefs(doc, "endnoteReference", "endnote", false))}");
    }

    private static HashSet<string> SchemaErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(w)
            .Select(e => $"{e.Id}@{e.Path?.XPath}: {e.Description}").ToHashSet();
    }

    private static void AssertNoNewSchemaErrors(WmlDocument baseDoc, WmlDocument produced)
    {
        // Sem_MissingReferenceElement is EXCLUDED: OpenXmlValidator does not resolve a note reference that lives
        // inside another note's definition body against the sibling notes part (a footnoteReference nested in an
        // endnote body, or vice versa) — it false-positives on such refs even when they resolve, and the base doc
        // already carries the same error (see DocxDiffFootnoteRobustnessTests.CrossKindNestedNoteReference_...).
        // The base-subtraction cannot cancel it either, because the renumber legitimately changes the id embedded
        // in the message string. Reference RESOLVABILITY is instead asserted by the structural AssertAllNoteRefsResolve
        // oracle; every OTHER schema error (duplicate ids, malformed markup, etc.) is still caught here.
        static HashSet<string> Real(WmlDocument d) =>
            SchemaErrors(d).Where(e => !e.Contains("Sem_MissingReferenceElement")).ToHashSet();
        var baseErrors = Real(baseDoc);
        var newErrors = Real(produced).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0, $"new schema errors: {string.Join(" | ", newErrors.Take(5))}");
    }

    // ------------------------------------------------------------------ 1. base cross-kind nesting survives N-way renumber

    /// <summary>
    /// The base has GAPPED note ids (fn 2/3, en 5/8) so the body-order renumber genuinely shifts every id, and
    /// two cross-kind nested references (footnote 2 cites endnote 5, endnote 8 cites footnote 2). Two reviewers
    /// each make an edit — Alice edits a note (which is what ENTERS the composite note-render path) and Bob edits
    /// body text — so the renumber + cross-kind nested-reference sweep runs through Consolidate. Every reference,
    /// including the two nested cross-kind ones, must still resolve on the merged doc AND on accept AND on reject.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void CrossKind_nested_base_note_refs_survive_consolidate_renumber(ConflictResolution policy)
    {
        // fn2 cites en5 (cross-kind); en8 cites fn2 (reverse cross-kind); fn3/en5 are plain. Ids are gapped.
        string body =
            Para("Alpha cites the outer footnote.", FnRef(2)) +
            Para("Beta cites the inner endnote.", EnRef(5)) +
            Para("Gamma cites the other endnote.", EnRef(8)) +
            Para("Delta cites the plain footnote.", FnRef(3));
        string fnDefs =
            Footnote(2, "Outer footnote citing ", EnRef(5)) +
            Footnote(3, "A plain footnote.");
        string enDefs =
            Endnote(5, "The inner endnote, also body-cited.") +
            Endnote(8, "Endnote citing ", FnRef(2));

        var b = Build(body, fnDefs, enDefs);
        // Alice edits the plain footnote's text -> a real note diff, so the composite note-render path runs.
        var alice = Build(body, Footnote(2, "Outer footnote citing ", EnRef(5)) + Footnote(3, "A plain footnote EDITED."), enDefs);
        // Bob edits body text elsewhere -> genuine N-way merge.
        var bob = Build(
            Para("Alpha cites the outer footnote.", FnRef(2)) +
            Para("Beta cites the inner endnote.", EnRef(5)) +
            Para("Gamma cites the other endnote REVISED.", EnRef(8)) +
            Para("Delta cites the plain footnote.", FnRef(3)),
            fnDefs, enDefs);

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionProcessor.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        AssertAllNoteRefsResolve(merged, "merged");
        AssertAllNoteRefsResolve(accepted, "accept");
        AssertAllNoteRefsResolve(rejected, "reject");
        AssertNoNewSchemaErrors(b, merged);

        // reject ≡ base body.
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(rejected));
    }

    // ------------------------------------------------------------------ 2. inserted footnote cites EXISTING endnote

    /// <summary>
    /// A reviewer INSERTS a footnote whose body cites an EXISTING (base) endnote — a cross-kind nested reference
    /// living inside a reviewer-inserted note definition. It lands under a fresh output id, and its nested
    /// endnote reference must be remapped to the endnote's renumbered id (else it dangles).
    /// </summary>
    [Fact]
    public void Reviewer_inserted_footnote_citing_existing_endnote_stays_resolvable()
    {
        string body =
            Para("Alpha cites the footnote.", FnRef(2)) +
            Para("Beta cites the endnote.", EnRef(5));
        string fnDefs = Footnote(2, "The base footnote.");
        string enDefs = Endnote(5, "The base endnote.");

        var b = Build(body, fnDefs, enDefs);
        // Alice adds a new paragraph with a NEW footnote (her id 3) whose body cites the existing endnote 5.
        var alice = Build(
            body + Para("Gamma adds a footnote.", FnRef(3)),
            fnDefs + Footnote(3, "Inserted footnote citing ", EnRef(5)),
            enDefs);
        // Bob edits body text -> N-way.
        var bob = Build(
            Para("Alpha cites the footnote DELTA.", FnRef(2)) + Para("Beta cites the endnote.", EnRef(5)),
            fnDefs, enDefs);

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionProcessor.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        AssertAllNoteRefsResolve(merged, "merged");
        AssertAllNoteRefsResolve(accepted, "accept");
        AssertAllNoteRefsResolve(rejected, "reject");
        AssertNoNewSchemaErrors(b, merged);
    }

    // ------------------------------------------------------------------ 3. inserted endnote cites EXISTING footnote

    /// <summary>Symmetric to #2: a reviewer inserts an ENDNOTE whose body cites an EXISTING footnote.</summary>
    [Fact]
    public void Reviewer_inserted_endnote_citing_existing_footnote_stays_resolvable()
    {
        string body =
            Para("Alpha cites the footnote.", FnRef(2)) +
            Para("Beta cites the endnote.", EnRef(5));
        string fnDefs = Footnote(2, "The base footnote.");
        string enDefs = Endnote(5, "The base endnote.");

        var b = Build(body, fnDefs, enDefs);
        // Alice adds a new paragraph with a NEW endnote (her id 8) whose body cites the existing footnote 2.
        var alice = Build(
            body + Para("Gamma adds an endnote.", EnRef(8)),
            fnDefs,
            enDefs + Endnote(8, "Inserted endnote citing ", FnRef(2)));
        var bob = Build(
            Para("Alpha cites the footnote DELTA.", FnRef(2)) + Para("Beta cites the endnote.", EnRef(5)),
            fnDefs, enDefs);

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionProcessor.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        AssertAllNoteRefsResolve(merged, "merged");
        AssertAllNoteRefsResolve(accepted, "accept");
        AssertAllNoteRefsResolve(rejected, "reject");
        AssertNoNewSchemaErrors(b, merged);
    }

    // ------------------------------------------------------------------ 4. inserted footnote cites reviewer-INSERTED endnote

    /// <summary>
    /// The hardest sub-case (the one the audit flagged as neither obviously handled nor exercised): a reviewer
    /// inserts a footnote whose body cites an endnote the SAME reviewer inserts. The cited endnote receives a
    /// FRESH output id, but the nested reference inside the inserted footnote definition still carries the
    /// reviewer's own id — which the body-reference rewrite never touches (it only rewrites body clones). Its
    /// nested reference must be rewritten to the endnote's fresh output id, then follow the renumber, or it
    /// dangles on merge/accept/reject.
    /// </summary>
    [Fact]
    public void Reviewer_inserted_footnote_citing_reviewer_inserted_endnote_stays_resolvable()
    {
        string body =
            Para("Alpha cites the footnote.", FnRef(2)) +
            Para("Beta cites the endnote.", EnRef(5));
        string fnDefs = Footnote(2, "The base footnote.");
        string enDefs = Endnote(5, "The base endnote.");

        var b = Build(body, fnDefs, enDefs);
        // Alice inserts a footnote (her id 3) whose body cites an endnote (her id 6) she ALSO inserts.
        var alice = Build(
            body +
                Para("Gamma adds a footnote.", FnRef(3)) +
                Para("Epsilon adds an endnote.", EnRef(6)),
            fnDefs + Footnote(3, "Inserted footnote citing ", EnRef(6)),
            enDefs + Endnote(6, "The inserted endnote."));
        var bob = Build(
            Para("Alpha cites the footnote DELTA.", FnRef(2)) + Para("Beta cites the endnote.", EnRef(5)),
            fnDefs, enDefs);

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionProcessor.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        AssertAllNoteRefsResolve(merged, "merged");
        AssertAllNoteRefsResolve(accepted, "accept");
        AssertAllNoteRefsResolve(rejected, "reject");
        AssertNoNewSchemaErrors(b, merged);
    }
}
