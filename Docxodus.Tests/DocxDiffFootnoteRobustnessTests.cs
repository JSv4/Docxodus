#nullable enable
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Footnote id↔reference↔text robustness beyond the scenario corpus: cases where a note's <c>w:id</c> does NOT
/// equal its body-reference ordinal, so the renumber pass cannot rely on a positional coincidence to keep a
/// reference-deleted-but-definition-preserved note resolvable. The synthetic scenario base uses ids {1,2} in
/// reference order, which masks this; the real NVCA contract (111 footnotes, gaps) does not.
/// </summary>
public class DocxDiffFootnoteRobustnessTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A doc whose footnotes carry NON-sequential ids (1 and 5), referenced in body order [1, 5].</summary>
    private static byte[] BuildDocWithGappedFootnoteIds()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">First footnote text.</w:t></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"5\"><w:p><w:r><w:t xml:space=\"preserve\">Fifth footnote text with a gapped id.</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Alpha paragraph references one.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Beta paragraph references five.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"5\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>");
        }
        return ms.ToArray();
    }

    /// <summary>Right = left with the footnote-5-referencing paragraph wholly rewritten (its w:footnoteReference
    /// dropped). The fn5 DEFINITION lingers in the part (unreferenced) — the reconcile case.</summary>
    private static byte[] RewriteFiveReferencingParagraph(byte[] left)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var p = body.Elements<Paragraph>().First(x =>
                string.Concat(x.Descendants<Text>().Select(t => t.Text)).Contains("Beta paragraph"));
            foreach (var r in p.Elements<Run>().ToList()) r.Remove();
            p.AppendChild(new Run(new Text("Beta paragraph has been wholly rewritten.")
                { Space = SpaceProcessingModeValues.Preserve }));
        }
        return ms.ToArray();
    }

    /// <summary>A doc with the THREE reserved boilerplate footnotes Word can emit — separator (-1),
    /// continuationSeparator (0), AND continuationNotice with a POSITIVE id (1, as the real NVCA contract uses) —
    /// plus real footnotes referenced in the body.</summary>
    private static byte[] BuildDocWithPositiveIdReservedFootnote()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                // continuationNotice with POSITIVE id 1 — the reserved note the NVCA contract carries.
                "<w:footnote w:type=\"continuationNotice\" w:id=\"1\"><w:p/></w:footnote>" +
                "<w:footnote w:id=\"2\"><w:p><w:r><w:t xml:space=\"preserve\">First real footnote.</w:t></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"3\"><w:p><w:r><w:t xml:space=\"preserve\">Second real footnote.</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Alpha references the first.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Beta references the second.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"3\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>");
        }
        return ms.ToArray();
    }

    private static byte[] ReplaceWord(byte[] left, string find, string repl)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var t = doc.MainDocumentPart!.Document!.Body!.Descendants<Text>().First(x => x.Text.Contains(find));
            t.Text = t.Text.Replace(find, repl);
            t.Space = SpaceProcessingModeValues.Preserve;
        }
        return ms.ToArray();
    }

    [Fact]
    public void ReservedContinuationNoticeAtPositiveId_DoesNotCollideWithRenumberedRealFootnote()
    {
        // Editing a body paragraph (NOT a footnote) on a doc whose continuationNotice boilerplate occupies id 1
        // must not make a renumbered real footnote ALSO land on id 1. Reproduces the NVCA-scale duplicate that
        // every edit (even format/body-only) tripped; the synthetic corpus (reserved ids only -1/0) masks it.
        var left = new WmlDocument("left.docx", BuildDocWithPositiveIdReservedFootnote());
        var right = new WmlDocument("right.docx", ReplaceWord(left.DocumentByteArray, "Alpha", "Gamma"));

        var result = DocxDiff.Compare(left, right);

        var ids = FootnoteIds(result);
        Assert.Equal(ids.Distinct().Count(), ids.Count);

        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0, $"schema errors: {string.Join(" | ", newErrors.Take(5))}");

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(ReferencedFootnoteTexts(right), ReferencedFootnoteTexts(accepted));
        Assert.Equal(ReferencedFootnoteTexts(left), ReferencedFootnoteTexts(rejected));
    }

    /// <summary>A doc where footnote 2's BODY cites footnote 5 (a note-in-note reference), and footnote 5 is
    /// ALSO referenced from the document body — so the body-reference renumber pass remaps footnote 5's
    /// definition id, and the nested reference inside footnote 2 must be remapped too or it dangles.</summary>
    private static byte[] BuildNoteInNoteDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn, $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                // footnote 2's body itself references footnote 5 (note-in-note)
                "<w:footnote w:id=\"2\"><w:p><w:r><w:t xml:space=\"preserve\">Outer note citing </w:t></w:r><w:r><w:footnoteReference w:id=\"5\"/></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"5\"><w:p><w:r><w:t xml:space=\"preserve\">Inner note, also body-cited.</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");
            WritePartXml(main, $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Alpha cites the outer note.</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Beta cites the inner note.</w:t></w:r><w:r><w:footnoteReference w:id=\"5\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return ms.ToArray();
    }

    /// <summary>Every footnote reference ANYWHERE (document body AND inside note definition bodies) that does
    /// NOT resolve to exactly one footnote definition in the part.</summary>
    private static List<string> AllUnresolvedFootnoteRefs(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var fnRoot = main.FootnotesPart?.GetXDocument().Root;
        var defCounts = (fnRoot?.Elements(System.Xml.Linq.XName.Get("footnote", W))
                            .Select(e => (string?)e.Attribute(System.Xml.Linq.XName.Get("id", W)))
                            .Where(x => x != null).Select(x => x!) ?? Enumerable.Empty<string>())
                        .GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var refs = new List<System.Xml.Linq.XElement>();
        var body = main.GetXDocument().Root?.Element(System.Xml.Linq.XName.Get("body", W));
        if (body != null) refs.AddRange(body.Descendants(System.Xml.Linq.XName.Get("footnoteReference", W)));
        if (fnRoot != null) refs.AddRange(fnRoot.Descendants(System.Xml.Linq.XName.Get("footnoteReference", W)));
        return refs.Select(r => (string?)r.Attribute(System.Xml.Linq.XName.Get("id", W)))
                   .Where(id => id == null || !defCounts.TryGetValue(id, out var n) || n != 1)
                   .Select(id => id ?? "(null)").ToList();
    }

    [Fact]
    public void NestedNoteReference_ToRenumberedNote_StaysResolvable()
    {
        // A note-in-note reference (footnote 2 -> footnote 5) must survive the body-reference renumber that
        // remaps footnote 5's definition id. Before the fix, the nested ref kept id 5 while the definition moved
        // to id 2, dangling on compare/accept/reject. (OpenXmlValidator does not resolve note-body refs, so the
        // structural check below — not the schema check — is the oracle for this gap.)
        var left = new WmlDocument("left.docx", BuildNoteInNoteDoc());
        var right = new WmlDocument("right.docx", ReplaceWord(left.DocumentByteArray, "Alpha", "Gamma"));

        var result = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        Assert.Equal(new List<string>(), AllUnresolvedFootnoteRefs(result));
        Assert.Equal(new List<string>(), AllUnresolvedFootnoteRefs(accepted));
        Assert.Equal(new List<string>(), AllUnresolvedFootnoteRefs(rejected));

        var ids = FootnoteIds(result);
        Assert.Equal(ids.Distinct().Count(), ids.Count);
    }

    /// <summary>A doc where footnote 2's BODY cites ENDNOTE 5 (a CROSS-KIND note-in-note reference), and
    /// endnote 5 is ALSO referenced from the document body — so the endnote renumber pass remaps endnote 5's
    /// definition id, and the nested endnote reference inside the FOOTNOTES part must be remapped too or it
    /// dangles. The same-kind nested remap (footnote-cites-footnote) cannot catch this: each kind's pass only
    /// sweeps its OWN part for nested references.</summary>
    private static byte[] BuildCrossKindNoteInNoteDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn, $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                // footnote 2's body references ENDNOTE 5 (cross-kind note-in-note)
                "<w:footnote w:id=\"2\"><w:p><w:r><w:t xml:space=\"preserve\">Footnote citing an endnote </w:t></w:r><w:r><w:endnoteReference w:id=\"5\"/></w:r></w:p></w:footnote>" +
                "</w:footnotes>");
            var en = main.AddNewPart<EndnotesPart>();
            WritePartXml(en, $"<w:endnotes xmlns:w=\"{W}\">" +
                "<w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>" +
                "<w:endnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>" +
                // endnote 8's body references FOOTNOTE 2 (the reverse cross-kind nesting)
                "<w:endnote w:id=\"5\"><w:p><w:r><w:t xml:space=\"preserve\">The endnote, also body-cited.</w:t></w:r></w:p></w:endnote>" +
                "<w:endnote w:id=\"8\"><w:p><w:r><w:t xml:space=\"preserve\">Endnote citing a footnote </w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p></w:endnote>" +
                "</w:endnotes>");
            WritePartXml(main, $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Alpha cites the footnote.</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Beta cites the endnote.</w:t></w:r><w:r><w:endnoteReference w:id=\"5\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Gamma cites the other endnote.</w:t></w:r><w:r><w:endnoteReference w:id=\"8\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return ms.ToArray();
    }

    /// <summary>Every note reference of <paramref name="refLocal"/> kind ANYWHERE (document body, footnotes
    /// part, endnotes part) that does NOT resolve to exactly one definition of <paramref name="defLocal"/> kind
    /// in its part. The cross-kind generalization of <see cref="AllUnresolvedFootnoteRefs"/>.</summary>
    private static List<string> AllUnresolvedNoteRefs(WmlDocument doc, string refLocal, string defLocal, bool defsInFootnotesPart)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var fnRoot = main.FootnotesPart?.GetXDocument().Root;
        var enRoot = main.EndnotesPart?.GetXDocument().Root;
        var defRoot = defsInFootnotesPart ? fnRoot : enRoot;
        var defCounts = (defRoot?.Elements(System.Xml.Linq.XName.Get(defLocal, W))
                            .Where(e => e.Attribute(System.Xml.Linq.XName.Get("type", W)) == null)
                            .Select(e => (string?)e.Attribute(System.Xml.Linq.XName.Get("id", W)))
                            .Where(x => x != null).Select(x => x!) ?? Enumerable.Empty<string>())
                        .GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var refs = new List<System.Xml.Linq.XElement>();
        var body = main.GetXDocument().Root?.Element(System.Xml.Linq.XName.Get("body", W));
        if (body != null) refs.AddRange(body.Descendants(System.Xml.Linq.XName.Get(refLocal, W)));
        if (fnRoot != null) refs.AddRange(fnRoot.Descendants(System.Xml.Linq.XName.Get(refLocal, W)));
        if (enRoot != null) refs.AddRange(enRoot.Descendants(System.Xml.Linq.XName.Get(refLocal, W)));
        return refs.Select(r => (string?)r.Attribute(System.Xml.Linq.XName.Get("id", W)))
                   .Where(id => id == null || !defCounts.TryGetValue(id, out var n) || n != 1)
                   .Select(id => $"{refLocal}:{id ?? "(null)"}").ToList();
    }

    [Fact]
    public void CrossKindNestedNoteReference_ToRenumberedNote_StaysResolvable()
    {
        // An ENDNOTE reference nested inside a FOOTNOTE body (and a footnote ref inside an endnote body) must
        // survive the body-reference renumber that remaps the cited definitions' ids. Before the fix, each
        // kind's renumber pass swept only its OWN part for nested references, so the cross-kind nested ref kept
        // the stale id while its definition moved — dangling on compare/accept/reject.
        var left = new WmlDocument("left.docx", BuildCrossKindNoteInNoteDoc());
        var right = new WmlDocument("right.docx", ReplaceWord(left.DocumentByteArray, "Alpha", "Delta"));

        var result = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        foreach (var doc in new[] { result, accepted, rejected })
        {
            Assert.Equal(new List<string>(), AllUnresolvedNoteRefs(doc, "endnoteReference", "endnote", defsInFootnotesPart: false));
            Assert.Equal(new List<string>(), AllUnresolvedNoteRefs(doc, "footnoteReference", "footnote", defsInFootnotesPart: true));
        }
    }

    [Fact]
    public void DeletingReferenceToGappedIdFootnote_KeepsStructureResolvable()
    {
        var left = new WmlDocument("left.docx", BuildDocWithGappedFootnoteIds());
        var right = new WmlDocument("right.docx", RewriteFiveReferencingParagraph(left.DocumentByteArray));

        var result = DocxDiff.Compare(left, right);

        // Unique footnote definition ids.
        var ids = FootnoteIds(result);
        Assert.Equal(ids.Distinct().Count(), ids.Count);

        // No NEW schema errors vs the left.
        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0, $"schema errors: {string.Join(" | ", newErrors.Take(5))}");

        // id → ref → text round-trip: accept ≡ right (only fn1 referenced), reject ≡ left (fn1 + fn5).
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(ReferencedFootnoteTexts(right), ReferencedFootnoteTexts(accepted));
        Assert.Equal(ReferencedFootnoteTexts(left), ReferencedFootnoteTexts(rejected));
    }

    // ---- comment-anchor round-trip (a commented paragraph that is EDITED) ----------------------

    private static byte[] BuildCommentedDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(new RunPropertiesDefault(
                new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var cp = main.AddNewPart<WordprocessingCommentsPart>();
            WritePartXml(cp, $"<w:comments xmlns:w=\"{W}\">" +
                "<w:comment w:id=\"0\" w:author=\"Reviewer\" w:date=\"2020-01-01T00:00:00Z\" w:initials=\"R\"><w:p><w:r><w:t xml:space=\"preserve\">The reviewer note.</w:t></w:r></w:p></w:comment>" +
                "</w:comments>");
            WritePartXml(main, $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t xml:space=\"preserve\">The commented sentence here.</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>" +
                "<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"0\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">A plain trailing paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return ms.ToArray();
    }

    /// <summary>The body must carry exactly ONE balanced comment anchor — one commentRangeStart, one
    /// commentRangeEnd, one commentReference, all sharing an id that RESOLVES to a comments.xml definition.
    /// (An extra orphaned definition in the comments part is schema-valid, exactly as a reference-deleted
    /// footnote leaves an orphaned definition.)</summary>
    private static bool OneResolvedCommentAnchor(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.GetXDocument().Root?.Element(System.Xml.Linq.XName.Get("body", W));
        List<string?> Ids(string n) => body?.Descendants(System.Xml.Linq.XName.Get(n, W))
            .Select(e => (string?)e.Attribute(System.Xml.Linq.XName.Get("id", W))).ToList() ?? new();
        var rs = Ids("commentRangeStart"); var re = Ids("commentRangeEnd"); var rf = Ids("commentReference");
        if (rs.Count != 1 || re.Count != 1 || rf.Count != 1) return false;
        if (rs[0] != re[0] || re[0] != rf[0]) return false;
        var defIds = main.WordprocessingCommentsPart?.Comments?.Elements<Comment>()
            .Select(c => c.Id?.Value).ToHashSet() ?? new HashSet<string?>();
        return defIds.Contains(rs[0]);
    }

    [Fact]
    public void EditingCommentedParagraph_PreservesCommentAnchorsOnRoundTrip()
    {
        // Editing a word inside a commented paragraph must not orphan the comment: accept ≡ right and
        // reject ≡ left must each carry exactly ONE balanced comment anchor (rangeStart + rangeEnd + reference)
        // resolving to a comments.xml definition. (The IR drops comment markers from paragraphs, so the fine
        // token-diff path loses them; the conservative whole-block fallback preserves them, and a comment-id
        // dedup pass keeps the duplicated del/ins copies schema-valid.)
        var left = new WmlDocument("left.docx", BuildCommentedDoc());
        var right = new WmlDocument("right.docx", ReplaceWord(left.DocumentByteArray, "commented", "annotated"));

        var result = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        Assert.True(OneResolvedCommentAnchor(accepted), "accept must keep one resolved comment anchor (right)");
        Assert.True(OneResolvedCommentAnchor(rejected), "reject must keep one resolved comment anchor (left)");

        // No NEW schema errors from the Compare output (the duplicated del/ins comment ids are deduped).
        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0, $"schema errors: {string.Join(" | ", newErrors.Take(5))}");
    }

    // ---- helpers (mirrors DocxDiffScenarioTests, scoped to footnotes) --------------------------

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    private static List<long> FootnoteIds(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var footnotes = w.MainDocumentPart?.FootnotesPart?.Footnotes;
        return footnotes is null
            ? new List<long>()
            : footnotes.Elements<Footnote>().Where(f => f.Id is not null).Select(f => f.Id!.Value).ToList();
    }

    private static List<string> ReferencedFootnoteTexts(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var defText = new Dictionary<long, string>();
        foreach (var f in main.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>())
            if (f.Id is not null)
                defText[f.Id.Value] = string.Concat(f.Descendants<Text>().Select(t => t.Text));
        var seq = new List<string>();
        foreach (var r in main.Document?.Body?.Descendants<FootnoteReference>() ?? Enumerable.Empty<FootnoteReference>())
        {
            var id = r.Id?.Value;
            seq.Add(id is not null && defText.TryGetValue(id.Value, out var t) ? t : "(unresolved)");
        }
        return seq;
    }

    private static HashSet<string> SchemaErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        return validator.Validate(w).Select(e => $"{e.Id}@{e.Path?.XPath}: {e.Description}").ToHashSet();
    }
}
