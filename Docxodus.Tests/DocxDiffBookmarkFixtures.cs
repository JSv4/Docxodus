#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WmlDocument = Docxodus.WmlDocument;

namespace DocxodusDiffParityFixtures;

/// <summary>
/// A synthetic corpus of bookmark / internal-cross-reference shapes a real legal contract exposes, used by
/// <c>DocxDiffBookmarkStructureTests</c> to assert <see cref="Docxodus.DocxDiff"/> preserves bookmark
/// id↔name↔reference integrity across edits to bookmark-bearing / bookmark-referencing paragraphs.
/// <para>Each scenario is a deliberately NON-coincidental shape (gapped ids, multi-bookmark paragraphs,
/// multi-paragraph ranges, hyperlink+REF on one bookmark, a TOC with <c>_Toc</c> bookmarks, and the
/// del/ins renumber-collision shape).</para>
/// </summary>
internal static class DocxDiffBookmarkFixtures
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // ---- public API -----------------------------------------------------------------------------

    public static IEnumerable<string> Names() => Catalog.Keys;

    public static (WmlDocument Left, WmlDocument Right) Build(string scenario)
    {
        if (!Catalog.TryGetValue(scenario, out var pair))
            throw new ArgumentException($"unknown bookmark scenario '{scenario}'", nameof(scenario));
        var (leftBody, mutate) = pair;
        var left = Doc(leftBody);
        var right = mutate(left);
        return (new WmlDocument($"{scenario}.left.docx", left), new WmlDocument($"{scenario}.right.docx", right));
    }

    // ---- the catalog ----------------------------------------------------------------------------
    // Each entry: body XML (left) + a mutation that produces the right bytes from the left bytes.

    private static readonly IReadOnlyDictionary<string, (string Body, Func<byte[], byte[]> Mutate)> Catalog =
        new Dictionary<string, (string, Func<byte[], byte[]>)>
        {
            // 1. Gapped (non-sequential) ids; edit the text AFTER one bookmark (fine path, should survive).
            ["gapped-ids-edit-after"] = (
                BkPara("7", "_Ref_Defs", "Definitions", " govern this Agreement.") +
                BkPara("20", "_Ref_Close", "Closing", " occurs on the date herein.") +
                RefPara("_Ref_Defs") + RefPara("_Ref_Close"),
                b => Edit(b, "govern this Agreement", "govern all of this Agreement")),

            // 2. Edit the BOOKMARKED TEXT itself (boundary zero-width drop risk).
            ["edit-bookmarked-text"] = (
                BkPara("7", "_Ref_Sec21", "Section 2.1", " for definitions.") + RefPara("_Ref_Sec21"),
                b => Edit(b, "Section 2.1", "Section 2.1(a)")),

            // 3. Whole-paragraph rewrite of a bookmark-bearing paragraph (both-boundary drop risk).
            ["whole-para-rewrite"] = (
                BkPara("7", "_Ref_Sec21", "Section 2.1 Definitions", "") + RefPara("_Ref_Sec21"),
                b => ReplaceWholePara(b, "Section 2.1 Definitions",
                    "An entirely different sentence about widgets, gadgets and gizmos.")),

            // 4. A paragraph carrying MULTIPLE bookmarks, edited.
            ["multi-bookmark-para"] = (
                "<w:p>" +
                Bk("3", "_Ref_A", "Alpha") + "<w:r><w:t xml:space=\"preserve\"> and </w:t></w:r>" +
                Bk("5", "_Ref_B", "Beta") + "<w:r><w:t xml:space=\"preserve\"> are defined terms.</w:t></w:r></w:p>" +
                RefPara("_Ref_A") + RefPara("_Ref_B"),
                b => Edit(b, "are defined terms", "are the defined terms")),

            // 5. Bookmark RANGE spanning paragraphs; edit the START paragraph (leading-zero-width drop risk).
            ["multi-para-range-edit-start"] = (
                "<w:p><w:bookmarkStart w:id=\"7\" w:name=\"_Ref_Clause\"/><w:r><w:t>Clause opening paragraph here.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Middle paragraph inside the clause range.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Closing paragraph of the clause.</w:t></w:r><w:bookmarkEnd w:id=\"7\"/></w:p>" +
                RefPara("_Ref_Clause"),
                b => Edit(b, "Clause opening paragraph here", "Clause opening paragraph here, as amended")),

            // 6. Bookmark RANGE spanning paragraphs; edit the MIDDLE paragraph.
            ["multi-para-range-edit-middle"] = (
                "<w:p><w:bookmarkStart w:id=\"7\" w:name=\"_Ref_Clause\"/><w:r><w:t>Clause opening paragraph here.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Middle paragraph inside the clause range.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Closing paragraph of the clause.</w:t></w:r><w:bookmarkEnd w:id=\"7\"/></w:p>" +
                RefPara("_Ref_Clause"),
                b => Edit(b, "Middle paragraph inside the clause range",
                    "Middle paragraph inside the clause range with new wording")),

            // 7. Bookmark RANGE spanning paragraphs; whole-rewrite the START paragraph (churned endpoint).
            ["multi-para-range-rewrite-start"] = (
                "<w:p><w:bookmarkStart w:id=\"7\" w:name=\"_Ref_Clause\"/><w:r><w:t>Clause opening paragraph here.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Middle paragraph inside the clause range.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Closing paragraph of the clause.</w:t></w:r><w:bookmarkEnd w:id=\"7\"/></w:p>" +
                RefPara("_Ref_Clause"),
                b => ReplaceWholePara(b, "Clause opening paragraph here",
                    "A wholly rewritten opening line with no resemblance to the original text.")),

            // 8. A hyperlink AND a REF field both target ONE bookmark; edit the bookmark paragraph.
            ["hyperlink-and-ref-one-bookmark"] = (
                BkPara("7", "_Ref_Term", "Defined Term", " has the meaning set forth herein.") +
                "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
                "<w:r><w:instrText xml:space=\"preserve\"> REF _Ref_Term \\h </w:instrText></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>Defined Term</w:t></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
                "<w:r><w:t xml:space=\"preserve\"> or click </w:t></w:r>" +
                "<w:hyperlink w:anchor=\"_Ref_Term\"><w:r><w:t>here</w:t></w:r></w:hyperlink>" +
                "<w:r><w:t>.</w:t></w:r></w:p>",
                b => Edit(b, "has the meaning set forth herein", "has the meaning set forth below")),

            // 9. The whole-block-bail renumber-collision shape: bookmark + comment in one paragraph, edited.
            ["renumber-collision-comment"] = (
                "<w:p><w:commentRangeStart w:id=\"0\"/>" +
                "<w:r><w:t xml:space=\"preserve\">The </w:t></w:r>" +
                Bk("7", "_Ref_Sec21", "Section 2.1") +
                "<w:r><w:t xml:space=\"preserve\"> governs definitions.</w:t></w:r>" +
                "<w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>" +
                RefPara("_Ref_Sec21"),
                b => Edit(b, "governs definitions", "governs all definitions")),

            // 10. PAGEREF + NOTEREF style fields targeting bookmarks; edit a bookmark paragraph.
            ["pageref-and-noteref"] = (
                BkPara("7", "_Ref_PageA", "Schedule A", " lists the parties.") +
                BkPara("9", "_Ref_NoteB", "Note B", " qualifies the schedule.") +
                "<w:p><w:r><w:t xml:space=\"preserve\">Page </w:t></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
                "<w:r><w:instrText xml:space=\"preserve\"> PAGEREF _Ref_PageA \\h </w:instrText></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>1</w:t></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
                "<w:r><w:t xml:space=\"preserve\">; see </w:t></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
                "<w:r><w:instrText xml:space=\"preserve\"> NOTEREF _Ref_NoteB \\h </w:instrText></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>2</w:t></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:t>.</w:t></w:r></w:p>",
                b => Edit(b, "lists the parties", "lists all the parties")),

            // 11. fldSimple (not fldChar) REF; edit the bookmark paragraph.
            ["fldsimple-ref"] = (
                BkPara("7", "_Ref_Sec21", "Section 2.1", " is controlling.") +
                "<w:p><w:r><w:t xml:space=\"preserve\">As stated in </w:t></w:r>" +
                "<w:fldSimple w:instr=\" REF _Ref_Sec21 \\h \"><w:r><w:t>Section 2.1</w:t></w:r></w:fldSimple>" +
                "<w:r><w:t>.</w:t></w:r></w:p>",
                b => Edit(b, "is controlling", "is the controlling provision")),

            // 12. A TOC with _Toc bookmarks on headings; edit a heading's text.
            ["toc-toc-bookmarks"] = (
                // TOC field referencing the headings.
                "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
                "<w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h </w:instrText></w:r>" +
                "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
                "<w:hyperlink w:anchor=\"_Toc001\"><w:r><w:t>1. Definitions</w:t></w:r></w:hyperlink>" +
                "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>" +
                "<w:p><w:bookmarkStart w:id=\"100\" w:name=\"_Toc001\"/><w:r><w:t>1. Definitions</w:t></w:r><w:bookmarkEnd w:id=\"100\"/></w:p>" +
                "<w:p><w:bookmarkStart w:id=\"101\" w:name=\"_Toc002\"/><w:r><w:t>2. Closing</w:t></w:r><w:bookmarkEnd w:id=\"101\"/></w:p>",
                b => Edit(b, "1. Definitions", "1. Definitions and Interpretation")),

            // 13. Bookmark whose ONLY reference is DELETED, but the bookmark lingers (reconcile case).
            //     accept ≡ right must keep the (now unreferenced) bookmark and drop the ref.
            ["ref-deleted-bookmark-lingers"] = (
                BkPara("7", "_Ref_Sec21", "Section 2.1", " has definitions.") +
                "<w:p><w:r><w:t xml:space=\"preserve\">As defined in </w:t></w:r>" +
                "<w:fldSimple w:instr=\" REF _Ref_Sec21 \\h \"><w:r><w:t>Section 2.1</w:t></w:r></w:fldSimple>" +
                "<w:r><w:t>.</w:t></w:r></w:p>",
                b => DeleteParaContaining(b, "As defined in")),

            // 15. A run carrying a w:noBreakHyphen (1-char in the IR, zero-width in the slicer) sits in the
            //     UNCHANGED tail of an edited bookmark paragraph — guards the char-accounting desync that dropped
            //     the adjacent "I" of "Company‑Controlled Intellectual" on reject.
            ["nobreakhyphen-in-edited-para"] = (
                "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>" +
                Bk("7", "_Ref_Sec21", "Section 2.1") +
                "<w:r><w:t xml:space=\"preserve\"> governs Company</w:t></w:r>" +
                "<w:r><w:noBreakHyphen/><w:t>Controlled Intellectual Property terms.</w:t></w:r></w:p>" +
                RefPara("_Ref_Sec21"),
                b => Edit(b, "See ", "See, per the recitals, ")),

            // 14. Bookmark DELETED while a reference SURVIVES (right itself has a dangling ref — we mirror it).
            ["bookmark-deleted-ref-survives"] = (
                BkPara("7", "_Ref_Sec21", "Section 2.1", " has definitions.") +
                "<w:p><w:r><w:t xml:space=\"preserve\">As defined in </w:t></w:r>" +
                "<w:fldSimple w:instr=\" REF _Ref_Sec21 \\h \"><w:r><w:t>Section 2.1</w:t></w:r></w:fldSimple>" +
                "<w:r><w:t>.</w:t></w:r></w:p>",
                b => DeleteParaContaining(b, "has definitions")),
        };

    // ---- body fragment helpers ------------------------------------------------------------------

    /// <summary>A bare bookmark wrapping <paramref name="text"/>: start, run, end.</summary>
    private static string Bk(string id, string name, string text) =>
        $"<w:bookmarkStart w:id=\"{id}\" w:name=\"{name}\"/><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r><w:bookmarkEnd w:id=\"{id}\"/>";

    /// <summary>A paragraph: "See " + bookmark(text) + trailing.</summary>
    private static string BkPara(string id, string name, string text, string trailing) =>
        $"<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>{Bk(id, name, text)}" +
        (trailing.Length > 0 ? $"<w:r><w:t xml:space=\"preserve\">{trailing}</w:t></w:r>" : "") + "</w:p>";

    /// <summary>A REF-field + internal-hyperlink paragraph targeting <paramref name="name"/>.</summary>
    private static string RefPara(string name) =>
        "<w:p><w:r><w:t xml:space=\"preserve\">As defined in </w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        $"<w:r><w:instrText xml:space=\"preserve\"> REF {name} \\h </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>X</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
        "<w:r><w:t xml:space=\"preserve\"> or </w:t></w:r>" +
        $"<w:hyperlink w:anchor=\"{name}\"><w:r><w:t>that section</w:t></w:r></w:hyperlink>" +
        "<w:r><w:t>.</w:t></w:r></w:p>";

    // ---- document builder + mutators ------------------------------------------------------------

    private static byte[] Doc(string bodyInner)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            if (bodyInner.Contains("commentReference"))
            {
                var cp = main.AddNewPart<WordprocessingCommentsPart>();
                WritePartXml(cp, $"<w:comments xmlns:w=\"{W}\"><w:comment w:id=\"0\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\"><w:p><w:r><w:t>a note</w:t></w:r></w:p></w:comment></w:comments>");
            }
            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\"><w:body>{bodyInner}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return ms.ToArray();
    }

    private static byte[] Edit(byte[] left, string find, string repl) =>
        Mutate(left, body =>
        {
            foreach (var t in body.Descendants<Text>())
                if (t.Text.Contains(find)) { t.Text = t.Text.Replace(find, repl); return; }
            throw new InvalidOperationException($"text '{find}' not found");
        });

    private static byte[] ReplaceWholePara(byte[] left, string contains, string newText) =>
        Mutate(left, body =>
        {
            var p = body.Elements<Paragraph>().First(x =>
                string.Concat(x.Descendants<Text>().Select(t => t.Text)).Contains(contains));
            // keep the bookmark markers; replace only the run text content.
            foreach (var r in p.Elements<Run>().ToList()) r.Remove();
            // re-insert a single run after any leading bookmarkStart (preserve marker positions roughly).
            var bkStart = p.Elements().FirstOrDefault(e => e.LocalName == "bookmarkStart");
            var run = new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
            if (bkStart != null) bkStart.InsertAfterSelf(run); else p.AppendChild(run);
        });

    private static byte[] DeleteParaContaining(byte[] left, string contains) =>
        Mutate(left, body =>
        {
            var p = body.Elements<Paragraph>().First(x =>
                string.Concat(x.Descendants<Text>().Select(t => t.Text)).Contains(contains));
            p.Remove();
        });

    private static byte[] Mutate(byte[] left, Action<Body> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
            mutate(doc.MainDocumentPart!.Document!.Body!);
        return ms.ToArray();
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }
}
