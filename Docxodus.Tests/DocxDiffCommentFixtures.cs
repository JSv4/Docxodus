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
/// A synthetic corpus of COMMENT shapes a real comment-dense contract exposes, used by
/// <c>DocxDiffCommentStructureTests</c> to assert <see cref="Docxodus.DocxDiff.Compare"/> preserves
/// comment id↔range↔reference↔definition integrity (and threaded <c>commentsExtended</c> reply links)
/// across edits to commented paragraphs — and renders them with FINE per-word markup rather than the
/// coarse whole-block bail.
/// <para>Each scenario is a deliberately NON-coincidental shape: multiple comments on one paragraph,
/// overlapping ranges, a range spanning paragraphs with one end edited, a comment whose anchored text is
/// itself edited (the del/ins comment-id collision), a threaded reply, a right-added comment, and a
/// comment anchored on deleted text.</para>
/// </summary>
internal static class DocxDiffCommentFixtures
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // ---- public API -----------------------------------------------------------------------------

    public static IEnumerable<string> Names() => Catalog.Keys;

    /// <summary>Shapes that are a PURE single-paragraph text edit (no paragraph add/delete), so the fine
    /// path must keep the body paragraph count equal to the right document's — the whole-block bail would
    /// duplicate the commented paragraph into a del-copy + ins-copy. Used by the fine-markup assertion.</summary>
    public static IEnumerable<string> SingleParaEditShapes() => new[]
    {
        "multi-comment-one-para", "overlapping-ranges", "edit-commented-text",
        "threaded-reply-anchor-edited", "right-added-comment",
    };

    public static (WmlDocument Left, WmlDocument Right) Build(string scenario)
    {
        if (!Catalog.TryGetValue(scenario, out var s))
            throw new ArgumentException($"unknown comment scenario '{scenario}'", nameof(scenario));
        var left = Doc(s.Body, s.Comments, s.CommentsEx);
        var right = s.Mutate(left);
        return (new WmlDocument($"{scenario}.left.docx", left), new WmlDocument($"{scenario}.right.docx", right));
    }

    private sealed record Scenario(string Body, string Comments, string? CommentsEx, Func<byte[], byte[]> Mutate);

    // ---- the catalog ----------------------------------------------------------------------------

    private static readonly IReadOnlyDictionary<string, Scenario> Catalog =
        new Dictionary<string, Scenario>
        {
            // 1. ONE paragraph carrying TWO distinct comments; edit text OUTSIDE both ranges (both survive, fine).
            ["multi-comment-one-para"] = new(
                "<w:p>" +
                Open("0") + Run("The first phrase ") + Close("0") +
                Run("links the ") +
                Open("1") + Run("second phrase") + Close("1") +
                Run(" plainly and clearly here.") + "</w:p>",
                Cmt("0", "Alice", "comment on the first phrase") +
                Cmt("1", "Bob", "comment on the second phrase"),
                null,
                b => Edit(b, "clearly", "obviously")),

            // 2. OVERLAPPING ranges (id 0 covers "alpha beta", id 1 covers "beta gamma"); edit the shared word.
            ["overlapping-ranges"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">Start </w:t></w:r>" +
                Open("0") + Run("alpha ") + Open("1") + Run("beta") + Close("0") + Run(" gamma") + Close("1") +
                RefRun("0") + RefRun("1") +
                "<w:r><w:t xml:space=\"preserve\"> end of line.</w:t></w:r></w:p>",
                Cmt("0", "Alice", "covers alpha and beta") + Cmt("1", "Bob", "covers beta and gamma"),
                null,
                b => Edit(b, "beta", "BETA")),

            // 3. A comment RANGE spanning paragraphs; edit the START paragraph (one range endpoint churned).
            ["range-spans-paras-edit-start"] = new(
                "<w:p>" + Open("0") + Run("Opening clause text here.") + "</w:p>" +
                "<w:p>" + Run("Closing clause text here.") + Close("0") + RefRun("0") + "</w:p>",
                Cmt("0", "Alice", "spans the whole clause"),
                null,
                b => Edit(b, "Opening clause text here", "Opening clause text here, as amended")),

            // 4. A comment whose ANCHORED TEXT ITSELF is edited — the del/ins comment-id collision (dedup B).
            ["edit-commented-text"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">The </w:t></w:r>" +
                Open("0") + Run("disputed") + Close("0") + RefRun("0") +
                "<w:r><w:t xml:space=\"preserve\"> clause stands.</w:t></w:r></w:p>",
                Cmt("0", "Alice", "is this term right?"),
                null,
                b => Edit(b, "disputed", "contested")),

            // 5. A commented paragraph that is WHOLLY rewritten (every run replaced).
            ["whole-rewrite-commented-para"] = new(
                "<w:p>" + Open("0") + Run("The original sentence about widgets.") + Close("0") + RefRun("0") + "</w:p>",
                Cmt("0", "Alice", "note on the original"),
                null,
                b => ReplaceWholePara(b, "original sentence about widgets",
                    "A wholly different sentence about gadgets and gizmos.")),

            // 6. A THREADED reply (parent id 0 + reply id 1 via commentsExtended paraIdParent); edit text
            //    OUTSIDE the anchored region — the unchanged thread must be carried wholesale.
            ["threaded-reply-unchanged"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">Prefix words. </w:t></w:r>" +
                Open("0") + Run("anchored text") + Close("0") + RefRun("0") + RefRun("1") +
                "<w:r><w:t xml:space=\"preserve\"> trailing words here.</w:t></w:r></w:p>",
                Cmt("0", "Alice", "parent question", "00000001") +
                Cmt("1", "Bob", "child reply", "00000002"),
                CmtEx("00000001", null) + CmtEx("00000002", "00000001"),
                b => Edit(b, "trailing words here", "trailing words right here")),

            // 7. A THREADED reply where the ANCHORED text is edited (dedup B with threading preserved).
            ["threaded-reply-anchor-edited"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">Prefix. </w:t></w:r>" +
                Open("0") + Run("anchored") + Close("0") + RefRun("0") + RefRun("1") +
                "<w:r><w:t xml:space=\"preserve\"> suffix.</w:t></w:r></w:p>",
                Cmt("0", "Alice", "parent question", "00000001") +
                Cmt("1", "Bob", "child reply", "00000002"),
                CmtEx("00000001", null) + CmtEx("00000002", "00000001"),
                b => Edit(b, "anchored", "amended")),

            // 8. RIGHT adds a brand-new comment on an edited paragraph (forces right-definition merge).
            ["right-added-comment"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">A clause with no comment yet here.</w:t></w:r></w:p>",
                "", // no comments on the left
                null,
                b => AddCommentAround(b, "no comment yet", id: "0", author: "Carol",
                    text: "newly added review note", edit: ("clause", "provision"))),

            // 8b. RIGHT adds a comment whose definition paragraph carries a w14:paraId, with NO commentsExtended
            //     part on either side — guards the eager-create-empty-commentsExtended-part schema bug.
            ["right-added-comment-paraid"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">Another clause with no comment yet here.</w:t></w:r></w:p>",
                "",
                null,
                b => AddCommentAround(b, "no comment yet", id: "0", author: "Carol",
                    text: "review note with a paraId", edit: ("clause", "provision"), paraId: "5A5A5A5A")),

            // 9. A comment anchored on text that is wholly DELETED in the right (comment becomes reject-only).
            ["comment-on-deleted-text"] = new(
                "<w:p><w:r><w:t xml:space=\"preserve\">Keep this. </w:t></w:r>" +
                Open("0") + Run("Delete this obsolete clause.") + Close("0") + RefRun("0") +
                "<w:r><w:t xml:space=\"preserve\"> Keep that.</w:t></w:r></w:p>",
                Cmt("0", "Alice", "flag for deletion"),
                null,
                b => Edit(b, "Delete this obsolete clause.", "")),
        };

    // ---- body fragment helpers ------------------------------------------------------------------

    private static string Run(string text) =>
        $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Open(string id) => $"<w:commentRangeStart w:id=\"{id}\"/>";
    private static string Close(string id) => $"<w:commentRangeEnd w:id=\"{id}\"/>";

    /// <summary>The reference-bearing run that follows a comment range close (carries the CommentReference style).</summary>
    private static string RefRun(string id) =>
        $"<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"{id}\"/></w:r>";

    // ---- comment definition helpers -------------------------------------------------------------

    private static string Cmt(string id, string author, string text, string? paraId = null) =>
        $"<w:comment w:id=\"{id}\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\" w:initials=\"{author[0]}\">" +
        $"<w:p{(paraId != null ? $" w14:paraId=\"{paraId}\"" : "")}><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:comment>";

    private static string CmtEx(string paraId, string? paraIdParent) =>
        $"<w15:commentEx w15:paraId=\"{paraId}\"{(paraIdParent != null ? $" w15:paraIdParent=\"{paraIdParent}\"" : "")} w15:done=\"0\"/>";

    // ---- document builder + mutators ------------------------------------------------------------

    private static byte[] Doc(string bodyInner, string commentsInner, string? commentsExInner)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))),
                new Style(new StyleName { Val = "annotation reference" }) { Type = StyleValues.Character, StyleId = "CommentReference" });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            if (commentsInner.Length > 0)
            {
                var cp = main.AddNewPart<WordprocessingCommentsPart>();
                WritePartXml(cp, $"<w:comments xmlns:w=\"{W}\" xmlns:w14=\"{W14}\">{commentsInner}</w:comments>");
            }
            if (commentsExInner != null)
            {
                var ce = main.AddNewPart<WordprocessingCommentsExPart>();
                WritePartXml(ce, $"<w15:commentsEx xmlns:w=\"{W}\" xmlns:w15=\"{W15}\">{commentsExInner}</w15:commentsEx>");
            }
            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:w14=\"{W14}\"><w:body>{bodyInner}" +
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
            foreach (var r in p.Elements<Run>().Where(r => r.Descendants<Text>().Any()).ToList())
                r.Remove();
            // re-insert one run after the first commentRangeStart (preserve marker positions roughly).
            var open = p.Elements().FirstOrDefault(e => e.LocalName == "commentRangeStart");
            var run = new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
            if (open != null) open.InsertAfterSelf(run); else p.AddChild(run);
        });

    /// <summary>RIGHT-only: wrap a word with a NEW comment range + add its definition (forces the
    /// right-definition merge), and apply a small text edit so the paragraph is a Modify. An optional
    /// <paramref name="paraId"/> stamps the comment-definition paragraph's <c>w14:paraId</c> (with NO
    /// commentsExtended part) — the shape that guards the eager-create-empty-commentsExtended schema bug.</summary>
    private static byte[] AddCommentAround(byte[] left, string word, string id, string author, string text,
        (string find, string repl) edit, string? paraId = null) =>
        MutateDoc(left, main =>
        {
            var body = main.Document!.Body!;
            var p = body.Elements<Paragraph>().First(x =>
                string.Concat(x.Descendants<Text>().Select(t => t.Text)).Contains(word));
            // Insert the comment range around the whole paragraph's run content.
            var firstRun = p.Elements<Run>().First();
            var lastRun = p.Elements<Run>().Last();
            firstRun.InsertBeforeSelf(new CommentRangeStart { Id = id });
            var endMarker = new CommentRangeEnd { Id = id };
            lastRun.InsertAfterSelf(endMarker);
            endMarker.InsertAfterSelf(new Run(
                new RunProperties(new RunStyle { Val = "CommentReference" }), new CommentReference { Id = id }));
            // Add the definition to a (possibly new) comments part.
            var cp = main.WordprocessingCommentsPart ?? main.AddNewPart<WordprocessingCommentsPart>();
            using (var stream = cp.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false)))
                writer.Write($"<w:comments xmlns:w=\"{W}\" xmlns:w14=\"{W14}\">{Cmt(id, author, text, paraId)}</w:comments>");
            // Apply the small text edit.
            foreach (var t in body.Descendants<Text>())
                if (t.Text.Contains(edit.find)) { t.Text = t.Text.Replace(edit.find, edit.repl); break; }
        });

    private static byte[] Mutate(byte[] left, Action<Body> mutate) =>
        MutateDoc(left, main => mutate(main.Document!.Body!));

    private static byte[] MutateDoc(byte[] left, Action<MainDocumentPart> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
            mutate(doc.MainDocumentPart!);
        return ms.ToArray();
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }
}
