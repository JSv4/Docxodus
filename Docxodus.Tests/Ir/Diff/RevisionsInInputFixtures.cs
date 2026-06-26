#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using WordType = DocumentFormat.OpenXml.WordprocessingDocumentType;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Shared fixtures + cross-scope revision inspection helpers for the "pre-existing tracked changes in input"
/// campaign — used by both the DEFAULT-behavior characterization suite
/// (<see cref="RevisionsInInputDefaultTests"/>) and the <c>PreAcceptInputRevisions</c> flag suite
/// (<see cref="PreAcceptInputRevisionsTests"/>). Centralizes the multi-scope revision-bearing document
/// builder and the all-parts revision/author/text projections so the two suites assert against the same shapes.
/// </summary>
internal static class RevisionsInInputFixtures
{
    public const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public const string Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace Wn = Wns;

    public static readonly XName[] RevisionNames =
    {
        Wn + "ins", Wn + "del", Wn + "moveFrom", Wn + "moveTo",
        Wn + "moveFromRangeStart", Wn + "moveFromRangeEnd",
        Wn + "moveToRangeStart", Wn + "moveToRangeEnd", Wn + "rPrChange", Wn + "pPrChange",
    };

    private static string Ins(string author, string text) =>
        $"<w:ins w:id=\"900\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\">" +
        $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:ins>";

    private static string Del(string author, string text) =>
        $"<w:del w:id=\"901\" w:author=\"{author}\" w:date=\"2020-01-01T00:00:00Z\">" +
        $"<w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:del>";

    private static string ReservedNote(string type, int id) =>
        $"<w:footnote w:type=\"{type}\" w:id=\"{id}\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>";

    private static void WritePart(OpenXmlPart part, string xml)
    {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = new StreamWriter(s);
        w.Write(xml);
    }

    /// <summary>
    /// A document that carries UN-accepted, pre-existing tracked-change markup in three scopes: the body
    /// (inline <c>w:ins</c>/<c>w:del</c> by <paramref name="bodyAuthor"/>), a default header, and the real
    /// footnote (id=1) — header and footnote insertions authored by <paramref name="priorAuthor"/>. Two
    /// documents built with the same header/footnote text are IDENTICAL in those carried-over scopes (so they
    /// exercise the leak path, not a real diff); the body text varies so the diff has work to do.
    /// </summary>
    public static WmlDocument MultiScopeRevisionDoc(
        string bodyLead, string bodyAuthor, string bodyInsText, string bodyDelText,
        string priorAuthor, string headerInsText, string footnoteInsText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            const string headerRelId = "rIdHdr1";
            var headerPart = main.AddNewPart<HeaderPart>(headerRelId);
            WritePart(headerPart,
                $"<w:hdr xmlns:w=\"{Wns}\"><w:p><w:r><w:t xml:space=\"preserve\">Header </w:t></w:r>" +
                Ins(priorAuthor, headerInsText) + "</w:p></w:hdr>");

            var fnPart = main.AddNewPart<FootnotesPart>();
            WritePart(fnPart,
                $"<w:footnotes xmlns:w=\"{Wns}\">" +
                ReservedNote("separator", -1) + ReservedNote("continuationSeparator", 0) +
                "<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">Note </w:t></w:r>" +
                Ins(priorAuthor, footnoteInsText) + "</w:p></w:footnote></w:footnotes>");

            WritePart(main,
                $"<w:document xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\"><w:body>" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{bodyLead} </w:t></w:r>" +
                Ins(bodyAuthor, bodyInsText) + Del(bodyAuthor, bodyDelText) +
                "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                $"<w:sectPr><w:headerReference w:type=\"default\" r:id=\"{headerRelId}\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>");
        }
        return new WmlDocument("multiscope.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose COMMENT definition (id 0, referenced from the body) carries an UN-accepted tracked
    /// insertion authored by <paramref name="commentRevAuthor"/>. The body text varies via
    /// <paramref name="bodyText"/> so two such documents differ for the diff while the commented range + comment
    /// definition are identical (a carry-over). Used to pin the one scope the flag does NOT clean:
    /// <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/> does not process the comments part.
    /// </summary>
    public static WmlDocument CommentWithRevisionDoc(string bodyText, string commentRevAuthor)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            WritePart(commentsPart,
                $"<w:comments xmlns:w=\"{Wns}\">" +
                "<w:comment w:id=\"0\" w:author=\"Commenter\" w:initials=\"C\" w:date=\"2020-01-01T00:00:00Z\">" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Comment </w:t></w:r>" +
                Ins(commentRevAuthor, "CMTEDIT") + "</w:p></w:comment></w:comments>");

            WritePart(main,
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                $"<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t xml:space=\"preserve\">{bodyText}</w:t></w:r>" +
                "<w:commentRangeEnd w:id=\"0\"/>" +
                "<w:r><w:commentReference w:id=\"0\"/></w:r></w:p>" +
                "</w:body></w:document>");
        }
        return new WmlDocument("comment-rev.docx", ms.ToArray());
    }

    public static IEnumerable<OpenXmlPart> ContentParts(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart!;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart != null) yield return main.FootnotesPart;
        if (main.EndnotesPart != null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart != null) yield return main.WordprocessingCommentsPart;
        if (main.StyleDefinitionsPart != null) yield return main.StyleDefinitionsPart;
        if (main.GlossaryDocumentPart != null) yield return main.GlossaryDocumentPart;
    }

    /// <summary>
    /// A document whose GLOSSARY (building-blocks / AutoText) document part carries an UN-accepted tracked
    /// insertion authored by <paramref name="glossaryRevAuthor"/>. The body varies via <paramref name="bodyText"/>.
    /// Used to pin a SECOND scope the flag does not clean: <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/>
    /// never processes the <c>GlossaryDocumentPart</c>, and the markup renderer clones the whole left package
    /// (glossary included) verbatim.
    /// </summary>
    public static WmlDocument GlossaryWithRevisionDoc(string bodyText, string glossaryRevAuthor)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var glossary = main.AddNewPart<GlossaryDocumentPart>();
            WritePart(glossary,
                $"<w:glossaryDocument xmlns:w=\"{Wns}\"><w:docParts><w:docPart>" +
                "<w:docPartPr><w:name w:val=\"GB1\"/>" +
                "<w:category><w:name w:val=\"General\"/><w:gallery w:val=\"placeholder\"/></w:category>" +
                "</w:docPartPr><w:docPartBody>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Block </w:t></w:r>" +
                Ins(glossaryRevAuthor, "GLOSSEDIT") + "</w:p>" +
                "</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>");

            WritePart(main,
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{bodyText}</w:t></w:r></w:p>" +
                "</w:body></w:document>");
        }
        return new WmlDocument("glossary-rev.docx", ms.ToArray());
    }

    /// <summary>The distinct <c>w:author</c> values on every revision element across every content part.</summary>
    public static HashSet<string> RevisionAuthorsAllScopes(WmlDocument d)
    {
        var authors = new HashSet<string>(StringComparer.Ordinal);
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        foreach (var part in ContentParts(doc))
        {
            XDocument xd;
            using (var s = part.GetStream(FileMode.Open, FileAccess.Read)) xd = XDocument.Load(s);
            foreach (var el in xd.Descendants().Where(e => RevisionNames.Contains(e.Name)))
            {
                var a = (string?)el.Attribute(Wn + "author");
                if (a != null) authors.Add(a);
            }
        }
        return authors;
    }

    /// <summary>The set of distinct revision element names present across every content part.</summary>
    public static HashSet<XName> AllRevisionElementNames(WmlDocument d)
    {
        var names = new HashSet<XName>();
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        foreach (var part in ContentParts(doc))
        {
            XDocument xd;
            using (var s = part.GetStream(FileMode.Open, FileAccess.Read)) xd = XDocument.Load(s);
            foreach (var el in xd.Descendants().Where(e => RevisionNames.Contains(e.Name)))
                names.Add(el.Name);
        }
        return names;
    }

    /// <summary>The concatenated <c>w:t</c> text of one named part scope (e.g. only the header parts).</summary>
    private static string PartText(OpenXmlPart part)
    {
        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
        var xd = XDocument.Load(s);
        return string.Concat(xd.Descendants(Wn + "t").Select(t => t.Value));
    }

    /// <summary>Body + header + footer + footnote + endnote VISIBLE text (<c>w:t</c> only — i.e. the post-accept
    /// content; callers pass accepted documents when comparing accepted views), each scope on its own line in a
    /// stable order. An all-scopes text projection for round-trip oracles.</summary>
    public static string AllScopesText(WmlDocument d)
    {
        var sb = new StringBuilder();
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        sb.Append("BODY:").Append(PartText(main)).Append('\n');
        foreach (var h in main.HeaderParts) sb.Append("HEADER:").Append(PartText(h)).Append('\n');
        foreach (var f in main.FooterParts) sb.Append("FOOTER:").Append(PartText(f)).Append('\n');
        if (main.FootnotesPart != null) sb.Append("FOOTNOTES:").Append(PartText(main.FootnotesPart)).Append('\n');
        if (main.EndnotesPart != null) sb.Append("ENDNOTES:").Append(PartText(main.EndnotesPart)).Append('\n');
        return sb.ToString();
    }

    /// <summary>The concatenated header <c>w:t</c> text (across all header parts) of a document.</summary>
    public static string HeaderText(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        return string.Concat(doc.MainDocumentPart!.HeaderParts.Select(PartText));
    }

    public static WmlDocument AcceptAll(WmlDocument d) => RevisionProcessor.AcceptRevisions(d);
}
