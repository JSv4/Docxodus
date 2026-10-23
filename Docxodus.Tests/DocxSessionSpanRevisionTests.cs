using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxSession.ReplaceTextAtSpan"/> inside and beside existing tracked revisions: the
/// span path the browser editor sends every keystroke through. Each test reads the paragraph back
/// as a compact shape — <c>"text"</c> for a run, <c>ins:Author[...]</c> for a revision envelope —
/// so the markup a reviewer would see is the assertion.
/// </summary>
public class DocxSessionSpanRevisionTests
{
    private const string Date = "2026-05-01T09:00:00Z";

    // ─── Deleting from the session author's own pending insertion (issue #1316) ───

    [Fact]
    public void Backspace_InOwnPendingInsertion_RemovesTheCharacter()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Parties "));
        Assert.True(session.ReplaceTextAtSpan(p, 8, 0, "bindingg").Success);

        var edit = session.ReplaceTextAtSpan(p, 15, 1, "");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Parties \" ins:Pat[\"binding\"]", Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:binding" }, Revisions(session));
        Assert.Equal("Parties binding", Accepted(session));
        Assert.Equal("Parties ", Rejected(session));
    }

    [Fact]
    public void Correction_InOwnPendingInsertion_StaysOneInsertion()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Fee: "));
        Assert.True(session.ReplaceTextAtSpan(p, 5, 0, "teh clause").Success);

        var edit = session.ReplaceTextAtSpan(p, 6, 2, "he");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Fee: \" ins:Pat[\"the clause\"]", Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:the clause" }, Revisions(session));
    }

    [Fact]
    public void DeletingAllOfOwnPendingInsertion_LeavesNoRevision()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Parties "));
        Assert.True(session.ReplaceTextAtSpan(p, 8, 0, "abc").Success);

        Assert.True(session.ReplaceTextAtSpan(p, 8, 3, "").Success);

        Assert.Equal("\"Parties \"", Shape(session, p));
        Assert.Empty(Revisions(session));
    }

    [Fact]
    public void Replacement_AcrossOriginalTextAndOwnInsertion_DeletesOnlyTheOriginal()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Parties "), Ins("Pat", 1, Run("binding")));

        // "s bi": the original "s " is deleted under tracking, Pat's own "bi" is simply removed.
        var edit = session.ReplaceTextAtSpan(p, 6, 4, "X");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("\"Partie\" del:Pat[\"s \"] ins:Pat[\"X\"] ins:Pat[\"nding\"]", Shape(session, p));
        Assert.Equal("PartieXnding", Accepted(session));
        Assert.Equal("Parties ", Rejected(session));
    }

    [Fact]
    public void Control_BackspaceInAnotherAuthorsInsertion_IsADeletionInsideIt()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Parties "), Ins("Alex", 1, Run("bindingg")));

        Assert.True(session.ReplaceTextAtSpan(p, 15, 1, "").Success);

        Assert.Equal("\"Parties \" ins:Alex[\"binding\" del:Pat[\"g\"]]", Shape(session, p));
        Assert.Equal("Parties binding", Accepted(session));
    }

    [Fact]
    public void Control_BackspaceInOwnMoveDestination_IsADeletionInsideIt()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat",
            Run("Parties "), new XElement(W.moveTo, Stamp("Pat", 1), Run("bindingg")));

        Assert.True(session.ReplaceTextAtSpan(p, 15, 1, "").Success);

        Assert.Equal("\"Parties \" moveTo:Pat[\"binding\" del:Pat[\"g\"]]", Shape(session, p));
    }

    // ─── Helpers ─────────────────────────────────────────────────────────

    private static XElement Run(string text, bool bold = false) => new(W.r,
        bold ? new XElement(W.rPr, new XElement(W.b)) : null,
        new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));

    private static object[] Stamp(string author, int id) => new object[]
    {
        new XAttribute(W.id, id), new XAttribute(W.author, author), new XAttribute(W.date, Date),
    };

    private static XElement Ins(string author, int id, params object[] content) =>
        new(W.ins, Stamp(author, id), content);

    private static XElement Del(string author, int id, string text) => new(W.del, Stamp(author, id),
        new XElement(W.r, new XElement(W.delText, new XAttribute(XNamespace.Xml + "space", "preserve"), text)));

    private static (DocxSession Session, string Anchor) Open(
        TrackedChangeMode tracking, string author, params object[] paragraphContent)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.body, new XElement(W.p, paragraphContent)))));
        }
        var session = new DocxSession(stream.ToArray(), new DocxSessionSettings
        {
            TrackedChanges = tracking,
            RevisionAuthor = author,
        });
        return (session, session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id);
    }

    /// <summary>The paragraph's content as runs and revision envelopes, ignoring properties,
    /// ids and dates: <c>"Fee " ins:Alex["amo"]</c>.</summary>
    private static string Shape(DocxSession session, string anchor) =>
        Shape(XElement.Parse(session.Raw.GetXml(anchor)));

    private static string Shape(XElement container) => string.Join(" ", container.Elements()
        .Where(e => e.Name != W.pPr && e.Name != W.rPr)
        .Select(e => e.Name == W.r
            ? "\"" + string.Concat(e.Elements().Where(t => t.Name == W.t || t.Name == W.delText).Select(t => t.Value)) + "\""
            : $"{e.Name.LocalName}{(e.Attribute(W.author) is { } a ? ":" + a.Value : "")}[{Shape(e)}]"));

    private static string[] Revisions(DocxSession session) =>
        session.ListRevisions().Select(r => $"{r.Type}:{r.Author}:{r.Text}").ToArray();

    private static string Accepted(DocxSession session) => ParagraphText(
        RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", session.Save())).DocumentByteArray);

    private static string Rejected(DocxSession session) => ParagraphText(
        RevisionProcessor.RejectRevisions(new WmlDocument("r.docx", session.Save())).DocumentByteArray);

    private static string ParagraphText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return string.Concat(document.MainDocumentPart!.GetXDocument().Descendants(W.t).Select(t => t.Value));
    }
}
