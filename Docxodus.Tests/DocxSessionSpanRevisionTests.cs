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

    [Theory]
    [InlineData(4, 1, "", "\"Pa \" ins:Pat[\"a⇥cd\"]")]
    [InlineData(5, 1, "X", "\"Pa \" ins:Pat[\"ab⇥Xd\"]")]
    public void EditInOwnPendingInsertion_KeepsATabBetweenTwoTextNodesInPlace(int start, int length, string replace, string shape)
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Pa "), Ins("Pat", 1,
            new XElement(W.r, new XElement(W.t, "ab"), new XElement(W.tab), new XElement(W.t, "cd"))));

        Assert.True(session.ReplaceTextAtSpan(p, start, length, replace).Success);

        Assert.Equal(shape, Shape(session, p));
    }

    [Fact]
    public void EditInOwnPendingInsertion_KeepsABreakBetweenTwoTextNodesInPlace()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Pa "), Ins("Pat", 1,
            new XElement(W.r, new XElement(W.t, "ab"), new XElement(W.br), new XElement(W.t, "cd"))));

        Assert.True(session.ReplaceTextAtSpan(p, 3, 2, "").Success);

        Assert.Equal("\"Pa \" ins:Pat[\"↵cd\"]", Shape(session, p));
    }

    [Fact]
    public void ReplaceTextRangeAndReplaceMatch_InOwnPendingInsertion_EditItInPlace()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Fee: "));
        Assert.True(session.ReplaceTextAtSpan(p, 5, 0, "teh clause").Success);

        Assert.True(Assert.Single(session.ReplaceTextRange(p, "teh", "the")).Success);
        Assert.True(session.ReplaceMatch(Assert.Single(session.Grep("clause")), "term").Success);

        Assert.Equal("\"Fee: \" ins:Pat[\"the term\"]", Shape(session, p));
        Assert.Equal(new[] { "insert:Pat:the term" }, Revisions(session));
    }

    [Fact]
    public void BackspaceInOwnPendingInsertion_UndoesAndRedoes()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Parties "));
        Assert.True(session.ReplaceTextAtSpan(p, 8, 0, "bindingg").Success);
        var typed = Shape(session, p);

        Assert.True(session.ReplaceTextAtSpan(p, 15, 1, "").Success);
        Assert.True(session.Undo());
        Assert.Equal(typed, Shape(session, p));
        Assert.True(session.Redo());

        Assert.Equal("\"Parties \" ins:Pat[\"binding\"]", Shape(session, p));
    }

    [Fact]
    public void BackspaceInOwnPendingInsertionInsideAHyperlink_RemovesTheCharacter()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("See "),
            new XElement(W.hyperlink, new XAttribute(W.anchor, "terms"), Ins("Pat", 1, Run("termss"))));

        Assert.True(session.ReplaceTextAtSpan(p, 9, 1, "").Success);

        Assert.Equal("\"See \" hyperlink[ins:Pat[\"terms\"]]", Shape(session, p));
    }

    [Fact]
    public void EditInOwnPendingInsertion_KeepsABookmarkInside()
    {
        var (session, p) = Open(TrackedChangeMode.RenderInline, "Pat", Run("Pa "),
            Ins("Pat", 1, Bookmark(W.bookmarkStart), Run("bindingg"), Bookmark(W.bookmarkEnd)));

        Assert.True(session.ReplaceTextAtSpan(p, 10, 1, "").Success);

        Assert.Equal("\"Pa \" ins:Pat[bookmarkStart[] \"binding\" bookmarkEnd[]]", Shape(session, p));
    }

    [Fact]
    public void Control_RemovingAPictureFromAnotherAuthorsInsertion_IsRefused()
    {
        // A control for the image ops' own-insertion check, now OwnPendingInsertion: another
        // author's insertion is refused before that check runs, so this passed before too.
        // Alex inserts a picture under tracking; Pat then tries to remove it with tracking on.
        byte[] bytes;
        using (var alex = new DocxSession(Build(Run("Logo: ")), new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Alex",
        }))
        {
            var anchor = alex.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
            Assert.True(alex.InsertImage(anchor, 6, DocxSessionImageTests.Png(2, 3)).Success);
            bytes = alex.Save();
        }
        using var pat = new DocxSession(bytes, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Pat",
        });
        var before = pat.GetPackageContentHash();

        var removed = pat.RemoveImage(Assert.Single(pat.ListImages()).Id);

        Assert.False(removed.Success);
        Assert.Equal(before, pat.GetPackageContentHash());
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

    private static XElement Bookmark(XName marker) => marker == W.bookmarkStart
        ? new XElement(marker, new XAttribute(W.id, 10), new XAttribute(W.name, "_Toc1"))
        : new XElement(marker, new XAttribute(W.id, 10));

    private static (DocxSession Session, string Anchor) Open(
        TrackedChangeMode tracking, string author, params object[] paragraphContent)
    {
        var session = new DocxSession(Build(paragraphContent), new DocxSessionSettings
        {
            TrackedChanges = tracking,
            RevisionAuthor = author,
        });
        return (session, session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id);
    }

    /// <summary>A document whose body is one paragraph holding <paramref name="paragraphContent"/>.</summary>
    internal static byte[] Build(params object[] paragraphContent)
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
        return stream.ToArray();
    }

    /// <summary>The paragraph's content as runs and revision envelopes, ignoring properties,
    /// ids and dates: <c>"Fee " ins:Alex["amo"]</c>.</summary>
    private static string Shape(DocxSession session, string anchor) =>
        Shape(XElement.Parse(session.Raw.GetXml(anchor)));

    private static string Shape(XElement container) => string.Join(" ", container.Elements()
        .Where(e => e.Name != W.pPr && e.Name != W.rPr)
        .Select(e => e.Name == W.r
            ? "\"" + string.Concat(e.Elements().Select(t => t.Name == W.t || t.Name == W.delText ? t.Value
                : t.Name == W.tab ? "⇥" : t.Name == W.br ? "↵" : "")) + "\""
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
