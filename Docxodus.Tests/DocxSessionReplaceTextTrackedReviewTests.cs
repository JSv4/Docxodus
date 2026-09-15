// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tracked whole-paragraph replacement (<see cref="DocxSession.ReplaceText"/> under
/// <see cref="TrackedChangeMode.RenderInline"/>) after issue #786: the markup it writes beside
/// earlier revisions, hyperlinks, inline controls and note references, and what every reader of
/// that markup — the projection, the revision list and both resolvers — makes of it.
/// </summary>
public class DocxSessionReplaceTextTrackedReviewTests
{
    private static int _nextRevisionId = 100;

    [Fact]
    public void Projection_TreatsContentDeletedInsideAnInsertionAsDeleted()
    {
        var bytes = Build(P("ZZMARK München."));
        bytes = Replace(bytes, "A", "ZZMARK Hamburg.");
        bytes = Replace(bytes, "B", "ZZMARK Köln.");

        using var accepted = new DocxSession(bytes);
        var acceptedMarkdown = accepted.Project().Markdown;
        Assert.Contains("ZZMARK Köln.", acceptedMarkdown);
        Assert.DoesNotContain("Hamburg", acceptedMarkdown);

        using var inline = new DocxSession(bytes, InlineProjection());
        var inlineMarkdown = inline.Project().Markdown;
        Assert.Contains("{-ZZMARK Hamburg.-}", inlineMarkdown);
        Assert.DoesNotContain("{+ZZMARK Hamburg.+}", inlineMarkdown);
        Assert.Contains("{+ZZMARK Köln.+}", inlineMarkdown);
    }

    [Theory]
    [InlineData("insertion")]
    [InlineData("hyperlink")]
    public void Tracked_ReplaceText_KeepsANoteReferenceNestedInEarlierMarkup(string shape)
    {
        byte[] bytes;
        if (shape == "insertion")
        {
            bytes = Build(P("Main text"));
            using var author = Open(bytes, "A");
            var inserted = author.InsertFootnote(Anchor(author, "p"), 9, "A note.");
            Assert.True(inserted.Success, inserted.Error?.Message);
            bytes = author.Save();
            Assert.Single(Body(bytes).Elements(W.p).Single().Elements(W.ins));
        }
        else
        {
            bytes = Build(
                new[] { E("p", Run("See "), E("hyperlink", A("anchor", "x"), Run("link"), NoteRef("1")), Run(" tail")) },
                WithFootnotes(Footnote("1", "A note.")));
        }
        Assert.Single(Body(bytes).Descendants(W.footnoteReference));

        bytes = Replace(bytes, "B", "Replaced.");
        AssertValid(bytes);

        foreach (var (resolver, resolve) in Resolvers())
        {
            var resolved = resolve(bytes, true);
            Assert.True(Body(resolved).Descendants(W.footnoteReference).Count() == 1,
                $"{resolver}: the note reference must survive acceptance");
            Assert.Contains("A note.", Text(Root(resolved, "word/footnotes.xml")));
        }
    }

    [Theory]
    [InlineData("customXml")]
    [InlineData("emptyField")]
    public void Tracked_ReplaceText_RefusesUnrecordableInlineContentBeforeMutation(string shape)
    {
        var unrecordable = shape == "customXml"
            ? E("customXml", A("element", "foo"), Run("inside"))
            : E("fldSimple", A("instr", " PAGE "));
        var bytes = Build(E("p", Run("Before "), unrecordable, Run(" after")));
        using var session = Open(bytes, "B");

        var result = session.ReplaceText(Anchor(session, "p"), "Replacement.");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.IncompatibleElementType, result.Error?.Code);
        AssertXmlEqual(Body(bytes), Body(session.Save()));
    }

    [Theory]
    [InlineData("leading", "[note]text", "[note]New")]
    [InlineData("trailing", "text[note]", "New[note]")]
    [InlineData("interleaved", "Main text[note] continued.", "[note]New")]
    [InlineData("hyperlink", "Link[note] text", "[note]New")]
    [InlineData("prior-deletion", "Old [note]text", "[note]New")]
    [InlineData("inserted-trailing", "Main text", "New[note]")]
    [InlineData("hyperlink-trailing", "See link[note]", "New[note]")]
    public void Tracked_ReplaceText_LeavesNoteReferencesWhereTheySit(string shape, string rejected, string accepted)
    {
        var paragraph = shape switch
        {
            "leading" => E("p", NoteRef("1"), Run("text")),
            "trailing" => E("p", Run("text"), NoteRef("1")),
            "interleaved" => E("p", Run("Main text"), NoteRef("1"), Run(" continued.")),
            "hyperlink" => E("p", E("hyperlink", A("anchor", "x"), Run("Link")), NoteRef("1"), Run(" text")),
            "inserted-trailing" => E("p", Run("Main text"), Ins("A", NoteRef("1"))),
            "hyperlink-trailing" => E("p", Run("See "), E("hyperlink", A("anchor", "x"), Run("link"), NoteRef("1"))),
            _ => E("p", Del("A", "Old "), NoteRef("1"), Run("text")),
        };
        var bytes = Build(new[] { paragraph }, WithFootnotes(Footnote("1", "A note.")));

        bytes = Replace(bytes, "B", "New");
        AssertValid(bytes);

        foreach (var (resolver, resolve) in Resolvers())
        {
            var afterReject = TextWithNotes(Body(resolve(bytes, false)));
            Assert.True(rejected == afterReject, $"{resolver} reject: expected '{rejected}' but got '{afterReject}'");
            var afterAccept = TextWithNotes(Body(resolve(bytes, true)));
            Assert.True(accepted == afterAccept, $"{resolver} accept: expected '{accepted}' but got '{afterAccept}'");
        }
    }

    [Fact]
    public void Tracked_ReplaceText_LeavesTextBoxRevisionsUntouched()
    {
        var bytes = Build(E("p", Run("Intro "), TextBox(Ins("A", Run("Boxed"))), Run(" tail")));

        bytes = Replace(bytes, "B", "Replacement.");
        AssertValid(bytes);

        // A's insertion inside the box is A's revision, not text this deletion owns.
        var boxed = Body(bytes).Descendants(W.txbxContent).Single();
        Assert.Empty(boxed.Descendants(W.delText));
        Assert.Equal("Boxed", Text(boxed));

        using var review = new DocxSession(bytes);
        foreach (var deletion in review.ListRevisions().Where(r => r.Author == "B" && r.Type == "delete").ToList())
            Assert.True(review.RejectRevision(deletion.Id).Success);
        foreach (var insertion in review.ListRevisions().Where(r => r.Author == "A").ToList())
            Assert.True(review.AcceptRevision(insertion.Id).Success);
        var resolved = Body(review.Save());
        Assert.Empty(resolved.Descendants(W.delText));
        Assert.Equal("Boxed", Text(resolved.Descendants(W.txbxContent).Single()));
    }

    [Fact]
    public void Tracked_ReplaceText_NestsInsertedLinkRunsInsideTheHyperlink()
    {
        var bytes = Build(P("Old."));

        bytes = Replace(bytes, "B", "See [the site](https://example.com/x) now.");
        AssertValid(bytes);

        var paragraph = Body(bytes).Element(W.p)!;
        Assert.DoesNotContain(paragraph.Elements(W.ins), i => i.Element(W.hyperlink) is not null);
        var hyperlink = Assert.Single(paragraph.Elements(W.hyperlink));
        Assert.Equal("the site", Text(Assert.Single(hyperlink.Elements(W.ins))));
        Assert.Equal("See the site now.", Text(Body(Accepted(bytes))));
        Assert.Equal("Old.", Text(Body(Rejected(bytes))));
    }

    [Fact]
    public void Tracked_ReplaceText_UnInsertsTheAuthorsOwnFieldAndNestedEnvelopes()
    {
        var bytes = Build(E("p", Run("Ref: "),
            E("fldSimple", A("instr", " REF x "), Ins("B", Run("cached"))),
            Ins("B", Ins("B", Run("twice")))));

        bytes = Replace(bytes, "B", "New");
        AssertValid(bytes);

        var paragraph = Body(bytes).Element(W.p)!;
        Assert.Empty(paragraph.Descendants(W.fldSimple));
        Assert.Equal("New", Text(Assert.Single(paragraph.Descendants(W.ins))));
        using var review = new DocxSession(bytes);
        Assert.All(review.ListRevisions(), r => Assert.NotEqual(string.Empty, r.Text));

        // The previous replacement must not have manufactured a shape the next one refuses.
        bytes = Replace(bytes, "B", "Again");
        Assert.Equal("Again", Text(Body(Accepted(bytes))));
    }

    [Theory]
    [InlineData("own-control")]
    [InlineData("existing-control")]
    public void Tracked_ReplaceText_UnInsertsInsideAControlWithoutBreakingIt(string shape)
    {
        var control = shape == "own-control"
            ? Ins("B", E("sdt", E("sdtPr", E("id", A("val", "7"))), E("sdtContent", Run("draft"))))
            : E("sdt", E("sdtPr", E("id", A("val", "7"))), E("sdtContent", Ins("B", Run("draft"))));
        var bytes = Build(E("p", Run("Amount: "), control));

        bytes = Replace(bytes, "B", "Final.");
        AssertValid(bytes);

        var paragraph = Body(bytes).Element(W.p)!;
        if (shape == "own-control")
            Assert.Empty(paragraph.Descendants(W.sdt));
        else
            Assert.NotNull(Assert.Single(paragraph.Descendants(W.sdt)).Element(W.sdtContent));
        Assert.Equal("Final.", Text(Body(Accepted(bytes))));
    }

    [Fact]
    public void Projection_TreatsMovedRunsAsRevisions()
    {
        var bytes = Build(
            E("p", Run("Visit "),
                E("moveFromRangeStart", A("id", "90"), A("name", "move1"), A("author", "A"), A("date", "2026-01-01T00:00:00Z")),
                E("hyperlink", A("anchor", "site"), E("moveFrom", Stamp("A"), DeletedRun("the site"))),
                E("moveFromRangeEnd", A("id", "90")),
                Run(" today")),
            E("p",
                E("moveToRangeStart", A("id", "91"), A("name", "move1"), A("author", "A"), A("date", "2026-01-01T00:00:00Z")),
                E("moveTo", Stamp("A"), Run("the site")),
                E("moveToRangeEnd", A("id", "91"))));

        using var accepted = new DocxSession(bytes);
        var acceptedMarkdown = accepted.Project().Markdown;
        Assert.DoesNotContain("[the site](#site)", acceptedMarkdown);
        Assert.Contains("the site", acceptedMarkdown);

        using var inline = new DocxSession(bytes, InlineProjection());
        var inlineMarkdown = inline.Project().Markdown;
        Assert.Contains("{-[the site](#site)-}", inlineMarkdown);
        Assert.Contains("{+the site+}", inlineMarkdown);
    }

    [Fact]
    public void Projection_KeepsTextBoxRunsInsideARevisionEnvelope()
    {
        // The flat text span operations address counts a text box's runs when the box rides
        // inside an envelope; the projection must show the same characters.
        var bytes = Build(E("p", Run("Intro"), Ins("A", TextBox(Run("Box text")))));

        using var accepted = new DocxSession(bytes);
        Assert.Contains("IntroBox text", accepted.Project().Markdown);
        using var inline = new DocxSession(bytes, InlineProjection());
        Assert.Contains("Intro{+Box text+}", inline.Project().Markdown);
    }

    [Fact]
    public void Tracked_ReplaceText_PrunesANoteWhoseOnlyReferenceItUnInserted()
    {
        var mixed = E("r", E("t", new XAttribute(XNamespace.Xml + "space", "preserve"), "mixed"),
            E("footnoteReference", A("id", "1")));
        var bytes = Build(new[] { E("p", Run("Lead "), Ins("B", mixed)) }, WithFootnotes(Footnote("1", "A note.")));
        using var session = Open(bytes, "B");

        var result = session.ReplaceText(Anchor(session, "p"), "New");

        Assert.True(result.Success, result.Error?.Message);
        Assert.NotEmpty(result.Removed);
        var edited = session.Save();
        Assert.Empty(Body(edited).Descendants(W.footnoteReference));
        Assert.DoesNotContain(Root(edited, "word/footnotes.xml").Descendants(W.footnote), f => f.Attribute(W.w + "type") is null);
        Assert.DoesNotContain(session.ListRevisions(), r => r.Text.Contains("A note", StringComparison.Ordinal));
    }

    [Fact]
    public void Tracked_ReplaceText_AllowsUnrecordableContentInsideATextBox()
    {
        var bytes = Build(E("p", Run("Body "), TextBox(Run("Page "), E("fldSimple", A("instr", " PAGE "))), Run(" tail")));
        AssertValid(bytes);

        bytes = Replace(bytes, "B", "Replaced.");

        AssertValid(bytes);
        Assert.Equal("Replaced.", Text(Body(Accepted(bytes))));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReplaceText_WithAnEmptyPayload_EmptiesTheParagraph(bool tracked)
    {
        var bytes = Build(P("Original."));
        using var session = new DocxSession(bytes, new DocxSessionSettings
        {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "B",
        });

        var result = session.ReplaceText(Anchor(session, "p"), string.Empty);

        Assert.True(result.Success, result.Error?.Message);
        var edited = session.Save();
        Assert.Equal(string.Empty, Text(Body(edited)));
        if (tracked) Assert.Equal("Original.", Text(Body(Rejected(edited))));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReplaceText_InACommentBody_KeepsTheAnnotationMark(bool tracked)
    {
        var bytes = Build(P("Body text."));
        using (var author = new DocxSession(bytes))
        {
            var added = author.AddComment(Anchor(author, "p"), new CharSpan(0, 4), "A", "Why?");
            Assert.True(added.Success, added.Error?.Message);
            bytes = author.Save();
        }
        Assert.Single(Root(bytes, "word/comments.xml").Descendants(W.annotationRef));
        using var session = new DocxSession(bytes, new DocxSessionSettings
        {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "B",
        });
        var anchor = session.Project().AnchorIndex.Values
            .Single(a => a.Anchor.Kind == "p" && a.Anchor.Scope == "cmt").Anchor.Id;

        var result = session.ReplaceText(anchor, "Edited comment.");

        Assert.True(result.Success, result.Error?.Message);
        var edited = session.Save();
        var resolutions = tracked
            ? new List<(string name, byte[] bytes)> { ("session", ResolveSession(edited, true)) }
            : new List<(string name, byte[] bytes)> { ("untracked", edited) };
        foreach (var (name, resolved) in resolutions)
        {
            var paragraph = Root(resolved, "word/comments.xml").Descendants(W.comment).Single().Element(W.p)!;
            Assert.True(paragraph.Descendants(W.annotationRef).Count() == 1, $"{name}: annotation mark lost");
            Assert.Equal("Edited comment.", Text(paragraph));
        }
    }

    [Fact]
    public void ListRevisions_ReportsTheTextOfAMoveSource()
    {
        var bytes = Build(P("Moved paragraph."), P("Anchor."));
        using var session = Open(bytes, "B");

        var moved = session.MoveBlock(Anchor(session, "p"), Anchor(session, "p", 1), Position.After);

        Assert.True(moved.Success, moved.Error?.Message);
        var entry = Assert.Single(session.ListRevisions(), r => r.Type == "move");
        Assert.Contains("Moved paragraph.", entry.Text);
    }

    [Fact]
    public void Tracked_ReplaceText_DeletesInsideAnInlineControlBelowTheEarlierInsertion()
    {
        var control = E("sdt", E("sdtPr", E("id", A("val", "7"))), E("sdtContent", Ins("A", Run("draft"))));
        var bytes = Build(E("p", Run("Amount: "), control));

        bytes = Replace(bytes, "B", "Final.");
        AssertValid(bytes);

        var paragraph = Body(bytes).Element(W.p)!;
        Assert.DoesNotContain(paragraph.Elements(W.del), d => d.Element(W.sdt) is not null);
        var deletion = Assert.Single(paragraph.Descendants(W.sdtContent).Single().Descendants(W.del));
        Assert.Equal(W.ins, deletion.Parent!.Name);
        Assert.Equal("A", (string?)deletion.Parent.Attribute(W.author));

        using var review = new DocxSession(bytes);
        foreach (var entry in review.ListRevisions().Where(r => r.Author == "B" && r.Type == "delete").ToList())
            Assert.True(review.RejectRevision(entry.Id).Success);
        var partial = Body(review.Save());
        Assert.Empty(partial.Descendants(W.delText));
        Assert.Equal("draft", Text(partial.Descendants(W.sdt).Single()));
        Assert.Contains(review.ListRevisions(), r => r.Author == "A" && r.Type == "insert");

        Assert.Equal("Final.", Text(Body(Accepted(bytes))));
        Assert.Equal("Amount: ", Text(Body(Rejected(bytes))));
    }

    [Fact]
    public void Projection_RendersRevisedHyperlinkRuns()
    {
        var bytes = Build(E("p", Run("Venue: "), E("hyperlink", A("anchor", "venue"), Run("München."))));
        bytes = Replace(bytes, "B", "Venue: Bonn.");

        using var accepted = new DocxSession(bytes);
        var acceptedMarkdown = accepted.Project().Markdown;
        Assert.DoesNotContain("München", acceptedMarkdown);
        Assert.Contains("Venue: Bonn.", acceptedMarkdown);

        using var inline = new DocxSession(bytes, InlineProjection());
        var inlineMarkdown = inline.Project().Markdown;
        Assert.Contains("{-[München.](#venue)-}", inlineMarkdown);
        Assert.Contains("{+Venue: Bonn.+}", inlineMarkdown);

        var inserted = Build(E("p", Run("Venue: "), E("hyperlink", A("anchor", "venue"), Ins("A", Run("Hamburg.")))));
        using var live = new DocxSession(inserted);
        Assert.Contains("[Hamburg.](#venue)", live.Project().Markdown);
    }

    [Fact]
    public void Tracked_ReplaceText_UnInsertsTheAuthorsOwnPendingText()
    {
        var bytes = Build(P("Original."));
        bytes = Replace(bytes, "B", "First try.");
        bytes = Replace(bytes, "B", "Second try.");
        AssertValid(bytes);

        var paragraph = Body(bytes).Element(W.p)!;
        var insertion = Assert.Single(paragraph.Descendants(W.ins));
        Assert.Equal("Second try.", Text(insertion));
        Assert.DoesNotContain("First try.", paragraph.ToString());
        Assert.Equal("Original.", Text(Body(Rejected(bytes))));
        Assert.Equal("Second try.", Text(Body(Accepted(bytes))));
        using var review = new DocxSession(bytes);
        Assert.Equal(new[] { "delete", "insert" }, review.ListRevisions().Select(r => r.Type));
    }

    [Fact]
    public void Tracked_ReplaceText_RecordsOneDeletionPerContiguousLiveSpan()
    {
        // Characterisation of the id granularity: a foreign revision between live runs splits the
        // replacement's deletion into independently addressable groups, as Word's own markup does.
        var bytes = Build(E("p", Run("a "), Del("A", "b "), Run("c")));
        bytes = Replace(bytes, "B", "z");

        using var review = new DocxSession(bytes);
        var revisions = review.ListRevisions();
        Assert.Equal(new[] { "delete/B/a ", "delete/A/b ", "delete/B/c", "insert/B/z" },
            revisions.Select(r => $"{r.Type}/{r.Author}/{r.Text}"));
        Assert.True(review.RejectRevision(revisions[0].Id).Success);
        Assert.Equal("a z", Text(Body(review.Save())));
    }

    [Fact]
    public void ListRevisions_ReportsTheTextOfASupersededInsertion()
    {
        var bytes = Build(P("ZZMARK München."));
        bytes = Replace(bytes, "A", "ZZMARK Hamburg.");
        bytes = Replace(bytes, "B", "ZZMARK Köln.");

        using var review = new DocxSession(bytes);
        var superseded = Assert.Single(review.ListRevisions(), r => r.Type == "insert" && r.Author == "A");
        Assert.Equal("ZZMARK Hamburg.", superseded.Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReplaceText_InANoteBody_KeepsTheNoteMark(bool tracked)
    {
        var note = E("footnote", A("id", "1"), E("p", E("r", E("footnoteRef")), Run(" Original note.")));
        var bytes = Build(new[] { E("p", Run("Body"), NoteRef("1")) }, WithFootnotes(note));
        using var session = new DocxSession(bytes, new DocxSessionSettings
        {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "B",
        });
        var anchor = session.Project().AnchorIndex.Values
            .Single(a => a.Anchor.Kind == "p" && a.Anchor.Scope == "fn").Anchor.Id;

        var result = session.ReplaceText(anchor, "Corrected note.");

        Assert.True(result.Success, result.Error?.Message);
        var edited = session.Save();
        var resolutions = tracked
            ? Resolvers().Select(r => (r.name, bytes: r.resolve(edited, true))).ToList()
            : new List<(string name, byte[] bytes)> { ("untracked", edited) };
        foreach (var (name, resolved) in resolutions)
        {
            var paragraph = Root(resolved, "word/footnotes.xml").Descendants(W.footnote).Single().Element(W.p)!;
            var withMark = string.Concat(paragraph.Descendants()
                .Where(e => e.Name == W.t || e.Name == W.footnoteRef)
                .Select(e => e.Name == W.footnoteRef ? "[mark]" : e.Value));
            Assert.True(withMark == "[mark]Corrected note.", $"{name}: got '{withMark}'");
        }
    }

    [Theory]
    [InlineData("fi-FI")]
    [InlineData("th-TH")]
    public void Tracked_ReplaceText_StampsInvariantRevisionDates(string culture)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            var bytes = Build(P("Original."));
            bytes = Replace(bytes, "B", "Replaced.");

            var dates = Body(bytes).Descendants().Attributes(W.date).Select(a => a.Value).ToList();
            Assert.NotEmpty(dates);
            Assert.All(dates, date =>
            {
                Assert.Matches(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$", date);
                Assert.StartsWith(DateTime.UtcNow.Year.ToString(CultureInfo.InvariantCulture), date);
            });
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("inline-control")]
    public void Tracked_ReplaceText_KeepsCommentMarkers(string shape)
    {
        object[] commented =
        {
            E("commentRangeStart", A("id", "0")), Run("$5"), E("commentRangeEnd", A("id", "0")),
            E("r", E("commentReference", A("id", "0"))),
        };
        var paragraph = shape == "paragraph"
            ? E("p", Run("Intro "), commented)
            : E("p", Run("Intro "), E("sdt", E("sdtPr", E("id", A("val", "8"))), E("sdtContent", commented)));
        var bytes = Build(new[] { paragraph }, WithComments("0"));

        bytes = Replace(bytes, "B", "Intro $6");
        AssertValid(bytes);

        foreach (var (resolver, resolve) in Resolvers())
        {
            var body = Body(resolve(bytes, true));
            Assert.True(body.Descendants(W.commentRangeStart).Count() == 1, $"{resolver}: range start lost");
            Assert.True(body.Descendants(W.commentRangeEnd).Count() == 1, $"{resolver}: range end lost");
            Assert.True(body.Descendants(W.commentReference).Count() == 1, $"{resolver}: reference lost");
            Assert.Equal("Intro $6", Text(body));
        }
    }

    // ─── Helpers ─────────────────────────────────────────────────────────

    private static XElement E(string name, params object?[] content) => new(W.w + name, content);
    private static XAttribute A(string name, string value) => new(W.w + name, value);
    private static XElement Run(string text) => E("r", E("t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    private static XElement P(string text) => E("p", Run(text));
    private static XElement NoteRef(string id) => E("r", E("footnoteReference", A("id", id)));
    private static XElement Footnote(string id, string text) => E("footnote", A("id", id), P(text));
    private static object[] Stamp(string author) => new object[]
    {
        A("id", (_nextRevisionId++).ToString(CultureInfo.InvariantCulture)),
        A("author", author), A("date", "2026-01-01T00:00:00Z"),
    };
    private static XElement Ins(string author, params object[] content) => E("ins", Stamp(author), content);
    private static XElement DeletedRun(string text) =>
        E("r", E("delText", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    private static XElement Del(string author, string text) => E("del", Stamp(author), DeletedRun(text));
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static XElement TextBox(params object[] paragraphContent) => E("r", E("pict",
        new XElement(Vml + "shape", new XAttribute("id", "TextBox1"), new XAttribute("style", "width:100pt;height:30pt"),
            new XElement(Vml + "textbox", E("txbxContent", E("p", paragraphContent))))));

    private static DocxSessionSettings InlineProjection() => new()
    {
        ProjectionSettings = new WmlToMarkdownConverterSettings { TrackedChanges = TrackedChangeMode.RenderInline },
    };

    private static DocxSession Open(byte[] bytes, string author) => new(bytes, new DocxSessionSettings
    {
        TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = author,
    });

    private static string Anchor(DocxSession session, string kind, int skip = 0) => session.Project().AnchorIndex.Values
        .Where(a => a.Anchor.Kind == kind && a.Anchor.Scope == "body").Skip(skip).First().Anchor.Id;

    private static byte[] Replace(byte[] bytes, string author, string text)
    {
        using var session = Open(bytes, author);
        var result = session.ReplaceText(Anchor(session, "p"), text);
        Assert.True(result.Success, result.Error?.Message);
        return session.Save();
    }

    private static Action<MainDocumentPart> WithFootnotes(params XElement[] footnotes) => main =>
        main.AddNewPart<FootnotesPart>().PutXDocument(new XDocument(E("footnotes", footnotes)));

    private static Action<MainDocumentPart> WithComments(string id) => main =>
        main.AddNewPart<WordprocessingCommentsPart>().PutXDocument(new XDocument(E("comments",
            E("comment", A("id", id), A("author", "A"), P("Why?")))));

    private static byte[] Build(params XElement[] blocks) => Build(blocks, null);

    private static byte[] Build(XElement[] blocks, Action<MainDocumentPart>? setup)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(E("styles")));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(E("settings")));
            setup?.Invoke(main);
            main.PutXDocument(new XDocument(E("document", new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "r", R.r), E("body", blocks))));
            document.Save();
        }
        return stream.ToArray();
    }

    private static XElement Root(byte[] bytes, string part = "word/document.xml")
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(part)!.Open();
        return XElement.Load(stream);
    }

    private static XElement Body(byte[] bytes) => Root(bytes).Element(W.body)!;

    private static string Text(XElement root) => string.Concat(root.Descendants(W.t).Select(t => t.Value));

    private static string TextWithNotes(XElement root) => string.Concat(root.Descendants()
        .Where(e => e.Name == W.t || e.Name == W.footnoteReference)
        .Select(e => e.Name == W.footnoteReference ? "[note]" : e.Value));

    private static byte[] Accepted(byte[] bytes) =>
        RevisionProcessor.AcceptRevisions(new WmlDocument("accepted.docx", bytes)).DocumentByteArray;

    private static byte[] Rejected(byte[] bytes) =>
        RevisionProcessor.RejectRevisions(new WmlDocument("rejected.docx", bytes)).DocumentByteArray;

    private static byte[] ResolveSession(byte[] bytes, bool accept)
    {
        using var session = new DocxSession(bytes);
        var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(session.ListRevisions());
        return session.Save();
    }

    private static IEnumerable<(string name, Func<byte[], bool, byte[]> resolve)> Resolvers() => new[]
    {
        ("processor", new Func<byte[], bool, byte[]>((bytes, accept) => accept ? Accepted(bytes) : Rejected(bytes))),
        ("session", ResolveSession),
    };

    private static void AssertXmlEqual(XElement expected, XElement actual)
    {
        static string Normalize(XElement value)
        {
            value = new XElement(value);
            value.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration || a.Name == PtOpenXml.Unid).Remove();
            return value.ToString(SaveOptions.DisableFormatting);
        }
        Assert.Equal(Normalize(expected), Normalize(actual));
    }

    private static void AssertValid(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }
}
