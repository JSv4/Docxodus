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
    public void Tracked_ReplaceText_LeavesNoteReferencesWhereTheySit(string shape, string rejected, string accepted)
    {
        var paragraph = shape switch
        {
            "leading" => E("p", NoteRef("1"), Run("text")),
            "trailing" => E("p", Run("text"), NoteRef("1")),
            "interleaved" => E("p", Run("Main text"), NoteRef("1"), Run(" continued.")),
            "hyperlink" => E("p", E("hyperlink", A("anchor", "x"), Run("Link")), NoteRef("1"), Run(" text")),
            _ => E("p", Del("A", "Old "), NoteRef("1"), Run("text")),
        };
        var bytes = Build(new[] { paragraph }, WithFootnotes(Footnote("1", "A note.")));

        bytes = Replace(bytes, "B", "New");
        AssertValid(bytes);

        foreach (var (resolver, resolve) in Resolvers())
        {
            Assert.True(rejected == TextWithNotes(Body(resolve(bytes, false))),
                $"{resolver} reject: expected '{rejected}' but got '{TextWithNotes(Body(resolve(bytes, false)))}'");
            Assert.True(accepted == TextWithNotes(Body(resolve(bytes, true))),
                $"{resolver} accept: expected '{accepted}' but got '{TextWithNotes(Body(resolve(bytes, true)))}'");
        }
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
    private static XElement Del(string author, string text) => E("del", Stamp(author),
        E("r", E("delText", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));

    private static DocxSessionSettings InlineProjection() => new()
    {
        ProjectionSettings = new WmlToMarkdownConverterSettings { TrackedChanges = TrackedChangeMode.RenderInline },
    };

    private static DocxSession Open(byte[] bytes, string author) => new(bytes, new DocxSessionSettings
    {
        TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = author,
    });

    private static string Anchor(DocxSession session, string kind) => session.Project().AnchorIndex.Values
        .First(a => a.Anchor.Kind == kind && a.Anchor.Scope == "body").Anchor.Id;

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
