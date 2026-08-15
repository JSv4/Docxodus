#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Surgical text replacements in <see cref="TrackedChangeMode.RenderInline"/>
/// (issue #330), and the visible-text contract those offsets are computed over
/// (DS409-DS410, found by the issue #435 acceptance smoke). Test IDs DS400-DS410.
/// </summary>
public class DocxSessionSurgicalTrackedChangesTests
{
    private static readonly XNamespace Xml = XNamespace.Xml;

    private static XElement Run(string text, params XElement[] properties)
    {
        var run = new XElement(W.r);
        if (properties.Length > 0) run.Add(new XElement(W.rPr, properties));
        run.Add(new XElement(W.t, new XAttribute(Xml + "space", "preserve"), text));
        return run;
    }

    private static byte[] BuildDocument(params object[] bodyChildren)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();

            var xDocument = main.GetXDocument();
            xDocument.Root!.Element(W.body)!.ReplaceNodes(bodyChildren);
            main.PutXDocument();
        }

        return stream.ToArray();
    }

    private static XElement MainRoot(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!);
    }

    private static XElement? SettingsRoot(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.DocumentSettingsPart?.GetXDocument().Root is { } root
            ? new XElement(root)
            : null;
    }

    private static string AcceptedText(byte[] bytes)
    {
        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("accepted.docx", bytes));
        return Text(MainRoot(accepted.DocumentByteArray));
    }

    private static string RejectedText(byte[] bytes)
    {
        var rejected = RevisionProcessor.RejectRevisions(new WmlDocument("rejected.docx", bytes));
        return Text(MainRoot(rejected.DocumentByteArray));
    }

    private static string Text(XElement root) =>
        string.Concat(root.Descendants(W.t).Select(t => t.Value));

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(document)
            .Select(e => $"{e.Part?.Uri}: {e.Description} ({e.Path?.XPath})")
            .ToArray();
        Assert.Empty(errors);
    }

    [Fact]
    public void DS400_ReplaceTextAtSpan_TrackedSingleRunUsesWordRevisionEnvelopes()
    {
        var source = BuildDocument(
            new XElement(W.p,
                Run("Prefix target suffix",
                    new XElement(W.i),
                    new XElement(W.color, new XAttribute(W.val, "336699")))));
        using var session = new DocxSession(source,
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Range Reviewer",
            });
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = session.ReplaceTextAtSpan(anchor, 7, 6, "replacement");

        Assert.True(result.Success, result.Error?.Message);
        var tracked = session.Save();
        var paragraph = MainRoot(tracked).Descendants(W.p).Single();
        var deletion = Assert.Single(paragraph.Elements(W.del));
        var insertion = Assert.Single(paragraph.Elements(W.ins));
        Assert.Equal("target", Assert.Single(deletion.Descendants(W.delText)).Value);
        Assert.Equal("replacement", Assert.Single(insertion.Descendants(W.t)).Value);
        Assert.Equal("Range Reviewer", (string?)deletion.Attribute(W.author));
        Assert.Equal("Range Reviewer", (string?)insertion.Attribute(W.author));
        Assert.Equal((string?)deletion.Attribute(W.date), (string?)insertion.Attribute(W.date));
        Assert.NotEqual((string?)deletion.Attribute(W.id), (string?)insertion.Attribute(W.id));
        Assert.Matches(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$",
            (string)deletion.Attribute(W.date)!);

        var insertedProperties = insertion.Descendants(W.r).Single().Element(W.rPr);
        Assert.NotNull(insertedProperties?.Element(W.i));
        Assert.Equal("336699", (string?)insertedProperties?.Element(W.color)?.Attribute(W.val));
        Assert.Equal("Prefix ", paragraph.Elements(W.r).First().Element(W.t)?.Value);
        Assert.Equal(" suffix", paragraph.Elements(W.r).Last().Element(W.t)?.Value);

        Assert.Equal("Prefix replacement suffix", AcceptedText(tracked));
        Assert.Equal("Prefix target suffix", RejectedText(tracked));
        Assert.NotNull(SettingsRoot(tracked)?.Element(W.trackRevisions));
        AssertSchemaValid(tracked);

        using (var selectiveAccept = new DocxSession(tracked))
        {
            while (selectiveAccept.ListRevisions().FirstOrDefault() is { } revision)
                Assert.True(selectiveAccept.AcceptRevision(revision.Id).Success);
            Assert.Equal("Prefix replacement suffix", Text(MainRoot(selectiveAccept.Save())));
        }

        using (var selectiveReject = new DocxSession(tracked))
        {
            while (selectiveReject.ListRevisions().FirstOrDefault() is { } revision)
                Assert.True(selectiveReject.RejectRevision(revision.Id).Success);
            Assert.Equal("Prefix target suffix", Text(MainRoot(selectiveReject.Save())));
        }
    }

    [Fact]
    public void DS408_TrackedEdit_CreatesMissingSettingsPartWithTrackRevisions()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("before")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.Document.Save();
        }

        using var session = new DocxSession(stream.ToArray(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = session.ReplaceText(anchor, "after");

        Assert.True(result.Success, result.Error?.Message);
        var tracked = session.Save();
        var settings = Assert.IsType<XElement>(SettingsRoot(tracked));
        Assert.Single(settings.Elements(W.trackRevisions));
        AssertSchemaValid(tracked);
    }

    [Fact]
    public void DS401_ReplaceTextRange_TrackedMultiRunPreservesPerRunFormatting()
    {
        var source = BuildDocument(
            new XElement(W.p,
                Run("Plain ", new XElement(W.u, new XAttribute(W.val, "single"))),
                Run("BOLD", new XElement(W.b)),
                Run(" tail", new XElement(W.i))));
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = Assert.Single(session.ReplaceTextRange(anchor, "ain BOLD ta", "REPL"));

        Assert.True(result.Success, result.Error?.Message);
        var tracked = session.Save();
        var paragraph = MainRoot(tracked).Descendants(W.p).Single();
        var deletion = Assert.Single(paragraph.Elements(W.del));
        var deletedRuns = deletion.Elements(W.r).ToArray();
        Assert.Equal(3, deletedRuns.Length);
        Assert.Equal("ain BOLD ta", string.Concat(deletedRuns
            .SelectMany(r => r.Elements(W.delText)).Select(t => t.Value)));
        Assert.NotNull(deletedRuns[0].Element(W.rPr)?.Element(W.u));
        Assert.NotNull(deletedRuns[1].Element(W.rPr)?.Element(W.b));
        Assert.NotNull(deletedRuns[2].Element(W.rPr)?.Element(W.i));

        var insertedRun = Assert.Single(Assert.Single(paragraph.Elements(W.ins)).Elements(W.r));
        Assert.Equal("REPL", insertedRun.Element(W.t)?.Value);
        Assert.NotNull(insertedRun.Element(W.rPr)?.Element(W.u));
        Assert.Null(insertedRun.Element(W.rPr)?.Element(W.b));
        Assert.Null(insertedRun.Element(W.rPr)?.Element(W.i));
        var suffix = paragraph.Elements(W.r).Single(r => r.Element(W.t)?.Value == "il");
        Assert.NotNull(suffix.Element(W.rPr)?.Element(W.i));

        Assert.Equal("PlREPLil", AcceptedText(tracked));
        Assert.Equal("Plain BOLD tail", RejectedText(tracked));
        AssertSchemaValid(tracked);
    }

    [Fact]
    public void DS402_ReplaceMatch_TrackedRepeatedMatchesWorkInReverseOffsetOrder()
    {
        var source = BuildDocument(new XElement(W.p, Run("one cat, two cat, three cat")));
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var matches = session.Grep("cat").OrderByDescending(m => m.Span.Start).ToArray();
        var replacements = new[] { "C", "B", "A" };

        for (int i = 0; i < matches.Length; i++)
        {
            var result = session.ReplaceMatch(matches[i], replacements[i]);
            Assert.True(result.Success, result.Error?.Message);
        }

        var tracked = session.Save();
        var root = MainRoot(tracked);
        Assert.Equal(3, root.Descendants(W.del).Count());
        Assert.Equal(3, root.Descendants(W.ins).Count());
        Assert.Equal(6, root.Descendants()
            .Where(e => e.Name == W.del || e.Name == W.ins)
            .Select(e => (string?)e.Attribute(W.id)).Distinct().Count());
        Assert.Equal("one A, two B, three C", AcceptedText(tracked));
        Assert.Equal("one cat, two cat, three cat", RejectedText(tracked));
        AssertSchemaValid(tracked);
    }

    [Fact]
    public void DS403_ReplaceTextRange_TrackedIsOneUndoableRedoableOperation()
    {
        var source = BuildDocument(new XElement(W.p, Run("cat cat cat")));
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        Assert.Equal(3, session.ReplaceTextRange(anchor, "cat", "dog").Count);
        Assert.Equal("dog dog dog", AcceptedText(session.Save()));

        Assert.True(session.Undo());
        Assert.Equal("cat cat cat", Text(MainRoot(session.Save())));
        Assert.Empty(MainRoot(session.Save()).Descendants()
            .Where(e => e.Name == W.del || e.Name == W.ins));

        Assert.True(session.Redo());
        Assert.Equal("dog dog dog", AcceptedText(session.Save()));
    }

    [Fact]
    public void DS404_ReplaceInner_TrackedUsesTheSameSurgicalPath()
    {
        var source = BuildDocument(new XElement(W.p, Run("Price: $[___].")));
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var match = Assert.Single(session.Grep(@"\$?\[_+\]"));

        var result = session.ReplaceInner(match, "0.20");

        Assert.True(result.Success, result.Error?.Message);
        var tracked = session.Save();
        var root = MainRoot(tracked);
        Assert.Equal("$[___]", Assert.Single(root.Descendants(W.delText)).Value);
        Assert.Equal("$0.20", Assert.Single(root.Descendants(W.ins).Descendants(W.t)).Value);
        Assert.Equal("Price: $0.20.", AcceptedText(tracked));
        Assert.Equal("Price: $[___].", RejectedText(tracked));
    }

    [Fact]
    public void DS405_TrackedReplacementStaysInsideHyperlinkAndSdtContainers()
    {
        var source = BuildDocument(
            new XElement(W.p,
                new XElement(W.bookmarkStart,
                    new XAttribute(W.id, 0), new XAttribute(W.name, "destination")),
                new XElement(W.bookmarkEnd, new XAttribute(W.id, 0)),
                new XElement(W.hyperlink,
                    new XAttribute(W.anchor, "destination"),
                    Run("linked target", new XElement(W.u, new XAttribute(W.val, "single"))))),
            new XElement(W.p,
                new XElement(W.sdt,
                    new XElement(W.sdtPr,
                        new XElement(W.tag, new XAttribute(W.val, "field"))),
                    new XElement(W.sdtContent, Run("controlled target")))));
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchors = session.Project().AnchorIndex.Values
            .Where(a => a.Anchor.Kind == "p").ToArray();

        Assert.True(Assert.Single(session.ReplaceTextRange(
            anchors[0].Anchor.Id, "target", "value")).Success);
        Assert.True(Assert.Single(session.ReplaceTextRange(
            anchors[1].Anchor.Id, "target", "value")).Success);

        var tracked = session.Save();
        var root = MainRoot(tracked);
        var hyperlink = root.Descendants(W.hyperlink).Single();
        Assert.Single(hyperlink.Elements(W.del));
        Assert.Single(hyperlink.Elements(W.ins));
        var content = root.Descendants(W.sdtContent).Single();
        Assert.Single(content.Elements(W.del));
        Assert.Single(content.Elements(W.ins));
        Assert.Equal("linked valuecontrolled value", AcceptedText(tracked));
        Assert.Equal("linked targetcontrolled target", RejectedText(tracked));
        AssertSchemaValid(tracked);
    }

    [Fact]
    public void DS406_TrackedReplacementRetainsBookmarksCommentsAndNoteReferences()
    {
        var source = BuildMarkerDocument();
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.Project().AnchorIndex.Values
            .Single(a => a.Anchor.Kind == "p" && a.Anchor.Scope == "body").Anchor.Id;

        var result = Assert.Single(session.ReplaceTextRange(
            anchor, "before target", "replacement"));

        Assert.True(result.Success, result.Error?.Message);
        var tracked = session.Save();
        AssertSemanticMarkers(MainRoot(tracked));
        Assert.Equal("replacement", AcceptedText(tracked));
        Assert.Equal("before target", RejectedText(tracked));
        AssertSchemaValid(tracked);

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("accepted.docx", tracked));
        var rejected = RevisionProcessor.RejectRevisions(new WmlDocument("rejected.docx", tracked));
        AssertSemanticMarkers(MainRoot(accepted.DocumentByteArray));
        AssertSemanticMarkers(MainRoot(rejected.DocumentByteArray));
    }

    [Fact]
    public void DS407_DirectModeKeepsTheExistingSurgicalMutationShape()
    {
        var source = BuildDocument(new XElement(W.p, Run("Prefix target suffix")));
        using var session = new DocxSession(source);
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = session.ReplaceTextAtSpan(anchor, 7, 6, "replacement");

        Assert.True(result.Success, result.Error?.Message);
        var root = MainRoot(session.Save());
        Assert.Equal("Prefix replacement suffix", Text(root));
        Assert.Empty(root.Descendants()
            .Where(e => e.Name == W.del || e.Name == W.ins || e.Name == W.delText));
    }

    /// <summary>
    /// Found by the issue #435 acceptance smoke. Text a tracked edit inserts lands inside
    /// <c>w:ins</c>, and <c>InlineRuns</c> only descended into the containers listed in
    /// <c>InlineContainerNames</c> — which omitted <c>w:ins</c>. Every offset-addressed
    /// surface built on it (Grep, ParagraphText, ReplaceTextRange, format-by-substring)
    /// therefore could not see the edit it had just made, while the markdown projection and
    /// the anchor's TextPreview both showed it. An agent could not re-find its own work.
    /// </summary>
    [Fact]
    public void DS409_TrackedInsertedText_IsVisibleToEveryOffsetAddressedSurface()
    {
        var source = BuildDocument(new XElement(W.p, Run("The name is [____].")));
        using var session = new DocxSession(source,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        Assert.True(Assert.Single(session.ReplaceTextRange(anchor, "[____]", "Northstar")).Success);

        // The projection always saw it; these three surfaces did not.
        Assert.Contains("Northstar", session.Project().Markdown, StringComparison.Ordinal);
        var match = Assert.Single(session.Grep("Northstar"));
        Assert.Equal("Northstar", match.Text);
        Assert.Equal(anchor, match.EnclosingAnchor?.Anchor.Id);

        // Flat text is contiguous across the w:del/w:ins pair: deleted text is w:delText and
        // stays out, inserted text is w:t and comes in, so offsets address the visible string.
        Assert.Equal("The name is Northstar.", session.Project().AnchorIndex[anchor].TextPreview);

        // And the inserted text is addressable by a follow-up surgical op in the same session.
        Assert.True(Assert.Single(session.ReplaceTextRange(anchor, "Northstar", "Southstar")).Success);
        Assert.Equal("The name is Southstar.", AcceptedText(session.Save()));
        Assert.Equal("The name is [____].", RejectedText(session.Save()));
        AssertSchemaValid(session.Save());
    }

    /// <summary>
    /// The same blindness applied to revisions already present in the input, so a redline
    /// arriving from Word had its inserted and move-destination spans silently skipped by
    /// every text search. Pins the whole split in one fixture: w:ins and w:moveTo are visible
    /// text, w:del and w:moveFrom are not.
    /// </summary>
    [Fact]
    public void DS410_PreExistingInsertionsAndMoveDestinations_ArePartOfTheVisibleText()
    {
        var source = BuildDocument(new XElement(W.p,
            Run("Stage "),
            new XElement(W.ins,
                new XAttribute(W.id, 1),
                new XAttribute(W.author, "Reviewer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                Run("IV in the chromatogram")),
            new XElement(W.del,
                new XAttribute(W.id, 2),
                new XAttribute(W.author, "Reviewer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                new XElement(W.r,
                    new XElement(W.delText,
                        new XAttribute(Xml + "space", "preserve"), "III"))),
            new XElement(W.moveToRangeStart,
                new XAttribute(W.id, 3), new XAttribute(W.name, "move1")),
            new XElement(W.moveTo,
                new XAttribute(W.id, 4),
                new XAttribute(W.author, "Reviewer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                Run(" per the addendum")),
            new XElement(W.moveToRangeEnd, new XAttribute(W.id, 3)),
            new XElement(W.moveFrom,
                new XAttribute(W.id, 5),
                new XAttribute(W.author, "Reviewer"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                new XElement(W.r,
                    new XElement(W.delText,
                        new XAttribute(Xml + "space", "preserve"), " per the schedule"))),
            Run(" is final.")));
        using var session = new DocxSession(source);
        var anchor = session.Project().AnchorIndex.Values.Single().Anchor.Id;

        Assert.Equal("Stage IV in the chromatogram per the addendum is final.",
            session.Project().AnchorIndex[anchor].TextPreview);

        foreach (var needle in new[] { "chromatogram", "per the addendum" })
        {
            var match = Assert.Single(session.Grep(needle));
            Assert.Equal(anchor, match.EnclosingAnchor?.Anchor.Id);
        }

        // Deleted text is NOT visible text: w:del and w:moveFrom hold w:delText, which never
        // joins the flat string, so a search cannot match what the document says was removed.
        Assert.Empty(session.Grep("III"));
        Assert.Empty(session.Grep("per the schedule"));
    }

    private static byte[] BuildMarkerDocument()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments(
                new Comment(
                    new Paragraph(new Run(new Text("Comment body."))))
                {
                    Id = "0",
                    Author = "Reviewer",
                    Initials = "R",
                });
            commentsPart.Comments.Save();

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new SeparatorMark()))) { Id = -1 },
                new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark()))) { Id = 0 },
                new Footnote(
                    new Paragraph(
                        new Run(new FootnoteReferenceMark()),
                        new Run(new Text("Note body."))))
                {
                    Id = 2,
                });
            footnotesPart.Footnotes.Save();

            main.Document.Save();
            var xDocument = main.GetXDocument();
            xDocument.Root!.Element(W.body)!.ReplaceNodes(
                new XElement(W.p,
                    new XElement(W.bookmarkStart,
                        new XAttribute(W.id, 0), new XAttribute(W.name, "range")),
                    new XElement(W.commentRangeStart, new XAttribute(W.id, 0)),
                    Run("before "),
                    new XElement(W.r,
                        new XElement(W.footnoteReference, new XAttribute(W.id, 2))),
                    Run("target"),
                    new XElement(W.commentRangeEnd, new XAttribute(W.id, 0)),
                    new XElement(W.r,
                        new XElement(W.rPr,
                            new XElement(W.rStyle, new XAttribute(W.val, "CommentReference"))),
                        new XElement(W.commentReference, new XAttribute(W.id, 0))),
                    new XElement(W.bookmarkEnd, new XAttribute(W.id, 0))));
            main.PutXDocument();
        }

        return stream.ToArray();
    }

    private static void AssertSemanticMarkers(XElement root)
    {
        Assert.Single(root.Descendants(W.bookmarkStart));
        Assert.Single(root.Descendants(W.bookmarkEnd));
        Assert.Single(root.Descendants(W.commentRangeStart));
        Assert.Single(root.Descendants(W.commentRangeEnd));
        Assert.Single(root.Descendants(W.commentReference));
        Assert.Single(root.Descendants(W.footnoteReference));
    }
}
