using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Session edits that copy a run's or a paragraph's properties must not copy the tracked-change
/// markers inside them verbatim: two live markers sharing a <c>w:id</c> make every revision that
/// carries the id ambiguous, so none of them can be accepted or rejected on its own (issue #1297).
/// Each test pins the ids AND the accept-all / reject-all outcome, because unique ids on the wrong
/// half of a split would pass an id-only check while reject undid the wrong thing.
/// </summary>
public class DocxSessionClonedRevisionTests
{
    private const string Alice = "w:author=\"Alice\" w:date=\"2025-01-02T09:00:00Z\"";

    /// <summary>One run whose bold is Alice's pending format change.</summary>
    private const string PendingBoldRun =
        "<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id=\"40\" " + Alice + "><w:rPr/></w:rPrChange></w:rPr>" +
        "<w:t>abcdefghi</w:t></w:r></w:p>";

    /// <summary>A centred paragraph whose mark Alice inserted and whose centring she has pending,
    /// followed by an ordinary paragraph.</summary>
    private const string InsertedMarkParagraph =
        "<w:p><w:pPr><w:jc w:val=\"center\"/>" +
        "<w:rPr><w:ins w:id=\"50\" " + Alice + "/></w:rPr>" +
        "<w:pPrChange w:id=\"51\" " + Alice + "><w:pPr/></w:pPrChange></w:pPr>" +
        "<w:r><w:t>first second</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>after</w:t></w:r></w:p>";

    [Fact]
    public void RID001_PartialRunFormat_GivesEachSliceOfAPendingFormatChangeItsOwnId()
    {
        using var session = Open(PendingBoldRun, TrackedChangeMode.RenderInline);

        var result = session.ApplyFormat(FirstParagraph(session), new CharSpan(3, 3), new FormatOp { Italic = true });

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 3);
        AssertOutcomes(session,
            accepted: new[] { "p: abc{b}|def{b,i}|ghi{b}" },
            rejected: new[] { "p: abc|def|ghi" });
    }

    [Fact]
    public void RID002_TrackedReplaceInsideRun_GivesEveryCopiedFormatChangeItsOwnId()
    {
        using var session = Open(PendingBoldRun, TrackedChangeMode.RenderInline);

        var result = session.ReplaceTextAtSpan(FirstParagraph(session), 3, 3, "XYZ");

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 6);
        AssertOutcomes(session,
            accepted: new[] { "p: abc{b}|XYZ{b}|ghi{b}" },
            rejected: new[] { "p: abc|def|ghi" });
    }

    /// <summary>
    /// An untracked split adds a paragraph mark that is ordinary content. The mark Alice
    /// inserted is the one that still ends "second", so it stays on the second half; copying it
    /// onto the new mark would let reject-all undo the untracked split too. Her pending centring
    /// covers both halves, each under its own id.
    /// </summary>
    [Fact]
    public void RID003_UntrackedSplit_KeepsTheInsertedMarkOnTheHalfItStillEnds()
    {
        using var session = Open(InsertedMarkParagraph, TrackedChangeMode.Accept);

        var result = session.SplitParagraph(FirstParagraph(session), 6);

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 3);
        var paragraphs = Body(session).Elements(W.p).ToList();
        Assert.Null(paragraphs[0].Element(W.pPr)?.Element(W.rPr)?.Element(W.ins));
        // The second half keeps the original mark, so its revisions keep their ids.
        Assert.Equal("50", (string?)paragraphs[1].Element(W.pPr)?.Element(W.rPr)?.Element(W.ins)?.Attribute(W.id));
        Assert.Equal("51", (string?)paragraphs[1].Element(W.pPr)?.Element(W.pPrChange)?.Attribute(W.id));
        AssertOutcomes(session,
            accepted: new[] { "p[center]: first ", "p[center]: second", "p: after" },
            rejected: new[] { "p: first ", "p: second|after" });
    }

    [Fact]
    public void RID004_UntrackedSplitAtTheEnd_MovesTheInsertedMarkToTheNewParagraph()
    {
        using var session = Open(InsertedMarkParagraph, TrackedChangeMode.Accept);

        var result = session.SplitParagraph(FirstParagraph(session), "first second".Length);

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 3);
        AssertOutcomes(session,
            accepted: new[] { "p[center]: first second", "p[center]:", "p: after" },
            rejected: new[] { "p: first second", "p: after" });
    }

    /// <summary>
    /// Enter at the end of a heading starts a paragraph in the heading's next style, and the
    /// heading's original mark, the one Alice inserted, now ends that paragraph. Rejecting her
    /// insertion joins the empty paragraph to the next one, so the heading keeps its style;
    /// leaving her mark on the heading would make the reject restyle the heading instead.
    /// </summary>
    [Fact]
    public void RID007_UntrackedEnterAfterAHeading_KeepsTheInsertedMarkWithTheNextStyleParagraph()
    {
        using var session = Open(
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/><w:rPr><w:ins w:id=\"50\" " + Alice + "/></w:rPr></w:pPr>" +
            "<w:r><w:t>Terms</w:t></w:r></w:p><w:p><w:r><w:t>after</w:t></w:r></w:p>",
            TrackedChangeMode.Accept, stylesXml: HeadingStyles);

        var result = session.SplitParagraph(FirstParagraph(session), "Terms".Length);

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 1);
        AssertOutcomes(session,
            accepted: new[] { "p[Heading1]: Terms", "p[Normal]:", "p: after" },
            rejected: new[] { "p[Heading1]: Terms", "p: after" });
    }

    [Fact]
    public void RID005_TrackedMove_GivesTheDestinationCopyItsOwnIds()
    {
        using var session = Open(PendingBoldRun + "<w:p><w:r><w:t>after</w:t></w:r></w:p>",
            TrackedChangeMode.RenderInline);
        var anchors = session.Project().AnchorIndex.Values.Select(v => v.Anchor).Where(a => a.Kind == "p")
            .Select(a => a.Id).ToList();

        var result = session.MoveBlock(anchors[0], anchors[1], Position.After);

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: null);
        AssertOutcomes(session,
            accepted: new[] { "p: after", "p: abcdefghi{b}" },
            rejected: new[] { "p: abcdefghi", "p: after" });
    }

    [Fact]
    public void RID006_TrackedImageReplace_GivesTheInsertedCopyItsOwnIds()
    {
        // Alice's pending bold covers a picture run; Bob then swaps the picture under tracking,
        // which deletes the old run and inserts a copy of it.
        byte[] inserted;
        using (var author = Open("<w:p><w:r><w:t>abc</w:t></w:r></w:p>", TrackedChangeMode.Accept))
        {
            Assert.True(author.InsertImage(FirstParagraph(author), 3, DocxSessionImageTests.Png(2, 3)).Success);
            inserted = author.Save();
        }
        var reviewed = WithMainDocument(inserted, body =>
        {
            var properties = body.Descendants(W.drawing).Single().Parent!.Element(W.rPr)!;
            properties.Add(XElement.Parse(
                $"<w:rPrChange xmlns:w=\"{W.w}\" w:id=\"40\" {Alice}><w:rPr><w:noProof/></w:rPr></w:rPrChange>"));
            properties.AddFirst(new XElement(W.b));
        });
        using var bob = new DocxSession(reviewed, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Bob",
        });

        var result = bob.ReplaceImage(Assert.Single(bob.ListImages()).Id, DocxSessionImageTests.Png(4, 5));

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(bob, expectedCount: 4);
        Assert.True(bob.RejectAllRevisions().Success);
        Assert.Null(Body(bob).Descendants(W.drawing).Single().Parent!.Element(W.rPr)!.Element(W.b));
    }

    [Fact]
    public void RID008_TrackedInlineControlFill_GivesTheInsertedTextItsOwnIds()
    {
        using var session = Open(
            "<w:p><w:sdt><w:sdtPr><w:id w:val=\"7\"/><w:text/></w:sdtPr><w:sdtContent>" +
            "<w:r><w:rPr><w:b/><w:rPrChange w:id=\"40\" " + Alice + "><w:rPr/></w:rPrChange></w:rPr>" +
            "<w:t>old</w:t></w:r></w:sdtContent></w:sdt></w:p>",
            TrackedChangeMode.RenderInline);

        var result = session.FillContentControlText(Assert.Single(session.ListContentControls()).AnchorId, "new");

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 4);
        AssertOutcomes(session, accepted: new[] { "p: new{b}" }, rejected: new[] { "p: old" });
    }

    [Fact]
    public void RID009_InsertedRow_GivesCopiedCellPropertyChangesTheirOwnIds()
    {
        using var session = Open(
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:fill=\"FF0000\"/>" +
            "<w:tcPrChange w:id=\"60\" " + Alice + "><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr></w:tcPrChange></w:tcPr>" +
            "<w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>",
            TrackedChangeMode.Accept);
        var cell = session.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "tc").Id;

        var result = session.InsertTableRow(cell, Position.After);

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: 2);
        AssertOutcomes(session, accepted: new[] { "FF0000", "FF0000" }, rejected: new[] { "none", "none" },
            describe: body => body.Descendants(W.tc)
                .Select(tc => (string?)tc.Element(W.tcPr)?.Element(W.shd)?.Attribute(W.fill) ?? "none").ToArray());
    }

    [Fact]
    public void RID018_TrackedMove_GivesTheCopiedCommentsRevisionsTheirOwnIds()
    {
        byte[] commented;
        using (var author = Open("<w:p><w:r><w:t>moved</w:t></w:r></w:p><w:p><w:r><w:t>after</w:t></w:r></w:p>",
                   TrackedChangeMode.Accept, "Alice"))
        {
            Assert.True(author.AddComment(FirstParagraph(author), null, "Alice", "note").Success);
            commented = author.Save();
        }
        var withRevision = WithPart(commented, document => document.MainDocumentPart!.WordprocessingCommentsPart!,
            root => root.Descendants(W.p).First().Add(XElement.Parse(
                $"<w:ins xmlns:w=\"{W.w}\" w:id=\"80\" {Alice}><w:r><w:t> extra</w:t></w:r></w:ins>")));
        using var bob = new DocxSession(withRevision, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Bob",
        });
        var paragraphs = bob.Project().AnchorIndex.Values.Select(v => v.Anchor)
            .Where(a => a.Kind == "p" && a.Scope == "body").Select(a => a.Id).ToList();

        var result = bob.MoveBlock(paragraphs[0], paragraphs[1], Position.After);

        Assert.True(result.Success, result.Error?.Message);
        var commentInsertions = Part(bob.Save(), document => document.MainDocumentPart!.WordprocessingCommentsPart!)
            .Descendants(W.ins).Select(ins => (string?)ins.Attribute(W.id)).ToList();
        Assert.Equal(2, commentInsertions.Count);
        Assert.Equal(2, commentInsertions.Distinct().Count());
        Assert.All(bob.ListRevisions(),
            revision => Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus));
    }

    /// <summary>Raw XML can bring in revision ids above the session's counter; the next id the
    /// session mints must still be unused.</summary>
    [Fact]
    public void RID019_RawMarkupAfterTheCounterIsSeeded_IsNotReusedByTheNextCopy()
    {
        using var session = Open(PendingBoldRun + "<w:p><w:r><w:t>zzz</w:t></w:r></w:p>", TrackedChangeMode.Accept);
        var paragraphs = session.Project().AnchorIndex.Values.Select(v => v.Anchor)
            .Where(a => a.Kind == "p").Select(a => a.Id).ToList();
        Assert.True(session.ApplyFormat(paragraphs[0], new CharSpan(3, 3), new FormatOp { Italic = true }).Success);
        // The counter is now seeded and has handed out 1001 and 1002 to the split slices; occupy
        // the next ids it would hand out. (Nothing may save the session in between: a save
        // re-seeds the counter and would hide the defect.)
        var runs = string.Concat(Enumerable.Range(1003, 10).Select(id =>
            $"<w:r><w:rPr><w:b/><w:rPrChange w:id=\"{id}\" w:author=\"Carol\" w:date=\"2025-01-02T09:00:00Z\">" +
            "<w:rPr/></w:rPrChange></w:rPr><w:t>raw</w:t></w:r>"));
        Assert.True(session.Raw.InsertXml(paragraphs[1], Position.After, $"<w:p xmlns:w=\"{W.w}\">{runs}</w:p>").Success);

        var result = session.SplitParagraph(paragraphs[0], 1);

        Assert.True(result.Success, result.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: null);
    }

    // ─── Shared fixture plumbing ──────────────────────────────────────────

    /// <summary>A heading whose next style is Normal, as in Word's built-in styles.</summary>
    internal const string HeadingStyles =
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/>" +
        "<w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/></w:style>";

    internal static DocxSession Open(string bodyXml, TrackedChangeMode mode, string author = "Bob",
        string? stylesXml = null)
    {
        var bytes = IrTestDocuments.FromBodyXml(bodyXml).DocumentByteArray;
        if (stylesXml is not null)
            bytes = WithPart(bytes, document => document.MainDocumentPart!.StyleDefinitionsPart!,
                root => root.Add(XElement.Parse($"<w:styles xmlns:w=\"{W.w}\">{stylesXml}</w:styles>").Elements()));
        return new DocxSession(bytes, new DocxSessionSettings
        {
            TrackedChanges = mode,
            RevisionAuthor = author,
        });
    }

    internal static string FirstParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind is "p" or "h").Id;

    internal static XElement Body(DocxSession session) => Body(session.Save());

    private static byte[] WithMainDocument(byte[] bytes, Action<XElement> editBody) =>
        WithPart(bytes, document => document.MainDocumentPart!, root => editBody(root.Element(W.body)!));

    private static byte[] WithPart(byte[] bytes, Func<WordprocessingDocument, OpenXmlPart> part,
        Action<XElement> editRoot)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var target = part(document);
            XDocument xml;
            using (var reader = target.GetStream()) xml = XDocument.Load(reader);
            editRoot(xml.Root!);
            using var writer = target.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(writer);
        }
        return stream.ToArray();
    }

    private static XElement Body(byte[] bytes) =>
        Part(bytes, document => document.MainDocumentPart!).Element(W.body)!;

    private static XElement Part(byte[] bytes, Func<WordprocessingDocument, OpenXmlPart> part)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        using var reader = part(document).GetStream();
        return XDocument.Load(reader).Root!;
    }

    /// <summary>Every revision marker in the body has its own id, and the session can resolve
    /// each listed revision on its own.</summary>
    internal static void AssertRevisionIdsDistinct(DocxSession session, int? expectedCount)
    {
        var ids = Body(session).Descendants()
            // A range end repeats its start's id by design; it is not a second marker.
            .Where(e => Docxodus.Internal.RevisionOps.RevisionIdBearingNames.Contains(e.Name)
                && !e.Name.LocalName.EndsWith("RangeEnd", StringComparison.Ordinal))
            .Select(e => $"{e.Name.LocalName}:{(string?)e.Attribute(W.id)}")
            .ToList();
        if (expectedCount is { } count) Assert.Equal(count, ids.Count);
        var numbers = ids.Select(id => id[(id.IndexOf(':') + 1)..]).ToList();
        Assert.True(numbers.Distinct().Count() == numbers.Count, "duplicate revision ids: " + string.Join(", ", ids));
        Assert.All(session.ListRevisions(),
            revision => Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus));
    }

    /// <summary>Accepting and rejecting the edited document, all at once and one revision at a
    /// time, produce the expected bodies.</summary>
    internal static void AssertOutcomes(DocxSession session,
        string[] accepted, string[] rejected, Func<XElement, string[]>? describe = null)
    {
        var edited = session.Save();
        describe ??= body => body.Elements(W.p).Select(Describe).ToArray();
        Assert.Equal(accepted, describe(Resolve(edited, accept: true, oneAtATime: false)));
        Assert.Equal(accepted, describe(Resolve(edited, accept: true, oneAtATime: true)));
        Assert.Equal(rejected, describe(Resolve(edited, accept: false, oneAtATime: false)));
        Assert.Equal(rejected, describe(Resolve(edited, accept: false, oneAtATime: true)));
    }

    private static XElement Resolve(byte[] bytes, bool accept, bool oneAtATime)
    {
        using var session = new DocxSession(bytes);
        if (oneAtATime)
        {
            // Resolve in listed order, re-listing after each step: a resolution may renumber
            // the revisions that remain, but it must never make one unresolvable.
            for (var listed = session.ListRevisions(); listed.Count > 0; listed = session.ListRevisions())
            {
                var step = accept ? session.AcceptRevision(listed[0].Id) : session.RejectRevision(listed[0].Id);
                Assert.True(step.Success, $"{listed[0].Id}: {step.Error?.Code} {step.Error?.Message}");
            }
        }
        else
        {
            var all = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
            Assert.True(all.Success, all.Error?.Message);
        }
        Assert.Empty(session.ListRevisions());
        return Body(session);
    }

    /// <summary>"p[Heading1,center]: abc{b,i}|def": the style and justification, then each run's text with the
    /// formatting toggles the tests change; nothing else in the resolved body matters here.</summary>
    private static string Describe(XElement paragraph)
    {
        var properties = new[] { W.pStyle, W.jc }
            .Select(name => (string?)paragraph.Element(W.pPr)?.Element(name)?.Attribute(W.val))
            .OfType<string>().ToList();
        var runs = paragraph.Descendants(W.r).Select(run =>
        {
            var text = string.Concat(run.Elements(W.t).Select(t => t.Value));
            var toggles = new[] { W.b, W.i }.Where(name => run.Element(W.rPr)?.Element(name) is not null)
                .Select(name => name.LocalName).ToList();
            return toggles.Count == 0 ? text : $"{text}{{{string.Join(",", toggles)}}}";
        }).Where(text => text.Length > 0);
        var head = properties.Count == 0 ? "p:" : $"p[{string.Join(",", properties)}]:";
        var body = string.Join("|", runs);
        return body.Length == 0 ? head : $"{head} {body}";
    }
}
