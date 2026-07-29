#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Footnote / endnote <em>authoring</em> on <see cref="DocxSession"/> (issue #276):
/// <see cref="DocxSession.InsertFootnote"/> / <see cref="DocxSession.InsertEndnote"/> create the
/// note definition (creating the <c>FootnotesPart</c>/<c>EndnotesPart</c> plus the two
/// Word-reserved separator notes when absent) and insert the body-side reference run at a
/// character offset. Editing and deleting an authored note goes through the existing
/// <see cref="DocxSession.ReplaceText"/> / <see cref="DocxSession.DeleteBlock"/> paths, which
/// already understand <c>fn</c>/<c>en</c> scopes — DS329 pins that end-to-end.
/// Test IDs use the DS32x range.
/// </summary>
public class DocxSessionNoteAuthoringTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private static XElement PartXml(byte[] docxBytes, Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var part = pick(doc.MainDocumentPart!);
        Assert.NotNull(part);
        return part!.GetXDocument().Root!;
    }

    private static XElement BodyXml(byte[] docxBytes) => PartXml(docxBytes, m => m);

    // ─── Creation: part, boilerplate, reference ─────────────────────────

    [Fact]
    public void DS320_InsertFootnote_CreatesPartWithReservedSeparatorsAndUserNote()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var result = session.InsertFootnote(anchor, 5, "A note about the first five characters.");

        Assert.True(result.Success, result.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        var all = footnotes.Elements(W + "footnote").ToList();

        // The two Word-reserved notes Word always writes into a fresh footnotes part.
        Assert.Contains(all, n => (string?)n.Attribute(W + "type") == "separator");
        Assert.Contains(all, n => (string?)n.Attribute(W + "type") == "continuationSeparator");

        // Plus exactly one user note carrying the payload text.
        var user = all.Where(n => n.Attribute(W + "type") is null).ToList();
        var note = Assert.Single(user);
        Assert.Contains("A note about the first five characters.", note.Descendants(W + "t").Select(t => (string)t));
    }

    [Fact]
    public void DS321_InsertFootnote_InsertsBodyReferenceRunAtOffset()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var before = session.GetBlockMetadata(anchor);
        Assert.NotNull(before);

        var result = session.InsertFootnote(anchor, 5, "Note.");
        Assert.True(result.Success, result.Error?.Message);

        var body = BodyXml(session.Save());
        var para = body.Descendants(W + "p")
            .First(p => p.Descendants(W + "footnoteReference").Any());

        // The reference run sits after exactly 5 characters of the paragraph's text.
        var textBeforeRef = string.Concat(
            para.Descendants(W + "r")
                .TakeWhile(r => !r.Elements(W + "footnoteReference").Any())
                .SelectMany(r => r.Elements(W + "t"))
                .Select(t => (string)t));
        Assert.Equal(5, textBeforeRef.Length);

        // …and it points at the id of the user note that was created.
        var refId = (string?)para.Descendants(W + "footnoteReference").First().Attribute(W + "id");
        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Contains(footnotes.Elements(W + "footnote"),
            n => (string?)n.Attribute(W + "id") == refId && n.Attribute(W + "type") is null);
    }

    [Fact]
    public void DS322_InsertFootnote_AllocatesIdAboveHighestExisting_NotCount()
    {
        // Existing user notes are ids 1, 5 and 9 — a "count + 1" allocator would pick 4 and
        // silently overwrite/alias an existing definition.
        using var session = new DocxSession(BuildDocWithSparseFootnoteIds());
        var anchor = FirstBodyParagraph(session);

        var result = session.InsertFootnote(anchor, 0, "Fresh.");
        Assert.True(result.Success, result.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        var ids = footnotes.Elements(W + "footnote")
            .Select(n => int.Parse((string)n.Attribute(W + "id")!))
            .ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Contains(10, ids);
    }

    [Fact]
    public void DS323_InsertFootnote_SplitsAStraddlingRunAtTheOffset()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var full = string.Concat(
            BodyXml(session.Save()).Descendants(W + "p").First()
                .Descendants(W + "t").Select(t => (string)t));
        Assert.True(full.Length > 4, "fixture paragraph is long enough to split");

        var result = session.InsertFootnote(anchor, 3, "Mid-run.");
        Assert.True(result.Success, result.Error?.Message);

        var body = BodyXml(session.Save());
        var para = body.Descendants(W + "p")
            .First(p => p.Descendants(W + "footnoteReference").Any());

        // Text is preserved verbatim across the split (the ref run itself contributes none).
        var after = string.Concat(para.Descendants(W + "t").Select(t => (string)t));
        Assert.Equal(full, after);
    }

    [Fact]
    public void DS324_InsertFootnote_OffsetOutOfRange_IsRejected()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var tooBig = session.InsertFootnote(anchor, 10_000, "Nope.");
        Assert.False(tooBig.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, tooBig.Error!.Code);

        var negative = session.InsertFootnote(anchor, -1, "Nope.");
        Assert.False(negative.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, negative.Error!.Code);
    }

    [Fact]
    public void DS325_InsertFootnote_RequiresABodyParagraphAnchor()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDocWithFootnotes());

        // A note-scope paragraph is not a legal host: Word does not allow a footnote
        // reference inside a footnote/endnote/header/footer story.
        var notePara = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "fn" && t.Anchor.Kind == "p").Anchor.Id;
        var inNote = session.InsertFootnote(notePara, 0, "Nested.");
        Assert.False(inNote.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, inNote.Error!.Code);

        // Neither is the note definition anchor itself.
        var noteDef = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "fn").Anchor.Id;
        var onDef = session.InsertFootnote(noteDef, 0, "Nested.");
        Assert.False(onDef.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, onDef.Error!.Code);
    }

    [Fact]
    public void DS326_InsertEndnote_CreatesEndnotesPartAndEndnoteReference()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var result = session.InsertEndnote(anchor, 0, "An endnote.");
        Assert.True(result.Success, result.Error?.Message);

        var saved = session.Save();
        var endnotes = PartXml(saved, m => m.EndnotesPart);
        Assert.Contains(endnotes.Elements(W + "endnote"), n => (string?)n.Attribute(W + "type") == "separator");
        var user = Assert.Single(endnotes.Elements(W + "endnote").Where(n => n.Attribute(W + "type") is null));
        Assert.Contains("An endnote.", user.Descendants(W + "t").Select(t => (string)t));

        var body = BodyXml(saved);
        Assert.Single(body.Descendants(W + "endnoteReference"));
        Assert.Empty(body.Descendants(W + "footnoteReference"));
    }

    // ─── Created anchors + projection ───────────────────────────────────

    [Fact]
    public void DS327_InsertFootnote_ReturnsNoteAnchorsAndProjectsTheNote()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var result = session.InsertFootnote(anchor, 0, "Projected note text.");
        Assert.True(result.Success, result.Error?.Message);

        // The note definition anchor and its paragraph anchor both come back as Created,
        // so a caller can immediately address the note for a follow-up edit.
        Assert.Contains(result.Created, a => a.Kind == "fn" && a.Scope == "fn");
        Assert.Contains(result.Created, a => a.Kind == "p" && a.Scope == "fn");
        Assert.Contains(result.Modified, a => a.Id == anchor);

        var markdown = session.Project().Markdown;
        Assert.Contains("# Footnotes", markdown);
        Assert.Contains("Projected note text.", markdown);
    }

    // ─── Undo / redo across the part create ─────────────────────────────

    [Fact]
    public void DS328_Undo_RemovesTheCreatedFootnotesPart_RedoRestoresIt()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.InsertFootnote(anchor, 0, "Undo me.").Success);

        Assert.True(session.Undo());
        using (var ms = new MemoryStream(session.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            Assert.Null(doc.MainDocumentPart!.FootnotesPart);
            Assert.Empty(doc.MainDocumentPart.GetXDocument().Root!.Descendants(W + "footnoteReference"));
        }

        Assert.True(session.Redo());
        using (var ms = new MemoryStream(session.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            Assert.NotNull(doc.MainDocumentPart!.FootnotesPart);
            Assert.Single(doc.MainDocumentPart.GetXDocument().Root!.Descendants(W + "footnoteReference"));
        }
    }

    // ─── Word-faithful markup ───────────────────────────────────────────

    [Fact]
    public void DS329_AuthoredFootnote_IsEditableAndDeletableThroughTheExistingOps()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var created = session.InsertFootnote(anchor, 0, "Original note.");
        Assert.True(created.Success, created.Error?.Message);
        var notePara = created.Created.First(a => a.Kind == "p" && a.Scope == "fn").Id;
        var noteDef = created.Created.First(a => a.Kind == "fn").Id;

        // Edit: ReplaceText already handles fn-scope paragraphs.
        var edited = session.ReplaceText(notePara, "Rewritten note.");
        Assert.True(edited.Success, edited.Error?.Message);
        Assert.Contains("Rewritten note.", session.Project().Markdown);

        // Delete: DeleteBlock removes the definition AND the body-side reference.
        var deleted = session.DeleteBlock(noteDef);
        Assert.True(deleted.Success, deleted.Error?.Message);
        var body = BodyXml(session.Save());
        Assert.Empty(body.Descendants(W + "footnoteReference"));
    }

    [Fact]
    public void DS330_AuthoredFootnote_CarriesWordsNoteRefMarkAndStyles()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.InsertFootnote(anchor, 0, "Styled.").Success);

        var saved = session.Save();

        // Note body: first paragraph opens with the auto-numbering mark, styled FootnoteReference.
        var footnotes = PartXml(saved, m => m.FootnotesPart);
        var user = footnotes.Elements(W + "footnote").First(n => n.Attribute(W + "type") is null);
        var firstPara = user.Elements(W + "p").First();
        Assert.Single(firstPara.Descendants(W + "footnoteRef"));
        Assert.Equal("FootnoteText", (string?)firstPara.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));

        // Body-side reference run carries the FootnoteReference character style.
        var body = BodyXml(saved);
        var refRun = body.Descendants(W + "r").First(r => r.Elements(W + "footnoteReference").Any());
        Assert.Equal("FootnoteReference",
            (string?)refRun.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val"));

        // Both styles are actually defined, so the reference is not a phantom.
        var styles = PartXml(saved, m => m.StyleDefinitionsPart);
        var ids = styles.Elements(W + "style").Select(s => (string?)s.Attribute(W + "styleId")).ToList();
        Assert.Contains("FootnoteReference", ids);
        Assert.Contains("FootnoteText", ids);
    }

    [Fact]
    public void DS331_CreatingTheFootnotesPart_DeclaresTheSeparatorsInSettings()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.InsertFootnote(anchor, 0, "Settings.").Success);

        var settings = PartXml(session.Save(), m => m.DocumentSettingsPart);
        var fnPr = settings.Element(W + "footnotePr");
        Assert.NotNull(fnPr);
        var declared = fnPr!.Elements(W + "footnote")
            .Select(f => (string?)f.Attribute(W + "id"))
            .ToList();
        Assert.Contains("-1", declared);
        Assert.Contains("0", declared);
    }

    [Fact]
    public void DS332_SecondInsert_ReusesThePartWithoutDuplicatingSeparators()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.InsertFootnote(anchor, 0, "First.").Success);
        Assert.True(session.InsertFootnote(anchor, 1, "Second.").Success);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Single(footnotes.Elements(W + "footnote").Where(n => (string?)n.Attribute(W + "type") == "separator"));
        Assert.Single(footnotes.Elements(W + "footnote").Where(n => (string?)n.Attribute(W + "type") == "continuationSeparator"));
        Assert.Equal(2, footnotes.Elements(W + "footnote").Count(n => n.Attribute(W + "type") is null));

        // Distinct ids, both cited from the body.
        var body = BodyXml(session.Save());
        var refIds = body.Descendants(W + "footnoteReference")
            .Select(r => (string?)r.Attribute(W + "id")).ToList();
        Assert.Equal(2, refIds.Distinct().Count());
    }

    [Fact]
    public void DS333_NotePayload_SupportsTheMarkdownSubset()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.InsertFootnote(anchor, 0, "See **Smith** v. *Jones*.").Success);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        var user = footnotes.Elements(W + "footnote").First(n => n.Attribute(W + "type") is null);
        Assert.Contains(user.Descendants(W + "r"),
            r => r.Element(W + "rPr")?.Element(W + "b") is not null
                 && r.Elements(W + "t").Any(t => (string)t == "Smith"));
        Assert.Contains(user.Descendants(W + "r"),
            r => r.Element(W + "rPr")?.Element(W + "i") is not null
                 && r.Elements(W + "t").Any(t => (string)t == "Jones"));
    }

    [Fact]
    public void DS334_AuthoredNotes_ProduceASchemaValidDocument()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.InsertFootnote(anchor, 0, "A footnote.").Success);
        Assert.True(session.InsertEndnote(anchor, 4, "An **endnote**.").Success);

        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator()
            .Validate(doc)
            .Select(e => $"{e.Part?.Uri}: {e.Description}")
            .ToList();
        Assert.Empty(errors);
    }

    /// <summary>
    /// Footnotes part whose user notes are ids 1, 5 and 9 (non-contiguous), so a "count + 1"
    /// id allocator would collide. Body cites all three.
    /// </summary>
    private static byte[] BuildDocWithSparseFootnoteIds()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Body cites three footnotes."),
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 },
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 5 },
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 9 }))));

            var fnPart = main.AddNewPart<FootnotesPart>();
            var fnXml = """
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="1"><w:p><w:r><w:t>One.</w:t></w:r></w:p></w:footnote>
                  <w:footnote w:id="5"><w:p><w:r><w:t>Five.</w:t></w:r></w:p></w:footnote>
                  <w:footnote w:id="9"><w:p><w:r><w:t>Nine.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """;
            using var s = fnPart.GetStream(FileMode.Create);
            using var w = new StreamWriter(s);
            w.Write(fnXml);
        }
        return ms.ToArray();
    }
}
