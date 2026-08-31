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

    [Fact]
    public void DS335_Undo_OfASecondNote_KeepsThePartAndRollsBackOnlyThatDefinition()
    {
        // The other half of the note-part reconcile: when the snapshot HAS the part, undo must
        // leave it alone and let content-restore drop just the second definition.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.InsertFootnote(anchor, 0, "Keep me.").Success);
        Assert.True(session.InsertFootnote(anchor, 1, "Roll me back.").Success);

        Assert.True(session.Undo());

        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        Assert.NotNull(main.FootnotesPart);

        var user = main.FootnotesPart!.GetXDocument().Root!
            .Elements(W + "footnote")
            .Where(n => n.Attribute(W + "type") is null)
            .ToList();
        var surviving = Assert.Single(user);
        Assert.Contains("Keep me.", surviving.Descendants(W + "t").Select(t => (string)t));
        Assert.DoesNotContain("Roll me back.", surviving.Descendants(W + "t").Select(t => (string)t));
        Assert.Single(main.GetXDocument().Root!.Descendants(W + "footnoteReference"));
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

    /// <summary>
    /// Note ids must ascend in REFERENCE order, the invariant every Word-authored document holds
    /// (verified across the TestFiles corpus, gaps and all). Renderers depend on it: LibreOffice
    /// numbers the body markers by citation position but pairs them with the id-sorted definition
    /// list, so a first-cited note holding the highest id silently renders the WRONG note text —
    /// the marker says "1" and points at somebody else's footnote. Allocating max(id)+1 is only
    /// correct when the new citation follows every existing one.
    /// </summary>
    [Fact]
    public void DS336_InsertingBeforeExistingCitations_KeepsIdsAscendingInReferenceOrder()
    {
        using var session = new DocxSession(BuildDocWithTwoCitedFootnotes());
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && (t.TextPreview ?? "").StartsWith("Alpha")).Anchor.Id;

        // Cite a new note from the FIRST paragraph — ahead of both existing citations.
        Assert.True(session.InsertFootnote(first, 5, "BRAND NEW.").Success);

        var body = BodyXml(session.Save());
        var refsInDocumentOrder = body.Descendants(W + "footnoteReference")
            .Select(r => int.Parse((string)r.Attribute(W + "id")!))
            .ToList();

        Assert.Equal(3, refsInDocumentOrder.Count);
        Assert.Equal(refsInDocumentOrder.OrderBy(i => i).ToList(), refsInDocumentOrder);

        // …and the new note's id still resolves to the new note's text, not a shifted neighbour.
        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        var newNote = footnotes.Elements(W + "footnote")
            .First(n => (string?)n.Attribute(W + "id") == refsInDocumentOrder[0].ToString());
        Assert.Contains("BRAND NEW.", newNote.Descendants(W + "t").Select(t => (string)t));

        // Every id is still unique and every citation still resolves to a definition.
        var defIds = footnotes.Elements(W + "footnote")
            .Select(n => (string)n.Attribute(W + "id")!).ToList();
        Assert.Equal(defIds.Count, defIds.Distinct().Count());
        foreach (var r in refsInDocumentOrder) Assert.Contains(r.ToString(), defIds);
    }

    [Fact]
    public void DS337_ShiftedNotes_KeepTheirOwnText()
    {
        using var session = new DocxSession(BuildDocWithTwoCitedFootnotes());
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && (t.TextPreview ?? "").StartsWith("Alpha")).Anchor.Id;
        Assert.True(session.InsertFootnote(first, 5, "BRAND NEW.").Success);

        var saved = session.Save();
        var body = BodyXml(saved);
        var footnotes = PartXml(saved, m => m.FootnotesPart);
        string TextOf(string id) => string.Concat(footnotes.Elements(W + "footnote")
            .First(n => (string?)n.Attribute(W + "id") == id)
            .Descendants(W + "t").Select(t => (string)t));

        // Walk each citing paragraph and assert its citation resolves to the right note.
        var paras = body.Descendants(W + "p")
            .Where(p => p.Descendants(W + "footnoteReference").Any()).ToList();
        foreach (var p in paras)
        {
            var text = string.Concat(p.Descendants(W + "t").Select(t => (string)t));
            var id = (string)p.Descendants(W + "footnoteReference").First().Attribute(W + "id")!;
            var expected = text.StartsWith("Alpha") ? "BRAND NEW."
                         : text.StartsWith("Beta") ? "EXISTING ONE."
                         : "EXISTING TWO.";
            Assert.Contains(expected, TextOf(id));
        }
    }

    /// <summary>
    /// The renumbering shift has to reach references that live OUTSIDE the main document part.
    /// An endnote can be cited from inside a footnote body, so inserting an endnote ahead of that
    /// citation must renumber the reference sitting in <c>footnotes.xml</c> too — and flush that
    /// part. If it didn't, the footnote's citation would silently point at the wrong endnote.
    /// </summary>
    [Fact]
    public void DS338_ShiftRenumbersNoteReferencesInPeerParts()
    {
        using var session = new DocxSession(BuildDocWithEndnoteCitedFromAFootnote());
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && (t.TextPreview ?? "").StartsWith("Alpha")).Anchor.Id;

        // Insert an endnote in the FIRST paragraph — ahead of the body's existing endnote citation.
        Assert.True(session.InsertEndnote(first, 5, "NEW ENDNOTE.").Success);

        var saved = session.Save();
        var endnotes = PartXml(saved, m => m.EndnotesPart);
        string EndnoteText(string id) => string.Concat(endnotes.Elements(W + "endnote")
            .First(n => (string?)n.Attribute(W + "id") == id)
            .Descendants(W + "t").Select(t => (string)t));

        // The body's own citation still resolves to the endnote it always meant.
        var body = BodyXml(saved);
        var bodyRefs = body.Descendants(W + "endnoteReference")
            .Select(r => (string)r.Attribute(W + "id")!).ToList();
        Assert.Equal(2, bodyRefs.Count);
        Assert.Contains("NEW ENDNOTE.", EndnoteText(bodyRefs[0]));
        Assert.Contains("BODY-CITED ENDNOTE.", EndnoteText(bodyRefs[1]));

        // …and so does the one buried inside the footnote body, in a different part.
        var footnotes = PartXml(saved, m => m.FootnotesPart);
        var inNote = footnotes.Descendants(W + "endnoteReference")
            .Select(r => (string)r.Attribute(W + "id")!).ToList();
        var citedFromFootnote = Assert.Single(inNote);
        Assert.Contains("NOTE-CITED ENDNOTE.", EndnoteText(citedFromFootnote));
    }

    /// <summary>
    /// Two endnotes: one cited from the body, one cited from inside a footnote's body — the
    /// cross-part citation the renumbering shift has to follow.
    /// </summary>
    private static byte[] BuildDocWithEndnoteCitedFromAFootnote()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Alpha paragraph with no note."))),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Beta cites a footnote and an endnote.")),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 }),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.EndnoteReference { Id = 1 }))));

            // The footnote's body cites endnote 2 — a reference in a peer part.
            var fnPart = main.AddNewPart<FootnotesPart>();
            using (var s = fnPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
                w.Write("""
                    <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                      <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                      <w:footnote w:id="1"><w:p><w:r><w:t>Footnote body.</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p></w:footnote>
                    </w:footnotes>
                    """);

            var enPart = main.AddNewPart<EndnotesPart>();
            using (var s = enPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
                w.Write("""
                    <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
                      <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
                      <w:endnote w:id="1"><w:p><w:r><w:t>BODY-CITED ENDNOTE.</w:t></w:r></w:p></w:endnote>
                      <w:endnote w:id="2"><w:p><w:r><w:t>NOTE-CITED ENDNOTE.</w:t></w:r></w:p></w:endnote>
                    </w:endnotes>
                    """);
        }
        return ms.ToArray();
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
    /// Regression found while reviewing PR #491. Once tracked insertions joined the visible-text
    /// walk, note offsets could resolve inside <c>w:ins</c>/<c>w:moveTo</c>. The run splitter cut
    /// the nested run, but the top-level inserter treated the revision wrapper as atomic and
    /// silently placed the citation after the entire revision. Refuse that boundary before the
    /// history snapshot instead of reporting a successful wrong-location edit.
    /// </summary>
    [Theory]
    [InlineData("ins", true)]
    [InlineData("moveTo", false)]
    public void DS339_InsertNote_InsideVisibleRevision_FailsClosedWithoutMutation(
        string revisionName, bool footnote)
    {
        using var session = new DocxSession(BuildDocWithVisibleRevision(revisionName));
        var anchor = FirstBodyParagraph(session);
        var before = session.GetPackageContentHash();
        var version = session.Version;
        var undoCount = session.UndoCount;
        var redoCount = session.RedoCount;

        var result = footnote
            ? session.InsertFootnote(anchor, 4, "Must not be misplaced.")
            : session.InsertEndnote(anchor, 4, "Must not be misplaced.");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.Equal(version, session.Version);
        Assert.Equal(undoCount, session.UndoCount);
        Assert.Equal(redoCount, session.RedoCount);

        using var stream = new MemoryStream(session.Save());
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Null(document.MainDocumentPart!.FootnotesPart);
        Assert.Null(document.MainDocumentPart.EndnotesPart);
        Assert.Empty(document.MainDocumentPart.GetXDocument().Descendants(W + "footnoteReference"));
        Assert.Empty(document.MainDocumentPart.GetXDocument().Descendants(W + "endnoteReference"));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(6)]
    public void DS340_InsertFootnote_AtVisibleRevisionEdge_UsesTheExactOffset(int offset)
    {
        using var session = new DocxSession(BuildDocWithVisibleRevision("ins"));
        var anchor = FirstBodyParagraph(session);

        var result = session.InsertFootnote(anchor, offset, "Edge note.");

        Assert.True(result.Success, result.Error?.Message);
        var paragraph = BodyXml(session.Save()).Descendants(W + "p")
            .Single(p => p.Descendants(W + "footnoteReference").Any());
        var reference = paragraph.Descendants(W + "footnoteReference").Single();
        var referenceRun = reference.Parent!;
        var textBeforeReference = string.Concat(paragraph.Descendants(W + "r")
            .TakeWhile(run => !ReferenceEquals(run, referenceRun))
            .SelectMany(run => run.Elements(W + "t"))
            .Select(text => (string)text));
        Assert.Equal(offset, textBeforeReference.Length);
    }

    private static byte[] BuildDocWithVisibleRevision(string revisionName)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles =
                new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings =
                new DocumentFormat.OpenXml.Wordprocessing.Settings();
            main.Document.Save();

            var revision = new XElement(W + revisionName,
                new XAttribute(W + "id", "10"),
                new XAttribute(W + "author", "Reviewer"),
                new XAttribute(W + "date", "2026-01-01T00:00:00Z"),
                new XElement(W + "r", new XElement(W + "t", "BBBB")));
            var paragraphChildren = new System.Collections.Generic.List<object>
            {
                new XElement(W + "r", new XElement(W + "t", "AA")),
            };
            if (revisionName == "moveTo")
                paragraphChildren.Add(new XElement(W + "moveToRangeStart",
                    new XAttribute(W + "id", "9"), new XAttribute(W + "name", "move9")));
            paragraphChildren.Add(revision);
            if (revisionName == "moveTo")
                paragraphChildren.Add(new XElement(W + "moveToRangeEnd", new XAttribute(W + "id", "9")));
            paragraphChildren.Add(new XElement(W + "r", new XElement(W + "t", "CC")));

            var xDocument = main.GetXDocument();
            xDocument.Root!.Element(W + "body")!.ReplaceNodes(new XElement(W + "p", paragraphChildren));
            main.PutXDocument();
        }
        return stream.ToArray();
    }

    /// <summary>
    /// Three paragraphs; the second and third each cite a footnote (ids 1 and 2, ascending in
    /// reference order as every Word document does). The first cites nothing — inserting there
    /// puts a new citation AHEAD of both existing ones.
    /// </summary>
    private static byte[] BuildDocWithTwoCitedFootnotes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Alpha paragraph with no note."))),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Beta cites one.")),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 })),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Gamma cites two.")),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 2 }))));

            var fnPart = main.AddNewPart<FootnotesPart>();
            var fnXml = """
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="1"><w:p><w:r><w:t>EXISTING ONE.</w:t></w:r></w:p></w:footnote>
                  <w:footnote w:id="2"><w:p><w:r><w:t>EXISTING TWO.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """;
            using var s = fnPart.GetStream(FileMode.Create);
            using var w = new StreamWriter(s);
            w.Write(fnXml);
        }
        return ms.ToArray();
    }

    // ─── Recording as a tracked change (issue #614) ──────────────────────
    //
    // Under TrackedChangeMode.RenderInline the CITATION is the reversible unit and the definition
    // follows it: rejecting the w:ins takes the reference away, which leaves the definition
    // unreferenced, and the note-lifecycle rule in Internal.NoteReferenceOps removes it in the same
    // resolve. Before #614 the op ignored the recording mode entirely — it wrote an untracked
    // citation and an untracked definition, so reject-all had nothing to reject and the "rejected"
    // document still shipped the note text.

    /// <summary>Insert one note under recording and return (baseline, redline).</summary>
    private static (byte[] Baseline, byte[] Redline) AuthorNoteUnderRecording(bool footnote)
    {
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(baseline, new DocxSessionSettings
        {
            PersistAnchorIds = false,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Note Author",
        });
        var anchor = FirstBodyParagraph(session);
        var result = footnote
            ? session.InsertFootnote(anchor, 5, "Negotiated on March 11.")
            : session.InsertEndnote(anchor, 5, "Negotiated on March 11.");
        Assert.True(result.Success, result.Error?.Message);
        return (baseline, session.Save(persistAnchorIds: false));
    }

    private static System.Collections.Generic.List<XElement> UserNotes(byte[] bytes, bool footnote) =>
        PartXml(bytes, m => footnote ? m.FootnotesPart : (OpenXmlPart?)m.EndnotesPart)
            .Elements(W + (footnote ? "footnote" : "endnote"))
            .Where(n => n.Attribute(W + "type") is null)
            .ToList();

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS366_NoteAuthoredUnderRecording_ResolvesBothWays(bool footnote)
    {
        var (baseline, redline) = AuthorNoteUnderRecording(footnote);
        var referenceName = W + (footnote ? "footnoteReference" : "endnoteReference");

        using (var review = new DocxSession(redline, new DocxSessionSettings { PersistAnchorIds = false }))
        {
            // The citation is a revision at all, which is what was missing.
            Assert.NotEmpty(review.ListRevisions());
            Assert.True(review.RejectAllRevisions().Success);
            var rejected = review.Save(persistAnchorIds: false);

            // The definition went with the citation: only Word's two reserved notes remain…
            Assert.Empty(UserNotes(rejected, footnote));
            Assert.Empty(BodyXml(rejected).Descendants(referenceName));

            // …and the document as a whole is the baseline again, which is the property that matters.
            Assert.Empty(DocxDiff.GetRevisions(
                new WmlDocument("baseline.docx", baseline), new WmlDocument("rejected.docx", rejected)));
        }

        using (var review = new DocxSession(redline, new DocxSessionSettings { PersistAnchorIds = false }))
        {
            Assert.True(review.AcceptAllRevisions().Success);
            var accepted = review.Save(persistAnchorIds: false);

            var note = Assert.Single(UserNotes(accepted, footnote));
            Assert.Contains("Negotiated on March 11.", note.Descendants(W + "t").Select(t => (string)t));
            Assert.Single(BodyXml(accepted).Descendants(referenceName));
            Assert.Empty(BodyXml(accepted).Descendants(W + "ins"));
        }
    }

    /// <summary>The citation is a note reference nested inside <c>w:ins</c> — a shape this op had
    /// never produced before — so pin that the diff engine still reports it.</summary>
    [Fact]
    public void DS367_RecordedNoteInsertion_IsReportedByTheDiffEngine()
    {
        var (baseline, redline) = AuthorNoteUnderRecording(footnote: true);

        var revisions = DocxDiff.GetRevisions(
            new WmlDocument("baseline.docx", baseline), new WmlDocument("redline.docx", redline));

        Assert.Contains(revisions, r => r.Text.Contains("Negotiated on March 11."));
    }

    /// <summary>
    /// What the reversibility proof reports, which is how #614 was found. The reject path keeps
    /// divergences a generated redline legitimately explains — the run the citation split, the
    /// <c>w:trackRevisions</c>/<c>w:footnotePr</c> declarations, the note styles — but the note
    /// STORY must not be among them any more. Before the fix the residue included
    /// <c>/word/footnotes.xml</c> carrying the whole note body.
    /// </summary>
    [Fact]
    public void DS368_RecordedNoteInsertion_LeavesNoNoteResidueOnTheRejectPath()
    {
        var (baseline, redline) = AuthorNoteUnderRecording(footnote: true);
        byte[] intendedFinal;
        using (var accepting = new DocxSession(redline, new DocxSessionSettings { PersistAnchorIds = false }))
        {
            Assert.True(accepting.AcceptAllRevisions().Success);
            intendedFinal = accepting.Save(persistAnchorIds: false);
        }

        var run = Docxodus.Verification.RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);
        var reject = run.Proof.RejectToBaseline;

        Assert.True(reject?.Completed, run.Proof.ToJson());
        // Non-vacuous: the redline-authoring residue IS still reported, so the filter below is
        // reading a populated list rather than an empty one.
        Assert.NotEmpty(reject!.Divergences);
        Assert.DoesNotContain(reject.Divergences,
            divergence => divergence.PartUri.Contains("notes.xml", StringComparison.Ordinal));
        Assert.Contains(reject.Divergences,
            divergence => divergence.PartUri == "/word/document.xml");
    }

    /// <summary>
    /// The definition's content records too (issue #636, revisiting #614's unmarked-definition
    /// choice): every run sits in <c>w:ins</c> and the paragraph mark is insertion-marked, which
    /// is what Word writes and what the diff engine's own redline carries for a wholly-inserted
    /// note. That marking is what lets the STATELESS resolutions read the redline: reject empties
    /// the definition, so the guarded note-lifecycle prune takes it (before this, the stateless
    /// reject needed an unguarded prune that also ate a baseline's own kept husk — the mirror of
    /// the #631 accept bug); accept unwraps everything and keeps the note.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS369_RecordedNoteDefinition_IsInsertionMarked_AndStatelessResolutionsReadIt(
        bool footnote)
    {
        var (baseline, redline) = AuthorNoteUnderRecording(footnote);

        var note = Assert.Single(UserNotes(redline, footnote));
        Assert.All(note.Descendants(W + "r").Where(r => r.Parent!.Name != W + "rPr"),
            run => Assert.Contains(run.Ancestors(), a => a.Name == W + "ins"));
        Assert.All(note.Elements(W + "p"),
            p => Assert.NotNull(p.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "ins")));

        var rejected = Docxodus.Internal.DocxDiffOps.RejectRevisions(redline);
        Assert.Empty(UserNotes(rejected, footnote));
        Assert.Empty(DocxDiff.GetRevisions(
            new WmlDocument("baseline.docx", baseline), new WmlDocument("rejected.docx", rejected)));

        var accepted = Docxodus.Internal.DocxDiffOps.AcceptRevisions(redline);
        var acceptedNote = Assert.Single(UserNotes(accepted, footnote));
        Assert.Contains("Negotiated on March 11.",
            acceptedNote.Descendants(W + "t").Select(t => (string)t));
        Assert.Empty(acceptedNote.Descendants(W + "ins"));
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
