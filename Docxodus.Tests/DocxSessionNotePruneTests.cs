#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Structural deletes must not orphan footnote/endnote definitions (issue #591). Deleting the
/// body content that carried a note's LAST reference removes the definition from the notes
/// part via the same Word-faithful pruner revision resolution has used since issue #516 —
/// scoped strictly to ids the op itself unreferenced, so a pre-existing dangling note is
/// untouched and the part (with its Word-reserved separator notes) always survives.
/// Test IDs use the DS64x range.
/// </summary>
public class DocxSessionNotePruneTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement PartXml(byte[] docxBytes, System.Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var part = pick(doc.MainDocumentPart!);
        Assert.NotNull(part);
        return part!.GetXDocument().Root!;
    }

    private static int[] UserNoteIds(XElement notesRoot, XName noteName) =>
        notesRoot.Elements(noteName)
            .Where(n => n.Attribute(W + "type") is null)
            .Select(n => (int)n.Attribute(W + "id")!)
            .OrderBy(id => id).ToArray();

    private static string AnchorByPreview(DocxSession session, string contains) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h"
                && t.TextPreview.Contains(contains)).Anchor.Id;

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(wDoc)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
        Assert.True(errors.Count == 0, "OOXML schema errors:\n" + string.Join("\n", errors));
    }

    /// <summary>
    /// Five paragraphs: one cites footnote 2 alone, two share footnote 1, one cites endnote 1,
    /// one is plain. Footnote 7 exists but is referenced by nothing (a pre-existing dangler).
    /// </summary>
    private static byte[] BuildMultiNoteDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Solo cite."),
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 2 })),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Shared cite one."),
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 })),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Shared cite two."),
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 })),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Endnote cite."),
                            new DocumentFormat.OpenXml.Wordprocessing.EndnoteReference { Id = 1 })),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Plain paragraph.")))));

            var fnPart = main.AddNewPart<FootnotesPart>();
            var fnXml = """
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="1"><w:p><w:r><w:t>Shared note.</w:t></w:r></w:p></w:footnote>
                  <w:footnote w:id="2"><w:p><w:r><w:t>Solo note.</w:t></w:r></w:p></w:footnote>
                  <w:footnote w:id="7"><w:p><w:r><w:t>Dangling note.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """;
            using (var s = fnPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s)) w.Write(fnXml);

            var enPart = main.AddNewPart<EndnotesPart>();
            var enXml = """
                <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
                  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
                  <w:endnote w:id="1"><w:p><w:r><w:t>The endnote.</w:t></w:r></w:p></w:endnote>
                </w:endnotes>
                """;
            using (var s = enPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s)) w.Write(enXml);
        }
        return ms.ToArray();
    }

    /// <summary>A 2×2 table whose first cell cites footnote 2, plus an intro paragraph.</summary>
    private static byte[] BuildTableNoteDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Intro paragraph."))),
                    new DocumentFormat.OpenXml.Wordprocessing.Table(
                        new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.TableWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto, Width = "0" }),
                        new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                            new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2400" },
                            new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2400" }),
                        new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                            new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                                        new DocumentFormat.OpenXml.Wordprocessing.Text("Cell with note."),
                                        new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 2 }))),
                            new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                                        new DocumentFormat.OpenXml.Wordprocessing.Text("Plain cell."))))),
                        new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                            new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                                        new DocumentFormat.OpenXml.Wordprocessing.Text("Second row a.")))),
                            new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                                        new DocumentFormat.OpenXml.Wordprocessing.Text("Second row b."))))))));

            var fnPart = main.AddNewPart<FootnotesPart>();
            var fnXml = """
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="2"><w:p><w:r><w:t>Table note.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """;
            using (var s = fnPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s)) w.Write(fnXml);
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS640_DeleteBlock_PrunesTheOrphanedFootnoteDefinition()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        var result = session.DeleteBlock(AnchorByPreview(session, "Solo cite."));
        Assert.True(result.Success, result.Error?.Message);

        var saved = session.Save();
        var footnotes = PartXml(saved, m => m.FootnotesPart);
        // Footnote 2 lost its only reference and is pruned; the shared note and the
        // pre-existing dangler survive, as do the two Word-reserved separator notes.
        Assert.Equal(new[] { 1, 7 }, UserNoteIds(footnotes, W + "footnote"));
        Assert.Contains(footnotes.Elements(W + "footnote"),
            n => (string?)n.Attribute(W + "type") == "separator");
        Assert.Contains(footnotes.Elements(W + "footnote"),
            n => (string?)n.Attribute(W + "type") == "continuationSeparator");

        // The pruned definition is reported as removed alongside the paragraph.
        Assert.Contains(result.Removed, a => a.Kind == "fn");
        AssertSchemaValid(saved);
    }

    [Fact]
    public void DS641_DeleteBlock_KeepsADefinitionStillReferencedElsewhere()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        var result = session.DeleteBlock(AnchorByPreview(session, "Shared cite one."));
        Assert.True(result.Success, result.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Equal(new[] { 1, 2, 7 }, UserNoteIds(footnotes, W + "footnote"));
    }

    [Fact]
    public void DS642_DeleteBlock_PrunesWhenTheLastSharedReferenceGoes()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        Assert.True(session.DeleteBlock(AnchorByPreview(session, "Shared cite one.")).Success);
        var second = session.DeleteBlock(AnchorByPreview(session, "Shared cite two."));
        Assert.True(second.Success, second.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Equal(new[] { 2, 7 }, UserNoteIds(footnotes, W + "footnote"));
    }

    [Fact]
    public void DS643_DeleteBlock_LeavesAPreexistingDanglingNoteAlone()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        var result = session.DeleteBlock(AnchorByPreview(session, "Plain paragraph."));
        Assert.True(result.Success, result.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Equal(new[] { 1, 2, 7 }, UserNoteIds(footnotes, W + "footnote"));
    }

    [Fact]
    public void DS644_DeleteBlock_PrunesTheOrphanedEndnoteDefinition()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        var result = session.DeleteBlock(AnchorByPreview(session, "Endnote cite."));
        Assert.True(result.Success, result.Error?.Message);

        var endnotes = PartXml(session.Save(), m => m.EndnotesPart);
        Assert.Empty(UserNoteIds(endnotes, W + "endnote"));
        Assert.Contains(result.Removed, a => a.Kind == "en");
    }

    [Fact]
    public void DS645_TrackedDelete_KeepsTheDefinitionUntilTheRevisionIsAccepted()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var tracked = session.DeleteBlock(AnchorByPreview(session, "Solo cite."));
        Assert.True(tracked.Success, tracked.Error?.Message);

        // The reference is only marked deleted, so the definition must survive.
        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Equal(new[] { 1, 2, 7 }, UserNoteIds(footnotes, W + "footnote"));

        var accept = session.AcceptAllRevisions();
        Assert.True(accept.Success, accept.Error?.Message);
        footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Equal(new[] { 1, 7 }, UserNoteIds(footnotes, W + "footnote"));
    }

    [Fact]
    public void DS646_DeleteRange_PrunesEveryNoteTheRangeUnreferenced()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        var result = session.DeleteRange(
            AnchorByPreview(session, "Solo cite."),
            AnchorByPreview(session, "Endnote cite."));
        Assert.True(result.Success, result.Error?.Message);

        var saved = session.Save();
        var footnotes = PartXml(saved, m => m.FootnotesPart);
        Assert.Equal(new[] { 7 }, UserNoteIds(footnotes, W + "footnote"));
        // The endnote's citing paragraph was the exclusive end — untouched.
        var endnotes = PartXml(saved, m => m.EndnotesPart);
        Assert.Equal(new[] { 1 }, UserNoteIds(endnotes, W + "endnote"));
    }

    [Fact]
    public void DS647_DeleteTableRow_PrunesTheCellsOrphanedNote()
    {
        using var session = new DocxSession(BuildTableNoteDoc());
        var cellAnchor = AnchorByPreview(session, "Cell with note.");
        var result = session.DeleteTableRow(cellAnchor);
        Assert.True(result.Success, result.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Empty(UserNoteIds(footnotes, W + "footnote"));
    }

    [Fact]
    public void DS648_DeleteTableColumn_PrunesTheCellsOrphanedNote()
    {
        using var session = new DocxSession(BuildTableNoteDoc());
        var cellAnchor = AnchorByPreview(session, "Cell with note.");
        var result = session.DeleteTableColumn(cellAnchor);
        Assert.True(result.Success, result.Error?.Message);

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Empty(UserNoteIds(footnotes, W + "footnote"));
    }

    [Fact]
    public void DS649_Undo_RestoresAPrunedDefinition()
    {
        using var session = new DocxSession(BuildMultiNoteDoc());
        Assert.True(session.DeleteBlock(AnchorByPreview(session, "Solo cite.")).Success);
        Assert.True(session.Undo());

        var footnotes = PartXml(session.Save(), m => m.FootnotesPart);
        Assert.Equal(new[] { 1, 2, 7 }, UserNoteIds(footnotes, W + "footnote"));
        Assert.Contains("Solo note.",
            footnotes.Descendants(W + "t").Select(t => (string)t));
    }
}
