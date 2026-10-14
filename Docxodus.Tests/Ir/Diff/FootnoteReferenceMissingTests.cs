#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Word offers to recover a file when a <c>w:footnoteReference</c> names an id that is not a
/// <c>w:footnote</c>. One reference is enough. Proved on docxodus redlines 1829 and 1830.
/// 1829 also has deleted text inside a text box; that defect is covered by
/// <see cref="TextBoxDeletedTextTests"/>. A content footnote at id 0, beside the reserved
/// continuation separator that also uses id 0, is a real note. A reference that names no footnote
/// is removed. References that name footnotes 1 and 2 stay on those footnotes.
/// </summary>
public class FootnoteReferenceMissingTests
{
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Boilerplate =
        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static readonly string ValidNotes =
        Boilerplate + Note("0", "Simoes") + Note("1", "Planes") + Note("2", "Baseado");

    private static readonly string CitingBody =
        "<w:p><w:r><w:t xml:space=\"preserve\">Os principios.</w:t></w:r>" +
        Ref("0") + Ref("1") + "</w:p>" +
        "<w:p><w:r><w:t xml:space=\"preserve\">A definicao.</w:t></w:r>" +
        Ref("2") + "</w:p>";

    /// <summary>
    /// 1829. The left document already has the separator and the continuation separator.
    /// The right document adds three content footnotes, the first at id 0, and cites them
    /// in order: Simoes then Planes, then Baseado.
    /// </summary>
    [Fact]
    public void Insert_keeps_every_content_footnote_cited_when_the_first_id_is_zero()
    {
        var left = Build(
            "<w:p><w:r><w:t xml:space=\"preserve\">intro</w:t></w:r></w:p>",
            Boilerplate);
        var right = Build(CitingBody, ValidNotes);

        var rendered = Render(left, right);
        AssertCitations(rendered, "Simoes", "Planes", "Baseado");
        AssertReservedBoilerplate(rendered);

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        AssertCitations(accepted, "Simoes", "Planes", "Baseado");
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Inspect(rejected, (body, notes) =>
        {
            AssertEveryReferenceResolves(body, notes);
            Assert.Empty(body.Descendants(W.footnoteReference));
            Assert.DoesNotContain(notes, note => note.Value.Contains("Simoes"));
            Assert.DoesNotContain(notes, note => note.Value.Contains("Planes"));
            Assert.DoesNotContain(notes, note => note.Value.Contains("Baseado"));
        });
    }

    /// <summary>
    /// 1830. Content footnotes are 0, 1, and 2. The body cites 1, 2, and 3.
    /// Reference 3 names nothing. Removing that one reference leaves 1 on Planes and 2 on Baseado.
    /// Footnote 0 stays, unreferenced.
    /// </summary>
    [Fact]
    public void Render_removes_a_footnote_reference_that_names_no_footnote()
    {
        var notes = Note("0", "Simoes") + Note("1", "Planes") + Note("2", "Baseado");
        var body =
            "<w:p><w:r><w:t xml:space=\"preserve\">Os principios.</w:t></w:r>" +
            Ref("1") + Ref("2") + "</w:p>" +
            "<w:p><w:r><w:t xml:space=\"preserve\">A definicao.</w:t></w:r>" +
            Ref("3") + "</w:p>";
        var document = Build(body, notes);

        var rendered = Render(document, document);
        Inspect(rendered, (renderedBody, renderedNotes) =>
        {
            AssertEveryReferenceResolves(renderedBody, renderedNotes);
            var first = renderedBody.Elements(W.p).Single(paragraph => paragraph.Value.Contains("Os principios"));
            var second = renderedBody.Elements(W.p).Single(paragraph => paragraph.Value.Contains("A definicao"));
            Assert.Equal(new[] { "Planes", "Baseado" }, Cited(first, renderedNotes));
            Assert.Empty(second.Descendants(W.footnoteReference));
            var zero = renderedNotes.Single(note => (string?)note.Attribute(W.id) == "0");
            Assert.Contains("Simoes", zero.Value);
            Assert.Equal(1, renderedNotes.Count(note => note.Value.Contains("Simoes")));
        });
    }

    /// <summary>
    /// 1830's dangling reference is an equal reference: both sides are the same document.
    /// A deletion is the other side. The left document cites id 3 and the right document does not.
    /// The deleted reference names no footnote, so the render removes it instead of keeping it
    /// inside <c>w:del</c> or pointing it at another note.
    /// </summary>
    [Fact]
    public void Render_removes_a_deleted_footnote_reference_that_names_no_footnote()
    {
        var notes = Note("0", "Simoes") + Note("1", "Planes") + Note("2", "Baseado");
        var cited =
            "<w:p><w:r><w:t xml:space=\"preserve\">Os principios.</w:t></w:r>" +
            Ref("1") + Ref("2") + "</w:p>";
        var left = Build(
            cited + "<w:p><w:r><w:t xml:space=\"preserve\">A definicao.</w:t></w:r>" + Ref("3") + "</w:p>",
            notes);
        var right = Build(
            cited + "<w:p><w:r><w:t xml:space=\"preserve\">A definicao.</w:t></w:r></w:p>",
            notes);

        var rendered = Render(left, right);
        Inspect(rendered, (body, renderedNotes) =>
        {
            AssertEveryReferenceResolves(body, renderedNotes);
            var opening = body.Elements(W.p).Single(paragraph => paragraph.Value.Contains("Os principios"));
            var closing = body.Elements(W.p).Single(paragraph => paragraph.Value.Contains("A definicao"));
            Assert.Equal(new[] { "Planes", "Baseado" }, Cited(opening, renderedNotes));
            Assert.Empty(closing.Descendants(W.footnoteReference));
            Assert.Equal(1, renderedNotes.Count(note => note.Value.Contains("Simoes")));
        });
    }

    /// <summary>
    /// Neither document has a footnotes part. A body <c>w:footnoteReference</c> then names nothing.
    /// <c>RenumberNoteIds</c> used to return before the missing-definition check, so rendering the
    /// document against itself left the reference in place and Word still offered to recover.
    /// </summary>
    [Fact]
    public void Render_removes_a_footnote_reference_when_neither_document_has_a_footnotes_part()
    {
        var body =
            "<w:p><w:r><w:t xml:space=\"preserve\">Os principios.</w:t></w:r>" +
            Ref("1") + "</w:p>";
        var document = Build(body, footnotesInnerXml: null);

        var rendered = Render(document, document);
        Inspect(rendered, (renderedBody, renderedNotes) =>
        {
            Assert.Empty(renderedNotes);
            Assert.Empty(renderedBody.Descendants(W.footnoteReference));
            Assert.Contains("Os principios", renderedBody.Value);
        });
    }

    /// <summary>
    /// The left document has no footnotes part. The right document's content footnote at id 0 sits
    /// beside continuationSeparator, which also uses id 0. Seeding the new part must copy the separator
    /// notes only, so the content note is inserted once and stays cited.
    /// </summary>
    [Fact]
    public void Insert_into_a_document_with_no_footnote_part_cites_the_id_zero_note_once()
    {
        var left = Build(
            "<w:p><w:r><w:t xml:space=\"preserve\">intro</w:t></w:r></w:p>",
            footnotesInnerXml: null);
        var right = Build(CitingBody, ValidNotes);

        var rendered = Render(left, right);
        AssertCitations(rendered, "Simoes", "Planes", "Baseado");
        AssertReservedBoilerplate(rendered);
    }

    /// <summary>
    /// Deleting the citing document must keep the id-0 content footnote paired with its reference.
    /// The deleted note text stays <c>w:delText</c>, and the deleted body paragraph stays too.
    /// </summary>
    [Fact]
    public void Delete_keeps_a_deleted_content_footnote_at_id_zero_resolved()
    {
        var left = Build(CitingBody, ValidNotes);
        var right = Build(
            "<w:p><w:r><w:t xml:space=\"preserve\">gone</w:t></w:r></w:p>",
            Boilerplate);

        var rendered = Render(left, right);
        Inspect(rendered, (body, notes) =>
        {
            AssertEveryReferenceResolves(body, notes);
            var deleted = notes.SelectMany(note => note.Descendants(W.delText)).Select(text => text.Value).ToList();
            Assert.Contains("Simoes", deleted);
            Assert.Contains("Planes", deleted);
            Assert.Contains("Baseado", deleted);
            Assert.Contains(body.Descendants(W.delText).Select(text => text.Value), text => text.Contains("Os principios"));
        });

        var rejected = RevisionProcessor.RejectRevisions(rendered);
        AssertCitations(rejected, "Simoes", "Planes", "Baseado");
    }

    private static void AssertCitations(WmlDocument rendered, string first, string second, string third)
    {
        Inspect(rendered, (body, notes) =>
        {
            AssertEveryReferenceResolves(body, notes);
            var opening = body.Elements(W.p).Single(paragraph => paragraph.Value.Contains("Os principios"));
            var closing = body.Elements(W.p).Single(paragraph => paragraph.Value.Contains("A definicao"));
            Assert.Equal(new[] { first, second }, Cited(opening, notes));
            Assert.Equal(new[] { third }, Cited(closing, notes));
            Assert.Equal(1, notes.Count(note => note.Value.Contains(first)));
        });
    }

    private static void AssertReservedBoilerplate(WmlDocument rendered)
    {
        Inspect(rendered, (_, notes) =>
        {
            Assert.Contains(notes, note =>
                (string?)note.Attribute(W.type) == "separator" && (string?)note.Attribute(W.id) == "-1");
            Assert.Contains(notes, note =>
                (string?)note.Attribute(W.type) == "continuationSeparator" && (string?)note.Attribute(W.id) == "0");
        });
    }

    private static void AssertEveryReferenceResolves(XElement body, IReadOnlyList<XElement> notes)
    {
        var noteIds = notes.Select(note => (string?)note.Attribute(W.id)).ToList();
        foreach (var reference in body.Descendants(W.footnoteReference))
        {
            var id = (string?)reference.Attribute(W.id);
            Assert.True(noteIds.Count(noteId => noteId == id) == 1,
                $"reference {id} does not name exactly one footnote; notes [{string.Join(",", noteIds)}]");
        }
    }

    private static List<string> Cited(XElement paragraph, IReadOnlyList<XElement> notes)
    {
        var cited = new List<string>();
        foreach (var reference in paragraph.Descendants(W.footnoteReference))
        {
            var id = (string?)reference.Attribute(W.id);
            var note = notes.SingleOrDefault(candidate =>
                (string?)candidate.Attribute(W.id) == id && candidate.Attribute(W.type) == null);
            cited.Add(note?.Value ?? "");
        }
        return cited;
    }

    private static void Inspect(WmlDocument rendered, System.Action<XElement, IReadOnlyList<XElement>> assert)
    {
        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var body = main.GetXDocument().Root!.Element(W.body)!;
        var notes = main.FootnotesPart?.GetXDocument().Root?.Elements(W.footnote).ToList()
            ?? new List<XElement>();
        assert(body, notes);
    }

    private static WmlDocument Render(WmlDocument left, WmlDocument right)
    {
        var settings = new IrDiffSettings();
        var irLeft = IrReader.Read(left, new IrReaderOptions { RetainSources = false, RevisionView = Docxodus.Ir.RevisionView.Accept });
        var irRight = IrReader.Read(right, new IrReaderOptions { RetainSources = false, RevisionView = Docxodus.Ir.RevisionView.Accept });
        var script = IrEditScriptBuilder.Build(irLeft, irRight, settings);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    private static WmlDocument Build(string bodyInnerXml, string? footnotesInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            if (footnotesInnerXml != null)
            {
                var footnotes = main.AddNewPart<FootnotesPart>();
                Write(footnotes, $"<w:footnotes xmlns:w=\"{Wns}\">{footnotesInnerXml}</w:footnotes>");
            }
            Write(main, $"<w:document xmlns:w=\"{Wns}\"><w:body>{bodyInnerXml}</w:body></w:document>");
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }

    private static string Note(string id, string text) =>
        $"<w:footnote w:id=\"{id}\"><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:footnote>";

    private static string Ref(string id) =>
        $"<w:r><w:footnoteReference w:id=\"{id}\"/></w:r>";
}
