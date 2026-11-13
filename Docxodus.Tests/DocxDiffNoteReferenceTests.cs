using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffNoteReferenceTests
{
    private const string Reserved =
        "<w:footnote w:type='separator' w:id='-1'><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type='continuationNotice' w:id='3'><w:p/></w:footnote>";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData(false, 0, null)]
    [InlineData(true, 0, null)]
    [InlineData(false, 7, "normal")]
    [InlineData(true, 7, "normal")]
    public void Body_edit_preserves_normal_notes(bool endnotes, int id, string? type)
    {
        var notes = Reserved + Note(id, "First note", type) + Note(11, "Second note");
        var left = Build(endnotes, Paragraph("Alpha", id, 11, id), notes);
        var right = Build(endnotes, Paragraph("Beta", id, 11, id), notes);

        AssertRoundTrip(left, right, DocxDiff.Compare(left, right), endnotes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Emptying_a_note_preserves_its_citation(bool endnotes)
    {
        var body = Paragraph("Alpha", 7, 11);
        var left = Build(endnotes, body, Reserved + Note(7, "Old note") + Note(11, "Other note"));
        var right = Build(endnotes, body, Reserved + Note(7, "") + Note(11, "Other note"));

        AssertRoundTrip(left, right, DocxDiff.Compare(left, right), endnotes);
    }

    [Theory]
    [InlineData(false, false, 0, null)]
    [InlineData(true, false, 0, null)]
    [InlineData(false, true, 0, null)]
    [InlineData(true, true, 0, null)]
    [InlineData(false, false, 7, "normal")]
    [InlineData(true, false, 7, "normal")]
    [InlineData(false, true, 7, "normal")]
    [InlineData(true, true, 7, "normal")]
    public void Inserting_a_note_into_a_missing_part_round_trips(
        bool endnotes, bool consolidate, int id, string? type)
    {
        var left = Build(endnotes, Paragraph("Alpha"), null);
        var right = Build(endnotes, Paragraph("Alpha", id), Reserved + Note(id, "Inserted note", type));
        var rendered = consolidate
            ? Consolidate(left, right)
            : DocxDiff.Compare(left, right);

        AssertRoundTrip(left, right, rendered, endnotes);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void A_content_note_sharing_a_separator_id_is_not_the_separator(bool endnotes, bool delete, bool consolidate)
    {
        const string separator =
            "<w:footnote w:type='continuationSeparator' w:id='0'><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";
        var empty = Build(endnotes, Paragraph("Alpha"), Reserved + separator);
        var malformed = Build(endnotes, Paragraph("Alpha", 0), Reserved + separator + Note(0, "Content note"));
        var repaired = Build(endnotes, Paragraph("Alpha", 7), Reserved + separator + Note(7, "Content note"));
        var left = delete ? malformed : empty;
        var right = delete ? empty : malformed;
        var rendered = consolidate ? Consolidate(left, right) : DocxDiff.Compare(left, right);

        AssertRoundTrip(delete ? repaired : empty, delete ? empty : repaired, rendered, endnotes);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void Missing_definitions_remove_only_the_dangling_reference(bool endnotes, bool hasPart, bool delete)
    {
        var notes = hasPart ? Reserved + Note(7, "Cited note") + Note(11, "Unreferenced note") : null;
        var validIds = hasPart ? new[] { 7 } : Array.Empty<int>();
        var left = Build(endnotes, Paragraph("Alpha", validIds) + Paragraph("Closing", 99), notes);
        var right = Build(endnotes, Paragraph("Beta", validIds) + Paragraph("Closing", delete ? Array.Empty<int>() : new[] { 99 }), notes);
        var repairedLeft = Build(endnotes, Paragraph("Alpha", validIds) + Paragraph("Closing"), notes);
        var repairedRight = Build(endnotes, Paragraph("Beta", validIds) + Paragraph("Closing"), notes);

        AssertRoundTrip(repairedLeft, repairedRight, DocxDiff.Compare(left, right), endnotes);
    }

    // Issue #1701. Both documents number their notes from 1, and the body diff deletes the original's
    // paragraphs and inserts the revised ones, so the redline writes all four citations.
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Notes_cited_only_from_deleted_and_inserted_text_are_not_paired(bool endnotes, bool consolidate)
    {
        var left = Build(endnotes, Orchard + Irrigation, Reserved + Note(1, "County records.") + Note(2, "Water board survey."));
        var right = Build(endnotes, Ferry + Tickets, Reserved + Note(1, "Port schedule.") + Note(2, "October fare table."));
        var rendered = consolidate ? Consolidate(left, right) : DocxDiff.Compare(left, right);

        AssertCitations(rendered, endnotes, citations: 4, modifiedNotes: 0);
        AssertRoundTrip(left, right, rendered, endnotes);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void A_rewritten_note_stays_one_note_when_its_citation_is_kept(bool endnotes, bool replaceOtherParagraph)
    {
        // The first paragraph changes one word, so its citation is kept and its rewritten note is one note
        // with tracked edits. A replaced second paragraph brings two notes of its own.
        var left = Build(endnotes, Orchard + Irrigation, Reserved + Note(1, "County records.") + Note(2, "Water board survey."));
        var right = Build(endnotes, Orchard.Replace("spring", "winter") + (replaceOtherParagraph ? Tickets : Irrigation),
            Reserved + Note(1, "Port schedule.") + Note(2, replaceOtherParagraph ? "October fare table." : "Water board survey."));
        var rendered = DocxDiff.Compare(left, right);

        AssertCitations(rendered, endnotes, citations: replaceOtherParagraph ? 3 : 2, modifiedNotes: 1);
        AssertRoundTrip(left, right, rendered, endnotes);
    }

    [Theory]
    [InlineData("<w:p>", "<w:p><w:pPr><w:jc w:val='center'/></w:pPr>", false)]                 // format-only paragraph
    [InlineData("spring.</w:t></w:r><w:r>", "winter.</w:t></w:r><w:r><w:rPr><w:b/></w:rPr>", false)] // edited paragraph
    [InlineData(" Irrigation", ParagraphBreak + "Irrigation", false)]                          // split in two
    [InlineData(" Irrigation", ParagraphBreak + "Irrigation", true)]                           // two merged into one
    public void A_citation_kept_through_a_format_change_edit_split_or_merge_keeps_its_note_paired(
        string find, string replacement, bool reverse)
    {
        var changed = TwoCitations.Replace(find, replacement);
        var left = Build(false, reverse ? changed : TwoCitations, Reserved + Note(1, "County records.") + Note(2, "Water board survey."));
        var right = Build(false, reverse ? TwoCitations : changed, Reserved + Note(1, "County archive.") + Note(2, "Water board map."));
        var rendered = DocxDiff.Compare(left, right);

        AssertCitations(rendered, false, citations: 2, modifiedNotes: 2);
        AssertRoundTrip(left, right, rendered, false);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Consolidating_a_kept_note_whose_id_a_deleted_base_note_shares_keeps_both_citations(bool endnotes)
    {
        // The reviewer edits the first paragraph and its note, which their document numbers 2, and deletes
        // the second paragraph with the base's note 2.
        var left = Build(endnotes, Orchard + Irrigation, Reserved + Note(1, "County records.") + Note(2, "Water board survey."));
        var right = Build(endnotes, Orchard.Replace("spring", "winter").Replace("w:id='1'", "w:id='2'"),
            Reserved + Note(2, "County archive."));
        var rendered = Consolidate(left, right);

        AssertCitations(rendered, endnotes, citations: 2, modifiedNotes: 1);
        AssertRoundTrip(left, right, rendered, endnotes);
    }

    [Fact]
    public void A_paragraph_rewritten_whole_keeps_none_of_its_citations()
    {
        // An index-entry field has no visible text, so the comparison replaces its edited paragraph whole:
        // the original's citation is deleted with it and the revised one inserted.
        const string indexEntry = "<w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText> XE \"orchard\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType='end'/></w:r>";
        var body = Orchard.Replace("<w:p>", "<w:p>" + indexEntry);
        var left = Build(false, body, Reserved + Note(1, "County records."));
        var right = Build(false, body.Replace("spring", "winter"), Reserved + Note(1, "County archive."));
        var rendered = DocxDiff.Compare(left, right);

        AssertCitations(rendered, false, citations: 2, modifiedNotes: 0);
        AssertRoundTrip(left, right, rendered, false);
    }

    /// <summary>One paragraph, two sentences, each with its citation.</summary>
    private const string TwoCitations =
        "<w:p><w:r><w:t>The orchard was replanted after the frost of the previous spring.</w:t></w:r>" +
        "<w:r><w:footnoteReference w:id='1'/></w:r><w:r><w:t xml:space='preserve'> Irrigation lines were " +
        "rerouted along the eastern slope.</w:t></w:r><w:r><w:footnoteReference w:id='2'/></w:r></w:p>";
    private const string ParagraphBreak = "</w:t></w:r></w:p><w:p><w:r><w:t>";
    private static readonly string Orchard = Paragraph("The orchard was replanted after the frost of the previous spring.", 1);
    private static readonly string Irrigation = Paragraph("Irrigation lines were rerouted along the eastern slope.", 2);
    private static readonly string Ferry = Paragraph("Ferry service resumes once the harbour dredging is finished.", 1);
    private static readonly string Tickets = Paragraph("Ticket prices stay unchanged through the winter season.", 2);

    /// <summary>Every citation in the redline names a note of its own, and exactly
    /// <paramref name="modifiedNotes"/> notes hold both inserted and deleted text: the rest belong to one
    /// document. Word refuses a redline in which a deleted and an inserted citation share a note.</summary>
    private static void AssertCitations(WmlDocument redline, bool endnotes, int citations, int modifiedNotes)
    {
        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var ids = main.GetXDocument().Descendants(W + (endnotes ? "endnoteReference" : "footnoteReference"))
            .Select(r => (string)r.Attribute(W + "id")!).ToList();
        Assert.Equal(citations, ids.Count);
        Assert.Equal(ids.Count, ids.Distinct().Count());

        var part = endnotes ? (OpenXmlPart)main.EndnotesPart! : main.FootnotesPart!;
        Assert.Equal(modifiedNotes, part.GetXDocument().Root!.Elements()
            .Count(n => n.Descendants(W + "ins").Any() && n.Descendants(W + "del").Any()));
    }

    private static void AssertRoundTrip(WmlDocument left, WmlDocument right, WmlDocument rendered, bool endnotes)
    {
        Assert.Equal(Read(right, endnotes).Reserved, Read(rendered, endnotes).Reserved);
        AssertNotesMatch(right, RevisionProcessor.AcceptRevisions(rendered), endnotes);
        AssertNotesMatch(left, RevisionProcessor.RejectRevisions(rendered), endnotes);
    }

    private static void AssertNotesMatch(WmlDocument expected, WmlDocument actual, bool endnotes)
    {
        var e = Read(expected, endnotes);
        var a = Read(actual, endnotes);
        Assert.Equal(e.Body, a.Body);
        Assert.Equal(e.Citations, a.Citations);
        Assert.Equal(e.Notes, a.Notes);
    }

    private static WmlDocument Consolidate(WmlDocument left, WmlDocument right) =>
        DocxDiff.Consolidate(left, new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = right },
            new DocxDiffReviewer { Author = "Bob", Document = left },
        });

    private static (string Body, string[] Citations, string[] Notes, string[] Reserved) Read(
        WmlDocument document, bool endnotes)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        PackageValidation.AssertValid(doc, FileFormatVersions.Office2019);
        var main = doc.MainDocumentPart!;
        var body = main.GetXDocument().Root!.Element(W + "body")!;
        var part = endnotes ? (OpenXmlPart?)main.EndnotesPart : main.FootnotesPart;
        var notes = part?.GetXDocument().Root!.Elements().ToList() ?? new List<XElement>();
        var content = notes.Where(n => (string?)n.Attribute(W + "type") is null or "normal").ToList();
        var citations = body.Descendants(W + (endnotes ? "endnoteReference" : "footnoteReference"))
            .Select(r => Assert.Single(content, n => (string?)n.Attribute(W + "id") == (string?)r.Attribute(W + "id")).Value)
            .ToArray();
        return (string.Concat(body.Descendants(W + "t").Select(t => t.Value)), citations,
            content.Select(n => n.Value).OrderBy(t => t).ToArray(),
            notes.Except(content).Select(n => $"{n.Attribute(W + "id")}:{n.Attribute(W + "type")}:{n.Value}").ToArray());
    }

    private static WmlDocument Build(bool endnotes, string body, string? notes)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            if (endnotes)
            {
                body = body.Replace("footnote", "endnote");
                notes = notes?.Replace("footnote", "endnote");
            }
            if (notes != null)
            {
                OpenXmlPart part = endnotes ? main.AddNewPart<EndnotesPart>() : main.AddNewPart<FootnotesPart>();
                var name = endnotes ? "endnotes" : "footnotes";
                Write(part, $"<w:{name} xmlns:w='{W}'>{notes}</w:{name}>");
            }
            Write(main, $"<w:document xmlns:w='{W}'><w:body>{body}</w:body></w:document>");
        }
        return new WmlDocument("notes.docx", stream.ToArray());
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
        writer.Write(xml);
    }

    private static string Paragraph(string text, params int[] references) =>
        $"<w:p><w:r><w:t>{text}</w:t></w:r>" +
        string.Concat(references.Select(id => $"<w:r><w:footnoteReference w:id='{id}'/></w:r>")) + "</w:p>";

    private static string Note(int id, string text, string? type = null) =>
        $"<w:footnote w:id='{id}'{(type == null ? "" : $" w:type='{type}'")}><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:footnote>";
}
