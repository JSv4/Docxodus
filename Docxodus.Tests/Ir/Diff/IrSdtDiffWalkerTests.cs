#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using WordType = DocumentFormat.OpenXml.WordprocessingDocumentType;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Regression coverage for the bookkeeping walkers that must see through a block-level SDT while the
/// aligner still treats the SDT envelope itself as atomic.
/// </summary>
public sealed class IrSdtDiffWalkerTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };
    private static readonly IrDiffSettings Default = new();

    [Fact]
    public void Block_sdt_note_ref_participates_in_note_correspondence()
    {
        var left = Read(FootnoteDoc(
            Sdt(Paragraph(FootnoteReference("1"))),
            ("1", Paragraph(Text("shared old")))));
        var right = Read(FootnoteDoc(
            Sdt(Paragraph(FootnoteReference("2"))),
            ("2", Paragraph(Text("shared new")))));

        var script = IrEditScriptBuilder.Build(left, right, Default);
        var note = Assert.Single(script.NoteOps!.Where(n => n.Kind == IrNoteKind.Footnote));

        Assert.Equal("1", note.LeftNoteId);
        Assert.Equal("2", note.NoteId);
        Assert.Contains(note.Ops, op => op.Kind == IrEditOpKind.ModifyBlock);
    }

    [Fact]
    public void Note_token_bag_descends_block_sdt()
    {
        // The content-based residue matcher must select 1 → 4. The reverse-order banana match then
        // cannot cross it, so it remains a delete+insert. Without SDT descent every note bag is empty,
        // ties are positional, and the incorrect 1 → 3 pairing wins instead.
        var left = Read(FootnoteDoc(
            Paragraph(FootnoteReference("1")) + Paragraph(FootnoteReference("2")),
            ("1", Sdt(Paragraph(Text("orange apple pear")))),
            ("2", Sdt(Paragraph(Text("banana"))))));
        var right = Read(FootnoteDoc(
            Paragraph(FootnoteReference("3")) + Paragraph(FootnoteReference("4")),
            ("3", Sdt(Paragraph(Text("banana changed")))),
            ("4", Sdt(Paragraph(Text("orange apple pear changed"))))));

        var script = IrEditScriptBuilder.Build(left, right, Default);
        var notes = script.NoteOps!.Where(n => n.Kind == IrNoteKind.Footnote).ToList();

        Assert.Contains(notes, n => n.LeftNoteId == "1" && n.NoteId == "4");
        Assert.DoesNotContain(notes, n => n.LeftNoteId == "1" && n.NoteId == "3");
        Assert.Contains(notes, n => n.LeftNoteId is null && n.NoteId == "3");
        Assert.Contains(notes, n => n.LeftNoteId == "2" && n.NoteId == "2");
    }

    [Fact]
    public void Table_similarity_descends_block_sdt_in_cell()
    {
        var left = Read(IrTestDocuments.FromBodyXml(
            OneCellTable(Sdt(Paragraph(Text("alpha beta"))))));
        var right = Read(IrTestDocuments.FromBodyXml(
            OneCellTable(Sdt(Paragraph(Text("alpha gamma"))))));

        var leftTable = Assert.IsType<IrTable>(left.Body.Blocks.Single());
        var rightTable = Assert.IsType<IrTable>(right.Body.Blocks.Single());
        double score = new IrBlockSimilarity(Default).Score(leftTable, rightTable);

        Assert.True(score is > 0 and < 1,
            $"Expected overlapping SDT cell text to yield a partial table similarity, got {score}.");
    }

    private static IrDocument Read(WmlDocument document) => IrReader.Read(document, NoSources);

    private static WmlDocument FootnoteDoc(string bodyInnerXml, params (string Id, string InnerXml)[] notes)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var footnotes = new StringBuilder($"<w:footnotes xmlns:w=\"{W}\">")
                .Append("<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>")
                .Append("<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>");
            foreach (var (id, innerXml) in notes)
                footnotes.Append($"<w:footnote w:id=\"{id}\">{innerXml}</w:footnote>");
            footnotes.Append("</w:footnotes>");

            WritePart(main.AddNewPart<FootnotesPart>(), footnotes.ToString());
            WritePart(main, $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>");
        }
        return new WmlDocument("sdt-note-walker.docx", ms.ToArray());
    }

    private static void WritePart(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(xml);
    }

    private static string Text(string value) =>
        $"<w:r><w:t xml:space=\"preserve\">{value}</w:t></w:r>";

    private static string Paragraph(string innerXml) => $"<w:p>{innerXml}</w:p>";

    private static string FootnoteReference(string id) => $"<w:r><w:footnoteReference w:id=\"{id}\"/></w:r>";

    private static string Sdt(string innerXml) =>
        $"<w:sdt><w:sdtPr/><w:sdtContent>{innerXml}</w:sdtContent></w:sdt>";

    private static string OneCellTable(string cellInnerXml) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
        $"<w:tr><w:tc><w:tcPr/>{cellInnerXml}</w:tc></w:tr></w:tbl>";
}
