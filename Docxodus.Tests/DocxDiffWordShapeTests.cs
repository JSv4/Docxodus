#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace Docxodus.Tests;

/// <summary>
/// Word-shaped arrangement of replace gaps in <see cref="DocxDiff.Compare"/> output. At a gap where
/// old blocks are deleted and new blocks inserted, Microsoft Word's own compare emits the INSERTED
/// content first and the deleted content after it; the legacy shape (deletions first) inverts the
/// entire page layout on heavily-rewritten documents. The reorder is a pure projection change: the
/// edit script, revision semantics, and the accept ≡ right / reject ≡ left contract are unchanged.
/// </summary>
public class DocxDiffWordShapeTests
{
    private static WmlDocument Doc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("test.docx", stream.ToArray());
    }

    private static WmlDocument SingleCellTableDoc(params string[] cellParagraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var cell = new WTableCell(cellParagraphs.Select(t => new Paragraph(new Run(new Text(t)))));
            var table = new WTable(
                new TableProperties(new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 })),
                new TableGrid(new GridColumn()),
                new WTableRow(cell));
            mainPart.Document = new Document(new Body(table, new Paragraph()));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("table.docx", stream.ToArray());
    }

    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    /// <summary>Per body paragraph: "ins" (has inserted runs, no deleted), "del" (deleted, no
    /// inserted), "mixed", or "plain" — with the paragraph's visible text.</summary>
    private static List<(string State, string Text)> ParagraphStates(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart!.Document.Body!;
        var result = new List<(string, string)>();
        foreach (var p in body.Elements<Paragraph>())
        {
            var hasIns = p.Descendants<InsertedRun>().Any();
            var hasDel = p.Descendants<DeletedRun>().Any();
            var state = hasIns && hasDel ? "mixed" : hasIns ? "ins" : hasDel ? "del" : "plain";
            result.Add((state, p.InnerText));
        }
        return result;
    }

    [Fact]
    public void TotalRewrite_InsertedContentRendersBeforeDeletedContent()
    {
        // Zero token overlap and multi-block residue on both sides → the aligner lowers the whole
        // gap to Deletes + Inserts. Word's arrangement: new content first; the LAST inserted
        // paragraph and the FIRST deleted paragraph share one w:p (the seam); remaining deleted
        // content after it.
        var left = Doc("Alpha ancient text.", "Bravo bygone text.");
        var right = Doc("Neon fresh words.", "Quantum modern words.", "Zephyr future words.");

        var result = DocxDiff.Compare(left, right);

        var states = ParagraphStates(result);
        var expected = new List<(string, string)>
        {
            ("ins", "Neon fresh words."),
            ("ins", "Quantum modern words."),
            ("mixed", "Zephyr future words.Alpha ancient text."),
            ("del", "Bravo bygone text."),
        };
        Assert.Equal(expected, states);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void MidDocumentReplaceGap_InsertsBeforeDeletes_BetweenUnchangedSpine()
    {
        var left = Doc("Shared opening paragraph.", "Alpha ancient text.", "Bravo bygone text.", "Shared closing paragraph.");
        var right = Doc("Shared opening paragraph.", "Neon fresh words.", "Quantum modern words.", "Shared closing paragraph.");

        var result = DocxDiff.Compare(left, right);

        var states = ParagraphStates(result);
        var expected = new List<(string, string)>
        {
            ("plain", "Shared opening paragraph."),
            ("ins", "Neon fresh words."),
            ("mixed", "Quantum modern words.Alpha ancient text."),
            ("del", "Bravo bygone text."),
            ("plain", "Shared closing paragraph."),
        };
        Assert.Equal(expected, states);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void SeamMerge_SingleDeletedParagraph_LiveMarkRoundTrips()
    {
        // m=1: the seam consumes the only deleted paragraph; its mark stays LIVE so accept ends the
        // inserted text at it (no bleed into following content) and reject restores the old paragraph.
        var left = Doc("Common head.", "Obsolete removed prose.", "Common tail.");
        var right = Doc("Common head.", "Fresh writing appears.", "Second novel sentence.", "Common tail.");

        var result = DocxDiff.Compare(left, right);

        var states = ParagraphStates(result);
        var expected = new List<(string, string)>
        {
            ("plain", "Common head."),
            ("ins", "Fresh writing appears."),
            ("mixed", "Second novel sentence.Obsolete removed prose."),
            ("plain", "Common tail."),
        };
        Assert.Equal(expected, states);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void IntraParagraphReplace_CoalescesToOneInsThenOneDel()
    {
        // Word renders a contiguous changed region inside a paragraph as ONE inserted block followed
        // by ONE deleted block ("Heading 2 Center1 Style Demo"), consuming the interior whitespace
        // into both sides — not per-word alternating del/ins pairs. The engine's edit script keeps
        // its token grain; this is a rendering-projection concern.
        var left = Doc("Heading 1 Style Demo");
        var right = Doc("Heading 2 Center Demo");

        var result = DocxDiff.Compare(left, right);

        using (var stream = new MemoryStream(result.DocumentByteArray))
        using (var wdoc = WordprocessingDocument.Open(stream, false))
        {
            var para = wdoc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
            // Walk direct children: expect plain-run(s), then ins-region, then del-region, then plain.
            var regions = new List<(string Kind, string Text)>();
            foreach (var child in para.ChildElements)
            {
                var (kind, text) = child switch
                {
                    InsertedRun ins => ("ins", ins.InnerText),
                    DeletedRun del => ("del", del.InnerText),
                    Run r => ("plain", r.InnerText),
                    _ => ("", ""),
                };
                if (kind == "")
                    continue;
                if (regions.Count > 0 && regions[^1].Kind == kind)
                    regions[^1] = (kind, regions[^1].Text + text);
                else
                    regions.Add((kind, text));
            }
            Assert.Equal(
                new List<(string, string)>
                {
                    ("plain", "Heading "),
                    ("ins", "2 Center"),
                    ("del", "1 Style"),
                    ("plain", " Demo"),
                },
                regions);
        }

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void IntraParagraphReplace_CoalescesAcrossFormatChangedWhitespace()
    {
        // When the two sides' run FORMATS differ (e.g. bold-italic → underline), the separator
        // spaces between replaced words are FormatChanged spans, not Equal — they must still act as
        // interior glue so the replacement coalesces into one ins region + one del region instead of
        // a word-by-word zip ("All Bold three italic styles text …").
        static WmlDocument FormattedDoc(bool underline, params string[] words)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                var runs = words.Select((w, i) =>
                {
                    var rPr = underline
                        ? new RunProperties(new Underline { Val = UnderlineValues.Single })
                        : new RunProperties(new Bold(), new Italic());
                    var text = i == words.Length - 1 ? w : w + " ";
                    return new Run(rPr, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                });
                mainPart.Document = new Document(new Body(new Paragraph(runs)));
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("fmt.docx", stream.ToArray());
        }

        var left = FormattedDoc(false, "Bold", "italic", "text", "creates", "emphasis");
        var right = FormattedDoc(true, "All", "three", "styles", "combined", "create");

        var result = DocxDiff.Compare(left, right);

        using var stream2 = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream2, false);
        var para = wdoc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        var regions = new List<string>();
        foreach (var child in para.ChildElements)
        {
            var kind = child switch
            {
                InsertedRun => "ins",
                DeletedRun => "del",
                Run => "plain",
                _ => "",
            };
            if (kind == "")
                continue;
            if (regions.Count == 0 || regions[^1] != kind)
                regions.Add(kind);
        }
        // ONE inserted region then ONE deleted region — no zip.
        Assert.Equal(new List<string> { "ins", "del" }, regions);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void SeamMerge_SkipsParagraphsCarryingPageBreaks()
    {
        // Word keeps a deleted page break PAGINATING — the deleted paragraph stays standalone so the
        // following struck content still starts on its own page. Merging it into the seam would
        // swallow the break and shift every subsequent page.
        static WmlDocument DocWithPageBreakPara(params string[] texts)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                var body = new Body();
                for (int i = 0; i < texts.Length; i++)
                {
                    var para = new Paragraph(new Run(new Text(texts[i])));
                    if (i == 0)
                        para.PrependChild(new ParagraphProperties(new PageBreakBefore()));
                    body.Append(para);
                }
                mainPart.Document = new Document(body);
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("pb.docx", stream.ToArray());
        }

        // 2 left × 1 right with zero token overlap: a pure del+ins gap (no similarity pairing, no
        // 1×1 residue), so the block seam is the only merge candidate — and must decline.
        var left = DocWithPageBreakPara("Obsolete removed prose.", "Ancient trailing words.");
        var right = Doc("Fresh writing appears.");

        var result = DocxDiff.Compare(left, right);

        var states = ParagraphStates(result);
        // No mixed (seam) paragraph: the page-break-carrying deleted paragraph stays standalone.
        Assert.DoesNotContain(states, s => s.State == "mixed");
        using (var stream = new MemoryStream(result.DocumentByteArray))
        using (var wdoc = WordprocessingDocument.Open(stream, false))
        {
            var pageBreakParas = wdoc.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Count(p => p.ParagraphProperties?.PageBreakBefore is not null);
            Assert.Equal(1, pageBreakParas);
        }

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void PureInsertAndPureDeleteGaps_AreUnaffected()
    {
        // Insert-only gap: no deletes to leapfrog; order = spine, ins, spine.
        var left = Doc("First shared.", "Last shared.");
        var right = Doc("First shared.", "Brand new middle.", "Last shared.");
        var insResult = DocxDiff.Compare(left, right);
        Assert.Equal(
            new List<(string, string)> { ("plain", "First shared."), ("ins", "Brand new middle."), ("plain", "Last shared.") },
            ParagraphStates(insResult));

        // Delete-only gap: order = spine, del, spine.
        var delResult = DocxDiff.Compare(right, left);
        Assert.Equal(
            new List<(string, string)> { ("plain", "First shared."), ("del", "Brand new middle."), ("plain", "Last shared.") },
            ParagraphStates(delResult));
    }

    [Fact]
    public void ReplaceGapInsideTableCell_InsertsBeforeDeletes()
    {
        static WmlDocument TableDoc(params string[] cellParagraphs)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                var cell = new WTableCell(cellParagraphs.Select(t => new Paragraph(new Run(new Text(t)))));
                var table = new WTable(
                    new TableProperties(new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 })),
                    new TableGrid(new GridColumn()),
                    new WTableRow(cell));
                mainPart.Document = new Document(new Body(table, new Paragraph()));
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("table.docx", stream.ToArray());
        }

        var left = TableDoc("Anchor cell line.", "Alpha ancient text.", "Bravo bygone text.");
        var right = TableDoc("Anchor cell line.", "Neon fresh words.", "Quantum modern words.");

        var result = DocxDiff.Compare(left, right);

        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var cellOut = wdoc.MainDocumentPart!.Document.Body!.Descendants<WTableCell>().Single();
        var cellStates = cellOut.Elements<Paragraph>()
            .Select(p =>
            {
                var hasIns = p.Descendants<InsertedRun>().Any();
                var hasDel = p.Descendants<DeletedRun>().Any();
                return hasIns && hasDel ? "mixed" : hasIns ? "ins" : hasDel ? "del" : "plain";
            })
            .ToList();
        var insIdx = cellStates.FindIndex(s => s == "ins");
        var delIdx = cellStates.FindIndex(s => s == "del");
        Assert.True(insIdx >= 0 && delIdx >= 0, $"expected ins and del cell paragraphs, got: {string.Join(",", cellStates)}");
        Assert.True(insIdx < delIdx, $"cell inserts must precede cell deletes; got: {string.Join(",", cellStates)}");

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void Interior_full_rewrite_keeps_separate_marked_paragraphs()
    {
        // A paired paragraph on either side makes the middle zero-lexical 1×1 residue explicit
        // body provenance. Word keeps the new and old paragraphs physically separate here.
        var left = Doc("Anchor title blue", "obsolete amber stanza", "shared trailing paragraph");
        var right = Doc("Anchor title bold", "fresh quantum clause", "shared trailing paragraph");

        var result = DocxDiff.Compare(left, right);
        var states = ParagraphStates(result);
        int ins = states.FindIndex(p => p.State == "ins" && p.Text == "fresh quantum clause");
        int del = states.FindIndex(p => p.State == "del" && p.Text == "obsolete amber stanza");
        Assert.True(ins >= 0 && del >= 0 && ins < del,
            $"expected separate inserted then deleted rewrite paragraphs, got: {string.Join(" | ", states)}");

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void Head_full_rewrite_before_a_real_body_pair_keeps_separate_marked_paragraphs()
    {
        // The common empty body paragraph is a real following paired block (unlike the trailing
        // section-break sentinel), matching the Word title-before-table family.
        var left = Doc("old unrelated heading words", "", "shared trailing paragraph");
        var right = Doc("new quantum report title", "", "shared trailing paragraph");

        var result = DocxDiff.Compare(left, right);
        var states = ParagraphStates(result);
        int ins = states.FindIndex(p => p.State == "ins" && p.Text == "new quantum report title");
        int del = states.FindIndex(p => p.State == "del" && p.Text == "old unrelated heading words");
        Assert.True(ins >= 0 && del >= 0 && ins < del,
            $"expected separate head rewrite paragraphs, got: {string.Join(" | ", states)}");

        Assert.Equal(BodyTexts(right), BodyTexts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(BodyTexts(left), BodyTexts(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void Tail_full_rewrite_before_section_break_keeps_the_normal_seam()
    {
        var left = Doc("Anchor title blue", "obsolete amber stanza");
        var right = Doc("Anchor title bold", "fresh quantum clause");

        var result = DocxDiff.Compare(left, right);
        var states = ParagraphStates(result);
        Assert.Contains(states, p => p.State == "mixed" &&
            p.Text.Contains("fresh quantum clause") && p.Text.Contains("obsolete amber stanza"));
        Assert.DoesNotContain(states, p => p.State == "ins" && p.Text == "fresh quantum clause");
        Assert.DoesNotContain(states, p => p.State == "del" && p.Text == "obsolete amber stanza");

        Assert.Equal(BodyTexts(right), BodyTexts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(BodyTexts(left), BodyTexts(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void Cell_full_rewrite_keeps_the_normal_seam()
    {
        // Cell alignment calls AlignBlocks, never the body-marking entry point. Even with real
        // paired neighbors, a cell's full lexical rewrite therefore stays on the established seam.
        var left = SingleCellTableDoc("Anchor title blue", "obsolete amber stanza", "shared trailing paragraph");
        var right = SingleCellTableDoc("Anchor title bold", "fresh quantum clause", "shared trailing paragraph");

        var result = DocxDiff.Compare(left, right);
        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var cell = wdoc.MainDocumentPart!.Document.Body!.Descendants<WTableCell>().Single();
        var target = cell.Elements<Paragraph>().Single(p => p.InnerText.Contains("fresh quantum clause"));
        Assert.NotEmpty(target.Descendants<InsertedRun>());
        Assert.NotEmpty(target.Descendants<DeletedRun>());

        Assert.Equal(BodyTexts(right), BodyTexts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(BodyTexts(left), BodyTexts(RevisionProcessor.RejectRevisions(result)));
    }
}
