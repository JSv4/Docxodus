#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Pins for the replace-gap paragraph-arrangement grammar, decoded from Word's compare output.
/// The arrangement is purely STRUCTURAL (no content-similarity discriminator):
/// <list type="bullet">
/// <item>Within an unmatched region Word emits all next-side content first (¶INS paragraphs),
/// then all base-side content (¶DEL paragraphs).</item>
/// <item>Fusion (a mixed ins+del paragraph) and shared (unmarked) pilcrows exist ONLY where the
/// region reaches the END of its story (body end / cell end), or via the empty-paragraph tail
/// chain — never for an interior wordful↔wordful replace.</item>
/// <item>The story's final pilcrow is immutable: a trailing replace fuses the last next paragraph's
/// inserted runs with the first base paragraph's deleted runs, and the LAST base paragraph keeps a
/// live (unmarked) pilcrow.</item>
/// <item>The tail chain pairs additional pilcrows backwards, opening only on an empty↔empty pair
/// and continuing while at least one side is empty; a table on either side blocks the chain.</item>
/// <item>A shared-pilcrow paragraph carries the NEXT side's paragraph properties, with
/// <c>w:pPrChange</c> recording the base side's when they differ. ¶INS/¶DEL-marked paragraphs
/// carry their own side's pPr and never a pPrChange.</item>
/// </list>
/// Accept ≡ right / reject ≡ left holds for every shape (pilcrow arithmetic: reject removes ¶INS
/// pilcrows, accept removes ¶DEL pilcrows, shared pilcrows survive both).
/// </summary>
public class DocxDiffGapArrangementTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ---------------------------------------------------------------- interior separation

    [Fact]
    public void InteriorReplace_BothWordful_StaysSeparateInsThenDel()
    {
        // anchor / L1 / anchor  →  anchor / R1 / anchor, zero shared words: Word keeps
        // [R1 ¶INS] [L1 ¶DEL] — no fusion mid-document.
        var left = Doc("shared head", "alpha bravo charlie", "shared tail");
        var right = Doc("shared head", "delta echo foxtrot", "shared tail");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void InteriorReplace_TwoByTwo_AllSeparate()
    {
        var left = Doc("shared head", "alpha bravo", "charlie delta", "shared tail");
        var right = Doc("shared head", "echo foxtrot", "golf hotel", "shared tail");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "INS:INS", "DEL:DEL", "DEL:DEL", "RET" }, shape);
        AssertRoundTrip(result, left, right);
    }

    // ---------------------------------------------------------------- trailing fusion

    [Fact]
    public void TrailingReplace_OneByOne_FusesWithLivePilcrow()
    {
        // Document-final 1×1 rewrite: one mixed paragraph, ins before del, UNMARKED pilcrow.
        var left = Doc("shared head", "alpha bravo charlie");
        var right = Doc("shared head", "delta echo foxtrot");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "MIXED:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void TrailingReplace_TwoByOne_FusedParagraphCarriesDelMark_LastBaseKeepsLivePilcrow()
    {
        // m=2, n=1: [R1-ins + L1-del ¶DEL] [L2-del, unmarked ¶].
        var left = Doc("shared head", "alpha bravo", "charlie delta");
        var right = Doc("shared head", "echo foxtrot");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "MIXED:DEL", "DEL:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void TrailingReplace_WordfulPairAboveStructuralPair_DoesNotChain()
    {
        // base [W1, E] → next [N1, E]: the structural final pair is the empty↔empty pair, and the
        // wordful pair above it does NOT chain (the structural pair does not itself open a chain,
        // and W1↔N1 is wordful↔wordful). Emission: [N1 ¶INS] [W1 ¶DEL] [empty, unmarked ¶].
        var left = Doc("shared head", "alpha bravo charlie", "");
        var right = Doc("shared head", "delta echo foxtrot", "");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "DEL:DEL", "EMPTY:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void TrailingSurplusEmptyDelete_StaysMarked_NoVirtualPairOnSingleSidedGap()
    {
        // base [W1, E1, E2] → next [N1, E]: a blank never anchors on its own evidence, so E1 does NOT
        // pair with E across the unrelated W1/N1 content; only the story-final pilcrows pair (E2 ↔ E,
        // structural). The region is ONE trailing replace whose tail chain cannot open (E1 ↔ N1 is
        // empty↔wordful): N1 ¶INS first, then W1 and the surplus E1 fully ¶DEL, then the shared
        // final pilcrow — no stray live empty paragraph survives into the accepted view.
        var left = Doc("shared head", "alpha bravo charlie", "", "");
        var right = Doc("shared head", "delta echo foxtrot", "");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "DEL:DEL", "EMPTY_DEL:DEL", "EMPTY:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void TrailingSurplusInsert_InteriorReplace_StaysSeparate()
    {
        // base [W1, E1, E2] → next [N1, N2, E]: only the story-final pilcrows pair (E2 ↔ E); the
        // unsupported blank E1 stays with its own side. One trailing replace whose chain cannot open
        // (E1 ↔ N2 is empty↔wordful): inserts before deletes, the surplus E1 fully ¶DEL.
        var left = Doc("shared head", "alpha bravo charlie", "", "");
        var right = Doc("shared head", "delta echo foxtrot", "golf hotel", "");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "INS:INS", "DEL:DEL", "EMPTY_DEL:DEL", "EMPTY:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void TrailingSurplusEmptyDeletes_AfterAnchoredEmpties_StayMarked()
    {
        // base [W1, E1, E2, E3] → next [N1, E1', E2']: no blank anchors on its own evidence, so the
        // whole region is one trailing replace and the tail chain runs backwards from the structural
        // final pair: E3 ↔ E2' (shared), E2 ↔ E1' (empty↔empty opens the chain), E1 ↔ N1 (continues
        // while one side is empty) — N1's inserted runs fuse into W1's ¶DEL paragraph and every
        // paired pilcrow is shared. Accept coalesces the fused runs forward onto the first shared
        // pilcrow (= next), reject drops the inserted runs and restores W1's mark (= base).
        var left = Doc("shared head", "alpha bravo charlie", "", "", "");
        var right = Doc("shared head", "delta echo foxtrot", "", "");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "MIXED:DEL", "EMPTY:-", "EMPTY:-", "EMPTY:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void WordfulFinalPair_BlocksTailChain_BlanksStayMarked()
    {
        // base [W1, E1, E2] → next [N1, E1', N2]: the structural final pair is E2 ↔ N2 with a
        // WORDFUL next member, so no earlier pilcrow can pair — N2's words sit between its pilcrow
        // and E1's. The empty↔empty candidate E1 ↔ E1' therefore does NOT open a chain: E1' stays
        // ¶INS ahead of the deletions and E1 stays ¶DEL, while N2's runs fuse at the head (into W1's
        // ¶DEL paragraph) and E2 keeps the shared final pilcrow — the reference output's shape.
        var left = Doc("shared head", "alpha bravo charlie", "", "");
        var right = Doc("shared head", "delta echo foxtrot", "", "golf hotel india");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "EMPTY_INS:INS", "MIXED:DEL", "EMPTY_DEL:DEL", "EMPTY:-" }, shape);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void RestyledBlankAfterEditedParagraph_PairsWithFormatChange()
    {
        // base [W1, E, W2] → next [W1', E(centered), W2']: the blank follows an EDITED pair, so it
        // pairs with the blank across from it although the next side restyled it — a retained mark
        // carrying the next side's pPr with a pPrChange (as Word writes it), never a deleted blank
        // followed by an inserted one.
        var left = Doc(("alpha bravo charlie delta", null), ("", null), ("echo foxtrot golf hotel", null));
        var right = Doc(("alpha bravo charlie ZULU", null), ("", "center"), ("echo foxtrot golf YANKEE", null));
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(3, shape.Length);
        Assert.Equal("EMPTY:-", shape[1]);
        var blank = BodyParas(result)[1];
        Assert.NotNull(blank.Element(W + "pPr")?.Element(W + "pPrChange"));
        Assert.Equal("center", (string?)blank.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void InteriorRegion_TrailingBlanks_ShareOnlyWhenParagraphCountsMatch()
    {
        // Interior regions end at the shared "tail" paragraph. Equal counts (3 vs 3: [W1, E, E] vs
        // [N1, N2, E]) open the backward chain: E ↔ E shared, then E ↔ N2 (one side empty) shared with
        // N2's runs fused into W1's ¶DEL paragraph. Unequal counts (2 vs 3: [W1, E] vs [N1, E, E])
        // block it: every mark stays on its own side, inserts first.
        var equalLeft = Doc("alpha bravo charlie", "", "", "shared tail words");
        var equalRight = Doc("delta echo foxtrot", "golf hotel india", "", "shared tail words");
        var equal = BodyShape(DocxDiff.Compare(equalLeft, equalRight));
        Assert.Equal(new[] { "INS:INS", "MIXED:DEL", "EMPTY:-", "EMPTY:-", "RET" }, equal);
        AssertRoundTrip(DocxDiff.Compare(equalLeft, equalRight), equalLeft, equalRight);

        var unequalLeft = Doc("alpha bravo charlie", "", "shared tail words");
        var unequalRight = Doc("delta echo foxtrot", "", "", "shared tail words");
        var unequal = BodyShape(DocxDiff.Compare(unequalLeft, unequalRight));
        Assert.Equal(new[] { "INS:INS", "EMPTY_INS:INS", "EMPTY_INS:INS", "DEL:DEL", "EMPTY_DEL:DEL", "RET" }, unequal);
        AssertRoundTrip(DocxDiff.Compare(unequalLeft, unequalRight), unequalLeft, unequalRight);
    }

    [Fact]
    public void InteriorRegion_OneSideOnlyABlank_SharesIt()
    {
        // [W1, W2, E] vs [E] before the shared tail: the next side is nothing but the trailing blank,
        // so it shares the base's trailing blank mark (the deletions stay ahead of it).
        var left = Doc("alpha bravo charlie", "delta echo foxtrot", "", "shared tail words");
        var right = Doc("", "shared tail words");
        var result = DocxDiff.Compare(left, right);
        Assert.Equal(new[] { "DEL:DEL", "DEL:DEL", "EMPTY:-", "RET" }, BodyShape(result));
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void UnrelatedContent_UniqueBlankNeverAnchors_RegionStaysWhole()
    {
        // base [W1, W2, E, W3] → next [E, N1, N2]: the only content-equal block pair is the blank
        // (unique on each side, so the exact-match spine would pin it). A blank has no words — its
        // pilcrow is "the same paragraph" only inside matched content — so it must NOT split these
        // unrelated regions into two halves with the deletions scattered between them. The whole
        // body is one replace: next content ¶INS first (the leading blank included), then the base
        // content ¶DEL, fused at the structural final pair. No live empty paragraph survives mid-body.
        var left = Doc("alpha bravo charlie", "delta echo", "", "foxtrot golf");
        var right = Doc("", "hotel india juliet", "kilo lima");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.DoesNotContain("EMPTY:-", shape);
        Assert.Equal("EMPTY_INS:INS", shape[0]);
        Assert.Equal(new[] { "INS:INS" }, shape.Skip(1).TakeWhile(c => c == "INS:INS").ToArray());
        Assert.DoesNotContain("DEL:DEL", shape.Take(2));
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void BlankBesideEditedNeighbours_StaysRetained()
    {
        // base [W1, E, W2] → next [W1', E, W2'] with each neighbour EDITED (shares words with its
        // counterpart): the blank sits inside matched content, so it keeps its retained (shared,
        // unmarked) pilcrow exactly as before — support comes from the paired neighbours, not from
        // the blank's own (empty) content.
        var left = Doc("alpha bravo charlie delta", "", "echo foxtrot golf hotel");
        var right = Doc("alpha bravo charlie ZULU", "", "echo foxtrot golf YANKEE");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(3, shape.Length);
        Assert.Equal("EMPTY:-", shape[1]);
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void BaseEndsWithTable_NextEndsWithBlank_FinalPilcrowLandsAfterDeletedTable()
    {
        // base [W1, TBL] → next [N1, E]: the base story physically ends with a table, so the virtual
        // structural pair applies — and with an EMPTY final next paragraph there is nothing to fuse:
        // N1 stays ¶INS ahead of the deletions, W1 and the whole table are deleted, and the next
        // side's final pilcrow lands AFTER the deleted table as the document's last paragraph (Word
        // keeps it live there; our contract marks it ¶INS so reject ≡ base exactly). Before this rule
        // the trailing blank was emitted with the other inserts and the document ENDED with the
        // deleted table.
        var left = TableDoc("shared head", "alpha bravo charlie", new[] { "cell one", "cell two" }, afterTable: null);
        var right = Doc("shared head", "delta echo foxtrot", "");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "INS:INS", "DEL:DEL", "EMPTY_INS:INS" }, shape);
        using (var stream = new MemoryStream(result.DocumentByteArray))
        using (var word = WordprocessingDocument.Open(stream, false))
        {
            var blocks = word.MainDocumentPart!.Document!.Body!.ChildElements
                .Where(e => e is DocumentFormat.OpenXml.Wordprocessing.Paragraph or DocumentFormat.OpenXml.Wordprocessing.Table)
                .ToList();
            Assert.IsType<DocumentFormat.OpenXml.Wordprocessing.Table>(blocks[^2]);
            Assert.IsType<DocumentFormat.OpenXml.Wordprocessing.Paragraph>(blocks[^1]);
        }
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void PairedRows_SurplusBaseCell_StaysInRowAsDeletedCell()
    {
        // base row [c1, c2, c3] → next row [c1, C2']: the rows pair, the first two cells pair, and the
        // third cell has no next counterpart. The table stays ONE modified table: the surplus base
        // cell is emitted in place, whole-marked deleted (w:tcPr/w:cellDel + struck content), so
        // accept removes it (= next) and reject restores it (= base). It used to lower the whole
        // table to a deleted table followed by an inserted one.
        var left = TableDoc("shared head", "before table", new[] { "cell one", "cell two", "cell three" }, "after table");
        var right = TableDoc("shared head", "before table", new[] { "cell one", "cell TWO" }, "after table");
        var result = DocxDiff.Compare(left, right);

        using (var stream = new MemoryStream(result.DocumentByteArray))
        using (var word = WordprocessingDocument.Open(stream, false))
        {
            var body = word.MainDocumentPart!.Document!.Body!;
            var table = Assert.Single(body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>());
            var row = Assert.Single(table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>());
            var cells = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList();
            Assert.Equal(3, cells.Count);
            Assert.Null(cells[0].TableCellProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.CellDeletion>());
            Assert.Null(cells[1].TableCellProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.CellDeletion>());
            Assert.NotNull(cells[2].TableCellProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.CellDeletion>());
            Assert.Equal("cell three", string.Concat(cells[2].Descendants<DocumentFormat.OpenXml.Wordprocessing.DeletedText>().Select(t => t.Text)));
        }
        Assert.Equal(BodyTextOf(right), BodyTextOf(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(BodyTextOf(left), BodyTextOf(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void TrailingReplace_AcrossDeletedTable_AcceptStaysPilcrowExact()
    {
        // base [W1, TBL, W2] → next [N1]: the trailing replace fuses N1's runs at the head and the
        // live pilcrow lands on W2, BEYOND the deleted table. Accept must coalesce the fused
        // content across the (entirely removed) table onto that live pilcrow — no stray empty
        // paragraph may survive into the accepted view.
        var left = TableDoc("shared head", "alpha bravo charlie", new[] { "cell one", "cell two" }, "delta echo");
        var right = Doc("shared head", "foxtrot golf hotel");
        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal(BodyTextOf(right), BodyTextOf(accepted));
        using (var stream = new MemoryStream(accepted.DocumentByteArray))
        using (var word = WordprocessingDocument.Open(stream, false))
        {
            var body = word.MainDocumentPart!.Document!.Body!;
            Assert.Equal(2, body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count());
            Assert.Empty(body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>());
        }
        Assert.Equal(BodyTextOf(left), BodyTextOf(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void TrailingReplace_BaseEndsWithTable_VirtualPairKeepsBaseMarkedAndInsertedTailPilcrow()
    {
        // Decoded from Word's compare output (the base story physically ENDS with a table): the
        // structural final pair goes VIRTUAL — the last next paragraph's runs still fuse into the
        // FIRST deleted paragraph (¶DEL), every remaining base block stays fully marked (¶DEL,
        // deleted table), and the NEXT side's final pilcrow is retained AFTER the deleted table as
        // an empty paragraph. Word leaves that pilcrow live (its own reject then strands a stray
        // empty — an infidelity our exact contract forbids); we mark it ¶INS instead, which renders
        // identically (an empty final line) while keeping reject ≡ left exact.
        var left = TableTailDoc("shared head", new[] { "alpha bravo", "charlie delta" }, new[] { "cell one", "cell two" });
        var right = Doc("shared head", "echo foxtrot");
        var result = DocxDiff.Compare(left, right);

        var shape = BodyShape(result);
        Assert.Equal(new[] { "RET", "MIXED:DEL", "DEL:DEL", "EMPTY_INS:INS" }, shape);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        Assert.Equal(BodyTextOf(right), BodyTextOf(accepted));
        using (var stream = new MemoryStream(accepted.DocumentByteArray))
        using (var word = WordprocessingDocument.Open(stream, false))
        {
            var body = word.MainDocumentPart!.Document!.Body!;
            Assert.Equal(2, body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count());
            Assert.Empty(body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>());
        }

        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTextOf(left), BodyTextOf(rejected));
        using (var stream = new MemoryStream(rejected.DocumentByteArray))
        using (var word = WordprocessingDocument.Open(stream, false))
        {
            var body = word.MainDocumentPart!.Document!.Body!;
            // Reject restores the base exactly: no stray empty paragraph after the table.
            Assert.Equal(3, body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count());
            Assert.Single(body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>());
        }
    }

    [Fact]
    public void TrailingReplace_BaseEndsWithTable_TailPilcrowCarriesNextPPrWithoutPPrChange()
    {
        // The ¶INS tail pilcrow is the NEXT side's paragraph mark: it carries the next paragraph's
        // own pPr and — like every marked paragraph — never a pPrChange.
        var left = TableTailDoc("shared head", new[] { "alpha bravo" }, new[] { "cell one" });
        var right = Doc(("shared head", null), ("echo foxtrot", "right"));
        var result = DocxDiff.Compare(left, right);

        var paras = BodyParas(result);
        var tail = paras[^1];
        Assert.Equal("INS", ParaMark(tail));
        Assert.Equal("right", (string?)tail.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        Assert.Null(tail.Element(W + "pPr")?.Element(W + "pPrChange"));
        AssertRoundTrip(result, left, right);
    }

    private static WmlDocument TableTailDoc(string head, string[] paras, string[] cells)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(head))));
            foreach (var text in paras)
                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(text))));
            var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
            foreach (var cell in cells)
                row.Append(new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                    new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth { Width = "2000", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa }),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cell)))));
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Table(
                new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.TableWidth { Width = "0", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto }),
                new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                    cells.Select(_ => new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" })),
                row));
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(body);
            doc.Save();
        }
        return new WmlDocument("gap-arrangement-table-tail.docx", stream.ToArray());
    }

    private static WmlDocument TableDoc(string head, string beforeTable, string[] cells, string? afterTable)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(head))));
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(beforeTable))));
            var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
            foreach (var cell in cells)
                row.Append(new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                    new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth { Width = "2000", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa }),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(cell)))));
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Table(
                new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.TableWidth { Width = "0", Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto }),
                new DocumentFormat.OpenXml.Wordprocessing.TableGrid(
                    cells.Select(_ => new DocumentFormat.OpenXml.Wordprocessing.GridColumn { Width = "2000" })),
                row));
            if (afterTable is not null)
                body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(afterTable))));
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(body);
            doc.Save();
        }
        return new WmlDocument("gap-arrangement-table.docx", stream.ToArray());
    }

    // ---------------------------------------------------------------- pPr / pPrChange discipline

    [Fact]
    public void SharedPilcrowParagraph_CarriesNextPPr_WithPPrChangeRecordingBase()
    {
        // Trailing 1×1 fuse where base is centered and next is right-aligned: the fused (shared
        // pilcrow) paragraph carries the NEXT side's jc with w:pPrChange recording the base side's.
        var left = Doc(("shared head", null), ("alpha bravo charlie", "center"));
        var right = Doc(("shared head", null), ("delta echo foxtrot", "right"));
        var result = DocxDiff.Compare(left, right);

        var paras = BodyParas(result);
        var fused = paras[^1];
        Assert.Equal("right", (string?)fused.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        var change = fused.Element(W + "pPr")?.Element(W + "pPrChange");
        Assert.NotNull(change);
        Assert.Equal("center", (string?)change!.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));
        AssertRoundTrip(result, left, right);
    }

    [Fact]
    public void MarkedGapParagraphs_NeverCarryPPrChange()
    {
        var left = Doc("shared head", "alpha bravo", "charlie delta", "shared tail");
        var right = Doc("shared head", "echo foxtrot", "golf hotel", "shared tail");
        var result = DocxDiff.Compare(left, right);

        foreach (var p in BodyParas(result))
        {
            var mark = ParaMark(p);
            if (mark is "INS" or "DEL")
                Assert.Null(p.Element(W + "pPr")?.Element(W + "pPrChange"));
        }
    }

    // ---------------------------------------------------------------- helpers

    private static WmlDocument Doc(params string[] paraTexts) =>
        Doc(paraTexts.Select(t => (t, (string?)null)).ToArray());

    private static WmlDocument Doc(params (string Text, string? Jc)[] paras)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            foreach (var (text, jc) in paras)
            {
                var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                if (jc is not null)
                {
                    p.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.Justification
                        {
                            Val = jc switch
                            {
                                "center" => DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center,
                                "right" => DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Right,
                                _ => DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left,
                            },
                        });
                }
                if (text.Length > 0)
                    p.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text(text)));
                body.Append(p);
            }
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(body);
            doc.Save();
        }
        return new WmlDocument("gap-arrangement.docx", stream.ToArray());
    }

    private static List<XElement> BodyParas(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(word.MainDocumentPart!.GetStream());
        var xdoc = XDocument.Parse(reader.ReadToEnd());
        return xdoc.Root!.Element(W + "body")!.Elements(W + "p").ToList();
    }

    private static string? ParaMark(XElement p)
    {
        var rPr = p.Element(W + "pPr")?.Element(W + "rPr");
        if (rPr?.Element(W + "ins") is not null) return "INS";
        if (rPr?.Element(W + "del") is not null) return "DEL";
        return null;
    }

    /// <summary>Category:mark per body paragraph — INS/DEL/MIXED/EMPTY[_INS|_DEL]/RET, mark INS/DEL/-.</summary>
    private static string[] BodyShape(WmlDocument doc)
    {
        var shapes = new List<string>();
        foreach (var p in BodyParas(doc))
        {
            string insText = string.Concat(p.Descendants(W + "ins").SelectMany(i => i.Descendants(W + "t")).Select(t => t.Value));
            string delText = string.Concat(p.Descendants(W + "delText").Select(t => t.Value));
            string retText = string.Concat(p.Descendants(W + "t")
                .Where(t => t.Ancestors(W + "ins").FirstOrDefault() is null)
                .Select(t => t.Value));
            var mark = ParaMark(p) ?? "-";
            string cat;
            if (insText.Trim().Length > 0 && delText.Trim().Length > 0) cat = "MIXED";
            else if (insText.Trim().Length > 0) cat = "INS";
            else if (delText.Trim().Length > 0) cat = "DEL";
            else if (retText.Trim().Length > 0) cat = "RET";
            else cat = mark switch { "INS" => "EMPTY_INS", "DEL" => "EMPTY_DEL", _ => "EMPTY" };
            shapes.Add(cat == "RET" ? "RET" : $"{cat}:{mark}");
        }
        return shapes.ToArray();
    }

    private static void AssertRoundTrip(WmlDocument redline, WmlDocument left, WmlDocument right)
    {
        Assert.Equal(BodyTextOf(right), BodyTextOf(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(BodyTextOf(left), BodyTextOf(RevisionProcessor.RejectRevisions(redline)));
    }

    private static string BodyTextOf(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        var body = word.MainDocumentPart!.Document!.Body!;
        return string.Join(" ", body
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
            .Select(p => string.Concat(p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text))));
    }
}
