using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// A formatting change inside a table cell whose text is unchanged is tracked as the same change in a body
/// paragraph is: run formatting, paragraph properties and the paragraph mark (issue #1797).
/// </summary>
public class TableCellFormatTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string Paragraph(string text, string pPr = "", string rPr = "") =>
        $"<w:p>{(pPr.Length == 0 ? "" : $"<w:pPr>{pPr}</w:pPr>")}" +
        $"<w:r>{(rPr.Length == 0 ? "" : $"<w:rPr>{rPr}</w:rPr>")}<w:t>{text}</w:t></w:r></w:p>";

    private static string Cell(string content) => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{content}</w:tc>";

    private static string Table(string tblPr, params string[] rows) =>
        $"<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>{tblPr}</w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>" +
        string.Concat(rows.Select(cells => $"<w:tr>{cells}</w:tr>")) + "</w:tbl>";

    private static WmlDocument Document(string body) => IrTestDocuments.FromBodyXml(body + Paragraph("After."));

    private static XElement Body(WmlDocument doc)
    {
        using var stream = new System.IO.MemoryStream(doc.DocumentByteArray);
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
    }

    /// <summary>
    /// Every cell paragraph as its properties and its runs' properties and text: adjacent runs with the same
    /// properties join, and an empty property container (as a rejected change leaves) reads as absent.
    /// </summary>
    private static string[] CellParagraphs(WmlDocument doc) => Body(doc).Descendants(W + "tc").Elements(W + "p")
        .Select(p =>
        {
            static string Props(XElement? e) => e is null ? "" : string.Concat(e.Elements().Select(c =>
                c.Name.LocalName is "rPr" or "pPr" && !c.HasElements ? "" :
                $"{c.Name.LocalName}({string.Join(",", c.Attributes().Select(a => a.Value))}{Props(c)})"));
            var runs = p.Descendants(W + "r").Select(r => (Props: Props(r.Element(W + "rPr")), Text: r.Value))
                .Aggregate(new System.Collections.Generic.List<(string Props, string Text)>(), (list, run) =>
                {
                    if (list.Count > 0 && list[^1].Props == run.Props) list[^1] = (run.Props, list[^1].Text + run.Text);
                    else list.Add(run);
                    return list;
                });
            return Props(p.Element(W + "pPr")) + string.Concat(runs.Select(r => $"[{r.Props}]{r.Text}"));
        }).ToArray();

    [Theory]
    [InlineData("", "<w:b/>", "rPrChange", DocxDiffFormatChangeScope.Run)]
    [InlineData("<w:jc w:val=\"center\"/>", "", "pPrChange", DocxDiffFormatChangeScope.Paragraph)]
    [InlineData("<w:rPr><w:sz w:val=\"72\"/></w:rPr>", "", "rPrChange", DocxDiffFormatChangeScope.Paragraph)]
    public void A_format_change_in_a_cell_is_tracked_and_rejects_to_the_original(
        string pPr, string rPr, string marker, DocxDiffFormatChangeScope scope)
    {
        var left = Document(Table("", Cell(Paragraph("Cell text."))));
        var right = Document(Table("", Cell(Paragraph("Cell text.", pPr, rPr))));

        var redline = DocxCompare.Compare(left, right);
        Assert.Single(Body(redline).Descendants(W + "tc").Descendants(W + marker));
        Assert.Equal(CellParagraphs(left), CellParagraphs(RevisionProcessor.RejectRevisions(redline)));
        Assert.Equal(CellParagraphs(right), CellParagraphs(RevisionProcessor.AcceptRevisions(redline)));
        foreach (var comparison in new[] { DocxDiffFormatComparison.ModeledOnly, DocxDiffFormatComparison.Full })
        {
            var revision = Assert.Single(DocxDiff.GetRevisions(left, right, new DocxDiffSettings { FormatComparison = comparison }));
            Assert.Equal(DocxDiffRevisionType.FormatChanged, revision.Type);
            Assert.Equal(scope, revision.FormatChange!.Scope);
        }
    }

    [Fact]
    public void An_unmodeled_change_in_a_cell_stays_untracked()
    {
        // Control: the default policy does not model a run's language, in a cell as in the body.
        var left = Document(Table("", Cell(Paragraph("Cell text."))));
        var right = Document(Table("", Cell(Paragraph("Cell text.", rPr: "<w:lang w:val=\"nb-NO\"/>"))));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(Body(DocxCompare.Compare(left, right)).Descendants(W + "rPrChange"));
    }

    [Fact]
    public void Format_changes_beside_a_text_edit_are_tracked_and_unchanged_cells_keep_no_marker()
    {
        // The first row's text changes, so its other cell is compared on its own; the second row's text does not.
        var left = Document(Table("",
            Cell(Paragraph("A1")) + Cell(Paragraph("B1")),
            Cell(Paragraph("A2")) + Cell(Paragraph("B2"))));
        var right = Document(Table("",
            Cell(Paragraph("A1 changed")) + Cell(Paragraph("B1", rPr: "<w:b/>")),
            Cell(Paragraph("A2")) + Cell(Paragraph("B2", rPr: "<w:i/>"))));

        var redline = DocxCompare.Compare(left, right);
        var cells = Body(redline).Descendants(W + "tc").ToList();
        Assert.Single(cells[0].Descendants(W + "ins"));
        Assert.Single(cells[1].Descendants(W + "rPrChange"));
        Assert.DoesNotContain(cells[2].Descendants(), e => e.Name.LocalName.EndsWith("Change") || e.Name == W + "ins" || e.Name == W + "del");
        Assert.Single(cells[3].Descendants(W + "rPrChange"));
        Assert.Equal(CellParagraphs(left), CellParagraphs(RevisionProcessor.RejectRevisions(redline)));
        Assert.Equal(CellParagraphs(right), CellParagraphs(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(new[] { "B1", "B2" }, DocxDiff.GetRevisions(left, right)
            .Where(r => r.Type == DocxDiffRevisionType.FormatChanged).Select(r => r.Text));
    }

    [Fact]
    public void A_table_property_change_and_a_cell_format_change_are_both_tracked()
    {
        var left = Document(Table("", Cell(Paragraph("Cell text."))));
        var right = Document(Table("<w:jc w:val=\"center\"/>", Cell(Paragraph("Cell text.", rPr: "<w:b/>"))));

        var redline = DocxCompare.Compare(left, right);
        Assert.Single(Body(redline).Descendants(W + "tblPrChange"));
        Assert.Single(Body(redline).Descendants(W + "tc").Descendants(W + "rPrChange"));
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(CellParagraphs(left), CellParagraphs(rejected));
        Assert.Null(Body(rejected).Element(W + "tbl")!.Element(W + "tblPr")!.Element(W + "jc"));
        Assert.Equal(new[] { DocxDiffFormatChangeScope.Run, DocxDiffFormatChangeScope.Table },
            DocxDiff.GetRevisions(left, right).Select(r => r.FormatChange!.Scope).OrderBy(s => s));
    }

    [Fact]
    public void A_format_change_in_a_nested_table_cell_is_tracked()
    {
        static WmlDocument Nested(string rPr) => Document(Table("",
            Cell(Table("", Cell(Paragraph("Inner text.", rPr: rPr))) + Paragraph(""))));
        var (left, right) = (Nested(""), Nested("<w:b/>"));

        var redline = DocxCompare.Compare(left, right);
        Assert.Single(Body(redline).Descendants(W + "rPrChange"));
        Assert.Equal(CellParagraphs(left), CellParagraphs(RevisionProcessor.RejectRevisions(redline)));
        Assert.Single(DocxDiff.GetRevisions(left, right));
    }

    [Fact]
    public void An_unmodeled_change_in_a_nested_table_leaves_the_table_unmodified()
    {
        // Control: a nested table's language change is not a format change of the outer table's cell, so the
        // outer table is not taken apart row by row for it.
        static WmlDocument Nested(string rPr) => Document(Table("",
            Cell(Table("", Cell(Paragraph("Inner text.", rPr: rPr))) + Paragraph(""))));
        var (left, right) = (Nested(""), Nested("<w:lang w:val=\"nb-NO\"/>"));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
        Assert.DoesNotContain("\"tableDiff\"", DocxDiff.GetEditScriptJson(left, right));
    }

    [Theory]
    [InlineData("<w:jc w:val=\"center\"/>", "", "", "", "tblPrChange")]
    [InlineData("", "2400", "", "", "tblGridChange")]
    [InlineData("", "", "<w:trPr><w:cantSplit/></w:trPr>", "", "trPrChange")]
    [InlineData("", "", "", "<w:tblPrEx><w:jc w:val=\"center\"/></w:tblPrEx>", "tblPrExChange")]
    public void A_property_change_of_a_nested_table_is_tracked_inside_its_cell(
        string tblPr, string gridWidth, string trPr, string tblPrEx, string marker)
    {
        static string Nested(string tblPr, string grid, string rowProperties) =>
            $"<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>{tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w=\"{grid}\"/></w:tblGrid>" +
            $"<w:tr>{rowProperties}{Cell(Paragraph("Inner text."))}</w:tr></w:tbl>";
        static XElement Shell(WmlDocument doc)
        {
            var nested = new XElement(Body(doc).Descendants(W + "tc").Descendants(W + "tbl").Single());
            nested.Elements(W + "tr").Elements(W + "tc").Remove();
            // Empty property containers (a rejected change can leave one) read as absent.
            foreach (var e in nested.Descendants().Reverse().ToList())
                if (!e.HasElements && !e.HasAttributes && e.Name.LocalName is "tr" or "trPr" or "tblPrEx")
                    e.Remove();
            return nested;
        }
        var left = Document(Table("", Cell(Nested("", "3000", "") + Paragraph(""))));
        var right = Document(Table("", Cell(Nested(tblPr, gridWidth.Length == 0 ? "3000" : gridWidth, tblPrEx + trPr) + Paragraph(""))));

        var redline = DocxCompare.Compare(left, right);
        Assert.Single(Body(redline).Descendants(W + "tc").Descendants(W + marker));
        Assert.Equal(Shell(left).ToString(SaveOptions.DisableFormatting), Shell(RevisionProcessor.RejectRevisions(redline)).ToString(SaveOptions.DisableFormatting));
    }

    [Fact]
    public void A_text_edit_beside_a_content_control_cell_stays_a_cell_level_edit()
    {
        // A cell a row-level content control holds is never compared for formatting, so the row keeps the cell-level
        // edit main makes of a text change elsewhere in the table rather than a whole-table replacement (issue #1804).
        static WmlDocument Revision(string text, string rPr) => Document(Table("",
            Cell(Paragraph(text)) + Cell(Paragraph("B1")),
            Cell(Paragraph("A2")) + ContentControl(Cell(Paragraph("B2", rPr: rPr)))));
        var (left, right) = (Revision("A1", ""), Revision("A1 changed", "<w:b/>"));

        var redline = DocxCompare.Compare(left, right);
        Assert.Single(Body(redline).Elements(W + "tbl"));
        Assert.Single(Body(redline).Descendants(W + "ins"));
        Assert.Equal(CellParagraphs(right), CellParagraphs(RevisionProcessor.AcceptRevisions(redline)));
    }

    [Fact]
    public void A_cell_paragraph_change_follows_the_body_when_block_format_tracking_is_off()
    {
        // A parity pin: main already leaves these untracked. A nested table's property change is not tracked either,
        // so it does not take the outer table apart row by row.
        static WmlDocument Nested(string tblPr) => Document(Table("",
            Cell(Table(tblPr, Cell(Paragraph("Inner text."))) + Paragraph(""))));
        Assert.DoesNotContain("\"tableDiff\"", DocxDiff.GetEditScriptJson(Nested(""), Nested("<w:jc w:val=\"center\"/>"),
            new DocxDiffSettings { TrackBlockFormatChanges = false }));
        var settings = new DocxDiffSettings { TrackBlockFormatChanges = false };
        foreach (var wrap in new System.Func<string, string>[] { p => p, p => Table("", Cell(p)) })
        {
            var (left, right) = (Document(wrap(Paragraph("Cell text."))), Document(wrap(Paragraph("Cell text.", "<w:jc w:val=\"center\"/>"))));
            Assert.Empty(DocxDiff.GetRevisions(left, right, settings));
            Assert.Empty(Body(DocxCompare.Compare(left, right, settings)).Descendants(W + "pPrChange"));
        }
    }

    private static string ContentControl(string content) =>
        $"<w:sdt><w:sdtPr><w:id w:val=\"9\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>";

    /// <summary>
    /// A table with a row or cell that a content control delivers is still copied whole (issue #1804): the
    /// modified-table renderer would write such a row twice, so its cell's change stays untracked but its rows
    /// come out once. <c>row</c> bolds a cell beside a row in a content control, <c>cell</c> bolds a cell a
    /// row-level content control holds, and <c>noise</c> changes only a nested table's language beside such a row.
    /// </summary>
    [Theory]
    [InlineData("row", false)]
    [InlineData("row", true)]
    [InlineData("cell", false)]
    [InlineData("noise", false)]
    public void A_table_with_a_content_control_row_or_cell_keeps_each_row_once(string shape, bool consolidate)
    {
        WmlDocument Revision(bool revised)
        {
            string bold = revised && shape != "noise" ? "<w:b/>" : "";
            string inner = Table("", Cell(Paragraph("Inner text.", rPr: revised && shape == "noise" ? "<w:lang w:val=\"nb-NO\"/>" : "")));
            string first = Cell(Paragraph("A1")) + (shape == "cell"
                ? ContentControl(Cell(Paragraph("B1", rPr: bold)))
                : Cell(Paragraph("B1", rPr: bold) + (shape == "noise" ? inner + Paragraph("") : "")));
            string table = Table("", first);
            return Document(shape == "cell" ? table : table.Replace("</w:tbl>",
                ContentControl($"<w:tr>{Cell(Paragraph("A2"))}{Cell(Paragraph("B2"))}</w:tr>") + "</w:tbl>"));
        }
        var (left, right) = (Revision(false), Revision(true));
        var redline = consolidate
            ? DocxCompare.Consolidate(left, new[] { new DocxDiffReviewer { Document = right, Author = "Reviewer A" } })
            : DocxCompare.Compare(left, right);

        static int Rows(WmlDocument doc) => Body(doc).Descendants(W + "tr").Count();
        Assert.DoesNotContain(Body(redline).Descendants(W + "tbl").Descendants(), e => e.Name == W + "ins" || e.Name == W + "del");
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        Assert.Equal(Rows(right), Rows(accepted));
        Assert.Equal(CellParagraphs(right), CellParagraphs(accepted));
        Assert.Equal(Rows(left), Rows(RevisionProcessor.RejectRevisions(redline)));
        if (shape == "noise")
            Assert.Empty(DocxDiff.GetRevisions(left, right));
    }

    [Fact]
    public void Consolidating_a_cell_format_change_tracks_it_under_the_reviewer()
    {
        var original = Document(Table("", Cell(Paragraph("A1")) + Cell(Paragraph("B1"))));
        var reviewed = Document(Table("", Cell(Paragraph("A1", rPr: "<w:b/>")) + Cell(Paragraph("B1"))));

        var merged = DocxCompare.Consolidate(original, new[] { new DocxDiffReviewer { Document = reviewed, Author = "Reviewer A" } });
        Assert.Equal("Reviewer A", (string?)Assert.Single(Body(merged).Descendants(W + "rPrChange")).Attribute(W + "author"));
        Assert.Equal(CellParagraphs(original), CellParagraphs(RevisionProcessor.RejectRevisions(merged)));
        Assert.Equal(CellParagraphs(reviewed), CellParagraphs(RevisionProcessor.AcceptRevisions(merged)));
    }

    [Fact]
    public void Consolidating_a_cell_format_change_and_a_text_edit_of_that_cell_is_a_conflict_as_in_the_body()
    {
        // One reviewer bolds the paragraph and another rewrites a word in it. In a body paragraph the merge records
        // one conflict there; the same edits in a cell must too, rather than drop the bold without a record.
        foreach (var wrap in new System.Func<string, string>[] { p => p, p => Table("", Cell(p)) })
        {
            var original = Document(wrap(Paragraph("Hello world here")));
            var reviewers = new[]
            {
                new DocxDiffReviewer { Document = Document(wrap(Paragraph("Hello world here", rPr: "<w:b/>"))), Author = "Reviewer A" },
                new DocxDiffReviewer { Document = Document(wrap(Paragraph("Hello there here"))), Author = "Reviewer B" },
            };
            var conflict = Assert.Single(DocxDiff.GetConflicts(original, reviewers));
            Assert.Equal(new[] { "Reviewer A", "Reviewer B" }, conflict.Competitors.Select(c => c.Author));
        }
    }
}
