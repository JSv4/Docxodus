using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// A table row or cell that a content control delivers (<c>w:tbl/w:sdt/w:tr</c>, as a repeating section writes, or
/// <c>w:tr/w:sdt/w:tc</c>) comes out of a comparison once, inside its control, with its edits tracked, and the
/// redline's revisions agree with <see cref="DocxDiff.GetRevisions(WmlDocument, WmlDocument, DocxDiffSettings?)"/>
/// (issue #1804).
/// </summary>
public class ContentControlTableTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string P(string text, string rPr = "") =>
        $"<w:p><w:r>{(rPr.Length == 0 ? "" : $"<w:rPr>{rPr}</w:rPr>")}<w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Cell(string content) => $"<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr>{content}</w:tc>";

    private static string Row(params string[] cells) =>
        "<w:tr>" + string.Concat(cells.Select(c => c.StartsWith("<w:") ? c : Cell(P(c)))) + "</w:tr>";

    private static string Control(int id, params string[] content) =>
        $"<w:sdt><w:sdtPr><w:alias w:val=\"Item {id}\"/><w:id w:val=\"{id}\"/></w:sdtPr>" +
        $"<w:sdtEndPr><w:rPr><w:i/></w:rPr></w:sdtEndPr><w:sdtContent>{string.Concat(content)}</w:sdtContent></w:sdt>";

    private static WmlDocument Table(params string[] children) => IrTestDocuments.FromBodyXml(
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
        string.Concat(children) + "</w:tbl>" + P("The register closes here."));

    private static XElement Body(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
    }

    /// <summary>The content controls around <paramref name="e"/>, up to its table or row, outermost first: each one's id,
    /// alias and end properties, which a rebuilt copy must keep.</summary>
    private static string Controls(XElement e) => string.Concat(e.Ancestors()
        .TakeWhile(a => a.Name != W + "tbl" && a.Name != W + "tr").Where(a => a.Name == W + "sdt").Reverse()
        .Select(a => $"[{a.Element(W + "sdtPr")!.Element(W + "id")?.Attribute(W + "val")?.Value}:" +
                     $"{a.Element(W + "sdtPr")!.Element(W + "alias")?.Attribute(W + "val")?.Value}:" +
                     $"{string.Join("/", a.Element(W + "sdtEndPr")?.Descendants().Select(d => d.Name.LocalName) ?? Array.Empty<string>())}]"));

    /// <summary>
    /// Every table row in order, as its controls and its cells, each cell as its controls and its runs' text with
    /// <c>*</c> before a bold run. Accepting a redline must give the revision's layout and rejecting it the
    /// original's: a row written twice, or a control lost, added or left empty, changes it.
    /// </summary>
    private static string[] Layout(WmlDocument doc)
    {
        var body = Body(doc);
        var rows = body.Descendants(W + "tr").Select(tr => Controls(tr) + "{" + string.Join("|", tr.Descendants(W + "tc")
            .Select(tc => Controls(tc) + string.Concat(tc.Descendants(W + "r").Select(r =>
                (r.Element(W + "rPr")?.Element(W + "b") is null ? "" : "*") + string.Concat(r.Elements(W + "t").Select(t => t.Value)))))) + "}");
        // Each table's own children in order, so a control that moves among the rows shows too.
        var tables = body.Descendants(W + "tbl").Select(t => "tbl:" + string.Join(",", t.Elements()
            .Where(e => e.Name == W + "tr" || e.Name == W + "sdt").Select(e => e.Name == W + "tr" ? "tr" : Controls(e.Elements().First()))));
        return rows.Concat(tables).Append($"controls:{body.Descendants(W + "sdt").Count()}").ToArray();
    }

    /// <summary>The text the redline marks inserted or deleted, in document order.</summary>
    private static string Marked(WmlDocument redline, bool inserted) => string.Concat(Body(redline)
        .Descendants(W + (inserted ? "ins" : "del")).Descendants(W + (inserted ? "t" : "delText")).Select(t => t.Value));

    private static void AssertValid(params WmlDocument[] docs)
    {
        foreach (var doc in docs)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
        }
    }

    /// <summary>
    /// Accept gives the revision and reject the original, row by row and control by control; the redline is one
    /// table whose marked text is the revisions' text; every view is schema-valid.
    /// </summary>
    private static void AssertRedline(WmlDocument left, WmlDocument right, WmlDocument redline)
    {
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(Layout(right), Layout(accepted));
        Assert.Equal(Layout(left), Layout(rejected));
        Assert.Single(Body(redline).Elements(W + "tbl"));
        var revisions = DocxDiff.GetRevisions(left, right);
        Assert.Equal(string.Concat(revisions.Where(r => r.Type == DocxDiffRevisionType.Inserted).Select(r => r.Text)), Marked(redline, true));
        Assert.Equal(string.Concat(revisions.Where(r => r.Type == DocxDiffRevisionType.Deleted).Select(r => r.Text)), Marked(redline, false));
        AssertValid(redline, accepted, rejected);
    }

    /// <summary>
    /// Rows a table-level content control delivers. <c>edit-beside</c> edits a plain row and <c>edit-inside</c> a
    /// row in the control; <c>row-added</c> adds a row to the control and <c>row-removed</c> removes one;
    /// <c>control-added</c> adds a control with two rows and <c>control-removed</c> removes it; <c>nested</c> edits
    /// a row in a control inside a control; <c>two-controls</c> edits a row between two controls and one inside
    /// the second.
    /// </summary>
    [Theory]
    [InlineData("edit-beside")]
    [InlineData("edit-inside")]
    [InlineData("row-added")]
    [InlineData("row-removed")]
    [InlineData("control-added")]
    [InlineData("control-removed")]
    [InlineData("nested")]
    [InlineData("two-controls")]
    [InlineData("section-added")]
    [InlineData("section-removed")]
    public void A_row_in_a_content_control_comes_out_once_inside_it(string shape)
    {
        string first = Row("Dock", "Berth one");
        string item = Row("Crane", "Hook two");
        (WmlDocument Left, WmlDocument Right) pair = shape switch
        {
            "edit-beside" => (Table(first, Control(9, item)), Table(Row("Dock", "Berth one, north"), Control(9, item))),
            "edit-inside" => (Table(first, Control(9, item)), Table(first, Control(9, Row("Crane", "Hook two, spare")))),
            "row-added" => (Table(first, Control(9, item)), Table(first, Control(9, item, Row("Barge", "Rope three")))),
            "row-removed" => (Table(first, Control(9, item, Row("Barge", "Rope three"))), Table(first, Control(9, item))),
            "control-added" => (Table(first), Table(first, Control(9, item, Row("Barge", "Rope three")))),
            "control-removed" => (Table(first, Control(9, item, Row("Barge", "Rope three"))), Table(first)),
            "nested" => (Table(first, Control(9, Control(10, item), Row("Barge", "Rope three"))),
                Table(first, Control(9, Control(10, Row("Crane", "Hook two, spare")), Row("Barge", "Rope three")))),
            "section-added" => (Table(first), Table(first, Control(9, Control(10, item), Control(11, Row("Barge", "Rope three"))))),
            "section-removed" => (Table(first, Control(9, Control(10, item), Control(11, Row("Barge", "Rope three")))), Table(first)),
            _ => (Table(first, Control(9, item), Row("Mast", "Sail four"), Control(11, Row("Keel", "Hull five"))),
                Table(first, Control(9, item), Row("Mast", "Sail four, patched"), Control(11, Row("Keel", "Hull five, tarred")))),
        };
        AssertRedline(pair.Left, pair.Right, DocxCompare.Compare(pair.Left, pair.Right));
    }

    [Fact]
    public void A_change_the_comparison_does_not_report_leaves_a_content_control_row_unmarked()
    {
        // The only difference is a run's language inside a block content control in a cell. A block control's hash
        // covers its content, so the table is compared row by row, but nothing is reported or marked.
        string Revision(string rPr) => Row("Dock", Cell(Control(7, P("Berth one", rPr))));
        var left = Table(Revision(""), Control(9, Row("Crane", "Hook two")));
        var right = Table(Revision("<w:lang w:val=\"fr-FR\"/>"), Control(9, Row("Crane", "Hook two")));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
        var redline = DocxCompare.Compare(left, right);
        AssertRedline(left, right, redline);
        Assert.DoesNotContain(Body(redline).Descendants(), e => e.Name.LocalName.EndsWith("Change"));
    }

    /// <summary>
    /// Cells a row-level content control delivers. <c>edit-same-row</c> edits another cell of the row, which used to
    /// replace the whole table; <c>edit-inside</c> edits the cell in the control.
    /// </summary>
    [Theory]
    [InlineData("edit-same-row")]
    [InlineData("edit-inside")]
    public void A_cell_in_a_content_control_comes_out_once_inside_it(string shape)
    {
        var left = Table(Row("Dock", "Berth one"), Row("Crane", Control(12, Cell(P("Hook two")))));
        var right = shape == "edit-same-row"
            ? Table(Row("Dock", "Berth one"), Row("Crane arm", Control(12, Cell(P("Hook two")))))
            : Table(Row("Dock", "Berth one"), Row("Crane", Control(12, Cell(P("Hook two, spare")))));
        AssertRedline(left, right, DocxCompare.Compare(left, right));
    }

    /// <summary>
    /// A formatting change in a content-control row or cell, or beside one in a table whose text is unchanged, is
    /// tracked as in any other cell (issue #1797 left these untracked until the controls could be rebuilt).
    /// </summary>
    [Theory]
    [InlineData("in-row")]
    [InlineData("beside-row")]
    [InlineData("in-cell")]
    public void A_format_change_in_or_beside_a_content_control_is_tracked(string shape)
    {
        WmlDocument Revision(string bold) => shape switch
        {
            "in-row" => Table(Row("Dock", "Berth one"), Control(9, Row("Crane", Cell(P("Hook two", bold))))),
            "beside-row" => Table(Row("Dock", Cell(P("Berth one", bold))), Control(9, Row("Crane", "Hook two"))),
            _ => Table(Row("Dock", "Berth one"), Row("Crane", Control(12, Cell(P("Hook two", bold))))),
        };
        var (left, right) = (Revision(""), Revision("<w:b/>"));
        var redline = DocxCompare.Compare(left, right);
        AssertRedline(left, right, redline);
        Assert.Single(Body(redline).Descendants(W + "rPrChange"));
        Assert.Equal(DocxDiffRevisionType.FormatChanged, Assert.Single(DocxDiff.GetRevisions(left, right)).Type);
    }

    [Fact]
    public void Consolidating_one_reviewers_edit_keeps_each_content_control_row_once()
    {
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var reviewed = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare"), Row("Barge", "Rope three")));
        var merged = DocxCompare.Consolidate(original, new[] { new DocxDiffReviewer { Document = reviewed, Author = "Reviewer A" } });
        Assert.Equal(Layout(reviewed), Layout(RevisionProcessor.AcceptRevisions(merged)));
        Assert.Equal(Layout(original), Layout(RevisionProcessor.RejectRevisions(merged)));
    }

    [Fact]
    public void Two_reviewers_editing_a_table_with_a_content_control_row_keep_each_row_once()
    {
        // Composing two reviewers' edits in such a table would also write the control's rows twice, so the merge
        // records the table as a conflict, as it does for any table it cannot compose, and keeps the original's.
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = Table(Row("Dock", "Berth one, north"), Control(9, Row("Crane", "Hook two"))), Author = "Reviewer A" },
            new DocxDiffReviewer { Document = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare"))), Author = "Reviewer B" },
        };
        var merged = DocxCompare.Consolidate(original, reviewers);
        Assert.Equal(Layout(original), Layout(RevisionProcessor.AcceptRevisions(merged)));
        Assert.Equal(Layout(original), Layout(RevisionProcessor.RejectRevisions(merged)));
        Assert.Single(DocxDiff.GetConflicts(original, reviewers));
    }

    [Fact]
    public void A_row_property_change_in_a_content_control_row_is_tracked()
    {
        // Rows pair as the reader lists them, so a row a control delivers gets its row-property history too.
        WmlDocument Revision(string trPr) => Table(Row("Dock", "Berth one"),
            Control(9, Row("Crane", "Hook two").Replace("<w:tr>", "<w:tr>" + trPr)));
        var (left, right) = (Revision(""), Revision("<w:trPr><w:cantSplit/></w:trPr>"));
        var redline = DocxCompare.Compare(left, right);
        Assert.Single(Body(redline).Descendants(W + "sdt").Descendants(W + "trPrChange"));
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(redline), RevisionProcessor.RejectRevisions(redline));
        Assert.Empty(Body(rejected).Descendants(W + "cantSplit"));
        Assert.Single(Body(accepted).Descendants(W + "cantSplit"));
        Assert.Equal(Layout(left), Layout(rejected));
        AssertValid(redline, accepted, rejected);
    }

    [Fact]
    public void A_row_moved_inside_a_content_control_keeps_its_move_ranges_inside_it()
    {
        static string R(string a, string b) => Row(a, b);
        var left = Table(Row("Dock", "Berth one"), Control(9, R("Crane", "Hook two"), R("Barge", "Rope three"),
            R("Mast", "Sail four"), R("Keel", "Hull five")));
        var right = Table(Row("Dock", "Berth one"), Control(9, R("Barge", "Rope three"), R("Mast", "Sail four"),
            R("Keel", "Hull five"), R("Crane", "Hook two")));
        var redline = DocxCompare.Compare(left, right);
        Assert.Equal(Layout(right), Layout(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(Layout(left), Layout(RevisionProcessor.RejectRevisions(redline)));
        var markers = Body(redline).Descendants().Where(e => e.Name.LocalName.StartsWith("move") && e.Name.LocalName.Contains("Range")).ToList();
        foreach (var start in markers.Where(e => e.Name.LocalName.EndsWith("Start")))
        {
            var end = Assert.Single(markers, e => e.Name.LocalName == start.Name.LocalName.Replace("Start", "End") &&
                (string?)e.Attribute(W + "id") == (string?)start.Attribute(W + "id"));
            Assert.Same(start.Parent, end.Parent);
        }
        AssertValid(redline);
    }

    /// <summary>
    /// Two reviewers who edit a content-control table are recorded as a conflict, as for any table the merge cannot
    /// compose, even when their tables look alike: <c>alike</c> makes the same edit in two separately saved copies, one
    /// with another <c>w14:textId</c>; <c>bookmark</c> also adds a bookmark in the second copy, which a shortcut that
    /// kept one reviewer's table would drop; <c>bookmark-only</c> has the second reviewer only add that bookmark, which the
    /// merge would otherwise hand to the one reviewer who edited the table and drop. Nothing is lost: the base is kept
    /// and the conflict names both reviewers.
    /// </summary>
    [Theory]
    [InlineData("alike")]
    [InlineData("bookmark")]
    [InlineData("bookmark-only")]
    public void Two_reviewers_editing_a_content_control_table_alike_record_a_conflict(string shape)
    {
        const string w14 = "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var a = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare")));
        var b = Table(Row("Dock", shape is "bookmark" or "bookmark-only"
                ? Cell("<w:p><w:bookmarkStart w:id=\"4\" w:name=\"Berth\"/><w:r><w:t>Berth one</w:t></w:r><w:bookmarkEnd w:id=\"4\"/></w:p>")
                : Cell(P("Berth one"))),
            Control(9, Row("Crane", Cell(P(shape == "bookmark-only" ? "Hook two" : "Hook two, spare")
                .Replace("<w:p>", $"<w:p {w14} w14:textId=\"22222222\">")))));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = a, Author = "Reviewer A" },
            new DocxDiffReviewer { Document = b, Author = "Reviewer B" },
        };
        var conflict = Assert.Single(DocxDiff.GetConflicts(original, reviewers));
        Assert.Equal(new[] { "Reviewer A", "Reviewer B" }, conflict.Competitors.Select(c => c.Author));
        var merged = DocxCompare.Consolidate(original, reviewers);
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
        Assert.Equal(Layout(original), Layout(accepted));
        Assert.Equal(Layout(original), Layout(rejected));
        AssertValid(merged, accepted, rejected);
    }

    [Fact]
    public void A_bookmark_only_reviewer_listed_first_wins_a_content_control_table_under_first_reviewer_wins()
    {
        // The reviewer who only adds a bookmark competes in the conflict in list order, so FirstReviewerWins keeps
        // that reviewer's table, bookmark and all, rather than the later reviewer's edit.
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var marked = Table(Row("Dock", Cell("<w:p><w:bookmarkStart w:id=\"4\" w:name=\"Berth\"/><w:r><w:t>Berth one</w:t></w:r><w:bookmarkEnd w:id=\"4\"/></w:p>")),
            Control(9, Row("Crane", "Hook two")));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = marked, Author = "Reviewer B" },
            new DocxDiffReviewer { Document = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare"))), Author = "Reviewer A" },
        };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.FirstReviewerWins };
        var conflict = Assert.Single(DocxDiff.GetConflicts(original, reviewers, settings));
        Assert.Equal(new[] { "Reviewer B", "Reviewer A" }, conflict.Competitors.Select(c => c.Author));
        var merged = DocxCompare.Consolidate(original, reviewers, settings);
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
        Assert.Equal(Layout(original), Layout(accepted));
        Assert.Single(Body(accepted).Descendants(W + "bookmarkStart"), b => (string?)b.Attribute(W + "name") == "Berth");
        Assert.Equal(Layout(original), Layout(rejected));
        AssertValid(merged, accepted, rejected);
    }

    [Fact]
    public void Consolidating_a_row_with_a_control_that_holds_no_cell_keeps_the_control_in_place()
    {
        // No row or cell is delivered by a control here, so two reviewers' edits compose; the composed row is built
        // from the base row's shell and keeps the empty control between the same cells.
        const string empty = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent/></w:sdt>";
        WmlDocument Make(string berth, string hook) => Table(
            "<w:tr>" + Cell(P("Dock")) + empty + Cell(P(berth)) + "</w:tr>", Row("Crane", hook));
        var original = Make("Berth one", "Hook two");
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = Make("Berth one, north", "Hook two"), Author = "Reviewer A" },
            new DocxDiffReviewer { Document = Make("Berth one", "Hook two, spare"), Author = "Reviewer B" },
        };
        var merged = DocxCompare.Consolidate(original, reviewers);
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
        static string[] FirstRow(WmlDocument doc) => Body(doc).Descendants(W + "tr").First().Elements()
            .Select(e => e.Name.LocalName + ":" + string.Concat(e.Descendants(W + "t").Select(t => t.Value))).ToArray();
        Assert.Equal(FirstRow(original), FirstRow(rejected));
        Assert.Equal(FirstRow(Make("Berth one, north", "Hook two, spare")), FirstRow(accepted));
        Assert.Equal(Layout(Make("Berth one, north", "Hook two, spare")), Layout(accepted));
        Assert.Equal(Layout(original), Layout(rejected));
        AssertValid(merged, accepted, rejected);
    }

    /// <summary>
    /// Controls the redline cannot rebuild exactly make it delete and insert the table whole, as for any table it
    /// cannot rebuild. <c>renamed</c>: the control around a row is renamed as the row beside it is edited, and a
    /// control's properties cannot be tracked. <c>row-leaves</c>: the revision replaces the control's last row with a
    /// row outside it, and the comparison pairs the two rows.
    /// </summary>
    [Theory]
    [InlineData("renamed")]
    [InlineData("row-leaves")]
    [InlineData("empty-control")]
    [InlineData("bookmark-control")]
    [InlineData("empty-control-removed")]
    [InlineData("empty-control-added")]
    public void A_table_whose_content_controls_cannot_be_rebuilt_is_replaced_whole(string shape)
    {
        // A control that holds no row (empty, or only a bookmark) cannot go back among the rows from the row ops, and
        // copying it from the revision would move it to the head of the table, so it too keeps the table whole.
        var empty = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent/></w:sdt>";
        var bookmark = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id=\"3\" w:name=\"Ledger\"/><w:bookmarkEnd w:id=\"3\"/></w:sdtContent></w:sdt>";
        string inner = shape == "bookmark-control" ? bookmark : empty;
        WmlDocument Plain(string first, bool control) => Table(Row("Dock", first), control ? inner : "", Row("Crane", "Hook two"));
        var (left, right) = shape switch
        {
            "renamed" => (Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two"), Row("Barge", "Rope three"))),
                Table(Row("Dock", "Berth one, north"), Control(9, Row("Crane", "Hook two"), Row("Barge", "Rope three")).Replace("Item 9", "Spare item 9"))),
            "row-leaves" => (Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two"), Row("Barge", "Rope three"))),
                Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")), Row("Lighthouse", "Lamp oil"))),
            "empty-control-removed" => (Plain("Berth one", true), Plain("Berth one, north", false)),
            "empty-control-added" => (Plain("Berth one", false), Plain("Berth one, north", true)),
            _ => (Plain("Berth one", true), Plain("Berth one, north", true)),
        };
        var redline = DocxCompare.Compare(left, right);
        Assert.Equal(2, Body(redline).Elements(W + "tbl").Count());
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(redline), RevisionProcessor.RejectRevisions(redline));
        Assert.Equal(Layout(right), Layout(accepted));
        Assert.Equal(Layout(left), Layout(rejected));
        AssertValid(redline, accepted, rejected);
    }

    [Fact]
    public void The_planner_refuses_a_control_whose_rows_would_come_out_apart()
    {
        // The row diff pairs a gap's rows in order, so no comparison reaches this today. A row only the original's
        // control holds, written after a row the revision adds outside it, would split the control in two.
        static XElement Parse(params string[] children) => Body(Table(children)).Element(W + "tbl")!;
        var left = Parse(Control(9, Row("Crane", "Hook two"), Row("Barge", "Rope three")));
        var right = Parse(Control(9, Row("Crane", "Hook two")), Row("Lighthouse", "Lamp oil"));
        var (leftRows, rightRows) = (left.Descendants(W + "tr").ToList(), right.Descendants(W + "tr").ToList());
        var apart = new List<(XElement?, XElement?)> { (leftRows[0], rightRows[0]), (null, rightRows[1]), (leftRows[1], null) };
        Assert.Null(Docxodus.Ir.Diff.IrMarkupRenderer.PlanControls(apart, left, right, W + "tr"));
        var adjacent = new List<(XElement?, XElement?)> { (leftRows[0], rightRows[0]), (leftRows[1], null), (null, rightRows[1]) };
        Assert.NotNull(Docxodus.Ir.Diff.IrMarkupRenderer.PlanControls(adjacent, left, right, W + "tr"));
    }
}
