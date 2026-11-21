using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
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

    public static IEnumerable<object[]> BookmarkAndEditCases()
    {
        foreach (var policy in Policies)
        foreach (bool controlled in new[] { false, true })
        foreach (bool editorFirst in new[] { false, true })
            yield return new object[] { policy, controlled, editorFirst };
    }

    [Theory]
    [MemberData(nameof(BookmarkAndEditCases))]
    public void A_bookmark_in_one_row_composes_with_an_edit_in_another(
        ConflictResolution policy, bool controlled, bool editorFirst)
    {
        const string bookmark = "<w:bookmarkStart w:id=\"4\" w:name=\"Berth\"/><w:r><w:t>Berth one</w:t></w:r><w:bookmarkEnd w:id=\"4\"/>";
        WmlDocument Make(bool marked, bool edited)
        {
            string first = Row("Dock", Cell(marked ? "<w:p>" + bookmark + "</w:p>" : P("Berth one")));
            string second = Row("Crane", edited ? "Hook two, spare" : "Hook two");
            return Table(first, controlled ? Control(9, second) : second);
        }

        var original = Make(false, false);
        var editor = new DocxDiffReviewer { Author = "Editor", Document = Make(false, true) };
        var marker = new DocxDiffReviewer { Author = "Marker", Document = Make(true, false) };
        var reviewers = editorFirst ? new[] { editor, marker } : new[] { marker, editor };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
        var redline = DocxCompare.Consolidate(original, reviewers, settings);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertValid(redline, accepted, rejected);
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
        Assert.Equal(Layout(Make(true, true)), Layout(accepted));
        Assert.Equal(Layout(original), Layout(rejected));
        Assert.Empty(Body(rejected).Descendants(W + "bookmarkStart"));
        Assert.Empty(Body(rejected).Descendants(W + "bookmarkEnd"));
        foreach (var doc in new[] { redline, accepted })
        {
            var body = Body(doc);
            var rows = body.Descendants(W + "tr").ToArray();
            Assert.Equal(2, rows.Length);
            var start = Assert.Single(body.Descendants(W + "bookmarkStart"));
            var end = Assert.Single(body.Descendants(W + "bookmarkEnd"));
            Assert.Same(rows[0], start.Ancestors(W + "tr").Single());
            Assert.Same(start.Ancestors(W + "p").Single(), end.Ancestors(W + "p").Single());
            Assert.Equal("Berth", (string?)start.Attribute(W + "name"));
            Assert.Equal((string?)start.Attribute(W + "id"), (string?)end.Attribute(W + "id"));
            var covered = rows[0].Descendants().SkipWhile(e => e != start).TakeWhile(e => e != end);
            Assert.Equal("Berth one", string.Concat(covered.Where(e => e.Name == W + "t").Select(e => e.Value)));
        }
    }

    [Fact]
    public void Two_reviewers_editing_a_table_with_a_content_control_row_keep_each_row_once()
    {
        // The composed table puts each row back in its control once, so both reviewers' edits compose (issue #1810).
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = Table(Row("Dock", "Berth one, north"), Control(9, Row("Crane", "Hook two"))), Author = "Reviewer A" },
            new DocxDiffReviewer { Document = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare"))), Author = "Reviewer B" },
        };
        var merged = DocxCompare.Consolidate(original, reviewers);
        Assert.Equal(Layout(Table(Row("Dock", "Berth one, north"), Control(9, Row("Crane", "Hook two, spare")))),
            Layout(RevisionProcessor.AcceptRevisions(merged)));
        Assert.Equal(Layout(original), Layout(RevisionProcessor.RejectRevisions(merged)));
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers));
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
    /// Two reviewers who make the same edit in a content-control table agree, as in a plain table (issue #1810), even
    /// when their tables differ elsewhere: <c>alike</c> makes the same edit in two separately saved copies, one with
    /// another <c>w14:textId</c>; <c>bookmark</c> also adds a bookmark in the second copy, which composing keeps.
    /// </summary>
    [Theory]
    [InlineData("alike")]
    [InlineData("bookmark")]
    public void Two_reviewers_editing_a_content_control_table_alike_agree(string shape)
    {
        const string w14 = "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var a = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare")));
        var b = Table(Row("Dock", shape == "bookmark"
                ? Cell("<w:p><w:bookmarkStart w:id=\"4\" w:name=\"Berth\"/><w:r><w:t>Berth one</w:t></w:r><w:bookmarkEnd w:id=\"4\"/></w:p>")
                : Cell(P("Berth one"))),
            Control(9, Row("Crane", Cell(P("Hook two, spare").Replace("<w:p>", $"<w:p {w14} w14:textId=\"22222222\">")))));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = a, Author = "Reviewer A" },
            new DocxDiffReviewer { Document = b, Author = "Reviewer B" },
        };
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers));
        var merged = DocxCompare.Consolidate(original, reviewers);
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
        Assert.Equal(Layout(a), Layout(accepted));
        Assert.Equal(shape == "bookmark" ? 1 : 0, Body(accepted).Descendants(W + "bookmarkStart").Count(s => s.Attribute(W + "name")?.Value == "Berth"));
        Assert.Equal(Layout(original), Layout(rejected));
        Assert.Empty(Body(rejected).Descendants(W + "bookmarkStart"));
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
    /// A control that holds no row (empty, or only a bookmark) added or removed as a row beside it is edited cannot be
    /// tracked among the rows, so the redline deletes and inserts the table whole, as for any table it cannot rebuild.
    /// </summary>
    [Theory]
    [InlineData("empty-control-removed")]
    [InlineData("empty-control-added")]
    public void A_table_whose_content_controls_cannot_be_rebuilt_is_replaced_whole(string shape)
    {
        var empty = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent/></w:sdt>";
        WmlDocument Plain(string first, bool control) => Table(Row("Dock", first), control ? empty : "", Row("Crane", "Hook two"));
        var (left, right) = shape switch
        {
            "empty-control-removed" => (Plain("Berth one", true), Plain("Berth one, north", false)),
            _ => (Plain("Berth one", false), Plain("Berth one, north", true)),
        };
        var redline = DocxCompare.Compare(left, right);
        Assert.Equal(2, Body(redline).Elements(W + "tbl").Count());
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(redline), RevisionProcessor.RejectRevisions(redline));
        Assert.Equal(Layout(right), Layout(accepted));
        Assert.Equal(Layout(left), Layout(rejected));
        AssertValid(redline, accepted, rejected);
    }

    private static WmlDocument CentredTable(params string[] children) => IrTestDocuments.FromBodyXml(
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:jc w:val=\"center\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
        string.Concat(children) + "</w:tbl>" + P("The register closes here."));

    private const string EmptyControl = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent/></w:sdt>";

    private const string BookmarkControl = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent>" +
        "<w:bookmarkStart w:id=\"3\" w:name=\"Ledger\"/><w:bookmarkEnd w:id=\"3\"/></w:sdtContent></w:sdt>";

    private static string CellLessRow(string label, string text) =>
        "<w:tr>" + Cell(P(label)) + "<w:sdt><w:sdtPr><w:id w:val=\"8\"/></w:sdtPr><w:sdtContent/></w:sdt>" + Cell(P(text)) + "</w:tr>";

    /// <summary>
    /// A change only to a table's content controls is reported and tracked (issue #1810): the row the controls change
    /// is deleted in its original controls and inserted in its revised ones, so rejecting gives the original's
    /// controls. <c>wrapped</c>/<c>unwrapped</c>: an unchanged row gains or loses a control; <c>renamed</c>: a control's
    /// alias changes; <c>placeholder-off</c>: a control stops showing its placeholder; <c>moves-to-next</c>: a row
    /// moves from one control to the next; <c>cell-wrapped</c> and <c>cell-renamed</c>: a cell gains a row-level
    /// control, or its control is renamed; <c>cell-less-added</c>: a row gains a row-level control that holds no
    /// cell.
    /// </summary>
    [Theory]
    [InlineData("wrapped")]
    [InlineData("unwrapped")]
    [InlineData("renamed")]
    [InlineData("placeholder-off")]
    [InlineData("moves-to-next")]
    [InlineData("cell-wrapped")]
    [InlineData("cell-renamed")]
    [InlineData("cell-less-added")]
    public void A_change_only_to_a_tables_content_controls_is_reported_and_tracked(string shape)
    {
        string first = Row("Dock", "Berth one");
        string item = Row("Crane", "Hook two");
        string third = Row("Barge", "Rope three");
        string cellControlled(string alias) => Row("Crane", Control(12, Cell(P("Hook two"))).Replace("Item 12", alias));
        var (left, right) = shape switch
        {
            "wrapped" => (Table(first, item), Table(first, Control(9, item))),
            "unwrapped" => (Table(first, Control(9, item)), Table(first, item)),
            "renamed" => (Table(first, Control(9, item)), Table(first, Control(9, item).Replace("Item 9", "Spare item 9"))),
            "placeholder-off" => (Table(first, Control(9, item).Replace("</w:sdtPr>", "<w:showingPlcHdr/></w:sdtPr>")), Table(first, Control(9, item))),
            "moves-to-next" => (Table(Control(9, first, item), Control(10, third)), Table(Control(9, first), Control(10, item, third))),
            "cell-wrapped" => (Table(first, item), Table(first, cellControlled("Item 12"))),
            "cell-renamed" => (Table(first, cellControlled("Item 12")), Table(first, cellControlled("Spare item 12"))),
            _ => (Table(first, item), Table(first, CellLessRow("Crane", "Hook two"))),
        };
        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        AssertRedline(left, right, DocxCompare.Compare(left, right));
    }

    /// <summary>
    /// An edit beside a change to a table's content controls stays row- and cell-precise: the edit is tracked in its
    /// row and the rows the controls change are deleted and inserted in their controls, all in one table that agrees
    /// with <see cref="DocxDiff.GetRevisions(WmlDocument, WmlDocument, DocxDiffSettings?)"/> (issue #1810).
    /// <c>renamed</c> and <c>placeholder-off</c>: the control around two rows changes as the row before it is edited;
    /// <c>item-replaced</c>: one repeating-section item replaces another, the row in it edited; <c>cell-renamed</c>: a
    /// cell's row-level control is renamed beside an edit; <c>row-leaves</c>: the revision replaces the control's last
    /// row with one outside it; <c>empty-control</c>, <c>bookmark-control</c> and <c>cell-less-control</c>: an
    /// unchanged control that holds no row or no cell sits beside the edit.
    /// </summary>
    [Theory]
    [InlineData("renamed")]
    [InlineData("placeholder-off")]
    [InlineData("item-replaced")]
    [InlineData("cell-renamed")]
    [InlineData("row-leaves")]
    [InlineData("empty-control")]
    [InlineData("bookmark-control")]
    [InlineData("cell-less-control")]
    [InlineData("cell-wrapped")]
    [InlineData("cell-less-added")]
    [InlineData("format-and-renamed")]
    [InlineData("controlled-then-plain")]
    [InlineData("empty-before-deleted-row")]
    [InlineData("empty-after-inserted-row")]
    public void An_edit_beside_a_content_control_change_stays_row_precise(string shape)
    {
        // empty-before-deleted-row: a control that holds no row sits before a row the revision deletes, and stays
        // there; empty-after-inserted-row: the revision inserts a row before such a control.

        // cell-less-control is a control (an unchanged control that holds no cell beside the edit, which main drew in
        // place). cell-wrapped: an unchanged row's cell gains a row-level control as another row is edited; cell-less-added: a
        // row gains a row-level control that holds no cell, likewise; format-and-renamed: the table is centred, nothing
        // else, as a control around a row is renamed; controlled-then-plain: a row leaves its control and the plain row
        // after it is edited, which must still pair with its own original.
        string edited = "Berth one, north";
        string pair = Control(9, Row("Crane", "Hook two"), Row("Barge", "Rope three"));
        string cellControlled(string alias) => Row("Crane", Control(12, Cell(P("Hook two"))).Replace("Item 12", alias));
        var (left, right) = shape switch
        {
            "renamed" => (Table(Row("Dock", "Berth one"), pair), Table(Row("Dock", edited), pair.Replace("Item 9", "Spare item 9"))),
            "placeholder-off" => (Table(Row("Dock", "Berth one"), pair.Replace("</w:sdtPr>", "<w:showingPlcHdr/></w:sdtPr>")),
                Table(Row("Dock", edited), pair)),
            "item-replaced" => (Table(Row("Dock", "Berth one"), Control(20, Control(21, Row("Crane", "Hook two")))),
                Table(Row("Dock", "Berth one"), Control(20, Control(22, Row("Crane", "Hook two, spare"))))),
            "cell-renamed" => (Table(Row("Dock", "Berth one"), cellControlled("Item 12")), Table(Row("Dock", edited), cellControlled("Spare item 12"))),
            "row-leaves" => (Table(Row("Dock", "Berth one"), pair),
                Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")), Row("Lighthouse", "Lamp oil"))),
            "empty-control" => (Table(Row("Dock", "Berth one"), EmptyControl, Row("Crane", "Hook two")),
                Table(Row("Dock", edited), EmptyControl, Row("Crane", "Hook two"))),
            "bookmark-control" => (Table(Row("Dock", "Berth one"), BookmarkControl, Row("Crane", "Hook two")),
                Table(Row("Dock", edited), BookmarkControl, Row("Crane", "Hook two"))),
            "cell-wrapped" => (Table(Row("Dock", "Berth one"), Row("Crane", "Hook two")), Table(Row("Dock", edited), cellControlled("Item 12"))),
            "cell-less-added" => (Table(Row("Dock", "Berth one"), Row("Crane", "Hook two")), Table(Row("Dock", edited), CellLessRow("Crane", "Hook two"))),
            "format-and-renamed" => (Table(Row("Dock", "Berth one"), pair),
                CentredTable(Row("Dock", "Berth one"), pair.Replace("Item 9", "Spare item 9"))),
            "controlled-then-plain" => (Table(Control(9, Row("Crane", "Hook two")), Row("Dock", "Berth one")),
                Table(Row("Crane", "Hook two"), Row("Dock", edited))),
            "empty-before-deleted-row" => (Table(Row("Dock", "Berth one"), Row("Crane", "Hook two"), EmptyControl, Row("Barge", "Rope three")),
                Table(Row("Dock", edited), Row("Crane", "Hook two"), EmptyControl)),
            "empty-after-inserted-row" => (Table(Row("Dock", "Berth one"), EmptyControl, Row("Crane", "Hook two")),
                Table(Row("Dock", edited), Row("Barge", "Rope three"), EmptyControl, Row("Crane", "Hook two"))),
            _ => (Table(Row("Dock", "Berth one"), CellLessRow("Crane", "Hook two")), Table(Row("Dock", edited), CellLessRow("Crane", "Hook two"))),
        };
        var redline = DocxCompare.Compare(left, right);
        AssertRedline(left, right, redline);
        if (shape == "bookmark-control")
        {
            Assert.Single(Body(RevisionProcessor.AcceptRevisions(redline)).Descendants(W + "sdt").Descendants(W + "bookmarkStart"));
            Assert.Single(Body(RevisionProcessor.RejectRevisions(redline)).Descendants(W + "sdt").Descendants(W + "bookmarkStart"));
        }
    }

    /// <summary>
    /// A change to a table's controls the rows cannot express (a control that holds no row added, removed or moved
    /// among the rows) replaces the table whole, and the comparison reports it the same way: the redline's marked text
    /// is the reported text (issue #1810).
    /// </summary>
    [Theory]
    [InlineData("added")]
    [InlineData("removed")]
    [InlineData("moved")]
    public void A_control_that_holds_no_row_changing_replaces_the_table_whole_as_reported(string shape)
    {
        string a = Row("Dock", "Berth one"), b = Row("Crane", "Hook two"), c = Row("Barge", "Rope three");
        var (left, right) = shape switch
        {
            "added" => (Table(a, b, c), Table(a, EmptyControl, b, c)),
            "removed" => (Table(a, EmptyControl, b, c), Table(a, b, c)),
            _ => (Table(a, EmptyControl, b, c), Table(a, b, EmptyControl, c)),
        };
        var redline = DocxCompare.Compare(left, right);
        Assert.Equal(2, Body(redline).Elements(W + "tbl").Count());
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(redline), RevisionProcessor.RejectRevisions(redline));
        Assert.Equal(Layout(right), Layout(accepted));
        Assert.Equal(Layout(left), Layout(rejected));
        var revisions = DocxDiff.GetRevisions(left, right);
        Assert.Equal(string.Concat(revisions.Where(r => r.Type == DocxDiffRevisionType.Inserted).Select(r => r.Text)), Marked(redline, true));
        Assert.Equal(string.Concat(revisions.Where(r => r.Type == DocxDiffRevisionType.Deleted).Select(r => r.Text)), Marked(redline, false));
        AssertValid(redline, accepted, rejected);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Table_control_plans_bind_to_fresh_sources_and_serialized_scripts(bool retain)
    {
        string MakeRow(string value) => Row("Crane", Control(12, Cell(P(value))));
        var left = Table(Row("Dock", "Berth one"), EmptyControl, Control(9, Control(10, MakeRow("Hook two"))));
        var right = Table(Row("Dock", "Berth one"), EmptyControl, Control(9, Control(10, MakeRow("Hook two, spare"))));
        var options = new IrReaderOptions { RetainSources = retain, RevisionView = RevisionView.Accept };
        var settings = new IrDiffSettings();
        var script = IrEditScriptBuilder.Build(IrReader.Read(left, options), IrReader.Read(right, options), settings);
        var diff = Assert.Single(script.Operations, op => op.TableDiff is not null).TableDiff!;
        Assert.False(diff.ReplaceWhole);
        Assert.NotNull(diff.ControlPlan);
        var restored = IrEditScriptJson.Read(IrEditScriptJson.Write(script));
        Assert.Equal(script, restored);
        foreach (var plan in new[] { script, restored })
            AssertRedline(left, right, IrMarkupRenderer.Render(plan, left, right, settings));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Composed_controls_bind_to_fresh_reviewer_sources(bool retain)
    {
        WmlDocument Make(string dock, string crane) => Table(Row("Dock", dock),
            Control(9, Row("Crane", Control(12, Cell(P(crane))))));
        var original = Make("Berth one", "Hook two");
        var editor = Make("Berth one, north", "Hook two");
        var inspector = Make("Berth one", "Hook two, spare");
        var expected = Make("Berth one, north", "Hook two, spare");
        var options = new IrReaderOptions { RetainSources = retain, RevisionView = RevisionView.Accept };
        var settings = new IrDiffSettings();
        var plan = IrCompositeMerger.Merge(IrReader.Read(original, options),
            new[] { ("Editor", IrReader.Read(editor, options)), ("Inspector", IrReader.Read(inspector, options)) },
            ConflictResolution.BaseWins, settings);
        Assert.Empty(plan.Conflicts);
        var table = Assert.Single(plan.Operations, op => op.AuthoredRows is not null);
        Assert.NotNull(table.ControlPlacement?.Rows);
        Assert.Single(table.ControlPlacement.Cells);
        var marked = IrCompositeMarkupRenderer.Render(plan, original, new[] { ("Editor", editor), ("Inspector", inspector) }, settings);
        var accepted = RevisionProcessor.AcceptRevisions(marked);
        var rejected = RevisionProcessor.RejectRevisions(marked);
        Assert.Equal(Layout(expected), Layout(accepted));
        Assert.Equal(Layout(original), Layout(rejected));
        AssertValid(marked, accepted, rejected);
    }

    private static readonly ConflictResolution[] Policies =
        { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll };

    private static DocxDiffReviewer[] Reviewers(WmlDocument a, WmlDocument b) => new[]
    {
        new DocxDiffReviewer { Document = a, Author = "Reviewer A" },
        new DocxDiffReviewer { Document = b, Author = "Reviewer B" },
    };

    /// <summary>
    /// Two reviewers who edit a table with content-control rows or cells compose as in a plain table (issue #1810):
    /// accepting gives both edits, each row once in its control, and nothing is recorded as a conflict.
    /// <c>different-rows</c>: one edits a plain row and one a row in a control; <c>row-added</c>: one adds a row to the
    /// control; <c>nested</c>: each edits a row at a different depth of nested controls; <c>cell-control</c>: one edits a
    /// cell a row-level control delivers; <c>renamed</c>: one renames the control, the other edits the plain row.
    /// </summary>
    [Theory]
    [InlineData("different-rows")]
    [InlineData("row-added")]
    [InlineData("nested")]
    [InlineData("cell-control")]
    [InlineData("renamed")]
    public void Two_reviewers_editing_a_content_control_table_compose(string shape)
    {
        string dock = Row("Dock", "Berth one"), dockEdited = Row("Dock", "Berth one, north");
        string crane = Row("Crane", "Hook two"), craneEdited = Row("Crane", "Hook two, spare");
        string barge = Row("Barge", "Rope three"), bargeEdited = Row("Barge", "Rope three, tarred");
        string cellRow(string text) => Row("Crane", Control(12, Cell(P(text))));
        var (original, a, b, expected) = shape switch
        {
            "different-rows" => (Table(dock, Control(9, crane)), Table(dockEdited, Control(9, crane)),
                Table(dock, Control(9, craneEdited)), Table(dockEdited, Control(9, craneEdited))),
            "row-added" => (Table(dock, Control(9, crane)), Table(dock, Control(9, crane, barge)),
                Table(dockEdited, Control(9, crane)), Table(dockEdited, Control(9, crane, barge))),
            "nested" => (Table(dock, Control(9, Control(10, crane), barge)), Table(dock, Control(9, Control(10, craneEdited), barge)),
                Table(dock, Control(9, Control(10, crane), bargeEdited)), Table(dock, Control(9, Control(10, craneEdited), bargeEdited))),
            "cell-control" => (Table(dock, cellRow("Hook two")), Table(dockEdited, cellRow("Hook two")),
                Table(dock, cellRow("Hook two, spare")), Table(dockEdited, cellRow("Hook two, spare"))),
            _ => (Table(dock, Control(9, crane)), Table(dock, Control(9, crane).Replace("Item 9", "Spare item 9")),
                Table(dockEdited, Control(9, crane)), Table(dockEdited, Control(9, crane).Replace("Item 9", "Spare item 9"))),
        };
        foreach (var policy in Policies)
            foreach (var reviewers in new[] { Reviewers(a, b), Reviewers(b, a) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
                Assert.Equal(Layout(expected), Layout(accepted));
                Assert.Equal(Layout(original), Layout(rejected));
                Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
                Assert.Single(Body(merged).Elements(W + "tbl"));
                AssertValid(merged, accepted, rejected);
            }
    }

    /// <summary>
    /// A reviewer who only adds a bookmark in a table with a content-control row keeps it when another reviewer edits
    /// the controlled row, in either order and under every policy (issue #1810).
    /// </summary>
    [Fact]
    public void A_reviewers_bookmark_in_a_content_control_table_another_reviewer_edits_is_kept()
    {
        string marked = Row("Dock", Cell("<w:p><w:bookmarkStart w:id=\"4\" w:name=\"Berth\"/><w:r><w:t>Berth one</w:t></w:r><w:bookmarkEnd w:id=\"4\"/></w:p>"));
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var editor = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare")));
        var marker = Table(marked, Control(9, Row("Crane", "Hook two")));
        foreach (var policy in Policies)
            foreach (var reviewers in new[] { Reviewers(editor, marker), Reviewers(marker, editor) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
                Assert.Equal(Layout(Table(marked, Control(9, Row("Crane", "Hook two, spare")))), Layout(accepted));
                Assert.Equal("Berth", Assert.Single(Body(accepted).Descendants(W + "bookmarkStart")).Attribute(W + "name")!.Value);
                Assert.Equal(Layout(original), Layout(rejected));
                Assert.Empty(Body(rejected).Descendants(W + "bookmarkStart"));
                Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
                AssertValid(merged, accepted, rejected);
            }
    }

    [Fact]
    public void A_reviewer_renaming_a_control_another_reviewer_edits_inside_is_a_conflict()
    {
        // The renamer's comparison deletes the row in the old control and inserts it in the renamed one, and the editor
        // modifies the same row: a delete-versus-modify of one row, recorded and resolved by policy as in a plain table.
        var original = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")));
        var renamer = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two")).Replace("Item 9", "Spare item 9"));
        var editor = Table(Row("Dock", "Berth one"), Control(9, Row("Crane", "Hook two, spare")));
        foreach (var policy in Policies)
            foreach (var reviewers in new[] { Reviewers(renamer, editor), Reviewers(editor, renamer) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
                Assert.NotEmpty(DocxDiff.GetConflicts(original, reviewers, settings));
                Assert.Equal(Layout(original), Layout(rejected));
                if (policy == ConflictResolution.BaseWins)
                    Assert.Equal(Layout(original), Layout(accepted));
                if (policy == ConflictResolution.FirstReviewerWins)
                    Assert.Equal(Layout(reviewers[0].Document), Layout(accepted));
                AssertValid(merged, accepted, rejected);
            }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void An_annotation_only_row_keeps_its_cell_structure_and_bookmark(bool controlled)
    {
        string MarkedRow(string cell) => "<w:tr><w:bookmarkStart w:id='1' w:name='Row'/>" + Cell(P("Crane")) +
            (controlled ? Control(12, Cell(cell)) : Cell(cell)) + "<w:bookmarkEnd w:id='1'/></w:tr>";
        string marked = "<w:p><w:bookmarkStart w:id='4' w:name='Hook'/><w:r><w:t>Hook two</w:t></w:r><w:bookmarkEnd w:id='4'/></w:p>";
        var original = Table(Row("Dock", "Berth one"), MarkedRow(P("Hook two")));
        var editor = Table(Row("Dock", "Berth one, north"), MarkedRow(P("Hook two")));
        var marker = Table(Row("Dock", "Berth one"), MarkedRow(marked));
        var expected = Table(Row("Dock", "Berth one, north"), MarkedRow(marked));
        foreach (var policy in Policies)
            foreach (var reviewers in new[] { Reviewers(editor, marker), Reviewers(marker, editor) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
                Assert.Equal(Layout(expected), Layout(accepted));
                Assert.Equal(Layout(original), Layout(rejected));
                foreach (var doc in new[] { merged, accepted })
                {
                    var row = Body(doc).Descendants(W + "tr").Last();
                    var start = Assert.Single(row.Descendants(W + "bookmarkStart"), e => (string?)e.Attribute(W + "name") == "Hook");
                    Assert.Equal("Hook two", start.Ancestors(W + "p").First().Value);
                    Assert.Equal(controlled, start.Ancestors(W + "sdt").Any());
                    Assert.Single(row.Elements(W + "bookmarkStart"), e => (string?)e.Attribute(W + "name") == "Row");
                }
                Assert.DoesNotContain(Body(rejected).Descendants(W + "bookmarkStart"), e => (string?)e.Attribute(W + "name") == "Hook");
                Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
                Assert.DoesNotContain(DocxDiff.GetConsolidatedRevisions(original, reviewers, settings), r => r.Text.Contains("Hook"));
                AssertValid(merged, accepted, rejected);
            }
    }

    /// <summary>
    /// Shapes the merge cannot compose with the controls rebuilt are recorded as a conflict, never composed with a row
    /// misplaced (issue #1810 review): <c>renamed-and-inserted</c> and <c>unwrapped-and-inserted</c>, one reviewer
    /// renames, or removes, the control around two rows as another inserts a row between them.
    /// </summary>
    [Theory]
    [InlineData("renamed-and-inserted")]
    [InlineData("unwrapped-and-inserted")]
    public void A_content_control_table_the_merge_cannot_compose_is_a_recorded_conflict(string shape)
    {
        string dock = Row("Dock", "Berth one");
        string crane = Row("Crane", "Hook two"), barge = Row("Barge", "Rope three"), mast = Row("Mast", "Sail four");
        var (original, a, b) = shape switch
        {
            "renamed-and-inserted" => (Table(dock, Control(9, crane, barge)), Table(dock, Control(9, crane, barge).Replace("Item 9", "Spare item 9")),
                Table(dock, Control(9, crane, mast, barge))),
            _ => (Table(dock, Control(9, crane, barge)), Table(dock, crane, barge), Table(dock, Control(9, crane, mast, barge))),
        };
        foreach (var policy in Policies)
            foreach (var reviewers in new[] { Reviewers(a, b), Reviewers(b, a) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(merged), RevisionProcessor.RejectRevisions(merged));
                Assert.NotEmpty(DocxDiff.GetConflicts(original, reviewers, settings));
                Assert.Equal(Layout(original), Layout(rejected));
                if (policy == ConflictResolution.BaseWins)
                    Assert.Equal(Layout(original), Layout(accepted));
                if (policy == ConflictResolution.FirstReviewerWins)
                    Assert.Equal(Layout(reviewers[0].Document), Layout(accepted));
                AssertValid(merged, accepted, rejected);
            }
    }

    [Fact]
    public void Insertions_on_opposite_sides_of_an_empty_control_keep_their_source_positions()
    {
        string first = Row("Dock", "Berth one"), last = Row("Crane", "Hook two");
        string before = Row("Barge", "Rope three"), after = Row("Mast", "Sail four");
        var original = Table(first, EmptyControl, last);
        var leftOfControl = Table(first, before, EmptyControl, last);
        var rightOfControl = Table(first, EmptyControl, after, last);
        var expected = Table(first, before, EmptyControl, after, last);
        foreach (var reviewers in new[] { Reviewers(leftOfControl, rightOfControl), Reviewers(rightOfControl, leftOfControl) })
            foreach (var policy in Policies)
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxDiff.Consolidate(original, reviewers, settings);
                Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
                Assert.Equal(Layout(expected), Layout(RevisionProcessor.AcceptRevisions(merged)));
                Assert.Equal(Layout(original), Layout(RevisionProcessor.RejectRevisions(merged)));
                var insertions = Body(merged).Descendants(W + "trPr").Elements(W + "ins").ToArray();
                Assert.Equal(2, insertions.Length);
                Assert.Equal(reviewers.Select(r => r.Author).Order(), insertions.Select(e => (string)e.Attribute(W + "author")).Order());
                AssertValid(merged);
            }
    }

    [Fact]
    public void Rows_two_reviewers_insert_into_neighbouring_controls_compose_in_either_order()
    {
        // One reviewer adds a row at the end of the first control, the other at the head of the next: both rows go in
        // at the same place among the base rows, and each goes next to its own control's rows whichever reviewer comes
        // first (issue #1810 review).
        string crane = Row("Crane", "Hook two"), barge = Row("Barge", "Rope three");
        var original = Table(Row("Dock", "Berth one"), Control(9, crane), Control(10, barge));
        var a = Table(Row("Dock", "Berth one"), Control(9, crane, Row("Mast", "Sail four")), Control(10, barge));
        var b = Table(Row("Dock", "Berth one"), Control(9, crane), Control(10, Row("Keel", "Hull five"), barge));
        var expected = Table(Row("Dock", "Berth one"), Control(9, crane, Row("Mast", "Sail four")), Control(10, Row("Keel", "Hull five"), barge));
        foreach (var policy in Policies)
            foreach (var reviewers in new[] { Reviewers(a, b), Reviewers(b, a) })
            {
                var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
                var merged = DocxCompare.Consolidate(original, reviewers, settings);
                Assert.Empty(DocxDiff.GetConflicts(original, reviewers, settings));
                Assert.Equal(Layout(expected), Layout(RevisionProcessor.AcceptRevisions(merged)));
                Assert.Equal(Layout(original), Layout(RevisionProcessor.RejectRevisions(merged)));
                AssertValid(merged);
            }
    }

    /// <summary>
    /// The redline agrees with <see cref="DocxDiff.GetRevisions(WmlDocument, WmlDocument, DocxDiffSettings?)"/> whether
    /// it draws the table in place or whole (issue #1810 review): the comparison asks the redline's own planner.
    /// <c>item-boundaries</c>: two repeating-section items whose rows next to each other both change their cell controls,
    /// drawn in place by putting each item's deleted and inserted rows together; <c>nested-cell-controls</c>: two outer
    /// controls whose rows lose their cell controls; <c>alike-merged</c> and <c>alike-split</c>: two controls with the same
    /// properties become one, or one becomes two, beside an edit; <c>alike-merged-only</c>: the same with no other change,
    /// which must still be reported.
    /// </summary>
    [Theory]
    [InlineData("item-boundaries")]
    [InlineData("nested-cell-controls")]
    [InlineData("alike-merged")]
    [InlineData("alike-split")]
    [InlineData("alike-merged-only")]
    public void The_redline_of_a_content_control_table_agrees_with_the_reported_revisions(string shape)
    {
        string CellControl(int id, string text, string alias = "") => Control(id, Cell(P(text))).Replace($"Item {id}", alias.Length > 0 ? alias : $"Item {id}");
        static string Alike(params string[] rows) => "<w:sdt><w:sdtPr><w:alias w:val=\"Entry\"/></w:sdtPr><w:sdtContent>" + string.Concat(rows) + "</w:sdtContent></w:sdt>";
        string crane = Row("Crane", "Hook two"), barge = Row("Barge", "Rope three");
        var (left, right) = shape switch
        {
            "item-boundaries" => (Table(Row("Dock", "Berth one"), Control(21, Row("Note", "one"), Row("Crane", CellControl(31, "Open"))),
                    Control(22, Row("Barge", CellControl(32, "Open")), Row("Note", "two"))),
                Table(Row("Dock", "Berth one"), Control(21, Row("Note", "one"), Row("Crane", CellControl(31, "Open", "Ticked"))),
                    Control(22, Row("Barge", CellControl(32, "Open", "Ticked")), Row("Note", "two")))),
            "nested-cell-controls" => (Table(Control(9, Row("Dock", "Berth one"), Row("Crane", CellControl(12, "Hook two"))),
                    Control(10, Row("Barge", CellControl(13, "Rope three")), Row("Mast", "Sail four"))),
                Table(Control(9, Row("Dock", "Berth one"), Row("Crane", "Hook two")), Control(10, Row("Barge", "Rope three"), Row("Mast", "Sail four")))),
            "alike-merged" => (Table(Row("Dock", "Berth one"), Alike(crane), Alike(barge)), Table(Row("Dock", "Berth one, north"), Alike(crane, barge))),
            "alike-split" => (Table(Row("Dock", "Berth one"), Alike(crane, barge)), Table(Row("Dock", "Berth one, north"), Alike(crane), Alike(barge))),
            _ => (Table(Row("Dock", "Berth one"), Alike(crane), Alike(barge)), Table(Row("Dock", "Berth one"), Alike(crane, barge))),
        };
        var redline = DocxCompare.Compare(left, right);
        var (accepted, rejected) = (RevisionProcessor.AcceptRevisions(redline), RevisionProcessor.RejectRevisions(redline));
        Assert.Equal(Layout(right), Layout(accepted));
        Assert.Equal(Layout(left), Layout(rejected));
        var revisions = DocxDiff.GetRevisions(left, right);
        Assert.NotEmpty(revisions);
        Assert.Equal(string.Concat(revisions.Where(r => r.Type == DocxDiffRevisionType.Inserted).Select(r => r.Text)), Marked(redline, true));
        Assert.Equal(string.Concat(revisions.Where(r => r.Type == DocxDiffRevisionType.Deleted).Select(r => r.Text)), Marked(redline, false));
        if (shape == "item-boundaries")
            Assert.Single(Body(redline).Elements(W + "tbl"));
        AssertValid(redline, accepted, rejected);
    }

    [Fact]
    public void The_planner_refuses_a_control_whose_rows_would_come_out_apart()
    {
        // The row diff pairs a gap's rows in order, so no comparison reaches this today. A row only the original's
        // control holds, written after a row the revision adds outside it, would split the control in two.
        var control = new IrTableControlId(-1, 0);
        var wrapped = new[] { control };
        var bare = System.Array.Empty<IrTableControlId>();
        Assert.False(IrTableControlPlanner.StayTogether(new[] { wrapped, bare, wrapped }));
        Assert.True(IrTableControlPlanner.StayTogether(new[] { wrapped, wrapped, bare }));
    }
}
