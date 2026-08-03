#nullable enable

using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace OxPt;

/// <summary>
/// Word Compare's "Tables" comparison option (<see cref="DocxDiffSettings.CompareTables"/>): whether a
/// table's CONTENT and STRUCTURE — rows, cells, per-cell text — are compared. Off, a changed table rides
/// through verbatim from the right document with no revision.
/// </summary>
public class DocxDiffTablesOptionTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static readonly DocxDiffSettings Off = new() { CompareTables = false };

    private static WmlDocument Doc(string bodyXml)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyXml}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("tables.docx", ms.ToArray());
    }

    private static string Row(params string[] cells) =>
        "<w:tr>" + string.Concat(cells.Select(c =>
            "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr>" +
            $"<w:p><w:r><w:t xml:space=\"preserve\">{c}</w:t></w:r></w:p></w:tc>")) + "</w:tr>";

    private static string Table(int columns, params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2500\"/>", columns)) +
        "</w:tblGrid>" + string.Concat(rows) + "</w:tbl>";

    /// <summary>The table between two unique anchoring paragraphs, so the pair aligns as Modified.</summary>
    private static WmlDocument WithTable(string tableXml) => Doc(
        "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" + tableXml +
        "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

    /// <summary>A one-cell table nested inside a one-cell table.</summary>
    private static WmlDocument NestedTable(string innerText) => WithTable(
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2500\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr>" +
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"1000\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/></w:tcPr>" +
        $"<w:p><w:r><w:t xml:space=\"preserve\">{innerText}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "<w:p/></w:tc></w:tr></w:tbl>");

    /// <summary>The table lives in a HEADER story; the body is identical on both sides.</summary>
    private static WmlDocument HeaderTable(string cellText)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var header = main.AddNewPart<HeaderPart>("rH1");
            using (var hs = header.GetStream(FileMode.Create, FileAccess.Write))
            using (var hw = new StreamWriter(hs, new UTF8Encoding(false)))
                hw.Write($"<w:hdr xmlns:w=\"{W}\">{Table(1, Row(cellText))}<w:p/></w:hdr>");

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rH1\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("header-table.docx", ms.ToArray());
    }

    /// <summary>Content-equal on both sides; only the column/cell WIDTH differs.</summary>
    private static WmlDocument ShellOnly(string width) => WithTable(
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        $"<w:tblGrid><w:gridCol w:w=\"{width}\"/></w:tblGrid>" +
        $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr>" +
        "<w:p><w:r><w:t>a1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");

    /// <summary>The table lives in a FOOTNOTE definition; the body is identical on both sides.</summary>
    private static WmlDocument FootnoteTable(string cellText)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var notes = main.AddNewPart<FootnotesPart>("rFn1");
            using (var fs = notes.GetStream(FileMode.Create, FileAccess.Write))
            using (var fw = new StreamWriter(fs, new UTF8Encoding(false)))
            {
                fw.Write($"<w:footnotes xmlns:w=\"{W}\">" +
                    "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                    "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\">" +
                    "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                    $"<w:footnote w:id=\"1\">{Table(1, Row(cellText))}<w:p/></w:footnote></w:footnotes>");
            }

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{Rel}\"><w:body>" +
                "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Cited here.</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("footnote-table.docx", ms.ToArray());
    }

    private static string NotesXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.FootnotesPart?.Footnotes.OuterXml ?? string.Empty;
    }

    /// <summary>The same table sits either after the intro or at the end — a relocation.</summary>
    private static WmlDocument MovableTable(bool atEnd)
    {
        var table = Table(1, Row("movable cell"));
        const string intro = "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>";
        const string middle = "<w:p><w:r><w:t>Beta unique middle paragraph.</w:t></w:r></w:p>";
        const string closing = "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>";
        return Doc(atEnd ? intro + middle + closing + table : intro + table + middle + closing);
    }

    /// <summary>
    /// A numbered list in the body (numId 1) plus a SEPARATE instance (numId 2) inside the table, over one
    /// abstract definition — the shape whose list identity a positional row zip would mis-harvest.
    /// </summary>
    private static WmlDocument NumberedTable(bool extraRow)
    {
        string ListParagraph(int numId, string text) =>
            $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"{numId}\"/></w:numPr></w:pPr>" +
            $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";
        string ListRow(int numId, string text) =>
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr>" +
            ListParagraph(numId, text) + "</w:tc></w:tr>";

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var numbering = main.AddNewPart<NumberingDefinitionsPart>("rNum");
            using (var ns = numbering.GetStream(FileMode.Create, FileAccess.Write))
            using (var nw = new StreamWriter(ns, new UTF8Encoding(false)))
            {
                nw.Write($"<w:numbering xmlns:w=\"{W}\">" +
                    "<w:abstractNum w:abstractNumId=\"0\"><w:multiLevelType w:val=\"singleLevel\"/>" +
                    "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>" +
                    "<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/></w:lvl></w:abstractNum>" +
                    "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>" +
                    "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>");
            }

            // extraRow marks the RIGHT side: it also switches the table to its own list instance, so a
            // false positional pairing has a wrong answer available to produce.
            int cellNumId = extraRow ? 2 : 1;
            var table = "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"2500\"/></w:tblGrid>" +
                (extraRow ? ListRow(cellNumId, "brand new extra row") : "") +
                ListRow(cellNumId, "cell list item") + "</w:tbl>";

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                ListParagraph(1, "Alpha unique body item.") + table +
                ListParagraph(1, "Omega unique body item.") +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("numbered-table.docx", ms.ToArray());
    }

    private static void AssertValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(wDoc).Select(e => e.Description));
    }

    private static string BodyXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.Document.Body!.OuterXml;
    }

    private static string HeaderXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return string.Concat(wDoc.MainDocumentPart!.HeaderParts.Select(p => p.Header.OuterXml));
    }

    /// <summary>Body paragraph texts, so a round trip is compared at the text level.</summary>
    private static string[] Paragraphs(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>().Select(p => p.InnerText).ToArray();
    }

    private static string[] Accepted(WmlDocument compared) =>
        Paragraphs(Docxodus.Internal.DocxDiffOps.AcceptRevisions(compared.DocumentByteArray));

    private static string[] Rejected(WmlDocument compared) =>
        Paragraphs(Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray));

    // --- default ------------------------------------------------------------------------------------

    [Fact]
    public void CompareTables_defaults_to_true()
    {
        Assert.True(new DocxDiffSettings().CompareTables);
    }

    // --- default (on): table content and structure are compared -------------------------------------

    [Fact]
    public void Default_reports_a_cell_text_change_inside_the_cell()
    {
        var left = WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2")));
        var right = WithTable(Table(2, Row("a1", "CHANGED"), Row("a2", "b2")));

        Assert.Equal(
            new[] { "Deleted=b1", "Inserted=CHANGED" },
            DocxDiff.GetRevisions(left, right).Select(r => $"{r.Type}={r.Text}").ToArray());
        Assert.Contains("rowOps", DocxDiff.GetEditScriptJson(left, right));
    }

    [Fact]
    public void Default_reports_an_added_row()
    {
        var revisions = DocxDiff.GetRevisions(
            WithTable(Table(2, Row("a1", "b1"))),
            WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2"))));

        Assert.Contains("a2b2", revisions.Select(r => r.Text));
    }

    [Fact]
    public void Default_reports_a_removed_row()
    {
        var revisions = DocxDiff.GetRevisions(
            WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2"))),
            WithTable(Table(2, Row("a1", "b1"))));

        Assert.Contains("a2b2", revisions.Select(r => r.Text));
    }

    [Fact]
    public void Default_reports_an_added_column_with_native_cellIns_markup()
    {
        var compared = DocxDiff.Compare(
            WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2"))),
            WithTable(Table(3, Row("a1", "b1", "c1"), Row("a2", "b2", "c2"))));

        AssertValid(compared);
        Assert.Contains("cellIns", BodyXml(compared));
    }

    // --- off: the table stops being compared --------------------------------------------------------

    [Theory]
    [InlineData("cell text")]
    [InlineData("row added")]
    [InlineData("row removed")]
    [InlineData("column added")]
    public void CompareTables_false_silences_every_table_change(string shape)
    {
        var (left, right) = TablePair(shape);

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.DoesNotContain("<w:ins ", body);
        Assert.DoesNotContain("<w:del ", body);
        Assert.DoesNotContain("cellIns", body);
    }

    [Theory]
    [InlineData("cell text")]
    [InlineData("row added")]
    [InlineData("row removed")]
    [InlineData("column added")]
    public void CompareTables_false_carries_the_right_table_through_untracked(string shape)
    {
        // An uncompared difference cannot be reversible. The output is the RIGHT table verbatim — the same
        // emit an unchanged table takes — so accept is still exact and reject keeps the right table.
        var (left, right) = TablePair(shape);
        var compared = DocxDiff.Compare(left, right, Off);

        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(right.DocumentByteArray), Rejected(compared));
        Assert.NotEqual(Paragraphs(left.DocumentByteArray), Rejected(compared));
    }

    private static (WmlDocument Left, WmlDocument Right) TablePair(string shape) => shape switch
    {
        "cell text" => (WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2"))),
                        WithTable(Table(2, Row("a1", "CHANGED"), Row("a2", "b2")))),
        "row added" => (WithTable(Table(2, Row("a1", "b1"))),
                        WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2")))),
        "row removed" => (WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2"))),
                          WithTable(Table(2, Row("a1", "b1")))),
        "column added" => (WithTable(Table(2, Row("a1", "b1"), Row("a2", "b2"))),
                           WithTable(Table(3, Row("a1", "b1", "c1"), Row("a2", "b2", "c2")))),
        _ => throw new System.ArgumentOutOfRangeException(nameof(shape), shape, null),
    };

    [Fact]
    public void CompareTables_false_drops_the_nested_row_cell_ops_from_the_edit_script()
    {
        var left = WithTable(Table(2, Row("a1", "b1")));
        var right = WithTable(Table(2, Row("a1", "CHANGED")));

        Assert.Contains("rowOps", DocxDiff.GetEditScriptJson(left, right));
        Assert.DoesNotContain("rowOps", DocxDiff.GetEditScriptJson(left, right, Off));
    }

    [Fact]
    public void CompareTables_false_reaches_a_nested_table()
    {
        var left = NestedTable("inner one");
        var right = NestedTable("inner two");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Contains("inner two", BodyXml(compared));
        Assert.DoesNotContain("inner one", BodyXml(compared));
    }

    [Fact]
    public void CompareTables_false_reaches_a_table_in_a_header_story()
    {
        var left = HeaderTable("hdr one");
        var right = HeaderTable("hdr two");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.DoesNotContain("<w:del ", HeaderXml(compared));
    }

    [Fact]
    public void CompareTables_false_still_tracks_a_shell_only_change_reversibly()
    {
        // A column width is a FORMATTING change, owned by TrackBlockFormatChanges — unchecking "Tables" must
        // not silence part of "Formatting". Such a pair reaches this gate at all only because a cell's shell
        // digest feeds its table's content identity (so the aligner says Modified, not FormatOnly); the gate
        // tests content equality and routes it to FormatOnlyBlock, which keeps it fully reversible.
        var left = ShellOnly("2500");
        var right = ShellOnly("4000");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Contains("tblGridChange", BodyXml(compared));

        // Reversible at the property level, and identical to what the box ON produces.
        Assert.Equal(BodyXml(DocxDiff.Compare(left, right)), BodyXml(compared));
        Assert.Contains("4000", BodyXml(new WmlDocument("a.docx",
            Docxodus.Internal.DocxDiffOps.AcceptRevisions(compared.DocumentByteArray))));
        var rejected = BodyXml(new WmlDocument("r.docx",
            Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
        Assert.Contains("2500", rejected);
        Assert.DoesNotContain("4000", rejected);
    }

    [Fact]
    public void CompareTables_false_keeps_the_LEFT_table_in_a_header_story()
    {
        // Per-scope surviving side, exactly as for CompareFields: a story part is only rebuilt when its story
        // produced ops, so with none produced the LEFT part rides through and it is REJECT that is exact here.
        var left = HeaderTable("hdr one");
        var right = HeaderTable("hdr two");
        var compared = DocxDiff.Compare(left, right, Off);

        AssertValid(compared);
        Assert.Contains("hdr one", HeaderXml(compared));
        Assert.DoesNotContain("hdr two", HeaderXml(compared));
        Assert.Contains("hdr one", HeaderXml(new WmlDocument("r.docx",
            Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray))));
    }

    [Fact]
    public void CompareTables_false_keeps_the_LEFT_table_in_a_footnote()
    {
        var left = FootnoteTable("fn one");
        var right = FootnoteTable("fn two");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Contains("fn one", NotesXml(compared));
        Assert.DoesNotContain("fn two", NotesXml(compared));
    }

    [Fact]
    public void CompareTables_false_still_reports_a_moved_table()
    {
        // The gate covers the in-place Modified alignment only. A relocated table is a block-level move, so it
        // survives the toggle — reported as a Moved pair and fully reversible.
        var left = MovableTable(atEnd: false);
        var right = MovableTable(atEnd: true);

        var revisions = DocxDiff.GetRevisions(left, right, Off);
        Assert.NotEmpty(revisions);
        Assert.All(revisions, r => Assert.Equal(DocxDiffRevisionType.Moved, r.Type));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(compared));
    }

    [Fact]
    public void CompareTables_false_gives_an_uncompared_table_a_fresh_list_import_with_no_revision_markup()
    {
        // The Uncompared flag's reason for existing. CollectAlignedNumIdPairs harvests list-identity evidence
        // from an EqualBlock pair by positionally zipping rows/cells/blocks — valid only for a content-ALIGNED
        // pair — so an uncompared table must contribute none, and the numId remap it then needs must not be
        // stamped as a w:pPrChange (or Compare would show formatting revisions on a table GetRevisions calls
        // unchanged).
        var left = NumberedTable(extraRow: false);
        var right = NumberedTable(extraRow: true);

        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.DoesNotContain("pPrChange", body);
        Assert.DoesNotContain("<w:ins ", body);
        Assert.DoesNotContain("<w:del ", body);

        // The right table's list must NOT be rebound onto the body's list instance: harvesting evidence from
        // this pair would zip right-row-0 against left-row-0 and conclude the table's list is the body's,
        // making the table continue the body's counter instead of restarting as it does in the right document.
        Assert.NotEqual(NumIdNear(body, "Alpha unique body item."), NumIdNear(body, "cell list item"));
    }

    /// <summary>The numId governing the paragraph that contains <paramref name="text"/>.</summary>
    private static string NumIdNear(string body, string text)
    {
        int at = body.IndexOf(text, System.StringComparison.Ordinal);
        Assert.True(at >= 0, $"'{text}' not found in the compared body");
        const string marker = "<w:numId w:val=\"";
        int start = body.LastIndexOf(marker, at, System.StringComparison.Ordinal);
        Assert.True(start >= 0, $"no w:numId precedes '{text}'");
        start += marker.Length;
        return body.Substring(start, body.IndexOf('"', start) - start);
    }

    [Fact]
    public void CompareTables_false_satisfies_the_edit_script_verifier()
    {
        // The verifier asserts an EqualBlock's two sides are ContentHash-equal, calling a mislabeled EqualBlock
        // over two differing tables a defect — so the off path must be distinguishable from a genuine equal.
        var left = Docxodus.Ir.IrReader.Read(WithTable(Table(2, Row("a1", "b1"))));
        var right = Docxodus.Ir.IrReader.Read(WithTable(Table(2, Row("a1", "CHANGED"))));
        var settings = new Docxodus.Ir.Diff.IrDiffSettings { CompareTables = false };

        var script = Docxodus.Ir.Diff.IrEditScriptBuilder.Build(left, right, settings);

        Docxodus.Tests.Ir.Diff.IrEditScriptVerifier.Verify(left, right, script, settings);
    }

    [Fact]
    public void CompareTables_false_marks_the_op_uncompared_on_the_data_surface()
    {
        // "Not compared" must not read as "equal" to a consumer of the edit script.
        var json = DocxDiff.GetEditScriptJson(
            WithTable(Table(2, Row("a1", "b1"))), WithTable(Table(2, Row("a1", "CHANGED"))), Off);

        Assert.Contains("\"uncompared\": true", json);
    }

    // --- the scope boundary: what "off" must NOT silence --------------------------------------------

    [Fact]
    public void CompareTables_false_still_reports_a_whole_table_added()
    {
        // A table appearing or vanishing is an ordinary block-level insertion/deletion, not a table
        // COMPARISON — so it survives the toggle, and it stays fully reversible.
        var left = Doc(
            "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");
        var right = WithTable(Table(2, Row("a1", "b1")));

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(compared));
    }

    [Fact]
    public void CompareTables_false_still_reports_a_whole_table_removed()
    {
        var left = WithTable(Table(2, Row("a1", "b1")));
        var right = Doc(
            "<w:p><w:r><w:t>Alpha unique intro paragraph.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

        Assert.NotEmpty(DocxDiff.GetRevisions(left, right, Off));

        var compared = DocxDiff.Compare(left, right, Off);
        AssertValid(compared);
        Assert.Equal(Paragraphs(right.DocumentByteArray), Accepted(compared));
        Assert.Equal(Paragraphs(left.DocumentByteArray), Rejected(compared));
    }

    [Fact]
    public void CompareTables_false_still_reports_a_body_change_beside_a_table()
    {
        // The table is byte-identical on both sides, so this cannot pass for a table-related reason.
        var left = Doc(
            "<w:p><w:r><w:t>Alpha intro one.</w:t></w:r></w:p>" + Table(2, Row("a1", "b1")) +
            "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");
        var right = Doc(
            "<w:p><w:r><w:t>Alpha intro two.</w:t></w:r></w:p>" + Table(2, Row("a1", "b1")) +
            "<w:p><w:r><w:t>Omega unique closing paragraph.</w:t></w:r></w:p>");

        Assert.Equal(
            DocxDiff.GetRevisions(left, right).Select(r => $"{r.Type}={r.Text}").ToArray(),
            DocxDiff.GetRevisions(left, right, Off).Select(r => $"{r.Type}={r.Text}").ToArray());
        Assert.Equal(BodyXml(DocxDiff.Compare(left, right)), BodyXml(DocxDiff.Compare(left, right, Off)));
    }

    [Fact]
    public void CompareTables_false_is_a_no_op_when_the_table_is_unchanged()
    {
        var left = WithTable(Table(2, Row("a1", "b1")));
        var right = WithTable(Table(2, Row("a1", "b1")));

        Assert.Empty(DocxDiff.GetRevisions(left, right, Off));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
    }

    [Fact]
    public void Consolidate_honors_CompareTables_per_reviewer()
    {
        // Passed through rather than forced (the CompareFields precedent). Off, a reviewer's table edit is
        // not merged and the base table survives.
        var baseDoc = WithTable(Table(2, Row("a1", "b1")));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = WithTable(Table(2, Row("a1", "CHANGED"))), Author = "Reviewer One" },
        };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers);
        AssertValid(merged);
        Assert.Contains("CHANGED", Accepted(merged));

        var mergedOff = DocxDiff.Consolidate(baseDoc, reviewers,
            new DocxDiffConsolidateSettings { Diff = Off });
        AssertValid(mergedOff);
        Assert.DoesNotContain("CHANGED", Accepted(mergedOff));
        Assert.Contains("b1", Accepted(mergedOff));
    }

    // --- wire ---------------------------------------------------------------------------------------

    [Fact]
    public void CompareTables_parses_from_the_settings_wire()
    {
        var left = WithTable(Table(2, Row("a1", "b1"))).DocumentByteArray;
        var right = WithTable(Table(2, Row("a1", "CHANGED"))).DocumentByteArray;

        Assert.Contains("rowOps",
            Docxodus.Internal.DocxDiffOps.GetEditScriptJson(left, right, null));
        Assert.DoesNotContain("rowOps",
            Docxodus.Internal.DocxDiffOps.GetEditScriptJson(left, right, "{\"compareTables\":false}"));
    }
}
