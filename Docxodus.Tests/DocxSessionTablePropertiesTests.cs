using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word's Table Properties on <see cref="DocxSession"/>: SetTableProperties (width, alignment,
/// indent, default cell margins, alt text), SetCellProperties (vertical alignment by cell, row,
/// column or table) and GetTableProperties. Test IDs use the TP0xx range.
/// </summary>
public class DocxSessionTablePropertiesTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement DocumentXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!;
    }

    private static (DocxSession session, string[] cells) NewTable(int rows, int cols, DocxSessionSettings? settings = null)
    {
        var session = settings is null
            ? new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs())
            : new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), settings);
        var anchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        var r = session.InsertTable(anchor, Position.After, rows, cols, new TableInsertOptions());
        Assert.True(r.Success, r.Error?.Message);
        return (session, r.Created.Select(a => a.Id).ToArray());
    }

    private static XElement Table(DocxSession session) =>
        DocumentXml(session.Save()).Descendants(W + "tbl").Single();

    private static void AssertSchemaValid(byte[] bytes) =>
        PackageValidation.AssertValid(bytes, DocumentFormat.OpenXml.FileFormatVersions.Office2007);

    [Fact]
    public void TP001_SetTableProperties_WritesEachPropertyInSchemaOrder_AndReadsBack()
    {
        var (session, cells) = NewTable(2, 2);
        var r = session.SetTableProperties(cells[0], new TablePropertiesOp
        {
            WidthType = TableWidthType.Percent,
            Width = 80,
            Alignment = "center",
            IndentTwips = 360,
            CellMargins = new TableCellMargins { TopTwips = 0, LeftTwips = 144, BottomTwips = 0, RightTwips = 144 },
            AltTitle = "Prices",
            AltDescription = "Prices by year",
        });
        Assert.True(r.Success, r.Error?.Message);

        var tblPr = Table(session).Element(W + "tblPr")!;
        Assert.Equal("4000", (string?)tblPr.Element(W + "tblW")!.Attribute(W + "w"));
        Assert.Equal("pct", (string?)tblPr.Element(W + "tblW")!.Attribute(W + "type"));
        Assert.Equal("center", (string?)tblPr.Element(W + "jc")!.Attribute(W + "val"));
        Assert.Equal("360", (string?)tblPr.Element(W + "tblInd")!.Attribute(W + "w"));
        var mar = tblPr.Element(W + "tblCellMar")!;
        Assert.Equal(new[] { "top", "left", "bottom", "right" }, mar.Elements().Select(e => e.Name.LocalName));
        Assert.Equal("144", (string?)mar.Element(W + "left")!.Attribute(W + "w"));
        Assert.Equal("Prices", (string?)tblPr.Element(W + "tblCaption")!.Attribute(W + "val"));
        // Alt text (w:tblCaption, w:tblDescription) is Word 2010's.
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Office2010);

        var info = session.GetTableProperties(cells[3])!;
        Assert.Equal(TableWidthType.Percent, info.WidthType);
        Assert.Equal(80, info.Width);
        Assert.Equal("center", info.Alignment);
        Assert.Equal(360, info.IndentTwips);
        Assert.Equal(144, info.CellMargins!.RightTwips);
        Assert.Equal("Prices by year", info.AltDescription);
        Assert.Equal(1, info.RowIndex);
        Assert.Equal(2, info.RowCount);
        Assert.Equal(1, info.ColumnIndex);
        Assert.Equal(2, info.ColumnWidthsTwips.Count);
    }

    [Fact]
    public void TP002_Defaults_RemoveTheirElements_AndOneUndoRestores()
    {
        var (session, cells) = NewTable(2, 2);
        Assert.True(session.SetTableProperties(cells[0], new TablePropertiesOp
        {
            Alignment = "right", IndentTwips = 200, AltTitle = "T",
        }).Success);
        var before = Table(session);

        var r = session.SetTableProperties(cells[0], new TablePropertiesOp
        {
            WidthType = TableWidthType.Twips, Width = 7200, Alignment = "left", IndentTwips = 0, AltTitle = "",
        });
        Assert.True(r.Success, r.Error?.Message);
        var tblPr = Table(session).Element(W + "tblPr")!;
        Assert.Null(tblPr.Element(W + "jc"));
        Assert.Null(tblPr.Element(W + "tblInd"));
        Assert.Null(tblPr.Element(W + "tblCaption"));
        Assert.Equal("7200", (string?)tblPr.Element(W + "tblW")!.Attribute(W + "w"));
        Assert.Equal("dxa", (string?)tblPr.Element(W + "tblW")!.Attribute(W + "type"));

        Assert.True(session.Undo());
        Assert.True(XNode.DeepEquals(before, Table(session)));
    }

    [Theory]
    [InlineData("middle", null, 50.0, null)]
    [InlineData(null, TableWidthType.Percent, 120.0, null)]
    [InlineData(null, TableWidthType.Twips, 0.0, null)]
    [InlineData(null, null, null, -5)]
    public void TP003_InvalidValues_AreRefused_AndChangeNothing(string? alignment, TableWidthType? type, double? width, int? margin)
    {
        var (session, cells) = NewTable(2, 2);
        var before = Table(session);
        long version = session.Version;
        var r = session.SetTableProperties(cells[0], new TablePropertiesOp
        {
            Alignment = alignment, WidthType = type, Width = width,
            CellMargins = margin is null ? null : new TableCellMargins { LeftTwips = margin },
        });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidTableStyling, r.Error!.Code);
        Assert.Equal(version, session.Version);
        Assert.True(XNode.DeepEquals(before, Table(session)));
    }

    [Fact]
    public void TP004_SetCellProperties_ByCellRowColumnAndTable()
    {
        var (session, cells) = NewTable(3, 3);
        string VAlign(int row, int col) =>
            (string?)Table(session).Elements(W + "tr").ElementAt(row).Elements(W + "tc").ElementAt(col)
                .Element(W + "tcPr")?.Element(W + "vAlign")?.Attribute(W + "val") ?? "-";

        Assert.True(session.SetCellProperties(cells[4], new CellPropertiesOp { VerticalAlignment = "center" }).Success);
        Assert.Equal("center", VAlign(1, 1));
        Assert.Equal("-", VAlign(1, 0));

        Assert.True(session.SetCellProperties(cells[4], new CellPropertiesOp { VerticalAlignment = "bottom" }, TableCellScope.Row).Success);
        Assert.Equal(new[] { "bottom", "bottom", "bottom" }, new[] { VAlign(1, 0), VAlign(1, 1), VAlign(1, 2) });
        Assert.Equal("-", VAlign(0, 0));

        Assert.True(session.SetCellProperties(cells[2], new CellPropertiesOp { VerticalAlignment = "center" }, TableCellScope.Column).Success);
        Assert.Equal(new[] { "center", "center", "center" }, new[] { VAlign(0, 2), VAlign(1, 2), VAlign(2, 2) });
        Assert.Equal("-", VAlign(0, 1));

        // "top" is Word's default: written as no vAlign at all.
        Assert.True(session.SetCellProperties(cells[0], new CellPropertiesOp { VerticalAlignment = "top" }, TableCellScope.Table).Success);
        Assert.Empty(Table(session).Descendants(W + "vAlign"));
        Assert.Equal("top", session.GetTableProperties(cells[4])!.VerticalAlignment ?? "top");
        AssertSchemaValid(session.Save());

        var bad = session.SetCellProperties(cells[0], new CellPropertiesOp { VerticalAlignment = "middle" });
        Assert.False(bad.Success);
        Assert.Equal(EditErrorCode.InvalidTableStyling, bad.Error!.Code);
    }

    [Fact]
    public void TP005_ColumnScope_TakesAMergedCellThatCoversTheColumn()
    {
        var (session, cells) = NewTable(2, 3);
        var merged = session.MergeCells(cells[0], 1, 2);
        Assert.True(merged.Success, merged.Error?.Message);
        var cell = session.GetTableMetadata(session.GetTableProperties(cells[4])!.TableAnchorId).Metadata!
            .Rows[1].Cells[1].Anchor.Id;
        Assert.True(session.SetCellProperties(cell, new CellPropertiesOp { VerticalAlignment = "bottom" }, TableCellScope.Column).Success);
        var rows = Table(session).Elements(W + "tr").ToList();
        Assert.Equal("bottom", (string?)rows[0].Elements(W + "tc").First().Element(W + "tcPr")!.Element(W + "vAlign")!.Attribute(W + "val"));
        Assert.Equal("bottom", (string?)rows[1].Elements(W + "tc").ElementAt(1).Element(W + "tcPr")!.Element(W + "vAlign")!.Attribute(W + "val"));
        Assert.Null(rows[1].Elements(W + "tc").ElementAt(2).Element(W + "tcPr")!.Element(W + "vAlign"));
    }

    [Fact]
    public void TP006_Tracked_RecordsTblPrChangeAndTcPrChange()
    {
        var (session, cells) = NewTable(2, 2, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Bob",
        });
        Assert.True(session.SetTableProperties(cells[0], new TablePropertiesOp { Alignment = "center" }).Success);
        Assert.True(session.SetCellProperties(cells[0], new CellPropertiesOp { VerticalAlignment = "center" }).Success);
        var tbl = Table(session);
        var change = tbl.Element(W + "tblPr")!.Element(W + "tblPrChange")!;
        Assert.Equal("Bob", (string?)change.Attribute(W + "author"));
        Assert.Null(change.Element(W + "tblPr")!.Element(W + "jc"));
        Assert.NotNull(tbl.Descendants(W + "tc").First().Element(W + "tcPr")!.Element(W + "tcPrChange"));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void TP007_GetTableProperties_ReadsRowOptions_AndIsNullOutsideATable()
    {
        var (session, cells) = NewTable(2, 2);
        Assert.True(session.SetTableRowOptions(cells[2], new TableRowOptions
        {
            HeightTwips = 600, HeightRule = TableRowHeightRule.Exact, AllowBreakAcrossPages = false, RepeatHeader = true,
        }).Success);
        var info = session.GetTableProperties(cells[3])!;
        Assert.Equal(600, info.RowHeightTwips);
        Assert.Equal(TableRowHeightRule.Exact, info.RowHeightRule);
        Assert.False(info.AllowBreakAcrossPages);
        Assert.True(info.RepeatHeader);
        Assert.Null(session.GetTableProperties(cells[0])!.RowHeightTwips);

        var paragraph = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" && !t.Anchor.Id.Contains("tc")).Anchor.Id;
        Assert.Null(session.GetTableProperties(paragraph));
    }

    [Fact]
    public void TP008_Json_ParsesTheOp_AndSerializesTheInfo()
    {
        var op = DocxSessionJson.ParseTablePropertiesOp(
            "{\"widthType\":\"percent\",\"width\":50,\"alignment\":\"center\",\"cellMargins\":{\"leftTwips\":72}}");
        Assert.Equal(TableWidthType.Percent, op.WidthType);
        Assert.Equal(50, op.Width);
        Assert.Equal(72, op.CellMargins!.LeftTwips);
        Assert.Null(op.IndentTwips);
        Assert.Equal(TableCellScope.Column, DocxSessionJson.ParseTableCellScope("column"));

        var (session, cells) = NewTable(1, 2);
        Assert.True(session.SetTableProperties(cells[0], op).Success);
        using var json = JsonDocument.Parse(DocxSessionJson.SerializeTablePropertiesInfo(session.GetTableProperties(cells[1])));
        var root = json.RootElement;
        Assert.Equal("percent", root.GetProperty("widthType").GetString());
        Assert.Equal(50, root.GetProperty("width").GetDouble());
        Assert.Equal(72, root.GetProperty("cellMargins").GetProperty("leftTwips").GetInt32());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("cellMargins").GetProperty("topTwips").ValueKind);
        Assert.Equal(2, root.GetProperty("columnWidthsTwips").GetArrayLength());
        Assert.Equal("null", DocxSessionJson.SerializeTablePropertiesInfo(null));
    }

    /// <summary>A document written as literal XML: <paramref name="bodyXml"/> in w:body and
    /// <paramref name="stylesXml"/> in w:styles.</summary>
    private static byte[] LiteralDocument(string bodyXml, string stylesXml = "")
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            void Write(OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            Write(main, $"<w:document {ns}><w:body>{bodyXml}<w:p/></w:body></w:document>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}>{stylesXml}</w:styles>");
        }
        return stream.ToArray();
    }

    private static string[] Anchors(DocxSession session, string kind) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind == kind).Select(t => t.Anchor.Id).ToArray();

    [Fact]
    public void TP009_GetTableProperties_FindsTheTableAsTheSettersDo()
    {
        // An outer table holding a centred nested table, and a right-aligned table whose row and cell
        // are each wrapped in a content control.
        const string grid = "<w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>";
        var session = new DocxSession(LiteralDocument(
            $"<w:tbl><w:tblPr/>{grid}<w:tr><w:tc><w:p><w:r><w:t>Outer</w:t></w:r></w:p>"
            + $"<w:tbl><w:tblPr><w:jc w:val=\"center\"/></w:tblPr>{grid}<w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "<w:p/></w:tc></w:tr></w:tbl><w:p/>"
            + $"<w:tbl><w:tblPr><w:jc w:val=\"right\"/></w:tblPr>{grid}<w:sdt><w:sdtContent><w:tr>"
            + "<w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>Wrapped</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>"
            + "</w:tr></w:sdtContent></w:sdt></w:tbl>"));

        // Each table's own anchor reads that table, the nested one included.
        var tables = Anchors(session, "tbl");
        Assert.Equal(3, tables.Length);
        foreach (var table in tables)
            Assert.Equal(table, session.GetTableProperties(table)!.TableAnchorId);
        Assert.Equal(new[] { "-", "center", "right" },
            tables.Select(t => session.GetTableProperties(t)!.Alignment ?? "-").OrderBy(a => a));

        // Each cell's anchor, and a paragraph in it, read the table the cell is in.
        var cellAlignments = Anchors(session, "tc").Select(tc => session.GetTableProperties(tc)!.Alignment ?? "-");
        Assert.Equal(new[] { "-", "center", "right" }, cellAlignments.OrderBy(a => a));
        var wrapped = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p")
            .Select(t => session.GetTableProperties(t.Anchor.Id))
            .Where(info => info?.Alignment == "right").ToList();
        Assert.Single(wrapped);
        Assert.Equal(1, wrapped[0]!.RowCount);
        Assert.Equal(0, wrapped[0]!.RowIndex);
    }

    [Fact]
    public void TP010_Defaults_AreWrittenWhenTheTableStyleSetsOtherValues()
    {
        // "Grid" is based on "Base"; "Plain" is the default table style, used by a table with no tblStyle.
        const string styles =
            "<w:style w:type=\"table\" w:styleId=\"Base\"><w:name w:val=\"Base\"/>"
            + "<w:tblPr><w:jc w:val=\"center\"/></w:tblPr><w:tcPr><w:vAlign w:val=\"bottom\"/></w:tcPr></w:style>"
            + "<w:style w:type=\"table\" w:styleId=\"Grid\"><w:name w:val=\"Grid\"/><w:basedOn w:val=\"Base\"/>"
            + "<w:tblPr><w:tblInd w:w=\"720\" w:type=\"dxa\"/></w:tblPr></w:style>"
            + "<w:style w:type=\"table\" w:default=\"1\" w:styleId=\"Plain\"><w:name w:val=\"Plain\"/>"
            + "<w:tblPr><w:jc w:val=\"right\"/></w:tblPr></w:style>";
        const string grid = "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>";
        const string row = "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>";
        var session = new DocxSession(LiteralDocument(
            $"<w:tbl><w:tblPr><w:tblStyle w:val=\"Grid\"/></w:tblPr>{grid}{row}</w:tbl><w:p/>"
            + $"<w:tbl><w:tblPr/>{grid}{row}</w:tbl>", styles));
        var cells = Anchors(session, "tc");
        Assert.Equal(4, cells.Length);
        XElement[] Tables() => DocumentXml(session.Save()).Descendants(W + "tbl").ToArray();

        Assert.True(session.SetTableProperties(cells[0], new TablePropertiesOp { Alignment = "left", IndentTwips = 0 }).Success);
        Assert.True(session.SetCellProperties(cells[0], new CellPropertiesOp { VerticalAlignment = "top" }, TableCellScope.Table).Success);
        Assert.True(session.SetTableProperties(cells[2], new TablePropertiesOp { Alignment = "left", IndentTwips = 0 }).Success);
        Assert.True(session.SetCellProperties(cells[2], new CellPropertiesOp { VerticalAlignment = "top" }, TableCellScope.Table).Success);

        // The styled table: left over the base style's centre, 0 over Grid's indent, top over bottom.
        var styled = Tables()[0];
        Assert.Equal("left", (string?)styled.Element(W + "tblPr")!.Element(W + "jc")!.Attribute(W + "val"));
        Assert.Equal("0", (string?)styled.Element(W + "tblPr")!.Element(W + "tblInd")!.Attribute(W + "w"));
        Assert.Equal(new[] { "top", "top" }, styled.Descendants(W + "vAlign").Select(v => (string?)v.Attribute(W + "val")));

        // The other table takes the default style: left is written over its right; it sets no
        // indent or vertical alignment, so those stay unwritten.
        var plain = Tables()[1];
        Assert.Equal("left", (string?)plain.Element(W + "tblPr")!.Element(W + "jc")!.Attribute(W + "val"));
        Assert.Null(plain.Element(W + "tblPr")!.Element(W + "tblInd"));
        Assert.Empty(plain.Descendants(W + "vAlign"));
        AssertSchemaValid(session.Save());
    }
}
