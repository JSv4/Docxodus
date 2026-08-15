#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Canonical table identity/coordinate coverage for issue #450 (including #471).</summary>
public class DocxSessionTableAddressingTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static byte[] BodyDoc(string bodyXml) =>
        IrTestDocuments.FromBodyXml(bodyXml).DocumentByteArray;

    private static string TableXml(string grid, string rows) =>
        $"<w:tbl><w:tblPr/>{grid}{rows}</w:tbl>";

    private static string CellXml(string text, string properties = "") =>
        $"<w:tc><w:tcPr>{properties}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Grid(int count) =>
        "<w:tblGrid>" + string.Concat(Enumerable.Range(0, count)
            .Select(index => $"<w:gridCol w:w=\"{1000 + index * 100}\"/>")) + "</w:tblGrid>";

    private static string AnchorId(DocxSession session, string kind, string? scope = null, int skip = 0) =>
        session.AnchorIndex().Values
            .Where(target => target.Anchor.Kind == kind && (scope is null || target.Anchor.Scope == scope))
            .Skip(skip).First().Anchor.Id;

    private static XElement MainXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.GetXDocument().Root!;
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(document)
            .Select(error => $"{error.Path?.XPath}: {error.Description}").ToList();
        Assert.True(errors.Count == 0, "OOXML schema errors:\n" + string.Join("\n", errors));
    }

    [Fact]
    public void DT250_RectangularMetadata_ExposesEveryIdentityAndRoundTripsCoordinates()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var bodyParagraph = AnchorId(session, "p");
        var inserted = session.InsertTable(bodyParagraph, Position.After, 2, 3);
        Assert.True(inserted.Success, inserted.Error?.Message);
        Assert.All(inserted.Created, anchor => Assert.Equal("tc", anchor.Kind));

        var tableId = AnchorId(session, "tbl");
        var result = session.GetTableMetadata(tableId);
        Assert.True(result.Success, result.Error?.Message);
        var table = result.Metadata!;
        Assert.Equal("tbl", table.Anchor.Kind);
        Assert.Equal(3, table.Columns.Count);
        Assert.All(table.Columns, column =>
        {
            Assert.Equal("col", column.Anchor.Kind);
            Assert.False(column.IsVirtual);
        });
        Assert.Equal(2, table.Rows.Count);
        Assert.All(table.Rows, row => Assert.Equal("tr", row.Anchor.Kind));
        Assert.All(table.Rows.SelectMany(row => row.Cells), cell => Assert.Equal("tc", cell.Anchor.Kind));

        foreach (var cell in table.Rows.SelectMany(row => row.Cells))
        {
            var byAnchor = session.ResolveTableCellAnchor(cell.Anchor.Id);
            var byCoordinate = session.ResolveTableCellCoordinate(
                table.Anchor.Id, cell.RowIndex, cell.ColumnIndex);
            Assert.True(byAnchor.Success, byAnchor.Error?.Message);
            Assert.True(byCoordinate.Success, byCoordinate.Error?.Message);
            Assert.Equal(cell.Anchor.Id, byAnchor.Cell!.Anchor.Id);
            Assert.Equal(cell.Anchor.Id, byCoordinate.Cell!.Anchor.Id);
        }

        var columnIds = table.Columns.Select(column => column.Anchor.Id).ToArray();
        Assert.True(session.SetColumnWidths(inserted.Created[0].Id, new[] { 2000, 3000, 4000 }).Success);
        var stable = session.GetTableMetadata(tableId).Metadata!;
        Assert.Equal(columnIds, stable.Columns.Select(column => column.Anchor.Id).ToArray());

        using var reopened = new DocxSession(session.Save(persistAnchorIds: true));
        var afterReopen = reopened.GetTableMetadata(tableId);
        Assert.True(afterReopen.Success, afterReopen.Error?.Message);
        Assert.Equal(stable.Anchor.Id, afterReopen.Metadata!.Anchor.Id);
        Assert.Equal(stable.Columns.Select(column => column.Anchor.Id),
            afterReopen.Metadata.Columns.Select(column => column.Anchor.Id));
        Assert.Equal(stable.Rows.Select(row => row.Anchor.Id),
            afterReopen.Metadata.Rows.Select(row => row.Anchor.Id));
        Assert.Equal(stable.Rows.SelectMany(row => row.Cells).Select(cell => cell.Anchor.Id),
            afterReopen.Metadata.Rows.SelectMany(row => row.Cells).Select(cell => cell.Anchor.Id));
    }

    [Fact]
    public void DT251_RaggedSpannedVerticalGrid_HasExactCoordinatesAndRowSpan()
    {
        var rows =
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:gridAfter w:val=\"1\"/></w:trPr>" +
            CellXml("lead", "<w:gridSpan w:val=\"2\"/><w:vMerge w:val=\"restart\"/>") + "</w:tr>" +
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:gridAfter w:val=\"1\"/></w:trPr>" +
            CellXml("", "<w:gridSpan w:val=\"2\"/><w:vMerge/>") + "</w:tr>";
        using var session = new DocxSession(BodyDoc(TableXml(Grid(4), rows) + "<w:p/>"));
        var tableId = AnchorId(session, "tbl");
        var table = session.GetTableMetadata(tableId).Metadata!;

        Assert.Equal((1, 1), (table.Rows[0].GridBefore, table.Rows[0].GridAfter));
        var lead = Assert.Single(table.Rows[0].Cells);
        var continuation = Assert.Single(table.Rows[1].Cells);
        Assert.Equal((1, 2, 2), (lead.ColumnIndex, lead.ColumnSpan, lead.RowSpan));
        Assert.Equal(TableVerticalMergeRole.Restart, lead.VerticalMerge);
        Assert.Equal((0, TableVerticalMergeRole.Continue),
            (continuation.RowSpan, continuation.VerticalMerge));
        Assert.Equal(lead.Anchor.Id,
            session.ResolveTableCellCoordinate(tableId, 0, 1).Cell!.Anchor.Id);
        Assert.Equal(lead.Anchor.Id,
            session.ResolveTableCellCoordinate(tableId, 0, 2).Cell!.Anchor.Id);
        var leadingGap = session.ResolveTableCellCoordinate(tableId, 0, 0);
        var trailingGap = session.ResolveTableCellCoordinate(tableId, 0, 3);
        Assert.False(leadingGap.Success);
        Assert.False(trailingGap.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, leadingGap.Error!.Code);
        Assert.Contains("(0, 0)", leadingGap.Error.Message);

        var structure = DocumentStructureAnalyzer.Analyze(
            IrTestDocuments.FromBodyXml(TableXml(Grid(4), rows) + "<w:p/>"));
        var structureTable = Assert.Single(structure.FindByType(DocumentElementType.Table));
        Assert.Equal(tableId, structureTable.AnchorId);
        var cells = structure.FindByType(DocumentElementType.TableCell).ToList();
        Assert.Equal((1, 2, 2), (cells[0].ColumnIndex, cells[0].ColumnSpan, cells[0].RowSpan));
        Assert.Equal(0, cells[1].RowSpan);
        Assert.Equal(lead.Anchor.Id, cells[0].AnchorId);
        Assert.Equal(continuation.Anchor.Id, cells[1].AnchorId);
        Assert.Equal(table.Columns.Select(column => column.Anchor.Id),
            structure.GetTableColumns(structureTable.Id).Select(column => column.AnchorId));

        // Inserting at the exact start of a trailing omission creates a physical cell; the
        // existing omitted suffix moves right and remains an omission. Deletion is its inverse.
        var trailingInsert = session.InsertTableColumn(lead.Anchor.Id, Position.After);
        Assert.True(trailingInsert.Success, trailingInsert.Error?.Message);
        Assert.Equal(2, trailingInsert.Created.Count);
        var withTrailingInsert = session.GetTableMetadata(tableId).Metadata!;
        Assert.Equal(5, withTrailingInsert.Columns.Count);
        Assert.All(withTrailingInsert.Rows, row =>
        {
            Assert.Equal(1, row.GridAfter);
            Assert.Equal(2, row.Cells.Count);
            Assert.Equal(3, row.Cells[^1].ColumnIndex);
        });
        var trailingDelete = session.DeleteTableColumn(trailingInsert.Created[0].Id);
        Assert.True(trailingDelete.Success, trailingDelete.Error?.Message);
        var afterTrailingDelete = session.GetTableMetadata(tableId).Metadata!;
        Assert.Equal(4, afterTrailingDelete.Columns.Count);
        Assert.All(afterTrailingDelete.Rows, row => Assert.Equal(1, row.GridAfter));

        // A row that omits the insertion coordinate through gridBefore adjusts that omission;
        // the addressed row receives a cell at its first physical boundary.
        var asymmetricRows =
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:gridAfter w:val=\"1\"/></w:trPr>" +
            CellXml("a") + CellXml("b") + "</w:tr>" +
            "<w:tr><w:trPr><w:gridBefore w:val=\"2\"/></w:trPr>" +
            CellXml("c") + CellXml("d") + "</w:tr>";
        using var leadingSession = new DocxSession(
            BodyDoc(TableXml(Grid(4), asymmetricRows) + "<w:p/>"));
        var leadingTableId = AnchorId(leadingSession, "tbl");
        var leadingBefore = leadingSession.GetTableMetadata(leadingTableId).Metadata!;
        var leadingInsert = leadingSession.InsertTableColumn(
            leadingBefore.Rows[0].Cells[0].Anchor.Id, Position.Before);
        Assert.True(leadingInsert.Success, leadingInsert.Error?.Message);
        Assert.Single(leadingInsert.Created);
        var withLeadingInsert = leadingSession.GetTableMetadata(leadingTableId).Metadata!;
        Assert.Equal((1, 3), (withLeadingInsert.Rows[0].GridBefore, withLeadingInsert.Rows[1].GridBefore));
        var leadingDelete = leadingSession.DeleteTableColumn(leadingInsert.Created[0].Id);
        Assert.True(leadingDelete.Success, leadingDelete.Error?.Message);
        var afterLeadingDelete = leadingSession.GetTableMetadata(leadingTableId).Metadata!;
        Assert.Equal((1, 2), (afterLeadingDelete.Rows[0].GridBefore, afterLeadingDelete.Rows[1].GridBefore));

        // At the far edge of a shorter row's trailing omission, the row remains ragged: the
        // appended grid column is omitted there and materialized only in the full-width row.
        var trailingEdgeRows =
            "<w:tr>" + CellXml("e") + CellXml("f") + CellXml("g") + CellXml("h") + "</w:tr>" +
            "<w:tr><w:trPr><w:gridAfter w:val=\"2\"/></w:trPr>" +
            CellXml("i") + CellXml("j") + "</w:tr>";
        using var trailingEdgeSession = new DocxSession(
            BodyDoc(TableXml(Grid(4), trailingEdgeRows) + "<w:p/>"));
        var trailingEdgeTableId = AnchorId(trailingEdgeSession, "tbl");
        var trailingEdgeBefore = trailingEdgeSession.GetTableMetadata(trailingEdgeTableId).Metadata!;
        var trailingEdgeInsert = trailingEdgeSession.InsertTableColumn(
            trailingEdgeBefore.Rows[0].Cells[^1].Anchor.Id, Position.After);
        Assert.True(trailingEdgeInsert.Success, trailingEdgeInsert.Error?.Message);
        Assert.Single(trailingEdgeInsert.Created);
        var withTrailingEdgeInsert = trailingEdgeSession.GetTableMetadata(trailingEdgeTableId).Metadata!;
        Assert.Equal((0, 3),
            (withTrailingEdgeInsert.Rows[0].GridAfter, withTrailingEdgeInsert.Rows[1].GridAfter));
        var trailingEdgeDelete = trailingEdgeSession.DeleteTableColumn(trailingEdgeInsert.Created[0].Id);
        Assert.True(trailingEdgeDelete.Success, trailingEdgeDelete.Error?.Message);
        var afterTrailingEdgeDelete = trailingEdgeSession.GetTableMetadata(trailingEdgeTableId).Metadata!;
        Assert.Equal((0, 2),
            (afterTrailingEdgeDelete.Rows[0].GridAfter, afterTrailingEdgeDelete.Rows[1].GridAfter));
    }

    [Fact]
    public void DT252_NestedTableCanonicalCell_NeverRetargetsOuterCell()
    {
        var nested = TableXml(Grid(1), "<w:tr>" + CellXml("inner") + "</w:tr>");
        var outerCell = "<w:tc><w:tcPr/><w:p><w:r><w:t>outer-before</w:t></w:r></w:p>" +
            nested + "<w:p><w:r><w:t>outer-after</w:t></w:r></w:p></w:tc>";
        using var session = new DocxSession(BodyDoc(
            TableXml(Grid(1), "<w:tr>" + outerCell + "</w:tr>") + "<w:p/>"));
        var outerTableId = AnchorId(session, "tbl", skip: 0);
        var innerTableId = AnchorId(session, "tbl", skip: 1);
        var outer = session.GetTableMetadata(outerTableId).Metadata!;
        var inner = session.GetTableMetadata(innerTableId).Metadata!;
        var outerCellMetadata = Assert.Single(Assert.Single(outer.Rows).Cells);
        var innerCellMetadata = Assert.Single(Assert.Single(inner.Rows).Cells);
        Assert.Equal(2, outerCellMetadata.ParagraphAnchors.Count);
        Assert.Single(innerCellMetadata.ParagraphAnchors);
        Assert.DoesNotContain(innerCellMetadata.ParagraphAnchors[0], outerCellMetadata.ParagraphAnchors);

        var replaced = session.ReplaceCellContent(innerCellMetadata.Anchor.Id, "inner-new");
        Assert.True(replaced.Success, replaced.Error?.Message);
        Assert.Equal(innerCellMetadata.Anchor.Id, Assert.Single(replaced.Modified).Id);
        var insertedInnerRow = session.InsertTableRow(innerCellMetadata.Anchor.Id, Position.After);
        Assert.True(insertedInnerRow.Success, insertedInnerRow.Error?.Message);
        Assert.Equal(innerTableId, insertedInnerRow.TableAnchors!.Retained
            .Single(entry => entry.Before.EntityKind == TableAnchorEntityKind.Table).After.Anchor.Id);
        Assert.Single(session.GetTableMetadata(outerTableId).Metadata!.Rows);
        Assert.Equal(2, session.GetTableMetadata(innerTableId).Metadata!.Rows.Count);
        var xml = MainXml(session.Save());
        var tables = xml.Descendants(W + "tbl").ToList();
        Assert.Single(tables[0].Elements(W + "tr"));
        Assert.Equal(2, tables[1].Elements(W + "tr").Count());
        Assert.Contains("outer-before", string.Concat(tables[0].Elements(W + "tr")
            .Elements(W + "tc").Elements(W + "p").Descendants(W + "t")));
        Assert.Equal("inner-new", string.Concat(tables[1].Descendants(W + "t").Select(text => text.Value)));

        var structural = session.SetCellShading(innerTableId, "FF0000");
        Assert.False(structural.Success);
        Assert.Equal(EditErrorCode.TableAnchorMigrationRequired, structural.Error!.Code);
        Assert.Contains("canonical tc anchor", structural.Error.Message);
    }

    [Fact]
    public void DT253_LegacyParagraphTranslation_IsNearestCellOnlyAndCanonicalizesResult()
    {
        using var session = new DocxSession(BodyDoc(
            TableXml(Grid(1), "<w:tr>" + CellXml("old") + "</w:tr>") + "<w:p><w:r><w:t>body</w:t></w:r></w:p>"));
        var cell = AnchorId(session, "tc");
        var cellParagraph = session.GetTableMetadata(AnchorId(session, "tbl"))
            .Metadata!.Rows[0].Cells[0].ParagraphAnchors[0].Id;
        var bodyParagraph = session.AnchorIndex().Values
            .First(target => target.Anchor.Kind == "p" && target.Anchor.Id != cellParagraph).Anchor.Id;
        var translated = session.ReplaceCellContent(cellParagraph, "new");
        Assert.True(translated.Success, translated.Error?.Message);
        Assert.Equal(cell, Assert.Single(translated.Modified).Id);

        var rejected = session.InsertTableRow(bodyParagraph, Position.After);
        Assert.False(rejected.Success);
        Assert.Equal(EditErrorCode.TableAnchorMigrationRequired, rejected.Error!.Code);
        Assert.Contains("GetTableMetadata", rejected.Error.Message);
    }

    [Fact]
    public void DT254_StructuralMappings_AreDeterministicAndCanonical()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var inserted = session.InsertTable(AnchorId(session, "p"), Position.After, 2, 2);
        var tableId = AnchorId(session, "tbl");
        var before = session.GetTableMetadata(tableId).Metadata!;
        var oldBottomLeft = before.Rows[1].Cells[0];

        var rowInsert = session.InsertTableRow(before.Rows[0].Cells[0].Anchor.Id, Position.After);
        Assert.True(rowInsert.Success, rowInsert.Error?.Message);
        Assert.All(rowInsert.Created, anchor => Assert.Equal("tc", anchor.Kind));
        var mapping = rowInsert.TableAnchors!;
        Assert.Equal(3, mapping.Added.Count); // one tr + two tc
        Assert.Empty(mapping.Invalidated);
        var shifted = mapping.Retained.Single(entry => entry.Before.Anchor.Id == oldBottomLeft.Anchor.Id);
        Assert.Equal(1, shifted.Before.RowIndex);
        Assert.Equal(2, shifted.After.RowIndex);

        var after = session.GetTableMetadata(tableId).Metadata!;
        var doomed = after.Rows[1].Cells[1].Anchor.Id;
        var columnDelete = session.DeleteTableColumn(doomed);
        Assert.True(columnDelete.Success, columnDelete.Error?.Message);
        Assert.All(columnDelete.Removed, anchor => Assert.Equal("tc", anchor.Kind));
        Assert.Contains(columnDelete.TableAnchors!.Invalidated,
            location => location.Anchor.Id == doomed && location.EntityKind == TableAnchorEntityKind.Cell);
        Assert.Equal(columnDelete.Removed.Select(anchor => anchor.Id),
            columnDelete.TableAnchors.Invalidated
                .Where(location => location.EntityKind == TableAnchorEntityKind.Cell)
                .Select(location => location.Anchor.Id));
    }

    [Fact]
    public void DT255_MissingGrid_MetadataIsReadOnlyAndColumnEditReplacesVirtualIdentities()
    {
        using var session = new DocxSession(BodyDoc(
            TableXml("", "<w:tr>" + CellXml("a") + CellXml("b") + "</w:tr>") + "<w:p/>"));
        var tableId = AnchorId(session, "tbl");
        var before = session.GetTableMetadata(tableId).Metadata!;
        Assert.Equal(2, before.Columns.Count);
        Assert.All(before.Columns, column => Assert.True(column.IsVirtual));
        Assert.All(before.Columns, column => Assert.Matches("^[0-9a-f]{32}$", column.Anchor.Unid));
        Assert.Null(MainXml(session.Save()).Descendants(W + "tblGrid").FirstOrDefault());

        var edit = session.InsertTableColumn(before.Rows[0].Cells[0].Anchor.Id, Position.After);
        Assert.True(edit.Success, edit.Error?.Message);
        Assert.All(before.Columns, column => Assert.Contains(edit.TableAnchors!.Invalidated,
            location => location.Anchor.Id == column.Anchor.Id && location.IsVirtual));
        var after = session.GetTableMetadata(tableId).Metadata!;
        Assert.Equal(3, after.Columns.Count);
        Assert.All(after.Columns, column => Assert.False(column.IsVirtual));
        Assert.All(after.Columns, column => Assert.Contains(edit.TableAnchors!.Added,
            location => location.Anchor.Id == column.Anchor.Id));
    }

    [Fact]
    public void DT256_HeaderTable_UsesScopedCanonicalAnchors()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("body")))));
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header();
            header.PutXDocument(XDocument.Parse(
                $"<w:hdr xmlns:w=\"{W}\">{TableXml(Grid(1), "<w:tr>" + CellXml("header") + "</w:tr>")}</w:hdr>"));
            var relationship = main.GetIdOfPart(header);
            main.Document.Body!.Append(new SectionProperties(
                new HeaderReference { Id = relationship, Type = HeaderFooterValues.Default }));
            main.Document.Save();
        }

        using var session = new DocxSession(stream.ToArray());
        var tableId = AnchorId(session, "tbl", "hdr1");
        var metadata = session.GetTableMetadata(tableId).Metadata!;
        Assert.Equal("hdr1", metadata.Anchor.Scope);
        Assert.Equal("hdr1", metadata.Rows[0].Cells[0].Anchor.Scope);
        var edit = session.ReplaceCellContent(metadata.Rows[0].Cells[0].Anchor.Id, "header-new");
        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("hdr1", Assert.Single(edit.Modified).Scope);
    }

    [Fact]
    public void DT257_TableInsideRevisionWrapper_HasCanonicalAnchorsAndNativeRowRevision()
    {
        var revisedRow =
            "<w:tr><w:trPr><w:ins w:id=\"7\" w:author=\"A\" " +
            "w:date=\"2026-01-01T00:00:00Z\"/></w:trPr>" + CellXml("revision-table") + "</w:tr>";
        var input = BodyDoc(TableXml(Grid(1), revisedRow) + "<w:p/>");
        AssertSchemaValid(input);
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var tableId = AnchorId(session, "tbl");
        var metadata = session.GetTableMetadata(tableId).Metadata!;
        var edit = session.InsertTableRow(metadata.Rows[0].Cells[0].Anchor.Id, Position.After);
        Assert.True(edit.Success, edit.Error?.Message);
        Assert.All(edit.Created, anchor => Assert.Equal("tc", anchor.Kind));

        var saved = session.Save();
        AssertSchemaValid(saved);
        var xml = MainXml(saved);
        var tableElement = xml.Descendants(W + "tbl").Single();
        var rows = tableElement.Elements(W + "tr").ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("7", (string?)Assert.Single(rows[0].Element(W + "trPr")!.Elements(W + "ins"))
            .Attribute(W + "id"));
        var inserted = Assert.Single(rows[1].Element(W + "trPr")!.Elements(W + "ins"));
        Assert.NotEqual("7", (string?)inserted.Attribute(W + "id"));
        Assert.Single(rows[1].Descendants(W + "rPr").Elements(W + "ins"));
    }
}
