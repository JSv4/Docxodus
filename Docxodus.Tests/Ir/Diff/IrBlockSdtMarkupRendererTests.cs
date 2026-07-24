#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// End-to-end coverage for block-level content-control envelopes. A control has OOXML-owned metadata
/// outside its payload, so its renderer must toggle both the wrapper (custom-XML range boundaries) and
/// its block payload (normal revision markup). These cases intentionally compare the whole canonical
/// envelope after accept/reject, not merely visible text.
/// </summary>
[Trait("Category", "Markup")]
public class IrBlockSdtMarkupRendererTests
{
    private static readonly IrReaderOptions ReadOptions =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    [Fact]
    public void Render_block_sdt_metadata_change_round_trips_exact_envelope()
    {
        var left = IrTestDocuments.FromBodyXml(BlockSdt("old-control", "same controlled payload"));
        var right = IrTestDocuments.FromBodyXml(BlockSdt("new-control", "same controlled payload"));

        var rendered = Render(left, right);
        var body = Body(rendered);
        var renderedSdts = body.Elements(W.sdt).ToList();
        Assert.Equal(2, renderedSdts.Count);

        var oldEnvelope = Assert.Single(renderedSdts.Where(s => Tag(s) == "old-control"));
        var newEnvelope = Assert.Single(renderedSdts.Where(s => Tag(s) == "new-control"));
        AssertEnvelopeRangeTopology(oldEnvelope, W.customXmlDelRangeStart, W.customXmlDelRangeEnd);
        AssertEnvelopeRangeTopology(newEnvelope, W.customXmlInsRangeStart, W.customXmlInsRangeEnd);
        AssertUniqueCustomXmlRangeStartIds(body);

        AssertExactTopLevelEnvelope(right, RevisionProcessor.AcceptRevisions(rendered));
        AssertExactTopLevelEnvelope(left, RevisionProcessor.RejectRevisions(rendered));
        AssertSchemaValid(rendered);
        AssertSchemaValid(RevisionProcessor.AcceptRevisions(rendered));
        AssertSchemaValid(RevisionProcessor.RejectRevisions(rendered));
    }

    [Fact]
    public void Render_nested_block_sdts_toggle_both_envelopes_and_round_trip_exactly()
    {
        var left = IrTestDocuments.FromBodyXml(NestedBlockSdt("old-outer", "old-inner", "old nested payload"));
        var right = IrTestDocuments.FromBodyXml(NestedBlockSdt("new-outer", "new-inner", "new nested payload"));

        var rendered = Render(left, right);
        var body = Body(rendered);
        var outerSdts = body.Elements(W.sdt).ToList();
        Assert.Equal(2, outerSdts.Count);

        var oldOuter = Assert.Single(outerSdts.Where(s => Tag(s) == "old-outer"));
        var newOuter = Assert.Single(outerSdts.Where(s => Tag(s) == "new-outer"));
        AssertEnvelopeAndNestedTopology(oldOuter, W.customXmlDelRangeStart, W.customXmlDelRangeEnd);
        AssertEnvelopeAndNestedTopology(newOuter, W.customXmlInsRangeStart, W.customXmlInsRangeEnd);
        AssertUniqueCustomXmlRangeStartIds(body);

        // This specifically proves that accepting/rejecting the outer range does not leave an empty
        // nested w:sdt behind after its payload is removed.
        AssertExactTopLevelEnvelope(right, RevisionProcessor.AcceptRevisions(rendered));
        AssertExactTopLevelEnvelope(left, RevisionProcessor.RejectRevisions(rendered));
        AssertSchemaValid(rendered);
    }

    [Fact]
    public void Render_block_sdt_in_table_cell_round_trips_exact_envelope_inside_one_table()
    {
        var left = IrTestDocuments.FromBodyXml(TableWithCellSdt("old-cell-control", "old cell payload"));
        var right = IrTestDocuments.FromBodyXml(TableWithCellSdt("new-cell-control", "new cell payload"));

        var rendered = Render(left, right);
        var body = Body(rendered);
        var table = Assert.Single(body.Elements(W.tbl));
        var cell = Assert.Single(table.Elements(W.tr)).Elements(W.tc).Single();
        var renderedSdts = cell.Elements(W.sdt).ToList();
        Assert.Equal(2, renderedSdts.Count);

        var oldEnvelope = Assert.Single(renderedSdts.Where(s => Tag(s) == "old-cell-control"));
        var newEnvelope = Assert.Single(renderedSdts.Where(s => Tag(s) == "new-cell-control"));
        AssertEnvelopeRangeTopology(oldEnvelope, W.customXmlDelRangeStart, W.customXmlDelRangeEnd);
        AssertEnvelopeRangeTopology(newEnvelope, W.customXmlInsRangeStart, W.customXmlInsRangeEnd);
        AssertUniqueCustomXmlRangeStartIds(body);

        AssertExactCellEnvelope(right, RevisionProcessor.AcceptRevisions(rendered));
        AssertExactCellEnvelope(left, RevisionProcessor.RejectRevisions(rendered));
        AssertSchemaValid(rendered);
        AssertSchemaValid(RevisionProcessor.AcceptRevisions(rendered));
        AssertSchemaValid(RevisionProcessor.RejectRevisions(rendered));
    }

    private static WmlDocument Render(WmlDocument left, WmlDocument right)
    {
        var settings = new IrDiffSettings();
        var leftIr = IrReader.Read(left, ReadOptions);
        var rightIr = IrReader.Read(right, ReadOptions);
        var script = IrEditScriptBuilder.Build(leftIr, rightIr, settings);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    private static XElement Body(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        return new XElement(wordDocument.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }

    private static void AssertExactTopLevelEnvelope(WmlDocument expected, WmlDocument actual)
    {
        var expectedEnvelope = TopLevelEnvelope(expected, "expected");
        var actualEnvelope = TopLevelEnvelope(actual, "actual");
        Assert.Equal(expectedEnvelope.EnvelopeDigest, actualEnvelope.EnvelopeDigest);
    }

    private static void AssertExactCellEnvelope(WmlDocument expected, WmlDocument actual)
    {
        var expectedEnvelope = CellEnvelope(expected, "expected");
        var actualEnvelope = CellEnvelope(actual, "actual");
        Assert.Equal(expectedEnvelope.EnvelopeDigest, actualEnvelope.EnvelopeDigest);
    }

    private static IrSdtBlock TopLevelEnvelope(WmlDocument document, string documentRole)
    {
        var blocks = IrReader.Read(document, ReadOptions).Body.Blocks;
        var block = Assert.Single(blocks);
        Assert.True(block is IrSdtBlock,
            $"{documentRole} document did not retain its block SDT envelope; found {block.GetType().Name}. " +
            $"Body: {Body(document)}");
        return (IrSdtBlock)block;
    }

    private static IrSdtBlock CellEnvelope(WmlDocument document, string documentRole)
    {
        var table = Assert.IsType<IrTable>(Assert.Single(IrReader.Read(document, ReadOptions).Body.Blocks));
        var cell = Assert.Single(table.Rows).Cells.Single();
        var block = Assert.Single(cell.Blocks);
        Assert.True(block is IrSdtBlock,
            $"{documentRole} cell did not retain its block SDT envelope; found {block.GetType().Name}. " +
            $"Body: {Body(document)}");
        return (IrSdtBlock)block;
    }

    private static void AssertEnvelopeAndNestedTopology(XElement outer, XName startName, XName endName)
    {
        AssertEnvelopeRangeTopology(outer, startName, endName);
        var nested = Assert.Single(outer.Element(W.sdtContent)!.Elements(W.sdt));
        AssertEnvelopeRangeTopology(nested, startName, endName);
    }

    /// <summary>
    /// Verify the cross-boundary shape RevisionProcessor recognizes: range A covers the opening sdt tag,
    /// range B covers the closing tag. The start elements own author/date; their matching end elements
    /// repeat only the corresponding id.
    /// </summary>
    private static void AssertEnvelopeRangeTopology(XElement sdt, XName startName, XName endName)
    {
        var parent = Assert.IsType<XElement>(sdt.Parent);
        var siblings = parent.Elements().ToList();
        int index = siblings.IndexOf(sdt);
        Assert.InRange(index, 1, siblings.Count - 2);

        var before = siblings[index - 1];
        var after = siblings[index + 1];
        Assert.Equal(startName, before.Name);
        Assert.Equal(endName, after.Name);
        Assert.NotNull(before.Attribute(W.id));
        Assert.NotNull(before.Attribute(W.author));
        Assert.NotNull(before.Attribute(W.date));

        var content = sdt.Element(W.sdtContent);
        Assert.NotNull(content);
        var children = content!.Elements().ToList();
        Assert.True(children.Count >= 2, "sdtContent needs both cross-boundary range markers");
        var openingEnd = children[0];
        var closingStart = children[^1];
        Assert.Equal(endName, openingEnd.Name);
        Assert.Equal(startName, closingStart.Name);
        Assert.Equal((string?)before.Attribute(W.id), (string?)openingEnd.Attribute(W.id));
        Assert.Equal((string?)closingStart.Attribute(W.id), (string?)after.Attribute(W.id));
        Assert.NotNull(closingStart.Attribute(W.author));
        Assert.NotNull(closingStart.Attribute(W.date));
    }

    private static void AssertUniqueCustomXmlRangeStartIds(XElement body)
    {
        var starts = body.Descendants()
            .Where(e => e.Name == W.customXmlDelRangeStart || e.Name == W.customXmlInsRangeStart)
            .Select(e => (string?)e.Attribute(W.id))
            .ToList();
        Assert.NotEmpty(starts);
        Assert.DoesNotContain(starts, id => string.IsNullOrEmpty(id));
        Assert.Equal(starts.Count, starts.Distinct(StringComparer.Ordinal).Count());
    }

    private static void AssertSchemaValid(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(wordDocument)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
            .ToList();
        Assert.True(errors.Count == 0,
            "Unexpected schema errors:\n" + string.Join("\n", errors.Select(e => e.Description)));
    }

    private static string? Tag(XElement sdt) =>
        (string?)sdt.Element(W.sdtPr)?.Element(W.tag)?.Attribute(W.val);

    private static string BlockSdt(string tag, string text) =>
        "<w:sdt><w:sdtPr><w:tag w:val=\"" + tag + "\"/></w:sdtPr><w:sdtContent>" +
        "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>" +
        "</w:sdtContent></w:sdt>";

    private static string NestedBlockSdt(string outerTag, string innerTag, string text) =>
        "<w:sdt><w:sdtPr><w:tag w:val=\"" + outerTag + "\"/></w:sdtPr><w:sdtContent>" +
        BlockSdt(innerTag, text) +
        "</w:sdtContent></w:sdt>";

    private static string TableWithCellSdt(string tag, string text) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid>" +
        "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>" +
        BlockSdt(tag, text) +
        "</w:tc></w:tr></w:tbl>";
}
