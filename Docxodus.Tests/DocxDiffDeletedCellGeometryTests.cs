using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffDeletedCellGeometryTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("widths_only")]
    [InlineData("widths_row_delete")]
    [InlineData("widths_row_insert")]
    [InlineData("merge_horizontal")]
    [InlineData("split_horizontal")]
    [InlineData("offset_shift")]
    [InlineData("offset_add_cell")]
    [InlineData("offset_remove_cell")]
    [InlineData("merge_vertical")]
    [InlineData("split_vertical")]
    [InlineData("merged_widths_row_insert")]
    [InlineData("merged_widths_row_delete")]
    [InlineData("merge_horizontal_stable_widths")]
    [InlineData("offset_remove_stable_widths")]
    public void StructuralTableChanges_PreserveBothEndpointGeometries(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles/Issue1071", name));
        var original = new WmlDocument(Path.Combine(directory, "left.docx"));
        var revised = new WmlDocument(Path.Combine(directory, "right.docx"));
        var comparison = DocxCompare.Compare(original, revised);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);

        foreach (var document in new[] { original, revised, comparison, accepted, rejected })
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
        Assert.Equal(Text(revised), Text(accepted));
        Assert.Equal(Text(original), Text(rejected));
        Assert.Equal(Geometry(original), Geometry(rejected));
        Assert.Equal(Geometry(revised), Geometry(accepted));
    }

    private static XDocument Body(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return new XDocument(package.MainDocumentPart!.GetXDocument());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void NativeDeletedCells_StillFillTheUnoccupiedGridColumns(int deletedSpan)
    {
        // Native Word merge markup leaves the surviving cell unexpanded until acceptance.
        // Exercise logical spans rather than counting physical deleted cells.
        var body = "<w:tbl><w:tblPr/><w:tblGrid>" +
            string.Concat(Enumerable.Repeat("<w:gridCol w:w='1800'/>", deletedSpan + 1)) +
            "</w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Canal gauge</w:t></w:r></w:p></w:tc>" +
            $"<w:tc><w:tcPr><w:gridSpan w:val='{deletedSpan}'/><w:cellDel w:id='1' w:author='Surveyor'/></w:tcPr>" +
            "<w:p/></w:tc></w:tr></w:tbl>";
        var original = Ir.IrTestDocuments.FromBodyXml(body);
        var accepted = RevisionProcessor.AcceptRevisions(original);
        var cell = Assert.Single(Body(accepted).Descendants(W + "tc"));
        Assert.Equal((deletedSpan + 1).ToString(), (string?)cell.Element(W + "tcPr")?.Element(W + "gridSpan")?.Attribute(W + "val"));
        using var stream = new MemoryStream(accepted.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
    }

    private static string[] Text(WmlDocument document) => Body(document).Descendants(W + "p")
        .Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value))).ToArray();

    private static string Geometry(WmlDocument document) => JsonSerializer.Serialize(
        Body(document).Descendants(W + "tbl").Select(table => new
        {
            grid = table.Element(W + "tblGrid")!.Elements(W + "gridCol").Select(c => (string?)c.Attribute(W + "w")).ToArray(),
            rows = table.Elements(W + "tr").Select(row => new
            {
                before = Value(row.Element(W + "trPr"), "gridBefore", "val", "0"),
                after = Value(row.Element(W + "trPr"), "gridAfter", "val", "0"),
                beforeWidth = Value(row.Element(W + "trPr"), "wBefore", "w", "0"),
                afterWidth = Value(row.Element(W + "trPr"), "wAfter", "w", "0"),
                cells = row.Elements(W + "tc").Select(cell => new
                {
                    width = Value(cell.Element(W + "tcPr"), "tcW", "w", ""),
                    widthType = Value(cell.Element(W + "tcPr"), "tcW", "type", ""),
                    span = Value(cell.Element(W + "tcPr"), "gridSpan", "val", "1"),
                    merge = cell.Element(W + "tcPr")?.Element(W + "vMerge") is { } merge
                        ? (string?)merge.Attribute(W + "val") ?? "continue" : "none",
                    text = string.Concat(cell.Descendants(W + "t").Select(t => t.Value)),
                }).ToArray(),
            }).ToArray(),
        }).ToArray());

    private static string Value(XElement? properties, string element, string attribute, string fallback) =>
        (string?)properties?.Element(W + element)?.Attribute(W + attribute) ?? fallback;
}
