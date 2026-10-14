using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffTableStyleMarginTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData("style")]
    [InlineData("direct")]
    [InlineData("table-override")]
    [InlineData("cell-override")]
    [InlineData("inherited")]
    [InlineData("conditional")]
    [InlineData("conditional-change")]
    [InlineData("conditional-change-override")]
    [InlineData("conditional-change-partial")]
    public void SharedTableMargins_AreReversibleAndRespectOverrides(string shape)
    {
        var left = Document(false, shape);
        var right = Document(true, shape);
        AssertClean(left, right);
        string rightTop = shape == "table-override" ? "90" : shape == "conditional-change-override" ? "60" :
            shape.StartsWith("conditional-change", StringComparison.Ordinal) ? "95" : shape == "conditional" ? "75" : "300";
        string rightBottom = shape is "cell-override" or "conditional-change-partial" ? "65" : "180";
        Assert.Equal(new[] { rightTop + "|120|" + rightBottom + "|120", "0|108|0|108" }, Margins(right));
        var comparison = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);
        Assert.Equal(Margins(right), Margins(accepted));
        Assert.Equal(Margins(left), Margins(rejected));
        Assert.Equal(Text(right), Text(accepted));
        Assert.Equal(Text(left), Text(rejected));
        AssertClean(comparison, accepted, rejected);
        if (shape != "direct")
            Assert.NotEmpty(Xml(comparison).Descendants(W + "tblPrChange"));
    }

    [Theory]
    [InlineData("body-row")]
    [InlineData("row-insert")]
    [InlineData("table-insert")]
    [InlineData("cell-edit")]
    [InlineData("nested")]
    public void MarginChanges_FollowTableContentPaths(string change)
    {
        var left = Document(false, "conditional-change");
        var right = Document(true, "conditional-change");
        if (change is "body-row" or "row-insert")
        {
            WmlDocument AddRow(WmlDocument doc) => Mutate(doc, root =>
            {
                var table = root.Descendants(W + "tbl").First();
                var row = new XElement(table.Element(W + "tr")!);
                row.Element(W + "trPr")?.Remove();
                row.Descendants(W + "t").Single().Value = "Copper lamp measures a wave.";
                table.Add(row);
            });
            if (change == "body-row") left = AddRow(left);
            right = AddRow(right);
        }
        if (change == "table-insert")
            right = Mutate(right, root =>
            {
                var table = new XElement(root.Descendants(W + "tbl").First());
                table.Descendants(W + "t").Single().Value = "Copper lamp measures a wave.";
                root.Element(W + "body")!.AddFirst(table, P("Indigo sign starts the next record."));
            });
        if (change == "cell-edit")
            right = Mutate(right, root => root.Descendants(W + "t").First().Value += " Copper lamps stay warm.");
        if (change == "nested")
        {
            WmlDocument Nest(WmlDocument doc) => Mutate(doc, root =>
            {
                var nested = new XElement(root.Descendants(W + "tbl").First());
                nested.Descendants(W + "t").Single().Value = "Copper lamp measures a wave.";
                root.Descendants(W + "tc").First().Add(nested, P("Indigo sign ends the inner record."));
            });
            left = Nest(left); right = Nest(right);
        }
        AssertEndpoints(left, right);
    }

    [Fact]
    public void RightOnlyChildOfChangedSharedStyle_InheritsRevisedMargins()
    {
        var left = Document(false, "style");
        var right = Document(true, "inherited");
        AssertEndpoints(left, right);
    }

    [Fact]
    public void LogicalMarginAxes_KeepTheirSchemaOrderAndValues()
    {
        WmlDocument Logical(WmlDocument doc, bool revised) => MutateStyles(doc, root =>
        {
            var margins = root.Descendants(W + "tblCellMar").Single();
            margins.Element(W + "left")!.Name = W + "start";
            margins.Element(W + "right")!.Name = W + "end";
            margins.Element(W + "start")!.SetAttributeValue(W + "w", revised ? 200 : 120);
            margins.Element(W + "end")!.SetAttributeValue(W + "w", revised ? 240 : 120);
        });
        AssertEndpoints(Logical(Document(false, "style"), false), Logical(Document(true, "style"), true));
    }

    [Fact]
    public void AssembledPartialMargins_RemainSchemaValid()
    {
        var document = Document(true, "conditional-change-partial");
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings { CreateHtmlConverterAnnotationAttributes = false });
        AssertClean(assembled);
        Assert.Equal("95|120|65|120", Margins(assembled).First());
    }

    [Fact]
    public void MarginChange_ReachesAnUnchangedHeaderTable()
    {
        WmlDocument WithHeader(WmlDocument document)
        {
            using var stream = new OpenXmlMemoryStreamDocument(document);
            using (var package = stream.GetWordprocessingDocument())
            {
                var main = package.MainDocumentPart!;
                var header = main.AddNewPart<HeaderPart>("rIdBeaconHeader");
                header.GetXDocument().Add(new XElement(W + "hdr", new XElement(main.GetXDocument().Descendants(W + "tbl").First())));
                header.PutXDocument();
                var body = main.GetXDocument().Root!.Element(W + "body")!;
                var section = body.Element(W + "sectPr");
                if (section is null) { section = new XElement(W + "sectPr"); body.Add(section); }
                section.AddFirst(new XElement(W + "headerReference", new XAttribute(W + "type", "default"),
                    new XAttribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id", "rIdBeaconHeader")));
                main.PutXDocument();
            }
            return stream.GetModifiedWmlDocument();
        }
        var left = WithHeader(Document(false, "conditional-change"));
        var right = WithHeader(Document(true, "conditional-change"));
        AssertEndpoints(left, right);
        var comparison = DocxCompare.Compare(left, right);
        Assert.Equal(Margins(right, header: true), Margins(RevisionProcessor.AcceptRevisions(comparison), header: true));
        Assert.Equal(Margins(left, header: true), Margins(RevisionProcessor.RejectRevisions(comparison), header: true));
    }

    [Fact]
    public void RemovedStyleMargins_RevertToBuiltInValues()
    {
        var left = Document(true, "style");
        var right = MutateStyles(Document(false, "style"), root => root.Descendants(W + "tblCellMar").Remove());
        AssertEndpoints(left, right);
    }

    [Fact]
    public void DefaultsAndMarginChanges_ComposeInTheSameComparison()
    {
        WmlDocument WithDefaults(WmlDocument document, bool revised) => MutateStyles(document, root =>
        {
            var font = revised ? "DejaVu Serif" : "DejaVu Sans";
            root.AddFirst(new XElement(W + "docDefaults",
                new XElement(W + "rPrDefault", new XElement(W + "rPr", new XElement(W + "rFonts",
                    new[] { "ascii", "hAnsi", "eastAsia", "cs" }.Select(slot => new XAttribute(W + slot, font))),
                    new XElement(W + "sz", new XAttribute(W + "val", revised ? 28 : 20)),
                    new XElement(W + "szCs", new XAttribute(W + "val", revised ? 28 : 20)))),
                new XElement(W + "pPrDefault", new XElement(W + "pPr", new XElement(W + "spacing",
                    new XAttribute(W + "before", 0), new XAttribute(W + "after", revised ? 360 : 0),
                    new XAttribute(W + "line", revised ? 360 : 240), new XAttribute(W + "lineRule", "auto"))))));
        });
        var left = WithDefaults(Document(false, "conditional-change"), false);
        var right = WithDefaults(Document(true, "conditional-change"), true);
        AssertEndpoints(left, right);
        var comparison = DocxCompare.Compare(left, right);
        Assert.Equal(Presentation(right), Presentation(RevisionProcessor.AcceptRevisions(comparison)));
        Assert.Equal(Presentation(left), Presentation(RevisionProcessor.RejectRevisions(comparison)));
    }

    private static string[] Presentation(WmlDocument document) => Xml(FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings()))
        .Descendants(W + "p").Select(p =>
        {
            var r = p.Descendants(W + "r").First().Element(W + "rPr")!;
            var spacing = p.Element(W + "pPr")?.Element(W + "spacing");
            return string.Join("|", new[] { "ascii", "hAnsi", "eastAsia", "cs" }.Select(slot => (string?)r.Element(W + "rFonts")?.Attribute(W + slot))
                .Concat(new[] { "sz", "szCs" }.Select(size => (string?)r.Element(W + size)?.Attribute(W + "val")))
                .Concat(new[] { "before", "after", "line", "lineRule" }.Select(axis => (string?)spacing?.Attribute(W + axis))));
        }).ToArray();

    private static void AssertEndpoints(WmlDocument left, WmlDocument right)
    {
        AssertClean(left, right);
        var comparison = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);
        Assert.Equal(Margins(right), Margins(accepted));
        Assert.Equal(Margins(left), Margins(rejected));
        Assert.Equal(Text(right), Text(accepted));
        Assert.Equal(Text(left), Text(rejected));
        AssertClean(comparison, accepted, rejected);
    }

    private static XElement P(string text) => new(W + "p", new XElement(W + "r", new XElement(W + "t", text)));

    private static WmlDocument Mutate(WmlDocument document, Action<XElement> mutate, bool styles = false)
    {
        using var stream = new OpenXmlMemoryStreamDocument(document);
        using (var package = stream.GetWordprocessingDocument())
        {
            OpenXmlPart part = styles ? package.MainDocumentPart!.StyleDefinitionsPart! : package.MainDocumentPart!;
            mutate(part.GetXDocument().Root!);
            part.PutXDocument();
        }
        return stream.GetModifiedWmlDocument();
    }

    private static WmlDocument MutateStyles(WmlDocument document, Action<XElement> mutate) => Mutate(document, mutate, styles: true);

    private static WmlDocument Document(bool revised, string shape)
    {
        string Margin(string axis, int value) => $"<w:{axis} w:w=\"{value}\" w:type=\"dxa\"/>";
        string M(int top, int bottom) => "<w:tblCellMar>" + Margin("top", top) + Margin("left", 120) + Margin("bottom", bottom) + Margin("right", 120) + "</w:tblCellMar>";
        int top = revised ? 300 : 0, bottom = revised ? 180 : 0;
        var styles = "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
            "<w:style w:type=\"table\" w:styleId=\"BeaconGrid\"><w:name w:val=\"Beacon Grid\"/><w:tblPr>" +
            M(shape == "direct" ? 0 : top, shape == "direct" ? 0 : bottom) + "</w:tblPr>";
        if (shape.StartsWith("conditional", StringComparison.Ordinal))
            styles += "<w:tblStylePr w:type=\"firstRow\"><w:tcPr><w:tcMar>" + Margin("top", revised && shape != "conditional" ? 95 : 75) + "</w:tcMar></w:tcPr></w:tblStylePr>";
        styles += "</w:style>";
        if (shape == "inherited")
            styles += "<w:style w:type=\"table\" w:styleId=\"BeaconChild\"><w:name w:val=\"Beacon Child\"/><w:basedOn w:val=\"BeaconGrid\"/></w:style>";
        string Table(string style, string sentence, bool affected)
        {
            var direct = affected && shape == "direct" ? M(top, bottom) : affected && shape == "table-override" ? "<w:tblCellMar>" + Margin("top", 90) + "</w:tblCellMar>" : "";
            var cell = affected && shape is "cell-override" or "conditional-change-partial" ? "<w:tcMar>" + Margin("bottom", 65) + "</w:tcMar>" : "";
            if (affected && shape == "conditional-change-override") cell = "<w:tcMar>" + Margin("top", 60) + "</w:tcMar>";
            return "<w:tbl><w:tblPr>" + (style.Length > 0 ? $"<w:tblStyle w:val=\"{style}\"/>" : "") +
                "<w:tblW w:w=\"4800\" w:type=\"dxa\"/>" + direct + "<w:tblLook w:firstRow=\"1\" w:noHBand=\"1\" w:noVBand=\"1\"/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"4800\"/></w:tblGrid><w:tr><w:trPr><w:cnfStyle w:val=\"100000000000\" w:firstRow=\"1\"/></w:trPr>" +
                "<w:tc><w:tcPr><w:tcW w:w=\"4800\" w:type=\"dxa\"/>" + cell + "</w:tcPr><w:p><w:r><w:t>" + sentence + "</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        }
        return IrTestDocuments.FromBodyAndStylesXml(Table(shape == "inherited" ? "BeaconChild" : "BeaconGrid", "Amber buoy counts six ripples.", true) +
            "<w:p><w:r><w:t>Olive flag separates the records.</w:t></w:r></w:p>" +
            Table("", "Lavender dock stays quiet.", false), styles);
    }

    private static string[] Margins(WmlDocument document, bool header = false)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings());
        return Xml(assembled, header).Descendants(W + "tc").Select(cell =>
        {
            var own = cell.Element(W + "tcPr")?.Element(W + "tcMar");
            var table = cell.Ancestors(W + "tbl").First().Element(W + "tblPr")?.Element(W + "tblCellMar");
            string Axis(string axis)
            {
                string logical = axis == "left" ? "start" : axis == "right" ? "end" : axis;
                return (string?)own?.Element(W + logical)?.Attribute(W + "w") ?? (string?)own?.Element(W + axis)?.Attribute(W + "w") ??
                    (string?)table?.Element(W + logical)?.Attribute(W + "w") ?? (string?)table?.Element(W + axis)?.Attribute(W + "w") ?? (axis is "top" or "bottom" ? "0" : "108");
            }
            return string.Join("|", new[] { Axis("top"), Axis("left"), Axis("bottom"), Axis("right") });
        }).ToArray();
    }

    private static XDocument Xml(WmlDocument document, bool header = false)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var input = header ? package.MainDocumentPart!.HeaderParts.Single().GetStream() : package.MainDocumentPart!.GetStream();
        return XDocument.Load(input);
    }

    private static string[] Text(WmlDocument document) => Xml(document).Descendants(W + "p")
        .Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value))).ToArray();

    private static void AssertClean(params WmlDocument[] documents)
    {
        foreach (var document in documents)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package).ToArray();
            Assert.True(errors.Length == 0, string.Join("\n", errors.Select(e => e.Description + " " + e.Node?.OuterXml)));
        }
    }
}
