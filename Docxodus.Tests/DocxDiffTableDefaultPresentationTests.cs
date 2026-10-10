using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffTableDefaultPresentationTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void DocumentDefaultChanges_RespectBodyAndCellPrecedence(bool table, bool directOverrides, bool conditionalStyle)
    {
        var left = Document(false, table, directOverrides, conditionalStyle);
        var right = Document(true, table, directOverrides, conditionalStyle);
        AssertClean(left, right);
        var comparison = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);
        Assert.Equal(Effective(right), Effective(accepted));
        Assert.Equal(Effective(left), Effective(rejected));
        Assert.Equal(Text(right), Text(accepted));
        Assert.Equal(Text(left), Text(rejected));
        AssertClean(left, right, comparison, accepted, rejected);
        if (table)
        {
            Assert.Equal(2, Xml(comparison).Descendants(W + "pPrChange").Count());
            Assert.Equal(2, Xml(comparison).Descendants(W + "r").Elements(W + "rPr").Elements(W + "rPrChange").Count());
        }
    }

    [Theory]
    [InlineData("body-edit")]
    [InlineData("cell-edit")]
    [InlineData("insert-delete")]
    [InlineData("split")]
    [InlineData("merge")]
    [InlineData("row-insert")]
    public void DefaultChanges_FollowExistingContentEmissionPaths(string change)
    {
        var left = Document(false, true, false, true);
        var right = Document(true, true, false, true);
        if (change == "body-edit" || change == "cell-edit")
            right = Mutate(right, root => root.Descendants(W + "p").ElementAt(change == "body-edit" ? 0 : 1)
                .Descendants(W + "t").Single().Value += " Violet gauges stay bright.");
        if (change == "insert-delete")
        {
            left = Mutate(left, root => root.Element(W + "body")!.AddFirst(P("Copper window closes early.")));
            right = Mutate(right, root => root.Element(W + "body")!.AddFirst(P("Violet orchard opens late.")));
        }
        if (change == "split" || change == "merge")
        {
            var one = "White lighthouse logs a gust. Violet gauges stay bright.";
            left = Mutate(left, root => root.Descendants(W + "p").First().Descendants(W + "t").Single().Value = one);
            right = Mutate(right, root =>
            {
                var first = root.Descendants(W + "p").First();
                first.Descendants(W + "t").Single().Value = "White lighthouse logs a gust. ";
                first.AddAfterSelf(P("Violet gauges stay bright."));
            });
            if (change == "merge") (left, right) = (right, left);
        }
        if (change == "row-insert")
            right = Mutate(right, root =>
            {
                var row = new XElement(root.Descendants(W + "tr").Single());
                row.Descendants(W + "t").Single().Value = "Violet buoy reports a ripple.";
                root.Descendants(W + "tbl").Single().Add(row);
            });
        AssertEndpoints(left, right);
    }

    [Fact]
    public void ReusedSnapshots_DoNotAcquireConsumerFormatting()
    {
        var left = Document(false, true, false, true);
        var right = Document(true, true, false, true);
        var l = DocxDiff.CreateSnapshot(left);
        var r = DocxDiff.CreateSnapshot(right);
        var expected = DocxDiff.Compare(left, right);
        var first = DocxDiff.CreateComparison(l, r).ToRedline();
        var second = DocxDiff.CreateComparison(l, r).ToRedline();
        Assert.Equal(expected.DocumentByteArray, first.DocumentByteArray);
        Assert.Equal(expected.DocumentByteArray, second.DocumentByteArray);
        Assert.Equal(right.DocumentByteArray, DocxDiff.CreateComparison(r, r).ToRedline().DocumentByteArray);
    }

    [Fact]
    public void EmptyParagraph_MarkFormattingFollowsDefaults()
    {
        var left = Mutate(Document(false, true, false, false), root => root.Element(W + "body")!.AddFirst(new XElement(W + "p")));
        var right = Mutate(Document(true, true, false, false), root => root.Element(W + "body")!.AddFirst(new XElement(W + "p")));
        AssertEndpoints(left, right);
    }

    [Fact]
    public void SharedStyleInheritance_StillOverridesOnlyItsOwnAxes()
    {
        WmlDocument Styled(WmlDocument document) => MutateStyles(Mutate(document, root =>
        {
            foreach (var p in root.Descendants(W + "p"))
            {
                p.AddFirst(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "BeaconBody"))));
                p.Descendants(W + "r").Single().AddFirst(new XElement(W + "rPr", new XElement(W + "rStyle", new XAttribute(W + "val", "BeaconWord"))));
            }
        }), styles =>
        {
            styles.Add(new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", "BeaconBase"),
                new XElement(W + "name", new XAttribute(W + "val", "Beacon Base")), new XElement(W + "basedOn", new XAttribute(W + "val", "Normal")),
                new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", "Liberation Sans")))));
            styles.Add(new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", "BeaconBody"),
                new XElement(W + "name", new XAttribute(W + "val", "Beacon Body")), new XElement(W + "basedOn", new XAttribute(W + "val", "BeaconBase")),
                new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "before", "60")))));
            styles.Add(new XElement(W + "style", new XAttribute(W + "type", "character"), new XAttribute(W + "styleId", "BeaconWord"),
                new XElement(W + "name", new XAttribute(W + "val", "Beacon Word")),
                new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "cs", "Liberation Mono")))));
        });
        AssertEndpoints(Styled(Document(false, true, false, true)), Styled(Document(true, true, false, true)));
    }

    [Fact]
    public void RemovedDefaultDeclarations_ResolveToBuiltInValues()
    {
        var left = Document(true, true, false, false);
        var right = MutateStyles(Document(false, true, false, false), styles => styles.Element(W + "docDefaults")!.Descendants(W + "rPr").Single().RemoveNodes());
        right = MutateStyles(right, styles => styles.Element(W + "docDefaults")!.Descendants(W + "pPr").Single().RemoveNodes());
        AssertEndpoints(left, right);
    }

    [Fact]
    public void UnchangedDefaults_DoNotCreateConsumerRevisions()
    {
        var document = Document(false, true, true, true);
        var comparison = DocxCompare.Compare(document, document);
        Assert.DoesNotContain(Xml(comparison).Descendants(), e => e.Name == W + "pPrChange" || e.Name == W + "rPrChange");
        AssertEndpoints(document, document);
    }

    private static XElement P(string text) => new(W + "p", new XElement(W + "r", new XElement(W + "t", text)));

    private static void AssertEndpoints(WmlDocument left, WmlDocument right)
    {
        AssertClean(left, right);
        var comparison = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);
        Assert.Equal(Effective(right), Effective(accepted));
        Assert.Equal(Effective(left), Effective(rejected));
        Assert.Equal(Text(right), Text(accepted));
        Assert.Equal(Text(left), Text(rejected));
        AssertClean(comparison, accepted, rejected);
        var revisions = Xml(comparison).Descendants().Where(e => e.Name == W + "pPrChange" || e.Name == W + "rPrChange").ToArray();
        Assert.Equal(revisions.Length, revisions.Select(e => (string?)e.Attribute(W + "id")).Distinct().Count());
    }

    private static WmlDocument Mutate(WmlDocument document, Action<XElement> mutate)
    {
        using var stream = new OpenXmlMemoryStreamDocument(document);
        using (var package = stream.GetWordprocessingDocument())
        {
            var xml = package.MainDocumentPart!.GetXDocument();
            mutate(xml.Root!);
            package.MainDocumentPart.PutXDocument();
        }
        return stream.GetModifiedWmlDocument();
    }

    private static WmlDocument MutateStyles(WmlDocument document, Action<XElement> mutate)
    {
        using var stream = new OpenXmlMemoryStreamDocument(document);
        using (var package = stream.GetWordprocessingDocument())
        {
            var xml = package.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument();
            mutate(xml.Root!);
            package.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
        }
        return stream.GetModifiedWmlDocument();
    }

    private static WmlDocument Document(bool revised, bool table, bool directOverrides, bool conditionalStyle)
    {
        var font = revised ? "DejaVu Serif" : "DejaVu Sans";
        var size = revised ? "28" : "20";
        var after = revised ? "360" : "0";
        var line = revised ? "360" : "240";
        string P(string text) => "<w:p>" + (directOverrides ? "<w:pPr><w:spacing w:after=\"80\"/></w:pPr>" : "") +
            "<w:r>" + (directOverrides ? "<w:rPr><w:rFonts w:ascii=\"Liberation Mono\"/><w:sz w:val=\"18\"/></w:rPr>" : "") +
            $"<w:t>{text}</w:t></w:r></w:p>";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr>" +
            $"<w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:eastAsia=\"{font}\" w:cs=\"{font}\"/>" +
            $"<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" +
            $"<w:spacing w:before=\"0\" w:after=\"{after}\" w:line=\"{line}\" w:lineRule=\"auto\"/>" +
            "</w:pPr></w:pPrDefault></w:docDefaults>" +
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>";
        if (conditionalStyle)
            styles += "<w:style w:type=\"table\" w:styleId=\"HarborGrid\"><w:name w:val=\"Harbor Grid\"/>" +
                "<w:tblStylePr w:type=\"firstRow\"><w:pPr><w:spacing w:after=\"120\"/></w:pPr>" +
                "<w:rPr><w:rFonts w:hAnsi=\"Liberation Serif\"/><w:sz w:val=\"24\"/></w:rPr></w:tblStylePr></w:style>";
        var body = P("White lighthouse logs a gust.");
        if (table)
            body += "<w:tbl><w:tblPr>" + (conditionalStyle ? "<w:tblStyle w:val=\"HarborGrid\"/>" : "") +
                "<w:tblW w:w=\"4800\" w:type=\"dxa\"/><w:tblLook w:firstRow=\"1\" w:noHBand=\"1\" w:noVBand=\"1\"/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"4800\"/></w:tblGrid><w:tr><w:trPr><w:cnfStyle w:val=\"100000000000\" w:firstRow=\"1\"/></w:trPr>" +
                "<w:tc><w:tcPr><w:tcW w:w=\"4800\" w:type=\"dxa\"/></w:tcPr>" + P("Mint compass records a bearing.") + "</w:tc></w:tr></w:tbl>";
        return IrTestDocuments.FromBodyAndStylesXml(body, styles);
    }

    private static string[] Effective(WmlDocument document)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings());
        var paragraphs = Xml(assembled).Descendants(W + "p");
        string? Value(XElement? props, string child, string attribute) => (string?)props?.Element(W + child)?.Attribute(W + attribute);
        return paragraphs.Select(p =>
        {
            var pPr = p.Element(W + "pPr");
            string RunKey(XElement? rPr) => string.Join("|", new[]
            {
                Value(rPr, "rFonts", "ascii") ?? "Times New Roman", Value(rPr, "rFonts", "hAnsi") ?? "Times New Roman",
                Value(rPr, "rFonts", "eastAsia") ?? "Times New Roman", Value(rPr, "rFonts", "cs") ?? "Times New Roman",
                Value(rPr, "sz", "val") ?? "20", Value(rPr, "szCs", "val") ?? "20",
            });
            return string.Join("|", new[]
            {
                Value(pPr, "spacing", "before") ?? "0", Value(pPr, "spacing", "after") ?? "0",
                Value(pPr, "spacing", "line") ?? "240", Value(pPr, "spacing", "lineRule") ?? "auto",
            }) + "¶" + RunKey(pPr?.Element(W + "rPr")) + "¶" + string.Join(";", p.Descendants(W + "r")
                .SelectMany(r => string.Concat(r.Descendants(W + "t").Select(t => t.Value)).Select(_ => RunKey(r.Element(W + "rPr")))));
        }).ToArray();
    }

    private static XDocument Xml(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var input = package.MainDocumentPart!.GetStream();
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
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
    }
}
