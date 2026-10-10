using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffDefaultStyleIdentityTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void DifferentDefaultStyleIds_KeepBothFormattingEndpoints(bool explicitStyle, bool editText)
    {
        var left = Document("BaseText", false, explicitStyle);
        var right = Document("FreshText", true, explicitStyle, editText);
        var comparison = DocxCompare.Compare(left, right);
        AssertEndpoint(RevisionProcessor.AcceptRevisions(comparison), right);
        AssertEndpoint(RevisionProcessor.RejectRevisions(comparison), left);
        var paragraph = Xml(comparison, false).Descendants(W + "p").Single();
        Assert.Equal("FreshText", (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.NotNull(paragraph.Element(W + "pPr")?.Element(W + "pPrChange"));
        AssertClean(left, right, comparison, RevisionProcessor.AcceptRevisions(comparison), RevisionProcessor.RejectRevisions(comparison));
    }

    [Fact]
    public void SameDefaultStyleId_StillTracksFontAndSpacing()
    {
        var left = Document("BaseText", false, false);
        var right = Document("BaseText", true, false);
        var comparison = DocxCompare.Compare(left, right);
        AssertEndpoint(RevisionProcessor.AcceptRevisions(comparison), right);
        AssertEndpoint(RevisionProcessor.RejectRevisions(comparison), left);
        AssertClean(comparison);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SameNameCustomStyles_RemainSeparateFromTheDefaultTransition(bool existingRightId)
    {
        var custom = "<w:style w:type=\"paragraph\" w:styleId=\"Archive\"><w:name w:val=\"Normal\"/>" +
            "<w:pPr><w:spacing w:before=\"0\" w:after=\"80\" w:line=\"260\" w:lineRule=\"auto\"/></w:pPr>" +
            "<w:rPr><w:rFonts w:ascii=\"Liberation Mono\" w:hAnsi=\"Liberation Mono\" w:eastAsia=\"Liberation Mono\" w:cs=\"Liberation Mono\"/>" +
            "<w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/></w:rPr></w:style>";
        var customParagraph = "<w:p><w:pPr><w:pStyle w:val=\"Archive\"/></w:pPr><w:r><w:t>Violet antenna records a pulse.</w:t></w:r></w:p>";
        var left = Document("BaseText", false, false, extraStyles: custom +
            (existingRightId ? custom.Replace("Archive", "FreshText") : ""), extraBody: customParagraph);
        var right = Document("FreshText", true, false, extraStyles: custom, extraBody: customParagraph, styleName: "New default label");
        var comparison = DocxCompare.Compare(left, right);
        AssertEndpoint(RevisionProcessor.AcceptRevisions(comparison), right);
        AssertEndpoint(RevisionProcessor.RejectRevisions(comparison), left);
        var styles = Xml(comparison, true).Root!;
        Assert.Equal(3, styles.Elements(W + "style").Count());
        Assert.Equal("20", (string?)styles.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == "BaseText")
            .Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"));
        AssertClean(comparison, RevisionProcessor.AcceptRevisions(comparison), RevisionProcessor.RejectRevisions(comparison));
    }

    [Theory]
    [InlineData("insert")]
    [InlineData("delete")]
    [InlineData("split")]
    [InlineData("merge")]
    [InlineData("move")]
    [InlineData("table")]
    public void DefaultTransition_SurvivesStructuralEdits(string change)
    {
        string P(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";
        var whole = P("The copper buoy watches the blue harbor.");
        var pieces = P("The copper buoy watches ") + P("the blue harbor.");
        var extra = P("Indigo lantern records the wind.");
        var third = P("The pale lighthouse opens at dawn.");
        var leftBody = whole;
        var rightBody = whole;
        switch (change)
        {
            case "insert": rightBody += extra; break;
            case "delete": leftBody += extra; break;
            case "split": rightBody = pieces; break;
            case "merge": leftBody = pieces; break;
            case "move": leftBody += extra + third; rightBody = extra + third + whole; break;
            case "table":
                var table = "<w:tbl><w:tblPr><w:tblW w:w=\"4800\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4800\"/></w:tblGrid>" +
                    "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4800\" w:type=\"dxa\"/></w:tcPr>" + extra + "</w:tc></w:tr></w:tbl>";
                leftBody += table; rightBody += table; break;
        }
        var left = WithBody(Document("BaseText", false, false), leftBody);
        var right = WithBody(Document("FreshText", true, false), rightBody);
        var comparison = DocxCompare.Compare(left, right);
        AssertEndpoint(RevisionProcessor.AcceptRevisions(comparison), right);
        AssertEndpoint(RevisionProcessor.RejectRevisions(comparison), left);
        AssertClean(left, right, comparison, RevisionProcessor.AcceptRevisions(comparison), RevisionProcessor.RejectRevisions(comparison));
    }

    [Fact]
    public void DefaultTransition_RespectsDirectPropertyOverrides()
    {
        var body = "<w:p><w:pPr><w:spacing w:after=\"80\"/></w:pPr><w:r>" +
            "<w:rPr><w:rFonts w:ascii=\"Liberation Mono\"/><w:sz w:val=\"18\"/></w:rPr>" +
            "<w:t>Silver weather vane turns slowly.</w:t></w:r></w:p>";
        var left = WithBody(Document("BaseText", false, false), body);
        var right = WithBody(Document("FreshText", true, false), body);
        var comparison = DocxCompare.Compare(left, right);
        AssertEndpoint(RevisionProcessor.AcceptRevisions(comparison), right);
        AssertEndpoint(RevisionProcessor.RejectRevisions(comparison), left);
        AssertClean(comparison);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultTransition_ResolvesInheritedProperties(bool sharedParentId)
    {
        WmlDocument Inherit(WmlDocument document, string parentId)
        {
            var styles = Xml(document, true).Root!;
            var style = styles.Elements(W + "style").Single();
            var parent = new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", parentId),
                new XElement(W + "name", new XAttribute(W + "val", "Parent")), style.Element(W + "pPr"), style.Element(W + "rPr"));
            style.Elements(W + "pPr").Remove();
            style.Elements(W + "rPr").Remove();
            style.Add(new XElement(W + "basedOn", new XAttribute(W + "val", parentId)));
            styles.Add(parent);
            return IrTestDocuments.FromBodyAndStylesXml(string.Concat(Xml(document, false).Root!.Element(W + "body")!.Elements()), string.Concat(styles.Elements()));
        }
        var left = Inherit(Document("BaseText", false, false), sharedParentId ? "Parent" : "OldParent");
        var right = Inherit(Document("FreshText", true, false), sharedParentId ? "Parent" : "NewParent");
        var comparison = DocxCompare.Compare(left, right);
        AssertEndpoint(RevisionProcessor.AcceptRevisions(comparison), right);
        AssertEndpoint(RevisionProcessor.RejectRevisions(comparison), left);
        AssertClean(comparison, RevisionProcessor.AcceptRevisions(comparison), RevisionProcessor.RejectRevisions(comparison));
    }

    private static WmlDocument WithBody(WmlDocument document, string body) =>
        IrTestDocuments.FromBodyAndStylesXml(body, string.Concat(Xml(document, true).Root!.Elements()));

    internal static WmlDocument Document(string id, bool revised, bool explicitStyle, bool editText = false,
        string extraStyles = "", string extraBody = "", string styleName = "Normal")
    {
        var pStyle = explicitStyle ? $"<w:pStyle w:val=\"{id}\"/>" : "";
        var body = $"<w:p><w:pPr>{pStyle}</w:pPr><w:r><w:t>Jade beacon remains {(editText ? "dim" : "bright")}.</w:t></w:r></w:p>" + extraBody;
        var font = revised ? "DejaVu Serif" : "DejaVu Sans";
        var size = revised ? "28" : "20";
        var after = revised ? "360" : "0";
        var line = revised ? "360" : "240";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>" +
            $"<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"{id}\"><w:name w:val=\"{styleName}\"/>" +
            $"<w:pPr><w:spacing w:before=\"0\" w:after=\"{after}\" w:line=\"{line}\" w:lineRule=\"auto\"/></w:pPr>" +
            $"<w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:eastAsia=\"{font}\" w:cs=\"{font}\"/>" +
            $"<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr></w:style>" + extraStyles;
        return IrTestDocuments.FromBodyAndStylesXml(body, styles);
    }

    internal static XDocument Xml(WmlDocument document, bool styles)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var input = (styles ? (OpenXmlPart)package.MainDocumentPart!.StyleDefinitionsPart! : package.MainDocumentPart!).GetStream();
        return XDocument.Load(input);
    }

    internal static void AssertEndpoint(WmlDocument actual, WmlDocument expected)
    {
        var a = Xml(actual, false).Descendants(W + "p").ToArray();
        var e = Xml(expected, false).Descendants(W + "p").ToArray();
        Assert.Equal(e.Length, a.Length);
        for (int i = 0; i < a.Length; i++)
        {
            Assert.Equal(string.Concat(e[i].Descendants(W + "t").Select(t => t.Value)), string.Concat(a[i].Descendants(W + "t").Select(t => t.Value)));
            foreach (var axis in new[] { "rFonts/ascii", "rFonts/hAnsi", "rFonts/eastAsia", "rFonts/cs", "sz/val", "szCs/val", "spacing/before", "spacing/after", "spacing/line", "spacing/lineRule" })
                Assert.Equal(Effective(Xml(expected, true).Root!, e[i], axis), Effective(Xml(actual, true).Root!, a[i], axis));
        }
    }

    private static string? Effective(XElement styles, XElement paragraph, string axis)
    {
        var parts = axis.Split('/');
        bool run = parts[0] != "spacing";
        var property = run ? "rPr" : "pPr";
        string? value = (string?)styles.Element(W + "docDefaults")?.Element(W + (run ? "rPrDefault" : "pPrDefault"))?.Element(W + property)?.Element(W + parts[0])?.Attribute(W + parts[1]);
        var id = (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") ??
            (string?)styles.Elements(W + "style").Single(s => (string?)s.Attribute(W + "type") == "paragraph" && (string?)s.Attribute(W + "default") == "1").Attribute(W + "styleId");
        var chain = new System.Collections.Generic.List<XElement>();
        var seen = new System.Collections.Generic.HashSet<string>();
        while (id is not null && seen.Add(id))
        {
            var style = styles.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == id);
            chain.Add(style);
            id = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
        }
        chain.Reverse();
        foreach (var style in chain)
            value = (string?)style.Element(W + property)?.Element(W + parts[0])?.Attribute(W + parts[1]) ?? value;
        return (string?)(run ? paragraph.Descendants(W + "r").First().Element(W + "rPr") : paragraph.Element(W + "pPr"))?.Element(W + parts[0])?.Attribute(W + parts[1]) ?? value;
    }

    internal static void AssertClean(params WmlDocument[] documents)
    {
        foreach (var document in documents)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
    }
}
