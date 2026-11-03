using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffInheritedPresentationTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("plain-control")]
    [InlineData("unused-numbering")]
    [InlineData("numbered-body")]
    [InlineData("changed-settings")]
    [InlineData("numbered-table")]
    public void DefaultsWithConcurrentPackageChanges_PreserveBothEndpoints(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles/Issue1070"));
        var original = new WmlDocument(Path.Combine(directory, name + "-original.docx"));
        var revised = new WmlDocument(Path.Combine(directory, name + "-revised.docx"));
        AssertEndpoints(original, revised);
    }

    [Theory]
    [InlineData("unused-numbering")]
    [InlineData("numbered-body")]
    [InlineData("changed-settings")]
    [InlineData("numbered-table")]
    public void DirectOverrides_KeepTheirPrecedenceOverChangedDefaults(string name)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles/Issue1070"));
        AssertEndpoints(WithOverrides(new WmlDocument(Path.Combine(directory, name + "-original.docx"))),
            WithOverrides(new WmlDocument(Path.Combine(directory, name + "-revised.docx"))));
    }

    [Theory]
    [InlineData("plain-control", "simple", false)]
    [InlineData("plain-control", "complex", false)]
    [InlineData("numbered-table", "simple", false)]
    [InlineData("numbered-table", "complex", false)]
    [InlineData("plain-control", "spanning", false)]
    [InlineData("numbered-table", "spanning", false)]
    [InlineData("plain-control", "simple", true)]
    [InlineData("numbered-table", "complex", true)]
    [InlineData("plain-control", "changed", false)]
    [InlineData("numbered-table", "changed", false)]
    public void Fields_DoNotSuppressChangedDefaults(string name, string shape, bool directOverrides)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles/Issue1070"));
        WmlDocument WithField(string suffix)
        {
            var input = new WmlDocument(Path.Combine(directory, name + suffix + ".docx"));
            using var stream = new OpenXmlMemoryStreamDocument(directOverrides ? WithOverrides(input) : input);
            using (var package = stream.GetWordprocessingDocument())
            {
                var main = package.MainDocumentPart!;
                var paragraph = main.GetXDocument().Descendants(W + "p").First();
                var result = new XElement(W + "r", new XElement(W + "t", "1"));
                if (directOverrides)
                    result.AddFirst(new XElement(W + "rPr",
                        new XElement(W + "rFonts", new XAttribute(W + "ascii", "Liberation Mono")),
                        new XElement(W + "sz", new XAttribute(W + "val", "18"))));
                string instruction = shape == "changed" && suffix == "-revised" ? " NUMPAGES " : " PAGE ";
                paragraph.Add(shape is "complex" or "spanning"
                    ? new XElement[] {
                        new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
                        new(W + "r", new XElement(W + "instrText", instruction)),
                        new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
                        result,
                        new(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))) }
                    : new[] { new XElement(W + "fldSimple", new XAttribute(W + "instr", instruction), result) });
                if (shape == "spanning")
                {
                    var end = paragraph.Elements(W + "r").Last();
                    end.Remove();
                    main.GetXDocument().Descendants(W + "p").Last().Add(end);
                }
                main.PutXDocument();
            }
            return stream.GetModifiedWmlDocument();
        }
        AssertEndpoints(WithField("-original"), WithField("-revised"));
    }

    private static WmlDocument WithOverrides(WmlDocument document)
    {
        using var stream = new OpenXmlMemoryStreamDocument(document);
        using (var package = stream.GetWordprocessingDocument())
        {
            var main = package.MainDocumentPart!;
            foreach (var paragraph in main.GetXDocument().Descendants(W + "p").Where(p => p.Descendants(W + "t").Any()))
            {
                paragraph.Element(W + "pPr")!.Add(new XElement(W + "spacing", new XAttribute(W + "after", "120")));
                foreach (var run in paragraph.Elements(W + "r"))
                    run.AddFirst(new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", "Liberation Mono"))));
            }
            main.PutXDocument();
        }
        return stream.GetModifiedWmlDocument();
    }

    private static void AssertEndpoints(WmlDocument original, WmlDocument revised)
    {
        var comparison = DocxCompare.Compare(original, revised);
        var accepted = RevisionProcessor.AcceptRevisions(comparison);
        var rejected = RevisionProcessor.RejectRevisions(comparison);

        Assert.Equal(Text(revised), Text(accepted));
        Assert.Equal(Text(original), Text(rejected));
        Assert.Equal(Fields(revised), Fields(accepted));
        Assert.Equal(Fields(original), Fields(rejected));
        // Check validity and rejection even when the acceptance assertion fails on the old code.
        foreach (var document in new[] { original, revised, comparison, accepted, rejected })
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
        }
        Assert.Equal(Presentation(original), Presentation(rejected));
        Assert.Equal(Presentation(revised), Presentation(accepted));
    }

    private static XDocument Body(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return new XDocument(package.MainDocumentPart!.GetXDocument());
    }

    private static string[] Text(WmlDocument document) => Body(document).Descendants(W + "p")
        .Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value))).ToArray();

    private static string[] Fields(WmlDocument document)
    {
        var body = Body(document);
        // A changed simple field is legally emitted as a complex field. Compare its instruction
        // and envelope at each paragraph boundary without requiring either serialization form.
        foreach (var field in body.Descendants(W + "fldSimple").Reverse().ToArray())
            field.ReplaceWith(new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
                new XElement(W + "r", new XElement(W + "instrText", (string?)field.Attribute(W + "instr"))),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
                field.Elements(),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
        return body.Descendants(W + "p").Select(p => string.Join("|", p.Descendants()
            .Where(e => e.Name == W + "fldChar" || e.Name == W + "instrText")
            .Select(e => e.Name.LocalName + ":" + ((string?)e.Attribute(W + "fldCharType") ?? e.Value)))).ToArray();
    }

    private static string Presentation(WmlDocument document)
    {
        var assembled = FormattingAssembler.AssembleFormatting(document, new FormattingAssemblerSettings
        {
            CreateHtmlConverterAnnotationAttributes = false,
            ClearStyles = false,
            RemoveStyleNamesFromParagraphAndRunProperties = false,
        });
        // Compare formatting per character so a legal run split cannot hide or invent a failure.
        return JsonSerializer.Serialize(Body(assembled).Descendants(W + "p")
            .Where(p => p.Descendants(W + "t").Any()).Select(p => new
            {
                before = Value(p.Element(W + "pPr"), "spacing", "before", "0"),
                after = Value(p.Element(W + "pPr"), "spacing", "after", "0"),
                line = Value(p.Element(W + "pPr"), "spacing", "line", "240"),
                lineRule = Value(p.Element(W + "pPr"), "spacing", "lineRule", "auto"),
                characters = p.Descendants(W + "r").SelectMany(r =>
                    string.Concat(r.Descendants(W + "t").Select(t => t.Value)).Select(character => new
                    {
                        character,
                        font = Value(r.Element(W + "rPr"), "rFonts", "ascii", ""),
                        size = Value(r.Element(W + "rPr"), "sz", "val", ""),
                    })).ToArray(),
            }).ToArray());
    }

    private static string Value(XElement? properties, string element, string attribute, string fallback) =>
        (string?)properties?.Element(W + element)?.Attribute(W + attribute) ?? fallback;
}
