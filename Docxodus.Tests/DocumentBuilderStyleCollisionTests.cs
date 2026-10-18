using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocumentBuilderStyleCollisionTests
{
    [Theory]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    [InlineData(false, false, false)]
    public void Build_KeepsEachSourcesStyleAndReferences(bool reverse, bool sameName, bool collide)
    {
        var first = StyledDocument("First", "Shared", "First style", reverse);
        var second = StyledDocument("Second", collide ? "Shared" : "Separate",
            sameName ? "First style" : "Second style", !reverse);
        AssertValid(first);
        AssertValid(second);

        var merged = DocumentBuilder.BuildDocument(new[] { new Source(first, false), new Source(second, false) }.ToList());

        AssertValid(merged);
        using var stream = new MemoryStream(merged.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart!.GetXDocument().Root!.Elements(W.style).ToList();
        Assert.Equal(styles.Count, styles.Select(s => (string?)s.Attribute(W.styleId)).Distinct().Count());
        foreach (var (text, character) in new[] { ("First", reverse), ("Second", !reverse) })
        {
            var paragraph = Assert.Single(main.GetXDocument().Descendants(W.p), p => p.Value == text);
            var reference = Assert.Single(paragraph.Descendants(character ? W.rStyle : W.pStyle));
            var style = Assert.Single(styles, s => (string?)s.Attribute(W.styleId) == (string?)reference.Attribute(W.val));
            Assert.Equal(character ? "character" : "paragraph", (string?)style.Attribute(W.type));
            Assert.Equal(character ? "C00000" : "0000C0", (string?)style.Element(W.rPr)?.Element(W.color)?.Attribute(W.val));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Build_RenamesInheritedAndStoryReferencesBeforeCopying(bool character)
    {
        var first = StyledDocument("First", "Shared", "First style", !character);
        var second = StyledDocument("Second", "Shared", "Second style", character);
        using var sourceStream = new MemoryStream();
        sourceStream.Write(second.DocumentByteArray);
        using (var package = WordprocessingDocument.Open(sourceStream, true))
        {
            var main = package.MainDocumentPart!;
            var styles = main.StyleDefinitionsPart!.GetXDocument();
            styles.Root!.Add(new XElement(W.style, new XAttribute(W.type, character ? "character" : "paragraph"),
                new XAttribute(W.styleId, "Shared1"), new XElement(W.name, new XAttribute(W.val, "Reserved"))));
            styles.Root.Add(new XElement(W.style, new XAttribute(W.type, character ? "character" : "paragraph"),
                new XAttribute(W.styleId, "Derived"), new XElement(W.name, new XAttribute(W.val, "Derived")),
                new XElement(W.basedOn, new XAttribute(W.val, "Shared"))));
            main.StyleDefinitionsPart.PutXDocument();
            var document = main.GetXDocument();
            var paragraph = document.Descendants(W.p).Single();
            var header = main.AddNewPart<HeaderPart>();
            header.PutXDocument(new XDocument(new XElement(W.hdr, new XElement(paragraph))));
            document.Descendants(W.sectPr).Single().AddFirst(new XElement(W.headerReference,
                new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(header))));
            var footnotes = main.AddNewPart<FootnotesPart>();
            footnotes.PutXDocument(new XDocument(new XElement(W.footnotes,
                new XElement(W.footnote, new XAttribute(W.id, 1), new XElement(paragraph)))));
            paragraph.Add(new XElement(W.r, new XElement(W.footnoteReference, new XAttribute(W.id, 1))));
            main.PutXDocument();
        }
        second = new WmlDocument("stories.docx", sourceStream.ToArray());
        AssertValid(second);

        var merged = DocumentBuilder.BuildDocument(new[] { new Source(first, true), new Source(second, true) }.ToList());

        AssertValid(merged);
        using var outputStream = new MemoryStream(merged.DocumentByteArray);
        using var output = WordprocessingDocument.Open(outputStream, false);
        var outputMain = output.MainDocumentPart!;
        var outputStyles = outputMain.StyleDefinitionsPart!.GetXDocument().Root!;
        var imported = Assert.Single(outputStyles.Elements(W.style), s => (string?)s.Element(W.name)?.Attribute(W.val) == "Second style");
        var id = (string)imported.Attribute(W.styleId)!;
        Assert.NotEqual("Shared", id);
        Assert.NotEqual("Shared1", id);
        Assert.Equal(id, (string?)outputStyles.Elements(W.style).Single(s => (string?)s.Attribute(W.styleId) == "Derived")
            .Element(W.basedOn)?.Attribute(W.val));
        var copiedHeaders = outputMain.HeaderParts.Where(p => p.GetXDocument().Descendants(W.t).Any(t => t.Value == "Second")).ToList();
        Assert.NotEmpty(copiedHeaders);
        foreach (var part in copiedHeaders.Cast<OpenXmlPart>().Append(outputMain.FootnotesPart!))
            Assert.Equal(id, (string?)Assert.Single(part.GetXDocument().Descendants(character ? W.rStyle : W.pStyle)).Attribute(W.val));
    }

    private static WmlDocument StyledDocument(string text, string id, string name, bool character)
    {
        var reference = character
            ? $"<w:rPr><w:rStyle w:val=\"{id}\"/></w:rPr>"
            : "";
        var paragraphProperties = character ? "" : $"<w:pPr><w:pStyle w:val=\"{id}\"/></w:pPr>";
        return IrTestDocuments.FromParts(
            $"<w:p>{paragraphProperties}<w:r>{reference}<w:t>{text}</w:t></w:r></w:p><w:sectPr/>",
            $"<w:style w:type=\"{(character ? "character" : "paragraph")}\" w:styleId=\"{id}\">" +
            $"<w:name w:val=\"{name}\"/><w:rPr><w:color w:val=\"{(character ? "C00000" : "0000C0")}\"/></w:rPr></w:style>");
    }

    private static void AssertValid(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package).Select(e => e.Description));
    }
}
