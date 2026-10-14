#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A complex-field <c>HYPERLINK</c> (fldChar begin / instrText / separate / result / end) and a
/// <c>w:hyperlink</c> element are the SAME link to a reader — Word Compare treats them as equal
/// when target and display text match (its output over such a pair is revision-free). The IR must
/// canonicalize the field form to <see cref="F:IrHyperlink"/> so both forms hash and tokenize
/// identically; without it, display-identical paragraphs mismatch wholesale (the comments-redline
/// corpus family diffs an entire table of links as phantom delete+insert).
/// </summary>
public class DocxDiffHyperlinkCanonTests
{
    private const string Url = "https://example.com/track-changes";

    private static WmlDocument ElementFormDoc()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var rel = main.AddHyperlinkRelationship(new Uri(Url), true);
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before the link paragraph."))),
                new Paragraph(
                    new Hyperlink(new Run(new Text("Open source"))) { Id = rel.Id }),
                new Paragraph(new Run(new Text("After the link paragraph.")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("el.docx", stream.ToArray());
    }

    private static WmlDocument FieldFormDoc()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Before the link paragraph."))),
                new Paragraph(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" HYPERLINK \"" + Url + "\" \\h ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("Open source")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End })),
                new Paragraph(new Run(new Text("After the link paragraph.")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("fld.docx", stream.ToArray());
    }

    [Fact]
    public void ElementForm_vs_FieldForm_SameLink_ProducesNoRevisions()
    {
        var result = DocxDiff.Compare(ElementFormDoc(), FieldFormDoc());

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        var body = XDocument.Parse(reader.ReadToEnd()).Root!.Element(w + "body")!;
        var insTexts = body.Descendants(w + "ins").SelectMany(i => i.Descendants(w + "t")).Select(t => (string)t).ToList();
        var delTexts = body.Descendants(w + "del").SelectMany(d => d.Descendants(w + "delText")).Select(t => (string)t).ToList();
        Assert.Empty(insTexts);
        Assert.Empty(delTexts);
    }

    [Fact]
    public void FieldForm_TargetChange_IsStillAContentChange()
    {
        // Canonicalization must NOT erase target semantics: same display text, different URL is a
        // change (the lnk: suffix carries the target).
        static WmlDocument FieldDoc(string url)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                        new Run(new FieldCode(" HYPERLINK \"" + url + "\" \\h ") { Space = SpaceProcessingModeValues.Preserve }),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                        new Run(new Text("Open source")),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.End })),
                    new Paragraph(new Run(new Text("Anchor paragraph text here.")))));
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("fld.docx", stream.ToArray());
        }

        var result = DocxDiff.Compare(FieldDoc("https://example.com/a"), FieldDoc("https://example.com/b"));

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        var body = XDocument.Parse(reader.ReadToEnd()).Root!.Element(w + "body")!;
        Assert.True(
            body.Descendants(w + "ins").Any() || body.Descendants(w + "del").Any(),
            "a target-only change must still surface as a revision");
    }
}
