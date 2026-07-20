#nullable enable

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
/// The redline output's surviving body is RIGHT-sourced (equal/inserted/modified blocks emit the
/// right document's XML), but its numbering part is seeded from the LEFT. When both sides define
/// the same <c>w:numId</c> with different content, the copy renumbers the right's definition to a
/// fresh id — and every surviving reference must be REBOUND to that fresh id, or the right's lists
/// silently render with the left's format (decimal→bullet swaps). Deleted (left-sourced)
/// paragraphs — the ones whose paragraph mark carries <c>w:del</c> — keep the left id untouched.
/// </summary>
public class DocxDiffNumberingRemapTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A doc whose numbering part defines numId 1 → abstractNum 0 with the given format,
    /// and whose paragraphs (one per <paramref name="texts"/> entry) are numbered with numId 1.</summary>
    private static WmlDocument NumberedDoc(string numFmt, params string[] texts)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var text in texts)
            {
                body.Append(new Paragraph(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 1 })),
                    new Run(new Text(text))));
            }
            mainPart.Document = new Document(body);
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = new EnumValue<NumberFormatValues>(
                            numFmt == "bullet" ? NumberFormatValues.Bullet : NumberFormatValues.Decimal) },
                        new LevelText { Val = numFmt == "bullet" ? "•" : "%1." },
                        new StartNumberingValue { Val = 1 })
                    { LevelIndex = 0 })
                { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            doc.Save();
        }
        return new WmlDocument("d.docx", stream.ToArray());
    }

    /// <summary>Add a right-only paragraph style which owns the numbering reference itself, rather
    /// than putting <c>w:numPr</c> directly on the paragraph.</summary>
    private static WmlDocument AddStyleOwnedNumberedParagraph(WmlDocument source, string styleId, string text)
    {
        using var streamDoc = new OpenXmlMemoryStreamDocument(source);
        using (var doc = streamDoc.GetWordprocessingDocument())
        {
            var mainPart = doc.MainDocumentPart!;
            mainPart.Document!.Body!.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
                new Run(new Text(text))));
            mainPart.StyleDefinitionsPart!.Styles!.Append(new Style(
                new StyleName { Val = styleId },
                new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 })))
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
            });
            doc.Save();
        }
        return streamDoc.GetModifiedWmlDocument();
    }

    private static (XDocument Main, XDocument Numbering) OpenParts(WmlDocument result)
    {
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        var main = wdoc.MainDocumentPart!;
        using var mr = new StreamReader(main.GetStream());
        using var nr = new StreamReader(main.NumberingDefinitionsPart!.GetStream());
        return (XDocument.Parse(mr.ReadToEnd()), XDocument.Parse(nr.ReadToEnd()));
    }

    private static XDocument OpenStyles(WmlDocument result)
    {
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.StyleDefinitionsPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }

    /// <summary>Resolve the numFmt a paragraph's numId reference lands on in the output package.</summary>
    private static string? ResolveNumFmt(XDocument numbering, string numId)
    {
        var num = numbering.Root!.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Attribute(W + "numId") == numId);
        var abstractId = (string?)num?.Element(W + "abstractNumId")?.Attribute(W + "val");
        var abstractNum = numbering.Root!.Elements(W + "abstractNum")
            .FirstOrDefault(a => (string?)a.Attribute(W + "abstractNumId") == abstractId);
        return (string?)abstractNum?.Element(W + "lvl")?.Element(W + "numFmt")?.Attribute(W + "val");
    }

    [Fact]
    public void CollidingNumId_InsertedListRebindsToRightDefinition_EqualKeepsLeft()
    {
        // Same numId 1 on both sides, different definitions: left bullet, right decimal. Word does
        // not compare numbering definitions — EQUAL paragraphs keep the left's id (bullet), but a
        // paragraph INSERTED from the right must resolve to the right's renumbered definition
        // (decimal), not silently rebind onto the left's colliding id.
        var left = NumberedDoc("bullet", "alpha item one", "beta item two", "gamma item three");
        var right = NumberedDoc("decimal", "alpha item one", "beta item two", "gamma item three",
            "delta item four entirely new");

        var result = DocxDiff.Compare(left, right);

        var (main, numbering) = OpenParts(result);
        var paras = main.Descendants(W + "p").ToList();
        string? NumIdOf(XElement p) =>
            (string?)p.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")?.Attribute(W + "val");
        bool IsInserted(XElement p) =>
            p.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "ins") is not null ||
            (p.Elements(W + "ins").Any() && !p.Elements(W + "r").Any() && !p.Elements(W + "del").Any());
        var inserted = paras.Where(IsInserted).Select(NumIdOf).Where(id => id is not null).Distinct().ToList();
        var equal = paras.Where(p => !IsInserted(p)).Select(NumIdOf).Where(id => id is not null).Distinct().ToList();
        Assert.NotEmpty(inserted);
        Assert.NotEmpty(equal);
        foreach (var id in inserted)
            Assert.Equal("decimal", ResolveNumFmt(numbering, id!));
        foreach (var id in equal)
            Assert.Equal("bullet", ResolveNumFmt(numbering, id!));
    }

    [Fact]
    public void CollidingNumId_RightOnlyStyleOwnedListRebindsToRightDefinition()
    {
        // The right-only paragraph has no direct numPr: its pStyle owns numId 1.  Because the
        // left's #1 is a bullet and the right's #1 is decimal, the copied right definition gets a
        // fresh id.  The copied STYLE must follow that new id without changing left-owned styles.
        const string importedStyleId = "ImportedDecimalList";
        var left = NumberedDoc("bullet", "alpha item one", "beta item two", "gamma item three");
        var right = AddStyleOwnedNumberedParagraph(
            NumberedDoc("decimal", "alpha item one", "beta item two", "gamma item three"),
            importedStyleId, "delta style-owned item entirely new");

        var result = DocxDiff.Compare(left, right);

        var (main, numbering) = OpenParts(result);
        var styles = OpenStyles(result);
        var styledParagraph = Assert.Single(main.Descendants(W + "p").Where(p =>
            (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == importedStyleId));
        Assert.Null(styledParagraph.Element(W + "pPr")?.Element(W + "numPr"));

        var importedStyle = Assert.Single(styles.Root!.Elements(W + "style").Where(style =>
            (string?)style.Attribute(W + "type") == "paragraph" &&
            (string?)style.Attribute(W + "styleId") == importedStyleId));
        var styleNumId = (string?)importedStyle.Element(W + "pPr")?.Element(W + "numPr")?
            .Element(W + "numId")?.Attribute(W + "val");
        Assert.NotNull(styleNumId);
        Assert.NotEqual("1", styleNumId);
        Assert.Equal("decimal", ResolveNumFmt(numbering, styleNumId!));
        Assert.Equal("bullet", ResolveNumFmt(numbering, "1"));
    }

    [Fact]
    public void CollidingNumId_DeletedParagraphKeepsLeftDefinition()
    {
        // Left has an extra numbered paragraph that the right deletes entirely; its deleted
        // rendering must keep resolving to the LEFT's bullet list.
        var left = NumberedDoc("bullet", "alpha item one", "beta item two", "vanishing entry gone");
        var right = NumberedDoc("decimal", "alpha item one", "beta item two");

        var result = DocxDiff.Compare(left, right);

        var (main, numbering) = OpenParts(result);
        var deleted = main.Descendants(W + "p")
            .Where(p => p.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "del") is not null)
            .Select(p => (string?)p.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")?.Attribute(W + "val"))
            .Where(id => id is not null)
            .ToList();
        Assert.NotEmpty(deleted);
        foreach (var id in deleted)
            Assert.Equal("bullet", ResolveNumFmt(numbering, id!));
    }

    [Fact]
    public void IdenticalDefinitions_AreNotDuplicated()
    {
        // Same numId, same definition on both sides — no rebind, no duplicate num entries.
        var left = NumberedDoc("decimal", "alpha item one", "beta item two");
        var right = NumberedDoc("decimal", "alpha item one", "beta item two changed");

        var result = DocxDiff.Compare(left, right);

        var (main, numbering) = OpenParts(result);
        Assert.Single(numbering.Root!.Elements(W + "num"));
        var ids = main.Descendants(W + "numId").Select(n => (string?)n.Attribute(W + "val")).Distinct().ToList();
        Assert.Equal(new[] { "1" }, ids);
    }
}
