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
/// Word-parity rules for how DocxDiff imports RIGHT-side list numbering into the LEFT-based output
/// package. An imported right-side list instance must get its own FRESH cloned <c>w:abstractNum</c>
/// UNLESS the diff proves the right list is genuinely the SAME list as a surviving left list (an
/// aligned paragraph pair carries the left numId on the left side and the right numId on the right
/// side, and the definitions agree). Word's compare output materializes a distinct abstractNum per
/// imported foreign list instance; content-based deduplication onto a left abstractNum is wrong
/// because LibreOffice keys list COUNTERS by abstractNumId — two different list instances sharing
/// one abstractNum CONTINUE numbering across each other where Word's output RESTARTS. The flip
/// side: when an inserted item joins a list that survives from the left document, the shared
/// definition must be KEPT so the counter continues — blanket always-cloning would break that.
/// </summary>
public class DocxDiffNumberingImportTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A doc whose numbering part defines numId 1 → abstractNum 0 as a single-level
    /// decimal list, and whose paragraphs (one per <paramref name="texts"/> entry) are numbered
    /// with numId 1.</summary>
    private static WmlDocument DecimalListDoc(params string[] texts)
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
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." },
                        new StartNumberingValue { Val = 1 })
                    { LevelIndex = 0 })
                { AbstractNumberId = 0 },
                new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
            doc.Save();
        }
        return new WmlDocument("d.docx", stream.ToArray());
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

    private static string? NumIdOf(XElement p) =>
        (string?)p.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")?.Attribute(W + "val");

    private static bool IsInserted(XElement p) =>
        p.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "ins") is not null ||
        (p.Elements(W + "ins").Any() && !p.Elements(W + "r").Any() && !p.Elements(W + "del").Any());

    private static bool IsDeleted(XElement p) =>
        p.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "del") is not null;

    /// <summary>The abstractNumId a num definition points at, by numId.</summary>
    private static string? AbstractRefOf(XDocument numbering, string numId) =>
        (string?)numbering.Root!.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Attribute(W + "numId") == numId)?
            .Element(W + "abstractNumId")?.Attribute(W + "val");

    /// <summary>Visible body paragraph texts of a revision-free document, in order, blanks dropped.</summary>
    private static string[] ParagraphTexts(WmlDocument doc)
    {
        using var s = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        var main = XDocument.Parse(reader.ReadToEnd());
        return main.Descendants(W + "p")
            .Select(p => string.Concat(p.Descendants(W + "t").Select(t => (string)t)))
            .Where(t => t.Length > 0)
            .ToArray();
    }

    [Fact]
    public void UnrelatedDocs_ContentEqualCollidingLists_ImportedListGetsFreshAbstractNum()
    {
        // Two UNRELATED documents, each with a single-level decimal list under the SAME numId 1
        // and CONTENT-EQUAL abstractNum definitions. No paragraph survives from left to right, so
        // the right list is NOT the left list — the imported instance must get its own fresh
        // cloned abstractNum. Sharing the left abstractNum would make LibreOffice CONTINUE the
        // counter from the deleted left items into the inserted right items ("1. a" … "2. x")
        // where Word's compare output RESTARTS ("1. a" … "1. x").
        var left = DecimalListDoc("alpha bravo charlie", "delta echo foxtrot");
        var right = DecimalListDoc("neun zehn elf", "zwoelf dreizehn vierzehn");

        var result = DocxDiff.Compare(left, right);

        var (main, numbering) = OpenParts(result);
        var nums = numbering.Root!.Elements(W + "num").ToList();
        Assert.Equal(2, nums.Count);

        // Every num maps to its own distinct abstractNum.
        var abstractRefs = nums
            .Select(n => (string?)n.Element(W + "abstractNumId")?.Attribute(W + "val"))
            .ToList();
        Assert.Equal(abstractRefs.Count, abstractRefs.Distinct().Count());

        // The deleted (left-sourced) list keeps the left identity ...
        var paras = main.Descendants(W + "p").ToList();
        var deletedIds = paras.Where(IsDeleted).Select(NumIdOf).Where(id => id is not null)
            .Distinct().ToList();
        var insertedIds = paras.Where(IsInserted).Select(NumIdOf).Where(id => id is not null)
            .Distinct().ToList();
        var deletedId = (string?)Assert.Single(deletedIds);
        var insertedId = (string?)Assert.Single(insertedIds);

        // ... and the imported (inserted, right-sourced) list resolves through its OWN
        // abstractNum, not the left's.
        Assert.NotEqual(deletedId, insertedId);
        var leftAbstract = AbstractRefOf(numbering, deletedId!);
        var importedAbstract = AbstractRefOf(numbering, insertedId!);
        Assert.NotNull(leftAbstract);
        Assert.NotNull(importedAbstract);
        Assert.NotEqual(leftAbstract, importedAbstract);
    }

    [Fact]
    public void UnrelatedDocs_ContentEqualCollidingLists_AcceptIsRightAndRejectIsLeft()
    {
        var left = DecimalListDoc("alpha bravo charlie", "delta echo foxtrot");
        var right = DecimalListDoc("neun zehn elf", "zwoelf dreizehn vierzehn");

        var result = DocxDiff.Compare(left, right);

        Assert.Equal(new[] { "neun zehn elf", "zwoelf dreizehn vierzehn" },
            ParagraphTexts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(new[] { "alpha bravo charlie", "delta echo foxtrot" },
            ParagraphTexts(RevisionProcessor.RejectRevisions(result)));
    }

    [Fact]
    public void InsertedItemJoiningSurvivingList_KeepsOneListIdentity_NoFreshAbstractNum()
    {
        // Base and next are versions of the SAME document: the next version appends a third item
        // to the SAME list (same numId). The inserted item joins the surviving left list — ONE
        // list identity, counter continues. No fresh abstractNum may be minted: blanket
        // always-cloning would split the inserted item onto its own list and restart its counter.
        var baseDoc = DecimalListDoc("alpha item one", "bravo item two");
        var nextDoc = DecimalListDoc("alpha item one", "bravo item two", "charlie item three");

        var result = DocxDiff.Compare(baseDoc, nextDoc);

        var (main, numbering) = OpenParts(result);
        var num = Assert.Single(numbering.Root!.Elements(W + "num"));
        Assert.Single(numbering.Root!.Elements(W + "abstractNum"));

        var paras = main.Descendants(W + "p").ToList();
        var inserted = paras.Where(IsInserted).ToList();
        Assert.NotEmpty(inserted);

        // The inserted item's numId resolves to the same num — and therefore the same
        // abstractNum — as its surviving siblings.
        var numId = (string?)num.Attribute(W + "numId");
        foreach (var p in paras)
            Assert.Equal(numId, NumIdOf(p));
    }

    [Fact]
    public void InsertedItemJoiningSurvivingList_AcceptIsRightAndRejectIsLeft()
    {
        var baseDoc = DecimalListDoc("alpha item one", "bravo item two");
        var nextDoc = DecimalListDoc("alpha item one", "bravo item two", "charlie item three");

        var result = DocxDiff.Compare(baseDoc, nextDoc);

        Assert.Equal(new[] { "alpha item one", "bravo item two", "charlie item three" },
            ParagraphTexts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(new[] { "alpha item one", "bravo item two" },
            ParagraphTexts(RevisionProcessor.RejectRevisions(result)));
    }
}
