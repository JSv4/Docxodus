using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Tests.Ir.Diff;
using Xunit;
using Section = Docxodus.Tests.Ir.Diff.HeaderFooterFixtures.Section;

namespace Docxodus.Tests;

/// <summary>
/// Every section of a comparison's output that has a header or footer in either input keeps a
/// reference that resolves to a part carrying that story's content, tracked where it changed
/// (issue #843). Two input shapes are pinned: an original whose section properties reference its
/// header/footer parts through broken relationships (an id with no relationship, or a relationship
/// whose target part is missing from the package), and a document with many sections, each with its
/// own header and footer.
/// </summary>
public class DocxDiffHeaderFooterReferenceTests
{
    private static readonly XNamespace W = HeaderFooterFixtures.Wns;
    private static readonly XNamespace R = HeaderFooterFixtures.Rns;

    /// <summary>The raw XML of the part a section's effective reference of <paramref name="kind"/>
    /// resolves to (OOXML inheritance from earlier sections applies), or null when none resolves.</summary>
    private static string? ReferencedStoryXml(WmlDocument d, bool isHeader, int sectionIndex, string kind = "default")
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        XDocument mainXd;
        using (var s = main.GetStream(FileMode.Open, FileAccess.Read)) mainXd = XDocument.Load(s);
        var refName = W + (isHeader ? "headerReference" : "footerReference");
        var reference = mainXd.Descendants(W + "sectPr").Take(sectionIndex + 1).Reverse()
            .SelectMany(sectPr => sectPr.Elements(refName))
            .FirstOrDefault(e => ((string?)e.Attribute(W + "type") ?? "default") == kind);
        var relId = (string?)reference?.Attribute(R + "id");
        if (relId is null || !main.Parts.Any(p => p.RelationshipId == relId))
            return null;
        using var ps = main.GetPartById(relId).GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(ps).ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Visible (non-deleted) text of <see cref="ReferencedStoryXml"/>: its <c>w:t</c> values
    /// concatenated, so revision markup that splits a run does not hide the text.</summary>
    private static string? ReferencedStoryText(WmlDocument d, bool isHeader, int sectionIndex) =>
        ReferencedStoryXml(d, isHeader, sectionIndex) is { } xml
            ? string.Concat(XDocument.Parse(xml).Descendants(W + "t").Select(t => t.Value))
            : null;

    private static int SectionCount(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Count();
    }

    /// <summary>Every header/footer reference in the output names a relationship that exists.</summary>
    private static void AssertEveryReferenceResolves(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var relIds = main.Parts.Select(p => p.RelationshipId).ToHashSet();
        XDocument mainXd;
        using (var s = main.GetStream(FileMode.Open, FileAccess.Read)) mainXd = XDocument.Load(s);
        foreach (var reference in mainXd.Descendants().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference"))
            Assert.Contains((string?)reference.Attribute(R + "id"), relIds);
    }

    private static WmlDocument OneSection(string body, string relId, string? header, string? footer)
    {
        var hp = new Dictionary<string, string[]>();
        var fp = new Dictionary<string, string[]>();
        if (header is not null) hp[relId + "H"] = new[] { header };
        if (footer is not null) fp[relId + "F"] = new[] { footer };
        return HeaderFooterFixtures.Build(
            new[] { new Section(new[] { body }, Headers: new[] { ("default", relId + "H") }, Footers: new[] { ("default", relId + "F") }) },
            hp, fp);
    }

    private static WmlDocument WithoutZipEntries(WmlDocument d, params string[] entries)
    {
        using var ms = new MemoryStream();
        ms.Write(d.DocumentByteArray);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            foreach (var entry in entries)
                zip.GetEntry(entry)!.Delete();
        return new WmlDocument("broken.docx", ms.ToArray());
    }

    private static void AssertInsertedStory(WmlDocument output, bool isHeader, int section, string text)
    {
        Assert.Equal(text, ReferencedStoryText(output, isHeader, section));
        Assert.Contains("<w:ins ", ReferencedStoryXml(output, isHeader, section));
    }

    [Fact]
    public void OriginalReferencesMissingRelationshipIds_OutputKeepsRevisedStoriesAsInsertions()
    {
        // The original's sectPr names rIdGoneH / rIdGoneF, which no relationship defines.
        var original = OneSection("Body text one.", "rIdGone", header: null, footer: null);
        var revised = OneSection("Body text one, revised.", "rIdOk", "Revised header", "Revised footer");

        var output = DocxCompare.Compare(original, revised);

        AssertEveryReferenceResolves(output);
        AssertInsertedStory(output, isHeader: true, 0, "Revised header");
        AssertInsertedStory(output, isHeader: false, 0, "Revised footer");
    }

    [Fact]
    public void OriginalReferencesRelationshipsWhoseTargetPartsAreMissing_OutputKeepsRevisedStories()
    {
        // The relationships exist, but the header and footer parts they target are gone from the zip.
        var original = WithoutZipEntries(
            OneSection("Body text one.", "rIdA", "Original header", "Original footer"),
            "word/header1.xml", "word/footer1.xml");
        var revised = OneSection("Body text one, revised.", "rIdA", "Revised header", "Revised footer");

        var output = DocxCompare.Compare(original, revised);

        AssertEveryReferenceResolves(output);
        AssertInsertedStory(output, isHeader: true, 0, "Revised header");
        AssertInsertedStory(output, isHeader: false, 0, "Revised footer");
    }

    [Fact]
    public void BrokenReferenceInOneSectionOfMany_EverySectionKeepsItsFooter()
    {
        const int sections = 4;
        Section[] Sections(bool breakSectionTwo) => Enumerable.Range(0, sections)
            .Select(i => new Section(new[] { $"Section {i} body." },
                Footers: new[] { ("default", breakSectionTwo && i == 2 ? "rIdGone" : $"rIdF{i}") }))
            .ToArray();
        var footers = Enumerable.Range(0, sections).ToDictionary(i => $"rIdF{i}", i => new[] { $"Footer {i}" });

        var output = DocxCompare.Compare(
            HeaderFooterFixtures.Build(Sections(true), null, footers),
            HeaderFooterFixtures.Build(Sections(false), null, footers));

        AssertEveryReferenceResolves(output);
        for (int i = 0; i < sections; i++)
            Assert.Equal($"Footer {i}", ReferencedStoryText(output, isHeader: false, i));
        Assert.Contains("<w:ins ", ReferencedStoryXml(output, isHeader: false, 2));
    }

    [Fact]
    public void ManySectionsWithOwnHeadersAndFooters_KeepEveryPart_AndTrackOnlyTheEditedFooter()
    {
        const int sections = 8;
        Section[] Sections(string edit) => Enumerable.Range(0, sections)
            .Select(i => new Section(new[] { $"Section {i} body{(i == 3 ? edit : "")}." },
                Headers: new[] { ("default", $"rIdH{i}") },
                Footers: new[] { ("default", $"rIdF{i}") }))
            .ToArray();
        Dictionary<string, string[]> Stories(string prefix, string edit) => Enumerable.Range(0, sections)
            .ToDictionary(i => $"rId{prefix}{i}", i => new[] { $"{prefix} story {i}{(i == 5 ? edit : "")}" });

        var original = HeaderFooterFixtures.Build(Sections(""), Stories("H", ""), Stories("F", ""));
        var revised = HeaderFooterFixtures.Build(Sections(" edited"), Stories("H", ""), Stories("F", " revised"));

        var output = DocxCompare.Compare(original, revised);

        AssertEveryReferenceResolves(output);
        Assert.Equal(sections, SectionCount(output));
        Assert.Equal(sections * 2, HeaderFooterFixtures.StoryPartsXml(output).Count);
        for (int i = 0; i < sections; i++)
        {
            Assert.Equal($"H story {i}", ReferencedStoryText(output, isHeader: true, i));
            Assert.Equal($"F story {i}{(i == 5 ? " revised" : "")}", ReferencedStoryText(output, isHeader: false, i));
            Assert.DoesNotContain("<w:ins ", ReferencedStoryXml(output, isHeader: true, i));
            Assert.Equal(i == 5, ReferencedStoryXml(output, isHeader: false, i)!.Contains("<w:ins "));
        }
    }

    [Fact]
    public void SectionInserted_EverySectionKeepsItsOwnFooter()
    {
        Section[] Sections(int count) => Enumerable.Range(0, count)
            .Select(i => new Section(new[] { $"Section {i} body text here.", $"Closing line {i}." },
                Footers: new[] { ("default", $"rIdF{i}") }))
            .ToArray();
        Dictionary<string, string[]> Footers(int count) =>
            Enumerable.Range(0, count).ToDictionary(i => $"rIdF{i}", i => new[] { $"Footer {i}" });

        var output = DocxCompare.Compare(
            HeaderFooterFixtures.Build(Sections(4), null, Footers(4)),
            HeaderFooterFixtures.Build(Sections(5), null, Footers(5)));

        AssertEveryReferenceResolves(output);
        for (int i = 0; i < 5; i++)
            Assert.Equal($"Footer {i}", ReferencedStoryText(output, isHeader: false, i));
    }
}
