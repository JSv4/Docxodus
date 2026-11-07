using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir.Diff;
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
        var reference = mainXd.Descendants(W + "sectPr").Where(s => s.Parent?.Name != W + "sectPrChange")
            .Take(sectionIndex + 1).Reverse()
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

    [Theory]
    [InlineData(false, "default", false)]
    [InlineData(true, "default", false)]
    [InlineData(false, "first", true)]
    [InlineData(true, "first", true)]
    [InlineData(false, "first", false)]
    [InlineData(true, "first", false)]
    [InlineData(false, "even", false)]
    [InlineData(true, "even", false)]
    public void RewrittenMergedSections_ShowInsertedStoriesBeforeDeletedBoundary(bool retainFirst, string kind, bool originalTitle)
    {
        var orchard = Enumerable.Range(1, 3).Select(n => $"Orchard note {n}: the pear rows were pruned before the first frost arrived this season.").ToArray();
        var cellar = Enumerable.Range(1, 3).Select(n => $"Cellar note {n}: the cider barrels were turned twice while the weather stayed mild.").ToArray();
        var weather = Enumerable.Range(1, 4).Select(n => $"Weather memo {n}: rainfall at the ridge station stayed under the seasonal average again.");
        var original = HeaderFooterFixtures.Build(new[] { new Section(orchard), new Section(cellar) },
            titlePg: originalTitle, evenAndOddHeaders: kind == "even");
        var revised = HeaderFooterFixtures.Build(new[] { new Section(
            (retainFirst ? orchard.Take(1).Concat(weather) : weather).ToArray(),
            Headers: new[] { (kind, "h") }, Footers: new[] { (kind, "f") }) },
            new Dictionary<string, string[]> { ["h"] = new[] { "RIDGE REGISTER" } },
            new Dictionary<string, string[]> { ["f"] = new[] { "Weather receipt" } },
            titlePg: kind == "first", evenAndOddHeaders: kind == "even");
        var result = DocxCompare.Compare(original, revised);
        AssertStoryEndpoints(original, revised, result, true, kind);
        AssertStoryEndpoints(original, revised, result, false, kind);
        foreach (var (header, text) in new[] { (true, "RIDGE REGISTER"), (false, "Weather receipt") })
        {
            var xml = ReferencedStoryXml(result, header, 0, kind);
            Assert.NotNull(xml);
            var story = XDocument.Parse(xml);
            Assert.Equal(text, string.Concat(story.Descendants(W + "t").Select(t => t.Value)));
            Assert.NotEmpty(story.Descendants(W + "ins"));
        }
    }

    private const string BoundaryText = "Orchard pears grew heavy before autumn arrived across the quiet mountain valley.";
    private const string MovedText = "The archive delivery includes numbered folders with blue stamps and sealed ribbons for northern offices.";
    private const string KeptText = "Stable inventory remains unchanged while maintenance personnel inspect emergency equipment and review safety instructions carefully.";
    private const string InsertedText = "Applicants must submit identity documents through certified postal channels every Friday morning.";

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, true)]
    [InlineData(false, false, true)]
    [InlineData(true, true, true)]
    [InlineData(false, true, true)]
    public void StoryVisibility_RespectsFormatTracking(bool track, bool unchangedStory, bool disableFirst)
    {
        string kind = disableFirst ? "default" : "first";
        var refs = unchangedStory ? new[] { (kind, "h") } : null;
        var original = HeaderFooterFixtures.Build(new[] { new Section(new[] { BoundaryText }, Headers: refs), new Section(new[] { MovedText }) },
            unchangedStory ? new Dictionary<string, string[]> { ["h"] = new[] { "FIRST PAGE STORY" } } : null, titlePg: disableFirst);
        var revised = HeaderFooterFixtures.Build(new[] { new Section(new[] { InsertedText }, Headers: new[] { (kind, "h") }) },
            new Dictionary<string, string[]> { ["h"] = new[] { "FIRST PAGE STORY" } }, titlePg: !disableFirst);
        var result = DocxCompare.Compare(original, revised, new DocxDiffSettings { TrackBlockFormatChanges = track });
        using var stream = new MemoryStream(result.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        var xml = package.MainDocumentPart!.GetXDocument();
        Assert.Equal(track, xml.Descendants(W + "sectPrChange").Any());
        Assert.All(xml.Descendants(W + "sectPr").Where(s => s.Parent?.Name != W + "sectPrChange"),
            s => Assert.Equal(!disableFirst, WordprocessingMLUtil.GetBoolProp(s, W + "titlePg")));
        if (track) AssertStoryEndpoints(original, revised, result, true, kind);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void MovedParagraph_BlocksStoryBindingAcrossDeletedBoundary(bool spanMove, bool crossParagraph)
    {
        const string divider = "Permanent divider remains here unchanged.";
        var original = HeaderFooterFixtures.Build(new[] {
            new Section(new[] { BoundaryText }),
            new Section(spanMove ? new[] { MovedText, divider, KeptText } : new[] { MovedText, KeptText }) });
        var revised = HeaderFooterFixtures.Build(new[] { new Section(
            spanMove ? new[] { InsertedText, divider, KeptText + " " + MovedText } : new[] { InsertedText, KeptText, MovedText },
            Headers: new[] { ("default", "h") }) },
            new Dictionary<string, string[]> { ["h"] = new[] { "REVISED STORY" } });
        var settings = new DocxDiffSettings { CrossParagraphTokenDiff = crossParagraph };
        var dataFirst = DocxDiff.CreateComparison(original, revised, settings);
        var json = dataFirst.GetEditScriptJson();
        var revisions = dataFirst.GetRevisions();
        var result = dataFirst.ToRedline();
        var redlineFirst = DocxDiff.CreateComparison(original, revised, settings);
        Assert.Equal(result.DocumentByteArray, redlineFirst.ToRedline().DocumentByteArray);
        Assert.Equal(json, redlineFirst.GetEditScriptJson());
        Assert.Equal(revisions.Select(r => r.ToString()), redlineFirst.GetRevisions().Select(r => r.ToString()));
        var script = IrEditScriptJson.Read(json);
        Assert.Equal(json, IrEditScriptJson.Write(script));
        Assert.Contains(script.Operations, op => spanMove ? op.RelocationGroupId is not null : op.MoveGroupId is not null);
        Assert.Null(ReferencedStoryXml(result, true, 0));
        AssertInsertedStory(result, true, 1, "REVISED STORY");
        AssertStoryEndpoints(original, revised, result, true, "default");
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void LaterReviewSection_UsesSourceFirstPageFlag(bool explicitOff, bool revisedTitle)
    {
        var original = HeaderFooterFixtures.Build(new[] {
            new Section(new[] { BoundaryText }), new Section(new[] { MovedText, KeptText,
                "Musicians practice difficult concert arrangements beside the old riverside theater." }),
            new Section(new[] { "Inventory balance remains recorded for the winter quarter." }) }, titlePg: explicitOff);
        if (explicitOff)
        {
            using var stream = new MemoryStream();
            stream.Write(original.DocumentByteArray);
            using (var package = WordprocessingDocument.Open(stream, true))
            {
                var main = package.MainDocumentPart!;
                foreach (var flag in main.GetXDocument().Descendants(W + "titlePg")) flag.SetAttributeValue(W + "val", "0");
                main.PutXDocument();
            }
            original = new WmlDocument("explicit-off.docx", stream.ToArray());
        }
        var revised = HeaderFooterFixtures.Build(new[] { new Section(new[] { InsertedText, KeptText, MovedText },
            Headers: new[] { ("first", "h") }) },
            new Dictionary<string, string[]> { ["h"] = new[] { "FIRST PAGE STORY" } }, titlePg: revisedTitle);
        var comparison = DocxDiff.CreateComparison(original, revised);
        var result = comparison.ToRedline();
        using var output = new MemoryStream(result.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(output, false);
        var sections = doc.MainDocumentPart!.GetXDocument().Descendants(W + "sectPr")
            .Where(s => s.Parent?.Name != W + "sectPrChange").ToArray();
        Assert.Equal(revisedTitle, WordprocessingMLUtil.GetBoolProp(sections[1], W + "titlePg"));
        Assert.All(sections, s => Assert.InRange(s.Elements(W + "sectPrChange").Count(), 0, 1));
        AssertStoryEndpoints(original, revised, result, true, "first");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReboundStoryMoves_ReserveBodyAndTokenRelocationIds(bool crossParagraph)
    {
        const string divider = "Permanent divider remains here unchanged.";
        const string fragment = "Visitors should photograph historic window patterns before entering the northern observatory wing.";
        const string first = "Curators catalogue artifacts carefully while volunteers prepare explanatory labels for incoming exhibitions.";
        const string last = "Engineers inspect roof drainage systems every spring and record necessary repairs in maintenance journals.";
        var refs = new[] { ("default", "h") };
        var original = HeaderFooterFixtures.Build(new[] { new Section(new[] { BoundaryText }, Headers: refs),
            new Section(new[] { MovedText, divider, KeptText, first + " " + fragment, last }, Headers: refs) },
            new Dictionary<string, string[]> { ["h"] = new[] { BoundaryText, MovedText, KeptText } });
        var revised = HeaderFooterFixtures.Build(new[] { new Section(
            new[] { InsertedText, divider, KeptText + " " + MovedText, first, last + " " + fragment }, Headers: refs) },
            new Dictionary<string, string[]> { ["h"] = new[] { MovedText, BoundaryText, KeptText } });
        var comparison = DocxDiff.CreateComparison(original, revised, new DocxDiffSettings { CrossParagraphTokenDiff = crossParagraph });
        var script = IrEditScriptJson.Read(comparison.GetEditScriptJson());
        var directIds = script.Operations.Where(op => op.RelocationGroupId is not null).Select(op => op.RelocationGroupId!.Value).ToArray();
        var tokenIds = script.Operations.Where(op => op.TokenDiff is not null)
            .SelectMany(op => op.TokenDiff!.Ops).Where(op => op.RelocationGroupId is not null)
            .Select(op => op.RelocationGroupId!.Value).ToArray();
        Assert.NotEmpty(directIds);
        Assert.Contains(tokenIds, id => !directIds.Contains(id));
        int maximum = directIds.Concat(tokenIds).Max();
        var headerMoves = script.HeaderFooterOps!.SelectMany(h => h.Ops).Where(op => op.MoveGroupId is not null).ToArray();
        Assert.NotEmpty(headerMoves);
        Assert.All(headerMoves, op => Assert.True(op.MoveGroupId > maximum));
        AssertStoryEndpoints(original, revised, comparison.ToRedline(), true, "default");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedBoundaryFallback_DoesNotInferVisibilityFromReviewOrdinals(bool legacyJson)
    {
        const string tail = "The permanent inventory continues in the final section after the revised boundary.";
        var original = HeaderFooterFixtures.Build(new[] { new Section(new[] { BoundaryText }), new Section(new[] { KeptText, tail }) });
        var refs = new[] { ("default", "h") };
        var revised = HeaderFooterFixtures.Build(new[] { new Section(new[] { KeptText, InsertedText }, Headers: refs),
            new Section(new[] { tail }, Headers: refs) },
            new Dictionary<string, string[]> { ["h"] = new[] { "SHARED DEFAULT STORY" } }, titlePg: true);
        using var stream = new MemoryStream();
        stream.Write(revised.DocumentByteArray);
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var main = package.MainDocumentPart!;
            main.GetXDocument().Root!.Element(W + "body")!.Element(W + "sectPr")!.Element(W + "titlePg")!.Remove();
            main.PutXDocument();
        }
        revised = new WmlDocument("mixed.docx", stream.ToArray());
        var result = DocxCompare.Compare(original, revised);
        if (legacyJson)
        {
            var script = IrEditScriptJson.Read(DocxDiff.GetEditScriptJson(original, revised));
            script = script with { HeaderFooterOps = Docxodus.Ir.IrNodeList.From(script.HeaderFooterOps!.Select(h => h with {
                ReferenceBindings = h.ReferenceBindings is null ? null : Docxodus.Ir.IrNodeList.From(
                    h.ReferenceBindings.Select(b => b with { RightSectionIndex = null })) })) };
            string json = IrEditScriptJson.Write(script);
            Assert.DoesNotContain("rightSectionIndex", json);
            result = IrMarkupRenderer.Render(IrEditScriptJson.Read(json), original, revised, new IrDiffSettings());
        }
        AssertStoryEndpoints(original, revised, result, true, "default");
    }

    [Theory]
    [InlineData(true, "default", false)]
    [InlineData(true, "default", true)]
    [InlineData(false, "default", false)]
    [InlineData(false, "default", true)]
    [InlineData(true, "first", false)]
    [InlineData(true, "first", true)]
    [InlineData(false, "first", false)]
    [InlineData(false, "first", true)]
    [InlineData(true, "even", false)]
    [InlineData(true, "even", true)]
    [InlineData(false, "even", false)]
    [InlineData(false, "even", true)]
    public void AddedOrRemovedOpeningSection_KeepsItsOwnStory(bool header, string kind, bool reverse)
    {
        Section Section(string[] text, string id) => new(text,
            Headers: header ? new[] { (kind, id) } : null,
            Footers: header ? null : new[] { (kind, id) });
        WmlDocument Document(Section[] sections, Dictionary<string, string[]> stories) =>
            HeaderFooterFixtures.Build(sections, header ? stories : null, header ? null : stories,
                titlePg: kind == "first", evenAndOddHeaders: kind == "even");
        var retained = new[] { "The maintenance summary stays in this section." };
        var original = Document(new[] { Section(retained, "old") },
            new() { ["old"] = new[] { "MAINTENANCE COPY" } });
        var revised = Document(new[] {
                Section(new[] { "Arrange the welcome badges by arrival time.", "Move spare chairs into the east alcove." }, "empty"),
                Section(retained, "kept") },
            new() { ["empty"] = new[] { "" }, ["kept"] = new[] { "MAINTENANCE COPY" } });

        if (reverse)
            (original, revised) = (revised, original);
        var result = DocxCompare.Compare(original, revised);
        var insertedStory = XDocument.Parse(ReferencedStoryXml(result, header, 0, kind)!);
        Assert.Equal("", string.Concat(insertedStory.Descendants(W + "delText").Select(t => t.Value)));
        Assert.Equal("", string.Concat(insertedStory.Descendants(W + "t").Select(t => t.Value)));
        Assert.Equal("MAINTENANCE COPY", string.Concat(XDocument.Parse(
            ReferencedStoryXml(result, header, 1, kind)!).Descendants(W + "t").Select(t => t.Value)));
        AssertStoryEndpoints(original, revised, result, header, kind);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AddedMiddleSection_PreservesLinkedAndEditedStories(bool reverse, bool edit)
    {
        var first = new Section(new[] { "The maintenance summary stays in this section." }, Footers: new[] { ("default", "a") });
        var retained = new Section(new[] { "The inspection ledger remains with the archivist." }, Footers: new[] { ("default", "b") });
        var linked = new Section(new[] { "The dispatch checklist follows the inspection ledger." });
        var added = new Section(new[] { "Arrange the welcome badges by arrival time.", "Move spare chairs into the east alcove." },
            Footers: new[] { ("default", "empty") });
        var original = HeaderFooterFixtures.Build(new[] { first, retained, linked }, footerParts:
            new Dictionary<string, string[]> { ["a"] = new[] { "MAINTENANCE COPY" }, ["b"] = new[] { "INSPECTION COPY" } });
        var revised = HeaderFooterFixtures.Build(new[] { first, added, retained, linked }, footerParts:
            new Dictionary<string, string[]> { ["a"] = new[] { "MAINTENANCE COPY" },
                ["b"] = new[] { edit ? "INSPECTION REVISED" : "INSPECTION COPY" }, ["empty"] = new[] { "" } });
        if (reverse)
            (original, revised) = (revised, original);
        var result = DocxCompare.Compare(original, revised);
        Assert.Equal("", XDocument.Parse(ReferencedStoryXml(result, false, 1)!).Root!.Value);
        Assert.Equal("MAINTENANCE COPY", ReferencedStoryText(result, false, 0));
        Assert.Equal(ReferencedStoryXml(result, false, 2), ReferencedStoryXml(result, false, 3));
        AssertStoryEndpoints(original, revised, result, false, "default");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void AddedBoundaryAfterRetainedText_DoesNotInventStoryRevisions(bool merge, bool reverse)
    {
        const string first = "The northern station records the morning temperature.";
        const string second = "The southern station records the evening temperature.";
        const string last = "The final receipt stays with the archivist.";
        var original = HeaderFooterFixtures.Build(new[] {
            new Section(merge ? new[] { first, second, last } : new[] { first, last },
                Footers: new[] { ("default", "a") }) }, footerParts: new Dictionary<string, string[]> { ["a"] = new[] { "MAINTENANCE COPY" } });
        var revised = HeaderFooterFixtures.Build(new[] {
            new Section(new[] { merge ? first + " " + second : first, "Arrange the welcome badges by arrival time." },
                Footers: new[] { ("default", "a") }),
            new Section(new[] { last }, Footers: new[] { ("default", "a") }) },
            footerParts: new Dictionary<string, string[]> { ["a"] = new[] { "MAINTENANCE COPY" } });
        if (reverse)
            (original, revised) = (revised, original);
        var result = DocxCompare.Compare(original, revised);
        var story = XDocument.Parse(ReferencedStoryXml(result, false, 0)!);
        Assert.Equal("", string.Concat(story.Descendants(W + "delText").Select(t => t.Value)));
        Assert.DoesNotContain(story.Descendants(W + "ins"),
            insertion => insertion.Descendants(W + "t").Any(t => t.Value.Length > 0));
        AssertStoryEndpoints(original, revised, result, false, "default");
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    public void RetainedParagraphSectionBoundary_RestoresBothTopologies(bool appendText, bool reverse, bool editText)
    {
        const string retained = "The maintenance summary stays in this section.";
        const string tail = "Arrange the welcome badges by arrival time.";
        var original = HeaderFooterFixtures.Build(new[] {
            new Section(appendText ? new[] { retained } : new[] { retained, tail },
                Footers: new[] { ("default", "a") }) },
            footerParts: new Dictionary<string, string[]> { ["a"] = new[] { "MAINTENANCE COPY" } });
        var revised = HeaderFooterFixtures.Build(new[] {
            new Section(new[] { editText ? retained.Replace("maintenance", "inspection") : retained },
                Footers: new[] { ("default", "a") }),
            new Section(new[] { tail }, Footers: new[] { ("default", "empty") }) },
            footerParts: new Dictionary<string, string[]> {
                ["a"] = new[] { "MAINTENANCE COPY" }, ["empty"] = new[] { "" } });
        if (reverse) (original, revised) = (revised, original);
        AssertStoryEndpoints(original, revised, DocxCompare.Compare(original, revised), false, "default");
    }

    [Fact]
    public void SectionBoundaryInsideControl_DoesNotSuppressLaterStoryEdits()
    {
        WmlDocument Document(string footer)
        {
            var input = HeaderFooterFixtures.Build(new[] {
                new Section(new[] { "The maintenance summary stays here." }, Footers: new[] { ("default", "a") }),
                new Section(new[] { "The inspection ledger remains with the archivist." }, Footers: new[] { ("default", "b") }) },
                footerParts: new Dictionary<string, string[]> { ["a"] = new[] { "MAINTENANCE COPY" }, ["b"] = new[] { footer } });
            using var stream = new OpenXmlMemoryStreamDocument(input);
            using (var package = stream.GetWordprocessingDocument())
            {
                var main = package.MainDocumentPart!;
                var first = main.GetXDocument().Root!.Element(W + "body")!.Element(W + "p")!;
                first.ReplaceWith(new XElement(W + "sdt", new XElement(W + "sdtPr"), new XElement(W + "sdtContent", new XElement(first))));
                main.PutXDocument();
            }
            return stream.GetModifiedWmlDocument();
        }
        var original = Document("INSPECTION COPY");
        var revised = Document("INSPECTION REVISED");
        AssertStoryEndpoints(original, revised, DocxCompare.Compare(original, revised), false, "default");
    }

    [Theory]
    [InlineData("whole-paragraph-replace")]
    [InlineData("body-delete-paragraph")]
    public void ReplacementBoundaryGaps_PreserveBothStoryEndpoints(string scenario)
    {
        var (left, right) = DocxodusDiffParityFixtures.DocxDiffScenarioFixtures.Build(scenario);
        var result = DocxCompare.Compare(left, right);
        AssertStoryEndpoints(left, right, result, true, "default");
        AssertStoryEndpoints(left, right, result, false, "default");
    }

    private static void AssertStoryEndpoints(WmlDocument original, WmlDocument revised, WmlDocument result, bool header, string kind)
    {
        string StoryText(WmlDocument document, int section) =>
            ReferencedStoryXml(document, header, section, kind) is { } xml ? XDocument.Parse(xml).Root!.Value : "";
        bool[] TitlePageFlags(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            return package.MainDocumentPart!.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>()
                .Select(section => section.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TitlePage>() is { } flag
                    && (flag.Val?.Value ?? true)).ToArray();
        }
        foreach (var (expected, resolved) in new[] {
            (original, RevisionProcessor.RejectRevisions(result)),
            (revised, RevisionProcessor.AcceptRevisions(result)) })
        {
            Assert.Equal(SectionCount(expected), SectionCount(resolved));
            Assert.Equal(TitlePageFlags(expected), TitlePageFlags(resolved));
            for (int i = 0; i < SectionCount(expected); i++)
                Assert.Equal(StoryText(expected, i), StoryText(resolved, i));
            AssertEveryReferenceResolves(resolved);
            Assert.Equal(BodyTexts(expected), BodyTexts(resolved));
        }
        foreach (var document in new[] { original, revised, result,
            RevisionProcessor.AcceptRevisions(result), RevisionProcessor.RejectRevisions(result) })
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
        }
    }

    private static string[] BodyTexts(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!.Elements(W + "p")
            .Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value))).ToArray();
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
