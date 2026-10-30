using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The XML depth door (issue #1389): every package read through the shared read step is scanned
/// once, and an XML part nested past its class limit — <see cref="XmlSemanticNormalizer.MaxStoryDepth"/>
/// for a WordprocessingML story or definition part, <see cref="XmlSemanticNormalizer.MaxElementDepth"/>
/// for every other part — is refused before any recursive walk can overflow the stack. Each part
/// class is swept through every public entry point on a 1 MB stack, at its limit and one level past it.
/// </summary>
public class DocxDepthLimitSweepTests
{
    private const int Story = XmlSemanticNormalizer.MaxStoryDepth;
    private const int Other = XmlSemanticNormalizer.MaxElementDepth;

    [Theory]
    [InlineData("body", Story)]
    [InlineData("header", Story)]
    [InlineData("footer", Story)]
    [InlineData("footnotes", Story)]
    [InlineData("endnotes", Story)]
    [InlineData("comments", Story)]
    [InlineData("styles", Story)]
    [InlineData("numbering", Story)]
    [InlineData("glossary", Story)]
    [InlineData("strict body with a deep shape", Story)]
    [InlineData("strict body the unfold makes well-formed", Story)]
    [InlineData("interleaved main part", Story)]
    [InlineData("main part typed application/xml", Story)]
    [InlineData("settings", Other)]
    [InlineData("commentsExtended", Other)]
    [InlineData("theme", Other)]
    [InlineData("fontTable", Other)]
    [InlineData("chart", Other)]
    [InlineData("custom XML data", Other)]
    [InlineData("custom XML data the compatibility normalizer rewrites", Other)]
    [InlineData("altChunk part", Other)]
    public void EveryEntryPoint_PassesAtTheLimit_AndRefusesOneLevelPast(string partClass, int limit)
    {
        // An entry point that rejects the shape itself (a shallow copy of it fails there too, as on
        // main) builds no tree from it; it is held only to not crashing.
        var shallow = DeepPackages.Build(partClass, 20, "Closing text");
        var rejectsShape = new HashSet<string>();
        OnThread(1024, () =>
        {
            foreach (var (name, run) in EntryPoints(shallow, DeepPackages.Build(partClass, 20, "Closing text, revised")))
            {
                try { if (run() is not null) rejectsShape.Add(name); }
                catch (Exception) { rejectsShape.Add(name); }
            }
        });

        var failures = new List<string>();
        foreach (var depth in new[] { limit, limit + 1 })
        {
            var bytes = DeepPackages.Build(partClass, depth, "Closing text");
            var changed = DeepPackages.Build(partClass, depth, "Closing text, revised");
            Assert.Equal(depth, DeepPackages.MaxDepth(bytes));
            bool over = depth > limit;
            OnThread(1024, () =>
            {
                foreach (var (name, run) in EntryPoints(bytes, changed))
                {
                    string? depthRefusal;
                    try { depthRefusal = run(); }
                    catch (Exception ex) when (ex.Message.Contains("xml_depth_limit_exceeded", StringComparison.Ordinal))
                    { depthRefusal = ex.GetType().Name; }
                    catch (Exception ex)
                    {
                        if (!rejectsShape.Contains(name))
                            failures.Add($"{partClass} at {depth}: {name} threw {ex.GetType().Name}: {ex.Message}");
                        continue;
                    }
                    if (over && depthRefusal is null && !rejectsShape.Contains(name))
                        failures.Add($"{partClass} at {depth}: {name} was not refused");
                    if (!over && depthRefusal is not null && !rejectsShape.Contains(name))
                        failures.Add($"{partClass} at {depth}: {name} was refused ({depthRefusal})");
                }
            });
        }
        Assert.True(failures.Count == 0, string.Join("\n", failures));
        // Strict markup is the byte read step's to unfold, so the bound counts the byte entry points.
        static bool FromBytes(string name) => !name.Contains("an open document", StringComparison.Ordinal);
        Assert.True(rejectsShape.Count(FromBytes) < 7, $"{partClass}: most entry points reject the shape itself: {string.Join(", ", rejectsShape)}");
    }

    /// <summary>Each entry point, returning a description of a depth refusal it reported in its
    /// result (or null), or throwing.</summary>
    internal static IEnumerable<(string Name, Func<string?> Run)> EntryPointsFor(byte[] bytes, byte[] changed) => EntryPoints(bytes, changed);

    private static IEnumerable<(string Name, Func<string?> Run)> EntryPoints(byte[] bytes, byte[] changed)
    {
        // A manifest describes the raw entries: one that is not XML until the strict unfold rewrites it
        // is reported malformed, which is a refusal to describe it too.
        static string? Finding(IEnumerable<string> codes) =>
            codes.FirstOrDefault(code => code.EndsWith("xml_depth_limit_exceeded", StringComparison.Ordinal)
                || code.EndsWith("malformed_xml", StringComparison.Ordinal));
        yield return ("session open + Project", () => { using var s = new DocxSession(bytes); s.Project(); return null; });
        yield return ("Save(false)", () => { using var s = new DocxSession(bytes); s.Save(persistAnchorIds: false); return null; });
        yield return ("Save(true)", () => { using var s = new DocxSession(bytes); s.Save(persistAnchorIds: true); return null; });
        yield return ("InsertParagraph + Undo", () =>
        {
            using var s = new DocxSession(bytes);
            var anchor = Assert.Single(s.Grep("Closing text")).EnclosingAnchor.Anchor.Id;
            Assert.True(s.InsertParagraph(anchor, Position.After, "Added").Success);
            Assert.True(s.Undo());
            return null;
        });
        yield return ("ListRevisions + AcceptAll", () =>
        {
            using var s = new DocxSession(bytes);
            s.ListRevisions();
            Assert.True(s.AcceptAllRevisions().Success);
            return null;
        });
        yield return ("DocxCompare", () => { DocxCompare.Compare(new WmlDocument("a.docx", bytes), new WmlDocument("b.docx", changed)); return null; });
        yield return ("HTML from WmlDocument", () => { WmlToHtmlConverter.ConvertToHtml(new WmlDocument("a.docx", bytes), new WmlToHtmlConverterSettings()); return null; });
        yield return ("HTML from bytes", () => { HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions()); return null; });
        // A package the caller opened never passes the byte read step: each public entry point that
        // takes an open document measures it.
        string? Opened(Action<WordprocessingDocument> run)
        {
            using var stream = new MemoryStream();
            stream.Write(bytes);
            using var document = WordprocessingDocument.Open(stream, true);
            // Only the byte read step unfolds strict markup; opened as is, such a main part is not XML,
            // and an entry point that never reads it builds no tree from it.
            using (var part = document.MainDocumentPart!.GetStream())
            using (var main = XmlReader.Create(part))
                while (main.Read()) { }
            run(document);
            return null;
        }
        yield return ("HTML from an open document", () => Opened(d => WmlToHtmlConverter.ConvertToHtml(d, new WmlToHtmlConverterSettings())));
        yield return ("markdown from an open document", () => Opened(d => WmlToMarkdownConverter.Convert(d, new WmlToMarkdownConverterSettings())));
        yield return ("accepting an open document", () => Opened(RevisionProcessor.AcceptRevisions));
        yield return ("rejecting an open document", () => Opened(RevisionProcessor.RejectRevisions));
        yield return ("simplifying an open document", () => Opened(d => MarkupSimplifier.SimplifyMarkup(d, new SimplifyMarkupSettings { RemoveComments = true, RemoveContentControls = true })));
        yield return ("single-character runs of an open document", () => Opened(MarkupSimplifier.TransformToSingleCharacterRuns));
        yield return ("formatting an open document", () => Opened(d => FormattingAssembler.AssembleFormatting(d, new FormattingAssemblerSettings())));
        yield return ("OpenContracts export of an open document", () => Opened(d => OpenContractExporter.Export(d)));
        yield return ("semantic changes", () => { DocxDiff.GetSemanticChanges(bytes, changed); return null; });
        yield return ("deliverable verification", () => Finding(DeliverableVerifier.VerifyDeliverable(bytes).Findings.Select(f => f.Code)));
        yield return ("OpenContracts export", () => { OpenContractExporter.Export(new WmlDocument("a.docx", bytes)); return null; });
        yield return ("metrics", () => { MetricsGetter.GetDocxMetrics(new WmlDocument("a.docx", bytes), new MetricsGetterSettings()); return null; });
        yield return ("metrics from a file", () =>
        {
            var path = Path.Combine(Path.GetTempPath(), $"depth-{Guid.NewGuid():N}.docx");
            File.WriteAllBytes(path, bytes);
            try { MetricsGetter.GetMetrics(path, new MetricsGetterSettings()); }
            finally { File.Delete(path); }
            return null;
        });
        yield return ("package manifest", () => Finding(PackageManifestGenerator.Generate(bytes).Findings.Select(f => f.Code)));
    }

    internal static void OnThread(int stackKb, Action body)
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try { body(); }
            catch (Exception ex) { failure = ex; }
        }, maxStackSize: stackKb * 1024);
        thread.Start();
        thread.Join();
        if (failure is not null) throw new Xunit.Sdk.XunitException(failure.ToString());
    }
}

/// <summary>Depth-door behaviour the sweep does not cover.</summary>
public class DocxSessionDeepNestingTests
{
    [Fact]
    public void NestedTablesAtTheStoryLimit_RenderToHtmlOnA1MbStack()
    {
        // Nested tables are the deepest-recursing shape measured (the HTML render); at the story
        // limit it still fits a 1 MB stack. Raising the limit to the old 256 overflows here.
        var bytes = LargeDocuments.NestedTables((XmlSemanticNormalizer.MaxStoryDepth - 5) / 3);
        Assert.InRange(LargeDocuments.Depth(bytes), XmlSemanticNormalizer.MaxStoryDepth - 2, XmlSemanticNormalizer.MaxStoryDepth);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            using (var session = new DocxSession(bytes))
                Assert.Single(session.Grep("Innermost text"));
            Assert.Contains("Innermost text", WmlToHtmlConverter.ConvertToHtml(
                new WmlDocument("d.docx", bytes), new WmlToHtmlConverterSettings()).ToString());
        });
    }

    [Fact]
    public void RefusingAnOverDeepDocument_StopsReadingAtTheLimit()
    {
        // Loading the whole 24,000-level tree first took several seconds (LINQ to XML's load is
        // superlinear in depth); the read step's scan refuses it at level 128 without building it.
        var bytes = LargeDocuments.NestedTables(levels: 8000);
        var stopwatch = Stopwatch.StartNew();
        var refusal = Assert.ThrowsAny<XmlException>(() => new DocxSession(bytes).Dispose());
        Assert.StartsWith("xml_depth_limit_exceeded:", refusal.Message, StringComparison.Ordinal);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(2), $"refusal took {stopwatch.Elapsed}");
    }

    [Fact]
    public void RawXmlEdits_CannotGrowTheTreePastTheLimit()
    {
        // Edits never pass the read step. Each payload is shallow, but wrapping the innermost
        // paragraph in one more w:customXml nests its text one level deeper: the edit that would put
        // it past the story limit is refused, so the saved document sits exactly at the limit.
        using var session = new DocxSession(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Innermost text") }));
        const string payload = "<w:customXml xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "w:element=\"wrap\"><w:p><w:r><w:t>Innermost text</w:t></w:r></w:p></w:customXml>";
        EditResult? refused = null;
        for (int edit = 0; edit < 200 && refused is null; edit++)
        {
            var anchor = Assert.Single(session.Grep("Innermost text")).EnclosingAnchor.Anchor.Id;
            var result = session.Raw.ReplaceXml(anchor, payload);
            if (!result.Success) refused = result;
        }
        Assert.NotNull(refused);
        Assert.Equal(EditErrorCode.ValidationFailed, refused!.Error!.Code);
        Assert.StartsWith("xml_depth_limit_exceeded:", refused.Error.Message, StringComparison.Ordinal);
        var saved = session.Save();
        Assert.Equal(XmlSemanticNormalizer.MaxStoryDepth, LargeDocuments.Depth(saved));
        using var reopened = new DocxSession(saved);
        Assert.Single(reopened.Grep("Innermost text"));
    }

    [Fact]
    public void ARawXmlPayloadFarPastTheLimit_IsRefusedOnA1MbStack()
    {
        using var session = new DocxSession(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Innermost text") }));
        var anchor = Assert.Single(session.Grep("Innermost text")).EnclosingAnchor.Anchor.Id;
        const string open = "<w:customXml w:element=\"wrap\">";
        var payload = "<w:customXml xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:element=\"wrap\">"
            + string.Concat(Enumerable.Repeat(open, 20_000)) + "<w:p><w:r><w:t>Deep</w:t></w:r></w:p>"
            + string.Concat(Enumerable.Repeat("</w:customXml>", 20_001));
        var saved = session.Save();
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            var result = session.Raw.ReplaceXml(anchor, payload);
            Assert.Equal(EditErrorCode.ValidationFailed, result.Error?.Code);
        });
        Assert.Equal(saved, session.Save());
    }

    [Theory]
    [InlineData("main part stored as .dat")]
    [InlineData("main part under a foreign root name")]
    public void TheDoorClassifiesByDeclaredContentType_NotByEntryNameOrRootName(string variant)
    {
        // The SDK loads a part by its declared content type, so the door does too: a main part named
        // word/document.dat, or rooted <x:document> in a foreign namespace, is still the main story.
        var bytes = DeepPackages.Build("body", XmlSemanticNormalizer.MaxStoryDepth + 1, "Closing text");
        bytes = variant.EndsWith(".dat", StringComparison.Ordinal)
            ? DeepPackages.RenameEntry(bytes, "word/document.xml", "word/document.dat")
            : DeepPackages.RewriteEntry(bytes, "word/document.xml", xml => xml
                .Replace("<w:document ", "<x:document xmlns:x=\"urn:example:foreign\" ", StringComparison.Ordinal)
                .Replace("</w:document>", "</x:document>", StringComparison.Ordinal));
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            AssertRefused(() => new DocxSession(bytes).Dispose());
            AssertRefused(() => HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions()));
        });
    }

    [Theory]
    [InlineData("a DOCTYPE")]
    [InlineData("a wrong encoding declaration")]
    public void AnEntryOnlyTheStrictUnfoldParses_CannotCrashIt(string variant)
    {
        // The strict-conformance unfold parses every .xml entry holding a strict namespace itself
        // (DTD allowed, text decoded as UTF-8), referenced or not, and re-homes its shapes recursively.
        // Its parse is depth-limited: 40,000 levels of strict-namespace shape markup in an unreferenced
        // entry neither crashes it nor refuses the package, since nothing else reads that entry.
        var prolog = variant == "a DOCTYPE" ? "<?xml version=\"1.0\"?><!DOCTYPE r>" : "<?xml version=\"1.0\" encoding=\"utf-16\"?>";
        var deep = prolog + "<r xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\"><a:graphicData uri=\"" +
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            string.Concat(Enumerable.Repeat("<a:x>", 40_000)) + string.Concat(Enumerable.Repeat("</a:x>", 40_000)) +
            "</a:graphicData></r>";
        var bytes = DeepPackages.AddEntry(DeepPackages.Build("strict body with a deep shape", 20, "Closing text"),
            "word/extra.xml", deep);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            using var session = new DocxSession(bytes);
            Assert.Contains("Closing text", session.Project().Markdown);
        });
    }

    [Theory]
    [InlineData("AlternateContent")]
    [InlineData("duplicate paragraph properties")]
    public void AnUnreferencedEntryTheCompatibilityPassRewrites_CannotCrashIt(string shape)
    {
        // DocxDiff's markup-compatibility pass parses every .xml entry carrying one of its shapes,
        // referenced or not, and rewrites it recursively; its parse is depth-limited, so 40,000 levels
        // in an entry nothing references are left alone.
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";
        var chain = string.Concat(Enumerable.Repeat("<w:x>", 40_000)) + string.Concat(Enumerable.Repeat("</w:x>", 40_000));
        var deep = shape == "AlternateContent"
            ? $"<w:document {ns}><w:body><w:p><w:r><mc:AlternateContent><mc:Fallback>{chain}</mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>"
            : $"<w:document {ns}><w:body><w:p><w:pPr/><w:pPr>{chain}</w:pPr></w:p></w:body></w:document>";
        var left = DeepPackages.AddEntry(DeepPackages.Build("body", 20, "Closing text"), "extra/deep.xml", deep);
        var right = DeepPackages.AddEntry(DeepPackages.Build("body", 20, "Closing text, revised"), "extra/deep.xml", deep);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
            Assert.NotEmpty(DocxCompare.Compare(new WmlDocument("a.docx", left), new WmlDocument("b.docx", right)).DocumentByteArray));
    }

    [Fact]
    public void ARelationshipCycle_IsWalkedOnce()
    {
        // A header whose relationships point back at itself and at the main document.
        var bytes = DeepPackages.Build("header", 20, "Closing text");
        bytes = DeepPackages.AddEntry(bytes, "word/_rels/header1.xml.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rIdSelf\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
            "<Relationship Id=\"rIdDoc\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"document.xml\"/>" +
            "</Relationships>");
        var header = DeepPackages.HeaderEntryName(bytes);
        Assert.NotNull(header);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            using var session = new DocxSession(bytes);
            Assert.Contains("Closing text", session.Project().Markdown);
        });
    }

    [Fact]
    public void AnExternalRelationshipNamingAnInternalPart_LoadsNothing()
    {
        // TargetMode="External" makes the target a URI, not a part: no layer loads it, so a 40,000-level
        // header behind it is never read, by the read step or by any later step.
        var deepHeader = "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            string.Concat(Enumerable.Repeat("<w:customXml w:element=\"x\">", 40_000)) + "<w:p/>" +
            string.Concat(Enumerable.Repeat("</w:customXml>", 40_000)) + "</w:hdr>";
        var bytes = DeepPackages.AddEntry(DeepPackages.Build("body", 20, "Closing text"), "word/header9.xml", deepHeader);
        bytes = DeepPackages.SetOverride(bytes, "/word/header9.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
        bytes = DeepPackages.RewriteEntry(bytes, "word/_rels/document.xml.rels", xml => xml.Replace("</Relationships>",
            "<Relationship Id=\"rIdExt\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" " +
            "Target=\"header9.xml\" TargetMode=\"External\"/></Relationships>", StringComparison.Ordinal));
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            using (var session = new DocxSession(bytes)) Assert.NotEmpty(session.Save());
            Assert.NotEmpty(DocxCompare.Compare(new WmlDocument("a.docx", bytes), new WmlDocument("b.docx", bytes)).DocumentByteArray);
            WmlToHtmlConverter.ConvertToHtml(new WmlDocument("a.docx", bytes), new WmlToHtmlConverterSettings());
        });
    }

    [Fact]
    public void ExtremeDepthInput_IsRefusedByTheWmlDocumentEntryPoints()
    {
        // 40,000 levels overflow the strict unfold and the typed DOM long before any later check:
        // DocxCompare unfolds a strict WmlDocument first; GetMetrics(path) forces every typed DOM first.
        var bytes = DeepPackages.Build("strict body with a deep shape", 40_000, "Closing text");
        var changed = DeepPackages.Build("strict body with a deep shape", 40_000, "Closing text, revised");
        var path = Path.Combine(Path.GetTempPath(), $"depth-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DeepPackages.Build("body", 40_000, "Closing text"));
        try
        {
            DocxDepthLimitSweepTests.OnThread(1024, () =>
            {
                AssertRefused(() => DocxCompare.Compare(new WmlDocument("a.docx", bytes), new WmlDocument("b.docx", changed)));
                AssertRefused(() => MetricsGetter.GetMetrics(path, new MetricsGetterSettings()));
            });
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void TheScanCache_RescansAnArrayWhoseContentChanged()
    {
        // Same length, different nesting: the cache must not trust the array it scanned before.
        var (shallow, deep) = DeepPackages.SameLengthPair(XmlSemanticNormalizer.MaxStoryDepth + 1);
        Assert.Equal(shallow.Length, deep.Length);
        var bytes = shallow.ToArray();
        new DocxSession(bytes).Dispose();
        Buffer.BlockCopy(deep, 0, bytes, 0, bytes.Length);
        AssertRefused(() => new DocxSession(bytes).Dispose());
    }

    [Fact]
    public void AnOpenDocument_IsMeasuredOncePerCall_AndTheLibrarysOwnNever()
    {
        var bytes = LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Measured text") });
        static int Measurements(Action run)
        {
            var before = PartDepthGuard.OpenPackageMeasurements;
            run();
            return PartDepthGuard.OpenPackageMeasurements - before;
        }

        using (var stream = new MemoryStream())
        {
            stream.Write(bytes);
            using var document = WordprocessingDocument.Open(stream, true);
            // The HTML render accepts, simplifies and assembles the same package in turn: one measure.
            Assert.Equal(1, Measurements(() => WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings())));
            // The caller can change the package between calls, so the next call measures it again.
            Assert.Equal(1, Measurements(() => WmlToMarkdownConverter.Convert(document, new WmlToMarkdownConverterSettings())));
        }
        Assert.Equal(0, Measurements(() => WmlToHtmlConverter.ConvertToHtml(new WmlDocument("a.docx", bytes), new WmlToHtmlConverterSettings())));
        Assert.Equal(0, Measurements(() => HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions())));
        using var session = new DocxSession(bytes);
        Assert.Equal(0, Measurements(() =>
        {
            var anchor = session.Project().AnchorIndex.Keys.First();
            HtmlConversionOps.RenderBlockHtml(session, anchor, new HtmlConversionOptions { StampAnchors = true });
            session.ListAnnotations();
        }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnOpenDocumentsLoadedTree_IsWhatIsMeasured(bool openedByTheLibrary)
    {
        // A caller that deepens a part's tree in memory, after the stream was read, is refused on it —
        // also in a package the library opened for it from bytes the door measured.
        var bytes = LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Loaded text") });
        using var stream = new MemoryStream();
        using var wrapper = new OpenXmlMemoryStreamDocument(new WmlDocument("a.docx", bytes));
        stream.Write(bytes);
        using var document = openedByTheLibrary ? wrapper.GetMeasuredWordprocessingDocument() : WordprocessingDocument.Open(stream, true);
        WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings());
        var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        XElement deep = LargeDocuments.Paragraph("Deep text");
        for (int level = 0; level < XmlSemanticNormalizer.MaxStoryDepth; level++)
            deep = new XElement(W.customXml, new XAttribute(W.element, "wrap"), deep);
        body.AddFirst(deep);
        AssertRefused(() => WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings()));
    }

    [Fact]
    public void AWrapperACallerFedNewBytes_IsMeasuredAgain()
    {
        // The public wrapper hands out a package whose streams its caller can rewrite, so it does not
        // claim them measured.
        using var wrapper = new OpenXmlMemoryStreamDocument(new WmlDocument("a.docx",
            LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Wrapped text") })));
        using var document = wrapper.GetWordprocessingDocument();
        document.MainDocumentPart!.FeedData(new MemoryStream(MainPartBytes(
            DeepPackages.Build("body", XmlSemanticNormalizer.MaxStoryDepth + 1, "Closing text"))));
        AssertRefused(() => WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings()));
    }

    [Fact]
    public void TheOpenContractsExport_NeverWalksAStreamTheDoorDidNotMeasure()
    {
        // The export read the main story through the SDK's typed DOM, which loads from the stream: a
        // caller's loaded (shallow) tree was measured while a deep stream fed in behind it was walked,
        // overflowing the stack. It now reads the measured tree; its annotation step, which reads the
        // package's bytes, then refuses the deep stream at the byte read step.
        using var stream = new MemoryStream();
        stream.Write(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Loaded text") }));
        using var document = WordprocessingDocument.Open(stream, true);
        document.MainDocumentPart!.GetXDocument();
        document.MainDocumentPart!.FeedData(new MemoryStream(MainPartBytes(DeepPackages.Build("body", 20_000, "Closing text"))));
        DocxDepthLimitSweepTests.OnThread(1024, () => AssertRefused(() => OpenContractExporter.Export(document)));
    }

    [Fact]
    public void TheOpenContractsExportOfAnOpenDocumentFarPastTheLimit_IsRefusedByTheOpenDocumentDoor()
    {
        // Far past the limit, so a refusal here is the open-document door's and not the byte read step's
        // the export reaches later.
        using var stream = new MemoryStream();
        stream.Write(DeepPackages.Build("body", 20_000, "Closing text"));
        using var document = WordprocessingDocument.Open(stream, true);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            var refusal = Assert.ThrowsAny<XmlException>(() => OpenContractExporter.Export(document));
            Assert.StartsWith("xml_depth_limit_exceeded:", refusal.Message, StringComparison.Ordinal);
        });
    }

    [Theory]
    [InlineData("body", XmlSemanticNormalizer.MaxStoryDepth, false)]
    [InlineData("body", XmlSemanticNormalizer.MaxStoryDepth + 1, true)]
    [InlineData("header", 20_000, false)]
    public void DocumentMetadataOfAnOpenDocument_MeasuresOnlyTheMainStory(string partClass, int depth, bool refused)
    {
        // Metadata walks the main story recursively (nested tables) and reads every other part flat.
        using var stream = new MemoryStream();
        stream.Write(DeepPackages.Build(partClass, depth, "Closing text"));
        using var document = WordprocessingDocument.Open(stream, true);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            if (refused) AssertRefused(() => WmlToHtmlConverter.GetDocumentMetadata(document));
            else Assert.NotNull(WmlToHtmlConverter.GetDocumentMetadata(document));
        });
    }

    [Fact]
    public void AnnotationsOfAnOpenDocument_NeedNoMeasuring()
    {
        // Listing annotations reads the main story flat, so even a tree far past the limit is no hazard.
        var annotated = AnnotationManager.AddAnnotation(new WmlDocument("a.docx",
                LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Annotated text") })),
            new DocumentAnnotation { Id = "a1", LabelId = "L", Label = "Label" }, AnnotationRange.FromSearch("Annotated"));
        using var stream = new MemoryStream();
        stream.Write(annotated.DocumentByteArray);
        using var document = WordprocessingDocument.Open(stream, true);
        var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        XElement content = new(W.customXml, new XAttribute(W.element, "wrap"), body.Elements().Where(e => e.Name != W.sectPr).ToList());
        body.Elements().Where(e => e.Name != W.sectPr).Remove();
        for (int level = 0; level < 20_000; level++)
            content = new XElement(W.customXml, new XAttribute(W.element, "wrap"), content);
        body.AddFirst(content);
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            Assert.Equal("Annotated", Assert.Single(AnnotationManager.GetAnnotations(document)).AnnotatedText);
            Assert.Equal(0, PartDepthGuard.OpenPackageMeasurements);
        });
    }

    private static byte[] MainPartBytes(byte[] package)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(package), false);
        using var part = document.MainDocumentPart!.GetStream();
        using var copy = new MemoryStream();
        part.CopyTo(copy);
        return copy.ToArray();
    }

    [Theory]
    [InlineData(XmlSemanticNormalizer.MaxStoryDepth, 0)]
    [InlineData(2000, 1)]
    public async Task TheDocx2HtmlTool_RefusesAnOverDeepFileWithAnError(int depth, int exitCode)
    {
        // The tool opens the file as a package itself; far past the limit the render used to overflow
        // the stack and abort the process.
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        var configuration = directory.Parent!.Name;
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Docxodus.Tests", "Docxodus.Tests.csproj")))
            directory = directory.Parent;
        var tool = Path.Combine(directory!.FullName, "tools", "docx2html", "bin", configuration, "net10.0", "docx2html.dll");
        Assert.True(File.Exists(tool), "Project-reference-built docx2html is missing: " + tool);
        var root = Directory.CreateTempSubdirectory("docx2html-depth-").FullName;
        try
        {
            var input = Path.Combine(root, "deep.docx");
            File.WriteAllBytes(input, DeepPackages.Build("body", depth, "Closing text"));
            var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
            { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            foreach (var value in new[] { tool, input, Path.Combine(root, "deep.html") }) start.ArgumentList.Add(value);
            using var process = Process.Start(start)!;
            var stdout = process.StandardOutput.ReadToEndAsync();
            var stderr = process.StandardError.ReadToEndAsync();
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            try { await process.WaitForExitAsync(timeout.Token); }
            catch { if (!process.HasExited) process.Kill(entireProcessTree: true); throw; }
            var text = await stdout + await stderr;
            Assert.True(process.ExitCode == exitCode, $"exit {process.ExitCode}: {text}");
            if (exitCode != 0) Assert.Contains("Error: xml_depth_limit_exceeded:", await stderr, StringComparison.Ordinal);
        }
        finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(200, false)]
    [InlineData(200, true)]
    [InlineData(20_000, false)]
    [InlineData(20_000, true)]
    public void ProjectingAnnotations_RefusesHtmlNestedPastTheLimit(int levels, bool fragment)
    {
        // The projector parses the caller's HTML itself. A fragment (several roots) is parsed again
        // inside a wrapper; a refusal must not be mistaken for a fragment.
        var doc = new WmlDocument("a.docx", LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Annotated text") }));
        var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "depth");
        set.TextLabels["LABEL"] = new AnnotationLabel { Id = "LABEL", Text = "Label", Color = "#FF0000" };
        set.LabelledText.Add(ExternalAnnotationManager.CreateAnnotationFromSearch("ann", "LABEL", set.Content, "Annotated", 1)!);
        var html = string.Concat(Enumerable.Repeat("<div>", levels)) + "<p>Annotated text</p>"
            + string.Concat(Enumerable.Repeat("</div>", levels)) + (fragment ? "<p>Second root</p>" : "");
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            if (levels < XmlSemanticNormalizer.MaxElementDepth)
                Assert.Contains("data-annotation-id=\"ann\"", ExternalAnnotationProjector.ProjectAnnotationsOntoHtml(html, set));
            else
                AssertRefused(() => ExternalAnnotationProjector.ProjectAnnotationsOntoHtml(html, set));
        });
    }

    [Theory]
    [InlineData(XmlSemanticNormalizer.MaxElementDepth, false)]
    [InlineData(XmlSemanticNormalizer.MaxElementDepth + 1, true)]
    public void FlatOpcDumps_HoldEveryPartToTheLimit(int depth, bool refused)
    {
        // The dumps read every part, including one nothing references, which the door never measures.
        var xml = string.Concat(Enumerable.Repeat("<n>", depth)) + string.Concat(Enumerable.Repeat("</n>", depth));
        var path = Path.Combine(Path.GetTempPath(), $"flat-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DeepPackages.AddEntry(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Dumped text") }),
            "customXml/unreferenced.xml", xml));
        try
        {
            foreach (var dump in new Action[]
            {
                () => FlatOpc.OpcToXmlDocument(path),
                () => FlatOpc.OpcToXDocument(path),
                () => FlatOpc.OpcToText(path),
            })
            {
                if (refused) AssertRefused(dump);
                else dump();
            }
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void TypedEdits_CannotNestAPartPastTheLimit()
    {
        // Edits never pass the read step. Each table nested in the previous one's only cell adds three
        // levels; the one that would put the story past its limit is refused atomically.
        using var session = new DocxSession(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Cell 000") }));
        var refused = NestTablesUntilRefused(session, out var innermost, out var accepted);
        Assert.Equal(EditErrorCode.ValidationFailed, refused.Error!.Code);
        Assert.StartsWith("xml_depth_limit_exceeded:", refused.Error.Message, StringComparison.Ordinal);
        var saved = session.Save();
        Assert.InRange(LargeDocuments.Depth(saved), XmlSemanticNormalizer.MaxStoryDepth - 2, XmlSemanticNormalizer.MaxStoryDepth);

        // Nothing changed: same version, same document, and the last undo step is the last accepted table.
        var version = session.Version;
        Assert.False(InsertCellTable(session, innermost, accepted + 1).Success);
        Assert.Equal(version, session.Version);
        Assert.Equal(saved, session.Save());
        Assert.True(session.Undo());
        Assert.Empty(session.Grep($"Cell {accepted:D3}"));

        // A restored tree is watched too.
        Assert.True(session.Redo());
        Assert.False(InsertCellTable(session, innermost, accepted + 1).Success);
        Assert.Single(session.Grep($"Cell {accepted:D3}"));
    }

    [Fact]
    public void EveryTreeAnOpEdits_CarriesExactlyOneWatch()
    {
        // Undo, redo and rollback install new trees; the next op's snapshot must watch them, once.
        using var session = SessionAtTheLimit(out var innermost);
        void Edit(string text) =>
            Assert.True(session.InsertParagraph(Assert.Single(session.Grep("Cell 000")).EnclosingAnchor.Anchor.Id,
                Position.After, text).Success);
        void AssertWatched() =>
            Assert.Single(session.LiveDocument.MainDocumentPart!.GetXDocument().Annotations<PartDepthWatch>());

        Edit("First");
        AssertWatched();
        Assert.True(session.Undo());
        Edit("After undo");
        AssertWatched();
        Assert.True(session.Undo());
        Assert.True(session.Redo());
        Edit("After redo");
        AssertWatched();
        var transaction = session.BeginTransaction();
        Edit("Rolled back");
        transaction.Rollback();
        Edit("After rollback");
        AssertWatched();
        Assert.False(InsertCellTable(session, innermost, 999).Success);
        Edit("After a refusal");
        AssertWatched();
    }

    [Fact]
    public void ASessionNestedToTheLimit_RendersEachBlockOnA1MbStack()
    {
        using var session = new DocxSession(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Cell 000") }));
        NestTablesUntilRefused(session, out _, out var accepted);
        var tables = session.Project().AnchorIndex.Keys.Where(a => a.StartsWith("tbl:", StringComparison.Ordinal)).ToList();
        Assert.Equal(accepted, tables.Count);
        // Every table holds the deepest cell, so the outermost one is the deepest render.
        DocxDepthLimitSweepTests.OnThread(1024, () =>
        {
            foreach (var anchor in tables)
                Assert.Contains($"Cell {accepted:D3}", HtmlConversionOps.RenderBlockHtml(session, anchor, new HtmlConversionOptions { StampAnchors = true }));
        });
    }

    [Fact]
    public void EditsBesideTheDeepestContent_AreAccepted()
    {
        // Control: content as deep as what is already there adds no depth.
        using var session = SessionAtTheLimit(out var innermost);
        Assert.True(session.InsertParagraph(innermost, Position.After, "Beside the deepest").Success);
        Assert.True(session.Raw.ReplaceXml(innermost,
            "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Replaced</w:t></w:r></w:p>").Success);
        Assert.Single(session.Grep("Replaced"));
    }

    [Fact]
    public void AListEditAtTheEdge_IsRefusedLikeEveryOtherEdit()
    {
        // A style-numbered item whose properties sit on the last allowed level: giving it its own
        // w:numPr would nest w:ilvl one level too deep, so the op is refused and nothing changes.
        XElement block = new(W.p, new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, "ListItem"))),
            new XElement(W.r, new XElement(W.t, "Edge item")));
        for (int level = 0; level < XmlSemanticNormalizer.MaxStoryDepth - 5; level++)
            block = new XElement(W.customXml, new XAttribute(W.element, "wrap"), block);
        using var session = new DocxSession(WithStyleNumbering(LargeDocuments.Build(new[] { block })));
        var item = Assert.Single(session.Grep("Edge item")).EnclosingAnchor.Anchor.Id;
        var saved = session.Save();

        var result = session.SetListLevel(item, 1);
        Assert.Equal(EditErrorCode.ValidationFailed, result.Error?.Code);
        Assert.Equal(saved, session.Save());
        Assert.False(session.Undo());
    }

    private static byte[] WithStyleNumbering(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            main.AddNewPart<NumberingDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.numbering,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.abstractNum, new XAttribute(W.abstractNumId, 0),
                    new XElement(W.lvl, new XAttribute(W.ilvl, 0), new XElement(W.numFmt, new XAttribute(W.val, "decimal"))),
                    new XElement(W.lvl, new XAttribute(W.ilvl, 1), new XElement(W.numFmt, new XAttribute(W.val, "decimal")))),
                new XElement(W.num, new XAttribute(W.numId, 1), new XElement(W.abstractNumId, new XAttribute(W.val, 0))))));
            main.StyleDefinitionsPart!.PutXDocument(new XDocument(new XElement(W.styles,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.style, new XAttribute(W.type, "paragraph"), new XAttribute(W.styleId, "ListItem"),
                    new XElement(W.name, new XAttribute(W.val, "List Item")),
                    new XElement(W.pPr, new XElement(W.numPr, new XElement(W.numId, new XAttribute(W.val, 1))))))));
        }
        return stream.ToArray();
    }

    [Fact]
    public void RollingBackInASessionAtTheLimit_KeepsTheSessionUsable()
    {
        using var session = SessionAtTheLimit(out var innermost);
        var transaction = session.BeginTransaction();
        Assert.True(session.InsertParagraph(innermost, Position.After, "Rolled back").Success);
        transaction.Rollback();
        Assert.Empty(session.Grep("Rolled back"));
        Assert.True(session.InsertParagraph(innermost, Position.After, "Kept").Success);
        Assert.Single(session.Grep("Kept"));
        Assert.NotEmpty(session.Save());
    }

    [Fact]
    public void ARefusalInsideATransaction_KeepsTheTransactionsOtherEdits()
    {
        using var session = SessionAtTheLimit(out var innermost);
        var transaction = session.BeginTransaction();
        Assert.True(session.InsertParagraph(innermost, Position.After, "Kept in the transaction").Success);
        Assert.Equal(EditErrorCode.ValidationFailed, InsertCellTable(session, innermost, 999).Error?.Code);
        transaction.Commit();
        Assert.Single(session.Grep("Kept in the transaction"));
        Assert.Empty(session.Grep("Cell 999"));
        Assert.True(session.Undo());
        Assert.Empty(session.Grep("Kept in the transaction"));
    }

    [Fact]
    public void PreviewsOfASessionAtTheLimit_ReadItsOwnOutput()
    {
        using var session = SessionAtTheLimit(out var innermost);
        var preview = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "insert_paragraph", s => s.InsertParagraph(innermost, Position.After, "Previewed")),
        });
        Assert.True(preview.Success);
        Assert.Empty(session.Grep("Previewed"));
    }

    [Fact]
    public void PreviewsThroughTheRegistry_ReadTheSessionsOwnOutput()
    {
        int handle = DocxSessionOps.OpenSession(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Cell 000") }), null);
        try
        {
            NestTablesUntilRefused(SessionRegistry.Get(handle), out _, out _);
            int shadow = SessionRegistry.CloneSessionForPreview(handle);
            SessionRegistry.CloseSession(shadow);
        }
        finally { DocxSessionOps.CloseSession(handle); }
    }

    [Fact]
    public void SemanticChangesOfASessionAtTheLimit_ReadAsTheSavedFileWould()
    {
        using var session = SessionAtTheLimit(out _);
        Assert.NotNull(session.GetSemanticChanges());
        new DocxSession(session.Save()).Dispose();
    }

    /// <summary>A session whose body a typed edit nested as deep as the story limit allows.</summary>
    private static DocxSession SessionAtTheLimit(out string innermost)
    {
        var session = new DocxSession(LargeDocuments.Build(new[] { LargeDocuments.Paragraph("Cell 000") }),
            new DocxSessionSettings { CaptureInitialProjection = true });
        NestTablesUntilRefused(session, out innermost, out _);
        return session;
    }

    /// <summary>Nest tables, each in the previous one's only cell, until an insert is refused. Returns
    /// the refusal; <paramref name="innermost"/> is the deepest cell's paragraph, holding
    /// <c>Cell {accepted}</c>.</summary>
    private static EditResult NestTablesUntilRefused(DocxSession session, out string innermost, out int accepted)
    {
        innermost = Assert.Single(session.Grep("Cell 000")).EnclosingAnchor.Anchor.Id;
        // Stop short of a depth the session's own recursive reads could not survive.
        for (accepted = 0; accepted < XmlSemanticNormalizer.MaxStoryDepth / 3 + 3; accepted++)
        {
            var result = InsertCellTable(session, innermost, accepted + 1);
            if (!result.Success) return result;
            innermost = Assert.Single(session.Grep($"Cell {accepted + 1:D3}")).EnclosingAnchor.Anchor.Id;
        }
        throw new Xunit.Sdk.XunitException($"{accepted} nested tables were all accepted");
    }

    private static EditResult InsertCellTable(DocxSession session, string anchor, int level) =>
        session.InsertTable(anchor, Position.After, 1, 1, new TableInsertOptions { CellContents = new[] { $"Cell {level:D3}" } });

    [Fact]
    public void ARepairThatTypesAPart_IsMeasuredAgain()
    {
        // An image part with no declared content type is not a part to the packaging layer until
        // open-time repair gives it the type its extension and relationship imply; the read step then
        // measures it. Here: an untyped SVG nested past the general limit.
        var svg = "<svg xmlns=\"http://www.w3.org/2000/svg\">" + string.Concat(Enumerable.Repeat("<g>", XmlSemanticNormalizer.MaxElementDepth))
            + string.Concat(Enumerable.Repeat("</g>", XmlSemanticNormalizer.MaxElementDepth)) + "</svg>";
        var bytes = DeepPackages.WithImage(DeepPackages.Build("body", 20, "Closing text"), "word/media/deep.svg", svg, typed: false);
        AssertRefused(() => new DocxSession(bytes).Dispose());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ADeepPartNothingWalksAsAStory_KeepsTheGeneralLimit(bool referenced)
    {
        // A 150-deep SVG image: unreferenced, nothing parses it; referenced, it is not a story part.
        // Either way it opens as on main.
        var svg = "<svg xmlns=\"http://www.w3.org/2000/svg\">" + string.Concat(Enumerable.Repeat("<g>", 150))
            + string.Concat(Enumerable.Repeat("</g>", 150)) + "</svg>";
        var package = DeepPackages.Build("body", 20, "Closing text");
        var bytes = referenced
            ? DeepPackages.WithImage(package, "word/media/deep.svg", svg, typed: true)
            : DeepPackages.AddEntry(package, "word/media/deep.svg", svg);
        using var session = new DocxSession(bytes);
        Assert.Contains("Closing text", session.Project().Markdown);
    }

    [Fact]
    public void UnreferencedBombs_CostNothing()
    {
        // Fifty billion-laughs entries and a whitespace bomb that nothing references: main never
        // reads them, and neither does the read step.
        const string laughs = "<?xml version=\"1.0\"?><!DOCTYPE l [<!ENTITY a \"aaaaaaaaaa\"><!ENTITY b \"&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;\">" +
            "<!ENTITY c \"&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;\"><!ENTITY d \"&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;\"><!ENTITY e \"&d;&d;&d;&d;&d;&d;&d;&d;&d;&d;\">" +
            "<!ENTITY f \"&e;&e;&e;&e;&e;&e;&e;&e;&e;&e;\"><!ENTITY g \"&f;&f;&f;&f;&f;&f;&f;&f;&f;&f;\">]><l>&g;</l>";
        var bytes = DeepPackages.Build("body", 20, "Closing text");
        for (int i = 0; i < 50; i++) bytes = DeepPackages.AddEntry(bytes, $"extra/laughs{i}.xml", laughs);
        bytes = DeepPackages.AddSpaces(bytes, "extra/spaces.xml", 256L * 1024 * 1024);
        var stopwatch = Stopwatch.StartNew();
        using (var session = new DocxSession(bytes))
        {
            Assert.Contains("Closing text", session.Project().Markdown);
            Assert.NotEmpty(session.Save());
        }
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5), $"took {stopwatch.Elapsed}");
    }

    [Theory]
    [InlineData("two parts with a DTD", "xml_size_limit_exceeded")]
    [InlineData("an inflating part", "compression_ratio_limit_exceeded")]
    public void ReferencedBombs_AreRefusedFast_ForAStatedReason(string bomb, string code)
    {
        const string laughs = "<?xml version=\"1.0\"?><!DOCTYPE l [<!ENTITY a \"aaaaaaaaaa\">]><l xmlns=\"urn:example:data\">&a;</l>";
        var bytes = DeepPackages.Build("body", 20, "Closing text");
        bytes = bomb == "two parts with a DTD"
            ? DeepPackages.WithCustomXml(DeepPackages.WithCustomXml(bytes, "customXml/item1.xml", laughs), "customXml/item2.xml", laughs)
            : DeepPackages.WithCustomXml(bytes, "customXml/item1.xml", null, spaces: 256L * 1024 * 1024);
        var stopwatch = Stopwatch.StartNew();
        var refusal = Assert.ThrowsAny<XmlException>(() => new DocxSession(bytes).Dispose());
        Assert.StartsWith(code + ":", refusal.Message, StringComparison.Ordinal);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(2), $"took {stopwatch.Elapsed}");
    }

    private static void AssertRefused(Action open)
    {
        var refusal = Assert.ThrowsAny<Exception>(open);
        Assert.Contains("xml_depth_limit_exceeded", refusal.Message, StringComparison.Ordinal);
    }

}

/// <summary>Packages with one XML part nested exactly <c>depth</c> elements deep, one level per step.</summary>
internal static class DeepPackages
{
    private static readonly XNamespace X = "urn:example:deep";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>A chain of <paramref name="levels"/> nested elements around <paramref name="leaf"/>.</summary>
    private static XElement Chain(XName name, int levels, XElement leaf)
    {
        for (int i = 0; i < levels; i++) leaf = new XElement(name, name.Namespace == W.w ? new XAttribute(W.element, "wrap") : null, leaf);
        return leaf;
    }

    private static XElement Leaf() => new(X + "leaf", "deep");

    internal static byte[] Build(string partClass, int depth, string closing)
    {
        if (partClass == "strict body with a deep shape") return Strict(depth, closing);
        if (partClass == "strict body the unfold makes well-formed")
            // A repeated attribute is not XML; the unfold drops every w:conformance="strict", curing it.
            return RewriteEntry(Strict(depth, closing), "word/document.xml", xml => xml.Replace("<w:body>",
                "<w:body w:conformance=\"strict\" w:conformance=\"strict\">", StringComparison.Ordinal));
        if (partClass == "interleaved main part") return Interleave(Build("body", depth, closing), "word/document.xml");
        if (partClass == "main part typed application/xml")
            return SetOverride(Build("body", depth, closing), "/word/document.xml", "application/xml");
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var stylesRoot = new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w));
            var settings = new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w));
            var blocks = new List<object>();
            // document(1) body(2) chain p r t
            if (partClass == "body") blocks.Add(Chain(W.customXml, depth - 5, LargeDocuments.Paragraph("Deep text")));
            XElement? sectPr = null;
            switch (partClass)
            {
                case "header":
                    var header = main.AddNewPart<HeaderPart>();
                    header.PutXDocument(new XDocument(new XElement(W.hdr, new XAttribute(XNamespace.Xmlns + "w", W.w),
                        Chain(W.customXml, depth - 4, LargeDocuments.Paragraph("Deep text")))));
                    sectPr = new XElement(W.sectPr, new XElement(W.headerReference,
                        new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(header))));
                    break;
                case "footer":
                    var footer = main.AddNewPart<FooterPart>();
                    footer.PutXDocument(new XDocument(new XElement(W.ftr, new XAttribute(XNamespace.Xmlns + "w", W.w),
                        Chain(W.customXml, depth - 4, LargeDocuments.Paragraph("Deep text")))));
                    sectPr = new XElement(W.sectPr, new XElement(W.footerReference,
                        new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(footer))));
                    break;
                case "footnotes":
                    main.AddNewPart<FootnotesPart>().PutXDocument(new XDocument(new XElement(W.footnotes,
                        new XAttribute(XNamespace.Xmlns + "w", W.w), new XElement(W.footnote, new XAttribute(W.id, 1),
                            Chain(W.customXml, depth - 5, LargeDocuments.Paragraph("Deep text"))))));
                    break;
                case "endnotes":
                    main.AddNewPart<EndnotesPart>().PutXDocument(new XDocument(new XElement(W.endnotes,
                        new XAttribute(XNamespace.Xmlns + "w", W.w), new XElement(W.endnote, new XAttribute(W.id, 1),
                            Chain(W.customXml, depth - 5, LargeDocuments.Paragraph("Deep text"))))));
                    break;
                case "comments":
                    main.AddNewPart<WordprocessingCommentsPart>().PutXDocument(new XDocument(new XElement(W.comments,
                        new XAttribute(XNamespace.Xmlns + "w", W.w), new XElement(W.comment, new XAttribute(W.id, 0),
                            new XAttribute(W.author, "Reviewer A"),
                            Chain(W.customXml, depth - 5, LargeDocuments.Paragraph("Deep text"))))));
                    break;
                case "styles":
                    stylesRoot.Add(Chain(X + "n", depth - 2, Leaf()));
                    break;
                case "numbering":
                    main.AddNewPart<NumberingDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.numbering,
                        new XAttribute(XNamespace.Xmlns + "w", W.w), Chain(X + "n", depth - 2, Leaf()))));
                    break;
                case "glossary":
                    main.AddNewPart<GlossaryDocumentPart>().PutXDocument(new XDocument(new XElement(W.glossaryDocument,
                        new XAttribute(XNamespace.Xmlns + "w", W.w), new XElement(W.docParts, new XElement(W.docPart,
                            new XElement(W.docPartBody, Chain(W.customXml, depth - 7, LargeDocuments.Paragraph("Deep text"))))))));
                    break;
                case "settings":
                    settings.Add(Chain(X + "n", depth - 2, Leaf()));
                    break;
                case "commentsExtended":
                    main.AddNewPart<WordprocessingCommentsExPart>().PutXDocument(new XDocument(new XElement(W15 + "commentsEx",
                        new XAttribute(XNamespace.Xmlns + "w15", W15), Chain(X + "n", depth - 2, Leaf()))));
                    break;
                case "theme":
                    var theme = SampleTheme();
                    theme.Add(Chain(X + "n", depth - 2, Leaf()));
                    main.AddNewPart<ThemePart>().PutXDocument(new XDocument(theme));
                    break;
                case "fontTable":
                    main.AddNewPart<FontTablePart>().PutXDocument(new XDocument(new XElement(W.fonts,
                        new XAttribute(XNamespace.Xmlns + "w", W.w), Chain(X + "n", depth - 2, Leaf()))));
                    break;
                case "chart":
                    main.AddNewPart<ChartPart>().PutXDocument(new XDocument(new XElement(C + "chartSpace",
                        new XAttribute(XNamespace.Xmlns + "c", C), Chain(X + "n", depth - 2, Leaf()))));
                    break;
                case "custom XML data the compatibility normalizer rewrites":
                    // root(1) AlternateContent(2) Fallback(3), then the chain: DocxDiff's markup
                    // compatibility pass resolves it in every .xml entry, recursively.
                    main.AddCustomXmlPart(CustomXmlPartType.CustomXml).PutXDocument(new XDocument(
                        new XElement(X + "root", new XAttribute(XNamespace.Xmlns + "mc", Mc), new XElement(Mc + "AlternateContent",
                            new XElement(Mc + "Fallback", Chain(X + "n", depth - 4, Leaf()))))));
                    break;
                case "altChunk part":
                    var chunk = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                    XNamespace xhtml = "http://www.w3.org/1999/xhtml";
                    chunk.PutXDocument(new XDocument(new XElement(xhtml + "html",
                        new XElement(xhtml + "body", Chain(xhtml + "div", depth - 3, new XElement(xhtml + "p", "deep"))))));
                    blocks.Add(new XElement(W.altChunk, new XAttribute(R.id, main.GetIdOfPart(chunk))));
                    break;
                case "custom XML data":
                    main.AddCustomXmlPart(CustomXmlPartType.CustomXml).PutXDocument(new XDocument(
                        new XElement(X + "root", Chain(X + "n", depth - 2, Leaf()))));
                    break;
            }
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(settings));
            stylesPart.PutXDocument(new XDocument(stylesRoot));
            blocks.Add(LargeDocuments.Paragraph(closing));
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w), new XAttribute(XNamespace.Xmlns + "r", R.r),
                new XElement(W.body, blocks, sectPr))));
        }
        return stream.ToArray();
    }

    /// <summary>A Strict-conformance package whose body holds a drawing with a deep shape:
    /// document(1) body(2) p(3) r(4) drawing(5) inline(6) graphic(7) graphicData(8) wsp(9), then the chain.</summary>
    private static byte[] Strict(int depth, string closing)
    {
        XNamespace sw = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        XNamespace swp = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
        XNamespace sa = "http://purl.oclc.org/ooxml/drawingml/main";
        XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        XElement P(string text) => new(sw + "p", new XElement(sw + "r", new XElement(sw + "t", text)));
        var drawing = new XElement(sw + "p", new XElement(sw + "r", new XElement(sw + "drawing",
            new XElement(swp + "inline",
                new XElement(swp + "extent", new XAttribute("cx", 914400), new XAttribute("cy", 914400)),
                new XElement(swp + "docPr", new XAttribute("id", 1), new XAttribute("name", "Shape 1")),
                new XElement(sa + "graphic", new XElement(sa + "graphicData",
                    new XAttribute("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
                    new XElement(wps + "wsp", Chain(wps + "extra", depth - 10, new XElement(wps + "leaf")))))))));
        var documentXml = new XDocument(new XElement(sw + "document",
            new XAttribute(XNamespace.Xmlns + "w", sw), new XAttribute(XNamespace.Xmlns + "wp", swp),
            new XAttribute(XNamespace.Xmlns + "a", sa), new XAttribute(XNamespace.Xmlns + "wps", wps),
            new XElement(sw + "body", drawing, P(closing))));
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using var writer = new StreamWriter(main.GetStream(FileMode.Create));
            writer.Write(documentXml.ToString(SaveOptions.DisableFormatting));
        }
        return stream.ToArray();
    }

    private static XElement SampleTheme()
    {
        using var source = WordprocessingDocument.Open(Path.Combine("../../../../TestFiles", "HW010-Tab01.docx"), false);
        return new XElement(source.MainDocumentPart!.ThemePart!.GetXDocument().Root!);
    }

    internal static byte[] RewriteEntry(byte[] package, string name, Func<string, string> rewrite) =>
        Edit(package, zip =>
        {
            var entry = zip.GetEntry(name)!;
            string text;
            using (var reader = new StreamReader(entry.Open())) text = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(name).Open());
            writer.Write(rewrite(text));
        });

    /// <summary>Stores <paramref name="name"/> as two interleaved pieces, which System.IO.Packaging
    /// reassembles into one part; neither piece alone is well-formed XML.</summary>
    internal static byte[] Interleave(byte[] package, string name) =>
        Edit(package, zip =>
        {
            var entry = zip.GetEntry(name)!;
            byte[] content;
            using (var source = entry.Open()) { using var copy = new MemoryStream(); source.CopyTo(copy); content = copy.ToArray(); }
            entry.Delete();
            int half = content.Length / 2;
            using (var first = zip.CreateEntry(name + "/[0].piece").Open()) first.Write(content, 0, half);
            using (var last = zip.CreateEntry(name + "/[1].last.piece").Open()) last.Write(content, half, content.Length - half);
        });

    /// <summary>Declares <paramref name="partName"/> <paramref name="contentType"/> by an Override,
    /// replacing any Override it had.</summary>
    internal static byte[] SetOverride(byte[] package, string partName, string contentType) =>
        RewriteEntry(package, "[Content_Types].xml", xml =>
        {
            var types = XDocument.Parse(xml);
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            types.Root!.Elements(ct + "Override")
                .Where(o => string.Equals((string?)o.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase))
                .Remove();
            types.Root.Add(new XElement(ct + "Override", new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
            return types.ToString(SaveOptions.DisableFormatting);
        });

    /// <summary>Adds an image part related from the main document, declared by an Override when
    /// <paramref name="typed"/>, else with no content type at all.</summary>
    internal static byte[] WithImage(byte[] package, string name, string content, bool typed)
    {
        package = AddEntry(package, name, content);
        package = RewriteEntry(package, "word/_rels/document.xml.rels", xml => xml.Replace("</Relationships>",
            "<Relationship Id=\"rIdDeep\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" " +
            $"Target=\"{name.Substring("word/".Length)}\"/></Relationships>", StringComparison.Ordinal));
        return typed ? SetOverride(package, "/" + name, "image/svg+xml") : package;
    }

    /// <summary>Adds a custom XML data part related from the main document, holding
    /// <paramref name="content"/> or, when null, <paramref name="spaces"/> spaces in a root element.</summary>
    internal static byte[] WithCustomXml(byte[] package, string name, string? content, long spaces = 0)
    {
        package = content is not null ? AddEntry(package, name, content) : AddSpaces(package, name, spaces);
        var id = "rId" + Math.Abs(name.GetHashCode());
        return RewriteEntry(package, "word/_rels/document.xml.rels", xml => xml.Replace("</Relationships>",
            $"<Relationship Id=\"{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" " +
            $"Target=\"../{name}\"/></Relationships>", StringComparison.Ordinal));
    }

    /// <summary>Adds an entry holding <c>&lt;s&gt;</c> plus <paramref name="count"/> spaces plus
    /// <c>&lt;/s&gt;</c>, deflated (it inflates over a thousand times).</summary>
    internal static byte[] AddSpaces(byte[] package, string name, long count) =>
        Edit(package, zip =>
        {
            using var stream = zip.CreateEntry(name, System.IO.Compression.CompressionLevel.SmallestSize).Open();
            stream.Write("<s>"u8);
            var block = new byte[1 << 20];
            Array.Fill(block, (byte)' ');
            for (long written = 0; written < count; written += block.Length) stream.Write(block, 0, (int)Math.Min(block.Length, count - written));
            stream.Write("</s>"u8);
        });

    internal static byte[] AddEntry(byte[] package, string name, string content) =>
        Edit(package, zip =>
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
            writer.Write(content);
        });

    /// <summary>Moves the main part to <paramref name="to"/>, keeping its declared content type and
    /// the package relationship that reaches it.</summary>
    internal static byte[] RenameEntry(byte[] package, string from, string to)
    {
        package = Edit(package, zip =>
        {
            var entry = zip.GetEntry(from)!;
            using var source = entry.Open();
            using (var target = zip.CreateEntry(to).Open()) source.CopyTo(target);
            source.Dispose();
            entry.Delete();
        });
        package = SetOverride(package, "/" + to,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
        return RewriteEntry(package, "_rels/.rels", xml => xml.Replace(from, to, StringComparison.Ordinal));
    }

    private static byte[] Edit(byte[] package, Action<System.IO.Compression.ZipArchive> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(package);
        using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
            edit(zip);
        return stream.ToArray();
    }

    /// <summary>Two stored (uncompressed) packages of exactly the same length: one whose body wraps a
    /// paragraph in <paramref name="levels"/> nested w:customXml, one with the same wrappers side by side.</summary>
    internal static (byte[] Shallow, byte[] Deep) SameLengthPair(int levels)
    {
        const string open = "<w:customXml w:element=\"x\">", close = "</w:customXml>";
        string Package(string body)
        {
            return "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                body + "<w:p><w:r><w:t>Closing text</w:t></w:r></w:p></w:body></w:document>";
        }
        byte[] Zip(string document)
        {
            using var stream = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                void Add(string name, string text)
                {
                    using var writer = new StreamWriter(zip.CreateEntry(name, System.IO.Compression.CompressionLevel.NoCompression).Open(),
                        new System.Text.UTF8Encoding(false));
                    writer.Write(text);
                }
                Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
                Add("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
                Add("word/document.xml", document);
            }
            return stream.ToArray();
        }
        var deep = string.Concat(Enumerable.Repeat(open, levels)) + string.Concat(Enumerable.Repeat(close, levels));
        var shallow = string.Concat(Enumerable.Repeat(open + close, levels));
        return (Zip(Package(shallow)), Zip(Package(deep)));
    }

    internal static string? HeaderEntryName(byte[] package)
    {
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package));
        return zip.Entries.FirstOrDefault(e => e.FullName == "word/header1.xml")?.FullName;
    }

    /// <summary>The deepest element nesting across the package's XML entries.</summary>
    internal static int MaxDepth(byte[] bytes)
    {
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        int deepest = 0;
        var parts = zip.Entries
            .Where(e => e.FullName.EndsWith(".xml", StringComparison.Ordinal) || e.FullName.EndsWith(".xhtml", StringComparison.Ordinal)
                || e.FullName.EndsWith(".piece", StringComparison.Ordinal))
            .GroupBy(e => System.Text.RegularExpressions.Regex.Replace(e.FullName, "/\\[\\d+\\](\\.last)?\\.piece$", ""));
        foreach (var part in parts)
        {
            var text = string.Concat(part.OrderBy(e => e.FullName, StringComparer.Ordinal).Select(e =>
            {
                using var reader = new StreamReader(e.Open());
                return reader.ReadToEnd();
            })).Replace(" w:conformance=\"strict\"", "", StringComparison.Ordinal);
            using var xml = XmlReader.Create(new StringReader(text));
            while (xml.Read())
                if (xml.NodeType == XmlNodeType.Element) deepest = Math.Max(deepest, xml.Depth + 1);
        }
        return deepest;
    }
}

internal static class LargeDocuments
{
    internal static XElement Paragraph(string text) => new(W.p, new XElement(W.r, new XElement(W.t, text)));

    /// <summary>Tables each nested in the previous one's only cell, as a fuzzer writes them:
    /// <c>3 * levels + 5</c> elements deep.</summary>
    internal static byte[] NestedTables(int levels, string closing = "After the tables") =>
        Build(new[] { NestedTablesElement(levels), Paragraph(closing) });

    internal static XElement NestedTablesElement(int levels)
    {
        XElement content = Paragraph("Innermost text");
        for (int level = levels; level > 0; level--)
            content = new XElement(W.tbl,
                new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, 2400))),
                new XElement(W.tr, new XElement(W.tc, Paragraph($"Level {level}"), content)));
        return content;
    }

    internal static int Depth(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var package = WordprocessingDocument.Open(stream, false);
        using var reader = XmlReader.Create(package.MainDocumentPart!.GetStream());
        int deepest = 0;
        while (reader.Read())
            if (reader.NodeType == XmlNodeType.Element) deepest = Math.Max(deepest, reader.Depth + 1);
        return deepest;
    }

    internal static byte[] Build(IEnumerable<XElement> blocks, XElement? header = null, XElement? customXml = null,
        XElement? glossary = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            XElement? sectPr = null;
            if (header is not null)
            {
                var headerPart = main.AddNewPart<HeaderPart>();
                headerPart.PutXDocument(new XDocument(new XElement(W.hdr,
                    new XAttribute(XNamespace.Xmlns + "w", W.w), header)));
                sectPr = new XElement(W.sectPr, new XElement(W.headerReference,
                    new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(headerPart))));
            }
            if (customXml is not null)
                main.AddCustomXmlPart(CustomXmlPartType.CustomXml).PutXDocument(new XDocument(customXml));
            if (glossary is not null)
                main.AddNewPart<GlossaryDocumentPart>().PutXDocument(new XDocument(new XElement(W.glossaryDocument,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XElement(W.docParts, new XElement(W.docPart, new XElement(W.docPartBody, glossary))))));
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w), new XAttribute(XNamespace.Xmlns + "r", R.r),
                new XElement(W.body, blocks, sectPr))));
        }
        return stream.ToArray();
    }
}
