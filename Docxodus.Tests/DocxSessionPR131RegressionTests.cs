#nullable enable

// Regression tests for the bugs uncovered by the PR131 smoke test on the
// NVCA-Model-COI fixture. Each test pins one defect and (once the fix lands)
// becomes a guardrail. Numbering continues from DocxSessionTests in the
// "phase 10 — PR131 smoke fixes" range (DS080-DS099).

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionPR131RegressionTests
{
    // ─── Fixture builders ────────────────────────────────────────────────

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>Single body paragraph containing: text "AB" + hyperlink("CD") + text "EF".
    /// Full visible text is "ABCDEF" (6 chars).</summary>
    internal static byte[] BuildParagraphWithHyperlink()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var rel = main.AddHyperlinkRelationship(new System.Uri("https://example.com/cd"), true);

            var paraXml = new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "AB")),
                new XElement(W + "hyperlink",
                    new XAttribute(R + "id", rel.Id),
                    new XElement(W + "r",
                        new XElement(W + "t",
                            new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "CD"))),
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "EF")));

            var docXml = new XElement(W + "document",
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "w", W.NamespaceName),
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "r", R.NamespaceName),
                new XElement(W + "body", paraXml));
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    /// <summary>Single body paragraph wrapped by a bookmarkStart/End named "MARK1".
    /// Layout: bookmarkStart(id=1, name="MARK1") + run("hello world") + bookmarkEnd(id=1).</summary>
    internal static byte[] BuildParagraphWithBookmark()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var paraXml = new XElement(W + "p",
                new XElement(W + "bookmarkStart",
                    new XAttribute(W + "id", "1"),
                    new XAttribute(W + "name", "MARK1")),
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "hello world")),
                new XElement(W + "bookmarkEnd",
                    new XAttribute(W + "id", "1")));

            var docXml = new XElement(W + "document",
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "w", W.NamespaceName),
                new XElement(W + "body", paraXml));
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    /// <summary>Two body paragraphs: "First." and "Second has a hyperlink to [TXT]".
    /// Second paragraph contains: text "Second has a hyperlink to " + hyperlink("TXT").</summary>
    internal static byte[] BuildTwoParagraphsSecondWithHyperlink()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var rel = main.AddHyperlinkRelationship(new System.Uri("https://example.com/txt"), true);

            var p1 = new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "First.")));
            var p2 = new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "Second has a hyperlink to ")),
                new XElement(W + "hyperlink",
                    new XAttribute(R + "id", rel.Id),
                    new XElement(W + "r",
                        new XElement(W + "t",
                            new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "TXT"))));

            var docXml = new XElement(W + "document",
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "w", W.NamespaceName),
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "r", R.NamespaceName),
                new XElement(W + "body", p1, p2));
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    private static AnchorTarget FirstBodyParagraph(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind is "p" or "h" or "li" && t.Anchor.Scope == "body");

    private static System.Collections.Generic.List<AnchorTarget> BodyParagraphs(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind is "p" or "h" or "li" && t.Anchor.Scope == "body")
            .ToList();

    // ─── B4: bookmark preservation ───────────────────────────────────────

    [Fact]
    public void DS080_ReplaceText_PreservesBookmarks()
    {
        using var session = new DocxSession(BuildParagraphWithBookmark());
        var anchor = FirstBodyParagraph(session);

        var r = session.ReplaceText(anchor.Anchor.Id, "replaced text");
        Assert.True(r.Success, r.Error?.Message);

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var bookmarks = doc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W + "bookmarkStart")
            .Select(b => (string?)b.Attribute(W + "name"))
            .ToList();
        Assert.Contains("MARK1", bookmarks);
    }

    // ─── B1/B5: SplitParagraph hyperlink-aware ───────────────────────────

    [Fact]
    public void DS081_SplitParagraph_OffsetBetweenHyperlinkAndTrailingText()
    {
        // "AB" + hyperlink("CD") + "EF"  →  split at offset 4 (just after "CD")
        // Expected first = "ABCD" (hyperlink stays with first half),
        //          second = "EF".
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        Assert.Equal("ABCDEF", session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview);

        var r = session.SplitParagraph(anchor.Anchor.Id, 4);
        Assert.True(r.Success, $"split failed: {r.Error?.Code}/{r.Error?.Message}");

        var firstText = session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview;
        var secondText = session.GetAnchorInfo(r.Created[0].Id)?.TextPreview;
        Assert.Equal("ABCD", firstText);
        Assert.Equal("EF", secondText);
    }

    [Fact]
    public void DS082_SplitParagraph_OffsetAtPreviewLengthSucceeds()
    {
        // "AB" + hyperlink("CD") + "EF"  →  split at offset 6 (end of text)
        // ParagraphText (used to validate) MUST include hyperlink text,
        // otherwise the agent gets OffsetOutOfRange on a valid offset.
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        Assert.Equal(6, session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview.Length);

        var r = session.SplitParagraph(anchor.Anchor.Id, 6);
        Assert.True(r.Success, $"split at preview length must succeed; got {r.Error?.Code}/{r.Error?.Message}");
    }

    [Fact]
    public void DS083_SplitParagraph_OffsetInsideHyperlink()
    {
        // "AB" + hyperlink("CD") + "EF"  →  split at offset 3 (between C and D)
        // Expected: hyperlink content "CD" is split — first contains "ABC",
        // second contains "DEF".
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);

        var r = session.SplitParagraph(anchor.Anchor.Id, 3);
        Assert.True(r.Success, r.Error?.Message);

        var firstText = session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview;
        var secondText = session.GetAnchorInfo(r.Created[0].Id)?.TextPreview;
        Assert.Equal("ABC", firstText);
        Assert.Equal("DEF", secondText);
    }

    /// <summary>Single body paragraph: text "AA" + <c>w:ins</c>("BBBB") + text "CC".
    /// Visible text is "AABBBBCC" (8 chars); offsets 3-5 are strictly inside the
    /// tracked insertion, and 2 and 6 are its edges.</summary>
    internal static byte[] BuildParagraphWithTrackedInsertion()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            XElement Run(string text) => new(W + "r",
                new XElement(W + "t",
                    new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), text));

            var paraXml = new XElement(W + "p",
                Run("AA"),
                new XElement(W + "ins",
                    new XAttribute(W + "id", 10),
                    new XAttribute(W + "author", "Reviewer"),
                    new XAttribute(W + "date", "2026-01-01T00:00:00Z"),
                    Run("BBBB")),
                Run("CC"));

            main.Document = new Document(new Body());
            main.Document.Save();
            var xDoc = main.GetXDocument();
            xDoc.Root!.Element(W + "body")!.ReplaceNodes(paraXml);
            main.PutXDocument();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Companion to the note-citation guard (DS339/DS340). Once tracked insertions joined the
    /// visible-text walk, a split offset could land strictly inside <c>w:ins</c>. The run
    /// splitter cut the nested run, but <c>MoveInlineChildrenAfter</c> only relocates children
    /// whose START is at or past the offset — so the whole revision wrapper stayed with the
    /// first half and the paragraph silently split at the wrapper's far edge while reporting
    /// success. Splitting a revision wrapper needs revision semantics this op does not own, so
    /// it refuses instead.
    /// </summary>
    [Theory]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void DS095_SplitParagraph_InsideTrackedInsertion_FailsClosedWithoutMutation(int offset)
    {
        using var session = new DocxSession(BuildParagraphWithTrackedInsertion());
        var anchor = FirstBodyParagraph(session).Anchor.Id;
        Assert.Equal("AABBBBCC", session.GetAnchorInfo(anchor)?.TextPreview);
        var before = session.GetPackageContentHash();
        var version = session.Version;
        var undoCount = session.UndoCount;

        var result = session.SplitParagraph(anchor, offset);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.Equal(version, session.Version);
        Assert.Equal(undoCount, session.UndoCount);
        Assert.Single(BodyParagraphs(session));
        Assert.Equal("AABBBBCC", session.GetAnchorInfo(anchor)?.TextPreview);
    }

    [Theory]
    [InlineData(2, "AA", "BBBBCC")]
    [InlineData(6, "AABBBB", "CC")]
    public void DS096_SplitParagraph_AtTrackedInsertionEdge_SplitsExactly(
        int offset, string expectedFirst, string expectedSecond)
    {
        using var session = new DocxSession(BuildParagraphWithTrackedInsertion());
        var anchor = FirstBodyParagraph(session).Anchor.Id;

        var result = session.SplitParagraph(anchor, offset);

        Assert.True(result.Success, $"{result.Error?.Code}/{result.Error?.Message}");
        Assert.Equal(expectedFirst, session.GetAnchorInfo(anchor)?.TextPreview);
        Assert.Equal(expectedSecond, session.GetAnchorInfo(result.Created[0].Id)?.TextPreview);
    }

    // ─── B2: MergeParagraphs preserves hyperlinks ────────────────────────

    [Fact]
    public void DS084_MergeParagraphs_PreservesHyperlinkInSecondParagraph()
    {
        using var session = new DocxSession(BuildTwoParagraphsSecondWithHyperlink());
        var paras = BodyParagraphs(session);
        Assert.True(paras.Count >= 2);
        var firstId = paras[0].Anchor.Id;
        var secondId = paras[1].Anchor.Id;

        var r = session.MergeParagraphs(firstId, secondId);
        Assert.True(r.Success, r.Error?.Message);

        var merged = session.GetAnchorInfo(firstId)?.TextPreview ?? "";
        Assert.Contains("TXT", merged);
    }

    // ─── B3: MergeParagraphs separator ───────────────────────────────────

    [Fact]
    public void DS085_MergeParagraphs_InsertsSeparator_WhenBothEndsAreNonWhitespace()
    {
        // Use the simple two-paragraph fixture ("First paragraph." + "Second paragraph.")
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paras = BodyParagraphs(session);
        var firstId = paras[0].Anchor.Id;
        var secondId = paras[1].Anchor.Id;

        var r = session.MergeParagraphs(firstId, secondId);
        Assert.True(r.Success, r.Error?.Message);

        var merged = session.GetAnchorInfo(firstId)?.TextPreview;
        Assert.Equal("First paragraph. Second paragraph.", merged);
    }

    [Fact]
    public void DS086_MergeParagraphs_NoDoubleSpace_WhenFirstEndsWithWhitespace()
    {
        // Build a two-para fixture where the first ends with a trailing space.
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            main.Document.Body = new Body();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Body.Append(new Paragraph(new Run(new Text("First. ") { Space = SpaceProcessingModeValues.Preserve })));
            main.Document.Body.Append(new Paragraph(new Run(new Text("Second."))));
            main.Document.Save();
        }
        using var session = new DocxSession(ms.ToArray());
        var paras = BodyParagraphs(session);
        var r = session.MergeParagraphs(paras[0].Anchor.Id, paras[1].Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);
        var merged = session.GetAnchorInfo(paras[0].Anchor.Id)?.TextPreview;
        Assert.Equal("First. Second.", merged);
    }

    // ─── B5: ApplyFormat hyperlink-aware ─────────────────────────────────

    [Fact]
    public void DS087_ApplyFormat_AcceptsSpanCoveringHyperlinkText()
    {
        // Preview is 6 chars ("ABCDEF"). ApplyFormat(0, 6) must succeed.
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        var r = session.ApplyFormat(anchor.Anchor.Id, new CharSpan(0, 6), new FormatOp { Bold = true });
        Assert.True(r.Success, $"got {r.Error?.Code}/{r.Error?.Message}");
    }

    [Fact]
    public void DS088_ApplyFormat_FormatsRunsInsideHyperlink()
    {
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        var r = session.ApplyFormat(anchor.Anchor.Id, null, new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var hyperlinkRun = doc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W + "hyperlink").First()
            .Element(W + "r");
        var bold = hyperlinkRun?.Element(W + "rPr")?.Element(W + "b");
        Assert.NotNull(bold);
    }

    // ─── B6: hyperlink dedup ─────────────────────────────────────────────

    [Fact]
    public void DS089_ReplaceText_DedupesHyperlinkRelationship_SameUrl()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        for (int i = 0; i < 5; i++)
            session.ReplaceText(anchor.Anchor.Id, $"Round {i} text [link](https://example.com/same)");

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var rels = doc.MainDocumentPart!.HyperlinkRelationships
            .Where(rl => rl.Uri.ToString() == "https://example.com/same")
            .ToList();
        Assert.Single(rels);
    }

    // ─── F1: bullet payload kind ─────────────────────────────────────────

    [Fact]
    public void DS090_InsertParagraph_BulletPayload_CreatedKindMatchesProjection()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        // Fixture has no numbering definitions and no preceding list, so the bullet
        // payload cannot inherit numPr. The returned kind MUST match what the
        // projector then reports for the created anchor — i.e. "p".
        var r = session.InsertParagraph(anchor.Anchor.Id, Position.After, "- bullet payload");
        Assert.True(r.Success, r.Error?.Message);

        var createdAnchor = r.Created[0];
        var projection = session.Project();
        var projectorTarget = projection.AnchorIndex.Values.FirstOrDefault(t => t.Unid == createdAnchor.Unid);
        Assert.NotNull(projectorTarget);
        Assert.Equal(projectorTarget!.Anchor.Kind, createdAnchor.Kind);
    }

    // ─── Bluth-Co smoke-test regressions ─────────────────────────────────

    /// <summary>Heading2 paragraph with both Heading2 style AND a decimal numPr,
    /// so the projector shows "## 1. text" with the auto-number as part of the visible heading.</summary>
    private static byte[] BuildNumberedHeadingFixture()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles =
                new Styles(new Style { Type = StyleValues.Paragraph, StyleId = "Heading2", StyleName = new StyleName { Val = "heading 2" } });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var numPart = main.AddNewPart<NumberingDefinitionsPart>();
            numPart.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." },
                        new StartNumberingValue { Val = 1 })
                    { LevelIndex = 0 })
                { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });

            var pPr = new ParagraphProperties(
                new ParagraphStyleId { Val = "Heading2" },
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            body.Append(new Paragraph(pPr, new Run(new Text("The name of this corporation."))));

            main.Document.Save();
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS091_ReplaceText_StripsResolvedAutoNumberPrefix()
    {
        // Repro from the NVCA-COI smoke test: the projector emits a heading as
        // "## Fourth The total number..." (auto-number + space + run text). An
        // agent that echoes the visible heading back as its replacement payload
        // ends up with "Fourth Fourth: ..." in the saved DOCX because Word still
        // applies the auto-number and the run text now also starts with "Fourth".
        //
        // Fix contract: ReplaceText strips a leading prefix from the payload that
        // matches the paragraph's resolved auto-number (followed by an optional
        // single separator space/tab/NBSP). Idempotent — if the agent omits the
        // prefix, no stripping happens.
        using var s = new DocxSession(BuildNumberedHeadingFixture());
        var proj = s.Project();
        var headingTarget = proj.AnchorIndex.Values.Single(t => t.Anchor.Kind == "h");

        // Sanity check: projection shows the auto-number prefix inline.
        Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} ## 1\. The name of this corporation\.", proj.Markdown);

        // Agent's natural payload: echoes the visible "1. " prefix.
        var r = s.ReplaceText(headingTarget.Anchor.Id, "1. Replaced heading text.");
        Assert.True(r.Success, r.Error?.Message);

        // The next projection must show the prefix EXACTLY ONCE.
        var after = s.Project().Markdown;
        Assert.Contains("1. Replaced heading text.", after);
        Assert.DoesNotContain("1. 1. ", after);
    }

    [Fact]
    public void DS091b_ReplaceText_AutoNumberStripping_NoOpWhenPayloadHasNoPrefix()
    {
        // The fix must be invisible when the agent omits the prefix.
        using var s = new DocxSession(BuildNumberedHeadingFixture());
        var headingTarget = s.Project().AnchorIndex.Values.Single(t => t.Anchor.Kind == "h");

        var r = s.ReplaceText(headingTarget.Anchor.Id, "Bare replacement text.");
        Assert.True(r.Success, r.Error?.Message);

        var after = s.Project().Markdown;
        Assert.Contains("1. Bare replacement text.", after);
    }

    [Fact]
    public void DS092_RawReplaceXml_ReportsModifiedOnUnidPreservingRoundTrip()
    {
        // Repro from the NVCA-COI smoke test: the Get→mutate→Replace recipe
        // documented in docs/architecture/docx_mutation_api.md preserves the
        // element's Unid. The old impl always put target.Anchor in Removed AND
        // re-added the (same-Unid) element to Created, so an agent that pattern-
        // matched on Created/Removed saw a phantom "deleted and recreated"
        // operation. Fix: classify by unid set intersection — overlap is Modified.
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();

        var xml = s.Raw.GetXml(anchor);
        var mutated = xml.Replace("First paragraph.", "EDITED: First paragraph.");
        var r = s.Raw.ReplaceXml(anchor, mutated);

        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Id == anchor);
        Assert.DoesNotContain(r.Removed, a => a.Id == anchor);
        Assert.DoesNotContain(r.Created, a => a.Id == anchor);
    }

    [Fact]
    public void DS092b_RawReplaceXml_FreshXml_StillReportsRemovedAndCreated()
    {
        // The complementary case: when the replacement XML has fresh Unids
        // (because it didn't come from GetXml), the old anchor is gone and the
        // new element gets a brand-new anchor. Removed/Created — NOT Modified.
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = s.Project().AnchorIndex.Keys.First();
        var freshXml = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Fresh.</w:t></w:r></w:p>";

        var r = s.Raw.ReplaceXml(anchor, freshXml);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Removed, a => a.Id == anchor);
        Assert.Empty(r.Modified);
        Assert.NotEmpty(r.Created);
        Assert.DoesNotContain(r.Created, a => a.Id == anchor); // new Unid, not the old one
    }

    /// <summary>
    /// Unzips a DOCX in-memory and concatenates every text part's XML into one string,
    /// so substring assertions over attribute presence work without a compressed-bytes
    /// false-negative.
    /// </summary>
    private static string ReadAllPartXml(byte[] docxBytes)
    {
        using var stream = new MemoryStream(docxBytes);
        using var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read);
        var sb = new System.Text.StringBuilder();
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase)) continue;
            using var es = entry.Open();
            using var reader = new StreamReader(es);
            sb.Append(reader.ReadToEnd());
        }
        return sb.ToString();
    }

    [Fact]
    public void DS093_Save_StripsInternalUnidsByDefault()
    {
        // Repro from the NVCA-COI smoke test: a 148 KB input round-tripped
        // through DocxSession came back as 588 KB (4×). The bloat was 14,000+
        // PtOpenXml:Unid attributes (~50 bytes each) persisted into the saved
        // DOCX. Per the comment in RawDocxOps.cs, this attribute is "internal-
        // only ... not in the OOXML schema" — so Save() now strips it from all
        // parts by default. Opt back in with DocxSessionSettings.PersistAnchorIds.
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        s.Project(); // assigns Unids
        var saved = s.Save();

        var allXml = ReadAllPartXml(saved);
        Assert.DoesNotContain("PtOpenXml:Unid", allXml);
        // The string `:Unid="` is how the attribute serializes (with the
        // namespace prefix). Tight enough to skip identifier text like "Unique".
        Assert.DoesNotContain(":Unid=\"", allXml);

        // Session must still operate correctly after Save (in-memory Unids restored).
        var proj = s.Project();
        Assert.Contains("First paragraph.", proj.Markdown);
        Assert.True(proj.AnchorIndex.Count >= 2);

        // Re-loading the stripped bytes must produce a working session
        // (with fresh Unids — anchor ids change across save/reload by default).
        using var s2 = new DocxSession(saved);
        Assert.Contains("First paragraph.", s2.Project().Markdown);
    }

    [Fact]
    public void DS094_Save_PersistAnchorIds_OptInKeepsUnids()
    {
        // The escape hatch for callers that need anchor ids to survive
        // save/reload (matches the spec's open-question scenario).
        var settings = new DocxSessionSettings { PersistAnchorIds = true };
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), settings);
        s.Project();
        var saved = s.Save();

        var allXml = ReadAllPartXml(saved);
        Assert.Contains(":Unid=\"", allXml);
    }
}
