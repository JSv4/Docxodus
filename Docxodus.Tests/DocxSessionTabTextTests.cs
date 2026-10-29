using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Tests.Ir;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>A <c>w:tab</c> is one character, <c>'\t'</c>, in the text every offset counts, every
/// search reads and the projection prints (issue #1387).</summary>
public class DocxSessionTabTextTests
{
    // "1.2<tab/>Alpha" in one run, the shape of a manually numbered heading.
    private static XElement NumberedRun() => new(W.r,
        new XElement(W.rPr, new XElement(W.i)),
        new XElement(W.t, "1.2"), new XElement(W.tab), new XElement(W.t, "Alpha"));

    [Fact]
    public void Grep_FindsATab_AndAPhraseAcrossIt()
    {
        using var session = Open(NumberedRun());

        var tab = Assert.Single(session.Grep("\t"));
        Assert.Equal(new CharSpan(3, 1), tab.Span);
        var phrase = Assert.Single(session.Grep(@"2\tAlp"));
        Assert.Equal(new CharSpan(2, 5), phrase.Span);
        Assert.Equal("2\tAlp", phrase.Text);
        Assert.Empty(session.Grep("2Alp"));
    }

    [Theory]
    [InlineData("# \tTitle\t  ", "\tTitle\t")]
    [InlineData("- \tItem\t  ", "\tItem\t")]
    [InlineData("2. \tItem\t  ", "\tItem\t")]
    [InlineData("# Title  ", "Title")]
    [InlineData("- Item  ", "Item")]
    public void ReplaceText_HeadingAndListPayloadsKeepEdgeTabsButTrimPadding(string payload, string expected)
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "Original")));

        AssertOk(session.ReplaceText(Paragraph(session), payload));

        Assert.Single(session.Grep("^" + System.Text.RegularExpressions.Regex.Escape(expected) + "$"));
        Assert.Equal(expected.Count(c => c == '\t'), Body(session.Save()).Descendants(W.tab).Count());
    }

    [Theory]
    [InlineData("\tValue\t", "| \tValue\t |")]
    [InlineData("  Value  ", "| Value |")]
    public void Project_TableCellsKeepEdgeTabsButTrimPadding(string text, string expected)
    {
        var table = new XElement(W.tbl, new XElement(W.tr, new XElement(W.tc,
            new XElement(W.p, new XElement(W.r, DocxSession.TextContent(text, W.t))))));
        var bytes = IrTestDocuments.FromBodyXml(table.ToString(SaveOptions.DisableFormatting)).DocumentByteArray;
        var settings = new WmlToMarkdownConverterSettings();

        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("tabs.docx", bytes), settings);
        var ir = IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument("tabs.docx", bytes)), settings);

        Assert.Contains(expected, oracle.Markdown);
        Assert.Equal(oracle.Markdown, ir.Markdown);
    }

    [Fact]
    public void Project_PrintsATab_SoAPhraseCopiedFromItMatches_InBothEmitters()
    {
        var bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "HW010-Tab01.docx"));
        using var session = new DocxSession(bytes);
        var line = session.Project().Markdown.Split('\n').Single(l => l.Contains("Text1"));
        Assert.Contains("Text1\tText2", line);
        Assert.Single(session.Grep(System.Text.RegularExpressions.Regex.Escape("Text1\tText2")));

        var settings = new WmlToMarkdownConverterSettings();
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("a.docx", bytes), settings);
        var ir = IrMarkdownEmitter.Emit(
            IrReader.Read(new WmlDocument("a.docx", bytes), new IrReaderOptions { RetainSources = false }),
            settings).ToProjection();
        Assert.Contains("Text1\tText2", oracle.Markdown);
        Assert.Equal(oracle.Markdown, ir.Markdown);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReplaceMatch_AcrossATab_WithATabInTheReplacement_LeavesOneTabInPlace(bool tracked)
    {
        using var session = Open(NumberedRun(), tracked);
        var match = Assert.Single(session.Grep(@"2\tAlpha"));
        AssertOk(session.ReplaceMatch(match, "2\tBeta"));
        AssertEdited(session, tracked, "1.2|Beta");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReplaceMatch_AcrossATab_WithoutATabInTheReplacement_ConsumesTheTab(bool tracked)
    {
        using var session = Open(NumberedRun(), tracked);
        var match = Assert.Single(session.Grep(@"2\tAlpha"));
        AssertOk(session.ReplaceMatch(match, "2 Beta"));
        AssertEdited(session, tracked, "1.2 Beta");
    }

    [Theory]
    [InlineData(3, 1, "X", "1.2XAlpha")]  // the tab alone
    [InlineData(3, 2, "X", "1.2Xlpha")]   // starts at the tab
    [InlineData(2, 2, "X", "1.XAlpha")]   // ends at the tab
    [InlineData(2, 3, "", "1.lpha")]      // a deletion covering it
    [InlineData(3, 0, "X", "1.2X|Alpha")] // an insertion just before it
    [InlineData(4, 0, "X", "1.2|XAlpha")] // and just after it
    [InlineData(1, 1, "X", "1X2|Alpha")]  // control: the tab is untouched
    [InlineData(5, 1, "X", "1.2|AXpha")]  // control: after the tab
    public void ReplaceTextAtSpan_CountsTheTabAsOneCharacter(int start, int length, string replace, string expected)
    {
        foreach (var tracked in new[] { false, true })
        {
            using var session = Open(NumberedRun(), tracked);
            AssertOk(session.ReplaceTextAtSpan(Paragraph(session), start, length, replace));
            AssertEdited(session, tracked, expected);
        }
    }

    [Fact]
    public void InsertedText_WritesATabAsATabElement()
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "Name")));
        AssertOk(session.ReplaceTextAtSpan(Paragraph(session), 4, 0, "\tValue"));
        AssertOk(session.ReplaceTextAtSpan(Paragraph(session), 0, 0, "1\t"));
        Assert.Equal("1|Name|Value", Shape(session.Save()));
        Assert.DoesNotContain(Body(session.Save()).Descendants(W.t), t => t.Value.Contains('\t'));

        using var markdown = Open(new XElement(W.r, new XElement(W.t, "Old")));
        AssertOk(markdown.ReplaceText(Paragraph(markdown), "Left\tRight"));
        Assert.Equal("Left|Right", Shape(markdown.Save()));
        Assert.Single(markdown.Grep(@"Left\tRight"));
        AssertOk(markdown.ReplaceText(Paragraph(markdown), "\tLead"));
        Assert.Equal("|Lead", Shape(markdown.Save()));
    }

    [Fact]
    public void ApplyFormat_AcrossATab_KeepsTheTabInPlace()
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "1.2"), new XElement(W.tab), new XElement(W.t, "Alpha")));
        AssertOk(session.ApplyFormat(Paragraph(session), new CharSpan(2, 3), new FormatOp { Bold = true }));
        Assert.Equal("1.2|Alpha", Shape(session.Save()));
        var bold = Body(session.Save()).Descendants(W.r).Where(r => r.Element(W.rPr)?.Element(W.b) is not null);
        Assert.Equal("2|A", string.Concat(bold.SelectMany(r => r.Elements()).Select(Char)));

        using var substring = Open(new XElement(W.r, new XElement(W.t, "1.2"), new XElement(W.tab), new XElement(W.t, "Alpha")));
        AssertOk(substring.ApplyFormatToSubstring(Paragraph(substring), "2\tA", new FormatOp { Bold = true }));
        Assert.Equal("2|A", string.Concat(Body(substring.Save()).Descendants(W.r)
            .Where(r => r.Element(W.rPr)?.Element(W.b) is not null).SelectMany(r => r.Elements()).Select(Char)));
    }

    [Fact]
    public void ApplyFormat_InARunWithABreak_KeepsTheBreakInPlace()
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "ab"), new XElement(W.br), new XElement(W.t, "cd")));
        AssertOk(session.ApplyFormat(Paragraph(session), new CharSpan(1, 1), new FormatOp { Bold = true }));
        Assert.Equal("ab/cd", string.Concat(Body(session.Save()).Descendants(W.r).SelectMany(r => r.Elements())
            .Select(e => e.Name == W.br ? "/" : Char(e))));
    }

    [Fact]
    public void AnnotationSpan_AcrossATab_CoversTheTab()
    {
        using var session = Open(NumberedRun());
        var match = Assert.Single(session.Grep(@"2\tAl"));
        AssertOk(session.AddAnnotation(match.EnclosingAnchor.Anchor.Id, match.Span,
            new DocumentAnnotation("tab-1", "L", "Label", "#FFEB3B")));
        Assert.Equal("2\tAl", Assert.Single(session.ListAnnotations()).AnnotatedText);
        Assert.Equal("1.2|Alpha", Shape(session.Save()));
    }

    [Fact]
    public void ExpectedTextRange_BuiltFromAMatchAcrossATab_Holds()
    {
        using var session = Open(NumberedRun());
        var match = Assert.Single(session.Grep(@"2\tAlpha"));
        var anchor = match.EnclosingAnchor.Anchor.Id;
        Assert.Equal("1.2\tAlpha", session.GetAnchorInfo(anchor)!.VisibleText);
        var result = session.ExecuteMutation(
            new MutationPreconditions
            {
                AnchorId = anchor,
                ExpectedTextRange = new TextRangePrecondition(match.Span.Start, match.Span.Length, match.Text),
            },
            s => s.ReplaceMatch(match, "2\tBeta"));
        AssertOk(result);
        Assert.Equal("1.2|Beta", Shape(session.Save()));
    }

    [Fact]
    public void FillPlaceholders_CoalescingBesideATab_KeepsTheTab()
    {
        using var session = Open(new XElement(W.r,
            new XElement(W.t, "Name:"), new XElement(W.tab), new XElement(W.t, "[_____]"), new XElement(W.tab), new XElement(W.t, "Date")));
        var result = session.FillPlaceholders(_ => string.Empty,
            new FillOptions { CoalesceWhitespaceAroundEmptyFill = true });
        Assert.Equal(1, result.Filled);
        Assert.Equal("Name:||Date", Shape(session.Save()));
    }

    [Fact]
    public void TabStopDefinitions_AndDeletedTabs_TakeNoRoom()
    {
        using var session = Open(new object[]
        {
            new XElement(W.pPr, new XElement(W.tabs, new XElement(W.tab,
                new XAttribute(W.val, "left"), new XAttribute(W.pos, "720")))),
            new XElement(W.r, new XElement(W.t, "Live")),
            new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Reviewer A"),
                new XElement(W.r, new XElement(W.tab), new XElement(W.delText, "Gone"))),
            new XElement(W.r, new XElement(W.t, "Text")),
        });
        var anchor = Paragraph(session);
        Assert.Equal("LiveText", session.GetAnchorInfo(anchor)!.VisibleText);
        Assert.Empty(session.Grep("\t"));
        Assert.Single(session.Grep("LiveText"));
    }

    // ---- review round: readers and writers outside the span ops ----------------------------

    // "ab", a tracked deletion of a tab and "gone", then "cd".
    private static object[] DeletedTabBetween() => new object[]
    {
        new XElement(W.r, new XElement(W.t, "ab")),
        new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Reviewer A"),
            new XElement(W.r, new XElement(W.tab), new XElement(W.delText, "gone"))),
        new XElement(W.r, new XElement(W.t, "cd")),
    };

    [Fact]
    public void Project_TableCellText_LeavesADeletedTabOut_InBothEmitters()
    {
        var bytes = IrTestDocuments.FromBodyXml("""
            <w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>
              <w:r><w:t>a</w:t></w:r><w:del w:id="1" w:author="Reviewer A"><w:r><w:tab/><w:delText>x</w:delText></w:r></w:del><w:r><w:t>b</w:t></w:r>
            </w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t><w:tab/><w:t>d</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
            <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
            """).DocumentByteArray;
        var settings = new WmlToMarkdownConverterSettings();
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("a.docx", bytes), settings);
        var ir = IrMarkdownEmitter.Emit(
            IrReader.Read(new WmlDocument("a.docx", bytes), new IrReaderOptions { RetainSources = false }),
            settings).ToProjection();
        Assert.Contains("| ab |", oracle.Markdown);
        Assert.Contains("c\td", oracle.Markdown); // control: a live tab in a cell prints
        Assert.Equal(oracle.Markdown, ir.Markdown);
    }

    [Fact]
    public void AnnotationText_LeavesADeletedTabOut()
    {
        using var session = Open(DeletedTabBetween());
        AssertOk(session.AddAnnotation(Paragraph(session), new CharSpan(1, 2),
            new DocumentAnnotation("del-1", "L", "Label", "#FFEB3B")));
        Assert.Equal("bc", Assert.Single(session.ListAnnotations()).AnnotatedText);
    }

    [Theory]
    [InlineData("1.2\tAlpha", false)]
    [InlineData("2\tAl", false)]
    [InlineData("2\tAl", true)]
    public void AnnotationManager_SearchesTheTextWithItsTabs_AndKeepsTheRunWhole(string search, bool inElement)
    {
        var doc = new WmlDocument("a.docx", DocxSessionSpanRevisionTests.Build(NumberedRun()));
        var annotation = new DocumentAnnotation("search-1", "L", "Label", "#FFEB3B");
        var result = inElement
            ? AnnotationManager.AddAnnotation(doc, annotation, AnnotationTarget.SearchInElement("doc/p-0", search))
            : AnnotationManager.AddAnnotation(doc, annotation, AnnotationRange.FromSearch(search));
        Assert.Equal(search, Assert.Single(AnnotationManager.GetAnnotations(result)).AnnotatedText);
        Assert.Equal("1.2|Alpha", Shape(result.DocumentByteArray));
        Assert.Throws<InvalidOperationException>(() => AnnotationManager.AddAnnotation(doc, annotation,
            AnnotationRange.FromSearch(search.Replace("\t", ""))));
    }

    [Theory]
    [InlineData("lph")]   // inside one text node of a run that also holds "1.2" and a tab
    [InlineData("Alpha")] // a whole text node after the tab
    public void AnnotationManager_SearchInsideAMixedRun_KeepsTheRestOfTheRun(string search)
    {
        var doc = new WmlDocument("a.docx", DocxSessionSpanRevisionTests.Build(NumberedRun()));
        var result = AnnotationManager.AddAnnotation(doc, new DocumentAnnotation("search-2", "L", "Label", "#FFEB3B"),
            AnnotationRange.FromSearch(search));
        Assert.Equal(search, Assert.Single(AnnotationManager.GetAnnotations(result)).AnnotatedText);
        Assert.Equal("1.2|Alpha", Shape(result.DocumentByteArray));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnnotationManager_SearchSkipsADeletedTab(bool inElement)
    {
        var doc = new WmlDocument("a.docx", DocxSessionSpanRevisionTests.Build(new object[]
        {
            new XElement(W.r, new XElement(W.t, "ab")),
            new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Reviewer A"),
                new XElement(W.r, new XElement(W.tab))),
            new XElement(W.r, new XElement(W.t, "cd")),
        }));
        WmlDocument Add(string search) => AnnotationManager.AddAnnotation(doc,
            new DocumentAnnotation("del-2", "L", "Label", "#FFEB3B"),
            inElement ? AnnotationTarget.SearchInElement("doc/p-0", search) : AnnotationTarget.Search(search));
        Assert.Equal("bc", Assert.Single(AnnotationManager.GetAnnotations(Add("bc"))).AnnotatedText);
        Assert.Throws<InvalidOperationException>(() => Add("b\tc"));
    }

    [Theory]
    [InlineData("br")]
    [InlineData("sym")]
    public void AnnotationManager_ContentBeforeTheMatch_StaysOutOfTheAnnotation(string marker)
    {
        var before = marker == "br" ? new XElement(W.br)
            : new XElement(W.sym, new XAttribute(W.font, "Symbol"), new XAttribute(W._char, "F0B7"));
        var doc = new WmlDocument("a.docx", DocxSessionSpanRevisionTests.Build(
            new XElement(W.r, new XElement(W.t, "Line1"), before, new XElement(W.t, "Line2"))));
        var result = AnnotationManager.AddAnnotation(doc, new DocumentAnnotation("zw-1", "L", "Label", "#FFEB3B"),
            AnnotationRange.FromSearch("Line2"));
        Assert.Equal("Line2", Assert.Single(AnnotationManager.GetAnnotations(result)).AnnotatedText);
        var order = Body(result.DocumentByteArray).Descendants()
            .Where(e => e.Name == W.w + marker || e.Name == W.bookmarkStart).Select(e => e.Name.LocalName);
        Assert.Equal(new[] { marker, "bookmarkStart" }, order);
    }

    [Fact]
    public void AnnotationManager_MatchAcrossABreakInOneRun_KeepsTheBreak()
    {
        var doc = new WmlDocument("a.docx", DocxSessionSpanRevisionTests.Build(
            new XElement(W.r, new XElement(W.t, "Line1"), new XElement(W.br), new XElement(W.t, "Line2"))));
        var result = AnnotationManager.AddAnnotation(doc, new DocumentAnnotation("br-1", "L", "Label", "#FFEB3B"),
            AnnotationRange.FromSearch("e1\nLi"));
        Assert.Equal("e1\nLi", Assert.Single(AnnotationManager.GetAnnotations(result)).AnnotatedText);
        Assert.Single(Body(result.DocumentByteArray).Descendants(W.br));
        Assert.Equal("Line1Line2", string.Concat(Body(result.DocumentByteArray).Descendants(W.t).Select(t => t.Value)));
    }

    // Moved-from text is not text a reader sees (issue #1282's rule, which every visible-text
    // reader now shares): a search finds the moved-to copy, and a cell leaves it out.
    [Fact]
    public void MovedFromText_IsLeftOutOfSearchAndCellText()
    {
        static XAttribute[] Stamp(int id) => new[] { new XAttribute(W.id, id), new XAttribute(W.author, "Reviewer A") };
        var body = new XElement(W.body,
            new XElement(W.p, new XElement(W.r, new XElement(W.t, "Before ")),
                new XElement(W.moveFrom, Stamp(1), new XElement(W.r, new XElement(W.t, "moved")))),
            new XElement(W.p, new XElement(W.moveTo, Stamp(2), new XElement(W.r, new XElement(W.t, "moved"))),
                new XElement(W.r, new XElement(W.t, " after"))),
            new XElement(W.tbl, new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W.w + "w", "2000"))),
                new XElement(W.tr, new XElement(W.tc, new XElement(W.p,
                    new XElement(W.r, new XElement(W.t, "keep ")),
                    new XElement(W.moveFrom, Stamp(3), new XElement(W.r, new XElement(W.t, "gone"))))))),
            new XElement(W.p, new XElement(W.r, new XElement(W.t, "Tail"))));
        var bytes = IrTestDocuments.FromBodyXml(string.Concat(body.Elements().Select(e => e.ToString()))).DocumentByteArray;

        var annotated = AnnotationManager.AddAnnotation(new WmlDocument("a.docx", bytes),
            new DocumentAnnotation("move-1", "L", "Label", "#FFEB3B"), AnnotationRange.FromSearch("moved"));
        var start = Body(annotated.DocumentByteArray).Descendants(W.bookmarkStart).Single();
        Assert.Contains(" after", start.Ancestors(W.p).Single().Value);

        var settings = new WmlToMarkdownConverterSettings();
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("a.docx", bytes), settings);
        var ir = IrMarkdownEmitter.Emit(
            IrReader.Read(new WmlDocument("a.docx", bytes), new IrReaderOptions { RetainSources = false }),
            settings).ToProjection();
        Assert.Contains("| keep |", oracle.Markdown);
        Assert.Contains("| keep |", ir.Markdown);
        Assert.Equal(WithoutAnchors(oracle.Markdown), WithoutAnchors(ir.Markdown));
    }

    private static string WithoutAnchors(string markdown) =>
        System.Text.RegularExpressions.Regex.Replace(markdown, @"\{#[^}]*\}", "");

    [Fact]
    public void TypingInOwnInsertion_JustBeforeATab_LandsBeforeIt()
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "Start")), tracked: true);
        AssertOk(session.ReplaceTextAtSpan(Paragraph(session), 5, 0, " ab\tcd"));
        AssertOk(session.ReplaceTextAtSpan(Paragraph(session), 8, 0, "X"));
        var saved = session.Save();
        Assert.Equal("Start abX|cd", Shape(RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", saved)).DocumentByteArray));
        Assert.Equal("Start", Shape(RevisionProcessor.RejectRevisions(new WmlDocument("r.docx", saved)).DocumentByteArray));
    }

    [Fact]
    public void InsertCrossReference_WritesTheBookmarkedTabAsATab()
    {
        using var session = Open(NumberedRun());
        var anchor = Paragraph(session);
        AssertOk(session.AddBookmark("Target", DocumentRange.In(anchor, new CharSpan(0, 9))));
        AssertOk(session.InsertCrossReference(anchor, 9, "Target"));
        Assert.DoesNotContain(Body(session.Save()).Descendants(W.t), t => t.Value.Contains('\t'));
        Assert.Equal("1.2|Alpha1.2|Alpha", Shape(session.Save()));
    }

    [Fact]
    public void ListComments_TextKeepsATab()
    {
        using var session = Open(NumberedRun());
        AssertOk(session.AddComment(Paragraph(session), new CharSpan(0, 3), "Reviewer A", "c\td"));
        Assert.Equal("c\td", Assert.Single(session.ListComments()).Text);
    }

    [Theory]
    [InlineData(2)] // the field's begin, just after the tab
    [InlineData(4)] // the field's end
    public void SplitParagraph_AtAFieldEdgeAfterATab_Succeeds(int offset)
    {
        // The field's marks share runs with the text either side, so the split has to cut them
        // free at the field's own offsets, which count the tab.
        static XElement Mark(string type) => new(W.fldChar, new XAttribute(W.fldCharType, type));
        using var session = Open(new object[]
        {
            new XElement(W.r, new XElement(W.t, "a"), new XElement(W.tab), Mark("begin")),
            new XElement(W.r, new XElement(W.instrText, " PAGE ")), new XElement(W.r, Mark("separate")),
            new XElement(W.r, new XElement(W.t, "12")),
            new XElement(W.r, Mark("end"), new XElement(W.t, "b")),
        });
        AssertOk(session.SplitParagraph(Paragraph(session), offset));
        Assert.Equal("a|12b", Shape(session.Save()));
    }

    [Fact]
    public void TrackedReplace_WritesALiteralTabInTheUntouchedTextAsATab()
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "a\tb c")), tracked: true);
        AssertOk(session.ReplaceTextAtSpan(Paragraph(session), 4, 1, "X"));
        Assert.DoesNotContain(Body(session.Save()).Descendants(W.t), t => t.Value.Contains('\t'));
        Assert.Equal("a|b X", Shape(RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", session.Save())).DocumentByteArray));
    }

    [Fact]
    public void Control_TabFreeText_KeepsItsOffsets()
    {
        using var session = Open(new XElement(W.r, new XElement(W.t, "Alpha Beta")));
        var match = Assert.Single(session.Grep("Beta"));
        Assert.Equal(new CharSpan(6, 4), match.Span);
        AssertOk(session.ReplaceMatch(match, "Gamma"));
        Assert.Equal("Alpha Gamma", Shape(session.Save()));
    }

    [Fact]
    public void RunTextMap_ForATabFreeBlock_AllocatesNoMoreThanBefore()
    {
        var paragraph = new XElement(W.p, Enumerable.Range(0, 40).Select(i =>
            new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, "Run text " + i))));
        RunTextMap.Build(paragraph); // warm up
        long before = GC.GetAllocatedBytesForCurrentThread();
        for (int i = 0; i < 100; i++) RunTextMap.Build(paragraph);
        long perBuild = (GC.GetAllocatedBytesForCurrentThread() - before) / 100;
        // What one Build of this block allocated before tabs took room in the text (measured on main).
        Assert.True(perBuild <= MainAllocatedBytesPerBuild, $"RunTextMap.Build allocated {perBuild} bytes");
    }

    private const long MainAllocatedBytesPerBuild = 12848;

    // ---- helpers --------------------------------------------------------------------------

    private static DocxSession Open(object content, bool tracked = false) =>
        new(DocxSessionSpanRevisionTests.Build(content), new DocxSessionSettings
        {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "Reviewer A",
        });

    private static string Paragraph(DocxSession session) =>
        session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;

    private static void AssertOk(EditResult result) => Assert.True(result.Success, result.Error?.Message);

    /// <summary>The live text untracked; under tracking, the accepted and rejected texts.</summary>
    private static void AssertEdited(DocxSession session, bool tracked, string expected)
    {
        var saved = session.Save();
        PackageValidation.AssertValid(saved, FileFormatVersions.Office2019);
        if (!tracked)
        {
            Assert.Equal(expected, Shape(saved));
            return;
        }
        Assert.Equal(expected, Shape(RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", saved)).DocumentByteArray));
        Assert.Equal("1.2|Alpha", Shape(RevisionProcessor.RejectRevisions(new WmlDocument("r.docx", saved)).DocumentByteArray));
    }

    private static XElement Body(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>The body's run content in order: text as itself, a tab as <c>|</c>.</summary>
    private static string Shape(byte[] bytes) =>
        string.Concat(Body(bytes).Descendants(W.r).SelectMany(r => r.Elements()).Select(Char));

    private static string Char(XElement e) => e.Name == W.t ? e.Value : e.Name == W.tab ? "|" : "";
}
