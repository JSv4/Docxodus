#nullable enable
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Native split/merge composition in the consolidate engine (the split/merge counterpart to
/// <see cref="IrCompositeMoveTests"/>). A reviewer's UNCONTESTED paragraph split/merge — the reviewer is
/// the sole toucher of every base paragraph the op consumes — passes through the composite merge as a
/// NATIVE <see cref="IrEditOpKind.SplitBlock"/>/<see cref="IrEditOpKind.MergeBlock"/> and renders as the
/// same paragraph-mark revisions as Compare. A contested split competes as one result at its base
/// paragraph. A contested merge lowers its consumed paragraphs to whole-block operations.
/// </summary>
public class IrCompositeSplitMergeTests
{
    // Sentence-shaped paragraphs (the proven split/merge detection shape: the segmenter needs sentence
    // boundaries) plus an anchor paragraph that keeps surrounding structure stable.
    private const string SplitSrc = "aaa bbb ccc ddd. eee fff ggg hhh.";
    private const string SplitA = "aaa bbb ccc ddd. ";
    private const string SplitB = "eee fff ggg hhh.";
    private const string M1 = "Alpha alpha words here.";
    private const string M2 = "Beta beta words here.";
    private const string Anchor = "anchor one two three four five.";

    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    private static IrCompositeScript Merge(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
    {
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Doc, ReadOpts))).ToList();
        return IrCompositeMerger.Merge(baseIr, revIr, policy, diff);
    }

    private static WmlDocument Consolidate(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.Consolidate(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static System.Collections.Generic.IReadOnlyList<DocxDiffConflict> Conflicts(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConflicts(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    // ---- 1. Uncontested split composes natively ----

    [Fact]
    public void Uncontested_split_survives_as_native_SplitBlock_authored_to_splitter()
    {
        var b = Docs.Para(SplitSrc, Anchor);
        var alice = Docs.Para(SplitA, SplitB, Anchor);            // Alice splits paragraph 1
        var bob = Docs.Para(SplitSrc, Anchor + " BOB tail");      // Bob edits the OTHER paragraph

        var s = Merge(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));

        var split = Assert.Single(s.Operations, o => o.Op.Kind == IrEditOpKind.SplitBlock);
        Assert.Equal("Alice", split.Author);
        Assert.Equal(2, split.Op.SplitMergeAnchors!.Count);
        Assert.NotNull(split.Op.SegmentDiffs);
        Assert.Empty(s.Conflicts);
        // The split did NOT also lower: no del/ins pair for the split base paragraph.
        Assert.DoesNotContain(s.Operations, o => o.Op.Kind == IrEditOpKind.DeleteBlock);
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Uncontested_split_composes_with_other_reviewers_edit_and_round_trips(ConflictResolution policy)
    {
        var b = Docs.Para(SplitSrc, Anchor);
        var alice = Docs.Para(SplitA, SplitB, Anchor);
        var bob = Docs.Para(SplitSrc, Anchor + " BOB tail");

        Assert.Empty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        // Accept: the split produced TWO paragraphs (Alice) and Bob's tail landed.
        var acceptText = Docs.PlainText(accepted);
        Assert.Contains("BOB tail", acceptText);
        Assert.Equal(3, acceptText.Split('\n').Length);   // SplitA / SplitB / anchor

        // Reject ≡ base: the two halves re-fuse into the single source paragraph.
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(rejected));
    }

    // ---- 2. Uncontested merge composes natively ----

    [Fact]
    public void Uncontested_merge_survives_as_native_MergeBlock_and_consumes_both_base_anchors()
    {
        var b = Docs.Para(M1, M2, Anchor);
        var alice = Docs.Para(M1 + " " + M2, Anchor);              // Alice merges paragraphs 1+2
        var bob = Docs.Para(M1, M2, Anchor + " BOB tail");         // Bob edits the anchor paragraph

        // Precondition: the pairwise aligner actually detects Alice's edit as a MergeBlock.
        Assert.Contains("\"kind\": \"MergeBlock\"", DocxDiff.GetEditScriptJson(b, alice));

        var s = Merge(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));

        var merge = Assert.Single(s.Operations, o => o.Op.Kind == IrEditOpKind.MergeBlock);
        Assert.Equal("Alice", merge.Author);
        Assert.Equal(2, merge.Op.SplitMergeAnchors!.Count);
        Assert.Empty(s.Conflicts);
        // The consumed base paragraphs are subsumed by the merge op: no separate EqualBlock/DeleteBlock
        // for either consumed anchor (the merge claims them exactly once).
        var consumed = merge.Op.SplitMergeAnchors!.ToHashSet();
        Assert.DoesNotContain(s.Operations, o =>
            !ReferenceEquals(o, merge) && o.Op.LeftAnchor is { } la && consumed.Contains(la));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Uncontested_merge_composes_with_other_reviewers_edit_and_round_trips(ConflictResolution policy)
    {
        var b = Docs.Para(M1, M2, Anchor);
        var alice = Docs.Para(M1 + " " + M2, Anchor);
        var bob = Docs.Para(M1, M2, Anchor + " BOB tail");

        Assert.Empty(Conflicts(b, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        // Accept: paragraphs 1+2 fused into one (2 paragraphs total) and Bob's tail landed.
        var acceptText = Docs.PlainText(accepted);
        Assert.Contains("BOB tail", acceptText);
        Assert.Contains("Alpha", acceptText);
        Assert.Contains("Beta", acceptText);
        Assert.Equal(2, acceptText.Split('\n').Length);   // merged para / anchor

        // Reject ≡ base: the merge's deleted pilcrows restore, re-separating the two paragraphs.
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(rejected));
    }

    // ---- 3. Insert following a merge slots AFTER the merged region ----

    [Fact]
    public void Insert_after_merged_region_lands_after_the_merge()
    {
        var b = Docs.Para(M1, M2, Anchor);
        var alice = Docs.Para(M1 + " " + M2, Anchor);                     // merges 1+2
        var bob = Docs.Para(M1, M2, "brand new bob paragraph.", Anchor);  // inserts AFTER the merged region

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var acceptText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));

        // Bob's insert must land AFTER the merged paragraph, not between its members (which would
        // corrupt the deleted-pilcrow reject reconstruction).
        Assert.True(acceptText.IndexOf("Beta", System.StringComparison.Ordinal)
                    < acceptText.IndexOf("brand new bob", System.StringComparison.Ordinal),
            $"insert should follow the merged region; accept body: {acceptText}");
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- 4. Contested splits and merges ----

    [Theory]
    [InlineData(ConflictResolution.BaseWins, false)]
    [InlineData(ConflictResolution.BaseWins, true)]
    [InlineData(ConflictResolution.FirstReviewerWins, false)]
    [InlineData(ConflictResolution.FirstReviewerWins, true)]
    [InlineData(ConflictResolution.StackAll, false)]
    [InlineData(ConflictResolution.StackAll, true)]
    public void A_split_and_an_edit_compete_as_complete_paragraph_results(ConflictResolution policy, bool editorFirst)
    {
        const string first = "The ferry leaves at nine.";
        const string second = "Tickets are sold on board.";
        const string original = first + " " + second;
        const string edited = "The ferry leaves at ten. " + second;
        const string split = first + "\n" + second;
        var baseDoc = Docs.Para(original, "End.");
        var splitter = new DocxDiffReviewer { Author = "Splitter", Document = Docs.Para(first, second, "End.") };
        var editor = new DocxDiffReviewer { Author = "Editor", Document = Docs.Para(edited, "End.") };
        var reviewers = editorFirst ? new[] { editor, splitter } : new[] { splitter, editor };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };

        Assert.Contains("\"kind\": \"SplitBlock\"", DocxDiff.GetEditScriptJson(baseDoc, splitter.Document));
        var results = editorFirst ? new[] { edited, split } : new[] { split, edited };
        AssertSplitConflict(baseDoc, reviewers, settings, original, results, "End.");
    }

    private static void AssertSplitConflict(WmlDocument baseDoc, DocxDiffReviewer[] reviewers,
        DocxDiffConsolidateSettings settings, string original, string[] results, string tail)
    {
        var policy = settings.ConflictResolution;
        var marked = DocxCompare.Consolidate(baseDoc, reviewers, settings);
        var accepted = RevisionProcessor.AcceptRevisions(marked);
        var rejected = RevisionProcessor.RejectRevisions(marked);
        foreach (var doc in new[] { marked, accepted, rejected })
            PackageValidation.AssertValid(doc, DocumentFormat.OpenXml.FileFormatVersions.Office2019);

        string expected = policy switch
        {
            ConflictResolution.BaseWins => original,
            ConflictResolution.FirstReviewerWins => results[0],
            _ => string.Join("\n", results),
        };
        Assert.Equal(original + "\n" + tail, Docs.PlainText(rejected));
        Assert.Equal(expected + "\n" + tail, Docs.PlainText(accepted));
        var conflict = Assert.Single(DocxDiff.GetConflicts(baseDoc, reviewers, settings));
        Assert.Equal(policy, conflict.AppliedPolicy);
        Assert.Equal(reviewers.Select(r => r.Author), conflict.Competitors.Select(c => c.Author));
        Assert.Equal(results, conflict.Competitors.Select(c => c.ResultText));

        var revisions = DocxDiff.GetConsolidatedRevisions(baseDoc, reviewers, settings);
        var xml = XDocument.Parse(Docs.MainPartXml(marked));
        XNamespace w = IrTestDocuments.W;
        var authors = policy == ConflictResolution.BaseWins ? System.Array.Empty<string>()
            : policy == ConflictResolution.FirstReviewerWins ? new[] { reviewers[0].Author }
            : reviewers.Select(r => r.Author).ToArray();
        Assert.Equal(authors, xml.Descendants().Attributes(w + "author").Select(a => a.Value).Distinct());
        Assert.Equal(authors, revisions.Select(r => r.Author).Distinct());
        Assert.All(revisions, r => Assert.Equal(conflict.Id, r.ConflictId));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void A_split_in_a_footnote_obeys_the_same_conflict_decision(ConflictResolution policy)
    {
        WmlDocument Note(params string[] paragraphs) => IrTestDocuments.FromBodyXmlWithFootnoteParagraph("",
            paragraphs.Select(p => $"<w:r><w:t xml:space=\"preserve\">{p}</w:t></w:r>").ToArray());
        string[] Paragraphs(WmlDocument document)
        {
            using var stream = new System.IO.MemoryStream(document.DocumentByteArray);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            var note = Assert.Single(package.MainDocumentPart!.FootnotesPart!.GetXDocument().Descendants(W.footnote),
                n => n.Attribute(W.type) == null);
            return note.Elements(W.p).Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))).ToArray();
        }

        const string edit = "aaa bbb ccc replaced. eee fff ggg hhh.";
        var original = Note(SplitSrc);
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Editor", Document = Note(edit) },
            new DocxDiffReviewer { Author = "Splitter", Document = Note(SplitA, SplitB) },
        };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };
        var marked = DocxCompare.Consolidate(original, reviewers, settings);
        var accepted = RevisionProcessor.AcceptRevisions(marked);
        var rejected = RevisionProcessor.RejectRevisions(marked);
        foreach (var doc in new[] { marked, accepted, rejected })
            PackageValidation.AssertValid(doc, DocumentFormat.OpenXml.FileFormatVersions.Office2019);
        var conflict = Assert.Single(DocxDiff.GetConflicts(original, reviewers, settings));
        Assert.Equal(new[] { edit, SplitA + "\n" + SplitB }, conflict.Competitors.Select(c => c.ResultText));
        Assert.Equal(Paragraphs(original), Paragraphs(rejected));
        Assert.Equal(policy switch
        {
            ConflictResolution.BaseWins => new[] { SplitSrc },
            ConflictResolution.FirstReviewerWins => new[] { edit },
            _ => new[] { edit, SplitA, SplitB },
        }, Paragraphs(accepted));
    }

    [Fact]
    public void Identical_splits_are_reported_once()
    {
        var original = Docs.Para(SplitSrc, Anchor);
        var result = Docs.Para(SplitA, SplitB, Anchor);
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "First", Document = result },
            new DocxDiffReviewer { Author = "Second", Document = Docs.Para(SplitA, SplitB, Anchor) },
        };
        var marked = DocxCompare.Consolidate(original, reviewers);
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers));
        Assert.Equal(Docs.PlainText(result), Docs.PlainText(RevisionProcessor.AcceptRevisions(marked)));
        Assert.Equal(Docs.PlainText(original), Docs.PlainText(RevisionProcessor.RejectRevisions(marked)));
    }

    private static string TextRun(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    private static string Paragraph(string content, string attributes = "") => $"<w:p {attributes}>{content}</w:p>";

    [Theory]
    [InlineData(ConflictResolution.BaseWins, false)]
    [InlineData(ConflictResolution.BaseWins, true)]
    [InlineData(ConflictResolution.FirstReviewerWins, false)]
    [InlineData(ConflictResolution.FirstReviewerWins, true)]
    [InlineData(ConflictResolution.StackAll, false)]
    [InlineData(ConflictResolution.StackAll, true)]
    public void A_whole_block_split_with_an_inline_control_competes_as_one_result(ConflictResolution policy, bool editorFirst)
    {
        static string Control(string text) =>
            "<w:sdt><w:sdtPr><w:tag w:val=\"A\"/></w:sdtPr><w:sdtContent>" + TextRun(text) + "</w:sdtContent></w:sdt>";
        string tail = Paragraph(TextRun(Anchor));
        var original = IrTestDocuments.FromBodyXml(Paragraph(Control(SplitA) + TextRun(SplitB)) + tail);
        var split = IrTestDocuments.FromBodyXml(Paragraph(Control(SplitA)) + Paragraph(TextRun(SplitB)) + tail);
        var edit = IrTestDocuments.FromBodyXml(Paragraph(Control(SplitA) + TextRun("eee fff altered hhh.")) + tail);
        var splitter = new DocxDiffReviewer { Author = "Splitter", Document = split };
        var editor = new DocxDiffReviewer { Author = "Editor", Document = edit };
        var reviewers = editorFirst ? new[] { editor, splitter } : new[] { splitter, editor };
        var results = reviewers.Select(r => string.Join("\n", Docs.PlainText(r.Document).Split('\n').SkipLast(1))).ToArray();
        AssertSplitConflict(original, reviewers, new DocxDiffConsolidateSettings { ConflictResolution = policy },
            SplitSrc, results, Anchor);
    }

    [Fact]
    public void A_split_consensus_ignores_Word_save_identifiers()
    {
        WmlDocument Saved(string id) => IrTestDocuments.FromBodyXml(Paragraph(TextRun(SplitA),
            $"xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" w14:paraId=\"{id}\" w14:textId=\"{id}\"")
            + Paragraph(TextRun(SplitB)) + Paragraph(TextRun(Anchor)));
        var original = Docs.Para(SplitSrc, Anchor);
        var reviewers = new[] { new DocxDiffReviewer { Author = "A", Document = Saved("12345678") },
            new DocxDiffReviewer { Author = "B", Document = Saved("23456789") } };
        Assert.Empty(DocxDiff.GetConflicts(original, reviewers));
        Assert.Equal(Docs.PlainText(reviewers[0].Document),
            Docs.PlainText(RevisionProcessor.AcceptRevisions(DocxCompare.Consolidate(original, reviewers))));
    }

    [Fact]
    public void A_winning_split_keeps_a_block_bookmark_between_its_members()
    {
        var original = Docs.Para(SplitSrc, Anchor);
        var split = IrTestDocuments.FromBodyXml(Paragraph(TextRun(SplitA)) + "<w:bookmarkStart w:id=\"1\" w:name=\"Range\"/>"
            + Paragraph(TextRun(SplitB)) + "<w:bookmarkEnd w:id=\"1\"/>" + Paragraph(TextRun(Anchor)));
        var reviewers = new[] { new DocxDiffReviewer { Author = "A", Document = split },
            new DocxDiffReviewer { Author = "B", Document = Docs.Para(SplitSrc.Replace("bbb", "BBB"), Anchor) } };
        var marked = DocxCompare.Consolidate(original, reviewers, new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.FirstReviewerWins });
        var accepted = XDocument.Parse(Docs.MainPartXml(RevisionProcessor.AcceptRevisions(marked)));
        XNamespace w = IrTestDocuments.W;
        Assert.Equal(new[] { "p", "bookmarkStart", "p", "bookmarkEnd", "p" },
            accepted.Descendants(w + "body").Single().Elements().Where(e => e.Name != w + "sectPr").Select(e => e.Name.LocalName));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_split_consensus_accounts_for_the_referenced_note(bool sameNote)
    {
        var original = Docs.Para(SplitSrc, Anchor);
        string body = Paragraph(TextRun(SplitA)) + Paragraph(TextRun(SplitB) + "<w:r><w:footnoteReference w:id=\"1\"/></w:r>") + Paragraph(TextRun(Anchor));
        var reviewers = new[] { new DocxDiffReviewer { Author = "A", Document = IrTestDocuments.FromBodyXmlWithFootnote(body, "First note.", false) },
            new DocxDiffReviewer { Author = "B", Document = IrTestDocuments.FromBodyXmlWithFootnote(body, sameNote ? "First note." : "Second note.", false) } };
        Assert.Equal(!sameNote, DocxDiff.GetConflicts(original, reviewers).Any(c => c.Competitors.Any(v => v.ResultText.Contains("eee"))));
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.StackAll };
        var marked = DocxCompare.Consolidate(original, reviewers, settings);
        string[] ReferencedNotes(WmlDocument doc)
        {
            using var stream = new System.IO.MemoryStream(doc.DocumentByteArray);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            var main = package.MainDocumentPart!;
            var definitions = main.FootnotesPart?.GetXDocument().Root?.Elements(W.footnote)
                .ToDictionary(n => (string)n.Attribute(W.id)!, n => string.Concat(n.Descendants(W.t).Select(t => t.Value)));
            return main.GetXDocument().Descendants(W.footnoteReference).Select(r => definitions![(string)r.Attribute(W.id)!]).ToArray();
        }
        Assert.Empty(ReferencedNotes(RevisionProcessor.RejectRevisions(marked)));
        Assert.Equal(sameNote ? new[] { "First note." } : new[] { "First note.", "Second note." },
            ReferencedNotes(RevisionProcessor.AcceptRevisions(marked)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_split_consensus_accounts_for_each_bookmarks_paired_end(bool sameEnds)
    {
        const string starts = "<w:bookmarkStart w:id=\"1\" w:name=\"First\"/><w:bookmarkStart w:id=\"2\" w:name=\"Second\"/>";
        static string Ends(string id) => $"<w:bookmarkEnd w:id=\"{id}\"/>";
        string Part(bool swap) => starts + TextRun("eee ") + Ends(swap ? "2" : "1") + TextRun("fff ") + Ends(swap ? "1" : "2") + TextRun("ggg hhh.");
        var original = IrTestDocuments.FromBodyXml(Paragraph(TextRun(SplitA) + Part(false)) + Paragraph(TextRun(Anchor)));
        WmlDocument Split(bool swap) => IrTestDocuments.FromBodyXml(Paragraph(TextRun(SplitA)) + Paragraph(Part(swap)) + Paragraph(TextRun(Anchor)));
        var reviewers = new[] { new DocxDiffReviewer { Author = "A", Document = Split(false) }, new DocxDiffReviewer { Author = "B", Document = Split(!sameEnds) } };
        Assert.Equal(!sameEnds, DocxDiff.GetConflicts(original, reviewers).Any());
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Contested_merge_lowers_and_records_conflict(ConflictResolution policy)
    {
        var b = Docs.Para(M1, M2, Anchor);
        var alice = Docs.Para(M1 + " " + M2, Anchor);                     // merges 1+2
        var bob = Docs.Para(M1, "Beta beta REPLACED here.", Anchor);      // edits consumed paragraph 2

        var s = Merge(b, policy, ("Alice", alice), ("Bob", bob));
        Assert.DoesNotContain(s.Operations, o => o.Op.Kind == IrEditOpKind.MergeBlock);
        Assert.NotEmpty(s.Conflicts);

        var merged = Consolidate(b, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- 5. Single-reviewer consolidate now renders split/merge natively (parity with Compare) ----

    [Fact]
    public void Single_reviewer_split_consolidate_matches_two_way_markup_shape()
    {
        var b = Docs.Para(SplitSrc, Anchor);
        var alice = Docs.Para(SplitA, SplitB, Anchor);

        var merged = Consolidate(b, ConflictResolution.BaseWins, ("Alice", alice));

        // Native split markup: an inserted paragraph mark (w:rPr/w:ins inside w:pPr) rather than a
        // whole-paragraph delete + two whole-paragraph inserts. The two-way Compare is the shape oracle.
        Assert.Equal(Docs.PlainText(alice), Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

        // No whole-block delete of the source paragraph: the split source's text must NOT appear inside
        // a w:delText (native split keeps the shared text bare, marking only the pilcrow).
        var xml = Docs.MainPartXml(merged);
        Assert.DoesNotContain("aaa bbb ccc ddd", GetDelText(xml));
    }

    /// <summary>Concatenated w:delText content of the main part.</summary>
    private static string GetDelText(string mainXml)
    {
        var doc = XDocument.Parse(mainXml);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        return string.Concat(doc.Descendants(w + "delText").Select(e => e.Value));
    }
}
