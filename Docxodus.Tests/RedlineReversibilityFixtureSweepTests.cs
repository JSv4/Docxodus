// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Fixture-scale evidence for the redline reversibility proof (#464), complementing the
/// scenario-focused cases in <see cref="RedlineReversibilityProofTests"/>.
///
/// Three redline sources, three claims:
///
/// 1. Word-authored redlines — every complete triple in <c>TestFiles/RP</c>
///    (<c>{stem}.docx</c> + <c>{stem}-Accepted.docx</c> + <c>{stem}-Rejected.docx</c>, the
///    long-standing RevisionProcessor oracle corpus). The proof's selective resolver must
///    accept every revision to the accepted oracle and reject every revision to the rejected
///    oracle at the story-text level, and each triple pins whether the far stronger
///    normalized whole-package equivalence also holds per direction. A guard test keeps the
///    pinned list in lock-step with the fixture directory so a new triple cannot be added
///    without being swept.
///
/// 2. Engine-generated redlines — the WCB comparison corpus (real Word documents under
///    <c>TestFiles/WC</c>, <c>CA</c>, <c>RC</c>). For each before/after pair the redline is
///    <c>DocxDiff.Compare(left, right)</c> and the proof must complete both paths, classify
///    every revision as generated, and recover <c>right</c> on accept / <c>left</c> on
///    reject at the story-text level. Pairs where that does not currently hold are pinned as
///    explicit known gaps with their exact failure signature, so the gap is visible here and
///    this file must be updated when it is fixed.
///
/// 3. Consolidated multi-author redlines — <c>DocxDiff.Consolidate</c> over two reviewers with
///    DISTINCT author names (RRS008/RRS009). Ownership classification compares revision
///    authors, so a redline carrying two reviewer identities exercises a path the
///    single-author corpora above never reach. The proven contract is Consolidate's own:
///    reject recovers the shared base, accept recovers the policy-resolved composite.
///
/// "Story text" means the concatenated <c>w:t</c> runs of the body plus every header,
/// footer, footnotes and endnotes part — presence of a story part counts, so an output that
/// grows an empty footnotes part the expected document never had does not pass.
/// </summary>
public class RedlineReversibilityFixtureSweepTests
{
    private static readonly string TestFilesDir = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles"));

    // ------------------------------------------------------------------
    // 1. Word-authored redlines: TestFiles/RP triples.
    // ------------------------------------------------------------------

    // Columns: fixture stem, accept path fully equivalent (normalized whole package),
    // reject path fully equivalent. Story-text equality with the oracle documents is
    // asserted unconditionally for every row; the booleans pin only the stronger
    // package-level claim per direction.
    [Theory]
    [InlineData("RP002-Deleted-Text", true, true)]
    [InlineData("RP003-Inserted-Text", true, true)]
    [InlineData("RP004-Deleted-Text-in-CC", true, true)]
    [InlineData("RP005-Deleted-Paragraph-Mark", false, true)]
    [InlineData("RP006-Inserted-Paragraph-Mark", true, false)]
    [InlineData("RP007-Multiple-Deleted-Para-Mark", false, true)]
    [InlineData("RP008-Multiple-Inserted-Para-Mark", true, false)]
    [InlineData("RP009-Deleted-Table-Row", true, true)]
    [InlineData("RP010-Inserted-Table-Row", true, true)]
    [InlineData("RP011-Multiple-Deleted-Rows", true, true)]
    [InlineData("RP012-Multiple-Inserted-Rows", true, true)]
    [InlineData("RP013-Deleted-Math-Control-Char", true, true)]
    [InlineData("RP014-Inserted-Math-Control-Char", true, true)]
    [InlineData("RP015-MoveFrom-MoveTo", false, false)]
    [InlineData("RP016-Deleted-CC", false, false)]
    [InlineData("RP017-Inserted-CC", true, true)]
    [InlineData("RP019-Deleted-Field-Code", false, false)]
    [InlineData("RP020-Inserted-Field-Code", true, false)]
    [InlineData("RP021-Inserted-Numbering-Properties", true, false)]
    [InlineData("RP022-NumberingChange", true, false)]
    [InlineData("RP023-NumberingChange", false, false)]
    [InlineData("RP024-ParagraphMark-rPr-Change", false, false)]
    [InlineData("RP025-Paragraph-Props-Change", false, false)]
    [InlineData("RP026-NumberingChange", false, false)]
    [InlineData("RP027-Change-Section", true, false)]
    [InlineData("RP028-Table-Grid-Change", true, false)]
    [InlineData("RP029-Table-Row-Props-Change", true, false)]
    [InlineData("RP030-Table-Row-Props-Change", true, false)]
    [InlineData("RP031-Table-Prop-Change", true, false)]
    [InlineData("RP032-Table-Prop-Change", true, false)]
    [InlineData("RP033-Table-Prop-Ex-Change", true, false)]
    [InlineData("RP034-Deleted-Cells", false, false)]
    [InlineData("RP035-Inserted-Cells", false, false)]
    [InlineData("RP036-Vert-Merged-Cells", true, false)]
    [InlineData("RP037-Changed-Style-Para-Props", true, false)]
    [InlineData("RP038-Inserted-Paras-at-End", false, false)]
    [InlineData("RP039-Inserted-Paras-at-End", false, false)]
    [InlineData("RP040-Deleted-Paras-at-End", false, false)]
    [InlineData("RP041-Cell-With-Empty-Paras-at-End", false, false)]
    [InlineData("RP042-Deleted-Para-Mark-at-End", false, false)]
    [InlineData("RP043-MERGEFORMAT-Field-Code", false, false)]
    [InlineData("RP044-MERGEFORMAT-Field-Code", false, false)]
    [InlineData("RP045-One-and-Half-Deleted-Lines-at-End", false, false)]
    [InlineData("RP046-Consecutive-Deleted-Ranges", false, false)]
    [InlineData("RP048-Deleted-Inserted-Para-Mark", false, false)]
    [InlineData("RP049-Deleted-Para-Before-Table", false, false)]
    [InlineData("RP050-Deleted-Footnote", false, false)]
    [InlineData("RP051-Arabic", false, false)]
    [InlineData("RP052-Deleted-Para-Mark", false, false)]
    public void RRS001_WordAuthoredRedline_AcceptAndRejectRecoverTheOracleDocuments(
        string stem,
        bool acceptEquivalent,
        bool rejectEquivalent)
    {
        var (redline, acceptedOracle, rejectedOracle) = RpTriple(stem);

        var run = RedlineReversibilityVerifier.Prove(rejectedOracle, acceptedOracle, redline);
        var proof = run.Proof;

        Assert.NotNull(proof.AcceptToFinal);
        Assert.NotNull(proof.RejectToBaseline);
        Assert.True(proof.AcceptToFinal!.Completed, proof.ToJson());
        Assert.True(proof.RejectToBaseline!.Completed, proof.ToJson());

        // A Word-authored redline over its fully-rejected baseline owns every revision.
        Assert.NotEmpty(proof.RevisionClassifications);
        Assert.All(proof.RevisionClassifications, classification =>
            Assert.Equal(RedlineRevisionDisposition.Generated, classification.Disposition));
        AssertResolutionClosure(proof.AcceptToFinal);
        AssertResolutionClosure(proof.RejectToBaseline);

        // The reversibility contract a reader actually cares about: the accepted output
        // reads as Word's accepted oracle, the rejected output as Word's rejected oracle.
        Assert.Equal(StoryTexts(acceptedOracle), StoryTexts(run.AcceptedPackageBytes!));
        Assert.Equal(StoryTexts(rejectedOracle), StoryTexts(run.RejectedPackageBytes!));

        // The stronger pinned claim, per direction.
        Assert.True(proof.AcceptToFinal.Equivalent == acceptEquivalent, proof.ToJson());
        Assert.True(proof.RejectToBaseline.Equivalent == rejectEquivalent, proof.ToJson());
        Assert.True(proof.Success == (acceptEquivalent && rejectEquivalent), proof.ToJson());

        // A non-equivalent path must say so with an error finding, never silently.
        AssertNonEquivalenceIsEvidenced(proof.AcceptToFinal);
        AssertNonEquivalenceIsEvidenced(proof.RejectToBaseline);
    }

    // Word-authored triples whose revision families the resolver does not support: the
    // proof must refuse to run either path rather than guess, and must name the reason.
    [Theory]
    [InlineData("RP018-MoveFrom-MoveTo-CC", "unsupported_custom_xml_move_range")]
    [InlineData("RP047-Inserted-and-Deleted-Paragraph-Mark", "unsupported_revision_family")]
    public void RRS002_WordAuthoredRedline_UnsupportedFamilyFailsClosedWithDiagnostic(
        string stem,
        string expectedDiagnosticCode)
    {
        var (redline, acceptedOracle, rejectedOracle) = RpTriple(stem);

        var run = RedlineReversibilityVerifier.Prove(rejectedOracle, acceptedOracle, redline);

        Assert.False(run.Proof.Success);
        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.Proof.RejectToBaseline);
        Assert.Null(run.AcceptedPackageBytes);
        Assert.Null(run.RejectedPackageBytes);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "generated_revision_not_resolvable");
        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.Diagnostic?.Code == expectedDiagnosticCode);
    }

    // Keeps RRS001 + RRS002 honest: every complete triple in TestFiles/RP must be pinned in
    // exactly one of them, so adding a fixture without sweeping it fails here.
    [Fact]
    public void RRS003_EveryCompleteFixtureTripleIsPinned()
    {
        var rpDir = Path.Combine(TestFilesDir, "RP");
        var files = new HashSet<string>(
            Directory.GetFiles(rpDir, "*.docx").Select(Path.GetFileName)!,
            StringComparer.Ordinal);
        var completeTriples = files
            .Where(name => !name.EndsWith("-Accepted.docx", StringComparison.Ordinal)
                && !name.EndsWith("-Rejected.docx", StringComparison.Ordinal))
            .Select(name => name[..^".docx".Length])
            .Where(stem => files.Contains(stem + "-Accepted.docx")
                && files.Contains(stem + "-Rejected.docx"))
            .OrderBy(stem => stem, StringComparer.Ordinal)
            .ToArray();

        var pinned = PinnedStems(nameof(RRS001_WordAuthoredRedline_AcceptAndRejectRecoverTheOracleDocuments))
            .Concat(PinnedStems(nameof(RRS002_WordAuthoredRedline_UnsupportedFamilyFailsClosedWithDiagnostic)))
            .OrderBy(stem => stem, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(completeTriples, pinned);
    }

    // ------------------------------------------------------------------
    // 2. Engine-generated redlines: the WCB comparison corpus.
    // ------------------------------------------------------------------

    // Every before/after pair from the WCB corpus except the known gaps pinned in
    // RRS005-RRS007 below. DocxDiff regenerates the package, so normalized whole-package
    // equivalence never holds here (asserted); the reversibility claim proven per pair is
    // content-level: accept recovers right's stories, reject recovers left's.
    [Theory]
    [InlineData("CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx")]
    [InlineData("WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx")]
    [InlineData("WC/WC004-Large.docx", "WC/WC004-Large-Mod.docx")]
    [InlineData("WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx")]
    [InlineData("WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx")]
    [InlineData("WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx")]
    [InlineData("WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx")]
    [InlineData("WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx")]
    [InlineData("WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx")]
    [InlineData("WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx")]
    [InlineData("WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx")]
    [InlineData("WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx")]
    [InlineData("WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx")]
    [InlineData("WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx")]
    [InlineData("WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx")]
    [InlineData("WC/WC011-Before.docx", "WC/WC011-After.docx")]
    [InlineData("WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx")]
    [InlineData("WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx")]
    [InlineData("WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx")]
    [InlineData("WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx")]
    [InlineData("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx")]
    [InlineData("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx")]
    [InlineData("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx")]
    [InlineData("WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx")]
    [InlineData("WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx")]
    [InlineData("WC/WC017-Image.docx", "WC/WC017-Image-After.docx")]
    [InlineData("WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx")]
    [InlineData("WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx")]
    [InlineData("WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx")]
    [InlineData("WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx")]
    [InlineData("WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx")]
    [InlineData("WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx")]
    [InlineData("WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx")]
    [InlineData("WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx")]
    [InlineData("WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx")]
    [InlineData("WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx")]
    [InlineData("WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx")]
    [InlineData("WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx")]
    [InlineData("WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx")]
    [InlineData("WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx")]
    [InlineData("WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx")]
    [InlineData("WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx")]
    [InlineData("WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx")]
    [InlineData("WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx")]
    [InlineData("WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx")]
    [InlineData("WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx")]
    [InlineData("WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx")]
    [InlineData("WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx")]
    [InlineData("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx")]
    [InlineData("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx")]
    [InlineData("WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx")]
    [InlineData("WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx")]
    [InlineData("WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx")]
    [InlineData("WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx")]
    [InlineData("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx")]
    [InlineData("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx")]
    [InlineData("WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx")]
    [InlineData("WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx")]
    [InlineData("WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx")]
    [InlineData("WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx")]
    [InlineData("WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx")]
    [InlineData("RC/RC001-Before.docx", "RC/RC001-After1.docx")]
    [InlineData("RC/RC002-Image.docx", "RC/RC002-Image-After1.docx")]
    public void RRS004_EngineGeneratedRedline_RoundTripsContentAcrossComparisonCorpus(
        string leftName,
        string rightName)
    {
        var left = Fixture(leftName);
        var right = Fixture(rightName);
        var redline = EngineRedline(left, right);

        var run = RedlineReversibilityVerifier.Prove(left, right, redline);
        var proof = run.Proof;

        Assert.NotNull(proof.AcceptToFinal);
        Assert.NotNull(proof.RejectToBaseline);
        Assert.True(proof.AcceptToFinal!.Completed, proof.ToJson());
        Assert.True(proof.RejectToBaseline!.Completed, proof.ToJson());
        Assert.NotEmpty(proof.RevisionClassifications);
        Assert.All(proof.RevisionClassifications, classification =>
            Assert.Equal(RedlineRevisionDisposition.Generated, classification.Disposition));
        AssertResolutionClosure(proof.AcceptToFinal);
        AssertResolutionClosure(proof.RejectToBaseline);

        // The content-level reversibility contract.
        Assert.Equal(StoryTexts(right), StoryTexts(run.AcceptedPackageBytes!));
        Assert.Equal(StoryTexts(left), StoryTexts(run.RejectedPackageBytes!));

        // DocxDiff rebuilds the output package, so the strict whole-package proof does not
        // hold for engine output today — and the proof must say so rather than succeed.
        Assert.False(proof.Success, proof.ToJson());
        AssertNonEquivalenceIsEvidenced(proof.AcceptToFinal);
        AssertNonEquivalenceIsEvidenced(proof.RejectToBaseline);
    }

    // ------------------------------------------------------------------
    // Known gaps in the engine corpus, pinned by exact failure signature. Each of these is
    // excluded from RRS004; when the underlying defect is fixed the pin fails and both the
    // pin and the RRS004 exclusion must be updated.
    // ------------------------------------------------------------------

    // WC004-Large regression pin for the selective-reject moveFrom fix: the comparison
    // engine serializes move-source text as w:delText inside w:moveFrom, and rejecting
    // those moves used to strand ~1.3k characters as orphaned w:delText once the wrapper
    // was unwrapped without the delText → t restore. Beyond RRS004's story-text oracle
    // (which WC004 now passes), assert the specific corruption shape can never return:
    // the rejected package contains no w:delText outside a delete-grade wrapper.
    [Fact]
    public void RRS005_RejectedMoveRestoresDeletedTextInsteadOfStrandingIt()
    {
        var left = Fixture("WC/WC004-Large.docx");
        var right = Fixture("WC/WC004-Large-Mod.docx");
        var redline = EngineRedline(left, right);

        var run = RedlineReversibilityVerifier.Prove(left, right, redline);

        Assert.True(run.Proof.RejectToBaseline?.Completed, run.Proof.ToJson());
        using var stream = new MemoryStream(run.RejectedPackageBytes!, writable: false);
        using var rejected = WordprocessingDocument.Open(stream, false);
        var body = rejected.MainDocumentPart!.Document.Body!;
        Assert.Empty(body.Descendants<DeletedText>());
        Assert.Empty(body.Descendants<DeletedRun>());
        Assert.Empty(body.Descendants<InsertedRun>());
    }

    // WC035: when a comparison adds the document's first footnote/endnote, the redline
    // package gains a notes part. Resolving toward the endpoint that has no notes leaves
    // that (now content-empty) part behind instead of removing it, so the output package
    // has a story the expected document never had. Body and note text themselves
    // round-trip; only the empty-part residue diverges.
    [Theory]
    [InlineData("WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx", "footnotes")]
    [InlineData("WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx", "footnotes")]
    [InlineData("WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx", "endnotes")]
    [InlineData("WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx", "endnotes")]
    public void RRS006_KnownGap_ResolvingAwayTheOnlyNoteLeavesAnEmptyNotesPart(
        string leftName,
        string rightName,
        string storyKind)
    {
        var left = Fixture(leftName);
        var right = Fixture(rightName);
        var redline = EngineRedline(left, right);

        var run = RedlineReversibilityVerifier.Prove(left, right, redline);

        Assert.True(run.Proof.AcceptToFinal?.Completed, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Completed, run.Proof.ToJson());

        // One endpoint has the note, the other has no notes part at all.
        var (noteless, noteBearing, notelessOutput, noteBearingOutput) =
            NotesText(left, storyKind) is null
                ? (left, right, run.RejectedPackageBytes!, run.AcceptedPackageBytes!)
                : (right, left, run.AcceptedPackageBytes!, run.RejectedPackageBytes!);
        Assert.Null(NotesText(noteless, storyKind));
        Assert.False(string.IsNullOrEmpty(NotesText(noteBearing, storyKind)));

        // Toward the note-bearing endpoint the round-trip is faithful.
        Assert.Equal(StoryTexts(noteBearing), StoryTexts(noteBearingOutput));

        // Toward the noteless endpoint the body text round-trips but an empty notes part
        // is left behind — the residue this pin exists for.
        Assert.Equal(BodyText(noteless), BodyText(notelessOutput));
        Assert.Equal(string.Empty, NotesText(notelessOutput, storyKind));
        Assert.False(run.Proof.Success);
    }

    // WC012-Math: the After document carries its own tracked insertions, and
    // DocxDiff.Compare does not reproduce that pre-existing review state in the redline.
    // Accepting such a redline can never recover the After document exactly, and the proof
    // must refuse to certify rather than resolve around the loss.
    [Fact]
    public void RRS007_KnownGap_RedlineMissingIntendedFinalReviewStateFailsClosed()
    {
        var left = Fixture("WC/WC012-Math-Before.docx");
        var right = Fixture("WC/WC012-Math-After.docx");
        Assert.True(ContainsNativeRevisions(right)); // the After doc's own review state
        var redline = EngineRedline(left, right);

        var run = RedlineReversibilityVerifier.Prove(left, right, redline);

        Assert.False(run.Proof.Success);
        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.Proof.RejectToBaseline);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "intended_final_revision_missing_from_redline");
    }

    // ------------------------------------------------------------------
    // 3. Consolidated multi-author redlines: DocxDiff.Consolidate.
    // ------------------------------------------------------------------

    // Two reviewers with disjoint edits of the same base, each stamped with a distinct
    // author name. The proof classifies every revision from BOTH reviewers as generated —
    // multi-author attribution must never read as conflicted ownership — and both paths
    // complete: accept recovers the composite (the accept-all of the consolidated redline,
    // Consolidate's own round-trip contract) and reject recovers the shared base, at the
    // story-text level AND at the modeled-semantic level. As with every engine-generated
    // redline (RRS004), Consolidate rebuilds the output package, so the strict
    // whole-package proof does not hold today and the proof must say so with evidence.
    [Fact]
    public void RRS008_ConsolidatedTwoReviewerRedline_RoundTripsWithPerAuthorClassification()
    {
        var baseBytes = Fixture("WC/WC002-Unmodified.docx");
        var consolidated = ConsolidatedRedline(
            baseBytes,
            Fixture("WC/WC002-DiffInMiddle.docx"),
            Fixture("WC/WC002-InsertAtEnd.docx"));
        var intendedComposite = RevisionProcessor.AcceptRevisions(consolidated).DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(
            baseBytes, intendedComposite, consolidated.DocumentByteArray);
        var proof = run.Proof;

        // Both reviewer identities are present, and every revision classifies as generated.
        Assert.NotEmpty(proof.RevisionClassifications);
        Assert.All(proof.RevisionClassifications, classification =>
            Assert.Equal(RedlineRevisionDisposition.Generated, classification.Disposition));
        Assert.Contains(proof.RevisionClassifications, classification =>
            classification.Redline?.Author == "Reviewer A");
        Assert.Contains(proof.RevisionClassifications, classification =>
            classification.Redline?.Author == "Reviewer B");

        Assert.NotNull(proof.AcceptToFinal);
        Assert.NotNull(proof.RejectToBaseline);
        Assert.True(proof.AcceptToFinal!.Completed, proof.ToJson());
        Assert.True(proof.RejectToBaseline!.Completed, proof.ToJson());
        AssertResolutionClosure(proof.AcceptToFinal);
        AssertResolutionClosure(proof.RejectToBaseline);

        // Accept recovers the composite, reject recovers the shared base.
        Assert.Equal(StoryTexts(intendedComposite), StoryTexts(run.AcceptedPackageBytes!));
        Assert.Equal(StoryTexts(baseBytes), StoryTexts(run.RejectedPackageBytes!));
        Assert.True(proof.AcceptToFinal.ModeledSemantic.Equivalent, proof.ToJson());
        Assert.True(proof.RejectToBaseline.ModeledSemantic.Equivalent, proof.ToJson());

        // The rebuilt package prevents the strict whole-package claim — evidenced, never silent.
        Assert.False(proof.Success, proof.ToJson());
        AssertNonEquivalenceIsEvidenced(proof.AcceptToFinal);
        AssertNonEquivalenceIsEvidenced(proof.RejectToBaseline);
    }

    // Two reviewers editing the SAME span differently — a genuine merge conflict, resolved
    // by the default BaseWins policy. The policy decides what markup survives into the
    // consolidated document (observed: the conflicted span resolves to base, so a
    // reviewer's competing edit may not appear at all); whatever survives must still
    // classify as generated for its own reviewer — a policy-resolved merge conflict is not
    // an ownership conflict — and the proof round-trips against the policy-resolved
    // composite exactly as in RRS008.
    [Fact]
    public void RRS009_ConsolidatedConflictingReviewers_ProveAgainstThePolicyResolvedComposite()
    {
        var baseBytes = Fixture("WC/WC002-Unmodified.docx");
        var reviewerA = Fixture("WC/WC002-DiffInMiddle.docx");
        var reviewerB = Fixture("WC/WC002-DeleteInMiddle.docx");

        // The pair genuinely conflicts: both reviewers appear as competitors on a base span.
        var conflicts = DocxDiff.GetConflicts(
            new WmlDocument("base.docx", baseBytes),
            Reviewers(reviewerA, reviewerB));
        Assert.NotEmpty(conflicts);
        Assert.Contains(conflicts.SelectMany(conflict => conflict.Competitors),
            competitor => competitor.Author == "Reviewer A");
        Assert.Contains(conflicts.SelectMany(conflict => conflict.Competitors),
            competitor => competitor.Author == "Reviewer B");

        var consolidated = ConsolidatedRedline(baseBytes, reviewerA, reviewerB);
        var intendedComposite = RevisionProcessor.AcceptRevisions(consolidated).DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(
            baseBytes, intendedComposite, consolidated.DocumentByteArray);
        var proof = run.Proof;

        Assert.NotEmpty(proof.RevisionClassifications);
        Assert.All(proof.RevisionClassifications, classification =>
        {
            Assert.Equal(RedlineRevisionDisposition.Generated, classification.Disposition);
            Assert.Contains(classification.Redline?.Author,
                new[] { "Reviewer A", "Reviewer B" });
        });

        Assert.NotNull(proof.AcceptToFinal);
        Assert.NotNull(proof.RejectToBaseline);
        Assert.True(proof.AcceptToFinal!.Completed, proof.ToJson());
        Assert.True(proof.RejectToBaseline!.Completed, proof.ToJson());
        AssertResolutionClosure(proof.AcceptToFinal);
        AssertResolutionClosure(proof.RejectToBaseline);

        Assert.Equal(StoryTexts(intendedComposite), StoryTexts(run.AcceptedPackageBytes!));
        Assert.Equal(StoryTexts(baseBytes), StoryTexts(run.RejectedPackageBytes!));
        Assert.True(proof.AcceptToFinal.ModeledSemantic.Equivalent, proof.ToJson());
        Assert.True(proof.RejectToBaseline.ModeledSemantic.Equivalent, proof.ToJson());

        Assert.False(proof.Success, proof.ToJson());
        AssertNonEquivalenceIsEvidenced(proof.AcceptToFinal);
        AssertNonEquivalenceIsEvidenced(proof.RejectToBaseline);
    }

    // ------------------------------------------------------------------
    // Helpers.
    // ------------------------------------------------------------------

    private static byte[] Fixture(string relativeName) =>
        File.ReadAllBytes(Path.Combine(TestFilesDir, relativeName));

    private static (byte[] Redline, byte[] AcceptedOracle, byte[] RejectedOracle) RpTriple(
        string stem)
    {
        var rpDir = Path.Combine(TestFilesDir, "RP");
        return (
            File.ReadAllBytes(Path.Combine(rpDir, stem + ".docx")),
            File.ReadAllBytes(Path.Combine(rpDir, stem + "-Accepted.docx")),
            File.ReadAllBytes(Path.Combine(rpDir, stem + "-Rejected.docx")));
    }

    private static byte[] EngineRedline(byte[] left, byte[] right) => DocxDiff.Compare(
        new WmlDocument("left.docx", left),
        new WmlDocument("right.docx", right),
        new DocxDiffSettings { AuthorForRevisions = "Docxodus Engine" }).DocumentByteArray;

    private static WmlDocument ConsolidatedRedline(
        byte[] baseBytes, byte[] reviewerA, byte[] reviewerB) => DocxDiff.Consolidate(
        new WmlDocument("base.docx", baseBytes),
        Reviewers(reviewerA, reviewerB));

    private static DocxDiffReviewer[] Reviewers(byte[] reviewerA, byte[] reviewerB) => new[]
    {
        new DocxDiffReviewer
        {
            Document = new WmlDocument("reviewer-a.docx", reviewerA),
            Author = "Reviewer A",
        },
        new DocxDiffReviewer
        {
            Document = new WmlDocument("reviewer-b.docx", reviewerB),
            Author = "Reviewer B",
        },
    };

    private static IEnumerable<string> PinnedStems(string theoryMethodName) =>
        typeof(RedlineReversibilityFixtureSweepTests)
            .GetMethod(theoryMethodName)!
            .GetCustomAttributes(typeof(InlineDataAttribute), inherit: false)
            .Cast<InlineDataAttribute>()
            .SelectMany(attribute => attribute.GetData(null!))
            .Select(row => (string)row[0]!);

    private static void AssertResolutionClosure(RedlineProofPathResult path)
    {
        var accountedFor = path.ResolvedRevisionIds
            .Concat(path.ImplicitlyResolvedRevisionIds)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(
            path.RequestedRevisionIds.OrderBy(id => id, StringComparer.Ordinal),
            accountedFor);
        Assert.Equal(accountedFor.Length, accountedFor.Distinct(StringComparer.Ordinal).Count());
    }

    private static void AssertNonEquivalenceIsEvidenced(RedlineProofPathResult path)
    {
        if (path.Equivalent)
            return;
        Assert.Contains(path.Findings, finding =>
            finding.Severity == VerificationFindingSeverity.Error);
        if (!path.NormalizedWholePackageEquivalent)
            Assert.NotNull(path.FirstDivergence);
    }

    /// <summary>
    /// Canonical story-text projection of a package: the concatenated <c>w:t</c> text of the
    /// body and of every header, footer, footnotes and endnotes part, keyed by story kind.
    /// Part presence counts — an empty part produces an empty entry, absence produces none —
    /// so growing or losing a story cannot pass unnoticed.
    /// </summary>
    private static string StoryTexts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var builder = new StringBuilder();
        builder.Append("body: ").Append(TextOf(main.Document)).Append('\n');
        AppendStory(builder, "headers", main.HeaderParts.Select(part => TextOf(part.Header)));
        AppendStory(builder, "footers", main.FooterParts.Select(part => TextOf(part.Footer)));
        AppendStory(builder, "footnotes", main.FootnotesPart is { } footnotes
            ? new[] { TextOf(footnotes.Footnotes) }
            : Array.Empty<string>());
        AppendStory(builder, "endnotes", main.EndnotesPart is { } endnotes
            ? new[] { TextOf(endnotes.Endnotes) }
            : Array.Empty<string>());
        return builder.ToString();
    }

    private static void AppendStory(
        StringBuilder builder, string kind, IEnumerable<string> texts)
    {
        foreach (var text in texts.OrderBy(text => text, StringComparer.Ordinal))
            builder.Append(kind).Append(": ").Append(text).Append('\n');
    }

    private static string TextOf(DocumentFormat.OpenXml.OpenXmlElement? root) =>
        root is null ? "" : string.Concat(root.Descendants<Text>().Select(text => text.Text));

    private static string BodyText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        return TextOf(document.MainDocumentPart!.Document.Body);
    }

    private static string? NotesText(byte[] bytes, string storyKind)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        return storyKind == "footnotes"
            ? main.FootnotesPart is { } footnotes ? TextOf(footnotes.Footnotes) : null
            : main.EndnotesPart is { } endnotes ? TextOf(endnotes.Endnotes) : null;
    }

    private static bool ContainsNativeRevisions(byte[] bytes)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document
            .Descendants()
            .Any(element => element.NamespaceUri == w
                && element.LocalName is "ins" or "del");
    }
}
