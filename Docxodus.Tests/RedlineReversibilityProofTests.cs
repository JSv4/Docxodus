// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using DocxodusDiffParityFixtures;
using Xunit;

namespace Docxodus.Tests;

public class RedlineReversibilityProofTests
{
    [Fact]
    public void RP001_GeneratedTextRevision_IsClassifiedAndResolvedBothWays()
    {
        var baseline = Document("The original clause.");
        var intendedFinal = Document("The revised clause.");
        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("final.docx", intendedFinal),
            new DocxDiffSettings { AuthorForRevisions = "Comparison Engine" }).DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.RevisionClassifications.Count > 0, run.Proof.ToJson());
        Assert.All(run.Proof.RevisionClassifications,
            item => Assert.Equal(RedlineRevisionDisposition.Generated, item.Disposition));
        Assert.NotNull(run.AcceptedPackageBytes);
        Assert.NotNull(run.RejectedPackageBytes);
        Assert.True(run.Proof.AcceptToFinal?.Completed);
        Assert.True(run.Proof.RejectToBaseline?.Completed);
        Assert.False(run.Proof.AcceptToFinal?.NormalizedWholePackageEquivalent);
        Assert.NotEmpty(run.Proof.AcceptToFinal?.Divergences ?? []);
        Assert.True(run.Proof.AcceptToFinal?.ModeledSemantic.Available);
        Assert.Equal(SemanticChangeSet.CurrentSchema,
            run.Proof.AcceptToFinal?.ModeledSemantic.Schema);
        Assert.False(run.Proof.AcceptToFinal?.ModeledSemantic.Equivalent);
        Assert.True(run.Proof.AcceptToFinal?.ModeledSemantic.ChangeCount > 0);
        var semanticFinding = Assert.Single(run.Proof.AcceptToFinal?.Findings ?? [],
            finding => finding.Code == "modeled_semantic_mismatch");
        Assert.NotNull(semanticFinding.Location?.EntryUri);
        Assert.NotNull(semanticFinding.Location?.PropertyPath);
        Assert.False(run.Proof.Success); // Whole-package differences still prevent a proof.
    }

    [Fact]
    public void RP002_InvalidPackage_FailsBeforeRevisionInventory()
    {
        var valid = Document("valid");

        var run = RedlineReversibilityVerifier.Prove(valid, valid, new byte[] { 1, 2, 3 });

        Assert.False(run.Proof.Success);
        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.AcceptedPackageBytes);
        Assert.Contains(run.Proof.Findings, item =>
            item.Code == "package_malformed_package"
            && item.Severity == VerificationFindingSeverity.Error);
    }

    [Fact]
    public void RP003_ProofJson_IsDeterministicAndReceiptEmbeddable()
    {
        var baseline = Document("A");
        var intendedFinal = Document("B");
        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("final.docx", intendedFinal)).DocumentByteArray;

        var first = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline).Proof;
        var second = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline).Proof;

        Assert.Equal(first.ToCanonicalJson(), second.ToCanonicalJson());
        using var parsed = JsonDocument.Parse(first.ToCanonicalJson());
        Assert.Equal(RedlineReversibilityProof.SchemaId,
            parsed.RootElement.GetProperty("schema").GetString());
        Assert.DoesNotContain("\"acceptedPackageBytes\"", first.ToCanonicalJson(), StringComparison.Ordinal);
    }

    [Fact]
    public void RP004_WordAuthoredComparison_ReportsNonReversiblePackageDifferences()
    {
        var root = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "../../../../TestFiles/WC"));
        var baseline = File.ReadAllBytes(Path.Combine(root, "WC001-Digits.docx"));
        var intendedFinal = File.ReadAllBytes(Path.Combine(root, "WC001-Digits-Mod.docx"));
        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("final.docx", intendedFinal)).DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.AcceptToFinal?.Completed, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Completed, run.Proof.ToJson());
        Assert.False(run.Proof.AcceptToFinal?.NormalizedWholePackageEquivalent);
        Assert.False(run.Proof.RejectToBaseline?.NormalizedWholePackageEquivalent);
        Assert.NotNull(run.Proof.AcceptToFinal?.FirstDivergence);
        Assert.NotNull(run.Proof.RejectToBaseline?.FirstDivergence);
        Assert.Contains(run.Proof.AcceptToFinal?.Findings ?? [], item =>
            item.Code == "normalized_whole_package_mismatch");
    }

    [Fact]
    public void RP005_NativeGeneratedRevision_RoundTripsNormalizedWholePackage()
    {
        var baseline = Document("The ", "original", " clause.");
        var intendedFinal = RewriteBody(baseline, new Paragraph(
            RunForText("The "),
            RunForText("revised"),
            RunForText(" clause.")));
        var redline = RewriteBody(baseline, new Paragraph(
            RunForText("The "),
            new DeletedRun(new Run(new DeletedText("original")))
            {
                Id = "1",
                Author = "Comparison Engine",
                Date = DateTime.Parse("2000-01-01T00:00:00Z",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.AdjustToUniversal),
            },
            new InsertedRun(new Run(new Text("revised")))
            {
                Id = "2",
                Author = "Comparison Engine",
                Date = DateTime.Parse("2000-01-01T00:00:00Z",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.AdjustToUniversal),
            },
            RunForText(" clause.")));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.AcceptToFinal?.NormalizedWholePackageEquivalent,
            run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.NormalizedWholePackageEquivalent,
            run.Proof.ToJson());
        Assert.Empty(run.Proof.AcceptToFinal?.Divergences
            .Where(item => item.UnknownOrUnmodeled) ?? []);
        Assert.Empty(run.Proof.RejectToBaseline?.Divergences
            .Where(item => item.UnknownOrUnmodeled) ?? []);
        Assert.True(run.Proof.AcceptToFinal?.ModeledSemantic.Equivalent,
            run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.ModeledSemantic.Equivalent,
            run.Proof.ToJson());
        Assert.All(run.Proof.AcceptToFinal?.Findings ?? [], finding =>
        {
            if (finding.Code == "raw_package_bytes_mismatch")
                Assert.Empty(finding.RevisionIds);
        });
        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP006_PreExistingMultiAuthorRevision_IsPreservedOnBothPaths()
    {
        var baseline = DocumentWithReviewBody(PriorReviewParagraph(),
            new Paragraph(RunForText("old")));
        var intendedFinal = RewriteBody(baseline, PriorReviewParagraph(),
            new Paragraph(RunForText("new")));
        var redline = RewriteBody(baseline, PriorReviewParagraph(), new Paragraph(
            new DeletedRun(new Run(new DeletedText("old")))
            {
                Id = "101",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            },
            new InsertedRun(RunForText("new"))
            {
                Id = "102",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            }));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.Contains(run.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.PreExisting
            && item.Redline?.Author == "Prior Reviewer");
        Assert.Contains(run.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Generated
            && item.Redline?.Author == "Comparison Engine");
        Assert.DoesNotContain(run.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Conflicted);
        Assert.True(run.Proof.AcceptToFinal?.PreExistingRevisionsPreserved,
            run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.PreExistingRevisionsPreserved,
            run.Proof.ToJson());
        Assert.Single(run.Proof.AcceptToFinal?.SurvivingPreExistingRevisions ?? []);
        Assert.Single(run.Proof.RejectToBaseline?.SurvivingPreExistingRevisions ?? []);
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Theory]
    [InlineData("RP015-MoveFrom-MoveTo", RevisionFamily.Move, false, false)]
    [InlineData("RP021-Inserted-Numbering-Properties", RevisionFamily.NumberingPropertiesInsert, true, false)]
    [InlineData("RP025-Paragraph-Props-Change", RevisionFamily.PropertiesChange, false, false)]
    [InlineData("RP027-Change-Section", RevisionFamily.PropertiesChange, true, false)]
    [InlineData("RP009-Deleted-Table-Row", RevisionFamily.RowDelete, true, true)]
    [InlineData("RP050-Deleted-Footnote", RevisionFamily.ContentDelete, false, false)]
    public void RP007_RealRevisionFamilies_ResolveAndEmitProofEvidence(
        string fixtureStem,
        RevisionFamily expectedFamily,
        bool acceptEquivalent,
        bool rejectEquivalent)
    {
        var root = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "../../../../TestFiles/RP"));
        var baseline = File.ReadAllBytes(Path.Combine(root, fixtureStem + "-Rejected.docx"));
        var intendedFinal = File.ReadAllBytes(Path.Combine(root, fixtureStem + "-Accepted.docx"));
        var redline = File.ReadAllBytes(Path.Combine(root, fixtureStem + ".docx"));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.Contains(run.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Generated
            && item.Redline?.Family == expectedFamily);
        Assert.True(run.Proof.AcceptToFinal?.Completed, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Completed, run.Proof.ToJson());
        AssertResolutionClosure(run.Proof.AcceptToFinal!);
        AssertResolutionClosure(run.Proof.RejectToBaseline!);
        Assert.NotNull(run.Proof.AcceptToFinal?.ActualPackage);
        Assert.NotNull(run.Proof.RejectToBaseline?.ActualPackage);
        AssertPathEvidenceCoherent(run.Proof.AcceptToFinal!, run.Proof.ToJson());
        AssertPathEvidenceCoherent(run.Proof.RejectToBaseline!, run.Proof.ToJson());
        Assert.True(
            run.Proof.AcceptToFinal!.Equivalent == acceptEquivalent,
            run.Proof.ToJson());
        Assert.True(
            run.Proof.RejectToBaseline!.Equivalent == rejectEquivalent,
            run.Proof.ToJson());
        Assert.True(
            run.Proof.Success == (acceptEquivalent && rejectEquivalent),
            run.Proof.ToJson());
        if (!acceptEquivalent)
        {
            Assert.Contains(run.Proof.AcceptToFinal.Findings,
                finding => finding.Code == "normalized_whole_package_mismatch");
        }

        if (!rejectEquivalent)
        {
            Assert.Contains(run.Proof.RejectToBaseline.Findings,
                finding => finding.Code == "normalized_whole_package_mismatch");
        }
    }

    [Fact]
    public void RP008_CommentedClause_PreservesDefinitionsAndRangeMarkers()
    {
        var (baseline, intendedFinal) = DocxDiffCommentFixtures.Build(
            "multi-comment-one-para");
        var redline = DocxDiff.Compare(
            baseline,
            intendedFinal,
            new DocxDiffSettings { AuthorForRevisions = "Comparison Engine" });

        var run = RedlineReversibilityVerifier.Prove(
            baseline.DocumentByteArray,
            intendedFinal.DocumentByteArray,
            redline.DocumentByteArray);

        Assert.True(run.Proof.AcceptToFinal?.Completed, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Completed, run.Proof.ToJson());
        Assert.Equal(CommentEvidence(intendedFinal.DocumentByteArray),
            CommentEvidence(run.AcceptedPackageBytes!));
        Assert.Equal(CommentEvidence(baseline.DocumentByteArray),
            CommentEvidence(run.RejectedPackageBytes!));
        Assert.True(run.Proof.AcceptToFinal?.ModeledSemantic.Available,
            run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.ModeledSemantic.Available,
            run.Proof.ToJson());
        AssertPathEvidenceCoherent(run.Proof.AcceptToFinal!, run.Proof.ToJson());
        AssertPathEvidenceCoherent(run.Proof.RejectToBaseline!, run.Proof.ToJson());
        Assert.False(run.Proof.AcceptToFinal!.Equivalent);
        Assert.False(run.Proof.RejectToBaseline!.Equivalent);
        Assert.False(run.Proof.Success);
        Assert.Contains(run.Proof.AcceptToFinal.Findings,
            finding => finding.Code == "modeled_semantic_mismatch");
        Assert.Contains(run.Proof.RejectToBaseline.Findings,
            finding => finding.Code == "modeled_semantic_mismatch");
    }

    [Fact]
    public void RP009_OpaquePartDifference_RemainsExplicitlyUnmodeled()
    {
        var baseline = AddOpaqueCustomXmlPart(Document("Unchanged body."), "baseline");
        var redline = RewriteOpaqueCustomXmlPart(baseline, "changed");

        var run = RedlineReversibilityVerifier.Prove(baseline, baseline, redline);

        Assert.True(run.Proof.AcceptToFinal?.Completed, run.Proof.ToJson());
        Assert.True(run.Proof.AcceptToFinal?.ModeledSemantic.Equivalent);
        var opaque = Assert.Single(run.Proof.AcceptToFinal?.Divergences ?? [], divergence =>
            divergence.PartUri.StartsWith("/customXml/", StringComparison.Ordinal)
            && divergence.UnknownOrUnmodeled);
        Assert.False(opaque.HasModeledSemanticChange);
        Assert.False(run.Proof.Success);
    }

    [Fact]
    public void RP010_HeaderAndFooterRevisions_RoundTripBothStoryParts()
    {
        var baseline = HeaderFooterDocument("Old header", "Old footer", tracked: false);
        var intendedFinal = RewriteHeaderFooter(
            baseline, "New header", "New footer", tracked: false);
        var redline = RewriteHeaderFooter(
            baseline, "New header", "New footer", tracked: true);

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.PartUri.StartsWith("/word/header", StringComparison.Ordinal)
                == true);
        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.PartUri.StartsWith("/word/footer", StringComparison.Ordinal)
                == true);
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP011_BookmarkAroundGeneratedText_IsPreservedOnBothPaths()
    {
        var baseline = BookmarkDocument("Old bookmarked text", tracked: false);
        var intendedFinal = RewriteBody(
            baseline, BookmarkParagraph("New bookmarked text", tracked: false));
        var redline = RewriteBody(
            baseline, BookmarkParagraph("New bookmarked text", tracked: true));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.Equal(new[] { "ContractClause" }, BookmarkNames(run.AcceptedPackageBytes!));
        Assert.Equal(new[] { "ContractClause" }, BookmarkNames(run.RejectedPackageBytes!));
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP012_AnchoredSemanticMismatch_ReportsOnlyApplicableRevisions()
    {
        var baseline = DocumentWithReviewBody(
            new Paragraph(RunForText("Old clause.")),
            new Paragraph(RunForText("Old schedule.")));
        var intendedFinal = RewriteBody(
            baseline,
            new Paragraph(RunForText("Expected clause.")),
            new Paragraph(RunForText("Expected schedule.")));
        var redline = RewriteBody(
            baseline,
            new Paragraph(
                new DeletedRun(new Run(new DeletedText("Old clause.")))
                {
                    Id = "501",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                },
                new InsertedRun(RunForText("Actual clause."))
                {
                    Id = "502",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                }),
            new Paragraph(
                new DeletedRun(new Run(new DeletedText("Old schedule.")))
                {
                    Id = "601",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                },
                new InsertedRun(RunForText("Expected schedule."))
                {
                    Id = "602",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                }));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        var modeledChanges = SemanticDiff.Compare(
                new WmlDocument("expected.docx", intendedFinal),
                new WmlDocument("actual.docx", run.AcceptedPackageBytes!))
            .Changes
            .Where(change => change.Family != SemanticChangeFamily.OpaquePackagePart)
            .ToArray();
        var firstSemanticChange = modeledChanges.FirstOrDefault(change =>
                change.RightAnchor is not null || change.LeftAnchor is not null)
            ?? Assert.Single(modeledChanges);
        var semanticAnchor = firstSemanticChange.RightAnchor ?? firstSemanticChange.LeftAnchor;
        Assert.NotNull(semanticAnchor);

        var finding = Assert.Single(run.Proof.AcceptToFinal?.Findings ?? [], item =>
            item.Code == "modeled_semantic_mismatch");
        Assert.Equal("/word/document.xml", finding.Location?.EntryUri);
        Assert.NotNull(finding.Location?.PropertyPath);
        Assert.Equal(semanticAnchor, finding.AnchorId);

        var generated = run.Proof.RevisionClassifications
            .Where(classification =>
                classification.Disposition == RedlineRevisionDisposition.Generated)
            .Select(classification => classification.Redline!)
            .ToArray();
        Assert.True(generated.Select(revision => revision.AnchorId)
            .Where(anchor => anchor is not null)
            .Distinct(StringComparer.Ordinal)
            .Count() >= 2, run.Proof.ToJson());
        var expectedApplicableIds = generated.Where(revision =>
                revision.ConstituentIds.Any(id => id is "501" or "502"))
            .Select(revision => revision.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        Assert.NotEmpty(expectedApplicableIds);
        Assert.Equal(expectedApplicableIds,
            finding.RevisionIds.OrderBy(id => id, StringComparer.Ordinal));

        var unrelated = generated.Where(revision => revision.ConstituentIds.Any(
                id => id is "601" or "602"))
            .ToArray();
        Assert.NotEmpty(unrelated);
        Assert.All(unrelated, revision =>
            Assert.DoesNotContain(revision.Id, finding.RevisionIds));
        Assert.False(run.Proof.AcceptToFinal?.Equivalent);
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
    }

    [Fact]
    public void RP013_RevisionElementBudget_FailsBeforePathExecution()
    {
        var baseline = Document("Base");
        var intendedFinal = Document("BaseAB");
        var redline = RewriteBody(baseline, new Paragraph(
            RunForText("Base"),
            new InsertedRun(RunForText("A"))
            {
                Id = "701",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            },
            new InsertedRun(RunForText("B"))
            {
                Id = "702",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            }));

        var run = RedlineReversibilityVerifier.Prove(
            baseline,
            intendedFinal,
            redline,
            new RedlineReversibilityProofOptions { MaxRevisionElements = 1 });

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.Proof.RejectToBaseline);
        Assert.Null(run.AcceptedPackageBytes);
        Assert.Null(run.RejectedPackageBytes);
        Assert.Empty(run.Proof.RevisionClassifications);
        var finding = Assert.Single(run.Proof.Findings,
            item => item.Code == "revision_element_limit_exceeded");
        Assert.Equal("redline/revisions", finding.Location?.PropertyPath);
    }

    [Fact]
    public void RP014_UncountedRevisionMarkers_CannotBypassLiveInventoryBudget()
    {
        var baseline = IrTestDocuments.Create("Base").DocumentByteArray;
        var redline = IrTestDocuments.FromBodyXml(
            "<w:p><w:moveFromRangeStart w:id=\"701\"/>"
            + "<w:moveFromRangeStart w:id=\"702\"/>"
            + "<w:r><w:t>Base</w:t></w:r></w:p>").DocumentByteArray;

        Assert.Equal(0, PackageManifestGenerator.Generate(redline).Facts.Revisions.Total);
        var run = RedlineReversibilityVerifier.Prove(
            baseline,
            baseline,
            redline,
            new RedlineReversibilityProofOptions { MaxRevisionElements = 1 });

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.Proof.RejectToBaseline);
        Assert.Empty(run.Proof.RevisionClassifications);
        var finding = Assert.Single(run.Proof.Findings,
            item => item.Code == "revision_element_limit_exceeded");
        Assert.Equal("redline/revisions", finding.Location?.PropertyPath);
    }

    [Fact]
    public void RP015_PartialPathProjection_DoesNotLabelGeneratedAsPreExisting()
    {
        var preExisting = RevisionIdentity("prior");
        var generated = RevisionIdentity("generated");

        var projected = RedlineReversibilityVerifier.SelectSurvivingPreExisting(
            new[] { preExisting },
            new[] { generated, preExisting });

        Assert.Equal(preExisting, Assert.Single(projected));
    }

    [Fact]
    public void RP016_PreExistingRevisionInEditedParagraph_IsNotOwnershipConflict()
    {
        var baseline = DocumentWithReviewBody(new Paragraph(
            RunForText("Base"),
            new InsertedRun(RunForText(" prior"))
            {
                Id = "800",
                Author = "Prior Reviewer",
                Date = FixedRevisionDate(),
            },
            RunForText(" old")));
        var intendedFinal = RewriteBody(baseline, new Paragraph(
            RunForText("Base"),
            new InsertedRun(RunForText(" prior"))
            {
                Id = "800",
                Author = "Prior Reviewer",
                Date = FixedRevisionDate(),
            },
            RunForText(" new")));
        var redline = RewriteBody(baseline, new Paragraph(
            RunForText("Base"),
            new InsertedRun(RunForText(" prior"))
            {
                Id = "800",
                Author = "Prior Reviewer",
                Date = FixedRevisionDate(),
            },
            new DeletedRun(new Run(new DeletedText(" old")
            {
                Space = SpaceProcessingModeValues.Preserve,
            }))
            {
                Id = "801",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            },
            new InsertedRun(RunForText(" new"))
            {
                Id = "802",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            }));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.Contains(run.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.PreExisting
            && item.Redline?.ConstituentIds.Contains("800", StringComparer.Ordinal) == true);
        Assert.DoesNotContain(run.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Conflicted);
        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP017_RawInputBudget_IsAppliedBeforePackageInspection()
    {
        var error = Assert.Throws<ArgumentException>(() =>
            RedlineReversibilityVerifier.Prove(
                new byte[4],
                new byte[4],
                new byte[4],
                new RedlineReversibilityProofOptions { MaxPackageBytes = 11 }));

        Assert.Contains("Aggregate baseline", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RP018_PackageDeltaLimit_FailsClosedWithoutPartialDivergences()
    {
        var baseline = Document("Base");
        var redline = AddOpaqueCustomXmlPart(baseline, "extra");

        var run = RedlineReversibilityVerifier.Prove(
            baseline,
            baseline,
            redline,
            new RedlineReversibilityProofOptions { MaxPackageChanges = 1 });

        Assert.False(run.Proof.AcceptToFinal?.DivergenceAnalysisCompleted);
        Assert.Empty(run.Proof.AcceptToFinal?.Divergences ?? []);
        Assert.Contains(run.Proof.AcceptToFinal?.Findings ?? [], item =>
            item.Code == "package_divergence_limit_exceeded");
        Assert.False(run.Proof.Success);
    }

    [Fact]
    public void RP019_SemanticChangeLimit_ReportsUnavailableEvidence()
    {
        var baseline = DocumentWithReviewBody(
            new Paragraph(RunForText("A")),
            new Paragraph(RunForText("B")),
            new Paragraph(RunForText("C")));
        var intendedFinal = RewriteBody(
            baseline,
            new Paragraph(RunForText("X")),
            new Paragraph(RunForText("Y")),
            new Paragraph(RunForText("Z")));

        var run = RedlineReversibilityVerifier.Prove(
            baseline,
            intendedFinal,
            baseline,
            new RedlineReversibilityProofOptions { MaxSemanticChanges = 1 });

        Assert.False(run.Proof.AcceptToFinal?.ModeledSemantic.Available);
        Assert.Null(run.Proof.AcceptToFinal?.ModeledSemantic.ChangeCount);
        Assert.Contains(run.Proof.AcceptToFinal?.Findings ?? [], item =>
            item.Code == "modeled_semantic_comparison_unavailable");
        Assert.False(run.Proof.Success);
    }

    [Fact]
    public void RP020_RevisionEvidenceLimit_FailsBeforePathExecution()
    {
        var baseline = Document("Base");
        var intendedFinal = Document("Base plus evidence");
        var redline = RewriteBody(baseline, new Paragraph(
            RunForText("Base"),
            new InsertedRun(RunForText(" plus evidence"))
            {
                Id = "900",
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            }));

        var run = RedlineReversibilityVerifier.Prove(
            baseline,
            intendedFinal,
            redline,
            new RedlineReversibilityProofOptions { MaxEvidenceTextCharacters = 8 });

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Empty(run.Proof.RevisionClassifications);
        Assert.Contains(run.Proof.Findings, item =>
            item.Code == "revision_evidence_limit_exceeded");
    }

    [Fact]
    public void RP021_CanonicalUtf8_IsCamelCaseAndMatchesCanonicalString()
    {
        var document = Document("same");
        var proof = RedlineReversibilityVerifier.Prove(document, document, document).Proof;

        Assert.Equal(proof.ToCanonicalJson(), Encoding.UTF8.GetString(proof.ToCanonicalUtf8Bytes()));
        Assert.Contains("\"direction\":\"acceptToFinal\"", proof.ToCanonicalJson(),
            StringComparison.Ordinal);
        Assert.DoesNotContain("AcceptToFinal", proof.ToCanonicalJson(), StringComparison.Ordinal);
    }

    [Fact]
    public void RP022_GeneratedGroupsMayMergeAfterSeparatorResolution()
    {
        const string stamp = " w:author=\"Comparison Engine\""
            + " w:date=\"2000-01-01T00:00:00Z\"";
        var baseline = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>A</w:t></w:r></w:p>").DocumentByteArray;
        var intendedFinal = RewriteBodyXml(
            baseline,
            "<w:p><w:r><w:t>X</w:t></w:r><w:r><w:t>Z</w:t></w:r></w:p>");
        var redline = RewriteBodyXml(
            baseline,
            "<w:p>"
            + $"<w:ins w:id=\"1\"{stamp}><w:r><w:t>X</w:t></w:r></w:ins>"
            + $"<w:del w:id=\"2\"{stamp}><w:r><w:delText>A</w:delText></w:r></w:del>"
            + $"<w:ins w:id=\"3\"{stamp}><w:r><w:t>Z</w:t></w:r></w:ins>"
            + "</w:p>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        AssertResolutionClosure(run.Proof.AcceptToFinal!);
        AssertResolutionClosure(run.Proof.RejectToBaseline!);
        Assert.Equal(3, run.Proof.RevisionClassifications.Count(classification =>
            classification.Disposition == RedlineRevisionDisposition.Generated));
    }

    [Fact]
    public void RP023_NestedPropertyRevisionFailsClosed()
    {
        var clean = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Clause</w:t></w:r></w:p><w:sectPr/>")
            .DocumentByteArray;
        var malformed = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Clause</w:t></w:r></w:p>"
            + "<w:sectPr><w:sectPrChange w:id=\"1\" w:author=\"Reviewer\">"
            + "<w:sectPr><w:sectPrChange w:id=\"2\" w:author=\"Reviewer\">"
            + "<w:sectPr/></w:sectPrChange></w:sectPr>"
            + "</w:sectPrChange></w:sectPr>").DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(clean, clean, malformed);

        Assert.Null(run.Proof.AcceptToFinal);
        var classification = Assert.Single(run.Proof.RevisionClassifications);
        Assert.Equal(RevisionResolutionStatus.Malformed,
            classification.Redline?.ResolutionStatus);
        Assert.Equal("malformed_properties_change",
            classification.Redline?.Diagnostic?.Code);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "generated_revision_not_resolvable");
    }

    [Fact]
    public void RP024_RejectPropertyChangeRestoresStoredAttributes()
    {
        var baseline = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Clause</w:t></w:r></w:p>"
            + "<w:sectPr w:rsidSect=\"AAAA\"><w:pgSz w:w=\"100\"/></w:sectPr>")
            .DocumentByteArray;
        var intendedFinal = RewriteBodyXml(
            baseline,
            "<w:p><w:r><w:t>Clause</w:t></w:r></w:p>"
            + "<w:sectPr w:rsidSect=\"BBBB\"><w:pgSz w:w=\"200\"/></w:sectPr>");
        var redline = RewriteBodyXml(
            baseline,
            "<w:p><w:r><w:t>Clause</w:t></w:r></w:p>"
            + "<w:sectPr w:rsidSect=\"BBBB\"><w:pgSz w:w=\"200\"/>"
            + "<w:sectPrChange w:id=\"1\" w:author=\"Comparison Engine\" "
            + "w:date=\"2000-01-01T00:00:00Z\">"
            + "<w:sectPr w:rsidSect=\"AAAA\"><w:pgSz w:w=\"100\"/></w:sectPr>"
            + "</w:sectPrChange></w:sectPr>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
    }

    [Fact]
    public void RP025_OrphanDeletionPayloadCannotBypassNativeCarrierBudget()
    {
        var clean = IrTestDocuments.Create("Base").DocumentByteArray;
        var malformed = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:delText>A</w:delText><w:delText>B</w:delText></w:r></w:p>")
            .DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(
            clean,
            clean,
            malformed,
            new RedlineReversibilityProofOptions { MaxRevisionElements = 1 });

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Empty(run.Proof.RevisionClassifications);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "revision_element_limit_exceeded");
    }

    [Fact]
    public void RP026_OfficeConflictRevisionIsExplicitlyUnsupported()
    {
        var clean = IrTestDocuments.Create("Base").DocumentByteArray;
        var conflicted = IrTestDocuments.FromBodyXml(
            "<w:p><w14:conflictIns xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
            + "w:id=\"1\" w:author=\"Reviewer\"><w:r><w:t>Base</w:t></w:r>"
            + "</w14:conflictIns></w:p>").DocumentByteArray;

        var run = RedlineReversibilityVerifier.Prove(clean, clean, conflicted);

        Assert.Null(run.Proof.AcceptToFinal);
        var classification = Assert.Single(run.Proof.RevisionClassifications);
        Assert.Equal(RevisionResolutionStatus.Unsupported,
            classification.Redline?.ResolutionStatus);
        Assert.Equal("unsupported_revision_family",
            classification.Redline?.Diagnostic?.Code);
    }

    [Fact]
    public void RP027_ProofResolutionNeverSweepsUnrelatedOrphanRelationships()
    {
        const string baselineBody = "<w:p><w:r><w:t>Old</w:t></w:r></w:p>";
        const string finalBody = "<w:p><w:r><w:t>New</w:t></w:r></w:p>";
        const string redlineBody = "<w:p>"
            + "<w:del w:id=\"1\" w:author=\"Comparison Engine\"><w:r><w:delText>Old</w:delText></w:r></w:del>"
            + "<w:ins w:id=\"2\" w:author=\"Comparison Engine\"><w:r><w:t>New</w:t></w:r></w:ins>"
            + "</w:p>";
        var baseline = IrTestDocuments.FromBodyXmlWithHyperlinks(
            baselineBody, ("rOrphan", "https://example.test/unused")).DocumentByteArray;
        var intendedFinal = RewriteBodyXml(baseline, finalBody);
        var redline = RewriteBodyXml(baseline, redlineBody);

        var preserved = RedlineReversibilityVerifier.Prove(
            baseline, intendedFinal, redline);

        Assert.True(preserved.Proof.Success, preserved.Proof.ToJson());

        var redlineWithUnownedExtra = AddOrphanHyperlink(
            redline, "rExtra", "https://example.test/unowned");
        var unowned = RedlineReversibilityVerifier.Prove(
            baseline, intendedFinal, redlineWithUnownedExtra);

        Assert.False(unowned.Proof.Success);
        Assert.False(unowned.Proof.AcceptToFinal?.NormalizedWholePackageEquivalent);
        Assert.False(unowned.Proof.RejectToBaseline?.NormalizedWholePackageEquivalent);
    }

    [Fact]
    public void RP028_DuplicateSiblingPropertyChangesFailClosed()
    {
        var clean = IrTestDocuments.Create("Clause").DocumentByteArray;
        var malformed = RewriteBodyXml(
            clean,
            "<w:p><w:pPr>"
            + "<w:pPrChange w:id=\"1\" w:author=\"Reviewer\"><w:pPr/></w:pPrChange>"
            + "<w:pPrChange w:id=\"2\" w:author=\"Reviewer\"><w:pPr/></w:pPrChange>"
            + "</w:pPr><w:r><w:t>Clause</w:t></w:r></w:p>");

        var run = RedlineReversibilityVerifier.Prove(clean, clean, malformed);

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.NotEmpty(run.Proof.RevisionClassifications);
        Assert.All(run.Proof.RevisionClassifications, classification =>
        {
            Assert.Equal(RevisionResolutionStatus.Malformed,
                classification.Redline?.ResolutionStatus);
            Assert.Equal("malformed_properties_change",
                classification.Redline?.Diagnostic?.Code);
        });
    }

    [Fact]
    public void RP029_InvalidOriginalCellMergeStateFailsClosed()
    {
        const string table = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:cellMerge w:id=\"1\" w:author=\"Reviewer\" "
            + "w:vMerge=\"rest\" w:vMergeOrig=\"bogus\"/></w:tcPr>"
            + "<w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:p><w:r><w:t>Survivor</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var clean = IrTestDocuments.Create("Cell Survivor").DocumentByteArray;
        var malformed = RewriteBodyXml(clean, table);

        var run = RedlineReversibilityVerifier.Prove(clean, clean, malformed);

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.ResolutionStatus == RevisionResolutionStatus.Malformed
            && classification.Redline.Diagnostic?.Code == "invalid_cell_merge_state");
    }

    [Fact]
    public void RP030_CrossedMoveRangesAcrossNamesFailClosed()
    {
        const string stamp = " w:author=\"Reviewer\" w:date=\"2000-01-01T00:00:00Z\"";
        var clean = IrTestDocuments.Create("AB").DocumentByteArray;
        var crossed = RewriteBodyXml(
            clean,
            "<w:p>"
            + $"<w:moveFromRangeStart w:id=\"1\" w:name=\"A\"{stamp}/>"
            + $"<w:moveFrom w:id=\"2\"{stamp}><w:r><w:t>A</w:t></w:r></w:moveFrom>"
            + $"<w:moveFromRangeStart w:id=\"3\" w:name=\"B\"{stamp}/>"
            + $"<w:moveFrom w:id=\"4\"{stamp}><w:r><w:t>B</w:t></w:r></w:moveFrom>"
            + "<w:moveFromRangeEnd w:id=\"1\"/><w:moveFromRangeEnd w:id=\"3\"/>"
            + $"<w:moveToRangeStart w:id=\"5\" w:name=\"A\"{stamp}/>"
            + $"<w:moveTo w:id=\"6\"{stamp}><w:r><w:t>A</w:t></w:r></w:moveTo>"
            + "<w:moveToRangeEnd w:id=\"5\"/>"
            + $"<w:moveToRangeStart w:id=\"7\" w:name=\"B\"{stamp}/>"
            + $"<w:moveTo w:id=\"8\"{stamp}><w:r><w:t>B</w:t></w:r></w:moveTo>"
            + "<w:moveToRangeEnd w:id=\"7\"/>"
            + "</w:p>");

        var run = RedlineReversibilityVerifier.Prove(clean, clean, crossed);

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.ResolutionStatus == RevisionResolutionStatus.Ambiguous
            && classification.Redline.Diagnostic?.Code == "crossed_move_range_topology");
    }

    [Fact]
    public void RP031_MoveCannotReuseOneIdForBothLogicalRanges()
    {
        const string stamp = " w:author=\"Reviewer\"";
        var clean = IrTestDocuments.Create("A").DocumentByteArray;
        var ambiguous = RewriteBodyXml(
            clean,
            "<w:p>"
            + $"<w:moveFromRangeStart w:id=\"1\" w:name=\"M\"{stamp}/>"
            + $"<w:moveFrom w:id=\"2\"{stamp}><w:r><w:t>A</w:t></w:r></w:moveFrom>"
            + "<w:moveFromRangeEnd w:id=\"1\"/>"
            + $"<w:moveToRangeStart w:id=\"1\" w:name=\"M\"{stamp}/>"
            + $"<w:moveTo w:id=\"3\"{stamp}><w:r><w:t>A</w:t></w:r></w:moveTo>"
            + "<w:moveToRangeEnd w:id=\"1\"/>"
            + "</w:p>");

        var run = RedlineReversibilityVerifier.Prove(clean, clean, ambiguous);

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.Diagnostic?.Code == "ambiguous_move_range_id");
    }

    [Fact]
    public void RP032_StrictOrphanRangeEndpointIsRejectedBeforeSessionOpen()
    {
        var clean = IrTestDocuments.Create("Base").DocumentByteArray;
        var strict = RewriteBodyXmlWithNamespace(
            clean,
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
            "<w:p><w:moveFromRangeEnd w:id=\"1\"/><w:r><w:t>Base</w:t></w:r></w:p>");

        var run = RedlineReversibilityVerifier.Prove(clean, clean, strict);

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "unsupported_strict_revision_markup");
    }

    [Fact]
    public void RP033_RevisionCarrierInUnmodeledWordPartFailsClosed()
    {
        var clean = IrTestDocuments.Create("Base").DocumentByteArray;
        var settingsRevision = RewriteSettingsXml(
            clean,
            "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:ins w:id=\"1\" w:author=\"Reviewer\"/></w:settings>");

        var run = RedlineReversibilityVerifier.Prove(
            settingsRevision, settingsRevision, settingsRevision);

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Empty(run.Proof.RevisionClassifications);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "unsupported_revision_part");
    }

    [Fact]
    public void RP034_FinalOnlyReviewStateGetsExplicitRejectPolicyEvidence()
    {
        var baseline = IrTestDocuments.Create("Base").DocumentByteArray;
        var final = RewriteBodyXml(
            baseline,
            "<w:p><w:ins w:id=\"7\" w:author=\"Reviewer B\">"
            + "<w:r><w:t>Final review</w:t></w:r></w:ins></w:p>");

        var run = RedlineReversibilityVerifier.Prove(baseline, final, final);

        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Disposition == RedlineRevisionDisposition.IntendedFinalPreExisting);
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.False(run.Proof.RejectToBaseline?.Equivalent);
        Assert.Contains(run.Proof.RejectToBaseline?.Findings ?? [], finding =>
            finding.Code == "intended_final_revision_survived_reject_path");
    }

    [Fact]
    public void RP035_UnsupportedConflictIdsAreReservedAndExhaustionFailsSafely()
    {
        var document = IrTestDocuments.FromBodyXml(
            "<w:p><w14:conflictIns xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
            + "w:id=\"2147483647\" w:author=\"Reviewer\">"
            + "<w:r><w:t>Base</w:t></w:r></w14:conflictIns></w:p>").DocumentByteArray;
        using var session = new DocxSession(document);

        var error = Assert.Throws<InvalidOperationException>(() => session.NextRevisionId());

        Assert.Contains("exhausted", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RP036_PropertyRevisionPreservesXmlNodesInBothImages()
    {
        const string prefix = "<w:p><w:pPr>";
        const string suffix = "</w:pPr><w:r><w:t>Clause</w:t></w:r></w:p>";
        var seed = IrTestDocuments.Create("Clause").DocumentByteArray;
        var baseline = RewriteBodyXml(seed, prefix + "<!--old-->" + suffix);
        var intendedFinal = RewriteBodyXml(seed, prefix + "<!--new-->" + suffix);
        var redline = RewriteBodyXml(
            seed,
            prefix + "<!--new-->"
            + "<w:pPrChange w:id=\"1\" w:author=\"Comparison Engine\">"
            + "<w:pPr><!--old--></w:pPr></w:pPrChange>" + suffix);

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
    }

    [Fact]
    public void RP037_NativeMoveRoundTripsBothEndpoints()
    {
        const string stamp = " w:author=\"Comparison Engine\""
            + " w:date=\"2000-01-01T00:00:00Z\"";
        var seed = IrTestDocuments.Create("AB").DocumentByteArray;
        var baseline = RewriteBodyXml(
            seed,
            "<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>");
        var intendedFinal = RewriteBodyXml(
            seed,
            "<w:p><w:r><w:t>B</w:t></w:r><w:r><w:t>A</w:t></w:r></w:p>");
        var redline = RewriteBodyXml(
            seed,
            "<w:p>"
            + $"<w:moveFromRangeStart w:id=\"1\" w:name=\"M\"{stamp}/>"
            + $"<w:moveFrom w:id=\"2\"{stamp}><w:r><w:t>A</w:t></w:r></w:moveFrom>"
            + "<w:moveFromRangeEnd w:id=\"1\"/>"
            + "<w:r><w:t>B</w:t></w:r>"
            + $"<w:moveToRangeStart w:id=\"3\" w:name=\"M\"{stamp}/>"
            + $"<w:moveTo w:id=\"4\"{stamp}><w:r><w:t>A</w:t></w:r></w:moveTo>"
            + "<w:moveToRangeEnd w:id=\"3\"/>"
            + "</w:p>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Redline?.Family == RevisionFamily.Move);
        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP038_BareInsertedRowRoundTripsWithoutEmptyPropertyHusk()
    {
        const string tableStart = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>";
        const string firstRow = "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>";
        const string secondRow = "<w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>";
        var seed = IrTestDocuments.Create("A").DocumentByteArray;
        var baseline = RewriteBodyXml(seed, tableStart + firstRow + "</w:tbl>");
        var intendedFinal = RewriteBodyXml(
            seed, tableStart + firstRow + secondRow + "</w:tbl>");
        var redline = RewriteBodyXml(
            seed,
            tableStart + firstRow
            + "<w:tr><w:trPr><w:ins w:id=\"1\" w:author=\"Comparison Engine\"/>"
            + "</w:trPr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>"
            + "</w:tbl>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.True(run.Proof.AcceptToFinal?.Equivalent, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.Equivalent, run.Proof.ToJson());
    }

    [Fact]
    public void RP039_EmptyStoredParagraphPropertiesRoundTripWithoutHusk()
    {
        var seed = IrTestDocuments.Create("Clause").DocumentByteArray;
        var baseline = RewriteBodyXml(
            seed, "<w:p><w:r><w:t>Clause</w:t></w:r></w:p>");
        var intendedFinal = RewriteBodyXml(
            seed,
            "<w:p><w:pPr><w:keepNext/></w:pPr>"
            + "<w:r><w:t>Clause</w:t></w:r></w:p>");
        var redline = RewriteBodyXml(
            seed,
            "<w:p><w:pPr><w:keepNext/>"
            + "<w:pPrChange w:id=\"1\" w:author=\"Comparison Engine\"><w:pPr/>"
            + "</w:pPrChange></w:pPr><w:r><w:t>Clause</w:t></w:r></w:p>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP040_InsertedCellRoundTripsWithoutEmptyPropertyHusk()
    {
        const string tableStart = "<w:tbl><w:tblPr/><w:tblGrid>"
            + "<w:gridCol/><w:gridCol/></w:tblGrid><w:tr>";
        const string firstCell = "<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>";
        const string secondCell = "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>";
        var seed = IrTestDocuments.Create("A").DocumentByteArray;
        var baseline = RewriteBodyXml(
            seed, tableStart + firstCell + "</w:tr></w:tbl>");
        var intendedFinal = RewriteBodyXml(
            seed, tableStart + firstCell + secondCell + "</w:tr></w:tbl>");
        var redline = RewriteBodyXml(
            seed,
            tableStart + firstCell
            + "<w:tc><w:tcPr><w:cellIns w:id=\"1\" w:author=\"Comparison Engine\"/>"
            + "</w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>"
            + "</w:tr></w:tbl>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP041_UnrelatedEmptyShellDoesNotProtectGeneratedPropertyHusk()
    {
        const string first = "<w:p><w:pPr/><w:r><w:t>First</w:t></w:r></w:p>";
        var seed = IrTestDocuments.Create("First Second").DocumentByteArray;
        var baseline = RewriteBodyXml(
            seed,
            first + "<w:p><w:r><w:t>Second</w:t></w:r></w:p>");
        var intendedFinal = RewriteBodyXml(
            seed,
            first + "<w:p><w:pPr><w:keepNext/></w:pPr>"
            + "<w:r><w:t>Second</w:t></w:r></w:p>");
        var redline = RewriteBodyXml(
            seed,
            first + "<w:p><w:pPr><w:keepNext/>"
            + "<w:pPrChange w:id=\"1\" w:author=\"Comparison Engine\"><w:pPr/>"
            + "</w:pPrChange></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP042_DistinctCarrierRolesMayReuseNumericRevisionId()
    {
        var seed = IrTestDocuments.Create("A").DocumentByteArray;
        var baseline = RewriteBodyXml(
            seed, "<w:p><w:r><w:t>A</w:t></w:r></w:p>");
        var intendedFinal = RewriteBodyXml(
            seed, "<w:p><w:r><w:t>B</w:t></w:r></w:p>");
        var redline = RewriteBodyXml(
            seed,
            "<w:p><w:del w:id=\"1\" w:author=\"Comparison Engine\">"
            + "<w:r><w:delText>A</w:delText></w:r></w:del>"
            + "<w:ins w:id=\"1\" w:author=\"Comparison Engine\">"
            + "<w:r><w:t>B</w:t></w:r></w:ins></w:p>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.All(run.Proof.RevisionClassifications, classification =>
            Assert.Equal(RevisionResolutionStatus.Supported,
                classification.Redline?.ResolutionStatus));
    }

    [Fact]
    public void RP043_RevisionAnchorEvidenceIsBoundedBeforePathExecution()
    {
        var baseline = Document("A");
        var intendedFinal = Document("AB");
        var redline = RewriteBodyXml(
            baseline,
            "<w:p><w:r><w:t>A</w:t></w:r>"
            + "<w:ins w:id=\"1\" w:author=\"Comparison Engine\">"
            + "<w:r><w:t>B</w:t></w:r></w:ins></w:p>");

        var run = RedlineReversibilityVerifier.Prove(
            baseline,
            intendedFinal,
            redline,
            new RedlineReversibilityProofOptions { MaxRevisionEvidenceItems = 2 });

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.Proof.RejectToBaseline);
        Assert.Empty(run.Proof.RevisionClassifications);
        Assert.Contains(run.Proof.Findings, item =>
            item.Code == "revision_evidence_limit_exceeded");
    }

    [Fact]
    public void RP044_HyperlinkOwnedByGeneratedRevisionRoundTripsRelationship()
    {
        var baseline = DocumentWithReviewBody(new Paragraph());
        var intendedFinal = HyperlinkDocument(baseline, tracked: false);
        var redline = HyperlinkDocument(baseline, tracked: true);

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.True(run.Proof.AcceptToFinal?.NormalizedWholePackageEquivalent);
        Assert.True(run.Proof.RejectToBaseline?.NormalizedWholePackageEquivalent);
    }

    [Fact]
    public void RP045_InsertedParagraphBesideBlockSdtRoundTripsBothEndpoints()
    {
        const string blockSdt = "<w:sdt><w:sdtContent><w:p><w:r><w:t>Existing"
            + "</w:t></w:r></w:p></w:sdtContent></w:sdt>";
        const string insertedParagraph = "<w:p><w:r><w:t>New</w:t></w:r></w:p>";
        const string trackedParagraph = "<w:p><w:pPr><w:rPr>"
            + "<w:ins w:id=\"1\" w:author=\"Comparison Engine\"/>"
            + "</w:rPr></w:pPr><w:ins w:id=\"2\" w:author=\"Comparison Engine\">"
            + "<w:r><w:t>New</w:t></w:r></w:ins></w:p>";
        var seed = Document("seed");
        var baseline = RewriteBodyXml(seed, blockSdt);
        var intendedFinal = RewriteBodyXml(seed, blockSdt + insertedParagraph);
        var redline = RewriteBodyXml(seed, blockSdt + trackedParagraph);

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP046_ReusedEndpointOrphanRelationshipIsPreserved()
    {
        var baseline = AddOrphanHyperlink(
            DocumentWithReviewBody(new Paragraph()),
            "rIdRevisionLink",
            "https://example.test/revision");
        var intendedFinal = HyperlinkDocument(baseline, tracked: false);
        var redline = HyperlinkDocument(baseline, tracked: true);

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.NormalizedWholePackageEquivalent);
    }

    [Fact]
    public void RP047_ReusedEndpointAbstractNumberingDefinitionIsPreserved()
    {
        var baseline = DocumentWithOrphanAbstractNumbering();
        var intendedFinal = DocumentUsingOrphanAbstractNumbering(baseline, tracked: false);
        var redline = DocumentUsingOrphanAbstractNumbering(baseline, tracked: true);

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.True(run.Proof.RejectToBaseline?.NormalizedWholePackageEquivalent);
    }

    [Fact]
    public void RP048_ResolvingOuterCellRevisionPreservesNestedTableRevision()
    {
        var seed = Document("seed");
        var baseline = RewriteBodyXml(
            seed, NestedCellMergeBody("<w:vMerge w:val=\"continue\"/>"));
        var intendedFinal = RewriteBodyXml(
            seed, NestedCellMergeBody("<w:vMerge w:val=\"restart\"/>"));
        var redline = RewriteBodyXml(
            seed,
            NestedCellMergeBody("<w:cellMerge w:id=\"1\""
                + " w:author=\"Comparison Engine\" w:date=\"2000-01-01T00:00:00Z\""
                + " w:vMerge=\"rest\" w:vMergeOrig=\"cont\"/>"));

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
        Assert.Contains(run.Proof.RevisionClassifications, classification =>
            classification.Disposition == RedlineRevisionDisposition.PreExisting
            && classification.Redline?.ConstituentIds.Contains("2") == true);
    }

    [Fact]
    public void RP049_RejectingOnlyBodyParagraphRestoresEmptyBody()
    {
        var seed = Document("seed");
        var baseline = RewriteBodyXml(seed, "<w:sectPr/>");
        var intendedFinal = RewriteBodyXml(
            seed, "<w:p><w:r><w:t>New</w:t></w:r></w:p><w:sectPr/>");
        var redline = RewriteBodyXml(
            seed,
            "<w:p><w:pPr><w:rPr><w:ins w:id=\"1\""
            + " w:author=\"Comparison Engine\"/></w:rPr></w:pPr>"
            + "<w:ins w:id=\"2\" w:author=\"Comparison Engine\">"
            + "<w:r><w:t>New</w:t></w:r></w:ins></w:p><w:sectPr/>");

        var run = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline);

        Assert.True(run.Proof.Success, run.Proof.ToJson());
    }

    [Fact]
    public void RP050_StructuralAnchorTraversalIsBounded()
    {
        var endpoint = Document("endpoint");
        var runs = string.Concat(Enumerable.Repeat("<w:r><w:t>x</w:t></w:r>", 4_500));
        var redline = RewriteBodyXml(
            endpoint,
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>"
            + "<w:tcPr><w:cellMerge w:id=\"1\" w:author=\"Comparison Engine\""
            + " w:vMerge=\"rest\"/></w:tcPr><w:p>" + runs
            + "</w:p></w:tc></w:tr></w:tbl>");

        var run = RedlineReversibilityVerifier.Prove(
            endpoint,
            endpoint,
            redline,
            new RedlineReversibilityProofOptions { MaxRevisionEvidenceItems = 1_000 });

        Assert.Null(run.Proof.AcceptToFinal);
        Assert.Null(run.Proof.RejectToBaseline);
        Assert.Contains(run.Proof.Findings, finding =>
            finding.Code == "revision_evidence_limit_exceeded");
    }

    [Fact]
    public void RP051_IsExactCanonical_RejectsUndefinedIntegerEnumValues()
    {
        var baseline = Document("The original clause.");
        var intendedFinal = Document("The revised clause.");
        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("final.docx", intendedFinal),
            new DocxDiffSettings { AuthorForRevisions = "Comparison Engine" }).DocumentByteArray;
        var proof = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline).Proof;
        var canonical = proof.ToCanonicalUtf8Bytes();
        Assert.True(RedlineReversibilityProof.IsExactCanonical(canonical));

        var json = Encoding.UTF8.GetString(canonical);
        Assert.Contains("\"disposition\":\"generated\"", json);
        var forged = Encoding.UTF8.GetBytes(json.Replace(
            "\"disposition\":\"generated\"", "\"disposition\":9999"));
        Assert.False(RedlineReversibilityProof.IsExactCanonical(forged));
    }

    private static void AssertResolutionClosure(RedlineProofPathResult path)
    {
        var accountedFor = path.ResolvedRevisionIds
            .Concat(path.ImplicitlyResolvedRevisionIds)
            .OrderBy(item => item, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(path.RequestedRevisionIds.OrderBy(item => item, StringComparer.Ordinal),
            accountedFor);
        Assert.Equal(accountedFor.Length, accountedFor.Distinct(StringComparer.Ordinal).Count());
    }

    private static void AssertPathEvidenceCoherent(
        RedlineProofPathResult path,
        string proofJson)
    {
        var expectedEquivalent = path.Completed
            && path.PreExistingRevisionsPreserved
            && path.ModeledSemantic.Available
            && path.ModeledSemantic.Equivalent == true
            && path.NormalizedWholePackageEquivalent
            && path.DivergenceAnalysisCompleted
            && path.Findings.All(finding =>
                finding.Severity != VerificationFindingSeverity.Error);
        Assert.Equal(expectedEquivalent, path.Equivalent);
        if (!path.NormalizedWholePackageEquivalent)
        {
            Assert.NotNull(path.FirstDivergence);
            Assert.Contains(path.Findings, finding =>
                finding.Code == "normalized_whole_package_mismatch");
        }
        if (!path.Equivalent)
        {
            Assert.Contains(path.Findings, finding =>
                finding.Severity == VerificationFindingSeverity.Error);
        }
        Assert.True(path.ModeledSemantic.Available, proofJson);
    }

    private static RedlineRevisionIdentity RevisionIdentity(string id) => new()
    {
        Id = id,
        PartUri = "/word/document.xml",
        Scope = "body",
        Type = "insert",
        Family = RevisionFamily.ContentInsert,
        ConstituentIds = new[] { id },
        ConstituentKeys = new[] { "w:ins:" + id },
        Author = "Reviewer",
        Date = null,
        DateUtc = null,
        Text = id,
        AnchorId = "p:doc:1",
        AffectedAnchorIds = new[] { "p:doc:1" },
        ResolutionStatus = RevisionResolutionStatus.Supported,
        Diagnostic = null,
    };

    private static byte[] Document(params string[] runs) =>
        DocumentWithReviewBody(new Paragraph(runs.Select(RunForText)));

    private static byte[] DocumentWithReviewBody(params Paragraph[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraphs));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            document.Save();
        }
        return stream.ToArray();
    }

    private static Run RunForText(string value)
    {
        var text = new Text(value);
        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
            text.Space = SpaceProcessingModeValues.Preserve;
        return new Run(text);
    }

    private static Paragraph PriorReviewParagraph() => new(
        RunForText("Base"),
        new InsertedRun(RunForText(" prior"))
        {
            Id = "90",
            Author = "Prior Reviewer",
            Date = FixedRevisionDate(),
        });

    private static DateTime FixedRevisionDate() => DateTime.Parse(
        "2000-01-01T00:00:00Z",
        System.Globalization.CultureInfo.InvariantCulture,
        System.Globalization.DateTimeStyles.AdjustToUniversal);

    private static string CommentEvidence(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var definitions = main.WordprocessingCommentsPart?.Comments?
            .Elements<Comment>()
            .Select(comment => $"{comment.Id}|{comment.Author}|{comment.InnerText}")
            .OrderBy(item => item, StringComparer.Ordinal)
            ?? Enumerable.Empty<string>();
        var markers = main.Document.Descendants()
            .Select(element => element switch
            {
                CommentRangeStart start => $"start:{start.Id}",
                CommentRangeEnd end => $"end:{end.Id}",
                CommentReference reference => $"reference:{reference.Id}",
                _ => null,
            })
            .Where(item => item is not null);
        return string.Join("\n", definitions)
            + "\n--markers--\n"
            + string.Join("\n", markers);
    }

    private static byte[] AddOpaqueCustomXmlPart(byte[] source, string value)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!
                .AddCustomXmlPart(CustomXmlPartType.CustomXml);
            var payload = OpaquePayload(value);
            using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
            partStream.Write(payload);
        }
        return stream.ToArray();
    }

    private static byte[] RewriteOpaqueCustomXmlPart(byte[] source, string value)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = Assert.Single(document.MainDocumentPart!.CustomXmlParts);
            using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
            partStream.Write(OpaquePayload(value));
        }
        return stream.ToArray();
    }

    private static byte[] OpaquePayload(string value) => Encoding.UTF8.GetBytes(
        $"<vendor:payload xmlns:vendor=\"urn:docxodus:test:vendor\">{value}</vendor:payload>");

    private static byte[] HeaderFooterDocument(
        string headerText,
        string footerText,
        bool tracked)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            header.Header = new Header(RevisionParagraph(
                "Old header", headerText, tracked, "301", "302"));
            var footer = main.AddNewPart<FooterPart>("rIdFooter");
            footer.Footer = new Footer(RevisionParagraph(
                "Old footer", footerText, tracked, "303", "304"));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document = new Document(new Body(
                new Paragraph(RunForText("Unchanged body.")),
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader",
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdFooter",
                    })));
            document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] RewriteHeaderFooter(
        byte[] source,
        string headerText,
        string footerText,
        bool tracked)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            document.MainDocumentPart!.HeaderParts.Single().Header = new Header(
                RevisionParagraph("Old header", headerText, tracked, "301", "302"));
            document.MainDocumentPart.FooterParts.Single().Footer = new Footer(
                RevisionParagraph("Old footer", footerText, tracked, "303", "304"));
            document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] BookmarkDocument(string text, bool tracked) =>
        DocumentWithReviewBody(BookmarkParagraph(text, tracked));

    private static Paragraph BookmarkParagraph(string text, bool tracked)
    {
        var paragraph = new Paragraph(new BookmarkStart
        {
            Id = "0",
            Name = "ContractClause",
        });
        if (tracked)
        {
            paragraph.Append(
                new DeletedRun(new Run(new DeletedText("Old bookmarked text")))
                {
                    Id = "401",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                },
                new InsertedRun(RunForText(text))
                {
                    Id = "402",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                });
        }
        else
        {
            paragraph.Append(RunForText(text));
        }
        paragraph.Append(new BookmarkEnd { Id = "0" });
        return paragraph;
    }

    private static Paragraph RevisionParagraph(
        string oldText,
        string newText,
        bool tracked,
        string deleteId,
        string insertId) => tracked
        ? new Paragraph(
            new DeletedRun(new Run(new DeletedText(oldText)))
            {
                Id = deleteId,
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            },
            new InsertedRun(RunForText(newText))
            {
                Id = insertId,
                Author = "Comparison Engine",
                Date = FixedRevisionDate(),
            })
        : new Paragraph(RunForText(newText));

    private static byte[] HyperlinkDocument(byte[] source, bool tracked)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            if (!main.HyperlinkRelationships.Any(relationship =>
                    relationship.Id == "rIdRevisionLink"))
                main.AddHyperlinkRelationship(
                    new Uri("https://example.test/revision", UriKind.Absolute),
                    true,
                    "rIdRevisionLink");
            var hyperlink = new Hyperlink { Id = "rIdRevisionLink" };
            if (tracked)
            {
                hyperlink.Append(new InsertedRun(RunForText("Link"))
                {
                    Id = "1",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                });
            }
            else
            {
                hyperlink.Append(RunForText("Link"));
            }
            main.Document.Body = new Body(new Paragraph(hyperlink));
            document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] DocumentWithOrphanAbstractNumbering()
    {
        var source = Document("Item");
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var numbering = document.MainDocumentPart!
                .AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1." })
                    { LevelIndex = 0 })
                { AbstractNumberId = 5 });
            numbering.Numbering.Save();
        }
        return stream.ToArray();
    }

    private static byte[] DocumentUsingOrphanAbstractNumbering(byte[] source, bool tracked)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            main.NumberingDefinitionsPart!.Numbering!.Append(
                new NumberingInstance(new AbstractNumId { Val = 5 }) { NumberID = 1 });
            main.NumberingDefinitionsPart.Numbering.Save();
            var numberingProperties = new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = 1 });
            if (tracked)
            {
                numberingProperties.Append(new Inserted
                {
                    Id = "1",
                    Author = "Comparison Engine",
                    Date = FixedRevisionDate(),
                });
            }
            main.Document.Body = new Body(new Paragraph(
                new ParagraphProperties(numberingProperties),
                RunForText("Item")));
            document.Save();
        }
        return stream.ToArray();
    }

    private static string NestedCellMergeBody(string outerCellProperty) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>"
        + "<w:tcPr>" + outerCellProperty + "</w:tcPr><w:p><w:r><w:t>Outer</w:t>"
        + "</w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>"
        + "<w:tr><w:tc><w:tcPr><w:cellMerge w:id=\"2\""
        + " w:author=\"Comparison Engine\" w:date=\"2000-01-01T00:00:00Z\""
        + " w:vMerge=\"cont\" w:vMergeOrig=\"rest\"/></w:tcPr>"
        + "<w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        + "<w:p/></w:tc></w:tr></w:tbl>";

    private static string[] BookmarkNames(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document.Descendants<BookmarkStart>()
            .Select(bookmark => bookmark.Name?.Value)
            .Where(name => name is not null)
            .Select(name => name!)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
    }

    private static byte[] RewriteBody(byte[] source, params Paragraph[] paragraphs)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            document.MainDocumentPart!.Document.Body = new Body(paragraphs);
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] RewriteBodyXml(byte[] source, string bodyInnerXml)
        => RewriteBodyXmlWithNamespace(source, IrTestDocuments.W, bodyInnerXml);

    private static byte[] RewriteBodyXmlWithNamespace(
        byte[] source,
        string wordNamespace,
        string bodyInnerXml)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var xml = $"<w:document xmlns:w=\"{wordNamespace}\">"
                + $"<w:body>{bodyInnerXml}</w:body></w:document>";
            using var partStream = document.MainDocumentPart!.GetStream(
                FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(partStream);
            writer.Write(xml);
        }
        return stream.ToArray();
    }

    private static byte[] RewriteSettingsXml(byte[] source, string settingsXml)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            using var partStream = document.MainDocumentPart!.DocumentSettingsPart!.GetStream(
                FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(partStream);
            writer.Write(settingsXml);
        }
        return stream.ToArray();
    }

    private static byte[] AddOrphanHyperlink(
        byte[] source, string relationshipId, string target)
    {
        using var stream = new MemoryStream();
        stream.Write(source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            document.MainDocumentPart!.AddHyperlinkRelationship(
                new Uri(target, UriKind.Absolute), true, relationshipId);
        }
        return stream.ToArray();
    }
}
