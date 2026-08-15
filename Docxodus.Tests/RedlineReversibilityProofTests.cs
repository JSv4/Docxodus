// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Verification;
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
        Assert.False(run.Proof.AcceptToFinal?.ModeledSemantic.Available);
        Assert.False(run.Proof.Success); // #457 is intentionally required before this can prove success.
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
    }

    [Theory]
    [InlineData("RP015-MoveFrom-MoveTo", RevisionFamily.Move)]
    [InlineData("RP025-Paragraph-Props-Change", RevisionFamily.PropertiesChange)]
    [InlineData("RP009-Deleted-Table-Row", RevisionFamily.RowDelete)]
    [InlineData("RP050-Deleted-Footnote", RevisionFamily.ContentDelete)]
    public void RP007_RealRevisionFamilies_ResolveAndEmitProofEvidence(
        string fixtureStem,
        RevisionFamily expectedFamily)
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
}
