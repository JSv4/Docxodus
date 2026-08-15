// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
    [InlineData("RP015-MoveFrom-MoveTo", RevisionFamily.Move)]
    [InlineData("RP021-Inserted-Numbering-Properties", RevisionFamily.NumberingPropertiesInsert)]
    [InlineData("RP025-Paragraph-Props-Change", RevisionFamily.PropertiesChange)]
    [InlineData("RP027-Change-Section", RevisionFamily.PropertiesChange)]
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
        AssertPathEvidenceCoherent(run.Proof.AcceptToFinal!, run.Proof.ToJson());
        AssertPathEvidenceCoherent(run.Proof.RejectToBaseline!, run.Proof.ToJson());
        Assert.False(run.Proof.AcceptToFinal!.Equivalent);
        Assert.False(run.Proof.RejectToBaseline!.Equivalent);
        Assert.False(run.Proof.Success);
        Assert.Contains(run.Proof.AcceptToFinal.Findings,
            finding => finding.Code == "normalized_whole_package_mismatch");
        Assert.Contains(run.Proof.RejectToBaseline.Findings,
            finding => finding.Code == "normalized_whole_package_mismatch");
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
                string.Equals(revision.PartUri, firstSemanticChange.PartUri,
                    StringComparison.Ordinal)
                && (string.Equals(revision.AnchorId, semanticAnchor, StringComparison.Ordinal)
                    || revision.AffectedAnchorIds.Contains(
                        semanticAnchor!, StringComparer.Ordinal)))
            .Select(revision => revision.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        Assert.Empty(expectedApplicableIds);
        Assert.Empty(finding.RevisionIds);
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
}
