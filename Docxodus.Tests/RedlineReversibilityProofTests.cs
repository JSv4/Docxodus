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

    private static byte[] Document(string text)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(text)))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            document.Save();
        }
        return stream.ToArray();
    }
}
