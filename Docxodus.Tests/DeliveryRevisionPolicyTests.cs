// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Delivery;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DeliveryRevisionPolicyTests
{
    public static TheoryData<DeliveryRevisionPolicy, DeliveryRevisionPolicy> PolicyMatrix => new()
    {
        { DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Preserve },
        { DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Accept },
        { DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Reject },
        { DeliveryRevisionPolicy.Accept, DeliveryRevisionPolicy.Preserve },
        { DeliveryRevisionPolicy.Accept, DeliveryRevisionPolicy.Accept },
        { DeliveryRevisionPolicy.Accept, DeliveryRevisionPolicy.Reject },
        { DeliveryRevisionPolicy.Reject, DeliveryRevisionPolicy.Preserve },
        { DeliveryRevisionPolicy.Reject, DeliveryRevisionPolicy.Accept },
        { DeliveryRevisionPolicy.Reject, DeliveryRevisionPolicy.Reject },
    };

    [Theory]
    [MemberData(nameof(PolicyMatrix))]
    public void DRP001_PolicyMatrix_ResolvesEachRevisionClassOnIsolatedClones(
        DeliveryRevisionPolicy preExisting,
        DeliveryRevisionPolicy generated)
    {
        var fixture = TwoRevisionClasses();
        var baselineSnapshot = (byte[])fixture.Baseline.Clone();
        var editedSnapshot = (byte[])fixture.Edited.Clone();

        var result = DeliveryRevisionPolicyProcessor.Apply(
            fixture.Baseline,
            fixture.Edited,
            Policy(preExisting, generated),
            requireReviewProof: false);

        Assert.Equal(baselineSnapshot, fixture.Baseline);
        Assert.Equal(editedSnapshot, fixture.Edited);
        Assert.Null(result.ReviewBytes);
        Assert.Null(result.ReviewProof);
        Assert.Contains(result.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.PreExisting);
        Assert.Contains(result.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Generated);
        Assert.DoesNotContain(result.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Conflicted);

        AssertRevisionAuthors(
            result.PolicyBaselineBytes,
            preservePrior: preExisting == DeliveryRevisionPolicy.Preserve,
            preserveGenerated: false);
        AssertRevisionAuthors(
            result.FinalBytes,
            preservePrior: preExisting == DeliveryRevisionPolicy.Preserve,
            preserveGenerated: generated == DeliveryRevisionPolicy.Preserve);
        AssertResolvedViews(
            result.PolicyBaselineBytes,
            preExisting,
            generated: DeliveryRevisionPolicy.Reject,
            includeGeneratedParagraph: false);
        AssertResolvedViews(result.FinalBytes, preExisting, generated,
            includeGeneratedParagraph: true);

        if (preExisting == DeliveryRevisionPolicy.Accept
            && generated == DeliveryRevisionPolicy.Accept)
        {
            WriteArtifact("matrix-accept-accept", "baseline.docx", fixture.Baseline);
            WriteArtifact("matrix-accept-accept", "edited.docx", fixture.Edited);
            WriteArtifact("matrix-accept-accept", "policy-baseline.docx",
                result.PolicyBaselineBytes);
            WriteArtifact("matrix-accept-accept", "final.docx", result.FinalBytes);
        }
    }

    [Fact]
    public void DRP002_UntrackedEdits_SurviveRejectingGeneratedRevisions()
    {
        var source = AppendParagraph(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            "Third paragraph.");
        var baseline = TrackedReplace(source, paragraph: 0, "Prior proposed.", "Prior Reviewer");
        var edited = TrackedReplace(baseline, paragraph: 1, "Generated proposed.", "Delivery Editor");
        edited = UntrackedReplace(edited, paragraph: 2, "Untracked retained.");

        var result = DeliveryRevisionPolicyProcessor.Apply(
            baseline,
            edited,
            Policy(DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Reject),
            requireReviewProof: false);

        var acceptedView = FullyResolve(result.FinalBytes, accept: true);
        var markdown = Markdown(acceptedView);
        Assert.Contains("Second paragraph.", markdown, StringComparison.Ordinal);
        Assert.Contains("Untracked retained.", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Generated proposed.", markdown, StringComparison.Ordinal);
        AssertRevisionAuthors(result.FinalBytes, preservePrior: true, preserveGenerated: false);
    }

    [Fact]
    public void DRP003_ProofRequest_ReturnsNativeReviewAndSuccessfulSelectiveProof()
    {
        var fixture = TwoRevisionClasses();
        var baselineSnapshot = (byte[])fixture.Baseline.Clone();
        var editedSnapshot = (byte[])fixture.Edited.Clone();

        var result = DeliveryRevisionPolicyProcessor.Apply(
            fixture.Baseline,
            fixture.Edited,
            Policy(DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Accept),
            requireReviewProof: true);

        Assert.Equal(baselineSnapshot, fixture.Baseline);
        Assert.Equal(editedSnapshot, fixture.Edited);
        Assert.NotNull(result.ReviewBytes);
        Assert.NotNull(result.ReviewProof);
        Assert.True(result.ReviewProof.Proof.Success, result.ReviewProof.Proof.ToJson());
        Assert.True(result.ReviewProof.Proof.AcceptToFinal?.Equivalent,
            result.ReviewProof.Proof.ToJson());
        Assert.True(result.ReviewProof.Proof.RejectToBaseline?.Equivalent,
            result.ReviewProof.Proof.ToJson());
        Assert.Contains(result.ReviewProof.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.PreExisting);
        Assert.Contains(result.ReviewProof.Proof.RevisionClassifications, item =>
            item.Disposition == RedlineRevisionDisposition.Generated);
        AssertRevisionAuthors(result.ReviewBytes!, preservePrior: true, preserveGenerated: true);
        using (var review = new DocxSession(result.ReviewBytes!, new DocxSessionSettings
               {
                   EmitMarkdownPatch = false,
                   CaptureInitialProjection = false,
               }))
        {
            Assert.Contains(review.ListRevisions(), item =>
                item.Author == "Delivery Editor");
        }
        WriteArtifact("proof-identity", "policy-baseline.docx", result.PolicyBaselineBytes);
        WriteArtifact("proof-identity", "final.docx", result.FinalBytes);
        WriteArtifact("proof-identity", "review.docx", result.ReviewBytes!);
        WriteArtifact("proof-identity", "reversibility-proof.json",
            System.Text.Encoding.UTF8.GetBytes(result.ReviewProof.Proof.ToJson()));
    }

    [Theory]
    [InlineData(DeliveryRevisionPolicy.Preserve)]
    [InlineData(DeliveryRevisionPolicy.Reject)]
    public void DRP004_ProofRequest_RequiresGeneratedAcceptAndDoesNotMutateInputs(
        DeliveryRevisionPolicy generated)
    {
        var fixture = TwoRevisionClasses();
        var baselineSnapshot = (byte[])fixture.Baseline.Clone();
        var editedSnapshot = (byte[])fixture.Edited.Clone();

        Assert.Throws<ArgumentException>(() => DeliveryRevisionPolicyProcessor.Apply(
            fixture.Baseline,
            fixture.Edited,
            Policy(DeliveryRevisionPolicy.Preserve, generated),
            requireReviewProof: true));

        Assert.Equal(baselineSnapshot, fixture.Baseline);
        Assert.Equal(editedSnapshot, fixture.Edited);
    }

    [Fact]
    public void DRP005_FullIdentityConflict_FailsClosedInsteadOfTrustingAuthorOrId()
    {
        var fixture = TwoRevisionClasses();
        var tampered = RewriteRevisionAuthor(fixture.Edited, "Prior Reviewer", "Impostor");
        var baselineSnapshot = (byte[])fixture.Baseline.Clone();
        var editedSnapshot = (byte[])tampered.Clone();

        var error = Assert.Throws<InvalidDataException>(() =>
            DeliveryRevisionPolicyProcessor.Apply(
                fixture.Baseline,
                tampered,
                Policy(DeliveryRevisionPolicy.Accept, DeliveryRevisionPolicy.Accept),
                requireReviewProof: false));

        Assert.Contains("revision identity", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(baselineSnapshot, fixture.Baseline);
        Assert.Equal(editedSnapshot, tampered);
    }

    [Fact]
    public void DRP006_InvalidOrOverLimitInput_FailsBoundedPreflightWithoutMutation()
    {
        var valid = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var validSnapshot = (byte[])valid.Clone();
        var malformed = new byte[] { 1, 2, 3, 4 };
        var malformedSnapshot = (byte[])malformed.Clone();
        var options = new PackageManifestOptions { MaxEntryCount = 1 };

        Assert.Throws<InvalidDataException>(() => DeliveryRevisionPolicyProcessor.Apply(
            valid,
            malformed,
            Policy(DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Preserve),
            requireReviewProof: false));
        Assert.Throws<InvalidDataException>(() => DeliveryRevisionPolicyProcessor.Apply(
            valid,
            valid,
            Policy(DeliveryRevisionPolicy.Preserve, DeliveryRevisionPolicy.Preserve),
            requireReviewProof: false,
            options));

        Assert.Equal(validSnapshot, valid);
        Assert.Equal(malformedSnapshot, malformed);
    }

    private static RevisionFixture TwoRevisionClasses()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var baseline = TrackedReplace(source, paragraph: 0, "Prior proposed.", "Prior Reviewer");
        var edited = TrackedReplace(
            baseline, paragraph: 1, "Generated proposed.", "Delivery Editor");
        return new RevisionFixture(baseline, edited);
    }

    private static byte[] TrackedReplace(
        byte[] input,
        int paragraph,
        string replacement,
        string author)
    {
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = author,
            PersistAnchorIds = false,
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        var edit = session.ReplaceText(BodyParagraphs(session)[paragraph], replacement);
        Assert.True(edit.Success, edit.Error?.Message);
        return session.Save(persistAnchorIds: false);
    }

    private static byte[] UntrackedReplace(byte[] input, int paragraph, string replacement)
    {
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.Accept,
            PersistAnchorIds = false,
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        var edit = session.ReplaceText(BodyParagraphs(session)[paragraph], replacement);
        Assert.True(edit.Success, edit.Error?.Message);
        return session.Save(persistAnchorIds: false);
    }

    private static byte[] FullyResolve(byte[] input, bool accept)
    {
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            PersistAnchorIds = false,
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        for (var attempt = 0; attempt < 100; attempt++)
        {
            var revision = session.ListRevisions().FirstOrDefault();
            if (revision is null)
                return session.Save(persistAnchorIds: false);
            var edit = accept
                ? session.AcceptRevision(revision.Id)
                : session.RejectRevision(revision.Id);
            Assert.True(edit.Success, edit.Error?.Message);
        }
        throw new InvalidOperationException("Test revision resolution did not converge.");
    }

    private static void AssertResolvedViews(
        byte[] bytes,
        DeliveryRevisionPolicy preExisting,
        DeliveryRevisionPolicy generated,
        bool includeGeneratedParagraph)
    {
        var acceptedMarkdown = Markdown(FullyResolve(bytes, accept: true));
        var rejectedMarkdown = Markdown(FullyResolve(bytes, accept: false));

        Assert.Contains(
            preExisting == DeliveryRevisionPolicy.Reject
                ? "First paragraph."
                : "Prior proposed.",
            acceptedMarkdown,
            StringComparison.Ordinal);
        Assert.Contains(
            preExisting == DeliveryRevisionPolicy.Accept
                ? "Prior proposed."
                : "First paragraph.",
            rejectedMarkdown,
            StringComparison.Ordinal);
        if (!includeGeneratedParagraph)
            return;

        Assert.Contains(
            generated == DeliveryRevisionPolicy.Reject
                ? "Second paragraph."
                : "Generated proposed.",
            acceptedMarkdown,
            StringComparison.Ordinal);
        Assert.Contains(
            generated == DeliveryRevisionPolicy.Accept
                ? "Generated proposed."
                : "Second paragraph.",
            rejectedMarkdown,
            StringComparison.Ordinal);
    }

    private static void AssertRevisionAuthors(
        byte[] bytes,
        bool preservePrior,
        bool preserveGenerated)
    {
        using var session = new DocxSession(bytes, new DocxSessionSettings
        {
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        var authors = session.ListRevisions().Select(item => item.Author).ToArray();
        Assert.Equal(preservePrior, authors.Contains("Prior Reviewer", StringComparer.Ordinal));
        Assert.Equal(preserveGenerated, authors.Contains("Delivery Editor", StringComparer.Ordinal));
    }

    private static string Markdown(byte[] bytes)
    {
        using var session = new DocxSession(bytes, new DocxSessionSettings
        {
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        return session.Project().Markdown;
    }

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:", StringComparison.Ordinal))
            .ToArray();

    private static byte[] AppendParagraph(byte[] input, string text)
    {
        using var stream = new MemoryStream();
        stream.Write(input);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var mainDocument = document.MainDocumentPart!.Document
                ?? throw new InvalidDataException("Test document has no main document root.");
            mainDocument.Body!.AppendChild(
                new Paragraph(new Run(new Text(text))));
            mainDocument.Save();
        }
        return stream.ToArray();
    }

    private static byte[] RewriteRevisionAuthor(byte[] input, string oldAuthor, string newAuthor)
    {
        using var stream = new MemoryStream();
        stream.Write(input);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var root = document.MainDocumentPart!.GetXDocument().Root!;
            foreach (var revision in root.Descendants().Where(element =>
                         element.Attribute(W.author)?.Value == oldAuthor))
            {
                revision.SetAttributeValue(W.author, newAuthor);
            }
            document.MainDocumentPart.PutXDocument();
        }
        return stream.ToArray();
    }

    private static DeliveryBundleRevisionPolicy Policy(
        DeliveryRevisionPolicy preExisting,
        DeliveryRevisionPolicy generated) => new()
    {
        PreExistingRevisions = preExisting,
        GeneratedRevisions = generated,
    };

    private static void WriteArtifact(string group, string fileName, byte[] bytes)
    {
        var root = Environment.GetEnvironmentVariable("DOCXODUS_TEST_ARTIFACT_DIR");
        if (string.IsNullOrWhiteSpace(root))
            return;
        var directory = Path.Combine(root, group);
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, fileName), bytes);
    }

    private sealed record RevisionFixture(byte[] Baseline, byte[] Edited);
}
