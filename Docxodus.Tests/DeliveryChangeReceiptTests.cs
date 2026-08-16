// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DeliveryChangeReceiptTests
{
    [Fact]
    public void DCR001_SingleEdit_ProducesDeterministicVerifiableReceipt()
    {
        var edit = SingleEdit("Replacement clause.", tracked: true);
        var builder = Builder(edit);
        var entryId = AddTransactionWithSemantic(builder, edit);

        var first = builder.Build();
        var second = builder.Build();

        Assert.Equal(first.ReceiptDigest, second.ReceiptDigest);
        Assert.Equal(first.ToJson(), second.ToJson());
        Assert.Equal(DeliveryChangeReceiptPayload.SchemaId, first.Payload.Schema);
        var entry = Assert.Single(first.Payload.Transactions);
        Assert.Equal(entryId, entry.EntryId);
        Assert.Equal(DeliveryTransactionStatus.Committed, entry.Status);
        Assert.Single(entry.Operations);
        Assert.Contains(entry.AuthoredChanges,
            change => change.EntityKind == DeliveryAuthoredEntityKind.Revision
                && change.Author == "Receipt Author");
        Assert.True(DeliveryChangeReceiptVerifier.Verify(first,
            RequiredArtifactBytes(edit)).IsValid);
        Assert.True(DeliveryChangeReceiptVerifier.VerifyJson(first.ToJsonBytes(),
            RequiredArtifactBytes(edit)).IsValid);
    }

    [Fact]
    public void DCR002_AtomicBatch_RetainsNormalizedStepOrderAndOneVersionTransition()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = Open(source);
        var anchors = BodyParagraphs(session);
        var beforeBytes = session.Save();
        var beforeManifest = Manifest(beforeBytes);
        var operations = new[]
        {
            DeliveryNormalizedOperation.Create("docx_edit", "replace_text",
                JsonSerializer.Serialize(new { anchorId = anchors[0], markdown = "First changed." })),
            DeliveryNormalizedOperation.Create("docx_edit", "replace_text",
                JsonSerializer.Serialize(new { anchorId = anchors[1], markdown = "Second changed." })),
        };
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[0], "First changed.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchors[1], "Second changed.")),
        });
        var afterBytes = session.Save();
        var afterManifest = Manifest(afterBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, beforeManifest, afterManifest, operations);
        var builder = new DeliveryChangeReceiptBuilder(beforeManifest, result.BaseVersion)
            .SetDeliveredDocument(afterManifest, result.ResultVersion);
        AddCleanDocx(builder, afterBytes, afterManifest, result.ResultVersion);
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(beforeBytes, afterBytes)));
        var entryId = builder.AddTransaction(contribution);
        AddTransactionSemantic(
            builder, entryId, beforeBytes, afterBytes, "semantic-source-to-delivered");

        var receipt = builder.Build();

        var entry = Assert.Single(receipt.Payload.Transactions);
        Assert.Equal(result.BaseVersion + 1, result.ResultVersion);
        Assert.Equal(new[] { 0, 1 }, entry.Operations.Select(operation => operation.Index));
        Assert.Equal(new[] { "First changed.", "Second changed." },
            operations.Select(operation =>
                operation.Arguments.GetProperty("markdown").GetString()));
        Assert.Equal(DeliveryTransactionStatus.Committed, entry.Status);
    }

    [Fact]
    public void DCR003_RetryIdentity_DeduplicatesAndConflictingFingerprintFails()
    {
        var fingerprint = Fingerprint("retry request");
        var identity = new DeliveryTransactionIdentity
        {
            TransactionId = "delivery-42",
            RequestFingerprint = fingerprint,
        };
        var edit = SingleEdit("Retry-safe replacement.", identity: identity);
        var builder = Builder(edit);

        var firstId = AddTransactionWithSemantic(builder, edit);
        var retryId = builder.AddTransaction(edit.Contribution);

        Assert.Equal(firstId, retryId);
        Assert.Single(builder.Build().Payload.Transactions);

        var conflict = SingleEdit("Retry-safe replacement.", identity: identity with
        {
            RequestFingerprint = Fingerprint("different request"),
        });
        var error = Assert.Throws<DeliveryReceiptValidationException>(
            () => builder.AddTransaction(conflict.Contribution));
        Assert.Equal("transaction_conflict", error.Code);
    }

    [Fact]
    public void DCR004_UndoRedo_AreLineageEventsNotDuplicateTransactions()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = Open(source);
        var anchor = BodyParagraphs(session)[0];
        var sourceBytes = session.Save();
        var sourceManifest = Manifest(sourceBytes);
        var operation = DeliveryNormalizedOperation.Create("docx_edit", "replace_text",
            JsonSerializer.Serialize(new { anchorId = anchor, markdown = "Lineage edit." }));
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchor, "Lineage edit.")),
        });
        var appliedBytes = session.Save();
        var appliedManifest = Manifest(appliedBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, sourceManifest, appliedManifest, new[] { operation });

        Assert.True(session.Undo());
        var undoVersion = session.Version;
        var undoBytes = session.Save();
        var undoManifest = Manifest(undoBytes);
        Assert.True(session.Redo());
        var redoVersion = session.Version;
        var redoBytes = session.Save();
        var redoManifest = Manifest(redoBytes);

        var builder = new DeliveryChangeReceiptBuilder(sourceManifest, result.BaseVersion);
        var entryId = builder.AddTransaction(contribution);
        builder.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Undo, entryId,
            appliedManifest, result.ResultVersion, undoManifest, undoVersion));
        builder.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Redo, entryId,
            undoManifest, undoVersion, redoManifest, redoVersion));
        builder.SetDeliveredDocument(redoManifest, redoVersion);
        AddCleanDocx(builder, redoBytes, redoManifest, redoVersion);
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(sourceBytes, redoBytes)));
        AddTransactionSemantic(
            builder, entryId, sourceBytes, appliedBytes, "semantic-transaction-1");

        var receipt = builder.Build();

        Assert.Single(receipt.Payload.Transactions);
        Assert.Equal(new[] { DeliveryLineageAction.Undo, DeliveryLineageAction.Redo },
            receipt.Payload.Lineage.Select(value => value.Action));
        Assert.All(receipt.Payload.Lineage,
            value => Assert.Equal(entryId, value.AffectedEntryId));
        Assert.Equal(redoVersion, receipt.Payload.DeliveredDocument.DocumentVersion);
    }

    [Fact]
    public void DCR012_EditAfterUndo_RetainsOneCrossStreamChronology()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = Open(source);
        var firstAnchor = BodyParagraphs(session)[0];
        var sourceBytes = session.Save();
        var sourceManifest = Manifest(sourceBytes);
        var firstOperation = DeliveryNormalizedOperation.Create("docx_edit", "replace_text",
            JsonSerializer.Serialize(new { anchorId = firstAnchor, markdown = "First edit." }));
        var firstResult = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(firstAnchor, "First edit.")),
        });
        var firstBytes = session.Save();
        var firstManifest = Manifest(firstBytes);
        var firstContribution = DeliveryTransactionContribution.FromMutationBatchResult(
            firstResult, sourceManifest, firstManifest, new[] { firstOperation });

        var builder = new DeliveryChangeReceiptBuilder(sourceManifest, firstResult.BaseVersion);
        var firstEntryId = builder.AddTransaction(firstContribution);
        Assert.True(session.Undo());
        var undoVersion = session.Version;
        var undoBytes = session.Save();
        var undoManifest = Manifest(undoBytes);
        builder.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Undo, firstEntryId,
            firstManifest, firstResult.ResultVersion, undoManifest, undoVersion));

        var secondAnchor = BodyParagraphs(session)[1];
        var secondOperation = DeliveryNormalizedOperation.Create("docx_edit", "replace_text",
            JsonSerializer.Serialize(new { anchorId = secondAnchor, markdown = "Second edit." }));
        var secondResult = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(secondAnchor, "Second edit.")),
        });
        var deliveredBytes = session.Save();
        var deliveredManifest = Manifest(deliveredBytes);
        var secondEntryId = builder.AddTransaction(
            DeliveryTransactionContribution.FromMutationBatchResult(
            secondResult, undoManifest, deliveredManifest, new[] { secondOperation }));
        builder.SetDeliveredDocument(deliveredManifest, secondResult.ResultVersion);
        AddCleanDocx(
            builder, deliveredBytes, deliveredManifest, secondResult.ResultVersion);
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(sourceBytes, deliveredBytes)));
        AddTransactionSemantic(
            builder, firstEntryId, sourceBytes, firstBytes, "semantic-transaction-1");
        AddTransactionSemantic(
            builder, secondEntryId, undoBytes, deliveredBytes, "semantic-transaction-2");

        var receipt = builder.Build();
        var chronology = receipt.Payload.Transactions
            .Select(transaction => (transaction.Sequence, Kind: "transaction"))
            .Concat(receipt.Payload.Lineage
                .Select(lineageEvent => (lineageEvent.Sequence, Kind: "undo")))
            .OrderBy(value => value.Sequence)
            .ToArray();

        Assert.Equal(new[]
        {
            (0L, "transaction"),
            (1L, "undo"),
            (2L, "transaction"),
        }, chronology);
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            receipt, new Dictionary<string, byte[]>
            {
                ["clean-docx"] = deliveredBytes,
                ["semantic-source-to-delivered"] =
                    SemanticChanges(sourceBytes, deliveredBytes).ToCanonicalUtf8Bytes(),
                ["semantic-transaction-1"] =
                    SemanticChanges(sourceBytes, firstBytes).ToCanonicalUtf8Bytes(),
                ["semantic-transaction-2"] =
                    SemanticChanges(undoBytes, deliveredBytes).ToCanonicalUtf8Bytes(),
            }).IsValid);
    }

    [Fact]
    public void DCR005_PrivacyProfiles_RedactBeforeCanonicalization()
    {
        const string secret = "CONFIDENTIAL-CUSTOMER-TERM-9381";
        var edit = SingleEdit(secret, tracked: true);

        var hashOnly = BuildWithProfile(edit, DeliveryReceiptPrivacyProfile.HashOnly);
        var summary = BuildWithProfile(edit, DeliveryReceiptPrivacyProfile.HashAndSummary);
        var full = BuildWithProfile(edit, DeliveryReceiptPrivacyProfile.FullEvidence);

        Assert.DoesNotContain(secret, hashOnly.ToJson(), StringComparison.Ordinal);
        Assert.DoesNotContain(secret, summary.ToJson(), StringComparison.Ordinal);
        Assert.Contains(secret, full.ToJson(), StringComparison.Ordinal);
        Assert.NotEqual(hashOnly.ReceiptDigest, summary.ReceiptDigest);
        Assert.NotEqual(summary.ReceiptDigest, full.ReceiptDigest);
        Assert.Null(hashOnly.Payload.Transactions[0].Operations[0].ArgumentsSummary);
        Assert.NotNull(summary.Payload.Transactions[0].Operations[0].ArgumentsSummary);
        Assert.NotNull(full.Payload.Transactions[0].Operations[0].Arguments);
    }

    [Fact]
    public void DCR006_Attribution_ReportsRequestedDerivedAndUnexpectedWithoutHidingChanges()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = Open(source);
        var anchor = BodyParagraphs(session)[0];
        var beforeBytes = session.Save();
        var beforeManifest = Manifest(beforeBytes);
        var operation = DeliveryNormalizedOperation.Create("docx_link", "add",
            JsonSerializer.Serialize(new
            {
                anchorId = anchor,
                start = 0,
                length = 5,
                target = "https://example.test/receipt",
            }));
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_link", "add", s => s.AddHyperlink(
                anchor, new CharSpan(0, 5),
                new HyperlinkTarget(HyperlinkKind.External, "https://example.test/receipt"))),
        });
        Assert.True(result.Success);
        var afterBytes = session.Save();
        var afterManifest = Manifest(afterBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, beforeManifest, afterManifest, new[] { operation });
        var addedRelationship = afterManifest.Relationships.Single(relationship =>
            !beforeManifest.Relationships.Any(before =>
                before.OwnerUri == relationship.OwnerUri && before.Id == relationship.Id));
        var builder = new DeliveryChangeReceiptBuilder(beforeManifest, result.BaseVersion)
            .SetDeliveredDocument(afterManifest, result.ResultVersion);
        var entryId = builder.AddTransaction(contribution);
        AddCleanDocx(builder, afterBytes, afterManifest, result.ResultVersion);
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(beforeBytes, afterBytes)));
        AddTransactionSemantic(
            builder, entryId, beforeBytes, afterBytes, "semantic-source-to-delivered");
        builder.AddAttributionRule(new DeliveryChangeAttributionRule
        {
            Kind = DeliveryPackageChangeKind.PartModified,
            EntryUri = "/word/document.xml",
            Disposition = DeliveryChangeDisposition.UserRequested,
            TransactionEntryId = entryId,
            RequestedOperationIndex = 0,
        });
        builder.AddAttributionRule(new DeliveryChangeAttributionRule
        {
            Kind = DeliveryPackageChangeKind.RelationshipAdded,
            OwnerUri = addedRelationship.OwnerUri,
            RelationshipId = addedRelationship.Id,
            Disposition = DeliveryChangeDisposition.Derived,
            TransactionEntryId = entryId,
            RequestedOperationIndex = 0,
            Derivation = "External hyperlink relationship required by requested hyperlink.",
        });

        var receipt = builder.Build();

        Assert.Contains(receipt.Payload.PackageChanges,
            change => change.Disposition == DeliveryChangeDisposition.UserRequested);
        Assert.Contains(receipt.Payload.PackageChanges,
            change => change.Disposition == DeliveryChangeDisposition.Derived);
        Assert.Contains(receipt.Payload.PackageChanges,
            change => change.Disposition == DeliveryChangeDisposition.Unexpected);
        Assert.True(receipt.Payload.HasUnexpectedChanges);
        builder.FailOnUnexpectedChanges = true;
        Assert.Equal("unexpected_package_change",
            Assert.Throws<DeliveryReceiptValidationException>(() => builder.Build()).Code);
    }

    [Fact]
    public void DCR007_MultiArtifactRedlineAndPageCitation_VerifyIndependently()
    {
        var edit = SingleEdit("Changed for redline delivery.");
        var redlineBytes = DocxDiff.Compare(
            new WmlDocument("before.docx", edit.BeforeBytes),
            new WmlDocument("after.docx", edit.AfterBytes)).DocumentByteArray;
        var redlineIdentity = DeliveryDocumentIdentity.FromManifest(
            Manifest(redlineBytes), edit.Result.ResultVersion);
        var deliveredIdentity = DeliveryDocumentIdentity.FromManifest(
            edit.AfterManifest, edit.Result.ResultVersion);
        var htmlBytes = Encoding.UTF8.GetBytes("<main>Changed for delivery.</main>");
        var pdfBytes = Encoding.ASCII.GetBytes("%PDF-1.7\nreceipt pagination fixture\n%%EOF");
        var validation = DeliverableVerifier.VerifyDeliverable(
            edit.BeforeBytes, edit.AfterBytes);
        var validationBytes = validation.ToCanonicalUtf8Bytes();
        var reversibility = RedlineReversibilityVerifier.Prove(
            edit.BeforeBytes, edit.AfterBytes, redlineBytes).Proof;
        var reversibilityBytes = Encoding.UTF8.GetBytes(reversibility.ToCanonicalJson());
        const string renderer = "chromium-140|fonts-v3|pagination-v1";
        var semanticBytes = SemanticChanges(
            edit.BeforeBytes, edit.AfterBytes).ToCanonicalUtf8Bytes();
        var pageMapBytes = PageMapBytes(
            edit.AnchorId, edit.Result.ResultVersion, renderer);
        var pageMapDigest = Digest(pageMapBytes);

        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "review-docx", DeliveryArtifactRole.ReviewDocx,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            redlineBytes) with
        {
            Document = redlineIdentity,
            RelativePath = "delivery/review.docx",
        });
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "delivery-html", DeliveryArtifactRole.Html, "text/html", htmlBytes));
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "delivery-page-map", DeliveryArtifactRole.PageMap,
            "application/json", pageMapBytes) with
        {
            Document = deliveredIdentity,
            RendererFingerprint = renderer,
        });
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "delivery-pdf", DeliveryArtifactRole.Pdf, "application/pdf", pdfBytes) with
        {
            Document = deliveredIdentity,
            RendererFingerprint = renderer,
            PageMapDigest = pageMapDigest,
            RelativePath = "delivery/document.pdf",
        });
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "validation", DeliveryArtifactRole.ValidationReport,
            "application/json", validationBytes));
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "reversibility", DeliveryArtifactRole.ReversibilityProof,
            "application/json", reversibilityBytes));
        builder.AddEvidence(new DeliveryEvidenceReference
        {
            Kind = DeliveryEvidenceKind.ValidationResult,
            Schema = DeliverableVerificationResult.SchemaId,
            Digest = Digest(validationBytes),
            ArtifactId = "validation",
            Summary = $"Deliverable verification decision: {validation.Decision}.",
        });
        builder.AddEvidence(new DeliveryEvidenceReference
        {
            Kind = DeliveryEvidenceKind.RedlineReversibility,
            Schema = RedlineReversibilityProof.SchemaId,
            Digest = Digest(reversibilityBytes),
            ArtifactId = "reversibility",
            Summary = $"Redline reversibility success: {reversibility.Success}.",
        });
        builder.AddPageCitation(new DeliveryPageCitationInput
        {
            Citation = Citation(edit.AnchorId, edit.Result.ResultVersion, renderer),
            Scope = "body",
            Document = deliveredIdentity,
            PageMapDigest = pageMapDigest,
            PageMapArtifactId = "delivery-page-map",
            RenderArtifactId = "delivery-pdf",
        });

        var receipt = builder.Build();
        var artifacts = new Dictionary<string, byte[]>
        {
            ["clean-docx"] = edit.AfterBytes,
            ["review-docx"] = redlineBytes,
            ["delivery-html"] = htmlBytes,
            ["delivery-page-map"] = pageMapBytes,
            ["delivery-pdf"] = pdfBytes,
            ["validation"] = validationBytes,
            ["reversibility"] = reversibilityBytes,
            ["semantic-source-to-delivered"] = semanticBytes,
        };
        var verification = DeliveryChangeReceiptVerifier.Verify(receipt, artifacts);

        Assert.True(verification.IsValid, string.Join(",", verification.Findings));
        Assert.Equal(artifacts.Keys.OrderBy(value => value, StringComparer.Ordinal),
            verification.Artifacts.Select(artifact => artifact.ArtifactId));
        Assert.All(verification.Artifacts,
            artifact => Assert.Equal(DeliveryArtifactVerificationStatus.Verified, artifact.Status));
        Assert.Single(receipt.Payload.PageCitations);
        Assert.Contains(receipt.Payload.SemanticChangeSets,
            evidence => evidence.Scope == DeliverySemanticComparisonScope.SourceToDelivered
                && evidence.Schema == SemanticChangeSet.CurrentSchema);
        Assert.Contains(receipt.Payload.Artifacts,
            artifact => artifact.Role == DeliveryArtifactRole.ReviewDocx);
        Assert.Contains(receipt.Payload.Evidence,
            evidence => evidence.Schema == DeliverableVerificationResult.SchemaId
                && evidence.Digest == Digest(validation.ToCanonicalUtf8Bytes()));
        Assert.Contains(receipt.Payload.Evidence,
            evidence => evidence.Schema == RedlineReversibilityProof.SchemaId
                && evidence.Digest == Digest(Encoding.UTF8.GetBytes(
                    reversibility.ToCanonicalJson())));
    }

    [Fact]
    public void DCR008_TamperedReceiptArtifactAndMissingArtifact_AreDetected()
    {
        var edit = SingleEdit("Tamper target.");
        var pdf = Encoding.ASCII.GetBytes("%PDF-1.7\noriginal\n%%EOF");
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "pdf", DeliveryArtifactRole.Pdf, "application/pdf", pdf));
        var receipt = builder.Build();
        var goodBytes = RequiredArtifactBytes(edit);
        goodBytes["pdf"] = pdf;

        var good = DeliveryChangeReceiptVerifier.Verify(receipt, goodBytes);
        Assert.True(good.IsValid);

        var tamperedArtifact = pdf.ToArray();
        tamperedArtifact[10] ^= 0x01;
        var tamperedBytes = RequiredArtifactBytes(edit);
        tamperedBytes["pdf"] = tamperedArtifact;
        var artifactResult = DeliveryChangeReceiptVerifier.Verify(receipt, tamperedBytes);
        Assert.False(artifactResult.IsValid);
        Assert.Equal(DeliveryArtifactVerificationStatus.DigestMismatch,
            Assert.Single(artifactResult.Artifacts,
                artifact => artifact.ArtifactId == "pdf").Status);

        var missing = DeliveryChangeReceiptVerifier.Verify(receipt,
            new Dictionary<string, byte[]>());
        Assert.False(missing.IsValid);
        Assert.Equal(DeliveryArtifactVerificationStatus.Missing,
            Assert.Single(missing.Artifacts,
                artifact => artifact.ArtifactId == "pdf").Status);

        var tamperedReceipt = receipt.ToJson().Replace(
            "hashAndSummary", "fullEvidence", StringComparison.Ordinal);
        var receiptResult = DeliveryChangeReceiptVerifier.VerifyJson(
            tamperedReceipt, goodBytes);
        Assert.False(receiptResult.IsValid);
        Assert.False(receiptResult.ReceiptDigestValid);
    }

    [Fact]
    public void DCR009_PageCitation_RejectsRendererPackageAndContinuousMismatches()
    {
        var edit = SingleEdit("Citation target.");
        var identity = DeliveryDocumentIdentity.FromManifest(
            edit.AfterManifest, edit.Result.ResultVersion);
        var pdf = Encoding.ASCII.GetBytes("%PDF-1.7\nfixture\n%%EOF");
        var pageMapBytes = PageMapBytes(
            edit.AnchorId, edit.Result.ResultVersion, "renderer-A");
        var pageMapDigest = Digest(pageMapBytes);
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "pdf", DeliveryArtifactRole.Pdf, "application/pdf", pdf) with
        {
            Document = identity,
            RendererFingerprint = "renderer-A",
            PageMapDigest = pageMapDigest,
        });
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "page-map", DeliveryArtifactRole.PageMap, "application/json", pageMapBytes) with
        {
            Document = identity,
            RendererFingerprint = "renderer-A",
        });
        builder.AddPageCitation(new DeliveryPageCitationInput
        {
            Citation = Citation(edit.AnchorId, edit.Result.ResultVersion, "renderer-B"),
            Scope = "body",
            Document = identity,
            PageMapDigest = pageMapDigest,
            PageMapArtifactId = "page-map",
            RenderArtifactId = "pdf",
        });

        Assert.Equal("citation_render_binding_mismatch",
            Assert.Throws<DeliveryReceiptValidationException>(() => builder.Build()).Code);

        var continuous = Builder(edit);
        AddTransactionWithSemantic(continuous, edit);
        continuous.AddArtifact(DeliveryArtifactInput.Available(
            "pdf", DeliveryArtifactRole.Pdf, "application/pdf", pdf) with
        {
            Document = identity,
            RendererFingerprint = "renderer-A",
            PageMapDigest = pageMapDigest,
        });
        continuous.AddArtifact(DeliveryArtifactInput.Available(
            "page-map", DeliveryArtifactRole.PageMap, "application/json", pageMapBytes) with
        {
            Document = identity,
            RendererFingerprint = "renderer-A",
        });
        continuous.AddPageCitation(new DeliveryPageCitationInput
        {
            Citation = Citation(edit.AnchorId, edit.Result.ResultVersion, "renderer-A") with
            {
                Availability = PageMapAvailability.Unavailable,
                UnavailableReason = PageCitationUnavailableReason.ContinuousMode,
                Pages = Array.Empty<PageMapPage>(),
                Fragments = Array.Empty<PageMapFragment>(),
            },
            Scope = "body",
            Document = identity,
            PageMapDigest = pageMapDigest,
            PageMapArtifactId = "page-map",
            RenderArtifactId = "pdf",
        });
        Assert.Equal("unavailable_page_citation",
            Assert.Throws<DeliveryReceiptValidationException>(() => continuous.Build()).Code);
    }

    [Fact]
    public void DCR010_OperationCanonicalization_SortsObjectsAndRejectsDuplicateKeys()
    {
        var left = DeliveryNormalizedOperation.Create("tool", "action", "{\"z\":1,\"a\":2}");
        var right = DeliveryNormalizedOperation.Create("tool", "action", "{ \"a\":2, \"z\":1 }");

        Assert.Equal(left.ArgumentsDigest, right.ArgumentsDigest);
        Assert.Equal("invalid_operation_arguments",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                DeliveryNormalizedOperation.Create("tool", "action", "{\"a\":1,\"a\":2}"))
                .Code);
    }

    [Fact]
    public void DCR011_UnavailableArtifact_IsPortableButCannotMasqueradeAsHashedOutput()
    {
        var edit = SingleEdit("Unavailable artifact.");
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        builder.AddArtifact(DeliveryArtifactInput.Unavailable(
            "pdf", DeliveryArtifactRole.Pdf, "application/pdf", "renderer unavailable") with
        {
            RelativePath = "delivery/document.pdf",
        });

        var receipt = builder.Build();
        var verification = DeliveryChangeReceiptVerifier.Verify(
            receipt, RequiredArtifactBytes(edit));

        Assert.True(verification.IsValid);
        var artifact = Assert.Single(receipt.Payload.Artifacts,
            value => value.ArtifactId == "pdf");
        Assert.Null(artifact.Digest);
        Assert.Null(artifact.ByteLength);
        Assert.Equal(DeliveryArtifactVerificationStatus.Unavailable,
            Assert.Single(verification.Artifacts,
                value => value.ArtifactId == "pdf").Status);
        Assert.Equal("unsafe_artifact_path",
            Assert.Throws<DeliveryReceiptValidationException>(() => Builder(edit).AddArtifact(
                DeliveryArtifactInput.Unavailable("bad", DeliveryArtifactRole.Pdf,
                    "application/pdf", "none") with { RelativePath = "../escape.pdf" })).Code);
    }

    [Fact]
    public void DCR013_RehashedMalformedContract_IsRejectedIndependentlyOfEnvelopeDigest()
    {
        var edit = SingleEdit("Malformed contract target.");
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        var original = builder.Build();
        var malformedPayload = original.Payload with
        {
            Transactions = new[]
            {
                original.Payload.Transactions[0] with
                {
                    Status = DeliveryTransactionStatus.Failed,
                },
            },
        };
        var malformed = new DeliveryChangeReceipt
        {
            Payload = malformedPayload,
            ReceiptDigest = Digest(
                DeliveryChangeReceiptSerializer.SerializePayload(malformedPayload)),
        };

        var verification = DeliveryChangeReceiptVerifier.Verify(
            malformed, RequiredArtifactBytes(edit));

        Assert.True(verification.ReceiptDigestValid);
        Assert.False(verification.ContractValid);
        Assert.False(verification.IsValid);
        Assert.Contains("failed_transaction_changed_document", verification.Findings);
    }

    [Fact]
    public void DCR014_PageCitation_RequiresReachableStateAndExactPageMapProjection()
    {
        const string renderer = "chromium-140|fonts-v3|pagination-v1";
        var edit = SingleEdit("Citation reachability target.");
        var deliveredIdentity = DeliveryDocumentIdentity.FromManifest(
            edit.AfterManifest, edit.Result.ResultVersion);
        var pdfBytes = Encoding.ASCII.GetBytes("%PDF-1.7\nreachable citation\n%%EOF");
        var pageMapBytes = PageMapBytes(
            edit.AnchorId, edit.Result.ResultVersion, renderer);

        var validBuilder = Builder(edit);
        AddTransactionWithSemantic(validBuilder, edit);
        AddPaginatedCitation(
            validBuilder,
            deliveredIdentity,
            Citation(edit.AnchorId, edit.Result.ResultVersion, renderer),
            pdfBytes,
            pageMapBytes);
        var validReceipt = validBuilder.Build();
        var originalCitation = Assert.Single(validReceipt.Payload.PageCitations);
        var forgedCitation = originalCitation with
        {
            Fragments = originalCitation.Fragments.Select(fragment => fragment with
            {
                Geometry = fragment.Geometry with { X = fragment.Geometry.X + 1 },
            }).ToArray(),
        };
        var forged = Rehash(validReceipt.Payload with
        {
            PageCitations = new[] { forgedCitation },
        });
        var artifacts = RequiredArtifactBytes(edit);
        artifacts["pdf"] = pdfBytes;
        artifacts["page-map"] = pageMapBytes;

        var verification = DeliveryChangeReceiptVerifier.Verify(forged, artifacts);

        Assert.True(verification.ReceiptDigestValid);
        Assert.False(verification.CitationBindingsValid);
        Assert.Contains(
            $"citation_page_map_projection_mismatch:{edit.AnchorId}",
            verification.Findings);

        var unreachableVersion = edit.Result.ResultVersion + 1;
        var unreachableIdentity = deliveredIdentity with
        {
            DocumentVersion = unreachableVersion,
        };
        var unreachablePageMap = PageMapBytes(
            edit.AnchorId, unreachableVersion, renderer);
        var unreachableBuilder = Builder(edit);
        AddTransactionWithSemantic(unreachableBuilder, edit);
        AddPaginatedCitation(
            unreachableBuilder,
            unreachableIdentity,
            Citation(edit.AnchorId, unreachableVersion, renderer),
            pdfBytes,
            unreachablePageMap);

        Assert.Equal("unreachable_citation_document",
            Assert.Throws<DeliveryReceiptValidationException>(
                () => unreachableBuilder.Build()).Code);

        var duplicatePropertyPageMap = Encoding.UTF8.GetBytes(
            Encoding.UTF8.GetString(pageMapBytes).Insert(1, "\"schemaVersion\":1,"));
        foreach (var invalidPageMap in new[]
                 {
                     duplicatePropertyPageMap,
                     new byte[] { 0xff, 0xfe, 0xfd },
                 })
        {
            var invalidMapBuilder = Builder(edit);
            AddTransactionWithSemantic(invalidMapBuilder, edit);
            AddPaginatedCitation(
                invalidMapBuilder,
                deliveredIdentity,
                Citation(edit.AnchorId, edit.Result.ResultVersion, renderer),
                pdfBytes,
                invalidPageMap);
            Assert.Equal("invalid_page_map_artifact",
                Assert.Throws<DeliveryReceiptValidationException>(
                    () => invalidMapBuilder.Build()).Code);
        }
    }

    [Fact]
    public void DCR015_CleanDocx_IsMandatoryAndExactlyMatchesDeliveredDocument()
    {
        var edit = SingleEdit("Clean artifact target.");
        Assert.Equal("invalid_package_manifest",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                new DeliveryChangeReceiptBuilder(
                    edit.BeforeManifest with { IsValid = false },
                    edit.Result.BaseVersion)).Code);
        var missing = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest, edit.Result.BaseVersion)
            .SetDeliveredDocument(edit.AfterManifest, edit.Result.ResultVersion);
        var missingEntryId = missing.AddTransaction(edit.Contribution);
        missing.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(edit.BeforeBytes, edit.AfterBytes)));
        AddTransactionSemantic(
            missing, missingEntryId, edit.BeforeBytes, edit.AfterBytes,
            "semantic-source-to-delivered");

        Assert.Equal("missing_clean_docx",
            Assert.Throws<DeliveryReceiptValidationException>(() => missing.Build()).Code);

        var validBuilder = Builder(edit);
        AddTransactionWithSemantic(validBuilder, edit);
        var valid = validBuilder.Build();
        var forgedArtifacts = valid.Payload.Artifacts.Select(artifact =>
            artifact.Role == DeliveryArtifactRole.CleanDocx
                ? artifact with
                {
                    DocumentVersion = edit.Result.ResultVersion + 1,
                }
                : artifact).ToArray();
        var forged = Rehash(valid.Payload with { Artifacts = forgedArtifacts });

        var verification = DeliveryChangeReceiptVerifier.Verify(
            forged, RequiredArtifactBytes(edit));

        Assert.True(verification.ReceiptDigestValid);
        Assert.False(verification.ContractValid);
        Assert.Contains("clean_docx_delivery_mismatch", verification.Findings);
    }

    [Fact]
    public void DCR016_Lineage_EnforcesLifoOrderVersionStepsAndNoOpRedoPreservation()
    {
        var edit = SingleEdit("Lineage invariant target.");

        (DeliveryChangeReceiptBuilder Builder, string EntryId) BareBuilder()
        {
            var builder = new DeliveryChangeReceiptBuilder(
                edit.BeforeManifest, edit.Result.BaseVersion);
            return (builder, builder.AddTransaction(edit.Contribution));
        }

        var (redoFirst, redoFirstEntryId) = BareBuilder();
        redoFirst.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Redo,
            redoFirstEntryId,
            edit.AfterManifest,
            edit.Result.ResultVersion,
            edit.AfterManifest,
            edit.Result.ResultVersion + 1));
        redoFirst.SetDeliveredDocument(
            edit.AfterManifest, edit.Result.ResultVersion + 1);
        Assert.Equal("invalid_redo_order",
            Assert.Throws<DeliveryReceiptValidationException>(() => redoFirst.Build()).Code);

        var (repeatedUndo, repeatedUndoEntryId) = BareBuilder();
        repeatedUndo.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Undo,
            repeatedUndoEntryId,
            edit.AfterManifest,
            edit.Result.ResultVersion,
            edit.BeforeManifest,
            edit.Result.ResultVersion + 1));
        repeatedUndo.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Undo,
            repeatedUndoEntryId,
            edit.BeforeManifest,
            edit.Result.ResultVersion + 1,
            edit.BeforeManifest,
            edit.Result.ResultVersion + 2));
        repeatedUndo.SetDeliveredDocument(
            edit.BeforeManifest, edit.Result.ResultVersion + 2);
        Assert.Equal("invalid_undo_order",
            Assert.Throws<DeliveryReceiptValidationException>(() => repeatedUndo.Build()).Code);

        var (skippedVersion, skippedVersionEntryId) = BareBuilder();
        skippedVersion.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Undo,
            skippedVersionEntryId,
            edit.AfterManifest,
            edit.Result.ResultVersion,
            edit.BeforeManifest,
            edit.Result.ResultVersion + 2));
        skippedVersion.SetDeliveredDocument(
            edit.BeforeManifest, edit.Result.ResultVersion + 2);
        Assert.Equal("invalid_lineage_version",
            Assert.Throws<DeliveryReceiptValidationException>(() => skippedVersion.Build()).Code);

        using (var session = Open(edit.BeforeBytes))
        {
            var anchors = BodyParagraphs(session);
            var sourceBytes = session.Save();
            var sourceManifest = Manifest(sourceBytes);
            var firstOperation = DeliveryNormalizedOperation.Create(
                "docx_edit", "replace_text",
                JsonSerializer.Serialize(new
                {
                    anchorId = anchors[0],
                    markdown = "First stacked edit.",
                }));
            var firstResult = session.ExecuteBatch(new[]
            {
                new MutationBatchStep("docx_edit", "replace_text",
                    s => s.ReplaceText(anchors[0], "First stacked edit.")),
            });
            var firstBytes = session.Save();
            var firstManifest = Manifest(firstBytes);
            var secondOperation = DeliveryNormalizedOperation.Create(
                "docx_edit", "replace_text",
                JsonSerializer.Serialize(new
                {
                    anchorId = anchors[1],
                    markdown = "Second stacked edit.",
                }));
            var secondResult = session.ExecuteBatch(new[]
            {
                new MutationBatchStep("docx_edit", "replace_text",
                    s => s.ReplaceText(anchors[1], "Second stacked edit.")),
            });
            var secondBytes = session.Save();
            var secondManifest = Manifest(secondBytes);
            var nonLifo = new DeliveryChangeReceiptBuilder(
                sourceManifest, firstResult.BaseVersion);
            var firstEntryId = nonLifo.AddTransaction(
                DeliveryTransactionContribution.FromMutationBatchResult(
                    firstResult,
                    sourceManifest,
                    firstManifest,
                    new[] { firstOperation }));
            _ = nonLifo.AddTransaction(
                DeliveryTransactionContribution.FromMutationBatchResult(
                    secondResult,
                    firstManifest,
                    secondManifest,
                    new[] { secondOperation }));
            nonLifo.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
                DeliveryLineageAction.Undo,
                firstEntryId,
                secondManifest,
                secondResult.ResultVersion,
                sourceManifest,
                secondResult.ResultVersion + 1));
            nonLifo.SetDeliveredDocument(
                sourceManifest, secondResult.ResultVersion + 1);

            Assert.Equal("invalid_undo_order",
                Assert.Throws<DeliveryReceiptValidationException>(() => nonLifo.Build()).Code);
        }

        var (noOpHistory, noOpHistoryEntryId) = BareBuilder();
        noOpHistory.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Undo,
            noOpHistoryEntryId,
            edit.AfterManifest,
            edit.Result.ResultVersion,
            edit.BeforeManifest,
            edit.Result.ResultVersion + 1));
        var noOpResult = new MutationBatchResult
        {
            Mode = MutationBatchMode.Atomic,
            Success = true,
            RolledBack = false,
            BaseVersion = edit.Result.ResultVersion + 1,
            ResultVersion = edit.Result.ResultVersion + 1,
        };
        _ = noOpHistory.AddTransaction(
            DeliveryTransactionContribution.FromMutationBatchResult(
                noOpResult,
                edit.BeforeManifest,
                edit.BeforeManifest,
                Array.Empty<DeliveryNormalizedOperation>()));
        noOpHistory.AddLineageEvent(DeliveryLineageEventInput.FromManifests(
            DeliveryLineageAction.Redo,
            noOpHistoryEntryId,
            edit.BeforeManifest,
            edit.Result.ResultVersion + 1,
            edit.AfterManifest,
            edit.Result.ResultVersion + 2));
        noOpHistory.SetDeliveredDocument(
            edit.AfterManifest, edit.Result.ResultVersion + 2);
        AddCleanDocx(
            noOpHistory,
            edit.AfterBytes,
            edit.AfterManifest,
            edit.Result.ResultVersion + 2);
        noOpHistory.AddSemanticChangeSet(
            DeliverySemanticChangeSetInput.ForSourceToDelivered(
                SemanticChanges(edit.BeforeBytes, edit.AfterBytes)));
        AddTransactionSemantic(
            noOpHistory,
            noOpHistoryEntryId,
            edit.BeforeBytes,
            edit.AfterBytes,
            "semantic-source-to-delivered");
        var validNoOpHistory = noOpHistory.Build();
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            validNoOpHistory, RequiredArtifactBytes(edit)).IsValid);

        var forgedLineage = validNoOpHistory.Payload.Lineage
            .Select((lineageEvent, index) => index == 0
                ? lineageEvent with
                {
                    AfterDocument = lineageEvent.AfterDocument with
                    {
                        DocumentVersion = lineageEvent.AfterDocument.DocumentVersion + 1,
                    },
                }
                : lineageEvent)
            .ToArray();
        var forged = Rehash(validNoOpHistory.Payload with { Lineage = forgedLineage });
        var verification = DeliveryChangeReceiptVerifier.Verify(
            forged, RequiredArtifactBytes(edit));
        Assert.True(verification.ReceiptDigestValid);
        Assert.Contains("invalid_lineage_version", verification.Findings);
    }

    [Fact]
    public void DCR017_Attribution_RejectsFailedOrRolledBackRequestedOperations()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = Open(source);
        var anchor = BodyParagraphs(session)[0];
        var beforeBytes = session.Save();
        var beforeManifest = Manifest(beforeBytes);
        var operations = new[]
        {
            DeliveryNormalizedOperation.Create(
                "docx_edit", "replace_text",
                JsonSerializer.Serialize(new
                {
                    anchorId = anchor,
                    markdown = "Retained best-effort edit.",
                })),
            DeliveryNormalizedOperation.Create(
                "docx_edit", "replace_text",
                JsonSerializer.Serialize(new
                {
                    anchorId = "p:body:missing",
                    markdown = "Must fail.",
                })),
        };
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchor, "Retained best-effort edit.")),
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText("p:body:missing", "Must fail.")),
        }, MutationBatchMode.BestEffort);
        Assert.False(result.Success);
        Assert.False(result.Steps[1].Success);
        var afterBytes = session.Save();
        var afterManifest = Manifest(afterBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, beforeManifest, afterManifest, operations);

        DeliveryChangeReceiptBuilder BuilderWithAttribution(
            int operationIndex,
            DeliveryChangeDisposition disposition = DeliveryChangeDisposition.UserRequested)
        {
            var builder = new DeliveryChangeReceiptBuilder(
                beforeManifest, result.BaseVersion)
                .SetDeliveredDocument(afterManifest, result.ResultVersion);
            var entryId = builder.AddTransaction(contribution);
            AddCleanDocx(builder, afterBytes, afterManifest, result.ResultVersion);
            builder.AddSemanticChangeSet(
                DeliverySemanticChangeSetInput.ForSourceToDelivered(
                    SemanticChanges(beforeBytes, afterBytes)));
            AddTransactionSemantic(
                builder,
                entryId,
                beforeBytes,
                afterBytes,
                "semantic-source-to-delivered");
            builder.AddAttributionRule(new DeliveryChangeAttributionRule
            {
                Kind = DeliveryPackageChangeKind.PartModified,
                EntryUri = "/word/document.xml",
                Disposition = disposition,
                TransactionEntryId = entryId,
                RequestedOperationIndex = operationIndex,
                Derivation = disposition == DeliveryChangeDisposition.Derived
                    ? "Derived from the retained request."
                    : null,
            });
            return builder;
        }

        Assert.Equal("invalid_attribution_operation",
            Assert.Throws<DeliveryReceiptValidationException>(
                () => BuilderWithAttribution(1).Build()).Code);
        Assert.Equal("invalid_attribution_operation",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                BuilderWithAttribution(
                    1, DeliveryChangeDisposition.Derived).Build()).Code);

        var valid = BuilderWithAttribution(0).Build();
        Assert.Contains(valid.Payload.PackageChanges,
            change => change.Disposition == DeliveryChangeDisposition.UserRequested);
        var forgedChanges = valid.Payload.PackageChanges.Select(change =>
            change.Disposition == DeliveryChangeDisposition.UserRequested
                ? change with { RequestedOperationIndex = 1 }
                : change).ToArray();
        var forged = Rehash(valid.Payload with { PackageChanges = forgedChanges });
        var artifacts = new Dictionary<string, byte[]>
        {
            ["clean-docx"] = afterBytes,
            ["semantic-source-to-delivered"] =
                SemanticChanges(beforeBytes, afterBytes).ToCanonicalUtf8Bytes(),
        };

        var verification = DeliveryChangeReceiptVerifier.Verify(forged, artifacts);

        Assert.True(verification.ReceiptDigestValid);
        Assert.Contains("invalid_package_change_attribution", verification.Findings);
    }

    [Fact]
    public void DCR018_SemanticChangeSets_RequireTypedExactAggregateAndTransactionCoverage()
    {
        var edit = SingleEdit("Semantic coverage target.");

        var missingAggregate = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest, edit.Result.BaseVersion)
            .SetDeliveredDocument(edit.AfterManifest, edit.Result.ResultVersion);
        var missingAggregateEntryId = missingAggregate.AddTransaction(edit.Contribution);
        AddCleanDocx(
            missingAggregate,
            edit.AfterBytes,
            edit.AfterManifest,
            edit.Result.ResultVersion);
        AddTransactionSemantic(
            missingAggregate,
            missingAggregateEntryId,
            edit.BeforeBytes,
            edit.AfterBytes,
            "semantic-transaction-1");
        Assert.Equal("missing_source_to_delivered_semantic_evidence",
            Assert.Throws<DeliveryReceiptValidationException>(
                () => missingAggregate.Build()).Code);

        var missingTransaction = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest, edit.Result.BaseVersion)
            .SetDeliveredDocument(edit.AfterManifest, edit.Result.ResultVersion);
        _ = missingTransaction.AddTransaction(edit.Contribution);
        AddCleanDocx(
            missingTransaction,
            edit.AfterBytes,
            edit.AfterManifest,
            edit.Result.ResultVersion);
        missingTransaction.AddSemanticChangeSet(
            DeliverySemanticChangeSetInput.ForSourceToDelivered(
                SemanticChanges(edit.BeforeBytes, edit.AfterBytes)));
        Assert.Equal("semantic_transaction_coverage_mismatch",
            Assert.Throws<DeliveryReceiptValidationException>(
                () => missingTransaction.Build()).Code);

        var genericEvidence = Builder(edit);
        Assert.Equal("semantic_evidence_requires_typed_factory",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                genericEvidence.AddEvidence(new DeliveryEvidenceReference
                {
                    Kind = DeliveryEvidenceKind.SemanticChangeSet,
                    Schema = SemanticChangeSet.CurrentSchema,
                    Digest = Digest(Encoding.UTF8.GetBytes("not semantic evidence")),
                })).Code);

        var validBuilder = Builder(edit);
        AddTransactionWithSemantic(validBuilder, edit);
        var valid = validBuilder.Build();
        var forgedBindings = valid.Payload.SemanticChangeSets.Select(binding =>
            binding.Scope == DeliverySemanticComparisonScope.SourceToDelivered
                ? binding with
                {
                    BeforeDocument = valid.Payload.DeliveredDocument,
                    Schema = "https://docxodus.dev/schemas/semantic-change-set/forged",
                    SchemaVersion = SemanticChangeSet.CurrentSchemaVersion + 1,
                }
                : binding).ToArray();
        var forged = Rehash(valid.Payload with { SemanticChangeSets = forgedBindings });

        var forgedVerification = DeliveryChangeReceiptVerifier.Verify(
            forged, RequiredArtifactBytes(edit));

        Assert.True(forgedVerification.ReceiptDigestValid);
        Assert.Contains("semantic_binding_identity_mismatch", forgedVerification.Findings);
        Assert.Contains("unsupported_semantic_change_set", forgedVerification.Findings);
        Assert.Contains(
            "semantic_artifact_binding_mismatch:semantic-source-to-delivered",
            forgedVerification.Findings);

        var tamperedArtifacts = RequiredArtifactBytes(edit);
        tamperedArtifacts["semantic-source-to-delivered"][^1] ^= 0x01;
        var byteVerification = DeliveryChangeReceiptVerifier.Verify(valid, tamperedArtifacts);
        Assert.Contains(
            "semantic_artifact_binding_mismatch:semantic-source-to-delivered",
            byteVerification.Findings);
    }

    [Fact]
    public void DCR019_SemanticArtifacts_MustBeCompleteTypedCanonicalBytes()
    {
        var edit = SingleEdit("Strict semantic bytes.");
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        var valid = builder.Build();
        var canonical = RequiredArtifactBytes(edit)["semantic-source-to-delivered"];

        DeliveryReceiptVerificationResult VerifyVariant(byte[] bytes, int changeCount)
        {
            var digest = Digest(bytes);
            var artifacts = valid.Payload.Artifacts.Select(artifact =>
                artifact.ArtifactId == "semantic-source-to-delivered"
                    ? artifact with { Digest = digest, ByteLength = bytes.LongLength }
                    : artifact).ToArray();
            var bindings = valid.Payload.SemanticChangeSets.Select(binding =>
                binding.ArtifactId == "semantic-source-to-delivered"
                    ? binding with { Digest = digest, ChangeCount = changeCount }
                    : binding).ToArray();
            var forged = Rehash(valid.Payload with
            {
                Artifacts = artifacts,
                SemanticChangeSets = bindings,
            });
            var supplied = RequiredArtifactBytes(edit);
            supplied["semantic-source-to-delivered"] = bytes;
            return DeliveryChangeReceiptVerifier.Verify(forged, supplied);
        }

        var whitespaceVariant = canonical.Concat(new byte[] { (byte)'\n' }).ToArray();
        var whitespaceVerification = VerifyVariant(
            whitespaceVariant, valid.Payload.SemanticChangeSets[0].ChangeCount);
        Assert.True(whitespaceVerification.ReceiptDigestValid);
        Assert.Contains("invalid_semantic_change_set", whitespaceVerification.Findings);
        Assert.Contains(
            "semantic_artifact_binding_mismatch:semantic-source-to-delivered",
            whitespaceVerification.Findings);

        var nullChange = Encoding.UTF8.GetBytes(
            $"{{\"schema\":\"{SemanticChangeSet.CurrentSchema}\","
            + $"\"schemaVersion\":{SemanticChangeSet.CurrentSchemaVersion},"
            + "\"changeCount\":1,\"changes\":[null]}");
        var nullVerification = VerifyVariant(nullChange, changeCount: 1);
        Assert.True(nullVerification.ReceiptDigestValid);
        Assert.Contains("invalid_semantic_change_set", nullVerification.Findings);
        Assert.Contains(
            "semantic_artifact_binding_mismatch:semantic-source-to-delivered",
            nullVerification.Findings);
    }

    [Fact]
    public void DCR020_CleanDocx_IdentityIsRecomputedFromExactWordPackageBytes()
    {
        var sourceBytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var manifest = Manifest(sourceBytes);
        const long version = 0;
        var arbitrary = Encoding.UTF8.GetBytes("not a DOCX package");
        var identity = DeliveryDocumentIdentity.FromManifest(manifest, version);

        var rejectingBuilder = new DeliveryChangeReceiptBuilder(manifest, version)
            .SetDeliveredDocument(manifest, version);
        var rejected = Assert.Throws<DeliveryReceiptValidationException>(() =>
            rejectingBuilder.AddArtifact(DeliveryArtifactInput.Available(
                "clean-docx",
                DeliveryArtifactRole.CleanDocx,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                arbitrary) with { Document = identity }));
        Assert.Equal("invalid_clean_docx", rejected.Code);

        var builder = new DeliveryChangeReceiptBuilder(manifest, version)
            .SetDeliveredDocument(manifest, version);
        AddCleanDocx(builder, sourceBytes, manifest, version);
        var emptySemantic = new SemanticChangeSet(Array.Empty<SemanticChange>());
        builder.AddSemanticChangeSet(
            DeliverySemanticChangeSetInput.ForSourceToDelivered(emptySemantic));
        var valid = builder.Build();

        var forgedDigest = Digest(arbitrary);
        var forgedIdentity = valid.Payload.DeliveredDocument with
        {
            RawPackageBytesDigest = forgedDigest,
        };
        var artifacts = valid.Payload.Artifacts.Select(artifact =>
            artifact.Role == DeliveryArtifactRole.CleanDocx
                ? artifact with
                {
                    ByteLength = arbitrary.LongLength,
                    Digest = forgedDigest,
                    PackageDigest = forgedDigest,
                }
                : artifact).ToArray();
        var semanticBindings = valid.Payload.SemanticChangeSets.Select(binding => binding with
        {
            BeforeDocument = forgedIdentity,
            AfterDocument = forgedIdentity,
        }).ToArray();
        var forged = Rehash(valid.Payload with
        {
            SourceDocument = forgedIdentity,
            DeliveredDocument = forgedIdentity,
            Artifacts = artifacts,
            SemanticChangeSets = semanticBindings,
        });
        var supplied = new Dictionary<string, byte[]>
        {
            ["clean-docx"] = arbitrary,
            ["semantic-source-to-delivered"] = emptySemantic.ToCanonicalUtf8Bytes(),
        };

        var verification = DeliveryChangeReceiptVerifier.Verify(forged, supplied);

        Assert.True(verification.ReceiptDigestValid);
        Assert.Contains("clean_docx_delivery_mismatch", verification.Findings);
        Assert.False(verification.IsValid);
    }

    [Fact]
    public void DCR021_VerifierRejectsDerivedFailedAttributionAndContradictoryResults()
    {
        var edit = SingleEdit("Attribution invariant.");
        var builder = Builder(edit);
        var entryId = AddTransactionWithSemantic(builder, edit);
        builder.AddAttributionRule(new DeliveryChangeAttributionRule
        {
            Kind = DeliveryPackageChangeKind.PartModified,
            EntryUri = "/word/document.xml",
            Disposition = DeliveryChangeDisposition.Derived,
            TransactionEntryId = entryId,
            RequestedOperationIndex = 0,
            Derivation = "Derived package representation.",
        });
        var valid = builder.Build();
        Assert.Contains(valid.Payload.PackageChanges,
            change => change.Disposition == DeliveryChangeDisposition.Derived);

        var transaction = Assert.Single(valid.Payload.Transactions);
        var operation = Assert.Single(transaction.Operations);
        var failedAttributionTransaction = transaction with
        {
            Operations = new[]
            {
                operation with
                {
                    ExecutionStatus = DeliveryOperationExecutionStatus.Failed,
                    Success = false,
                },
            },
        };
        var failedAttribution = Rehash(valid.Payload with
        {
            Transactions = new[] { failedAttributionTransaction },
        });
        var attributionVerification = DeliveryChangeReceiptVerifier.Verify(
            failedAttribution, RequiredArtifactBytes(edit));
        Assert.Contains(
            "invalid_package_change_attribution", attributionVerification.Findings);
        Assert.Contains("transaction_outcome_mismatch", attributionVerification.Findings);

        var result = Assert.Single(operation.Results);
        var contradictoryTransaction = transaction with
        {
            Operations = new[]
            {
                operation with
                {
                    Results = new[] { result with { Success = false } },
                },
            },
        };
        var contradictory = Rehash(valid.Payload with
        {
            Transactions = new[] { contradictoryTransaction },
        });
        var resultVerification = DeliveryChangeReceiptVerifier.Verify(
            contradictory, RequiredArtifactBytes(edit));
        Assert.Contains("operation_success_mismatch", resultVerification.Findings);
        Assert.Contains("invalid_operation_result", resultVerification.Findings);
    }

    [Fact]
    public void DCR022_FailedAtomicBatch_RetainsFullRequestWithExplicitAbsentOperations()
    {
        (DeliveryChangeReceipt Receipt, Dictionary<string, byte[]> Artifacts) BuildFailure(
            bool preflightFailure)
        {
            var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
            using var session = Open(source);
            var anchors = BodyParagraphs(session);
            var beforeBytes = session.Save();
            var beforeManifest = Manifest(beforeBytes);
            var operations = new[]
            {
                DeliveryNormalizedOperation.Create("docx_edit", "first", "{}"),
                DeliveryNormalizedOperation.Create("docx_edit", "failure", "{}"),
                DeliveryNormalizedOperation.Create("docx_edit", "later", "{}"),
            };
            var steps = new[]
            {
                new MutationBatchStep("docx_edit", "first",
                    s => s.ReplaceText(anchors[0], "Temporary atomic edit.")),
                new MutationBatchStep("docx_edit", "failure",
                    s => s.ReplaceText("p:body:missing", "Must fail."),
                    preflightFailure
                        ? _ => new EditError(EditErrorCode.AnchorNotFound, "preflight failure")
                        : null),
                new MutationBatchStep("docx_edit", "later",
                    s => s.ReplaceText(anchors[1], "Must not execute.")),
            };
            var result = session.ExecuteBatch(steps, MutationBatchMode.Atomic);
            Assert.False(result.Success);
            var afterBytes = session.Save();
            var afterManifest = Manifest(afterBytes);
            var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
                result, beforeManifest, afterManifest, operations);
            var builder = new DeliveryChangeReceiptBuilder(
                beforeManifest, result.BaseVersion)
                .SetDeliveredDocument(afterManifest, result.ResultVersion);
            _ = builder.AddTransaction(contribution);
            AddCleanDocx(builder, afterBytes, afterManifest, result.ResultVersion);
            var semantic = SemanticChanges(beforeBytes, afterBytes);
            builder.AddSemanticChangeSet(
                DeliverySemanticChangeSetInput.ForSourceToDelivered(semantic));
            return (builder.Build(), new Dictionary<string, byte[]>
            {
                ["clean-docx"] = afterBytes,
                ["semantic-source-to-delivered"] = semantic.ToCanonicalUtf8Bytes(),
            });
        }

        var executed = BuildFailure(preflightFailure: false);
        Assert.Equal(new[]
        {
            DeliveryOperationExecutionStatus.SucceededRolledBack,
            DeliveryOperationExecutionStatus.FailedRolledBack,
            DeliveryOperationExecutionStatus.NotExecuted,
        }, executed.Receipt.Payload.Transactions[0].Operations
            .Select(operation => operation.ExecutionStatus));
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            executed.Receipt, executed.Artifacts).IsValid);

        var preflight = BuildFailure(preflightFailure: true);
        Assert.Equal(new[]
        {
            DeliveryOperationExecutionStatus.NotExecuted,
            DeliveryOperationExecutionStatus.FailedRolledBack,
            DeliveryOperationExecutionStatus.NotExecuted,
        }, preflight.Receipt.Payload.Transactions[0].Operations
            .Select(operation => operation.ExecutionStatus));
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            preflight.Receipt, preflight.Artifacts).IsValid);
    }

    [Fact]
    public void DCR023_ResourceLimitsRejectBeforeUnboundedArtifactOrJsonProcessing()
    {
        var defaults = new DeliveryReceiptLimits();
        Assert.Equal(16 * 1024 * 1024, defaults.MaxReceiptJsonBytes);
        Assert.Equal(64 * 1024 * 1024, defaults.MaxSemanticEvidenceBytes);
        Assert.Equal(64 * 1024 * 1024, defaults.MaxPageMapBytes);
        Assert.Equal(256 * 1024 * 1024, defaults.MaxArtifactBytes);
        Assert.Equal(512L * 1024 * 1024, defaults.MaxTotalArtifactBytes);
        Assert.Equal(128, defaults.MaxJsonDepth);
        Assert.Equal(100_000, defaults.MaxCollectionItems);
        Assert.Equal(10_000, defaults.MaxTransactions);
        Assert.Equal(10_000, defaults.MaxOperationsPerTransaction);
        Assert.Equal(1_024, defaults.MaxArtifacts);
        Assert.Equal(1024 * 1024, defaults.MaxStringLength);

        var edit = SingleEdit("Resource limits.");
        var smallArtifactBuilder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest,
            edit.Result.BaseVersion,
            limits: new DeliveryReceiptLimits { MaxArtifactBytes = 8 })
            .SetDeliveredDocument(edit.AfterManifest, edit.Result.ResultVersion);
        Assert.Equal("artifact_resource_limit",
            Assert.Throws<DeliveryReceiptValidationException>(() => AddCleanDocx(
                smallArtifactBuilder,
                edit.AfterBytes,
                edit.AfterManifest,
                edit.Result.ResultVersion)).Code);

        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        var valid = builder.Build();
        var artifacts = RequiredArtifactBytes(edit);
        var receiptLimit = DeliveryChangeReceiptVerifier.VerifyJson(
            valid.ToJsonBytes(), artifacts, new DeliveryReceiptVerificationOptions
            {
                Limits = new DeliveryReceiptLimits { MaxReceiptJsonBytes = 128 },
            });
        Assert.Contains("receipt_resource_limit", receiptLimit.Findings);

        var semanticLimit = DeliveryChangeReceiptVerifier.Verify(
            valid, artifacts, new DeliveryReceiptVerificationOptions
            {
                Limits = new DeliveryReceiptLimits { MaxSemanticEvidenceBytes = 1 },
            });
        Assert.Contains("semantic_resource_limit", semanticLimit.Findings);

        var pageMapBytes = Encoding.UTF8.GetBytes("{}");
        var pageBuilder = Builder(edit);
        AddTransactionWithSemantic(pageBuilder, edit);
        pageBuilder.AddArtifact(DeliveryArtifactInput.Available(
            "bounded-page-map", DeliveryArtifactRole.PageMap,
            "application/json", pageMapBytes));
        var pageReceipt = pageBuilder.Build();
        var pageArtifacts = RequiredArtifactBytes(edit);
        pageArtifacts["bounded-page-map"] = pageMapBytes;
        var pageLimit = DeliveryChangeReceiptVerifier.Verify(
            pageReceipt, pageArtifacts, new DeliveryReceiptVerificationOptions
            {
                Limits = new DeliveryReceiptLimits { MaxPageMapBytes = 1 },
            });
        Assert.Contains("page_map_resource_limit", pageLimit.Findings);

        long suppliedBytes = artifacts.Values.Sum(value => value.LongLength);
        var totalLimit = DeliveryChangeReceiptVerifier.Verify(
            valid, artifacts, new DeliveryReceiptVerificationOptions
            {
                Limits = new DeliveryReceiptLimits
                {
                    MaxTotalArtifactBytes = suppliedBytes - 1,
                },
            });
        Assert.Contains("artifact_resource_limit", totalLimit.Findings);
    }

    [Fact]
    public void DCR024_RehashedArrayReorderingAndDuplicatesAreRejected()
    {
        var edit = SingleEdit("Canonical collection ordering.", tracked: true);
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        builder.AddWarning("first receipt warning");
        builder.AddWarning("second receipt warning");
        builder.AddEvidence(new DeliveryEvidenceReference
        {
            Kind = DeliveryEvidenceKind.ValidationResult,
            Schema = "https://example.test/validation/v1",
            Digest = Digest(Encoding.UTF8.GetBytes("validation")),
            Summary = "validation evidence",
        });
        builder.AddEvidence(new DeliveryEvidenceReference
        {
            Kind = DeliveryEvidenceKind.RedlineReversibility,
            Schema = "https://example.test/reversibility/v1",
            Digest = Digest(Encoding.UTF8.GetBytes("reversibility")),
            Summary = "reversibility evidence",
        });
        var valid = builder.Build();
        Assert.True(valid.Payload.Artifacts.Count >= 2);
        Assert.True(valid.Payload.PackageChanges.Count >= 1);
        Assert.Equal(2, valid.Payload.Evidence.Count);
        Assert.Equal(2, valid.Payload.Warnings.Count);

        var transaction = Assert.Single(valid.Payload.Transactions);
        var authored = Assert.IsType<DeliveryAuthoredChange>(
            transaction.AuthoredChanges.First());
        var operation = Assert.Single(transaction.Operations);
        var operationResult = Assert.Single(operation.Results);
        var objectChange = Assert.Single(operationResult.ObjectChanges);
        var duplicateAuthored = authored with
        {
            AffectedAnchorIds = new[]
            {
                "p:body:duplicate",
                "p:body:duplicate",
            },
        };
        var duplicateResult = operationResult with
        {
            ObjectChanges = new[] { objectChange, objectChange },
        };
        var forgedTransaction = transaction with
        {
            Operations = new[]
            {
                operation with { Results = new[] { duplicateResult } },
            },
            AuthoredChanges = new[] { duplicateAuthored, duplicateAuthored },
        };
        var citation = new DeliveryPageCitation
        {
            AnchorId = edit.AnchorId,
            Scope = "body",
            DocumentVersion = valid.Payload.DeliveredDocument.DocumentVersion,
            PackageDigest = valid.Payload.DeliveredDocument.RawPackageBytesDigest,
            RendererFingerprint = "forged-renderer",
            PageMapDigest = Digest(Encoding.UTF8.GetBytes("map")),
            PageMapArtifactId = "missing-page-map",
            RenderArtifactId = "missing-render",
            RenderArtifactDigest = Digest(Encoding.UTF8.GetBytes("render")),
        };
        var firstPackageChange = valid.Payload.PackageChanges[0];
        var duplicateLineage = new DeliveryLineageEvent
        {
            Sequence = 1,
            Action = DeliveryLineageAction.Undo,
            AffectedEntryId = transaction.EntryId,
            BeforeDocument = valid.Payload.DeliveredDocument,
            AfterDocument = valid.Payload.DeliveredDocument,
        };
        var forged = Rehash(valid.Payload with
        {
            Transactions = new[] { forgedTransaction, forgedTransaction },
            Lineage = new[] { duplicateLineage, duplicateLineage },
            PackageChanges = new[] { firstPackageChange, firstPackageChange },
            Artifacts = valid.Payload.Artifacts.Reverse().ToArray(),
            PageCitations = new[] { citation, citation },
            Evidence = valid.Payload.Evidence.Reverse().ToArray(),
            Warnings = valid.Payload.Warnings.Reverse().ToArray(),
        });

        var verification = DeliveryChangeReceiptVerifier.Verify(
            forged, RequiredArtifactBytes(edit));

        Assert.True(verification.ReceiptDigestValid);
        Assert.Contains("transaction_order_mismatch", verification.Findings);
        Assert.Contains("lineage_order_mismatch", verification.Findings);
        Assert.Contains("artifact_order_mismatch", verification.Findings);
        Assert.Contains("package_change_order_mismatch", verification.Findings);
        Assert.Contains("citation_order_mismatch", verification.Findings);
        Assert.Contains("evidence_order_mismatch", verification.Findings);
        Assert.Contains("warning_order_mismatch", verification.Findings);
        Assert.Contains("authored_change_order_mismatch", verification.Findings);
        Assert.Contains("affected_anchor_order_mismatch", verification.Findings);
        Assert.Contains("object_change_order_mismatch", verification.Findings);
        Assert.False(verification.IsValid);
    }

    [Fact]
    public void DCR025_CleanDocxInventory_IsTheAuthoritativeDeliveredManifest()
    {
        var edit = SingleEdit("Authoritative clean package inventory.");
        var expectedBuilder = Builder(edit);
        AddTransactionWithSemantic(expectedBuilder, edit);
        var expected = expectedBuilder.Build();
        var forgedManifest = edit.AfterManifest with
        {
            Entries = Array.Empty<PackageManifestEntry>(),
            Relationships = Array.Empty<PackageRelationship>(),
        };
        var actualBuilder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest, edit.Result.BaseVersion)
            .SetDeliveredDocument(forgedManifest, edit.Result.ResultVersion);
        AddCleanDocx(
            actualBuilder, edit.AfterBytes, forgedManifest, edit.Result.ResultVersion);
        actualBuilder.AddSemanticChangeSet(
            DeliverySemanticChangeSetInput.ForSourceToDelivered(
                SemanticChanges(edit.BeforeBytes, edit.AfterBytes)));
        AddTransactionWithSemantic(actualBuilder, edit);

        var actual = actualBuilder.Build();

        Assert.NotEmpty(actual.Payload.PackageChanges);
        Assert.Equal(expected.Payload.PackageChanges, actual.Payload.PackageChanges);
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            actual, RequiredArtifactBytes(edit)).IsValid);
    }

    [Fact]
    public void DCR026_ObjectAndTypedSemanticResourceLimits_FailClosedBeforeOutput()
    {
        var edit = SingleEdit("Bound object serialization.");
        var validBuilder = Builder(edit);
        AddTransactionWithSemantic(validBuilder, edit);
        var valid = validBuilder.Build();
        var artifacts = RequiredArtifactBytes(edit);
        var objectLimit = DeliveryChangeReceiptVerifier.Verify(
            valid, artifacts, new DeliveryReceiptVerificationOptions
            {
                Limits = new DeliveryReceiptLimits { MaxReceiptJsonBytes = 512 },
            });
        Assert.Contains("receipt_resource_limit", objectLimit.Findings);

        var bounded = Assert.Throws<DeliveryReceiptValidationException>(() =>
            DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
                new { Value = new string('x', 4_096) },
                new DeliveryReceiptLimits { MaxReceiptJsonBytes = 128 },
                128,
                "receipt_resource_limit"));
        Assert.Equal("receipt_resource_limit", bounded.Code);

        var originalOperation = Assert.Single(edit.Contribution.Operations);
        var oversizedResult = edit.Result with
        {
            Steps = new[]
            {
                new MutationBatchStepResult(
                    0,
                    originalOperation.Tool,
                    originalOperation.Action,
                    new[]
                    {
                        new EditResult
                        {
                            Success = true,
                            Patch = new MarkdownPatch(
                                edit.AnchorId, new string('z', 80 * 1024)),
                        },
                    },
                    false),
            },
            Failure = null,
        };
        var oversizedContribution =
            DeliveryTransactionContribution.FromMutationBatchResult(
                oversizedResult,
                edit.BeforeManifest,
                edit.AfterManifest,
                new[] { originalOperation });
        var resultBuilder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest,
            edit.Result.BaseVersion,
            limits: new DeliveryReceiptLimits { MaxReceiptJsonBytes = 64 * 1024 });
        Assert.Equal("receipt_resource_limit",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                resultBuilder.AddTransaction(oversizedContribution)).Code);

        SemanticValue deepValue = SemanticValue.String("leaf");
        for (int i = 0; i < 4; i++)
            deepValue = SemanticValue.Array(new[] { deepValue });
        var deepSet = SemanticSet(deepValue);
        var deepBuilder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest,
            edit.Result.BaseVersion,
            limits: new DeliveryReceiptLimits { MaxJsonDepth = 8 });
        Assert.Equal("semantic_resource_limit",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                deepBuilder.AddSemanticChangeSet(
                    DeliverySemanticChangeSetInput.ForSourceToDelivered(deepSet))).Code);

        var aggregateSet = SemanticSet(SemanticValue.String(new string('y', 256)));
        var aggregateBuilder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest,
            edit.Result.BaseVersion,
            limits: new DeliveryReceiptLimits { MaxSemanticEvidenceBytes = 512 });
        Assert.Equal("semantic_resource_limit",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                aggregateBuilder.AddSemanticChangeSet(
                    DeliverySemanticChangeSetInput.ForSourceToDelivered(aggregateSet))).Code);
    }

    [Fact]
    public void DCR027_ActualEmptySuccessfulStep_ProducesPortableNoOpEvidence()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = Open(source);
        var beforeBytes = session.Save();
        var beforeManifest = Manifest(beforeBytes);
        var operation = DeliveryNormalizedOperation.Create("docx_edit", "no_op", "{}");
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docx_edit", "no_op", _ => Array.Empty<EditResult>()),
        });
        var afterBytes = session.Save();
        var afterManifest = Manifest(afterBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, beforeManifest, afterManifest, new[] { operation });
        var builder = new DeliveryChangeReceiptBuilder(beforeManifest, result.BaseVersion)
            .SetDeliveredDocument(afterManifest, result.ResultVersion);
        AddCleanDocx(builder, afterBytes, afterManifest, result.ResultVersion);
        var semantic = SemanticChanges(beforeBytes, afterBytes);
        builder.AddSemanticChangeSet(
            DeliverySemanticChangeSetInput.ForSourceToDelivered(semantic));
        builder.AddTransaction(contribution);

        var receipt = builder.Build();
        var transaction = Assert.Single(receipt.Payload.Transactions);
        var operationEvidence = Assert.Single(transaction.Operations);
        var supplied = new Dictionary<string, byte[]>
        {
            ["clean-docx"] = afterBytes,
            ["semantic-source-to-delivered"] = semantic.ToCanonicalUtf8Bytes(),
        };

        Assert.True(result.Success);
        Assert.Equal(result.BaseVersion, result.ResultVersion);
        Assert.Equal(DeliveryTransactionStatus.Committed, transaction.Status);
        Assert.Equal(DeliveryOperationExecutionStatus.Succeeded,
            operationEvidence.ExecutionStatus);
        Assert.Empty(operationEvidence.Results);
        Assert.True(DeliveryChangeReceiptVerifier.Verify(receipt, supplied).IsValid);
        Assert.True(DeliveryChangeReceiptVerifier.VerifyJson(
            receipt.ToJsonBytes(), supplied).IsValid);
    }

    [Fact]
    public void DCR028_ArtifactPaths_AreNormalizedAndRootFormsArePortable()
    {
        var edit = SingleEdit("Portable artifact paths.");
        var builder = Builder(edit);
        AddTransactionWithSemantic(builder, edit);
        builder.AddArtifact(DeliveryArtifactInput.Unavailable(
            "pdf", DeliveryArtifactRole.Pdf, "application/pdf", "renderer unavailable") with
        {
            RelativePath = @"delivery\document.pdf",
        });

        var receipt = builder.Build();

        Assert.Equal("delivery/document.pdf", Assert.Single(
            receipt.Payload.Artifacts, artifact => artifact.ArtifactId == "pdf").RelativePath);
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            receipt, RequiredArtifactBytes(edit)).IsValid);

        var invalidBuilder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest, edit.Result.BaseVersion);
        foreach (var path in new[]
        {
            "/tmp/document.pdf",
            @"\\server\share\document.pdf",
            @"\rooted\document.pdf",
            @"C:\temp\document.pdf",
            @"C:relative\document.pdf",
            "../escape.pdf",
            "delivery//document.pdf",
        })
        {
            Assert.Equal("unsafe_artifact_path",
                Assert.Throws<DeliveryReceiptValidationException>(() =>
                    invalidBuilder.AddArtifact(DeliveryArtifactInput.Unavailable(
                        "bad", DeliveryArtifactRole.Pdf, "application/pdf", "none") with
                    {
                        RelativePath = path,
                    })).Code);
        }
    }

    [Fact]
    public void DCR029_CollectionLimits_AreChargedBeforeJsonItemsAreRetained()
    {
        var receiptLimits = new DeliveryReceiptLimits
        {
            MaxCollectionItems = 3,
            MaxReceiptJsonBytes = 4 * 1024,
        };
        foreach (var json in new[]
        {
            "[0,0,0,0]",
            "{\"a\":0,\"b\":0,\"c\":0,\"d\":0}",
        })
        {
            Assert.Equal("receipt_resource_limit",
                Assert.Throws<DeliveryReceiptValidationException>(() =>
                    DeliveryReceiptCanonicalJson.CanonicalizeBounded(
                        Encoding.UTF8.GetBytes(json),
                        receiptLimits,
                        receiptLimits.MaxReceiptJsonBytes,
                        "receipt_resource_limit")).Code);
        }

        var array = SemanticValue.Array(Enumerable.Repeat(SemanticValue.Absent, 8));
        var semanticBytes = SemanticSet(array).ToCanonicalUtf8Bytes();
        var semanticLimits = new DeliveryReceiptLimits
        {
            MaxCollectionItems = 20,
            MaxSemanticEvidenceBytes = 64 * 1024,
        };
        Assert.Equal("semantic_resource_limit",
            Assert.Throws<DeliveryReceiptValidationException>(() =>
                DeliverySemanticChangeSetAdapter.InspectExact(
                    semanticBytes, semanticLimits)).Code);
    }

    [Fact]
    public void DCR030_PackageChanges_ProjectTheSharedPackageDeltaExactly()
    {
        var edit = SingleEdit("Shared package delta projection.");
        var shared = PackageDelta.Compare(edit.BeforeManifest, edit.AfterManifest);
        var projected = DeliveryPackageManifestAdapter.Compare(
            edit.BeforeManifest, edit.AfterManifest);

        Assert.Equal(shared.Count, projected.Count);
        for (int index = 0; index < shared.Count; index++)
        {
            var expected = shared[index];
            var actual = projected[index];
            Assert.Equal(expected.Kind switch
            {
                PackageDeltaChangeKind.EntryAdded => DeliveryPackageChangeKind.PartAdded,
                PackageDeltaChangeKind.EntryRemoved => DeliveryPackageChangeKind.PartRemoved,
                PackageDeltaChangeKind.EntryModified => DeliveryPackageChangeKind.PartModified,
                PackageDeltaChangeKind.RelationshipAdded =>
                    DeliveryPackageChangeKind.RelationshipAdded,
                PackageDeltaChangeKind.RelationshipRemoved =>
                    DeliveryPackageChangeKind.RelationshipRemoved,
                PackageDeltaChangeKind.RelationshipModified =>
                    DeliveryPackageChangeKind.RelationshipModified,
                _ => throw new ArgumentOutOfRangeException(),
            }, actual.Kind);
            Assert.Equal(expected.Location, actual.Location);
            Assert.Equal(expected.BeforeValue, actual.Before);
            Assert.Equal(expected.AfterValue, actual.After);
        }
    }

    [Fact]
    public void DCR031_PackageChangeVerification_UsesSharedTargetAwareOrdering()
    {
        var edit = SingleEdit("Target-aware package change order.");
        var duplicateRelationships = new[]
        {
            new PackageRelationship
            {
                OwnerUri = "/word/document.xml",
                Id = "rDuplicate",
                Type = "a-type",
                Target = "z.xml",
                TargetMode = "Internal",
                ResolvedTargetUri = "/word/z.xml",
                IsTargetPresent = false,
            },
            new PackageRelationship
            {
                OwnerUri = "/word/document.xml",
                Id = "rDuplicate",
                Type = "z-type",
                Target = "a.xml",
                TargetMode = "Internal",
                ResolvedTargetUri = "/word/a.xml",
                IsTargetPresent = false,
            },
        };
        var sourceManifest = edit.BeforeManifest with
        {
            Relationships = edit.BeforeManifest.Relationships
                .Concat(duplicateRelationships).ToArray(),
        };
        var builder = new DeliveryChangeReceiptBuilder(
            sourceManifest, edit.Result.BaseVersion)
            .SetDeliveredDocument(edit.AfterManifest, edit.Result.ResultVersion);
        AddCleanDocx(builder, edit.AfterBytes, edit.AfterManifest, edit.Result.ResultVersion);
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(edit.BeforeBytes, edit.AfterBytes)));
        var entryId = builder.AddTransaction(edit.Contribution);
        AddTransactionSemantic(
            builder, entryId, edit.BeforeBytes, edit.AfterBytes,
            "semantic-source-to-delivered");

        var receipt = builder.Build();
        var duplicateChanges = receipt.Payload.PackageChanges.Where(change =>
            change.Location.RelationshipId == "rDuplicate").ToArray();

        Assert.Equal(
            new[] { "/word/a.xml", "/word/z.xml" },
            duplicateChanges.Select(change => change.Location.TargetUri));
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            receipt, RequiredArtifactBytes(edit)).IsValid);
    }

    [Fact]
    public void DCR032_RetryIdentity_RejectsDifferentResultEvidence()
    {
        var sourceBytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var manifest = Manifest(sourceBytes);
        var operation = DeliveryNormalizedOperation.Create("docx_edit", "replace_text");
        var identity = new DeliveryTransactionIdentity
        {
            TransactionId = "delivery-conflicting-result",
            RequestFingerprint = Fingerprint("same retry request"),
        };
        MutationBatchResult FailedResult(string message)
        {
            var error = new EditError(EditErrorCode.PreconditionFailed, message);
            return new MutationBatchResult
            {
                Mode = MutationBatchMode.Atomic,
                Success = false,
                RolledBack = true,
                BaseVersion = 7,
                ResultVersion = 7,
                Steps = new[]
                {
                    new MutationBatchStepResult(0, operation.Tool, operation.Action,
                        new[] { new EditResult { Success = false, Error = error } }, true),
                },
                Failure = new MutationBatchFailure(
                    0, operation.Tool, operation.Action, error, true),
            };
        }
        DeliveryTransactionContribution Contribution(string message) =>
            DeliveryTransactionContribution.FromMutationBatchResult(
                FailedResult(message), manifest, manifest, new[] { operation }, identity);
        var builder = new DeliveryChangeReceiptBuilder(manifest, 7);

        builder.AddTransaction(Contribution("first failure evidence"));
        var error = Assert.Throws<DeliveryReceiptValidationException>(() =>
            builder.AddTransaction(Contribution("different failure evidence")));

        Assert.Equal("retry_result_conflict", error.Code);
    }

    [Fact]
    public void DCR033_AuthoredChanges_DisambiguateDuplicateRevisionIdsAcrossParts()
    {
        var edit = SingleEdit("Duplicate revision identity.", tracked: true);
        var revision = edit.Result.RevisionChanges.Added.First();
        var duplicate = revision with
        {
            PartUri = "/word/header1.xml",
            Scope = "hdr1",
        };
        var result = edit.Result with
        {
            RevisionChanges = new MutationBatchChangeSet<RevisionListEntry>(
                new[] { revision, duplicate },
                Array.Empty<RevisionListEntry>(),
                Array.Empty<RevisionListEntry>()),
        };
        var operation = DeliveryNormalizedOperation.Create(
            "docx_edit", "replace_text",
            JsonSerializer.Serialize(new
            {
                anchorId = edit.AnchorId,
                markdown = "Duplicate revision identity.",
            }));
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, edit.BeforeManifest, edit.AfterManifest, new[] { operation });
        var builder = Builder(edit);
        var entryId = builder.AddTransaction(contribution);
        AddTransactionSemantic(
            builder, entryId, edit.BeforeBytes, edit.AfterBytes,
            "semantic-source-to-delivered");

        var receipt = builder.Build();
        var duplicateIds = Assert.Single(receipt.Payload.Transactions).AuthoredChanges
            .Where(change => change.EntityKind == DeliveryAuthoredEntityKind.Revision
                && change.EntityId == revision.Id)
            .ToArray();

        Assert.Equal(2, duplicateIds.Length);
        Assert.Equal(
            new[] { revision.PartUri, duplicate.PartUri }.OrderBy(value => value),
            duplicateIds.Select(change => change.PartUri));
        Assert.True(DeliveryChangeReceiptVerifier.Verify(
            receipt, RequiredArtifactBytes(edit)).IsValid);
    }

    [Fact]
    public void DCR034_DistinctNoOpTransactions_HaveDistinctEntryIdentities()
    {
        var documentBytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var manifest = Manifest(documentBytes);
        var operation = DeliveryNormalizedOperation.Create("docx_edit", "no_op");
        var result = new MutationBatchResult
        {
            Mode = MutationBatchMode.Atomic,
            Success = true,
            RolledBack = false,
            BaseVersion = 7,
            ResultVersion = 7,
            Steps = new[]
            {
                new MutationBatchStepResult(
                    0, operation.Tool, operation.Action, Array.Empty<EditResult>(), false),
            },
        };
        DeliveryTransactionContribution Contribution(string? transactionId) =>
            DeliveryTransactionContribution.FromMutationBatchResult(
                result, manifest, manifest, new[] { operation },
                transactionId is null
                    ? null
                    : new DeliveryTransactionIdentity
                    {
                        TransactionId = transactionId,
                        RequestFingerprint = Fingerprint("shared no-op request"),
                    });
        var builder = new DeliveryChangeReceiptBuilder(manifest, 7)
            .SetDeliveredDocument(manifest, 7);
        AddCleanDocx(builder, documentBytes, manifest, 7);
        var emptySemantic = new SemanticChangeSet(Array.Empty<SemanticChange>());
        builder.AddSemanticChangeSet(
            DeliverySemanticChangeSetInput.ForSourceToDelivered(emptySemantic));

        var entryIds = new[]
        {
            builder.AddTransaction(Contribution("delivery-no-op-a")),
            builder.AddTransaction(Contribution("delivery-no-op-b")),
            builder.AddTransaction(Contribution(null)),
            builder.AddTransaction(Contribution(null)),
        };
        var receipt = builder.Build();

        Assert.Equal(entryIds.Length, entryIds.Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(entryIds, receipt.Payload.Transactions.Select(entry => entry.EntryId));
        Assert.True(DeliveryChangeReceiptVerifier.Verify(receipt, new Dictionary<string, byte[]>
        {
            ["clean-docx"] = documentBytes,
            ["semantic-source-to-delivered"] = emptySemantic.ToCanonicalUtf8Bytes(),
        }).IsValid);
    }

    private static DeliveryChangeReceipt BuildWithProfile(
        EditFixture edit,
        DeliveryReceiptPrivacyProfile profile)
    {
        var builder = Builder(edit, profile);
        AddTransactionWithSemantic(builder, edit);
        return builder.Build();
    }

    private static DeliveryChangeReceiptBuilder Builder(
        EditFixture edit,
        DeliveryReceiptPrivacyProfile profile = DeliveryReceiptPrivacyProfile.HashAndSummary)
    {
        var builder = new DeliveryChangeReceiptBuilder(
            edit.BeforeManifest, edit.Result.BaseVersion, profile)
            .SetDeliveredDocument(edit.AfterManifest, edit.Result.ResultVersion);
        AddCleanDocx(builder, edit.AfterBytes, edit.AfterManifest, edit.Result.ResultVersion);
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
            SemanticChanges(edit.BeforeBytes, edit.AfterBytes)));
        return builder;
    }

    private static string AddTransactionWithSemantic(
        DeliveryChangeReceiptBuilder builder,
        EditFixture edit)
    {
        var entryId = builder.AddTransaction(edit.Contribution);
        AddTransactionSemantic(
            builder, entryId, edit.BeforeBytes, edit.AfterBytes,
            "semantic-source-to-delivered");
        return entryId;
    }

    private static void AddTransactionSemantic(
        DeliveryChangeReceiptBuilder builder,
        string entryId,
        byte[] beforeBytes,
        byte[] afterBytes,
        string artifactId)
    {
        builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForTransaction(
            entryId, SemanticChanges(beforeBytes, afterBytes), artifactId));
    }

    private static void AddCleanDocx(
        DeliveryChangeReceiptBuilder builder,
        byte[] deliveredBytes,
        PackageManifest deliveredManifest,
        long deliveredVersion)
    {
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "clean-docx",
            DeliveryArtifactRole.CleanDocx,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            deliveredBytes) with
        {
            Document = DeliveryDocumentIdentity.FromManifest(
                deliveredManifest, deliveredVersion),
            RelativePath = "delivery/clean.docx",
        });
    }

    private static void AddPaginatedCitation(
        DeliveryChangeReceiptBuilder builder,
        DeliveryDocumentIdentity identity,
        PageCitation citation,
        byte[] pdfBytes,
        byte[] pageMapBytes)
    {
        var pageMapDigest = Digest(pageMapBytes);
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "pdf", DeliveryArtifactRole.Pdf, "application/pdf", pdfBytes) with
        {
            Document = identity,
            RendererFingerprint = citation.RendererFingerprint,
            PageMapDigest = pageMapDigest,
        });
        builder.AddArtifact(DeliveryArtifactInput.Available(
            "page-map", DeliveryArtifactRole.PageMap, "application/json", pageMapBytes) with
        {
            Document = identity,
            RendererFingerprint = citation.RendererFingerprint,
        });
        builder.AddPageCitation(new DeliveryPageCitationInput
        {
            Citation = citation,
            Scope = "body",
            Document = identity,
            PageMapDigest = pageMapDigest,
            PageMapArtifactId = "page-map",
            RenderArtifactId = "pdf",
        });
    }

    private static DeliveryChangeReceipt Rehash(DeliveryChangeReceiptPayload payload) => new()
    {
        Payload = payload,
        ReceiptDigest = Digest(DeliveryChangeReceiptSerializer.SerializePayload(payload)),
    };

    private static SemanticChangeSet SemanticChanges(byte[] before, byte[] after) =>
        SemanticDiff.Compare(
            new WmlDocument("before.docx", before),
            new WmlDocument("after.docx", after));

    private static SemanticChangeSet SemanticSet(SemanticValue before) => new(new[]
    {
        new SemanticChange
        {
            Id = "ignored",
            Operation = SemanticChangeOperation.Modify,
            Family = SemanticChangeFamily.Text,
            PartUri = "/word/document.xml",
            Path = "body/paragraph[1]",
            Before = before,
            After = SemanticValue.Absent,
        },
    });

    private static Dictionary<string, byte[]> RequiredArtifactBytes(EditFixture edit) => new()
    {
        ["clean-docx"] = edit.AfterBytes,
        ["semantic-source-to-delivered"] =
            SemanticChanges(edit.BeforeBytes, edit.AfterBytes).ToCanonicalUtf8Bytes(),
    };

    private static EditFixture SingleEdit(
        string replacement,
        bool tracked = false,
        DeliveryTransactionIdentity? identity = null)
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(source, new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "Receipt Author",
        });
        var anchor = BodyParagraphs(session)[0];
        var beforeBytes = session.Save();
        var beforeManifest = Manifest(beforeBytes);
        var operation = DeliveryNormalizedOperation.Create("docx_edit", "replace_text",
            JsonSerializer.Serialize(new { anchorId = anchor, markdown = replacement }));
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text",
                s => s.ReplaceText(anchor, replacement)),
        });
        Assert.True(result.Success,
            result.Failure is null ? "batch failed" : result.Failure.Error.Message);
        var afterBytes = session.Save();
        var afterManifest = Manifest(afterBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result, beforeManifest, afterManifest, new[] { operation }, identity);
        return new EditFixture(
            anchor, beforeBytes, afterBytes, beforeManifest, afterManifest, result, contribution);
    }

    private static PageCitation Citation(string anchorId, long version, string renderer) => new()
    {
        AnchorId = anchorId,
        Availability = PageMapAvailability.Available,
        DocumentVersion = version,
        RendererFingerprint = renderer,
        Pages = new[]
        {
            new PageMapPage
            {
                PageNumber = 1,
                PageInSection = 1,
                Width = 612,
                Height = 792,
                SectionIndex = 0,
                PageName = "page-1",
            },
        },
        Fragments = new[]
        {
            new PageMapFragment
            {
                FragmentId = "page-1-fragment-0",
                AnchorId = anchorId,
                FragmentIndex = 0,
                PageNumber = 1,
                Geometry = new PageMapRect(72, 72, 300, 24),
                Story = PageMapStory.Body,
            },
        },
    };

    private static byte[] PageMapBytes(string anchorId, long version, string renderer)
    {
        var citation = Citation(anchorId, version, renderer);
        return JsonSerializer.SerializeToUtf8Bytes(new PageMap
        {
            Mode = PageMapMode.Paginated,
            Availability = PageMapAvailability.Available,
            DocumentVersion = version,
            RendererFingerprint = renderer,
            Pages = citation.Pages,
            Fragments = citation.Fragments,
        }, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Converters =
            {
                new System.Text.Json.Serialization.JsonStringEnumConverter(
                    JsonNamingPolicy.CamelCase, allowIntegerValues: false),
            },
        });
    }

    private static DocxSession Open(byte[] bytes) => new(bytes, new DocxSessionSettings
    {
        PersistAnchorIds = true,
    });

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:", StringComparison.Ordinal))
            .ToArray();

    private static PackageManifest Manifest(byte[] bytes) =>
        PackageManifestGenerator.Generate(bytes);

    private static VerificationDigest Digest(byte[] bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    private static string Fingerprint(string value) => "sha256:" + Digest(
        Encoding.UTF8.GetBytes(value)).Value;

    private sealed record EditFixture(
        string AnchorId,
        byte[] BeforeBytes,
        byte[] AfterBytes,
        PackageManifest BeforeManifest,
        PackageManifest AfterManifest,
        MutationBatchResult Result,
        DeliveryTransactionContribution Contribution);
}
