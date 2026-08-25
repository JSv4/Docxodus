// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using Docxodus.Verification;
using DeliverableAvailability = Docxodus.Verification.DeliverableArtifactAvailability;
using ReceiptArtifactAvailability = Docxodus.Verification.DeliveryArtifactAvailability;
using ReceiptArtifactInput = Docxodus.Verification.DeliveryArtifactInput;
using ReceiptArtifactRole = Docxodus.Verification.DeliveryArtifactRole;

namespace Docxodus.Delivery;

/// <summary>
/// Builds one immutable delivery bundle from exact baseline and working package snapshots. The
/// service owns dependency closure and evidence composition; rendering remains an injected,
/// capability-declared adapter.
/// </summary>
public sealed class DeliveryBundleService
{
    private const string DocxMediaType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private readonly IDeliveryArtifactRenderer? _renderer;

    public DeliveryBundleService(IDeliveryArtifactRenderer? renderer = null) =>
        _renderer = renderer;

    /// <summary>Return a verified in-memory bundle without publishing filesystem state.</summary>
    public async ValueTask<DeliveryBundle> BuildAsync(
        DeliveryBundleBuildRequest request,
        DeliveryBundleBuildOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        cancellationToken.ThrowIfCancellationRequested();
        options ??= new DeliveryBundleBuildOptions();
        ValidateOptions(options);
        var requested = ValidateRequests(request.ArtifactSnapshot);
        var plans = PlanArtifacts(requested);

        bool needsReview = plans.Any(plan =>
            plan.Request.Kind is DeliveryArtifactKind.ReviewDocx
                or DeliveryArtifactKind.ReversibilityProof
            || plan.Request.ReviewProfile == DeliveryReviewProfile.Markup);
        var policy = DeliveryRevisionPolicyProcessor.Apply(
            request.Baseline.CopyBytes(),
            request.Working.CopyBytes(),
            request.RevisionPolicy,
            needsReview,
            options.PackageManifestOptions);

        var policyBaseline = new DeliveryDocumentSnapshot(
            "policy-baseline",
            request.Baseline.DocumentVersion,
            policy.PolicyBaselineBytes);
        var final = new DeliveryDocumentSnapshot(
            request.FinalDocumentName,
            request.FinalDocumentVersion,
            policy.FinalBytes);
        var review = policy.ReviewBytes is null
            ? null
            : new DeliveryDocumentSnapshot(
                "review",
                request.FinalDocumentVersion,
                policy.ReviewBytes);
        var materializedRequest = new DeliveryBundleRequest(
            request.Baseline,
            request.Working,
            final,
            request.RevisionPolicy,
            plans.Where(plan => !plan.IsImplicit).Select(plan => plan.Request));

        var baselineManifest = PackageManifestGenerator.Generate(
            request.Baseline.CopyBytes(), options.PackageManifestOptions);
        var finalManifest = PackageManifestGenerator.Generate(
            final.CopyBytes(), options.PackageManifestOptions);
        EnsureValidManifest("baseline", baselineManifest);
        EnsureValidManifest("final", finalManifest);
        var semantic = SemanticDiff.Compare(
            new WmlDocument(request.Baseline.Name, request.Baseline.CopyBytes()),
            new WmlDocument(final.Name, final.CopyBytes()),
            new SemanticDiffOptions { PackageOptions = options.PackageManifestOptions });
        var packageDelta = DeliveryPackageDeltaReport.Create(baselineManifest, finalManifest);

        var outputs = new Dictionary<string, DeliveryBundleArtifactInput>(StringComparer.Ordinal);
        var renderStates = new List<RenderState>();
        foreach (var plan in plans.Where(plan => !IsDeferred(plan.Request.Kind)))
        {
            AddCoreArtifact(
                outputs,
                plan,
                request,
                policyBaseline,
                final,
                review,
                policy,
                baselineManifest,
                finalManifest,
                semantic,
                packageDelta);
        }

        var rendered = await RenderAllAsync(
            plans.Where(plan => DeliveryBundleManifest.IsProfiledRenderKind(plan.Request.Kind)),
            policyBaseline,
            final,
            review,
            options.PackageManifestOptions,
            cancellationToken).ConfigureAwait(false);
        foreach (var state in rendered)
        {
            outputs.Add(state.Plan.Request.ArtifactId, state.Output);
            renderStates.Add(state);
        }

        DeliverableVerificationResult? finalVerification = null;
        foreach (var plan in plans.Where(plan =>
                     plan.Request.Kind == DeliveryArtifactKind.ValidationReport))
        {
            var verificationOptions = options.DeliverableVerificationOptions with
            {
                PackageManifestOptions = options.PackageManifestOptions,
            };
            var finalReport = DeliverableVerifier.VerifyDeliverable(
                new DeliverableVerificationRequest
                {
                    BaselineBytes = request.Baseline.CopyBytes(),
                    DeliverableBytes = final.CopyBytes(),
                    ExpectedSemanticChanges = semantic,
                },
                verificationOptions);
            var cohortReports = renderStates
                .GroupBy(state => new RenderProfileKey(
                    state.Plan.Request.ReviewProfile!.Value,
                    state.Plan.Request.CommentProfile!.Value))
                .OrderBy(group => group.Key.ReviewProfile)
                .ThenBy(group => group.Key.CommentProfile)
                .Select(group =>
                {
                    var source = RenderSource(
                        group.Key.ReviewProfile, policyBaseline, final, review);
                    var sourceBytes = source.CopyBytes();
                    return new DeliveryRenderCohortValidation
                    {
                        ReviewProfile = group.Key.ReviewProfile,
                        CommentProfile = group.Key.CommentProfile,
                        SourceDocument = new DeliveryBundleDocumentIdentity
                        {
                            Name = source.Name,
                            DocumentVersion = source.DocumentVersion,
                            ByteLength = sourceBytes.LongLength,
                            Digest = DeliveryBundleCanonicalJson.Digest(sourceBytes),
                        },
                        Verification = DeliverableVerifier.VerifyDeliverable(
                            new DeliverableVerificationRequest
                            {
                                BaselineBytes = sourceBytes,
                                DeliverableBytes = sourceBytes,
                                CompanionArtifacts = group.Select(ToDeliverableCompanion).ToArray(),
                            },
                            verificationOptions),
                    };
                })
                .ToArray();
            var report = new DeliveryBundleValidationReport
            {
                Decision = CombineVerificationDecisions(
                    new[] { finalReport.Decision }.Concat(
                        cohortReports.Select(cohort => cohort.Verification.Decision))),
                FinalDeliverable = finalReport,
                RenderCohorts = cohortReports,
            };
            finalVerification = finalReport;
            if (options.FailOnDeliverableValidationFailure
                && report.Decision == DeliverableVerificationDecision.Failed)
            {
                var failures = finalReport.Decision == DeliverableVerificationDecision.Failed
                    ? new[]
                    {
                        "final:" + string.Join(",", finalReport.Findings
                            .Where(finding => finding.BlocksDelivery)
                            .Select(finding => finding.Code)
                            .Distinct(StringComparer.Ordinal)
                            .OrderBy(code => code, StringComparer.Ordinal)),
                    }
                    : Array.Empty<string>();
                failures = failures.Concat(cohortReports
                        .Where(cohort => cohort.Verification.Decision
                            == DeliverableVerificationDecision.Failed)
                        .Select(cohort =>
                            $"{Kebab(cohort.ReviewProfile)}/{Kebab(cohort.CommentProfile)}:"
                            + string.Join(",", cohort.Verification.Findings
                                .Where(finding => finding.BlocksDelivery)
                                .Select(finding => finding.Code)
                                .Distinct(StringComparer.Ordinal)
                                .OrderBy(code => code, StringComparer.Ordinal))))
                    .ToArray();
                throw new DeliveryBundleException(
                    "deliverable_validation_failed",
                    "The bundle failed the selected deliverable-verification policy: "
                    + string.Join("; ", failures));
            }
            outputs.Add(plan.Request.ArtifactId,
                Available(plan, report.ToCanonicalUtf8Bytes()));
        }

        foreach (var plan in plans.Where(plan =>
                     plan.Request.Kind == DeliveryArtifactKind.ChangeReceipt).ToArray())
        {
            BuildReceipt(
                request,
                options,
                plan,
                plans,
                outputs,
                baselineManifest,
                finalManifest,
                final,
                review,
                semantic,
                policy.ReviewProof,
                finalVerification);
        }

        EnforceRequiredAvailability(outputs.Values, options.ReturnIncompleteBundle);
        var relationships = BuildRelationships(plans, outputs.Values, renderStates);
        return DeliveryBundle.Create(
            materializedRequest,
            outputs.Values,
            relationships,
            limits: options.BundleVerificationLimits);
    }

    private static DeliverableVerificationDecision CombineVerificationDecisions(
        IEnumerable<DeliverableVerificationDecision> decisions)
    {
        var materialized = decisions.ToArray();
        if (materialized.Any(decision => decision == DeliverableVerificationDecision.Failed))
            return DeliverableVerificationDecision.Failed;
        if (materialized.Any(decision =>
                decision == DeliverableVerificationDecision.NotEvaluated))
            return DeliverableVerificationDecision.NotEvaluated;
        return materialized.Any(decision =>
                decision == DeliverableVerificationDecision.PassedWithPreExistingFindings)
            ? DeliverableVerificationDecision.PassedWithPreExistingFindings
            : DeliverableVerificationDecision.Passed;
    }

    private static void AddCoreArtifact(
        IDictionary<string, DeliveryBundleArtifactInput> outputs,
        ArtifactPlan plan,
        DeliveryBundleBuildRequest request,
        DeliveryDocumentSnapshot policyBaseline,
        DeliveryDocumentSnapshot final,
        DeliveryDocumentSnapshot? review,
        DeliveryRevisionPolicyResult policy,
        PackageManifest baselineManifest,
        PackageManifest finalManifest,
        SemanticChangeSet semantic,
        DeliveryPackageDeltaReport packageDelta)
    {
        byte[] bytes = plan.Request.Kind switch
        {
            DeliveryArtifactKind.BaselineDocx => request.Baseline.CopyBytes(),
            DeliveryArtifactKind.PolicyBaselineDocx => policyBaseline.CopyBytes(),
            DeliveryArtifactKind.WorkingDocx => request.Working.CopyBytes(),
            DeliveryArtifactKind.ReviewDocx => review?.CopyBytes()
                ?? throw new DeliveryBundleException(
                    "review_proof_unavailable", "Review DOCX proof was not produced."),
            DeliveryArtifactKind.FinalDocx => final.CopyBytes(),
            DeliveryArtifactKind.BaselinePackageManifest => baselineManifest.ToJsonBytes(),
            DeliveryArtifactKind.FinalPackageManifest => finalManifest.ToJsonBytes(),
            DeliveryArtifactKind.SemanticDelta => semantic.ToCanonicalUtf8Bytes(),
            DeliveryArtifactKind.PackageDelta => packageDelta.ToCanonicalUtf8Bytes(),
            DeliveryArtifactKind.ReversibilityProof => Encoding.UTF8.GetBytes(
                policy.ReviewProof?.Proof.ToCanonicalJson()
                ?? throw new DeliveryBundleException(
                    "review_proof_unavailable", "Review reversibility proof was not produced.")),
            _ => throw new ArgumentOutOfRangeException(nameof(plan), plan.Request.Kind, null),
        };
        outputs.Add(plan.Request.ArtifactId, Available(plan, bytes));
    }

    private async ValueTask<IReadOnlyList<RenderState>> RenderAllAsync(
        IEnumerable<ArtifactPlan> renderPlans,
        DeliveryDocumentSnapshot policyBaseline,
        DeliveryDocumentSnapshot final,
        DeliveryDocumentSnapshot? review,
        PackageManifestOptions packageManifestOptions,
        CancellationToken cancellationToken)
    {
        var plans = renderPlans.ToArray();
        var states = new Dictionary<string, RenderState>(StringComparer.Ordinal);
        var jobs = new List<RenderJob>();
        foreach (var plan in plans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var request = plan.Request;
            var reviewProfile = request.ReviewProfile!.Value;
            var commentProfile = request.CommentProfile!.Value;
            var source = RenderSource(reviewProfile, policyBaseline, final, review);
            var sourceDigest = PackageManifestGenerator.Generate(
                    source.CopyBytes(), packageManifestOptions)
                .RawPackageBytesDigest;
            var job = new RenderJob(
                plan,
                new DeliveryRenderRequest(
                    request.ArtifactId,
                    request.Kind,
                    reviewProfile,
                    commentProfile,
                    source),
                sourceDigest);
            if (_renderer is null)
            {
                states.Add(request.ArtifactId, new RenderState(plan, UnavailableRender(job,
                    "No delivery renderer was supplied."), sourceDigest, null));
                continue;
            }
            if (!_renderer.Capabilities.Supports(
                    request.Kind, reviewProfile, commentProfile))
            {
                states.Add(request.ArtifactId, new RenderState(plan, UnavailableRender(job,
                    $"Renderer '{_renderer.Capabilities.RendererId}' does not advertise this artifact/profile combination."),
                    sourceDigest, null));
                continue;
            }
            jobs.Add(job);
        }

        if (jobs.Count != 0 && _renderer is IDeliveryArtifactBatchRenderer batchRenderer)
        {
            try
            {
                var results = await batchRenderer.RenderBatchAsync(
                    jobs.Select(job => job.Request).ToArray(), cancellationToken)
                    .ConfigureAwait(false);
                ArgumentNullException.ThrowIfNull(results);
                var expectedIds = jobs.Select(job => job.Request.ArtifactId)
                    .ToHashSet(StringComparer.Ordinal);
                if (!expectedIds.SetEquals(results.Keys)
                    || results.Any(pair => pair.Value is null))
                    throw new InvalidDataException(
                        "The renderer batch result does not match the requested artifact IDs.");
                foreach (var job in jobs)
                    states.Add(job.Request.ArtifactId,
                        RenderStateFromResult(job, results[job.Request.ArtifactId]));
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                foreach (var job in jobs)
                    states.Add(job.Request.ArtifactId, RendererFailure(job, ex));
            }
        }
        else
        {
            foreach (var job in jobs)
            {
                try
                {
                    var result = await _renderer!.RenderAsync(job.Request, cancellationToken)
                        .ConfigureAwait(false);
                    states.Add(job.Request.ArtifactId,
                        RenderStateFromResult(job, result
                            ?? throw new InvalidDataException("Renderer returned a null result.")));
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    states.Add(job.Request.ArtifactId, RendererFailure(job, ex));
                }
            }
        }

        return plans.Select(plan => states[plan.Request.ArtifactId]).ToArray();
    }

    private static RenderState RenderStateFromResult(RenderJob job, DeliveryRenderResult result)
    {
        var plan = job.Plan;
        var request = job.Request;
        if (result.Availability == DeliveryArtifactAvailability.Unavailable)
        {
            return new RenderState(plan, DeliveryBundleArtifactInput.Unavailable(
                request.ArtifactId,
                request.Kind,
                RelativePath(plan),
                result.MediaType,
                result.UnavailableReason ?? "Renderer did not produce artifact bytes.",
                plan.IsImplicit,
                plan.Request.Requiredness,
                RenderMetadata(job, result)),
                job.SourceDigest,
                result);
        }
        return new RenderState(plan, DeliveryBundleArtifactInput.Available(
            request.ArtifactId,
            request.Kind,
            RelativePath(plan),
            result.MediaType,
            result.CopyBytes() ?? throw new InvalidDataException(
                "An available renderer result has no bytes."),
            plan.IsImplicit,
            plan.Request.Requiredness,
            RenderMetadata(job, result)),
            job.SourceDigest,
            result);
    }

    private static RenderState RendererFailure(RenderJob job, Exception exception) => new(
        job.Plan,
        UnavailableRender(job,
            RendererFailureReason(exception)),
        job.SourceDigest,
        null);

    private static string RendererFailureReason(Exception exception)
    {
        var reason = $"Renderer failed with {exception.GetType().Name}: {exception.Message}";
        return reason.Length <= 4096 ? reason : reason[..4096];
    }

    private static DeliveryArtifactRenderMetadataInput RenderMetadata(
        RenderJob job,
        DeliveryRenderResult? result = null) => new()
        {
            ReviewProfile = job.Request.ReviewProfile,
            CommentProfile = job.Request.CommentProfile,
            SourceDocumentName = job.Request.SourceDocumentName,
            SourceDocumentVersion = job.Request.SourceDocumentVersion,
            SourcePackageDigest = job.SourceDigest,
            RendererFingerprint = result?.RendererFingerprint,
            PageCount = result?.PageCount,
            Warnings = result?.Diagnostics.ToArray()
                ?? Array.Empty<DeliverableRenderDiagnostic>(),
        };

    private static DeliverableCompanionArtifactInput ToDeliverableCompanion(RenderState state)
    {
        var output = state.Output;
        var result = state.Result;
        return new DeliverableCompanionArtifactInput
        {
            ArtifactId = output.ArtifactId,
            Role = output.Kind switch
            {
                DeliveryArtifactKind.StandaloneHtml => DeliverableArtifactRole.Html,
                DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf =>
                    DeliverableArtifactRole.Pdf,
                DeliveryArtifactKind.PageMap => DeliverableArtifactRole.PageMap,
                DeliveryArtifactKind.RenderReport => DeliverableArtifactRole.RenderReport,
                _ => DeliverableArtifactRole.Other,
            },
            MediaType = output.MediaType,
            Availability = output.Availability == DeliveryArtifactAvailability.Available
                ? DeliverableAvailability.Available
                : DeliverableAvailability.Unavailable,
            Bytes = output.CopyBytes(),
            UnavailableReason = output.UnavailableReason,
            PageCount = result?.PageCount,
            RendererFingerprint = result?.RendererFingerprint,
            SourcePackageDigest = state.SourceDigest,
            PageMapDigest = result?.CopyPageMapBytes() is { } pageMap
                ? DeliveryBundleCanonicalJson.Digest(pageMap)
                : null,
            RenderDiagnostics = result?.Diagnostics ?? Array.Empty<DeliverableRenderDiagnostic>(),
        };
    }

    private static void BuildReceipt(
        DeliveryBundleBuildRequest request,
        DeliveryBundleBuildOptions options,
        ArtifactPlan receiptPlan,
        List<ArtifactPlan> plans,
        Dictionary<string, DeliveryBundleArtifactInput> outputs,
        PackageManifest baselineManifest,
        PackageManifest finalManifest,
        DeliveryDocumentSnapshot final,
        DeliveryDocumentSnapshot? review,
        SemanticChangeSet semantic,
        RedlineReversibilityProofRun? proof,
        DeliverableVerificationResult? finalVerification)
    {
        if (request.ReceiptContext is null)
        {
            outputs.Add(receiptPlan.Request.ArtifactId, Unavailable(receiptPlan,
                "Authoritative transaction evidence was not supplied."));
            return;
        }

        try
        {
            var context = request.ReceiptContext;
            var stagedPlans = new List<ArtifactPlan>();
            var stagedOutputs = new Dictionary<string, DeliveryBundleArtifactInput>(
                StringComparer.Ordinal);
            var builder = new DeliveryChangeReceiptBuilder(
                baselineManifest,
                request.Baseline.DocumentVersion,
                context.PrivacyProfile,
                options.DeliveryReceiptLimits)
                .SetDeliveredDocument(finalManifest, final.DocumentVersion);
            builder.FailOnUnexpectedChanges = context.FailOnUnexpectedChanges;

            var finalIdentity = DeliveryDocumentIdentity.FromManifest(
                finalManifest, final.DocumentVersion);
            var finalPlan = plans.Single(plan =>
                plan.Request.Kind == DeliveryArtifactKind.FinalDocx);
            builder.AddArtifact(ReceiptArtifactInput.Available(
                finalPlan.Request.ArtifactId,
                ReceiptArtifactRole.CleanDocx,
                DocxMediaType,
                final.CopyBytes()) with
            {
                RelativePath = RelativePath(finalPlan),
                Document = finalIdentity,
            });

            var semanticPlan = plans.Single(plan =>
                plan.Request.Kind == DeliveryArtifactKind.SemanticDelta);
            builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
                semantic,
                semanticPlan.Request.ArtifactId,
                RelativePath(semanticPlan)));

            foreach (var evidence in context.TransactionSnapshot)
            {
                ValidateTransactionSnapshots(evidence, options.PackageManifestOptions);
                var entryId = builder.AddTransaction(evidence.Contribution);
                if (!Equals(evidence.Contribution.BeforeDocument.RawPackageBytesDigest,
                        evidence.Contribution.AfterDocument.RawPackageBytesDigest))
                {
                    var transactionSemantic = SemanticDiff.Compare(
                        new WmlDocument(evidence.Before.Name, evidence.Before.CopyBytes()),
                        new WmlDocument(evidence.After.Name, evidence.After.CopyBytes()),
                        new SemanticDiffOptions
                        {
                            PackageOptions = options.PackageManifestOptions,
                        });
                    var transactionId = ReserveId(
                        plans.Select(value => value.Request.ArtifactId)
                            .Concat(outputs.Keys)
                            .Concat(stagedPlans.Select(value => value.Request.ArtifactId)),
                        "semantic-transaction-" + entryId);
                    var transactionPlan = new ArtifactPlan(new DeliveryArtifactRequest
                    {
                        ArtifactId = transactionId,
                        Kind = DeliveryArtifactKind.SemanticDelta,
                        Requiredness = DeliveryArtifactRequiredness.Required,
                    }, true);
                    stagedPlans.Add(transactionPlan);
                    var transactionBytes = transactionSemantic.ToCanonicalUtf8Bytes();
                    stagedOutputs.Add(
                        transactionId,
                        Available(transactionPlan, transactionBytes));
                    builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForTransaction(
                        entryId,
                        transactionSemantic,
                        transactionId,
                        RelativePath(transactionPlan)));
                }
            }

            foreach (var lineage in context.LineageSnapshot)
                builder.AddLineageEvent(lineage);
            foreach (var rule in context.AttributionRuleSnapshot)
                builder.AddAttributionRule(rule);
            foreach (var warning in context.WarningSnapshot)
                builder.AddWarning(warning);

            foreach (var output in outputs.Values.Where(output =>
                         output.Kind is not DeliveryArtifactKind.FinalDocx
                             and not DeliveryArtifactKind.SemanticDelta
                             and not DeliveryArtifactKind.ChangeReceipt))
            {
                builder.AddArtifact(ToReceiptArtifact(output));
            }

            // The receipt contract pins ValidationResult evidence to the exact canonical
            // deliverable-verification document, not to the bundle's aggregate validation
            // report — so the final deliverable's own verification result is staged as an
            // implicit artifact and the evidence binds to those bytes. The aggregate report
            // still enters the receipt above, as a plain report artifact.
            var hasAggregateValidation = outputs.Values.Any(output =>
                output.Kind == DeliveryArtifactKind.ValidationReport
                && output.Availability == DeliveryArtifactAvailability.Available);
            if (hasAggregateValidation && finalVerification is not null)
            {
                var verificationBytes = finalVerification.ToCanonicalUtf8Bytes();
                var verificationId = ReserveId(
                    plans.Select(value => value.Request.ArtifactId)
                        .Concat(outputs.Keys)
                        .Concat(stagedPlans.Select(value => value.Request.ArtifactId)),
                    "final-verification");
                var verificationPlan = new ArtifactPlan(new DeliveryArtifactRequest
                {
                    ArtifactId = verificationId,
                    Kind = DeliveryArtifactKind.ValidationReport,
                    Requiredness = DeliveryArtifactRequiredness.Required,
                }, true);
                stagedPlans.Add(verificationPlan);
                var verificationOutput = Available(verificationPlan, verificationBytes);
                stagedOutputs.Add(verificationId, verificationOutput);
                builder.AddArtifact(ReceiptArtifactInput.Available(
                    verificationId,
                    ReceiptArtifactRole.ValidationReport,
                    verificationOutput.MediaType,
                    verificationBytes) with
                {
                    RelativePath = RelativePath(verificationPlan),
                });
                builder.AddEvidence(new DeliveryEvidenceReference
                {
                    Kind = DeliveryEvidenceKind.ValidationResult,
                    Schema = DeliverableVerificationResult.SchemaId,
                    Digest = DeliveryBundleCanonicalJson.Digest(verificationBytes),
                    ArtifactId = verificationId,
                    Summary = "Final deliverable verification result.",
                });
            }
            var proofOutput = outputs.Values.FirstOrDefault(output =>
                output.Kind == DeliveryArtifactKind.ReversibilityProof
                && output.Availability == DeliveryArtifactAvailability.Available);
            if (proof is not null && proofOutput?.CopyBytes() is { } proofBytes)
            {
                builder.AddEvidence(new DeliveryEvidenceReference
                {
                    Kind = DeliveryEvidenceKind.RedlineReversibility,
                    Schema = RedlineReversibilityProof.SchemaId,
                    Digest = DeliveryBundleCanonicalJson.Digest(proofBytes),
                    ArtifactId = proofOutput.ArtifactId,
                    Summary = "Review DOCX reversibility proof.",
                });
            }

            var receipt = builder.Build();
            var receiptArtifactBytes = receipt.Payload.Artifacts
                .Where(value => value.Availability == ReceiptArtifactAvailability.Available)
                .ToDictionary(
                    value => value.ArtifactId,
                    value => (outputs.TryGetValue(value.ArtifactId, out var output)
                            ? output
                            : stagedOutputs[value.ArtifactId]).CopyBytes()
                        ?? throw new InvalidDataException(
                            $"Receipt artifact '{value.ArtifactId}' has no bytes."),
                    StringComparer.Ordinal);
            var verification = DeliveryChangeReceiptVerifier.Verify(
                receipt, receiptArtifactBytes);
            if (!verification.IsValid)
                throw new InvalidDataException(
                    $"Delivery receipt verification failed: {verification.Findings[0]}");
            plans.AddRange(stagedPlans);
            foreach (var staged in stagedOutputs)
                outputs.Add(staged.Key, staged.Value);
            outputs.Add(receiptPlan.Request.ArtifactId,
                Available(receiptPlan, receipt.ToJsonBytes()));
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidDataException)
        {
            outputs.Add(receiptPlan.Request.ArtifactId, Unavailable(receiptPlan,
                $"Authoritative receipt evidence was rejected ({ex.GetType().Name})."));
        }
    }

    private static ReceiptArtifactInput ToReceiptArtifact(DeliveryBundleArtifactInput output)
    {
        var role = output.Kind switch
        {
            DeliveryArtifactKind.ReviewDocx => ReceiptArtifactRole.ReviewDocx,
            DeliveryArtifactKind.StandaloneHtml => ReceiptArtifactRole.Html,
            DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf =>
                ReceiptArtifactRole.Pdf,
            DeliveryArtifactKind.PageMap => ReceiptArtifactRole.PageMap,
            DeliveryArtifactKind.BaselinePackageManifest
                or DeliveryArtifactKind.FinalPackageManifest => ReceiptArtifactRole.PackageManifest,
            // The aggregate bundle validation report is a plain receipt report; the typed
            // ValidationReport role is reserved for the staged canonical verification result
            // the ValidationResult evidence binds to.
            DeliveryArtifactKind.ReversibilityProof => ReceiptArtifactRole.ReversibilityProof,
            DeliveryArtifactKind.RenderReport => ReceiptArtifactRole.RenderReport,
            _ => ReceiptArtifactRole.OtherReport,
        };
        ReceiptArtifactInput input = output.Availability == DeliveryArtifactAvailability.Available
            ? ReceiptArtifactInput.Available(
                output.ArtifactId, role, output.MediaType,
                output.CopyBytes() ?? throw new InvalidDataException(
                    $"Artifact '{output.ArtifactId}' has no bytes."))
            : ReceiptArtifactInput.Unavailable(
                output.ArtifactId, role, output.MediaType,
                output.UnavailableReason ?? "Artifact unavailable.");
        return input with
        {
            RelativePath = output.RelativePath,
            RendererFingerprint = output.RenderMetadata?.RendererFingerprint,
        };
    }

    private static void ValidateTransactionSnapshots(
        DeliveryReceiptTransactionEvidence evidence,
        PackageManifestOptions options)
    {
        var before = PackageManifestGenerator.Generate(evidence.Before.CopyBytes(), options);
        var after = PackageManifestGenerator.Generate(evidence.After.CopyBytes(), options);
        var beforeIdentity = DeliveryDocumentIdentity.FromManifest(
            before, evidence.Before.DocumentVersion);
        var afterIdentity = DeliveryDocumentIdentity.FromManifest(
            after, evidence.After.DocumentVersion);
        if (beforeIdentity != evidence.Contribution.BeforeDocument
            || afterIdentity != evidence.Contribution.AfterDocument)
        {
            throw new DeliveryReceiptValidationException(
                "transaction_snapshot_identity_mismatch",
                "Receipt transaction snapshots do not match their authoritative contribution identities.");
        }
    }

    private static List<ArtifactPlan> PlanArtifacts(
        IReadOnlyList<DeliveryArtifactRequest> requested)
    {
        var plans = requested
            .OrderBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Select(value => new ArtifactPlan(value, false))
            .ToList();
        EnsureImplicit(plans, DeliveryArtifactKind.FinalDocx,
            DeliveryArtifactRequiredness.Required, "final-docx");
        EnsureImplicit(plans, DeliveryArtifactKind.ValidationReport,
            DeliveryArtifactRequiredness.Required, "validation-report");
        bool needsReview = plans.Any(plan =>
            plan.Request.Kind is DeliveryArtifactKind.ReviewDocx
                or DeliveryArtifactKind.ReversibilityProof
            || plan.Request.ReviewProfile == DeliveryReviewProfile.Markup);
        bool needsPolicyBaseline = needsReview || plans.Any(plan =>
            plan.Request.ReviewProfile == DeliveryReviewProfile.Original);
        if (needsPolicyBaseline)
            EnsureImplicit(plans, DeliveryArtifactKind.PolicyBaselineDocx,
                DeliveryArtifactRequiredness.Required, "policy-baseline-docx");
        if (needsReview)
        {
            EnsureImplicit(plans, DeliveryArtifactKind.ReviewDocx,
                DeliveryArtifactRequiredness.Required, "review-docx");
            EnsureImplicit(plans, DeliveryArtifactKind.ReversibilityProof,
                DeliveryArtifactRequiredness.Required, "reversibility-proof");
        }
        if (plans.Any(plan => plan.Request.Kind == DeliveryArtifactKind.ChangeReceipt))
        {
            EnsureImplicit(plans, DeliveryArtifactKind.SemanticDelta,
                DeliveryArtifactRequiredness.Required, "semantic-source-to-delivered");
        }
        var renderGroups = plans
            .Where(plan => DeliveryBundleManifest.IsProfiledRenderKind(plan.Request.Kind))
            .GroupBy(plan => new RenderProfileKey(
                plan.Request.ReviewProfile!.Value,
                plan.Request.CommentProfile!.Value))
            .Select(group => new
            {
                group.Key,
                Requiredness = group.Any(plan =>
                    plan.Request.Requiredness == DeliveryArtifactRequiredness.Required)
                    ? DeliveryArtifactRequiredness.Required
                    : DeliveryArtifactRequiredness.Optional,
            })
            .ToArray();
        foreach (var group in renderGroups)
        {
            EnsureProfiledImplicit(
                plans,
                DeliveryArtifactKind.PageMap,
                group.Key,
                group.Requiredness,
                $"page-map-{Kebab(group.Key.ReviewProfile)}-{Kebab(group.Key.CommentProfile)}");
            EnsureProfiledImplicit(
                plans,
                DeliveryArtifactKind.RenderReport,
                group.Key,
                group.Requiredness,
                $"render-report-{Kebab(group.Key.ReviewProfile)}-{Kebab(group.Key.CommentProfile)}");
        }
        return plans;
    }

    private static IReadOnlyList<DeliveryArtifactRequest> ValidateRequests(
        IReadOnlyList<DeliveryArtifactRequest> requested)
    {
        var snapshot = requested.ToArray();
        if (snapshot.Any(value => value is null))
            throw new ArgumentException("Artifact requests cannot contain null entries.");
        if (snapshot.GroupBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Any(group => group.Count() != 1))
            throw new ArgumentException("Artifact request IDs must be unique.");
        if (snapshot.Where(value => !DeliveryBundleManifest.IsProfiledRenderKind(value.Kind))
            .GroupBy(value => value.Kind).Any(group => group.Count() != 1))
            throw new ArgumentException("Non-render artifact kinds can be requested only once.");
        foreach (var value in snapshot)
        {
            if (string.IsNullOrWhiteSpace(value.ArtifactId)
                || !Enum.IsDefined(value.Kind)
                || !Enum.IsDefined(value.Requiredness))
                throw new ArgumentException("Artifact request identity is invalid.");
            if (DeliveryBundleManifest.IsProfiledRenderKind(value.Kind))
            {
                if (value.ReviewProfile is null || !Enum.IsDefined(value.ReviewProfile.Value)
                    || value.CommentProfile is null || !Enum.IsDefined(value.CommentProfile.Value))
                    throw new ArgumentException(
                        $"Render artifact '{value.ArtifactId}' requires explicit review and comment profiles.");
                if (value.Kind == DeliveryArtifactKind.FinalPdf
                    && value.ReviewProfile != DeliveryReviewProfile.Final)
                    throw new ArgumentException(
                        $"Final PDF artifact '{value.ArtifactId}' requires the final review profile.");
                if (value.Kind == DeliveryArtifactKind.ReviewPdf
                    && value.ReviewProfile != DeliveryReviewProfile.Markup)
                    throw new ArgumentException(
                        $"Review PDF artifact '{value.ArtifactId}' requires the markup review profile.");
            }
            else if (value.ReviewProfile is not null || value.CommentProfile is not null)
            {
                throw new ArgumentException(
                    $"Non-render artifact '{value.ArtifactId}' cannot select render profiles.");
            }
        }
        return snapshot;
    }

    private static void EnsureImplicit(
        List<ArtifactPlan> plans,
        DeliveryArtifactKind kind,
        DeliveryArtifactRequiredness requiredness,
        string preferredId)
    {
        if (plans.Any(plan => plan.Request.Kind == kind))
            return;
        plans.Add(new ArtifactPlan(new DeliveryArtifactRequest
        {
            ArtifactId = ReserveId(plans.Select(value => value.Request.ArtifactId), preferredId),
            Kind = kind,
            Requiredness = requiredness,
        }, true));
    }

    private static void EnsureProfiledImplicit(
        List<ArtifactPlan> plans,
        DeliveryArtifactKind kind,
        RenderProfileKey profile,
        DeliveryArtifactRequiredness requiredness,
        string preferredId)
    {
        var existingIndex = plans.FindIndex(plan => plan.Request.Kind == kind
            && plan.Request.ReviewProfile == profile.ReviewProfile
            && plan.Request.CommentProfile == profile.CommentProfile);
        if (existingIndex >= 0)
        {
            var existing = plans[existingIndex];
            if (requiredness == DeliveryArtifactRequiredness.Required
                && existing.Request.Requiredness == DeliveryArtifactRequiredness.Optional)
            {
                plans[existingIndex] = existing with
                {
                    Request = existing.Request with
                    {
                        Requiredness = DeliveryArtifactRequiredness.Required,
                    },
                };
            }
            return;
        }
        plans.Add(new ArtifactPlan(new DeliveryArtifactRequest
        {
            ArtifactId = ReserveId(plans.Select(value => value.Request.ArtifactId), preferredId),
            Kind = kind,
            Requiredness = requiredness,
            ReviewProfile = profile.ReviewProfile,
            CommentProfile = profile.CommentProfile,
        }, true));
    }

    private static IReadOnlyList<DeliveryArtifactRelationship> BuildRelationships(
        IReadOnlyList<ArtifactPlan> plans,
        IEnumerable<DeliveryBundleArtifactInput> outputs,
        IReadOnlyList<RenderState> renderStates)
    {
        var availableIds = outputs.Select(value => value.ArtifactId)
            .ToHashSet(StringComparer.Ordinal);
        var relationships = new List<DeliveryArtifactRelationship>();
        void Add(
            ArtifactPlan? from,
            DeliveryArtifactRelationshipKind kind,
            ArtifactPlan? to)
        {
            if (from is null || to is null
                || !availableIds.Contains(from.Request.ArtifactId)
                || !availableIds.Contains(to.Request.ArtifactId))
                return;
            relationships.Add(new DeliveryArtifactRelationship
            {
                RelationshipId = $"rel-{relationships.Count + 1:D4}",
                Kind = kind,
                FromArtifactId = from.Request.ArtifactId,
                ToArtifactId = to.Request.ArtifactId,
            });
        }

        ArtifactPlan? First(DeliveryArtifactKind kind) => plans.FirstOrDefault(plan =>
            plan.Request.Kind == kind);
        var baseline = First(DeliveryArtifactKind.BaselineDocx);
        var working = First(DeliveryArtifactKind.WorkingDocx);
        var policyBaseline = First(DeliveryArtifactKind.PolicyBaselineDocx);
        var review = First(DeliveryArtifactKind.ReviewDocx);
        var final = First(DeliveryArtifactKind.FinalDocx);
        Add(working, DeliveryArtifactRelationshipKind.DerivedFrom, baseline);
        Add(policyBaseline, DeliveryArtifactRelationshipKind.DerivedFrom, baseline);
        Add(final, DeliveryArtifactRelationshipKind.DerivedFrom, working);
        Add(review, DeliveryArtifactRelationshipKind.DerivedFrom, policyBaseline);
        Add(review, DeliveryArtifactRelationshipKind.Describes, final);
        Add(First(DeliveryArtifactKind.BaselinePackageManifest),
            DeliveryArtifactRelationshipKind.Describes, baseline);
        Add(First(DeliveryArtifactKind.FinalPackageManifest),
            DeliveryArtifactRelationshipKind.Describes, final);
        Add(First(DeliveryArtifactKind.SemanticDelta),
            DeliveryArtifactRelationshipKind.Describes, final);
        Add(First(DeliveryArtifactKind.PackageDelta),
            DeliveryArtifactRelationshipKind.Describes, final);
        Add(First(DeliveryArtifactKind.ValidationReport),
            DeliveryArtifactRelationshipKind.Validates, final);
        if (plans.Any(plan =>
                plan.Request.ReviewProfile == DeliveryReviewProfile.Original))
            Add(First(DeliveryArtifactKind.ValidationReport),
                DeliveryArtifactRelationshipKind.Validates, policyBaseline);
        if (plans.Any(plan =>
                plan.Request.ReviewProfile == DeliveryReviewProfile.Markup))
            Add(First(DeliveryArtifactKind.ValidationReport),
                DeliveryArtifactRelationshipKind.Validates, review);
        Add(First(DeliveryArtifactKind.ReversibilityProof),
            DeliveryArtifactRelationshipKind.Proves, review);
        Add(First(DeliveryArtifactKind.ChangeReceipt),
            DeliveryArtifactRelationshipKind.ReceiptFor, final);
        foreach (var render in plans.Where(plan =>
                     DeliveryBundleManifest.IsProfiledRenderKind(plan.Request.Kind))
                     .OrderBy(plan => plan.Request.ArtifactId, StringComparer.Ordinal))
        {
            var source = render.Request.ReviewProfile switch
            {
                DeliveryReviewProfile.Final => final,
                DeliveryReviewProfile.Original => policyBaseline,
                DeliveryReviewProfile.Markup => review,
                _ => null,
            };
            Add(render, DeliveryArtifactRelationshipKind.RenderedFrom, source);
        }
        var pageMaps = renderStates.Where(state =>
                state.Plan.Request.Kind == DeliveryArtifactKind.PageMap
                && state.Output.Availability == DeliveryArtifactAvailability.Available)
            .Select(state => new
            {
                State = state,
                Bytes = state.Output.CopyBytes(),
            })
            .Where(candidate => candidate.Bytes is not null)
            .Select(candidate => new
            {
                candidate.State,
                Digest = DeliveryBundleCanonicalJson.Digest(candidate.Bytes!),
            })
            .ToArray();
        var layouts = renderStates.Where(state =>
                (state.Plan.Request.Kind is DeliveryArtifactKind.StandaloneHtml
                    or DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf)
                && state.Output.Availability == DeliveryArtifactAvailability.Available)
            .Select(state => new
            {
                State = state,
                PageMapBytes = state.Result?.CopyPageMapBytes(),
            })
            .Where(layout => layout.PageMapBytes is not null)
            .OrderBy(layout => layout.State.Plan.Request.ArtifactId, StringComparer.Ordinal);
        foreach (var layout in layouts)
        {
            var digest = DeliveryBundleCanonicalJson.Digest(
                layout.PageMapBytes!);
            foreach (var pageMap in pageMaps.Where(candidate =>
                         candidate.Digest == digest
                         && candidate.State.SourceDigest == layout.State.SourceDigest
                         && string.Equals(
                             candidate.State.Result?.RendererFingerprint,
                             layout.State.Result!.RendererFingerprint,
                             StringComparison.Ordinal)))
            {
                Add(layout.State.Plan, DeliveryArtifactRelationshipKind.UsesPageMap,
                    pageMap.State.Plan);
            }
        }
        return relationships;
    }

    private static DeliveryDocumentSnapshot RenderSource(
        DeliveryReviewProfile profile,
        DeliveryDocumentSnapshot policyBaseline,
        DeliveryDocumentSnapshot final,
        DeliveryDocumentSnapshot? review) => profile switch
        {
            DeliveryReviewProfile.Final => final,
            DeliveryReviewProfile.Original => policyBaseline,
            DeliveryReviewProfile.Markup => review
                ?? throw new DeliveryBundleException(
                    "review_proof_unavailable", "Markup rendering requires a proven review DOCX."),
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, null),
        };

    private static DeliveryBundleArtifactInput Available(ArtifactPlan plan, byte[] bytes) =>
        DeliveryBundleArtifactInput.Available(
            plan.Request.ArtifactId,
            plan.Request.Kind,
            RelativePath(plan),
            MediaType(plan.Request.Kind),
            bytes,
            plan.IsImplicit,
            plan.Request.Requiredness);

    private static DeliveryBundleArtifactInput Unavailable(ArtifactPlan plan, string reason) =>
        DeliveryBundleArtifactInput.Unavailable(
            plan.Request.ArtifactId,
            plan.Request.Kind,
            RelativePath(plan),
            MediaType(plan.Request.Kind),
            reason,
            plan.IsImplicit,
            plan.Request.Requiredness);

    private static DeliveryBundleArtifactInput UnavailableRender(
        RenderJob job,
        string reason) =>
        DeliveryBundleArtifactInput.Unavailable(
            job.Plan.Request.ArtifactId,
            job.Plan.Request.Kind,
            RelativePath(job.Plan),
            MediaType(job.Plan.Request.Kind),
            reason,
            job.Plan.IsImplicit,
            job.Plan.Request.Requiredness,
            RenderMetadata(job));

    private static string RelativePath(ArtifactPlan plan)
    {
        var stem = Slug(plan.Request.ArtifactId);
        var extension = plan.Request.Kind switch
        {
            DeliveryArtifactKind.BaselineDocx or DeliveryArtifactKind.PolicyBaselineDocx
                or DeliveryArtifactKind.WorkingDocx or DeliveryArtifactKind.ReviewDocx
                or DeliveryArtifactKind.FinalDocx => ".docx",
            DeliveryArtifactKind.StandaloneHtml => ".html",
            DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf => ".pdf",
            _ => ".json",
        };
        var directory = plan.Request.Kind switch
        {
            DeliveryArtifactKind.BaselineDocx or DeliveryArtifactKind.PolicyBaselineDocx
                or DeliveryArtifactKind.WorkingDocx or DeliveryArtifactKind.ReviewDocx
                or DeliveryArtifactKind.FinalDocx => "documents",
            DeliveryArtifactKind.StandaloneHtml or DeliveryArtifactKind.FinalPdf
                or DeliveryArtifactKind.ReviewPdf or DeliveryArtifactKind.PageMap
                or DeliveryArtifactKind.RenderReport => "renders",
            _ => "evidence",
        };
        return $"{directory}/{stem}{extension}";
    }

    private static string Slug(string artifactId)
    {
        var normalized = new string(artifactId.Select(value =>
                char.IsAsciiLetterOrDigit(value) || value is '-' or '_' or '.' ? value : '-')
            .ToArray()).Trim('-', '.', ' ');
        if (normalized.Length > 48)
            normalized = normalized[..48];
        if (normalized.Length == 0)
            normalized = "artifact";
        var suffix = Convert.ToHexString(
                SHA256.HashData(Encoding.UTF8.GetBytes(artifactId)))
            .ToLowerInvariant()[..8];
        return $"{normalized}-{suffix}";
    }

    private static string Kebab<T>(T value)
        where T : struct, Enum
    {
        var name = value.ToString();
        var builder = new StringBuilder(name.Length + 8);
        for (var index = 0; index < name.Length; index++)
        {
            var character = name[index];
            if (index > 0 && char.IsUpper(character))
                builder.Append('-');
            builder.Append(char.ToLowerInvariant(character));
        }
        return builder.ToString();
    }

    private static string MediaType(DeliveryArtifactKind kind) => kind switch
    {
        DeliveryArtifactKind.BaselineDocx or DeliveryArtifactKind.PolicyBaselineDocx
            or DeliveryArtifactKind.WorkingDocx or DeliveryArtifactKind.ReviewDocx
            or DeliveryArtifactKind.FinalDocx => DocxMediaType,
        DeliveryArtifactKind.StandaloneHtml => "text/html",
        DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf => "application/pdf",
        _ => "application/json",
    };

    private static string ReserveId(IEnumerable<string> existingIds, string preferred)
    {
        var existing = existingIds.ToHashSet(StringComparer.Ordinal);
        if (!existing.Contains(preferred))
            return preferred;
        for (var suffix = 2; suffix < 10_000; suffix++)
        {
            var candidate = $"{preferred}-{suffix}";
            if (!existing.Contains(candidate))
                return candidate;
        }
        throw new InvalidOperationException("Could not allocate a stable implicit artifact ID.");
    }

    private static bool IsDeferred(DeliveryArtifactKind kind) =>
        DeliveryBundleManifest.IsProfiledRenderKind(kind)
        || kind is DeliveryArtifactKind.ValidationReport or DeliveryArtifactKind.ChangeReceipt;

    private static void EnforceRequiredAvailability(
        IEnumerable<DeliveryBundleArtifactInput> outputs,
        bool returnIncompleteBundle)
    {
        var missing = outputs.FirstOrDefault(value =>
            value.Availability == DeliveryArtifactAvailability.Unavailable
            && value.ImplicitRequiredness == DeliveryArtifactRequiredness.Required);
        if (missing is not null && !returnIncompleteBundle)
        {
            throw new DeliveryBundleException(
                "required_artifact_unavailable",
                $"Required artifact '{missing.ArtifactId}' is unavailable: {missing.UnavailableReason}");
        }
    }

    private static void EnsureValidManifest(string name, PackageManifest manifest)
    {
        if (!manifest.IsValid)
            throw new DeliveryBundleException(
                "invalid_package_manifest",
                $"The {name} package failed bounded manifest validation.");
    }

    private static void ValidateOptions(DeliveryBundleBuildOptions options)
    {
        ArgumentNullException.ThrowIfNull(options.PackageManifestOptions);
        ArgumentNullException.ThrowIfNull(options.DeliverableVerificationOptions);
        ArgumentNullException.ThrowIfNull(options.DeliveryReceiptLimits);
        ArgumentNullException.ThrowIfNull(options.BundleVerificationLimits);
        options.PackageManifestOptions.Validate();
        options.DeliverableVerificationOptions.Validate();
        options.BundleVerificationLimits.Validate();
    }

    private sealed record ArtifactPlan(DeliveryArtifactRequest Request, bool IsImplicit);

    private sealed record RenderProfileKey(
        DeliveryReviewProfile ReviewProfile,
        DeliveryCommentProfile CommentProfile);

    private sealed record RenderJob(
        ArtifactPlan Plan,
        DeliveryRenderRequest Request,
        VerificationDigest SourceDigest);

    private sealed record RenderState(
        ArtifactPlan Plan,
        DeliveryBundleArtifactInput Output,
        VerificationDigest SourceDigest,
        DeliveryRenderResult? Result);
}
