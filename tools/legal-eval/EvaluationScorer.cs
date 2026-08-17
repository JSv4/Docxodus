// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Verification;

namespace LegalEval;

public sealed class EvaluationScorer
{
    private const long MaximumInvariantPartBytes = 32L * 1024 * 1024;
    private const int MaximumProofXmlParts = 256;
    private const long MaximumProofXmlBytes = 64L * 1024 * 1024;
    private const int MaximumSemanticChanges = 4096;
    private const string WNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly IEvaluationPackageValidator _packageValidator;
    private readonly IEvaluationArtifactRenderer? _artifactRenderer;

    public EvaluationScorer(
        IEvaluationPackageValidator? packageValidator = null,
        IEvaluationArtifactRenderer? artifactRenderer = null)
    {
        _packageValidator = packageValidator ?? new EvaluationPackageValidator();
        _artifactRenderer = artifactRenderer;
    }

    public EvaluationScore Score(
        LegalScenario scenario,
        BaselineExecution baseline,
        byte[] candidate,
        ScoreKind kind,
        string artifactRoot,
        ArtifactRenderMode renderMode = ArtifactRenderMode.Disabled)
    {
        var metrics = new List<MetricResult>();
        string? packageSafetyError = null;
        PackageManifest? inputManifest = baseline.InputManifest;
        PackageManifest? expectedManifest = baseline.ExpectedManifest;
        PackageManifest? candidateManifest = null;
        SemanticChangeSet? candidateChanges = null;
        SemanticChangeSet? targetChanges = null;
        SemanticChangeSet? candidateTargetChanges = null;
        DeliverableVerificationResult? verification = null;
        RedlineReversibilityProofRun? redlineProof = null;
        byte[]? redlineBytes = null;
        DeliveryChangeReceipt? deliveryReceipt = null;
        IReadOnlyDictionary<string, byte[]> deliveryReceiptArtifacts =
            new Dictionary<string, byte[]>(StringComparer.Ordinal);
        string? deliveryReceiptUnavailableReason = null;
        try
        {
            inputManifest ??= _packageValidator.Inspect(baseline.Input, "evaluation input");
            expectedManifest ??= _packageValidator.Inspect(
                baseline.Expected, "pinned expected document");
            candidateManifest = _packageValidator.Inspect(candidate, "candidate document");
            metrics.Add(new MetricResult("document-validity.package-safety", "document_validity",
                "passed", "candidate passed the shared #456 bounded package inspection", 1));
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            packageSafetyError = exception.Message;
            metrics.Add(new MetricResult("document-validity.package-safety", "document_validity",
                "failed", exception.Message, 0));
        }
        (MetricResult Metric, string? CandidateHtml) render;
        if (packageSafetyError is null)
        {
            candidateChanges = CompareSemantic(baseline.Input, candidate);
            targetChanges = CompareSemantic(baseline.Input, baseline.Expected);
            candidateTargetChanges = CompareSemantic(candidate, baseline.Expected);
            verification = VerifyCandidate(
                baseline.Input, candidate, baseline.Expected);
            metrics.AddRange(EvaluateInvariants(scenario, candidate));
            metrics.Add(EvaluateTargetPrecision(candidateTargetChanges));
            metrics.AddRange(SafeMetrics(
                new[]
                {
                    ("unintended-change.parts", "unintended_change"),
                    ("unintended-change.relationships", "unintended_change"),
                    ("unintended-change.anchors", "unintended_change"),
                },
                () => EvaluateChangeBudget(scenario, verification, candidateChanges)));
            metrics.Add(EvaluateValidity(verification));
            (redlineBytes, redlineProof) = BuildRedlineProof(baseline.Input, candidate);
            metrics.Add(EvaluateRedlineReversibility(scenario, redlineProof));
            render = EvaluateRendering(baseline.Expected, candidate);
            metrics.Add(render.Metric);

            var receipt = BuildDeliveryReceipt(
                baseline, candidate, kind, inputManifest!, candidateManifest!, candidateChanges,
                verification, redlineProof);
            deliveryReceipt = receipt.Receipt;
            deliveryReceiptArtifacts = receipt.Artifacts;
            deliveryReceiptUnavailableReason = receipt.UnavailableReason;
        }
        else
        {
            metrics.AddRange(BlockedMetrics(scenario, packageSafetyError));
            metrics.Add(new MetricResult("redline-reversibility.full-proof",
                "redline_reversibility", "failed",
                $"blocked by package safety validation: {packageSafetyError}", 0));
            render = (new MetricResult("rendering-regression.html-projection",
                "rendering_regression", "failed",
                $"blocked by package safety validation: {packageSafetyError}", 0), null);
            metrics.Add(render.Metric);
            deliveryReceiptUnavailableReason =
                $"blocked by package safety validation: {packageSafetyError}";
        }

        var status = metrics.Any(value => value.Status == "failed") ? "failed" : "passed";
        var artifactDirectory = ArtifactWriter.ResolveScoreDirectory(artifactRoot, scenario.Id, kind);
        var publication = ArtifactWriter.Write(
            artifactDirectory,
            scenario.Id,
            kind,
            status,
            metrics,
            scenario.ExpectedOutputs,
            PublishedOperationLog(kind, baseline.OperationLog),
            new EvaluationEvidence
            {
                ScenarioContractJson = ScenarioContract(scenario),
                Input = baseline.Input,
                Candidate = candidate,
                Expected = baseline.Expected,
                InputManifest = inputManifest,
                CandidateManifest = candidateManifest,
                ExpectedManifest = expectedManifest,
                CandidateChanges = candidateChanges,
                TargetChanges = targetChanges,
                CandidateTargetChanges = candidateTargetChanges,
                DeliverableVerification = verification,
                RedlineBytes = redlineBytes,
                RedlineProofRun = redlineProof,
                DeliveryReceipt = deliveryReceipt,
                DeliveryReceiptArtifacts = deliveryReceiptArtifacts,
                DeliveryReceiptUnavailableReason = deliveryReceiptUnavailableReason,
                CandidateHtml = render.CandidateHtml,
                FailureReason = packageSafetyError,
                CandidateSafetyError = packageSafetyError,
            },
            renderMode,
            allowExternalRenderer: kind == ScoreKind.EngineBaseline,
            _artifactRenderer);

        return new EvaluationScore(scenario.Id, kind, publication.Status,
            publication.Metrics, publication.Artifacts, artifactDirectory);
    }

    private static IReadOnlyList<MetricResult> BlockedMetrics(
        LegalScenario scenario, string packageSafetyError)
    {
        var detail = $"blocked by package safety validation: {packageSafetyError}";
        var results = scenario.Invariants.Select(value =>
                new MetricResult(value.Id, value.Metric, "failed", detail, 0))
            .ToList();
        results.Add(new MetricResult("target-precision.reference-equivalence", "target_precision",
            "failed", detail, 0));
        results.Add(new MetricResult("unintended-change.parts", "unintended_change",
            "failed", detail, 0));
        results.Add(new MetricResult("unintended-change.relationships", "unintended_change",
            "failed", detail, 0));
        results.Add(new MetricResult("unintended-change.anchors", "unintended_change",
            "failed", detail, 0));
        results.Add(new MetricResult("document-validity.openxml", "document_validity",
            "failed", detail, 0));
        return results;
    }

    private static MetricResult SafeMetric(
        string id, string category, Func<MetricResult> evaluate)
    {
        try
        {
            return evaluate();
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return new MetricResult(id, category, "failed",
                $"evaluation error: {exception.Message}", 0);
        }
    }

    private static IReadOnlyList<MetricResult> SafeMetrics(
        IReadOnlyList<(string Id, string Category)> expected,
        Func<IReadOnlyList<MetricResult>> evaluate)
    {
        try
        {
            return evaluate();
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return expected.Select(value => new MetricResult(
                value.Id, value.Category, "failed",
                $"evaluation error: {exception.Message}", 0)).ToList();
        }
    }

    private static IReadOnlyList<MetricResult> EvaluateInvariants(
        LegalScenario scenario, byte[] candidate)
    {
        var results = new List<MetricResult>(scenario.Invariants.Count);
        foreach (var invariant in scenario.Invariants)
        {
            try
            {
                var observed = Probe(candidate, invariant.Probe);
                var passed = Compare(observed, invariant.Operator, invariant.Expected);
                results.Add(new MetricResult(
                    invariant.Id,
                    invariant.Metric,
                    passed ? "passed" : "failed",
                    $"operator={invariant.Operator}; expected={Json(invariant.Expected)}; observed={Json(observed)}",
                    passed ? 1 : 0));
            }
            catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
            {
                results.Add(new MetricResult(invariant.Id, invariant.Metric, "failed",
                    $"probe error: {exception.Message}", 0));
            }
        }
        return results;
    }

    private static JsonNode Probe(byte[] candidate, JsonObject probe)
    {
        var kind = RequiredString(probe, "kind");
        return kind switch
        {
            "xmlCount" => JsonValue.Create(XmlCount(candidate,
                RequiredString(probe, "part"), RequiredString(probe, "xpath")))!,
            "partExists" => JsonValue.Create(PartExists(candidate,
                RequiredString(probe, "part")))!,
            "textCount" => JsonValue.Create(TextCount(candidate,
                RequiredString(probe, "text")))!,
            _ => throw new ScenarioValidationException($"unknown invariant probe kind '{kind}'"),
        };
    }

    private static int XmlCount(byte[] bytes, string part, string xpath)
    {
        var payload = PartBytes(bytes, part, required: true)!;
        using var reader = XmlReader.Create(new MemoryStream(payload), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        });
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        var namespaces = new XmlNamespaceManager(new NameTable());
        namespaces.AddNamespace("w", WNamespace);
        namespaces.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        namespaces.AddNamespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
        namespaces.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        namespaces.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");
        namespaces.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships");
        var value = document.XPathEvaluate(xpath, namespaces);
        if (value is IEnumerable<object> objects) return objects.Count();
        if (value is IEnumerable<XElement> elements) return elements.Count();
        if (value is IEnumerable<XAttribute> attributes) return attributes.Count();
        if (value is double number) return checked((int)number);
        throw new InvalidOperationException($"XPath must return a node-set or number, got {value?.GetType().Name}");
    }

    private static int TextCount(byte[] bytes, string needle)
    {
        var projection = WmlToMarkdownConverter.Convert(
            new WmlDocument("candidate.docx", bytes),
            new WmlToMarkdownConverterSettings
            {
                AnchorMode = AnchorRenderMode.None,
                Scopes = ProjectionScopes.All,
                TrackedChanges = TrackedChangeMode.RenderInline,
            });
        var count = 0;
        var cursor = 0;
        while ((cursor = projection.Markdown.IndexOf(needle, cursor, StringComparison.Ordinal)) >= 0)
        {
            count++;
            cursor += needle.Length;
        }
        return count;
    }

    private static bool Compare(JsonNode observed, string operation, JsonNode? expected) =>
        operation switch
        {
            "equals" => Json(observed) == Json(expected),
            "atLeast" => Number(observed) >= Number(expected),
            "contains" => observed.GetValue<string>().Contains(
                expected?.GetValue<string>() ?? string.Empty, StringComparison.Ordinal),
            "setEquals" => Set(observed).SetEquals(Set(expected)),
            _ => throw new ScenarioValidationException($"unknown invariant operator '{operation}'"),
        };

    private static MetricResult EvaluateTargetPrecision(SemanticChangeSet candidateTargetChanges)
    {
        var passed = candidateTargetChanges.Changes.Count == 0;
        return new MetricResult("target-precision.reference-equivalence", "target_precision",
            passed ? "passed" : "failed",
            passed
                ? "candidate is #457 semantic-change equivalent to the scripted expected output"
                : $"candidate differs from scripted expected output in {candidateTargetChanges.Changes.Count} semantic changes",
            passed ? 1 : 0);
    }

    private static IReadOnlyList<MetricResult> EvaluateChangeBudget(
        LegalScenario scenario,
        DeliverableVerificationResult verification,
        SemanticChangeSet candidateChanges)
    {
        var packageDelta = verification.Checks.SingleOrDefault(
            check => check.Check == "package_delta");
        if (packageDelta?.Status != DeliverableCheckStatus.Completed)
        {
            throw new ScenarioValidationException(
                "#463 package-delta analysis did not complete: "
                + (packageDelta?.Diagnostic ?? "package_delta evidence is absent"));
        }
        var changedParts = verification.PackageChanges
            .Where(change => change.Kind is DeliverablePackageChangeKind.EntryAdded
                or DeliverablePackageChangeKind.EntryRemoved
                or DeliverablePackageChangeKind.EntryModified)
            .Select(change => change.Location.EntryUri
                ?? throw new ScenarioValidationException(
                    $"#463 entry change '{change.ChangeId}' has no entry URI"))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();
        var unexpectedParts = changedParts
            .Where(value => !scenario.ChangeBudget.AllowedChangedParts.Contains(value))
            .OrderBy(value => value, StringComparer.Ordinal).ToList();
        var partPass = unexpectedParts.Count == 0;

        var relationshipOwners = verification.PackageChanges
            .Where(change => change.Kind is DeliverablePackageChangeKind.RelationshipAdded
                or DeliverablePackageChangeKind.RelationshipRemoved
                or DeliverablePackageChangeKind.RelationshipModified)
            .Select(change => change.Location.OwnerUri
                ?? throw new ScenarioValidationException(
                    $"#463 relationship change '{change.ChangeId}' has no owner URI"))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();
        var unexpectedRelationshipOwners = relationshipOwners
            .Where(value => !scenario.ChangeBudget.AllowedRelationshipOwners.Contains(value))
            .ToList();
        var relationshipPass = unexpectedRelationshipOwners.Count == 0;

        var anchors = candidateChanges.Changes
            .SelectMany(value => new[] { value.LeftAnchor, value.RightAnchor })
            .Where(value => value is not null).Cast<string>()
            // #457 exposes both paragraph (p:...) and inline (li:...) views of the same stable
            // scope/hash. Count that logical anchor once while still counting before and after
            // identities separately. Otherwise every ordinary text replacement consumes four
            // budget slots and the corpus's declared logical-anchor budgets cannot pass.
            .Select(NormalizeSemanticAnchor)
            .Distinct(StringComparer.Ordinal).ToList();
        var anchorPass = anchors.Count <= scenario.ChangeBudget.MaximumChangedAnchors;

        return new[]
        {
            new MetricResult("unintended-change.parts", "unintended_change",
                partPass ? "passed" : "failed",
                partPass
                    ? $"changed parts are within budget: {string.Join(", ", changedParts)}"
                    : $"unexpected changed parts: {string.Join(", ", unexpectedParts)}",
                partPass ? 1 : 0),
            new MetricResult("unintended-change.relationships", "unintended_change",
                relationshipPass ? "passed" : "failed",
                relationshipPass
                    ? $"relationship owners are within budget: {string.Join(", ", relationshipOwners)}"
                    : $"unexpected relationship owners: {string.Join(", ", unexpectedRelationshipOwners)}",
                relationshipPass ? 1 : 0),
            new MetricResult("unintended-change.anchors", "unintended_change",
                anchorPass ? "passed" : "failed",
                $"distinct #457 semantic anchors={anchors.Count}; maximum={scenario.ChangeBudget.MaximumChangedAnchors}",
                anchorPass ? 1 : 0),
        };
    }

    private static string NormalizeSemanticAnchor(string anchor)
    {
        var separator = anchor.IndexOf(':');
        return separator < 0 ? anchor : anchor[(separator + 1)..];
    }

    private static MetricResult EvaluateValidity(DeliverableVerificationResult verification)
    {
        var requiredChecks = new HashSet<string>(StringComparer.Ordinal)
        {
            "deliverable.package_manifest",
            "deliverable.open_xml",
            "deliverable.wordprocessing_closure",
            "deliverable.workflow_and_revision_registry",
        };
        var incomplete = verification.Checks
            .Where(check => requiredChecks.Contains(check.Check)
                && check.Status != DeliverableCheckStatus.Completed)
            .ToList();
        // The verifier was configured with the pinned target's exact semantic and package
        // deltas. A blocking delta therefore means the candidate is not the declared delivery,
        // even when its broader scenario change budget happens to permit the affected part.
        var blocking = verification.Findings.Where(finding => finding.BlocksDelivery).ToList();
        var passed = incomplete.Count == 0 && blocking.Count == 0
            && verification.Decision is DeliverableVerificationDecision.Passed
                or DeliverableVerificationDecision.PassedWithPreExistingFindings;
        return new MetricResult("document-validity.openxml", "document_validity",
            passed ? "passed" : "failed",
            passed
                ? $"#463 bounded package/Open XML validation completed; decision={verification.Decision}"
                : string.Join(" | ", incomplete.Select(check =>
                        $"{check.Check}: {check.Status} {check.Diagnostic}")
                    .Concat(blocking.Select(finding => $"{finding.Code}: {finding.Message}"))
                    .Take(5)),
            passed ? 1 : 0);
    }

    private static MetricResult EvaluateRedlineReversibility(
        LegalScenario scenario,
        RedlineReversibilityProofRun proofRun)
    {
        var proof = proofRun.Proof;
        if (scenario.RedlineReversibility.Applicability
            == RedlineReversibilityApplicability.NotApplicable)
        {
            return new MetricResult("redline-reversibility.not-applicable",
                "redline_reversibility", "not_applicable",
                scenario.RedlineReversibility.Reason
                    + "; the diagnostic #464 proof is retained but is not scored",
                null);
        }
        return new MetricResult("redline-reversibility.full-proof",
            "redline_reversibility", proof.Success ? "passed" : "failed",
            proof.Success
                ? "#464 full reversibility proof succeeded and preserved pre-existing revisions on both paths"
                : "#464 proof failed: " + string.Join(", ", proof.Findings
                    .Concat(proof.AcceptToFinal?.Findings
                        ?? Array.Empty<RedlineProofFinding>())
                    .Concat(proof.RejectToBaseline?.Findings
                        ?? Array.Empty<RedlineProofFinding>())
                    .Select(finding => finding.Code).Distinct(StringComparer.Ordinal).Take(8)),
            proof.Success ? 1 : 0);
    }

    private static (MetricResult Metric, string? CandidateHtml) EvaluateRendering(
        byte[] expected, byte[] candidate)
    {
        try
        {
            var expectedHtml = RenderHtml(expected);
            var candidateHtml = RenderHtml(candidate);
            var passed = NormalizeHtml(expectedHtml) == NormalizeHtml(candidateHtml);
            return (new MetricResult("rendering-regression.html-projection", "rendering_regression",
                    passed ? "passed" : "failed",
                    passed
                        ? "sanitized Docxodus HTML projection matches the pinned expected output; this is not a visual-layout proof"
                        : "sanitized Docxodus HTML projection differs from the pinned expected output; this is not a visual-layout proof",
                    passed ? 1 : 0),
                candidateHtml);
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return (new MetricResult("rendering-regression.html-projection", "rendering_regression",
                "failed", exception.Message, 0), null);
        }
    }

    private SemanticChangeSet CompareSemantic(
        byte[] before, byte[] after, bool includePackageChanges = true) =>
        SemanticDiff.CompareBounded(
            new WmlDocument("before.docx", before),
            new WmlDocument("after.docx", after),
            new SemanticDiffOptions
            {
                IncludePackageChanges = includePackageChanges,
                PackageOptions = _packageValidator.ManifestOptions,
            },
            MaximumSemanticChanges);

    internal EvaluationEvidence AnalyzeAvailableEvidence(
        BaselineExecution baseline,
        byte[]? candidate,
        ScoreKind kind,
        string reason,
        string? inputSafetyError,
        string? candidateSafetyError,
        string? expectedSafetyError)
    {
        PackageManifest? inputManifest = baseline.InputManifest;
        PackageManifest? candidateManifest = null;
        PackageManifest? expectedManifest = baseline.ExpectedManifest;
        SemanticChangeSet? candidateChanges = null;
        SemanticChangeSet? targetChanges = null;
        SemanticChangeSet? candidateTargetChanges = null;
        DeliverableVerificationResult? verification = null;
        byte[]? redline = null;
        RedlineReversibilityProofRun? proof = null;
        DeliveryChangeReceipt? receipt = null;
        IReadOnlyDictionary<string, byte[]> receiptArtifacts =
            new Dictionary<string, byte[]>(StringComparer.Ordinal);
        string? receiptReason = baseline.DeliveryReceiptUnavailableReason;

        try
        {
            if (inputSafetyError is null)
                inputManifest ??= _packageValidator.Inspect(baseline.Input, "evaluation input");
            if (expectedSafetyError is null)
                expectedManifest ??= _packageValidator.Inspect(
                    baseline.Expected, "pinned expected document");
            if (candidate is not null && candidateSafetyError is null)
                candidateManifest = _packageValidator.Inspect(candidate, "candidate document");
            if (inputSafetyError is null && expectedSafetyError is null)
                targetChanges = CompareSemantic(baseline.Input, baseline.Expected);
            if (candidate is not null && inputSafetyError is null && candidateSafetyError is null)
                candidateChanges = CompareSemantic(baseline.Input, candidate);
            if (candidate is not null && candidateSafetyError is null && expectedSafetyError is null)
                candidateTargetChanges = CompareSemantic(candidate, baseline.Expected);
            if (candidate is not null && targetChanges is not null
                && inputSafetyError is null && candidateSafetyError is null)
                verification = VerifyCandidate(
                    baseline.Input, candidate, baseline.Expected);
            if (candidate is not null && inputSafetyError is null && candidateSafetyError is null)
                (redline, proof) = BuildRedlineProof(baseline.Input, candidate);
            if (candidate is not null && inputManifest is not null && candidateManifest is not null
                && candidateChanges is not null && verification is not null && proof is not null)
            {
                var delivery = BuildDeliveryReceipt(baseline, candidate, kind,
                    inputManifest, candidateManifest, candidateChanges, verification, proof);
                receipt = delivery.Receipt;
                receiptArtifacts = delivery.Artifacts;
                receiptReason = delivery.UnavailableReason;
            }
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            receiptReason ??= $"evidence analysis stopped after {exception.GetType().Name}: {exception.Message}";
        }

        return new EvaluationEvidence
        {
            Input = baseline.Input,
            Candidate = candidate,
            Expected = baseline.Expected,
            InputManifest = inputManifest,
            CandidateManifest = candidateManifest,
            ExpectedManifest = expectedManifest,
            CandidateChanges = candidateChanges,
            TargetChanges = targetChanges,
            CandidateTargetChanges = candidateTargetChanges,
            DeliverableVerification = verification,
            RedlineBytes = redline,
            RedlineProofRun = proof,
            DeliveryReceipt = receipt,
            DeliveryReceiptArtifacts = receiptArtifacts,
            DeliveryReceiptUnavailableReason = receiptReason,
            FailureReason = reason,
            InputSafetyError = inputSafetyError,
            CandidateSafetyError = candidateSafetyError,
            ExpectedSafetyError = expectedSafetyError,
        };
    }

    private DeliverableVerificationResult VerifyCandidate(
        byte[] input,
        byte[] candidate,
        byte[] expected)
    {
        var options = new DeliverableVerificationOptions
        {
            PackageManifestOptions = _packageValidator.ManifestOptions,
            RequireNoPlaceholders = false,
        };
        var target = DeliverableVerifier.VerifyDeliverable(input, expected, options);
        var expectedPackageChanges = target.PackageChanges.Select(change =>
            new DeliverablePackageChangeExpectation
            {
                Kind = change.Kind,
                Location = change.Location,
                BeforeDigest = change.BeforeDigest,
                AfterDigest = change.AfterDigest,
                BeforeValue = change.BeforeValue,
                AfterValue = change.AfterValue,
            }).ToArray();
        // DeliverableVerifier computes its modeled semantic delta with package supplements
        // disabled; package/relationship facts are matched independently below. Feed it the same
        // semantic surface so bookmark/revision supplements are not falsely reported as missing.
        var expectedModeledChanges = CompareSemantic(
            input, expected, includePackageChanges: false);
        return DeliverableVerifier.VerifyDeliverable(new DeliverableVerificationRequest
        {
            BaselineBytes = input,
            DeliverableBytes = candidate,
            ExpectedSemanticChanges = expectedModeledChanges,
            ExpectedPackageChanges = expectedPackageChanges,
        }, options with { FailOnUnexpectedChanges = true });
    }

    private (byte[] RedlineBytes, RedlineReversibilityProofRun ProofRun) BuildRedlineProof(
        byte[] input,
        byte[] candidate)
    {
        var proofBaseline = NormalizeProofPackage(input);
        var proofCandidate = NormalizeProofPackage(candidate);
        var baselineRevisionIds = RevisionIds(proofBaseline);
        var candidateRevisionIds = RevisionIds(proofCandidate);
        var candidateIsRedline = candidateRevisionIds.Except(
            baselineRevisionIds, StringComparer.Ordinal).Any();
        var intendedFinal = candidateIsRedline
            ? BuildResolvedFinalOnBaselineShell(
                proofBaseline,
                ResolveGeneratedCandidateRevisions(
                    proofCandidate, baselineRevisionIds, accept: true),
                proofCandidate)
            : proofCandidate;

        byte[] redline;
        if (TryBuildSimpleParagraphRedline(
                proofBaseline, intendedFinal, proofCandidate, out var simpleRedline))
        {
            redline = simpleRedline;
        }
        else
        {
            var settings = DiffSettings("Legal Evaluation Redline");
            settings.PreAcceptInputRevisions = false;
            settings.PreserveInputRevisions = true;
            var redlineSource = DocxDiff.Compare(
                new WmlDocument("input.docx", proofBaseline),
                new WmlDocument("candidate.docx", intendedFinal), settings).DocumentByteArray;
            redline = BuildRedlineOnBaselineShell(proofBaseline, redlineSource);
        }

        redline = NormalizeProofPackage(
            RestoreUnchangedPreExistingRevisionParagraphs(
                proofBaseline, intendedFinal, redline));
        var proof = RedlineReversibilityVerifier.Prove(
            proofBaseline, intendedFinal, redline, new RedlineReversibilityProofOptions
            {
                PackageManifestOptions = _packageValidator.ManifestOptions,
                MaxRevisionElements = 2_000,
            });
        return (redline, proof);
    }

    private static byte[] BuildResolvedFinalOnBaselineShell(
        byte[] baseline,
        byte[] resolvedCandidate,
        byte[] candidateWithRevisions)
    {
        var baselineParts = ReadXmlParts(baseline);
        var resolvedParts = ReadXmlParts(resolvedCandidate);
        var candidateParts = ReadXmlParts(candidateWithRevisions);
        using var output = new MemoryStream();
        output.Write(baseline);
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, true))
        {
            foreach (var part in EnumerateParts(package.MainDocumentPart!))
            {
                var uri = part.Uri.ToString();
                if (!IsXmlPart(part)
                    || !baselineParts.TryGetValue(uri, out var baselineXml)
                    || !candidateParts.TryGetValue(uri, out var candidateXml)
                    || !resolvedParts.TryGetValue(uri, out var resolvedXml))
                    continue;
                var baselineRevisions = baselineXml.Descendants()
                    .Where(value => TrackedRevisionNames.Contains(value.Name)).ToList();
                var hasGeneratedRevision = candidateXml.Descendants()
                    .Where(value => TrackedRevisionNames.Contains(value.Name))
                    .Any(value => !baselineRevisions.Any(baselineRevision =>
                        XNode.DeepEquals(baselineRevision, value)));
                if (hasGeneratedRevision)
                    part.PutXDocument(new XDocument(resolvedXml));
            }
        }
        return NormalizeProofPackage(
            GeneratedPackageNormalizer.RestoreUnchangedReviewParagraphs(
                baseline, output.ToArray()));
    }

    /// <summary>
    /// Build a package-exact redline when every edit is confined to the direct runs of an existing
    /// paragraph.  Tracking the complete old/new run groups is intentionally coarser than a word
    /// diff, but it makes both resolution trees exact: reject unwraps the original runs and accept
    /// unwraps the intended runs.  Any structural, nested-inline, or paragraph-count change opts
    /// out and uses DocxDiff instead.
    /// </summary>
    private static bool TryBuildSimpleParagraphRedline(
        byte[] baseline,
        byte[] intendedFinal,
        byte[] candidateWithAuthorship,
        out byte[] redline)
    {
        var finalParts = ReadXmlParts(intendedFinal);
        var authoredParts = ReadXmlParts(candidateWithAuthorship);
        using var output = new MemoryStream();
        output.Write(baseline);
        output.Position = 0;
        var revisionId = 900_000_000;
        var changedParagraphs = 0;
        using (var package = WordprocessingDocument.Open(output, true))
        {
            foreach (var part in EnumerateParts(package.MainDocumentPart!))
            {
                if (!IsXmlPart(part)
                    || !finalParts.TryGetValue(part.Uri.ToString(), out var finalXml)
                    || !authoredParts.TryGetValue(part.Uri.ToString(), out var authoredXml))
                    continue;

                XDocument baselineXml;
                try
                {
                    baselineXml = part.GetXDocument();
                }
                catch (XmlException)
                {
                    continue;
                }
                var baselineParagraphs = baselineXml.Descendants(W + "p").ToList();
                var finalParagraphs = finalXml.Descendants(W + "p").ToList();
                var authoredParagraphs = authoredXml.Descendants(W + "p").ToList();
                if (baselineParagraphs.Count != finalParagraphs.Count
                    || baselineParagraphs.Count != authoredParagraphs.Count)
                {
                    redline = Array.Empty<byte>();
                    return false;
                }

                var partChanged = false;
                for (var index = 0; index < baselineParagraphs.Count; index++)
                {
                    var before = baselineParagraphs[index];
                    var after = finalParagraphs[index];
                    if (XNode.DeepEquals(before, after)) continue;
                    if (!TryBuildSimpleReplacementParagraph(
                            before, after, authoredParagraphs[index],
                            ref revisionId, out var replacement))
                    {
                        redline = Array.Empty<byte>();
                        return false;
                    }
                    before.ReplaceWith(replacement);
                    changedParagraphs++;
                    partChanged = true;
                }
                if (partChanged) part.PutXDocument();
            }
        }
        redline = output.ToArray();
        return changedParagraphs != 0;
    }

    private static bool TryBuildSimpleReplacementParagraph(
        XElement before,
        XElement after,
        XElement authored,
        ref int revisionId,
        out XElement replacement)
    {
        replacement = null!;
        if (ContainsTrackedRevision(before) || ContainsTrackedRevision(after)) return false;
        if (!AttributesEqual(before, after)) return false;

        var beforeChildren = before.Elements().ToList();
        var afterChildren = after.Elements().ToList();
        var beforeOther = beforeChildren.Where(value => value.Name != W + "r").ToList();
        var afterOther = afterChildren.Where(value => value.Name != W + "r").ToList();
        if (beforeOther.Count != afterOther.Count
            || beforeOther.Where((value, index) => !XNode.DeepEquals(value, afterOther[index])).Any())
            return false;
        var beforeGroups = SplitDirectRunGroups(beforeChildren);
        var afterGroups = SplitDirectRunGroups(afterChildren);
        if (beforeGroups.Count != afterGroups.Count
            || !beforeGroups.Zip(afterGroups).Any(pair => !NodesEqual(pair.First, pair.Second)))
            return false;

        var authoredRevisions = authored.Descendants()
            .Where(value => TrackedRevisionNames.Contains(value.Name)
                && value.Attribute(W + "author") is not null)
            .ToList();
        var baselineRevisionShapes = before.Document?.Descendants()
            .Where(value => TrackedRevisionNames.Contains(value.Name)).ToList()
            ?? new List<XElement>();
        var authors = authoredRevisions
            .Where(value => !baselineRevisionShapes.Any(baselineRevision =>
                XNode.DeepEquals(baselineRevision, value)))
            .Select(value => (string?)value.Attribute(W + "author"))
            .Where(value => !string.IsNullOrWhiteSpace(value)).Cast<string>()
            .Distinct(StringComparer.Ordinal).ToList();
        var author = authors.Count == 1 ? authors[0] : "Legal Evaluation Redline";
        var date = authoredRevisions.Select(value => (string?)value.Attribute(W + "date"))
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
            ?? "2026-01-15T12:00:00Z";

        var children = new List<object>();
        for (var index = 0; index < beforeGroups.Count; index++)
        {
            var beforeGroup = beforeGroups[index];
            var afterGroup = afterGroups[index];
            if (NodesEqual(beforeGroup, afterGroup))
            {
                children.AddRange(beforeGroup.Select(value => (object)new XElement(value)));
            }
            else
            {
                if (beforeGroup.Count != 0)
                {
                children.Add(new XElement(W + "del",
                    new XAttribute(W + "id", revisionId++),
                    new XAttribute(W + "author", author),
                    new XAttribute(W + "date", date),
                        beforeGroup.Select(ToDeletedRun)));
                }
                if (afterGroup.Count != 0)
                {
                children.Add(new XElement(W + "ins",
                    new XAttribute(W + "id", revisionId++),
                    new XAttribute(W + "author", author),
                    new XAttribute(W + "date", date),
                        afterGroup.Select(value => new XElement(value))));
                }
            }
            if (index < beforeOther.Count)
                children.Add(new XElement(beforeOther[index]));
        }
        replacement = new XElement(before.Name, before.Attributes(), children);
        return true;
    }

    private static XElement ToDeletedRun(XElement run)
    {
        var result = new XElement(run);
        foreach (var text in result.Descendants(W + "t").ToList())
            text.Name = W + "delText";
        foreach (var instruction in result.Descendants(W + "instrText").ToList())
            instruction.Name = W + "delInstrText";
        return result;
    }

    private static bool AttributesEqual(XElement left, XElement right) =>
        left.Attributes().OrderBy(value => value.Name.ToString(), StringComparer.Ordinal)
            .Select(value => (value.Name, value.Value))
            .SequenceEqual(right.Attributes()
                .OrderBy(value => value.Name.ToString(), StringComparer.Ordinal)
                .Select(value => (value.Name, value.Value)));

    private static IReadOnlyList<IReadOnlyList<XElement>> SplitDirectRunGroups(
        IReadOnlyList<XElement> children)
    {
        var groups = new List<IReadOnlyList<XElement>>();
        var current = new List<XElement>();
        foreach (var child in children)
        {
            if (child.Name == W + "r")
            {
                current.Add(child);
                continue;
            }
            groups.Add(current);
            current = new List<XElement>();
        }
        groups.Add(current);
        return groups;
    }

    private static bool NodesEqual(
        IReadOnlyList<XElement> left,
        IReadOnlyList<XElement> right) =>
        left.Count == right.Count
        && left.Where((value, index) => !XNode.DeepEquals(value, right[index])).Any() is false;

    private static HashSet<string> RevisionIds(byte[] bytes)
    {
        using var session = new DocxSession(bytes, new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            EmitMarkdownPatch = false,
        });
        return session.ListRevisions().Select(value => value.Id)
            .ToHashSet(StringComparer.Ordinal);
    }

    private static byte[] ResolveGeneratedCandidateRevisions(
        byte[] redline,
        IReadOnlySet<string> baselineRevisionIds,
        bool accept)
    {
        using var session = new DocxSession(redline, new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            EmitMarkdownPatch = false,
        });
        var generated = session.ListRevisions()
            .Where(value => !baselineRevisionIds.Contains(value.Id))
            .Select(value => value.Id).ToList();
        foreach (var id in generated)
        {
            var result = accept ? session.AcceptRevision(id) : session.RejectRevision(id);
            if (!result.Success)
                throw new InvalidOperationException(
                    $"could not {(accept ? "accept" : "reject")} generated revision '{id}': "
                    + (result.Error?.Message ?? "unknown revision resolver failure"));
        }
        return NormalizeProofPackage(session.Save(persistAnchorIds: false));
    }

    private static byte[] NormalizeProofPackage(byte[] bytes)
    {
        var normalized = GeneratedPackageNormalizer.Normalize(bytes);
        using var output = new MemoryStream();
        output.Write(normalized);
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, true))
        {
            foreach (var part in EnumerateParts(package.MainDocumentPart!))
            {
                if (!IsXmlPart(part)) continue;
                XDocument xml;
                try
                {
                    xml = part.GetXDocument();
                }
                catch (XmlException)
                {
                    continue;
                }
                var changed = false;
                foreach (var text in xml.Descendants().Where(value =>
                             value.Name is var name
                             && (name == W + "t" || name == W + "delText"
                                 || name == W + "instrText" || name == W + "delInstrText")))
                {
                    var space = text.Attribute(XNamespace.Xml + "space");
                    if (space?.Value == "preserve"
                        && text.Value.Length != 0
                        && !char.IsWhiteSpace(text.Value[0])
                        && !char.IsWhiteSpace(text.Value[^1]))
                    {
                        space.Remove();
                        changed = true;
                    }
                }
                if (changed) part.PutXDocument();
            }
        }
        return ZipPackageOutputNormalizer.NormalizeDeterministic(output.ToArray());
    }

    /// <summary>
    /// DocxDiff owns revision markup, not unrelated package normalization.  Start from the
    /// normalized baseline package and replace only XML story parts that actually carry revision
    /// markup.  This keeps the reject path attributable to the selected baseline and prevents
    /// comparer-added font/theme scaffolding from masquerading as a legal edit.
    /// </summary>
    private static byte[] BuildRedlineOnBaselineShell(byte[] baseline, byte[] revisionSource)
    {
        var sourceParts = ReadXmlParts(revisionSource);
        using var output = new MemoryStream();
        output.Write(baseline);
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, true))
        {
            foreach (var part in EnumerateParts(package.MainDocumentPart!))
            {
                if (!IsXmlPart(part)
                    || !sourceParts.TryGetValue(part.Uri.ToString(), out var sourceXml)
                    || !sourceXml.Descendants().Any(value =>
                        TrackedRevisionNames.Contains(value.Name)))
                    continue;
                part.PutXDocument(new XDocument(sourceXml));
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// DocxDiff's preserve-input-revisions v1 contract intentionally flattens foreign markup in a
    /// modified block.  Its alignment can also classify an otherwise unchanged dirty paragraph as
    /// modified when edits elsewhere shift a list.  For proof generation we repair only the narrow,
    /// mechanically safe case: a revision-bearing candidate paragraph that is byte-for-byte equal
    /// to a baseline paragraph and has one unambiguous accepted-text peer in the redline.  Paragraphs
    /// containing the evaluated edit are never eligible, so no generated change can be hidden.
    /// </summary>
    private static byte[] RestoreUnchangedPreExistingRevisionParagraphs(
        byte[] baseline,
        byte[] candidate,
        byte[] redline)
    {
        var baselineParts = ReadXmlParts(baseline);
        var candidateParts = ReadXmlParts(candidate);
        using var output = new MemoryStream();
        output.Write(redline);
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, true))
        {
            foreach (var part in EnumerateParts(package.MainDocumentPart!))
            {
                var uri = part.Uri.ToString();
                if (!baselineParts.TryGetValue(uri, out var baselineXml)
                    || !candidateParts.TryGetValue(uri, out var candidateXml)
                    || !IsXmlPart(part))
                    continue;

                XDocument redlineXml;
                try
                {
                    redlineXml = part.GetXDocument();
                }
                catch (XmlException)
                {
                    continue;
                }

                var baselineParagraphs = baselineXml.Descendants(W + "p").ToList();
                var redlineParagraphs = redlineXml.Descendants(W + "p").ToList();
                var changed = false;
                foreach (var paragraph in candidateXml.Descendants(W + "p")
                    .Where(ContainsTrackedRevision))
                {
                    if (!baselineParagraphs.Any(value => XNode.DeepEquals(value, paragraph)))
                        continue;

                    var acceptedText = AcceptedParagraphText(paragraph);
                    var candidates = redlineParagraphs
                        .Where(value => string.Equals(
                            AcceptedParagraphText(value), acceptedText, StringComparison.Ordinal))
                        .ToList();
                    if (candidates.Count != 1 || ContainsTrackedRevision(candidates[0]))
                        continue;

                    var replacement = new XElement(paragraph);
                    candidates[0].ReplaceWith(replacement);
                    redlineParagraphs[redlineParagraphs.IndexOf(candidates[0])] = replacement;
                    changed = true;
                }

                if (changed)
                    part.PutXDocument();
            }
        }
        return output.ToArray();
    }

    private static readonly XNamespace W = WNamespace;

    private static readonly HashSet<XName> TrackedRevisionNames = new()
    {
        W + "ins", W + "del", W + "moveFrom", W + "moveTo",
        W + "moveFromRangeStart", W + "moveFromRangeEnd",
        W + "moveToRangeStart", W + "moveToRangeEnd",
        W + "customXmlDelRangeStart", W + "customXmlDelRangeEnd",
        W + "customXmlInsRangeStart", W + "customXmlInsRangeEnd",
        W + "pPrChange", W + "rPrChange", W + "tblPrChange",
        W + "tblGridChange", W + "trPrChange", W + "tcPrChange",
        W + "sectPrChange", W + "cellIns", W + "cellDel", W + "cellMerge",
    };

    private static bool ContainsTrackedRevision(XElement paragraph) =>
        paragraph.DescendantsAndSelf().Any(value => TrackedRevisionNames.Contains(value.Name));

    private static string AcceptedParagraphText(XElement paragraph) =>
        string.Concat(paragraph.Descendants(W + "t").Select(value => value.Value));

    private static IReadOnlyDictionary<string, XDocument> ReadXmlParts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var package = WordprocessingDocument.Open(stream, false);
        var result = new Dictionary<string, XDocument>(StringComparer.Ordinal);
        long totalXmlBytes = 0;
        foreach (var part in EnumerateParts(package.MainDocumentPart!))
        {
            if (!IsXmlPart(part)) continue;
            if (result.Count == MaximumProofXmlParts)
                throw new ScenarioValidationException(
                    $"proof XML analysis exceeds the {MaximumProofXmlParts}-part limit");
            try
            {
                using var partStream = part.GetStream(FileMode.Open, FileAccess.Read);
                if (partStream.Length > MaximumInvariantPartBytes)
                    throw new ScenarioValidationException(
                        $"proof XML part '{part.Uri}' exceeds the {MaximumInvariantPartBytes}-byte limit");
                totalXmlBytes = checked(totalXmlBytes + partStream.Length);
                if (totalXmlBytes > MaximumProofXmlBytes)
                    throw new ScenarioValidationException(
                        $"proof XML analysis exceeds the {MaximumProofXmlBytes}-byte aggregate limit");
                using var reader = XmlReader.Create(partStream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = MaximumInvariantPartBytes,
                });
                result.Add(part.Uri.ToString(), XDocument.Load(reader, LoadOptions.PreserveWhitespace));
            }
            catch (XmlException)
            {
                // Opaque XML extension parts are outside the modeled revision surface.
            }
        }
        return result;
    }

    private static IReadOnlyList<OpenXmlPart> EnumerateParts(OpenXmlPart root)
    {
        var result = new List<OpenXmlPart>();
        var pending = new Queue<OpenXmlPart>();
        var seen = new HashSet<OpenXmlPart>();
        pending.Enqueue(root);
        while (pending.Count != 0)
        {
            var part = pending.Dequeue();
            if (!seen.Add(part)) continue;
            result.Add(part);
            foreach (var child in part.Parts)
                pending.Enqueue(child.OpenXmlPart);
        }
        return result;
    }

    private static bool IsXmlPart(OpenXmlPart part) =>
        part.ContentType.EndsWith("+xml", StringComparison.Ordinal)
        || part.ContentType.EndsWith("/xml", StringComparison.Ordinal);

    private (DeliveryChangeReceipt? Receipt,
        IReadOnlyDictionary<string, byte[]> Artifacts,
        string? UnavailableReason) BuildDeliveryReceipt(
        BaselineExecution baseline,
        byte[] candidate,
        ScoreKind kind,
        PackageManifest inputManifest,
        PackageManifest candidateManifest,
        SemanticChangeSet candidateChanges,
        DeliverableVerificationResult verification,
        RedlineReversibilityProofRun proofRun)
    {
        if (kind != ScoreKind.EngineBaseline)
            return (null, new Dictionary<string, byte[]>(),
                "model-planning candidates do not carry the scripted executor's authoritative batch trace");
        var traces = baseline.TransactionTraces ?? Array.Empty<BaselineTransactionTrace>();
        if (traces.Count == 0)
            return (null, new Dictionary<string, byte[]>(),
                baseline.DeliveryReceiptUnavailableReason
                ?? "the baseline produced no authoritative ExecuteBatch trace");
        if (!traces[0].BeforeBytes.AsSpan().SequenceEqual(baseline.Input))
            return (null, new Dictionary<string, byte[]>(),
                "the executor's first batch snapshot does not match the published evaluation input");
        if (!traces[^1].AfterBytes.AsSpan().SequenceEqual(candidate))
            return (null, new Dictionary<string, byte[]>(),
                "the executor's final batch snapshot does not match the published candidate");

        try
        {
            var builder = new DeliveryChangeReceiptBuilder(
                inputManifest, traces[0].Result.BaseVersion)
                .SetDeliveredDocument(candidateManifest, traces[^1].Result.ResultVersion);
            var artifacts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            builder.AddArtifact(DeliveryArtifactInput.Available(
                "clean-docx", DeliveryArtifactRole.CleanDocx,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                candidate) with
            {
                Document = DeliveryDocumentIdentity.FromManifest(
                    candidateManifest, traces[^1].Result.ResultVersion),
                RelativePath = "candidate.docx",
            });
            artifacts.Add("clean-docx", candidate);

            builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForSourceToDelivered(
                candidateChanges, "semantic-change-set-v1", "semantic-change-set-v1.json"));
            artifacts.Add("semantic-change-set-v1", candidateChanges.ToCanonicalUtf8Bytes());

            for (var index = 0; index < traces.Count; index++)
            {
                var trace = traces[index];
                var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
                    trace.Result, trace.BeforeManifest, trace.AfterManifest,
                    new[] { trace.Operation });
                var entryId = builder.AddTransaction(contribution);
                var transactionChanges = CompareSemantic(trace.BeforeBytes, trace.AfterBytes);
                var artifactId = $"transaction-semantic-change-set-{index + 1:D3}";
                builder.AddSemanticChangeSet(DeliverySemanticChangeSetInput.ForTransaction(
                    entryId, transactionChanges, artifactId,
                    $"receipt/{artifactId}.json"));
                artifacts.Add(artifactId, transactionChanges.ToCanonicalUtf8Bytes());
            }

            var verificationBytes = verification.ToCanonicalUtf8Bytes();
            builder.AddArtifact(DeliveryArtifactInput.Available(
                "deliverable-verification-v1", DeliveryArtifactRole.ValidationReport,
                "application/json", verificationBytes) with
            {
                RelativePath = "deliverable-verification-v1.json",
            });
            artifacts.Add("deliverable-verification-v1", verificationBytes);
            builder.AddEvidence(new DeliveryEvidenceReference
            {
                Kind = DeliveryEvidenceKind.ValidationResult,
                Schema = DeliverableVerificationResult.SchemaId,
                Digest = Digest(verificationBytes),
                ArtifactId = "deliverable-verification-v1",
                Summary = $"Deliverable verification decision: {verification.Decision}.",
            });

            var proofBytes = Encoding.UTF8.GetBytes(proofRun.Proof.ToCanonicalJson());
            builder.AddArtifact(DeliveryArtifactInput.Available(
                "redline-reversibility-proof-v1", DeliveryArtifactRole.ReversibilityProof,
                "application/json", proofBytes) with
            {
                RelativePath = "redline-reversibility-proof-v1.json",
            });
            artifacts.Add("redline-reversibility-proof-v1", proofBytes);
            builder.AddEvidence(new DeliveryEvidenceReference
            {
                Kind = DeliveryEvidenceKind.RedlineReversibility,
                Schema = RedlineReversibilityProof.SchemaId,
                Digest = Digest(proofBytes),
                ArtifactId = "redline-reversibility-proof-v1",
                Summary = $"Redline reversibility success: {proofRun.Proof.Success}.",
            });

            var receipt = builder.Build();
            var receiptVerification = DeliveryChangeReceiptVerifier.Verify(receipt, artifacts);
            if (!receiptVerification.IsValid)
            {
                return (null, artifacts,
                    "the generated #458 receipt failed independent verification: "
                    + string.Join(", ", receiptVerification.Findings));
            }
            return (receipt, artifacts, null);
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return (null, new Dictionary<string, byte[]>(),
                $"#458 receipt construction failed: {exception.GetType().Name}: {exception.Message}");
        }
    }

    private static VerificationDigest Digest(ReadOnlySpan<byte> bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    internal static IReadOnlyList<string> PublishedOperationLog(
        ScoreKind kind, IReadOnlyList<string> baselineOperationLog) =>
        kind == ScoreKind.EngineBaseline
            ? baselineOperationLog
            : Array.Empty<string>();

    internal static string ScenarioContract(LegalScenario scenario) =>
        JsonSerializer.Serialize(new
        {
            schemaVersion = "docxodus.legal-evaluation-scenario-contract/1.0",
            sourceSchemaVersion = "2.0",
            evaluatorVersion = typeof(EvaluationScorer).Assembly.GetName().Version?.ToString()
                ?? "unknown",
            docxodusVersion = typeof(DocxSession).Assembly.GetName().Version?.ToString()
                ?? "unknown",
            scenario.Id,
            scenario.Title,
            tier = scenario.Tier == EvalTier.Fast ? "fast" : "full",
            fixture = new
            {
                provenanceId = scenario.Fixture.ProvenanceId,
                sourceSha256 = scenario.Fixture.SourceSha256,
            },
            expectedDocument = new
            {
                provenanceId = scenario.ExpectedDocument.ProvenanceId,
                sourceSha256 = scenario.ExpectedDocument.SourceSha256,
            },
            scenario.Instruction,
            scenario.Constraints,
            redlineReversibility = new
            {
                applicability = scenario.RedlineReversibility.Applicability
                    == RedlineReversibilityApplicability.Required
                        ? "required"
                        : "notApplicable",
                scenario.RedlineReversibility.Reason,
            },
            expectedOutputs = scenario.ExpectedOutputs,
            baseline = new
            {
                executor = "scripted-session-v1",
                operations = scenario.BaselineOperations,
            },
            invariants = scenario.Invariants,
            changeBudget = new
            {
                allowedChangedParts = scenario.ChangeBudget.AllowedChangedParts
                    .Order(StringComparer.Ordinal),
                allowedRelationshipOwners = scenario.ChangeBudget.AllowedRelationshipOwners
                    .Order(StringComparer.Ordinal),
                scenario.ChangeBudget.MaximumChangedAnchors,
            },
        }, ScenarioContractJsonOptions);

    internal static string RenderHtml(byte[] bytes)
    {
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = true,
            CssClassPrefix = "legal-eval-",
            PageTitle = "Docxodus legal evaluation",
        };
        var html = WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument("candidate.docx", bytes), settings);
        HardenPreview(html);
        return html.ToString(SaveOptions.DisableFormatting);
    }

    private static void HardenPreview(XElement html)
    {
        var head = html.Elements().FirstOrDefault(value => value.Name.LocalName == "head");
        head?.AddFirst(new XElement(html.Name.Namespace + "meta",
            new XAttribute("http-equiv", "Content-Security-Policy"),
            new XAttribute("content",
                "default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'; "
                + "base-uri 'none'; form-action 'none'; frame-ancestors 'none'")));

        foreach (var element in html.DescendantsAndSelf())
        {
            foreach (var attribute in element.Attributes()
                .Where(value => value.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                .ToList())
                attribute.Remove();

            var href = element.Attribute("href");
            if (href is not null && !href.Value.StartsWith('#'))
            {
                element.SetAttributeValue("title", "External link suppressed in evaluation preview");
                href.Value = "#";
            }
            var source = element.Attribute("src");
            if (source is not null && !source.Value.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                source.Remove();
        }
    }

    private static string NormalizeHtml(string html) =>
        html.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();

    private static (string? Json, string? Error) TrySemanticDiff(byte[] input, byte[] candidate)
    {
        try
        {
            return (SemanticDiff.CompareBounded(
                new WmlDocument("before.docx", input),
                new WmlDocument("after.docx", candidate),
                new SemanticDiffOptions(),
                MaximumSemanticChanges).ToCanonicalJson(), null);
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return (null, exception.Message);
        }
    }

    internal static (string Json, bool Succeeded) SemanticDiffArtifact(
        byte[] input, byte[] candidate, string artifact)
    {
        var result = TrySemanticDiff(input, candidate);
        return (result.Json ?? ErrorEnvelope(artifact, "failed",
            result.Error ?? "semantic diff generation failed"), result.Json is not null);
    }

    private static readonly JsonSerializerOptions ScenarioContractJsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    internal static string ErrorEnvelope(string artifact, string status, string detail) =>
        JsonSerializer.Serialize(new
        {
            schemaVersion = "docxodus.evaluation-error/1.0",
            artifact,
            status,
            detail,
        }, new JsonSerializerOptions { WriteIndented = true });

    private static DocxDiffSettings DiffSettings(string author) => new()
    {
        AuthorForRevisions = author,
        Deterministic = true,
        DateTimeForRevisions = "2026-01-15T12:00:00Z",
        CompareHeadersFooters = true,
        TrackBlockFormatChanges = true,
        PreAcceptInputRevisions = true,
    };

    private static byte[]? PartBytes(byte[] bytes, string part, bool required)
    {
        var normalized = part.TrimStart('/');
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read, leaveOpen: false);
        var entry = archive.GetEntry(normalized);
        if (entry is null)
        {
            if (required) throw new InvalidOperationException($"package part '{part}' is absent");
            return null;
        }
        if (entry.Length > MaximumInvariantPartBytes)
            throw new ScenarioValidationException(
                $"invariant part '{part}' exceeds the {MaximumInvariantPartBytes}-byte limit");
        using var stream = entry.Open();
        using var copy = new MemoryStream(checked((int)entry.Length));
        stream.CopyTo(copy);
        if (copy.Length > MaximumInvariantPartBytes)
            throw new ScenarioValidationException(
                $"invariant part '{part}' expanded beyond the {MaximumInvariantPartBytes}-byte limit");
        return copy.ToArray();
    }

    private static bool PartExists(byte[] bytes, string part)
    {
        using var archive = new ZipArchive(
            new MemoryStream(bytes), ZipArchiveMode.Read, leaveOpen: false);
        return archive.GetEntry(part.TrimStart('/')) is not null;
    }

    private static string RequiredString(JsonObject parent, string name) =>
        parent[name]?.GetValue<string>()
            ?? throw new ScenarioValidationException($"probe property '{name}' must be a string");

    private static string Json(JsonNode? node) =>
        node?.ToJsonString(new JsonSerializerOptions { WriteIndented = false }) ?? "null";

    private static double Number(JsonNode? node)
    {
        if (node is not JsonValue value)
            throw new ScenarioValidationException("numeric invariant operand is null or non-scalar");
        if (value.TryGetValue<int>(out var intValue)) return intValue;
        if (value.TryGetValue<long>(out var longValue)) return longValue;
        if (value.TryGetValue<double>(out var doubleValue)) return doubleValue;
        if (value.TryGetValue<decimal>(out var decimalValue)) return (double)decimalValue;
        throw new ScenarioValidationException("numeric invariant operand is not a number");
    }

    private static HashSet<string> Set(JsonNode? node) =>
        node is JsonArray array
            ? array.Select(value => value?.GetValue<string>()
                ?? throw new ScenarioValidationException("set members must be strings"))
                .ToHashSet(StringComparer.Ordinal)
            : throw new ScenarioValidationException("set operand must be an array");
}
