// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Docxodus.Verification;

namespace LegalEval;

public static class ScenarioLoader
{
    private const int MaximumJsonBytes = 4 * 1024 * 1024;
    private const int MaximumProvenanceDocumentBytes = 64 * 1024 * 1024;
    private const int MaximumArrayItems = 256;
    private const int MaximumBaselineOperations = 32;
    private const int MaximumConsolidationReviewers = 16;
    private const int MaximumStringCharacters = 16 * 1024;
    private static readonly Regex ScenarioSlug = new(
        "^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$",
        RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);
    private static readonly HashSet<string> ReservedWindowsFileNames = new(
        new[]
        {
            "con", "prn", "aux", "nul",
            "com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8", "com9",
            "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
        }, StringComparer.OrdinalIgnoreCase);

    private static readonly HashSet<string> Metrics = new(StringComparer.Ordinal)
    {
        "task_completion",
        "target_precision",
        "unintended_change",
        "document_validity",
        "redline_reversibility",
        "rendering_regression",
    };

    private static readonly IReadOnlyDictionary<string, (string MediaType, string Role)>
        ExpectedOutputContracts =
        new Dictionary<string, (string MediaType, string Role)>(StringComparer.Ordinal)
        {
            ["candidate-docx"] =
                ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "candidate"),
            ["semantic-change-set-v1"] = ("application/json", "verification"),
            ["after-html"] = ("text/html", "review"),
            ["redline-docx"] =
                ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "review"),
            ["candidate-pdf"] = ("application/pdf", "review"),
            ["candidate-package-manifest-v1"] = ("application/json", "verification"),
            ["deliverable-verification-v1"] = ("application/json", "verification"),
            ["delivery-change-receipt-v1"] = ("application/json", "verification"),
            ["redline-reversibility-proof-v1"] = ("application/json", "verification"),
        };

    public static LegalCorpus LoadCorpus(string corpusPath)
    {
        var fullCorpusPath = LegalEvaluationRunner.ResolveCanonicalPath(corpusPath);
        var root = Path.GetDirectoryName(fullCorpusPath)
            ?? throw new ScenarioValidationException($"Corpus path has no directory: {corpusPath}");
        var corpus = ReadObject(fullCorpusPath, "corpus");
        RejectUnknown(corpus, fullCorpusPath, "schemaVersion", "provenance", "scenarios");
        RequireExactVersion(corpus, fullCorpusPath, "1.0");

        var provenancePath = ResolveUnderRoot(root,
            RequireString(corpus, "provenance", fullCorpusPath), fullCorpusPath);
        var provenance = LoadProvenance(provenancePath, root);
        var scenarioNodes = RequireArray(corpus, "scenarios", fullCorpusPath);
        if (scenarioNodes.Count == 0)
            throw Error(fullCorpusPath, "scenarios must contain at least one path");

        var scenarios = new List<LegalScenario>(scenarioNodes.Count);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in scenarioNodes)
        {
            var relative = RequireNodeString(node, fullCorpusPath, "scenario path");
            var path = ResolveUnderRoot(root, relative, fullCorpusPath);
            var scenario = LoadScenario(path, root, provenance.Fixtures,
                provenance.ExpectedDocuments);
            if (!ids.Add(scenario.Id))
                throw Error(path, $"duplicate scenario id '{scenario.Id}'");
            scenarios.Add(scenario);
        }

        return new LegalCorpus(root, scenarios, provenance.Fixtures,
            provenance.ExpectedDocuments);
    }

    public static LegalScenario LoadScenario(
        string scenarioPath,
        string corpusRoot,
        IReadOnlyDictionary<string, FixtureProvenance> provenance,
        IReadOnlyDictionary<string, ExpectedDocumentProvenance> expectedDocumentProvenance)
    {
        var fullPath = LegalEvaluationRunner.ResolveCanonicalPath(scenarioPath);
        var node = ReadObject(fullPath, "scenario");
        RejectUnknown(node, fullPath, "schemaVersion", "id", "title", "tier", "fixture",
            "expectedDocument", "instruction", "redlineReversibility", "expectedOutputs",
            "baseline", "invariants", "changeBudget");
        RequireExactVersion(node, fullPath, "2.0");
        var id = RequireString(node, "id", fullPath);
        if (id.Length > 128 || !ScenarioSlug.IsMatch(id)
            || ReservedWindowsFileNames.Contains(id))
            throw Error(fullPath, $"id '{id}' is not a safe scenario slug");
        var title = RequireString(node, "title", fullPath);
        var tier = RequireString(node, "tier", fullPath) switch
        {
            "fast" => EvalTier.Fast,
            "full" => EvalTier.Full,
            var value => throw Error(fullPath, $"tier must be 'fast' or 'full', got '{value}'"),
        };

        var fixtureNode = RequireObject(node, "fixture", fullPath);
        RejectUnknown(fixtureNode, fullPath, "path", "provenanceId", "sourceSha256");
        var fixtureRelative = RequireString(fixtureNode, "path", fullPath);
        var fixturePath = ResolveUnderRoot(corpusRoot, fixtureRelative, fullPath);
        if (!File.Exists(fixturePath))
            throw Error(fullPath, $"fixture does not exist: {fixtureRelative}");
        var provenanceId = RequireString(fixtureNode, "provenanceId", fullPath);
        if (!provenance.TryGetValue(provenanceId, out var provenanceEntry))
            throw Error(fullPath, $"unknown fixture provenanceId '{provenanceId}'");
        var sourceSha = RequireSha256(fixtureNode, "sourceSha256", fullPath);
        var actualSha = Sha256File(fixturePath);
        if (!string.Equals(sourceSha, actualSha, StringComparison.Ordinal))
            throw Error(fullPath,
                $"fixture sourceSha256 mismatch for {fixtureRelative}: expected {sourceSha}, actual {actualSha}");
        if (!string.Equals(provenanceEntry.SourcePath, fixturePath, PathComparison)
            || !string.Equals(provenanceEntry.SourceSha256, actualSha, StringComparison.Ordinal))
            throw Error(fullPath, $"fixture provenance mismatch for '{provenanceId}'");

        var expectedNode = RequireObject(node, "expectedDocument", fullPath);
        RejectUnknown(expectedNode, fullPath, "path", "provenanceId", "sourceSha256");
        var expectedRelative = RequireString(expectedNode, "path", fullPath);
        var expectedPath = ResolveUnderRoot(corpusRoot, expectedRelative, fullPath);
        if (!File.Exists(expectedPath))
            throw Error(fullPath, $"expected document does not exist: {expectedRelative}");
        var expectedSha = RequireSha256(expectedNode, "sourceSha256", fullPath);
        var expectedProvenanceId = RequireString(expectedNode, "provenanceId", fullPath);
        if (!expectedDocumentProvenance.TryGetValue(expectedProvenanceId,
                out var expectedProvenanceEntry))
            throw Error(fullPath,
                $"unknown expectedDocument provenanceId '{expectedProvenanceId}'");
        var actualExpectedSha = Sha256File(expectedPath);
        if (!string.Equals(expectedSha, actualExpectedSha, StringComparison.Ordinal))
            throw Error(fullPath,
                $"expected document sourceSha256 mismatch for {expectedRelative}: expected {expectedSha}, actual {actualExpectedSha}");
        if (!string.Equals(expectedProvenanceEntry.SourcePath, expectedPath, PathComparison)
            || !string.Equals(expectedProvenanceEntry.SourceSha256, actualExpectedSha,
                StringComparison.Ordinal))
            throw Error(fullPath,
                $"expected document provenance mismatch for '{expectedProvenanceId}'");
        if (!string.Equals(expectedProvenanceEntry.ScenarioId, id, StringComparison.Ordinal))
            throw Error(fullPath,
                $"expected document provenance '{expectedProvenanceId}' belongs to scenario '{expectedProvenanceEntry.ScenarioId}'");

        var instructionNode = RequireObject(node, "instruction", fullPath);
        RejectUnknown(instructionNode, fullPath, "text", "constraints");
        var instruction = RequireString(instructionNode, "text", fullPath);
        var constraints = RequireArray(instructionNode, "constraints", fullPath)
            .Select((value, index) => RequireNodeString(
                value, fullPath, $"instruction.constraints[{index}]"))
            .ToList();
        if (constraints.Count == 0)
            throw Error(fullPath, "instruction.constraints must be explicit and non-empty");

        var reversibilityNode = RequireObject(node, "redlineReversibility", fullPath);
        RejectUnknown(reversibilityNode, fullPath, "applicability", "reason");
        var reversibilityApplicability = RequireString(
            reversibilityNode, "applicability", fullPath) switch
        {
            "required" => RedlineReversibilityApplicability.Required,
            "notApplicable" => RedlineReversibilityApplicability.NotApplicable,
            var value => throw Error(fullPath,
                $"redlineReversibility.applicability is unsupported: '{value}'"),
        };
        var reversibilityReason = reversibilityNode["reason"] is null
            ? null
            : RequireString(reversibilityNode, "reason", fullPath);
        if (reversibilityApplicability == RedlineReversibilityApplicability.NotApplicable
            && reversibilityReason is null)
            throw Error(fullPath,
                "redlineReversibility.reason is required when applicability is notApplicable");
        if (reversibilityApplicability == RedlineReversibilityApplicability.Required
            && reversibilityReason is not null)
            throw Error(fullPath,
                "redlineReversibility.reason is allowed only when applicability is notApplicable");
        var reversibility = new RedlineReversibilityPolicy(
            reversibilityApplicability, reversibilityReason);

        var outputsNode = RequireArray(node, "expectedOutputs", fullPath);
        if (outputsNode.Count == 0)
            throw Error(fullPath, "expectedOutputs must be explicit and non-empty");
        var outputs = outputsNode.Select((value, index) =>
        {
            var output = value as JsonObject
                ?? throw Error(fullPath, $"expectedOutputs[{index}] must be an object");
            RejectUnknown(output, fullPath, "id", "mediaType", "role", "required");
            var outputId = RequireString(output, "id", fullPath);
            var mediaType = RequireString(output, "mediaType", fullPath);
            if (!ExpectedOutputContracts.TryGetValue(outputId, out var contract))
                throw Error(fullPath,
                    $"expectedOutputs[{index}].id '{outputId}' is not a canonical artifact id");
            if (!string.Equals(mediaType, contract.MediaType, StringComparison.Ordinal))
                throw Error(fullPath,
                    $"expectedOutputs[{index}] mediaType for '{outputId}' must be '{contract.MediaType}'");
            var role = RequireString(output, "role", fullPath);
            if (!string.Equals(role, contract.Role, StringComparison.Ordinal))
                throw Error(fullPath,
                    $"expectedOutputs[{index}] role for '{outputId}' must be '{contract.Role}'");
            return new ExpectedArtifact(
                outputId,
                mediaType,
                role,
                RequireBool(output, "required", fullPath));
        }).ToList();
        if (outputs.Select(value => value.Id).Distinct(StringComparer.Ordinal).Count() != outputs.Count)
            throw Error(fullPath, "expectedOutputs ids must be unique within a scenario");

        var baseline = RequireObject(node, "baseline", fullPath);
        RejectUnknown(baseline, fullPath, "executor", "operations");
        if (RequireString(baseline, "executor", fullPath) != "scripted-session-v1")
            throw Error(fullPath, "baseline.executor must be 'scripted-session-v1'");
        var operationNodes = RequireArray(baseline, "operations", fullPath);
        if (operationNodes.Count == 0)
            throw Error(fullPath, "baseline.operations must be explicit and non-empty");
        if (operationNodes.Count > MaximumBaselineOperations)
            throw Error(fullPath,
                $"baseline.operations exceeds the {MaximumBaselineOperations}-operation limit");
        var operations = operationNodes.Select((value, index) =>
            value as JsonObject
                ?? throw Error(fullPath, $"baseline.operations[{index}] must be an object"))
            .Select(value => (JsonObject)value.DeepClone()).ToList();
        foreach (var operation in operations)
            ValidateOperation(operation, fullPath, nestedReviewer: false);
        var consolidationCount = operations.Count(value =>
            value["op"]?.GetValue<string>() == "consolidate");
        if (consolidationCount != 0 && operations.Count != 1)
            throw Error(fullPath,
                "consolidate must be the scenario's sole baseline operation");

        if (outputs.Count != ExpectedOutputContracts.Count
            || ExpectedOutputContracts.Keys.Except(
                outputs.Select(value => value.Id), StringComparer.Ordinal).Any())
            throw Error(fullPath,
                "expectedOutputs must declare every canonical artifact exactly once");
        foreach (var output in outputs)
        {
            var required = output.Id switch
            {
                "candidate-pdf" => false,
                "delivery-change-receipt-v1" => consolidationCount == 0,
                "redline-reversibility-proof-v1" =>
                    reversibility.Applicability == RedlineReversibilityApplicability.Required,
                _ => true,
            };
            if (output.Required != required)
                throw Error(fullPath,
                    $"expectedOutputs required policy for '{output.Id}' must be {required.ToString().ToLowerInvariant()}");
        }

        var invariantNodes = RequireArray(node, "invariants", fullPath);
        if (invariantNodes.Count == 0)
            throw Error(fullPath,
                "invariants must contain at least one explicit deterministic invariant; visual inspection is not an invariant");
        var invariants = invariantNodes.Select((value, index) =>
        {
            var invariant = value as JsonObject
                ?? throw Error(fullPath, $"invariants[{index}] must be an object");
            RejectUnknown(invariant, fullPath, "id", "metric", "deterministic", "probe",
                "operator", "expected");
            if (!RequireBool(invariant, "deterministic", fullPath))
                throw Error(fullPath, $"invariants[{index}].deterministic must be true");
            var metric = RequireString(invariant, "metric", fullPath);
            if (!Metrics.Contains(metric))
                throw Error(fullPath, $"invariants[{index}].metric '{metric}' is not recognized");
            if (!invariant.ContainsKey("expected"))
                throw Error(fullPath, $"invariants[{index}].expected is required");
            var probe = (JsonObject)RequireObject(invariant, "probe", fullPath).DeepClone();
            var operation = RequireString(invariant, "operator", fullPath);
            var expected = invariant["expected"]?.DeepClone();
            ValidateInvariant(probe, operation, expected, fullPath, index);
            return new DeterministicInvariant(
                RequireString(invariant, "id", fullPath),
                metric,
                probe,
                operation,
                expected);
        }).ToList();
        if (invariants.Select(value => value.Id).Distinct(StringComparer.Ordinal).Count() != invariants.Count)
            throw Error(fullPath, "invariant ids must be unique within a scenario");

        var budgetNode = RequireObject(node, "changeBudget", fullPath);
        RejectUnknown(budgetNode, fullPath, "allowedChangedParts",
            "allowedRelationshipOwners", "maximumChangedAnchors");
        var allowedParts = RequireArray(budgetNode, "allowedChangedParts", fullPath)
            .Select((value, index) => RequireNodeString(
                value, fullPath, $"changeBudget.allowedChangedParts[{index}]"))
            .ToHashSet(StringComparer.Ordinal);
        if (allowedParts.Count == 0)
            throw Error(fullPath, "changeBudget.allowedChangedParts must be explicit and non-empty");
        if (allowedParts.Any(value => string.IsNullOrWhiteSpace(value)
                || !value.StartsWith('/') || value.Contains("..", StringComparison.Ordinal)))
            throw Error(fullPath,
                "changeBudget.allowedChangedParts entries must be absolute safe OPC part names");
        var allowedRelationshipOwners = RequireArray(
                budgetNode, "allowedRelationshipOwners", fullPath)
            .Select((value, index) => RequireNodeString(
                value, fullPath, $"changeBudget.allowedRelationshipOwners[{index}]"))
            .ToHashSet(StringComparer.Ordinal);
        if (allowedRelationshipOwners.Any(value => string.IsNullOrWhiteSpace(value)
                || !value.StartsWith('/') || value.Contains("..", StringComparison.Ordinal)))
            throw Error(fullPath,
                "changeBudget.allowedRelationshipOwners entries must be absolute safe OPC owner names");
        var maximumChangedAnchors = RequireInt(budgetNode, "maximumChangedAnchors", fullPath);
        if (maximumChangedAnchors < 0)
            throw Error(fullPath, "changeBudget.maximumChangedAnchors must be non-negative");

        return new LegalScenario(fullPath, id, title, tier,
            new FixtureReference(fixturePath, provenanceId, sourceSha),
            new ExpectedDocumentReference(expectedPath, expectedProvenanceId, expectedSha),
            instruction, constraints, reversibility,
            outputs, operations, invariants,
            new ChangeBudget(allowedParts, allowedRelationshipOwners, maximumChangedAnchors));
    }

    private static ProvenanceCatalog LoadProvenance(
        string path, string corpusRoot)
    {
        var root = ReadObject(path, "provenance");
        RejectUnknown(root, path, "schemaVersion", "fixtures", "expectedDocuments");
        RequireExactVersion(root, path, "1.0");
        var records = RequireArray(root, "fixtures", path);
        var result = new Dictionary<string, FixtureProvenance>(StringComparer.Ordinal);
        foreach (var value in records)
        {
            var record = value as JsonObject ?? throw Error(path, "fixtures entries must be objects");
            RejectUnknown(record, path, "id", "title", "origin", "author", "created",
                "license", "redistributionPermission", "sourcePath", "sourceSha256",
                "recipePath", "recipeSha256");
            var sourcePath = ResolveUnderRoot(corpusRoot, RequireString(record, "sourcePath", path), path);
            var sourceSha = RequireSha256(record, "sourceSha256", path);
            if (!File.Exists(sourcePath))
                throw Error(path, $"provenance source does not exist: {sourcePath}");
            var actualSha = Sha256File(sourcePath);
            if (!string.Equals(sourceSha, actualSha, StringComparison.Ordinal))
                throw Error(path, $"provenance sourceSha256 mismatch: expected {sourceSha}, actual {actualSha}");
            string? recipePath = null;
            string? recipeSha = null;
            if (record["recipePath"] is not null || record["recipeSha256"] is not null)
            {
                recipePath = ResolveUnderRoot(corpusRoot,
                    RequireString(record, "recipePath", path), path);
                recipeSha = RequireSha256(record, "recipeSha256", path);
                if (!File.Exists(recipePath))
                    throw Error(path, $"provenance recipe does not exist: {recipePath}");
                var actualRecipeSha = Sha256File(recipePath);
                if (!string.Equals(recipeSha, actualRecipeSha, StringComparison.Ordinal))
                    throw Error(path,
                        $"provenance recipeSha256 mismatch: expected {recipeSha}, actual {actualRecipeSha}");
            }
            var entry = new FixtureProvenance(
                RequireString(record, "id", path),
                RequireString(record, "title", path),
                RequireString(record, "origin", path),
                RequireString(record, "author", path),
                RequireString(record, "created", path),
                RequireString(record, "license", path),
                RequireString(record, "redistributionPermission", path),
                sourcePath,
                sourceSha,
                recipePath,
                recipeSha);
            if (string.IsNullOrWhiteSpace(entry.RedistributionPermission))
                throw Error(path, $"fixture '{entry.Id}' has no redistribution permission");
            if (!result.TryAdd(entry.Id, entry))
                throw Error(path, $"duplicate fixture provenance id '{entry.Id}'");
        }
        if (result.Count == 0)
            throw Error(path, "fixtures provenance must not be empty");
        var expectedRecords = RequireArray(root, "expectedDocuments", path);
        var expectedResult = new Dictionary<string, ExpectedDocumentProvenance>(StringComparer.Ordinal);
        foreach (var value in expectedRecords)
        {
            var record = value as JsonObject
                ?? throw Error(path, "expectedDocuments entries must be objects");
            RejectUnknown(record, path, "id", "scenarioId", "origin", "generatedBy",
                "created", "reviewStatus", "reviewNotes", "license",
                "redistributionPermission", "sourcePath", "sourceSha256");
            var sourcePath = ResolveUnderRoot(corpusRoot,
                RequireString(record, "sourcePath", path), path);
            var sourceSha = RequireSha256(record, "sourceSha256", path);
            if (!File.Exists(sourcePath))
                throw Error(path, $"expected document provenance source does not exist: {sourcePath}");
            var actualSha = Sha256File(sourcePath);
            if (!string.Equals(sourceSha, actualSha, StringComparison.Ordinal))
                throw Error(path,
                    $"expected document provenance sourceSha256 mismatch: expected {sourceSha}, actual {actualSha}");
            var entry = new ExpectedDocumentProvenance(
                RequireString(record, "id", path),
                RequireString(record, "scenarioId", path),
                RequireString(record, "origin", path),
                RequireString(record, "generatedBy", path),
                RequireString(record, "created", path),
                RequireString(record, "reviewStatus", path),
                RequireString(record, "reviewNotes", path),
                RequireString(record, "license", path),
                RequireString(record, "redistributionPermission", path),
                sourcePath,
                sourceSha);
            if (string.IsNullOrWhiteSpace(entry.RedistributionPermission)
                || string.IsNullOrWhiteSpace(entry.ReviewStatus)
                || string.IsNullOrWhiteSpace(entry.ReviewNotes))
                throw Error(path,
                    $"expected document '{entry.Id}' has incomplete review or redistribution metadata");
            if (!expectedResult.TryAdd(entry.Id, entry))
                throw Error(path, $"duplicate expected document provenance id '{entry.Id}'");
        }
        if (expectedResult.Count == 0)
            throw Error(path, "expected document provenance must not be empty");
        return new ProvenanceCatalog(result, expectedResult);
    }

    private sealed record ProvenanceCatalog(
        IReadOnlyDictionary<string, FixtureProvenance> Fixtures,
        IReadOnlyDictionary<string, ExpectedDocumentProvenance> ExpectedDocuments);

    private static void ValidateOperation(JsonObject operation, string path, bool nestedReviewer)
    {
        var kind = RequireString(operation, "op", path);
        if (nestedReviewer && kind is not ("replaceText" or "replaceTableCell"))
            throw Error(path,
                $"consolidate reviewer operation '{kind}' is unsupported; use replaceText or replaceTableCell");
        switch (kind)
        {
            case "replaceText":
                RejectUnknown(operation, path, "op", "anchorContains", "find", "replace",
                    "tracked", "author");
                _ = RequireString(operation, "anchorContains", path);
                _ = RequireString(operation, "find", path);
                _ = RequireString(operation, "replace", path);
                if (operation["author"] is not null) _ = RequireString(operation, "author", path);
                if (operation["tracked"] is JsonValue tracked
                    && (!tracked.TryGetValue<bool>(out var trackedEnabled)
                        || (trackedEnabled && operation["author"] is null)))
                    throw Error(path, "tracked replaceText requires a boolean tracked value and author");
                if (operation["tracked"] is not null && operation["tracked"] is not JsonValue)
                    throw Error(path, "replaceText.tracked must be a boolean");
                if (nestedReviewer) _ = RequireString(operation, "author", path);
                break;
            case "insertNumberedClause":
                RejectUnknown(operation, path, "op", "afterAnchorContains", "styleId", "text");
                _ = RequireString(operation, "afterAnchorContains", path);
                _ = RequireString(operation, "styleId", path);
                _ = RequireString(operation, "text", path);
                break;
            case "replaceTableCell":
                RejectUnknown(operation, path, "op", "author", "row", "column", "text");
                if (RequireInt(operation, "row", path) < 0
                    || RequireInt(operation, "column", path) < 0)
                    throw Error(path, "replaceTableCell row and column must be non-negative");
                _ = RequireString(operation, "text", path);
                if (operation["author"] is not null) _ = RequireString(operation, "author", path);
                if (nestedReviewer) _ = RequireString(operation, "author", path);
                break;
            case "addReviewBundle":
                RejectUnknown(operation, path, "op", "anchorContains", "spanText", "author",
                    "comment", "initials", "replyAuthor", "reply", "replyInitials", "footnote",
                    "bookmark", "linkText");
                foreach (var property in new[]
                {
                    "anchorContains", "spanText", "author", "comment", "initials", "replyAuthor",
                    "reply", "replyInitials", "footnote", "bookmark", "linkText",
                })
                    _ = RequireString(operation, property, path);
                break;
            case "fillContentControl":
                RejectUnknown(operation, path, "op", "tag", "text");
                _ = RequireString(operation, "tag", path);
                _ = RequireString(operation, "text", path);
                break;
            case "consolidate":
                if (nestedReviewer)
                    throw Error(path, "nested consolidate operations are not supported");
                RejectUnknown(operation, path, "op", "reviewers");
                var reviewers = RequireArray(operation, "reviewers", path);
                if (reviewers.Count < 2)
                    throw Error(path, "consolidate.reviewers must contain at least two operations");
                if (reviewers.Count > MaximumConsolidationReviewers)
                    throw Error(path,
                        $"consolidate.reviewers exceeds the {MaximumConsolidationReviewers}-reviewer limit");
                foreach (var reviewerNode in reviewers)
                {
                    var reviewer = reviewerNode as JsonObject
                        ?? throw Error(path, "consolidate.reviewers entries must be objects");
                    ValidateOperation(reviewer, path, nestedReviewer: true);
                }
                break;
            default:
                throw Error(path, $"unsupported baseline operation '{kind}'");
        }
    }

    private static void ValidateInvariant(
        JsonObject probe,
        string operation,
        JsonNode? expected,
        string path,
        int index)
    {
        var prefix = $"invariants[{index}]";
        var kind = RequireString(probe, "kind", path);
        switch (kind)
        {
            case "xmlCount":
                RejectUnknown(probe, path, "kind", "part", "xpath");
                if (!RequireString(probe, "part", path).StartsWith("/", StringComparison.Ordinal))
                    throw Error(path, $"{prefix}.probe.part must be an absolute OPC part name");
                _ = RequireString(probe, "xpath", path);
                RequireNumericComparison(operation, expected, path, prefix);
                break;
            case "textCount":
                RejectUnknown(probe, path, "kind", "text");
                _ = RequireString(probe, "text", path);
                RequireNumericComparison(operation, expected, path, prefix);
                break;
            case "partExists":
                RejectUnknown(probe, path, "kind", "part");
                if (!RequireString(probe, "part", path).StartsWith("/", StringComparison.Ordinal))
                    throw Error(path, $"{prefix}.probe.part must be an absolute OPC part name");
                if (operation != "equals" || expected is not JsonValue boolean
                    || !boolean.TryGetValue<bool>(out _))
                    throw Error(path, $"{prefix} partExists requires equals with a boolean expected value");
                break;
            default:
                throw Error(path, $"{prefix}.probe.kind '{kind}' is not supported");
        }
    }

    private static void RequireNumericComparison(
        string operation, JsonNode? expected, string path, string prefix)
    {
        if (operation is not ("equals" or "atLeast"))
            throw Error(path, $"{prefix} numeric probe requires equals or atLeast");
        if (expected is not JsonValue number
            || !number.TryGetValue<long>(out var expectedCount)
            || expectedCount < 0)
            throw Error(path,
                $"{prefix} numeric probe requires a non-negative integer expected value");
    }

    private static void RejectUnknown(JsonObject node, string path, params string[] allowed)
    {
        var allowedSet = allowed.ToHashSet(StringComparer.Ordinal);
        var unknown = node.Select(property => property.Key)
            .Where(property => !allowedSet.Contains(property))
            .Order(StringComparer.Ordinal)
            .ToList();
        if (unknown.Count != 0)
            throw Error(path, $"unknown properties: {string.Join(", ", unknown)}");
    }

    private static JsonObject ReadObject(string path, string kind)
    {
        if (!File.Exists(path))
            throw new ScenarioValidationException($"{kind} file does not exist: {path}");
        try
        {
            return JsonNode.Parse(ReadUtf8Bounded(path, MaximumJsonBytes, kind)) as JsonObject
                ?? throw Error(path, $"{kind} root must be a JSON object");
        }
        catch (ScenarioValidationException) { throw; }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            throw Error(path, $"invalid JSON: {exception.Message}");
        }
    }

    private static void RequireExactVersion(JsonObject node, string path, string expected)
    {
        var version = RequireString(node, "schemaVersion", path);
        if (version != expected)
            throw Error(path, $"unsupported schemaVersion '{version}'; expected '{expected}'");
    }

    private static JsonObject RequireObject(JsonObject parent, string name, string path) =>
        parent[name] as JsonObject ?? throw Error(path, $"{name} must be an object");

    private static JsonArray RequireArray(JsonObject parent, string name, string path)
    {
        var result = parent[name] as JsonArray
            ?? throw Error(path, $"{name} must be an array");
        if (result.Count > MaximumArrayItems)
            throw Error(path, $"{name} exceeds the {MaximumArrayItems}-item limit");
        return result;
    }

    private static string RequireString(JsonObject parent, string name, string path)
    {
        if (parent[name] is not JsonValue value || !value.TryGetValue<string>(out var result)
            || string.IsNullOrWhiteSpace(result))
            throw Error(path, $"{name} must be a non-empty string");
        if (result.Length > MaximumStringCharacters)
            throw Error(path,
                $"{name} exceeds the {MaximumStringCharacters}-character limit");
        return result;
    }

    private static string RequireNodeString(JsonNode? node, string path, string label)
    {
        if (node is not JsonValue value || !value.TryGetValue<string>(out var result)
            || string.IsNullOrWhiteSpace(result))
            throw Error(path, $"{label} must be a non-empty string");
        if (result.Length > MaximumStringCharacters)
            throw Error(path,
                $"{label} exceeds the {MaximumStringCharacters}-character limit");
        return result;
    }

    private static bool RequireBool(JsonObject parent, string name, string path)
    {
        if (parent[name] is not JsonValue value || !value.TryGetValue<bool>(out var result))
            throw Error(path, $"{name} must be a boolean");
        return result;
    }

    private static string RequireSha256(JsonObject parent, string name, string path)
    {
        var value = RequireString(parent, name, path);
        if (value.Length != 64 || value.Any(character => character is not
                (>= '0' and <= '9' or >= 'a' and <= 'f')))
            throw Error(path, $"{name} must be a lowercase 64-character hexadecimal SHA-256");
        return value;
    }

    private static int RequireInt(JsonObject parent, string name, string path)
    {
        if (parent[name] is not JsonValue value || !value.TryGetValue<int>(out var result))
            throw Error(path, $"{name} must be an integer");
        return result;
    }

    private static string ResolveUnderRoot(string root, string relative, string source)
    {
        if (Path.IsPathRooted(relative))
            throw Error(source, $"paths must be corpus-relative, got '{relative}'");
        var canonicalRoot = Path.TrimEndingDirectorySeparator(
            LegalEvaluationRunner.ResolveCanonicalPath(root));
        var fullRoot = Path.EndsInDirectorySeparator(canonicalRoot)
            ? canonicalRoot
            : canonicalRoot + Path.DirectorySeparatorChar;
        var fullPath = LegalEvaluationRunner.ResolveCanonicalPath(
            Path.Combine(canonicalRoot, relative));
        if (!fullPath.StartsWith(fullRoot, PathComparison))
            throw Error(source, $"path escapes corpus root: '{relative}'");
        return fullPath;
    }

    internal static string Sha256File(string path) =>
        Convert.ToHexString(SHA256.HashData(
            ReadBytesBounded(path, MaximumProvenanceDocumentBytes, "provenance document")))
            .ToLowerInvariant();

    private static string ReadUtf8Bounded(string path, int maximumBytes, string label) =>
        new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
            .GetString(ReadBytesBounded(path, maximumBytes, label));

    private static byte[] ReadBytesBounded(string path, int maximumBytes, string label)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 81920, FileOptions.SequentialScan);
        if (stream.Length > maximumBytes)
            throw Error(path, $"{label} exceeds the {maximumBytes}-byte read limit");
        var bytes = new byte[checked((int)stream.Length)];
        stream.ReadExactly(bytes);
        if (stream.ReadByte() != -1)
            throw Error(path, $"{label} changed while it was being read");
        return bytes;
    }

    private static ScenarioValidationException Error(string path, string message) =>
        new($"{path}: {message}");

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
