// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Eval;

/// <summary>
/// The deterministic, fast subset of the #466 workflow evaluation suite: every scenario in
/// <c>eval/scenarios</c>, executed by the scripted caller and scored against its declared
/// invariants. Nothing here depends on a PDF, a browser, or a network, so it runs on every push;
/// the larger opt-in corpus and the agent-scored runs build on this contract rather than replacing
/// it.
/// </summary>
public sealed class WorkflowEvalTests
{
    public static TheoryData<string> Scenarios()
    {
        var data = new TheoryData<string>();
        foreach (var file in EvalHarness.ScenarioFiles())
            data.Add(Path.GetFileNameWithoutExtension(file));
        return data;
    }

    [Theory]
    [MemberData(nameof(Scenarios))]
    public void EV001_ScenarioMeetsItsDeclaredInvariants(string id)
    {
        var scenario = EvalHarness.LoadScenario(id);
        var outcome = EvalHarness.Run(scenario);
        var failures = new List<string>();
        var invariants = scenario.GetProperty("invariants");

        CheckTaskCompletion(invariants, outcome, failures);
        CheckTargetPrecision(invariants, outcome, failures);
        CheckCollateral(invariants, outcome, failures);
        CheckValidity(invariants, outcome, failures);
        CheckReversibility(invariants, outcome, failures);
        CheckRendering(invariants, outcome, failures);

        if (failures.Count == 0)
            return;

        var artifacts = WriteArtifacts(outcome);
        Assert.Fail(
            $"scenario '{id}' violated {failures.Count} invariant(s):{Environment.NewLine}"
            + string.Join(Environment.NewLine, failures.Select(failure => "  - " + failure))
            + $"{Environment.NewLine}artifacts: {artifacts}");
    }

    /// <summary>
    /// A scenario with no scoreable expectation is a scenario that cannot fail. Adding one has to
    /// mean writing down what "done" and "only that" look like, not eyeballing the output.
    /// </summary>
    [Theory]
    [MemberData(nameof(Scenarios))]
    public void EV002_ScenarioDeclaresScoreableExpectations(string id)
    {
        var scenario = EvalHarness.LoadScenario(id);

        Assert.Equal(id, scenario.GetProperty("id").GetString());
        Assert.False(string.IsNullOrWhiteSpace(scenario.GetProperty("title").GetString()));
        Assert.NotEmpty(scenario.GetProperty("steps").EnumerateArray());

        var invariants = scenario.GetProperty("invariants");
        Assert.True(
            invariants.TryGetProperty("taskCompletion", out var completion),
            $"scenario '{id}' must declare taskCompletion: what change had to land.");
        Assert.NotEmpty(completion.GetProperty("textPresent").EnumerateArray());
        Assert.True(
            invariants.TryGetProperty("targetPrecision", out _),
            $"scenario '{id}' must declare targetPrecision: how much it was allowed to touch.");
        Assert.True(
            invariants.TryGetProperty("collateral", out var collateral),
            $"scenario '{id}' must declare collateral: what had to survive untouched.");
        Assert.NotEmpty(collateral.GetProperty("textPreserved").EnumerateArray());
    }

    /// <summary>
    /// Every fixture a scenario names must build, and must be reproducible: the same script twice
    /// has to produce the same <em>document</em>. A corpus that drifts between runs cannot support
    /// a baseline that later agent runs are scored against.
    ///
    /// <para>Reproducible means same content and same package shape, not same bytes and not the
    /// same package digest. Anchor ids and revision-save ids are minted per build, so two honest
    /// builds of one script differ at the digest layer while being the same document. Comparing
    /// digests here would assert that bookkeeping is stable, which it is not and need not be.</para>
    /// </summary>
    [Fact]
    public void EV003_FixturesAreDeterministicAndComplete()
    {
        var names = EvalHarness.ScenarioFiles()
            .Select(file => EvalHarness.LoadJson(file).GetProperty("fixture").GetString()!)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        Assert.NotEmpty(names);

        foreach (var name in names)
        {
            var first = EvalHarness.BuildFixture(name);
            var second = EvalHarness.BuildFixture(name);
            Assert.True(first.Length > 1_000, $"fixture '{name}' did not build a real package");
            Assert.Equal(PartUris(first), PartUris(second));
            Assert.Equal(EvalHarness.TextProjection(first), EvalHarness.TextProjection(second));
        }
    }

    private static IReadOnlyList<string> PartUris(byte[] packageBytes)
    {
        using var manifest = JsonDocument.Parse(
            VerificationOps.GeneratePackageManifest(packageBytes));
        return manifest.RootElement.GetProperty("entries").EnumerateArray()
            .Select(entry => entry.GetProperty("uri").GetString() ?? string.Empty)
            .OrderBy(uri => uri, StringComparer.Ordinal)
            .ToList();
    }

    private static void CheckTaskCompletion(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("taskCompletion", out var completion))
            return;

        foreach (var needle in Strings(completion, "textPresent"))
        {
            if (Matches(outcome, needle) == 0)
                failures.Add($"taskCompletion: expected text not found: \"{needle}\"");
        }

        foreach (var needle in Strings(completion, "textAbsent"))
        {
            if (Matches(outcome, needle) > 0)
                failures.Add(
                    $"taskCompletion: text should have been replaced, still in "
                    + $"{Matches(outcome, needle)} block(s): \"{needle}\"");
        }
    }

    private static void CheckTargetPrecision(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("targetPrecision", out var precision))
            return;

        var changed = outcome.ChangedAnchors;
        if (precision.TryGetProperty("changedAnchorsAtLeast", out var atLeast)
            && changed.Count < atLeast.GetInt32())
            failures.Add(
                $"targetPrecision: {changed.Count} anchor(s) changed, expected at least "
                + $"{atLeast.GetInt32()} [{string.Join(", ", changed)}]");

        if (precision.TryGetProperty("changedAnchorsAtMost", out var atMost)
            && changed.Count > atMost.GetInt32())
            failures.Add(
                $"targetPrecision: {changed.Count} anchor(s) changed, expected at most "
                + $"{atMost.GetInt32()} [{string.Join(", ", changed)}]");

        if (precision.TryGetProperty("changedPartsAtMost", out var partsAtMost)
            && outcome.ChangedParts.Count > partsAtMost.GetInt32())
            failures.Add(
                $"targetPrecision: {outcome.ChangedParts.Count} part(s) changed, expected at most "
                + $"{partsAtMost.GetInt32()} [{string.Join(", ", outcome.ChangedParts)}]");
    }

    private static void CheckCollateral(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("collateral", out var collateral))
            return;

        if (collateral.TryGetProperty("partsAdded", out var added)
            && outcome.PartsAdded != added.GetInt32())
            failures.Add(
                $"collateral: {outcome.PartsAdded} package part(s) added, expected {added.GetInt32()}");

        if (collateral.TryGetProperty("partsRemoved", out var removed)
            && outcome.PartsRemoved != removed.GetInt32())
            failures.Add(
                $"collateral: {outcome.PartsRemoved} package part(s) removed, expected {removed.GetInt32()}");

        foreach (var needle in Strings(collateral, "textPreserved"))
        {
            if (Matches(outcome, needle) == 0)
                failures.Add($"collateral: unrelated content was damaged, lost: \"{needle}\"");
        }
    }

    private static void CheckValidity(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("validity", out var validity)
            || !validity.TryGetProperty("decisionIn", out var accepted))
            return;

        var allowed = accepted.EnumerateArray()
            .Select(value => value.GetString() ?? string.Empty)
            .ToList();
        if (!allowed.Contains(outcome.ValidityDecision))
            failures.Add(
                $"validity: deliverable gate decided '{outcome.ValidityDecision}', "
                + $"expected one of [{string.Join(", ", allowed)}]");
    }

    private static void CheckReversibility(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("reversibility", out var reversibility)
            || !reversibility.TryGetProperty("mode", out var mode)
            || mode.GetString() == "none")
            return;

        if (outcome.Reversibility is not { } proof)
        {
            failures.Add("reversibility: requested but no proof was produced");
            return;
        }

        if (reversibility.TryGetProperty("generatedRevisionsAtLeast", out var atLeast)
            && outcome.GeneratedRevisionCount < atLeast.GetInt32())
            failures.Add(
                $"reversibility: {outcome.GeneratedRevisionCount} generated revision(s), expected "
                + $"at least {atLeast.GetInt32()} — the edit was not recorded as tracked changes");

        if (Flag(reversibility, "pathsMustComplete") && !BothPathsCompleted(proof))
            failures.Add(
                "reversibility: a proof path did not complete "
                + $"[acceptToFinal={Describe(proof.AcceptToFinal)}, "
                + $"rejectToBaseline={Describe(proof.RejectToBaseline)}]");

        if (Flag(reversibility, "preExistingMustBePreserved")
            && (proof.AcceptToFinal?.PreExistingRevisionsPreserved == false
                || proof.RejectToBaseline?.PreExistingRevisionsPreserved == false))
            failures.Add(
                "reversibility: resolving the generated revisions consumed pre-existing "
                + "review state");

        // Full package equivalence is opt-in. See eval/README.md: a session-authored redline is
        // not yet expected to reach it, and asserting it here would be asserting the derivation
        // of the intended final rather than the engine's reversibility.
        if (!reversibility.TryGetProperty("mustSucceed", out var required)
            || !required.GetBoolean()
            || proof.Success)
            return;

        var reasons = proof.Findings
            .Where(finding => finding.Severity == VerificationFindingSeverity.Error)
            .Select(finding => finding.Code)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        // The first divergent part is usually the whole diagnosis, so name it in the failure.
        if (proof.AcceptToFinal?.FirstDivergence?.PartUri is { Length: > 0 } divergentPart)
            reasons.Add(divergentPart);
        failures.Add(
            "reversibility: the redline does not accept to the intended final and reject to the "
            + $"baseline [{string.Join(", ", reasons)}]");
    }

    /// <summary>A reversibility flag that defaults to on when the scenario does not name it.</summary>
    private static bool Flag(JsonElement reversibility, string property) =>
        !reversibility.TryGetProperty(property, out var value) || value.GetBoolean();

    private static bool BothPathsCompleted(RedlineReversibilityProof proof) =>
        proof.AcceptToFinal?.Completed == true && proof.RejectToBaseline?.Completed == true;

    private static string Describe(RedlineProofPathResult? path) =>
        path is null ? "absent" : path.Completed ? "completed" : "incomplete";

    private static void CheckRendering(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("rendering", out var rendering))
            return;

        foreach (var needle in Strings(rendering, "htmlMustContain"))
        {
            if (!outcome.Html.Contains(needle, StringComparison.Ordinal))
                failures.Add($"rendering: HTML projection is missing \"{needle}\"");
        }
    }

    private static int Matches(EvalOutcome outcome, string needle) =>
        outcome.TextMatchCounts.TryGetValue(needle, out var count) ? count : 0;

    private static IEnumerable<string> Strings(JsonElement parent, string property) =>
        parent.TryGetProperty(property, out var array)
            ? array.EnumerateArray().Select(value => value.GetString() ?? string.Empty)
            : Enumerable.Empty<string>();

    /// <summary>
    /// Preserve everything needed to diagnose a failure without re-running it: both packages, the
    /// rendered HTML, the semantic change set, and the reversibility proof when one was produced.
    /// </summary>
    private static string WriteArtifacts(EvalOutcome outcome)
    {
        var directory = Path.Combine(
            Environment.GetEnvironmentVariable("DOCXODUS_EVAL_ARTIFACTS")
                ?? Path.Combine(Path.GetTempPath(), "docxodus-eval-artifacts"),
            outcome.Id);
        Directory.CreateDirectory(directory);

        File.WriteAllBytes(Path.Combine(directory, "opening.docx"), outcome.OpeningBytes);
        File.WriteAllBytes(Path.Combine(directory, "delivered.docx"), outcome.DeliverableBytes);
        File.WriteAllText(Path.Combine(directory, "delivered.html"), outcome.Html, Encoding.UTF8);
        File.WriteAllText(Path.Combine(directory, "delivered.txt"), outcome.Text, Encoding.UTF8);
        File.WriteAllText(
            Path.Combine(directory, "semantic-changes.json"),
            outcome.SemanticChangesJson,
            Encoding.UTF8);
        if (outcome.Reversibility is { } proof)
        {
            File.WriteAllText(
                Path.Combine(directory, "reversibility-proof.json"),
                proof.ToJson(),
                Encoding.UTF8);
        }

        return directory;
    }
}
