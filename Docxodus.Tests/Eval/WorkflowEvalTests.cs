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
    /// <summary>
    /// Scenarios the current run executes: the fast subset always, plus the corpus tier when
    /// <c>DOCXODUS_RUN_EVAL_CORPUS=1</c>.
    /// </summary>
    public static TheoryData<string> Scenarios()
    {
        var data = new TheoryData<string>();
        foreach (var file in ExecutableScenarioFiles())
            data.Add(Path.GetFileNameWithoutExtension(file));
        return data;
    }

    /// <summary>
    /// Every scenario in the repository, both tiers. Declaration checks (vacuity, schema
    /// conformance) are cheap, so a corpus scenario cannot sit malformed until the weekly run.
    /// </summary>
    public static TheoryData<string> DeclaredScenarios()
    {
        var data = new TheoryData<string>();
        foreach (var file in AllScenarioFiles())
            data.Add(Path.GetFileNameWithoutExtension(file));
        return data;
    }

    private static IEnumerable<string> ExecutableScenarioFiles() =>
        EvalHarness.RunCorpusTier ? AllScenarioFiles() : EvalHarness.ScenarioFiles();

    private static IEnumerable<string> AllScenarioFiles() =>
        EvalHarness.ScenarioFiles().Concat(EvalHarness.CorpusScenarioFiles());

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
        CheckTrackedRevisions(invariants, outcome, failures);
        CheckComments(invariants, outcome, failures);
        CheckReversibility(invariants, outcome, failures);
        CheckRendering(invariants, outcome, failures);

        // The scorecard is written on success too: it is the machine-readable engine baseline a
        // later agent-scored run of the same scenario is compared against.
        WriteScorecard(outcome, failures);
        if (failures.Count == 0)
            return;

        var artifacts = WriteArtifacts(outcome);
        Assert.Fail(
            $"scenario '{id}' violated {failures.Count} invariant(s):{Environment.NewLine}"
            + string.Join(Environment.NewLine, failures.Select(failure => "  - " + failure))
            + $"{Environment.NewLine}artifacts: {artifacts}");
    }

    /// <summary>
    /// The invariant keys each group may use. This is the checkers' vocabulary: a key outside it
    /// is silently ignored by every checker, so EV002 rejects it rather than letting a typo turn
    /// an invariant off.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string[]> InvariantVocabulary =
        new Dictionary<string, string[]>(StringComparer.Ordinal)
        {
            ["taskCompletion"] = ["textPresent", "textAbsent"],
            ["targetPrecision"] =
                ["changedAnchorsAtLeast", "changedAnchorsAtMost", "changedPartsAtMost"],
            ["collateral"] =
                ["partsAdded", "partsRemoved", "textPreserved", "changedPartsMustBeWithin"],
            ["validity"] = ["decisionIn"],
            ["trackedRevisions"] = ["countAtLeast", "countAtMost", "authorsMustInclude"],
            ["comments"] =
                ["countAtLeast", "countAtMost", "authorsMustInclude", "repliesAtLeast", "resolvedAtLeast"],
            ["reversibility"] =
            [
                "mode", "mustSucceed", "generatedRevisionsAtLeast", "pathsMustComplete",
                "preExistingMustBePreserved", "rejectMustRestoreBaseline",
            ],
            ["rendering"] = ["htmlMustContain"],
        };

    /// <summary>
    /// A scenario with no scoreable expectation is a scenario that cannot fail. Adding one has to
    /// mean writing down what "done" and "only that" look like, not eyeballing the output.
    /// </summary>
    [Theory]
    [MemberData(nameof(DeclaredScenarios))]
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
            invariants.TryGetProperty("targetPrecision", out var precision),
            $"scenario '{id}' must declare targetPrecision: how much it was allowed to touch.");
        Assert.True(
            precision.EnumerateObject().Any(),
            $"scenario '{id}' declares an empty targetPrecision: it bounds nothing.");
        Assert.True(
            invariants.TryGetProperty("collateral", out var collateral),
            $"scenario '{id}' must declare collateral: what had to survive untouched.");
        Assert.NotEmpty(collateral.GetProperty("textPreserved").EnumerateArray());

        // Every declared invariant must be one the checkers actually read. The schema says
        // additionalProperties: false; this enforces it where the tests run.
        foreach (var group in invariants.EnumerateObject())
        {
            Assert.True(
                InvariantVocabulary.TryGetValue(group.Name, out var known),
                $"scenario '{id}' declares unknown invariant group '{group.Name}'.");
            Assert.True(
                group.Value.EnumerateObject().Any(),
                $"scenario '{id}' declares an empty '{group.Name}': it asserts nothing.");
            foreach (var property in group.Value.EnumerateObject())
                Assert.True(
                    known!.Contains(property.Name, StringComparer.Ordinal),
                    $"scenario '{id}': '{group.Name}.{property.Name}' is not an invariant any "
                    + "checker reads — a typo here would silently switch the check off.");
        }
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
        var names = AllScenarioFiles()
            .Select(file => EvalHarness.LoadJson(file).GetProperty("fixture").GetString()!)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        Assert.NotEmpty(names);

        foreach (var name in names)
        {
            var script = EvalHarness.LoadJson(
                Path.Combine(EvalHarness.CorpusRoot, "fixtures", $"{name}.json"));
            // Each fixture declares its own content anchor: identical-but-empty projections must
            // not count as reproducible, and a hardcoded anchor here would block a second fixture.
            var expected = script.GetProperty("expectedContent").EnumerateArray()
                .Select(value => value.GetString()!)
                .ToList();
            Assert.NotEmpty(expected);

            var first = EvalHarness.BuildFixture(name);
            var second = EvalHarness.BuildFixture(name);
            Assert.True(first.Length > 1_000, $"fixture '{name}' did not build a real package");
            Assert.Equal(EvalHarness.PartUris(first), EvalHarness.PartUris(second));
            var projection = EvalHarness.TextProjection(first);
            foreach (var needle in expected)
                Assert.Contains(needle, projection, StringComparison.Ordinal);
            Assert.Equal(projection, EvalHarness.TextProjection(second));
        }
    }

    /// <summary>
    /// Every scenario file, both tiers, must satisfy the published schema. Until now the schema
    /// was only the corpus-root marker: its enums, patterns, and closed objects were prose. A
    /// negative control proves the validator can actually reject (an unknown invariant group and
    /// a bad target mode), so a regression that turns it into a yes-machine fails here.
    /// </summary>
    [Fact]
    public void EV004_ScenariosConformToTheirSchema()
    {
        var schema = EvalHarness.LoadJson(
            Path.Combine(EvalHarness.CorpusRoot, "scenario.schema.json"));

        foreach (var file in AllScenarioFiles())
        {
            var errors = EvalSchemaValidator.Validate(EvalHarness.LoadJson(file), schema);
            Assert.True(
                errors.Count == 0,
                $"{Path.GetFileName(file)} violates scenario.schema.json:{Environment.NewLine}"
                + string.Join(Environment.NewLine, errors.Select(error => "  - " + error)));
        }

        using var invalid = JsonDocument.Parse("""
            {
              "id": "Bad Id With Spaces",
              "title": "x",
              "fixture": "master-services-agreement",
              "steps": [{ "tool": "docxodus_edit", "target": { "mode": "nonsense", "query": "q" }, "args": {} }],
              "invariants": { "notAGroup": { "x": 1 } }
            }
            """);
        var rejected = EvalSchemaValidator.Validate(invalid.RootElement, schema);
        Assert.Contains(rejected, error => error.Contains("pattern", StringComparison.Ordinal));
        Assert.Contains(rejected, error => error.Contains("mode", StringComparison.Ordinal));
        Assert.Contains(rejected, error => error.Contains("notAGroup", StringComparison.Ordinal));
    }

    /// <summary>
    /// The schema's invariant vocabulary and the checkers' vocabulary are two views of one
    /// contract. The C# side stays authoritative for "what a checker actually reads"; this pins
    /// the schema to it in both directions, so neither can grow a key the other ignores.
    /// </summary>
    [Fact]
    public void EV005_SchemaAndCheckerInvariantVocabulariesAgree()
    {
        var schema = EvalHarness.LoadJson(
            Path.Combine(EvalHarness.CorpusRoot, "scenario.schema.json"));
        var groups = schema.GetProperty("$defs").GetProperty("invariants").GetProperty("properties");

        var schemaGroups = groups.EnumerateObject().Select(group => group.Name)
            .OrderBy(name => name, StringComparer.Ordinal).ToList();
        Assert.Equal(
            InvariantVocabulary.Keys.OrderBy(name => name, StringComparer.Ordinal).ToList(),
            schemaGroups);

        foreach (var group in groups.EnumerateObject())
        {
            var schemaKeys = group.Value.GetProperty("properties").EnumerateObject()
                .Select(property => property.Name)
                .OrderBy(name => name, StringComparer.Ordinal).ToList();
            Assert.Equal(
                InvariantVocabulary[group.Name].OrderBy(name => name, StringComparer.Ordinal).ToList(),
                schemaKeys);
        }
    }

    /// <summary>
    /// Negative controls for the invariant checkers added for the #466 close-out: each one must
    /// record a failure when its expectation is violated and stay silent when it is met. A
    /// checker that cannot fail is decoration, not an invariant.
    /// </summary>
    [Fact]
    public void EV006_NewInvariantChecksFailWhenViolated()
    {
        var outcome = SyntheticOutcome();

        AssertCheck(
            outcome,
            """{ "collateral": { "changedPartsMustBeWithin": ["/word/document.xml"] } }""",
            CheckCollateral,
            expectFailure: true,
            "the change set touched /word/header1.xml");
        AssertCheck(
            outcome,
            """{ "collateral": { "changedPartsMustBeWithin": ["/word/document.xml", "/word/header1.xml"] } }""",
            CheckCollateral,
            expectFailure: false,
            "both changed parts are allowlisted");

        AssertCheck(
            outcome,
            """{ "trackedRevisions": { "countAtLeast": 2, "authorsMustInclude": ["Prior Reviewer", "Absent Author"] } }""",
            CheckTrackedRevisions,
            expectFailure: true,
            "only one revision, and Absent Author has none");
        AssertCheck(
            outcome,
            """{ "trackedRevisions": { "countAtLeast": 1, "countAtMost": 1, "authorsMustInclude": ["Prior Reviewer"] } }""",
            CheckTrackedRevisions,
            expectFailure: false,
            "the live revision matches every bound");

        AssertCheck(
            outcome,
            """{ "comments": { "repliesAtLeast": 2, "resolvedAtLeast": 2 } }""",
            CheckComments,
            expectFailure: true,
            "one reply and one resolved comment fall short of two");
        AssertCheck(
            outcome,
            """{ "comments": { "countAtLeast": 2, "authorsMustInclude": ["First Reviewer"], "repliesAtLeast": 1, "resolvedAtLeast": 1 } }""",
            CheckComments,
            expectFailure: false,
            "the comment thread matches every bound");
    }

    private static void AssertCheck(
        EvalOutcome outcome,
        string invariantsJson,
        Action<JsonElement, EvalOutcome, List<string>> checker,
        bool expectFailure,
        string because)
    {
        using var document = JsonDocument.Parse(invariantsJson);
        var failures = new List<string>();
        checker(document.RootElement, outcome, failures);
        if (expectFailure)
            Assert.True(failures.Count > 0, $"expected a failure ({because}), got none");
        else
            Assert.True(
                failures.Count == 0,
                $"expected no failure ({because}), got: {string.Join("; ", failures)}");
    }

    private static EvalOutcome SyntheticOutcome() => new()
    {
        Id = "synthetic-negative-control",
        OpeningBytes = [],
        DeliverableBytes = [],
        ChangedAnchors = ["p:body:00000001"],
        ChangedParts = ["/word/document.xml", "/word/header1.xml"],
        PartsAdded = 0,
        PartsRemoved = 0,
        ValidityDecision = "passed",
        VerificationJson = "{}",
        TextMatchCounts = new Dictionary<string, int>(StringComparer.Ordinal),
        Text = string.Empty,
        Html = string.Empty,
        SemanticChangesJson = "{}",
        RevisionsJson = """
            {"revisions":[{"id":"r1","type":"insert","family":"run","constituentIds":[],
            "constituentKeys":[],"author":"Prior Reviewer","text":"x","partUri":"word/document.xml",
            "scope":"body","affectedAnchors":[],"resolutionStatus":"open"}]}
            """,
        CommentsJson = """
            {"comments":[
            {"anchorId":"cmt:cmt:1","author":"First Reviewer","text":"Q?","resolved":true},
            {"anchorId":"cmt:cmt:2","author":"Second Reviewer","text":"A.","parentAnchorId":"cmt:cmt:1"}]}
            """,
    };

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

        // Identity, not cardinality: an edit that stayed within its part budget but landed in
        // the wrong part (a header instead of the body) fails here.
        if (collateral.TryGetProperty("changedPartsMustBeWithin", out var allowlist))
        {
            var allowed = allowlist.EnumerateArray()
                .Select(value => value.GetString() ?? string.Empty)
                .ToHashSet(StringComparer.Ordinal);
            var outside = outcome.ChangedParts
                .Where(part => !allowed.Contains(part))
                .ToList();
            if (outside.Count > 0)
                failures.Add(
                    "collateral: change set touched part(s) outside the allowlist "
                    + $"[{string.Join(", ", outside)}], allowed [{string.Join(", ", allowed)}]");
        }
    }

    /// <summary>
    /// Both of the next two groups are agent-surface facts: they parse the same
    /// docxodus_track_changes/docxodus_comment list responses an agent would read, captured from
    /// the delivered session, not the typed .NET view.
    /// </summary>
    private static void CheckTrackedRevisions(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("trackedRevisions", out var expected))
            return;

        using var document = JsonDocument.Parse(outcome.RevisionsJson);
        var revisions = document.RootElement.GetProperty("revisions").EnumerateArray().ToList();
        CheckCountBounds(expected, revisions.Count, "trackedRevisions", "revision", failures);
        CheckAuthors(
            expected,
            revisions.Select(item => item.GetProperty("author").GetString() ?? string.Empty),
            "trackedRevisions",
            failures);
    }

    private static void CheckComments(
        JsonElement invariants, EvalOutcome outcome, List<string> failures)
    {
        if (!invariants.TryGetProperty("comments", out var expected))
            return;

        using var document = JsonDocument.Parse(outcome.CommentsJson);
        var comments = document.RootElement.GetProperty("comments").EnumerateArray().ToList();
        CheckCountBounds(expected, comments.Count, "comments", "comment", failures);
        CheckAuthors(
            expected,
            comments.Select(item => item.GetProperty("author").GetString() ?? string.Empty),
            "comments",
            failures);

        var replies = comments.Count(item => item.TryGetProperty("parentAnchorId", out _));
        if (expected.TryGetProperty("repliesAtLeast", out var repliesAtLeast)
            && replies < repliesAtLeast.GetInt32())
            failures.Add(
                $"comments: {replies} repl(ies) present, expected at least {repliesAtLeast.GetInt32()}");

        var resolved = comments.Count(item =>
            item.TryGetProperty("resolved", out var flag) && flag.ValueKind == JsonValueKind.True);
        if (expected.TryGetProperty("resolvedAtLeast", out var resolvedAtLeast)
            && resolved < resolvedAtLeast.GetInt32())
            failures.Add(
                $"comments: {resolved} resolved comment(s), expected at least "
                + $"{resolvedAtLeast.GetInt32()}");
    }

    private static void CheckCountBounds(
        JsonElement expected, int count, string group, string noun, List<string> failures)
    {
        if (expected.TryGetProperty("countAtLeast", out var atLeast)
            && count < atLeast.GetInt32())
            failures.Add($"{group}: {count} {noun}(s) present, expected at least {atLeast.GetInt32()}");
        if (expected.TryGetProperty("countAtMost", out var atMost)
            && count > atMost.GetInt32())
            failures.Add($"{group}: {count} {noun}(s) present, expected at most {atMost.GetInt32()}");
    }

    private static void CheckAuthors(
        JsonElement expected, IEnumerable<string> actual, string group, List<string> failures)
    {
        if (!expected.TryGetProperty("authorsMustInclude", out var required))
            return;

        var present = actual.ToHashSet(StringComparer.Ordinal);
        foreach (var author in required.EnumerateArray())
        {
            var name = author.GetString() ?? string.Empty;
            if (!present.Contains(name))
                failures.Add(
                    $"{group}: no entry by author \"{name}\" "
                    + $"[present: {string.Join(", ", present.OrderBy(a => a, StringComparer.Ordinal))}]");
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

        // Restoration, not just completion: rejecting only the generated revisions must be
        // semantically the opening package. This is the half of reversibility the derived
        // intended final cannot blur — the baseline is stated up front, never derived.
        if (Flag(reversibility, "rejectMustRestoreBaseline")
            && proof.RejectToBaseline?.ModeledSemantic.Equivalent != true)
            failures.Add(
                "reversibility: rejecting the generated revisions does not restore the baseline "
                + $"[{Describe(proof.RejectToBaseline)}, semanticChanges="
                + $"{proof.RejectToBaseline?.ModeledSemantic.ChangeCount}]");

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
        var directory = ArtifactDirectory(outcome.Id);
        File.WriteAllBytes(Path.Combine(directory, "opening.docx"), outcome.OpeningBytes);
        File.WriteAllBytes(Path.Combine(directory, "delivered.docx"), outcome.DeliverableBytes);
        File.WriteAllText(Path.Combine(directory, "delivered.html"), outcome.Html, Encoding.UTF8);
        File.WriteAllText(Path.Combine(directory, "delivered.txt"), outcome.Text, Encoding.UTF8);
        File.WriteAllText(
            Path.Combine(directory, "semantic-changes.json"),
            outcome.SemanticChangesJson,
            Encoding.UTF8);
        File.WriteAllText(
            Path.Combine(directory, "verification.json"),
            outcome.VerificationJson,
            Encoding.UTF8);
        File.WriteAllText(
            Path.Combine(directory, "revisions.json"), outcome.RevisionsJson, Encoding.UTF8);
        File.WriteAllText(
            Path.Combine(directory, "comments.json"), outcome.CommentsJson, Encoding.UTF8);
        if (outcome.Reversibility is { } proof)
        {
            File.WriteAllText(
                Path.Combine(directory, "reversibility-proof.json"),
                proof.ToJson(),
                Encoding.UTF8);
        }

        return directory;
    }

    private static string ArtifactDirectory(string id)
    {
        var directory = Path.Combine(
            Environment.GetEnvironmentVariable("DOCXODUS_EVAL_ARTIFACTS")
                ?? Path.Combine(Path.GetTempPath(), "docxodus-eval-artifacts"),
            id);
        Directory.CreateDirectory(directory);
        return directory;
    }

    /// <summary>
    /// The per-scenario engine baseline, written pass or fail: every metric the run measured,
    /// in one machine-readable file. An agent-scored run of the same scenario is compared
    /// against this — same fixture, same invariants, measured by the same scripted caller.
    /// </summary>
    private static void WriteScorecard(EvalOutcome outcome, IReadOnlyList<string> failures)
    {
        var scorecard = new
        {
            scenario = outcome.Id,
            passed = failures.Count == 0,
            invariantFailures = failures,
            metrics = new
            {
                changedAnchors = outcome.ChangedAnchors,
                changedParts = outcome.ChangedParts,
                partsAdded = outcome.PartsAdded,
                partsRemoved = outcome.PartsRemoved,
                validityDecision = outcome.ValidityDecision,
                textMatchCounts = outcome.TextMatchCounts,
                generatedRevisionCount = outcome.GeneratedRevisionCount,
                reversibility = outcome.Reversibility is { } proof
                    ? new
                    {
                        acceptCompleted = proof.AcceptToFinal?.Completed,
                        rejectCompleted = proof.RejectToBaseline?.Completed,
                        rejectSemanticallyRestoresBaseline =
                            proof.RejectToBaseline?.ModeledSemantic.Equivalent,
                        preExistingPreserved =
                            proof.AcceptToFinal?.PreExistingRevisionsPreserved == true
                            && proof.RejectToBaseline?.PreExistingRevisionsPreserved == true,
                        fullPackageSuccess = proof.Success,
                    }
                    : null,
            },
        };
        File.WriteAllText(
            Path.Combine(ArtifactDirectory(outcome.Id), "scorecard.json"),
            JsonSerializer.Serialize(
                scorecard, new JsonSerializerOptions { WriteIndented = true }),
            Encoding.UTF8);
    }
}
