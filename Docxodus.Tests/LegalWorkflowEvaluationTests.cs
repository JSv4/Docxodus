// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using LegalEval;
using Xunit;

namespace Docxodus.Tests;

[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class LegalWorkflowEvaluationCollection
{
    public const string Name = "Legal workflow evaluation artifacts";
}

[Collection(LegalWorkflowEvaluationCollection.Name)]
public sealed class LegalWorkflowEvaluationTests
{
    private static readonly string RepositoryRoot = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "../../../../"));
    private static readonly string CorpusPath = Path.Combine(
        RepositoryRoot, "eval", "legal", "corpus.json");
    private static readonly string ArtifactRoot = Path.Combine(
        RepositoryRoot, "TestResults", "legal-eval", "xunit");

    [Fact]
    public void Corpus_has_nine_explicit_provenanced_scenarios()
    {
        var corpus = ScenarioLoader.LoadCorpus(CorpusPath);

        Assert.Equal(new[]
        {
            "compare-consolidate",
            "content-control-fill",
            "defined-term-targeting",
            "numbered-clause-insertion",
            "preexisting-review-state",
            "preserve-package-structures",
            "review-thread-note-cross-reference",
            "table-economics",
            "tracked-notice-amendment",
        }, corpus.Scenarios.Select(value => value.Id).Order(StringComparer.Ordinal));
        Assert.Equal(6, corpus.Scenarios.Count(value => value.Tier == EvalTier.Fast));
        Assert.Equal(3, corpus.Scenarios.Count(value => value.Tier == EvalTier.Full));
        Assert.All(corpus.Scenarios, scenario =>
        {
            Assert.NotEmpty(scenario.Constraints);
            Assert.NotEmpty(scenario.ExpectedOutputs);
            Assert.NotEmpty(scenario.BaselineOperations);
            Assert.NotEmpty(scenario.Invariants);
            Assert.NotEmpty(scenario.ChangeBudget.AllowedChangedParts);
            Assert.True(corpus.Provenance.ContainsKey(scenario.Fixture.ProvenanceId));
            Assert.True(corpus.ExpectedDocumentProvenance.ContainsKey(
                scenario.ExpectedDocument.ProvenanceId));
            Assert.True(File.Exists(scenario.Fixture.Path));
            Assert.True(File.Exists(scenario.ExpectedDocument.Path));
            Assert.Equal(scenario.Fixture.SourceSha256, Sha256File(scenario.Fixture.Path));
            Assert.Equal(scenario.ExpectedDocument.SourceSha256,
                Sha256File(scenario.ExpectedDocument.Path));
            Assert.All(scenario.ExpectedOutputs, output => Assert.Contains(output.Id, new[]
            {
                "candidate-docx",
                "semantic-diff",
                "after-html",
                "redline-docx",
                "candidate-pdf",
                "redline-proof-v1",
            }));
        });
        Assert.All(corpus.Provenance.Values, provenance =>
        {
            Assert.False(string.IsNullOrWhiteSpace(provenance.License));
            Assert.False(string.IsNullOrWhiteSpace(provenance.RedistributionPermission));
            Assert.True(File.Exists(provenance.SourcePath));
            Assert.NotNull(provenance.RecipePath);
            Assert.True(File.Exists(provenance.RecipePath));
        });
        Assert.Equal(9, corpus.ExpectedDocumentProvenance.Count);
        Assert.All(corpus.ExpectedDocumentProvenance.Values, provenance =>
        {
            Assert.Equal("not-independently-reviewed", provenance.ReviewStatus);
            Assert.Contains("not an independently authored or cross-engine correctness oracle",
                provenance.ReviewNotes, StringComparison.Ordinal);
            Assert.False(string.IsNullOrWhiteSpace(provenance.RedistributionPermission));
            Assert.True(File.Exists(provenance.SourcePath));
            Assert.Equal(provenance.SourceSha256, Sha256File(provenance.SourcePath));
        });
    }

    [Fact]
    public void Loader_rejects_a_scenario_without_deterministic_invariants()
    {
        var corpus = ScenarioLoader.LoadCorpus(CorpusPath);
        var source = corpus.Scenarios.Single(value => value.Id == "defined-term-targeting");
        var node = JsonNode.Parse(File.ReadAllText(source.FilePath))!.AsObject();
        node["invariants"] = new JsonArray();
        var inspectionDirectory = Path.Combine(ArtifactRoot, "loader-validation");
        Directory.CreateDirectory(inspectionDirectory);
        var invalidPath = Path.Combine(inspectionDirectory, "missing-invariants.scenario.json");
        File.WriteAllText(invalidPath, node.ToJsonString(new() { WriteIndented = true }));

        var exception = Assert.Throws<ScenarioValidationException>(() =>
            ScenarioLoader.LoadScenario(invalidPath, corpus.RootDirectory, corpus.Provenance,
                corpus.ExpectedDocumentProvenance));

        Assert.Contains("at least one explicit deterministic invariant", exception.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Fast_scripted_baselines_pass_and_always_leave_inspectable_artifacts()
    {
        var root = Path.Combine(ArtifactRoot, "fast");
        var outcome = new LegalEvaluationRunner().Run(new EvaluationRunOptions(
            CorpusPath, "fast", null, null, root, null));

        Assert.Equal(0, outcome.ExitCode);
        Assert.True(outcome.ArtifactsPublished);
        Assert.Equal(6, outcome.Results.Count);
        foreach (var result in outcome.Results)
        {
            var score = result.EngineBaseline;

            Assert.Equal("passed", score.Status);
            Assert.Equal(new[]
            {
                "document_validity",
                "redline_reversibility",
                "rendering_regression",
                "target_precision",
                "task_completion",
                "unintended_change",
            }, score.Metrics.Select(value => value.Category)
                .Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal));
            AssertInspectableArtifacts(score);
        }
        var summaryPath = Assert.IsType<string>(outcome.SummaryPath);
        Assert.True(File.Exists(summaryPath));
        Assert.True(File.Exists(Path.Combine(root, "index.html")));
        var summary = File.ReadAllText(summaryPath);
        Assert.DoesNotContain(RepositoryRoot, summary, StringComparison.Ordinal);
        Assert.Contains("\"artifactRoot\": \".\"", summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Whole_run_publication_removes_scenarios_excluded_by_a_filtered_rerun()
    {
        var root = Path.Combine(ArtifactRoot, "full-then-filtered");
        var runner = new LegalEvaluationRunner();
        var full = runner.Run(new EvaluationRunOptions(
            CorpusPath, "full", null, null, root, null));
        Assert.Equal(0, full.ExitCode);
        Assert.Equal(9, full.Results.Count);
        Assert.True(Directory.Exists(Path.Combine(root, "compare-consolidate")));

        var filtered = runner.Run(new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, null));

        Assert.Equal(0, filtered.ExitCode);
        Assert.Single(filtered.Results);
        Assert.False(Directory.Exists(Path.Combine(root, "compare-consolidate")));
        Assert.Equal(new[] { "defined-term-targeting" }, Directory.GetDirectories(root)
            .Select(Path.GetFileName).Order(StringComparer.Ordinal));
        var score = filtered.Results.Single().EngineBaseline;
        Assert.StartsWith(Path.GetFullPath(root) + Path.DirectorySeparatorChar,
            score.ArtifactDirectory!, StringComparison.Ordinal);
        Assert.DoesNotContain(".stage-", score.ArtifactDirectory!, StringComparison.Ordinal);
        Assert.All(score.Artifacts.Where(value => value.Path is not null), artifact =>
        {
            Assert.StartsWith(Path.GetFullPath(root) + Path.DirectorySeparatorChar,
                artifact.Path!, StringComparison.Ordinal);
            Assert.DoesNotContain(".stage-", artifact.Path!, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void Whole_run_publication_removes_model_evidence_when_candidates_are_not_requested()
    {
        var corpus = ScenarioLoader.LoadCorpus(CorpusPath);
        var scenario = corpus.Scenarios.Single(value => value.Id == "defined-term-targeting");
        var root = Path.Combine(ArtifactRoot, "candidate-then-none");
        var candidateDirectory = Path.Combine(ArtifactRoot, "candidate-then-none-input");
        Directory.CreateDirectory(candidateDirectory);
        File.Copy(scenario.ExpectedDocument.Path,
            Path.Combine(candidateDirectory, scenario.Id + ".docx"), overwrite: true);
        var runner = new LegalEvaluationRunner();

        var withCandidate = runner.Run(new EvaluationRunOptions(
            CorpusPath, "full", scenario.Id, candidateDirectory, root, null));
        Assert.Equal(0, withCandidate.ExitCode);
        Assert.NotNull(withCandidate.Results.Single().ModelPlanning);
        Assert.True(Directory.Exists(Path.Combine(root, scenario.Id, "model-planning")));

        var withoutCandidate = runner.Run(new EvaluationRunOptions(
            CorpusPath, "full", scenario.Id, null, root, null));

        Assert.Equal(0, withoutCandidate.ExitCode);
        Assert.Null(withoutCandidate.Results.Single().ModelPlanning);
        Assert.False(Directory.Exists(Path.Combine(root, scenario.Id, "model-planning")));
        Assert.True(Directory.Exists(Path.Combine(root, scenario.Id, "engine-baseline")));
    }

    [Fact]
    public void Artifact_root_validation_refuses_protected_or_unowned_directories()
    {
        var runner = new LegalEvaluationRunner();
        var options = new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null,
            Directory.GetCurrentDirectory(), null);

        var currentDirectoryError = Assert.Throws<ScenarioValidationException>(() =>
            runner.Run(options));
        Assert.Contains("current working directory", currentDirectoryError.Message,
            StringComparison.Ordinal);
        var corpusRoot = Path.GetDirectoryName(CorpusPath)!;
        var corpusError = Assert.Throws<ScenarioValidationException>(() =>
            runner.Run(options with { ArtifactRoot = corpusRoot }));
        Assert.Contains("corpus", corpusError.Message, StringComparison.Ordinal);

        var candidateDirectory = Path.Combine(
            ArtifactRoot, "protected-candidates-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(candidateDirectory);
        var candidatePath = Path.Combine(candidateDirectory, "defined-term-targeting.docx");
        File.Copy(Path.Combine(corpusRoot, "expected", "defined-term-targeting.docx"),
            candidatePath);
        var candidateError = Assert.Throws<ScenarioValidationException>(() => runner.Run(
            options with
            {
                ArtifactRoot = candidateDirectory,
                CandidateDirectory = candidateDirectory,
            }));
        Assert.Contains("candidate directory", candidateError.Message, StringComparison.Ordinal);
        Assert.True(File.Exists(candidatePath));

        var unowned = Path.Combine(
            ArtifactRoot, "unowned-root-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(unowned);
        var sentinel = Path.Combine(unowned, "must-survive.txt");
        File.WriteAllText(sentinel, "not evaluation output");
        var ownershipError = Assert.Throws<ScenarioValidationException>(() =>
            runner.Run(options with { ArtifactRoot = unowned }));
        Assert.Contains("not owned by legal-eval", ownershipError.Message,
            StringComparison.Ordinal);
        Assert.True(File.Exists(sentinel));
        Assert.Empty(Directory.GetFileSystemEntries(Path.GetDirectoryName(unowned)!,
            "." + Path.GetFileName(unowned) + ".stage-*"));
    }

    [Fact]
    public void Markerless_legacy_lookalike_root_is_refused_without_changing_its_files()
    {
        var root = Path.Combine(
            ArtifactRoot, "legacy-lookalike-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var summary = Path.Combine(root, "run-summary.json");
        var index = Path.Combine(root, "index.html");
        var sentinel = Path.Combine(root, "must-survive.txt");
        var summaryContents = """
            {
              "documentKind": "docxodus.legal-evaluation-run-summary",
              "schemaVersion": "1.0",
              "corpus": "corpus.json",
              "subset": "full",
              "artifactRoot": ".",
              "results": []
            }
            """;
        File.WriteAllText(summary, summaryContents);
        File.WriteAllText(index, "unrelated index");
        File.WriteAllText(sentinel, "unrelated sentinel");

        var exception = Assert.Throws<ScenarioValidationException>(() =>
            new LegalEvaluationRunner().Run(new EvaluationRunOptions(
                CorpusPath, "full", "defined-term-targeting", null, root, null)));

        Assert.Contains("not owned by legal-eval", exception.Message, StringComparison.Ordinal);
        Assert.Equal(summaryContents, File.ReadAllText(summary));
        Assert.Equal("unrelated index", File.ReadAllText(index));
        Assert.Equal("unrelated sentinel", File.ReadAllText(sentinel));
        Assert.False(File.Exists(Path.Combine(root, ".docxodus-legal-eval-root")));
        Assert.Empty(Directory.GetFileSystemEntries(Path.GetDirectoryName(root)!,
            "." + Path.GetFileName(root) + ".stage-*"));
    }

    [Fact]
    public void Symlinked_ancestors_cannot_alias_artifacts_or_reports_into_protected_trees()
    {
        var corpusRoot = Path.GetDirectoryName(CorpusPath)!;
        var corpusAlias = Path.Combine(
            ArtifactRoot, "corpus-alias-" + Guid.NewGuid().ToString("N"));
        CreateDirectorySymlinkOrSkip(corpusAlias, corpusRoot);
        var runner = new LegalEvaluationRunner();
        var options = new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null,
            Path.Combine(ArtifactRoot, "safe-root-" + Guid.NewGuid().ToString("N")), null);

        var aliasedArtifactTarget = Path.Combine(
            corpusAlias, "must-not-create-" + Guid.NewGuid().ToString("N"));
        var artifactException = Assert.Throws<ScenarioValidationException>(() =>
            runner.Run(options with { ArtifactRoot = aliasedArtifactTarget }));
        Assert.Contains("corpus", artifactException.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(Path.Combine(
            corpusRoot, Path.GetFileName(aliasedArtifactTarget))));

        var nestedAliasParent = Path.Combine(
            ArtifactRoot, "nested-alias-parent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(nestedAliasParent);
        var nestedInnerAlias = Path.Combine(nestedAliasParent, "inner");
        CreateDirectorySymlinkOrSkip(nestedInnerAlias, corpusRoot);
        var nestedOuterAlias = Path.Combine(
            ArtifactRoot, "nested-outer-alias-" + Guid.NewGuid().ToString("N"));
        CreateDirectorySymlinkOrSkip(nestedOuterAlias,
            Path.Combine(nestedInnerAlias, "scenarios"));
        var nestedTargetName = "must-not-create-" + Guid.NewGuid().ToString("N");
        var nestedArtifactException = Assert.Throws<ScenarioValidationException>(() =>
            runner.Run(options with
            {
                ArtifactRoot = Path.Combine(nestedOuterAlias, nestedTargetName),
            }));
        Assert.Contains("corpus", nestedArtifactException.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(Path.Combine(corpusRoot, "scenarios", nestedTargetName)));

        var corpusAliasArtifactTarget = Path.Combine(
            corpusRoot, "must-not-create-" + Guid.NewGuid().ToString("N"));
        var corpusException = Assert.Throws<ScenarioValidationException>(() => runner.Run(
            options with
            {
                CorpusPath = Path.Combine(corpusAlias, "corpus.json"),
                ArtifactRoot = corpusAliasArtifactTarget,
            }));
        Assert.Contains("corpus", corpusException.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(corpusAliasArtifactTarget));

        var reportTarget = Path.Combine(corpusRoot, "provenance.json");
        var reportHash = Sha256File(reportTarget);
        var reportRoot = Path.Combine(
            ArtifactRoot, "report-alias-root-" + Guid.NewGuid().ToString("N"));
        var reportException = Assert.Throws<ScenarioValidationException>(() => runner.Run(
            options with
            {
                ArtifactRoot = reportRoot,
                ReportPath = Path.Combine(corpusAlias, "provenance.json"),
            }));
        Assert.Contains("corpus", reportException.Message, StringComparison.Ordinal);
        Assert.Equal(reportHash, Sha256File(reportTarget));
        Assert.False(Directory.Exists(reportRoot));

        var candidateDirectory = Path.Combine(
            ArtifactRoot, "candidate-target-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(candidateDirectory);
        var candidateAlias = Path.Combine(
            ArtifactRoot, "candidate-alias-" + Guid.NewGuid().ToString("N"));
        CreateDirectorySymlinkOrSkip(candidateAlias, candidateDirectory);
        var candidateArtifactTarget = Path.Combine(
            candidateDirectory, "must-not-create-" + Guid.NewGuid().ToString("N"));
        var candidateException = Assert.Throws<ScenarioValidationException>(() => runner.Run(
            options with
            {
                ArtifactRoot = candidateArtifactTarget,
                CandidateDirectory = candidateAlias,
            }));
        Assert.Contains("candidate directory", candidateException.Message,
            StringComparison.Ordinal);
        Assert.False(Directory.Exists(candidateArtifactTarget));
    }

    [Fact]
    public void In_root_report_collision_is_rejected_before_staging_but_summary_alias_is_allowed()
    {
        var root = Path.Combine(
            ArtifactRoot, "report-collision-" + Guid.NewGuid().ToString("N"));
        var collision = Path.Combine(
            root, "defined-term-targeting", "engine-baseline", "candidate.docx");
        var options = new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, collision);
        var runner = new LegalEvaluationRunner();

        var exception = Assert.Throws<ScenarioValidationException>(() => runner.Run(options));

        Assert.Contains("canonical run-summary.json", exception.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(root));
        Assert.Empty(Directory.GetFileSystemEntries(Path.GetDirectoryName(root)!,
            "." + Path.GetFileName(root) + ".stage-*"));

        var alias = runner.Run(options with
        {
            ReportPath = Path.Combine(root, "run-summary.json"),
        });
        Assert.Equal(0, alias.ExitCode);
        Assert.True(alias.ArtifactsPublished);
        Assert.Equal(Path.Combine(Path.GetFullPath(root), "run-summary.json"),
            Assert.IsType<string>(alias.SummaryPath));
    }

    [Fact]
    public void External_report_is_staged_and_committed_without_changing_the_sealed_root()
    {
        var root = Path.Combine(
            ArtifactRoot, "external-report-success-" + Guid.NewGuid().ToString("N"));
        var reportPath = Path.Combine(
            ArtifactRoot, "external-report-success-" + Guid.NewGuid().ToString("N") + ".json");

        var outcome = new LegalEvaluationRunner().Run(new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, reportPath));

        Assert.Equal(0, outcome.ExitCode);
        Assert.True(outcome.ArtifactsPublished);
        Assert.Null(outcome.ReportError);
        var summaryPath = Assert.IsType<string>(outcome.SummaryPath);
        Assert.Equal(File.ReadAllBytes(summaryPath), File.ReadAllBytes(reportPath));
        var report = JsonNode.Parse(File.ReadAllText(reportPath))!;
        Assert.Equal("docxodus.legal-evaluation-run-summary",
            report["documentKind"]!.GetValue<string>());
        Assert.True(File.Exists(Path.Combine(root, ".docxodus-legal-eval-root")));
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(reportPath)!,
            Path.GetFileName(reportPath) + ".stage-*"));

        var update = new LegalEvaluationRunner().Run(new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, reportPath));
        Assert.Equal(0, update.ExitCode);
        Assert.True(update.ArtifactsPublished);
        Assert.Equal(File.ReadAllBytes(update.SummaryPath!), File.ReadAllBytes(reportPath));
    }

    [Fact]
    public void External_report_refuses_an_unowned_existing_file_without_changing_it()
    {
        var root = Path.Combine(
            ArtifactRoot, "external-report-refusal-" + Guid.NewGuid().ToString("N"));
        var reportPath = Path.Combine(
            ArtifactRoot, "unowned-report-" + Guid.NewGuid().ToString("N") + ".json");
        const string Sentinel = "unrelated external report sentinel\n";
        File.WriteAllText(reportPath, Sentinel);

        var exception = Assert.Throws<ScenarioValidationException>(() =>
            new LegalEvaluationRunner().Run(new EvaluationRunOptions(
                CorpusPath, "full", "defined-term-targeting", null, root, reportPath)));

        Assert.Contains("not owned by legal-eval", exception.Message, StringComparison.Ordinal);
        Assert.Equal(Sentinel, File.ReadAllText(reportPath));
        Assert.False(Directory.Exists(root));
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(reportPath)!,
            Path.GetFileName(reportPath) + ".stage-*"));
    }

    [Fact]
    public void Failed_prepublication_preserves_the_previous_root_without_dangling_outcome_paths()
    {
        var root = Path.Combine(ArtifactRoot, "publication-failure-preserves-root");
        var options = new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, null);
        var runner = new LegalEvaluationRunner();
        var first = runner.Run(options);
        Assert.Equal(0, first.ExitCode);
        Assert.True(first.ArtifactsPublished);
        var firstSummaryPath = Assert.IsType<string>(first.SummaryPath);
        var originalSummaryHash = Sha256File(firstSummaryPath);
        var marker = Path.Combine(root, "prior-root-marker.txt");
        File.WriteAllText(marker, "prior visible root");
        var reportTarget = Path.Combine(ArtifactRoot, "publication-failure-report-target");
        Directory.CreateDirectory(reportTarget);

        var failed = runner.Run(options with { ReportPath = reportTarget });

        Assert.Equal(2, failed.ExitCode);
        Assert.False(failed.ArtifactsPublished);
        Assert.Null(failed.SummaryPath);
        Assert.Empty(failed.Results);
        Assert.Null(failed.ReportError);
        Assert.Contains("before artifact root publication", failed.FatalError,
            StringComparison.Ordinal);
        Assert.True(File.Exists(marker));
        Assert.Equal(originalSummaryHash, Sha256File(firstSummaryPath));
        Assert.False(Directory.GetFileSystemEntries(Path.GetDirectoryName(root)!,
            "." + Path.GetFileName(root) + ".stage-*").Any());
    }

    [Fact]
    public void External_report_commit_failure_keeps_the_published_root_truthful()
    {
        var root = Path.Combine(ArtifactRoot, "external-report-commit-failure");
        var reportPath = Path.Combine(
            ArtifactRoot, "external-report-" + Guid.NewGuid().ToString("N") + ".json");
        var runner = new LegalEvaluationRunner(
            externalReportCommitter: (_, _) =>
                throw new IOException("induced external report commit failure"));

        var outcome = runner.Run(new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, reportPath));

        Assert.Equal(2, outcome.ExitCode);
        Assert.True(outcome.ArtifactsPublished);
        Assert.Null(outcome.FatalError);
        Assert.Contains("artifact root remains valid", outcome.ReportError,
            StringComparison.Ordinal);
        var summaryPath = Assert.IsType<string>(outcome.SummaryPath);
        Assert.True(File.Exists(summaryPath));
        var score = Assert.Single(outcome.Results).EngineBaseline;
        Assert.All(score.Artifacts.Where(value => value.Path is not null), artifact =>
            Assert.True(File.Exists(artifact.Path), artifact.Path));
        Assert.False(File.Exists(reportPath));
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(reportPath)!,
            Path.GetFileName(reportPath) + ".stage-*"));
        var summary = JsonNode.Parse(File.ReadAllText(summaryPath))!;
        Assert.Null(summary["fatalError"]);
        Assert.Equal(1, summary["engineBaselinePassed"]!.GetValue<int>());
    }

    [Fact]
    public void Engine_baseline_and_model_planning_are_scored_separately()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios
            .Single(value => value.Id == "defined-term-targeting");
        var baseline = new ScriptedBaselineExecutor().Execute(scenario);
        var scorer = new EvaluationScorer();
        var root = Path.Combine(ArtifactRoot, "score-separation");

        var engine = scorer.Score(scenario, baseline, baseline.Output,
            ScoreKind.EngineBaseline, root);
        var model = scorer.Score(scenario, baseline, baseline.Input,
            ScoreKind.ModelPlanning, root);

        Assert.Equal("passed", engine.Status);
        Assert.Equal("failed", model.Status);
        Assert.Contains(model.Metrics, value =>
            value.Id == "target-precision.reference-equivalence" && value.Status == "failed");
        Assert.NotEqual(engine.ArtifactDirectory, model.ArtifactDirectory);
        AssertInspectableArtifacts(engine);
        AssertInspectableArtifacts(model);
    }

    [Fact]
    public void Required_unavailable_artifact_fails_task_completion()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios
            .Single(value => value.Id == "defined-term-targeting");
        scenario = scenario with
        {
            ExpectedOutputs = new[]
            {
                new ExpectedArtifact("redline-proof-v1", "application/json",
                    "verification", Required: true),
            },
        };
        var baseline = new ScriptedBaselineExecutor().Execute(scenario);

        var score = new EvaluationScorer().Score(scenario, baseline, baseline.Output,
            ScoreKind.EngineBaseline, Path.Combine(ArtifactRoot, "required-artifact"));

        Assert.Equal("failed", score.Status);
        var metric = Assert.Single(score.Metrics,
            value => value.Id == "task-completion.required-artifacts");
        Assert.Equal("failed", metric.Status);
        Assert.Contains("redline-proof-v1", metric.Detail, StringComparison.Ordinal);
        var proof = Assert.Single(score.Artifacts, value => value.Id == "redline-proof-v1");
        Assert.Equal("unavailable", proof.Status);
        var metrics = JsonNode.Parse(File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "metrics-json").Path!))!;
        Assert.Equal("failed", metrics["status"]!.GetValue<string>());
        Assert.Equal("failed", metrics["metrics"]!.AsArray().Single(value =>
            value!["id"]!.GetValue<string>() == "task-completion.required-artifacts")!["status"]!
            .GetValue<string>());
        var summary = File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "scenario-summary").Path!);
        Assert.Contains("Status: `failed`", summary, StringComparison.Ordinal);
        var receipt = JsonNode.Parse(File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "evaluation-receipt").Path!))!;
        Assert.Equal("failed", receipt["status"]!.GetValue<string>());
        Assert.Equal("failed", receipt["metricStatus"]!.AsArray().Single(value =>
            value!["id"]!.GetValue<string>() == "task-completion.required-artifacts")!["status"]!
            .GetValue<string>());
    }

    [Fact]
    public void Malformed_model_candidate_fails_without_losing_evidence()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios
            .Single(value => value.Id == "defined-term-targeting");
        var baseline = new ScriptedBaselineExecutor().Execute(scenario);
        var malformed = Encoding.UTF8.GetBytes("not an OPC package");

        var score = new EvaluationScorer().Score(scenario, baseline, malformed,
            ScoreKind.ModelPlanning, Path.Combine(ArtifactRoot, "malformed-candidate"));

        Assert.Equal("failed", score.Status);
        Assert.Contains(score.Metrics, value => value.Status == "failed"
            && value.Detail.Contains("blocked by package safety validation", StringComparison.Ordinal));
        Assert.Contains(score.Metrics, value =>
            value.Id == "task-completion.required-artifacts" && value.Status == "failed");
        Assert.NotNull(score.ArtifactDirectory);
        var candidate = Assert.Single(score.Artifacts, value => value.Id == "candidate-docx");
        Assert.Equal("available", candidate.Status);
        Assert.Equal(malformed, File.ReadAllBytes(candidate.Path!));
        var preview = Assert.Single(score.Artifacts, value => value.Id == "after-html");
        Assert.Equal("failed", preview.Status);
        Assert.Contains("Unavailable preview", File.ReadAllText(preview.Path!),
            StringComparison.Ordinal);
        foreach (var id in new[] { "metrics-json", "scenario-summary", "artifact-status" })
        {
            var artifact = Assert.Single(score.Artifacts, value => value.Id == id);
            Assert.Equal("available", artifact.Status);
            Assert.True(File.Exists(artifact.Path), artifact.Path);
        }
    }

    [Fact]
    public void Generated_fixture_and_baseline_are_byte_deterministic()
    {
        var scenarios = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios;
        var executor = new ScriptedBaselineExecutor();

        foreach (var scenario in scenarios)
        {
            var first = executor.Execute(scenario);
            var second = executor.Execute(scenario);
            var inspectionDirectory = Path.Combine(ArtifactRoot, "determinism", scenario.Id);
            Directory.CreateDirectory(inspectionDirectory);
            File.WriteAllBytes(Path.Combine(inspectionDirectory, "first-input.docx"), first.Input);
            File.WriteAllBytes(Path.Combine(inspectionDirectory, "second-input.docx"), second.Input);
            File.WriteAllBytes(Path.Combine(inspectionDirectory, "first-output.docx"), first.Output);
            File.WriteAllBytes(Path.Combine(inspectionDirectory, "second-output.docx"), second.Output);

            Assert.Equal(first.Input, second.Input);
            Assert.Equal(first.Output, second.Output);
            AssertDeterministicContainer(first.Input);
            AssertDeterministicContainer(first.Output);
        }
    }

    [Fact]
    public void Readable_fixture_recipe_reproduces_the_pinned_input_fixture()
    {
        var corpus = ScenarioLoader.LoadCorpus(CorpusPath);
        var provenance = Assert.Single(corpus.Provenance.Values);

        var generated = LegalFixtureFactory.Build(provenance.RecipePath!);
        var pinned = File.ReadAllBytes(provenance.SourcePath);

        Assert.Equal(provenance.SourceSha256, Sha256File(provenance.SourcePath));
        Assert.Equal(pinned, generated);
    }

    [Fact]
    public void Shared_runner_preserves_checkpointed_failure_bundle_and_replaces_stale_evidence()
    {
        var root = Path.Combine(ArtifactRoot, "checkpointed-failure");
        var options = new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, null);
        var passing = new LegalEvaluationRunner().Run(options);
        Assert.Equal(0, passing.ExitCode);
        var scoreDirectory = passing.Results.Single().EngineBaseline.ArtifactDirectory!;
        File.WriteAllText(Path.Combine(scoreDirectory, "stale-from-prior-run.txt"), "stale");
        File.WriteAllText(Path.Combine(scoreDirectory, "candidate.pdf"), "stale pdf");
        File.WriteAllText(Path.Combine(scoreDirectory, "candidate-page-1.png"), "stale png");

        var renderer = new StubArtifactRenderer(available: true);
        var failing = new LegalEvaluationRunner(scenario => scenario with
        {
            BaselineOperations = scenario.BaselineOperations.Concat(new[]
            {
                new JsonObject
                {
                    ["op"] = "failForTest",
                    ["message"] = "induced checkpoint failure",
                },
            }).ToList(),
        }, renderer).Run(options with { RenderMode = ArtifactRenderMode.TrustedDocuments });

        Assert.Equal(1, failing.ExitCode);
        var score = failing.Results.Single().EngineBaseline;
        Assert.Equal("failed", score.Status);
        Assert.False(File.Exists(Path.Combine(scoreDirectory, "stale-from-prior-run.txt")));
        Assert.False(File.Exists(Path.Combine(scoreDirectory, "candidate-page-1.png")));
        foreach (var id in new[]
        {
            "input-docx", "candidate-docx", "expected-docx", "semantic-diff",
            "before-html", "after-html", "target-html", "metrics-json",
            "evaluation-receipt", "artifact-index-html", "artifact-status",
        })
        {
            var artifact = Assert.Single(score.Artifacts, value => value.Id == id);
            Assert.NotNull(artifact.Path);
            Assert.True(File.Exists(artifact.Path), artifact.Path);
        }
        var candidatePdf = Assert.Single(score.Artifacts, value => value.Id == "candidate-pdf");
        Assert.Equal("available", candidatePdf.Status);
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(File.ReadAllBytes(candidatePdf.Path!)),
            StringComparison.Ordinal);
        Assert.Equal(3, renderer.RenderCalls);
        Assert.Equal(1, renderer.CompareCalls);
        var afterHtml = File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "after-html").Path!);
        Assert.Contains("Availability Credit", afterHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Unavailable preview", afterHtml, StringComparison.Ordinal);
        var receipt = File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "evaluation-receipt").Path!);
        Assert.Contains("induced checkpoint failure", receipt, StringComparison.Ordinal);
        Assert.Contains("\"status\": \"failed\"", receipt, StringComparison.Ordinal);
        Assert.True(File.Exists(Assert.IsType<string>(failing.SummaryPath)));
    }

    [Fact]
    public void Pinned_scripted_golden_detects_divergence_from_the_declared_edit()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios
            .Single(value => value.Id == "defined-term-targeting");
        var wrongExpected = scenario with
        {
            ExpectedDocument = new ExpectedDocumentReference(
                scenario.Fixture.Path, scenario.ExpectedDocument.ProvenanceId,
                scenario.Fixture.SourceSha256),
        };
        var baseline = new ScriptedBaselineExecutor().Execute(wrongExpected);

        var score = new EvaluationScorer().Score(wrongExpected, baseline, baseline.Output,
            ScoreKind.EngineBaseline, Path.Combine(ArtifactRoot, "scripted-golden-divergence"));

        Assert.Equal("failed", score.Status);
        Assert.Contains(score.Metrics, value =>
            value.Id == "target-precision.reference-equivalence" && value.Status == "failed");
    }

    [Fact]
    public void Multi_operation_target_diff_covers_the_complete_input_to_target_change()
    {
        var root = Path.Combine(ArtifactRoot, "complete-target-diff");
        var outcome = new LegalEvaluationRunner().Run(new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", null, root, null));
        var score = outcome.Results.Single().EngineBaseline;
        var candidatePath = Assert.Single(score.Artifacts, value => value.Id == "semantic-diff").Path!;
        var targetPath = Assert.Single(score.Artifacts, value => value.Id == "target-semantic-diff").Path!;
        var candidate = JsonNode.Parse(File.ReadAllText(candidatePath));
        var target = JsonNode.Parse(File.ReadAllText(targetPath));

        Assert.True(JsonNode.DeepEquals(candidate, target));
        var operations = target!["operations"]!.AsArray();
        Assert.Equal(2, operations.Count(value =>
            value?["kind"]?.GetValue<string>() == "ModifyBlock"));
    }

    [Fact]
    public void Loader_rejects_unsafe_ids_unknown_properties_and_invalid_probe_operands()
    {
        var corpus = ScenarioLoader.LoadCorpus(CorpusPath);
        var source = corpus.Scenarios.Single(value => value.Id == "defined-term-targeting");
        var directory = Path.Combine(ArtifactRoot, "adversarial-loader");
        Directory.CreateDirectory(directory);

        var unsafeId = JsonNode.Parse(File.ReadAllText(source.FilePath))!.AsObject();
        unsafeId["id"] = "../../escape";
        AssertInvalidScenario(unsafeId, directory, "unsafe-id.json", corpus, "safe scenario slug");

        var unknown = JsonNode.Parse(File.ReadAllText(source.FilePath))!.AsObject();
        unknown["surprise"] = true;
        AssertInvalidScenario(unknown, directory, "unknown-property.json", corpus, "unknown properties");

        var invalidOperand = JsonNode.Parse(File.ReadAllText(source.FilePath))!.AsObject();
        var invariant = invalidOperand["invariants"]!.AsArray()[0]!.AsObject();
        invariant["operator"] = "contains";
        invariant["expected"] = null;
        AssertInvalidScenario(invalidOperand, directory, "invalid-operand.json", corpus,
            "numeric probe requires");

        var escapingFixture = JsonNode.Parse(File.ReadAllText(source.FilePath))!.AsObject();
        escapingFixture["fixture"]!["path"] = "../outside.docx";
        AssertInvalidScenario(escapingFixture, directory, "escaping-fixture.json", corpus,
            "path escapes corpus root");

        var consolidationSource = corpus.Scenarios
            .Single(value => value.Id == "compare-consolidate");
        var invalidReviewer = JsonNode.Parse(File.ReadAllText(consolidationSource.FilePath))!.AsObject();
        invalidReviewer["baseline"]!["operations"]![0]!["reviewers"]![0] = new JsonObject
        {
            ["op"] = "fillContentControl",
            ["tag"] = "ClientName",
            ["text"] = "Unsafe nested operation",
        };
        AssertInvalidScenario(invalidReviewer, directory, "invalid-reviewer.json", corpus,
            "reviewer operation 'fillContentControl' is unsupported");
    }

    [Fact]
    public void Interim_package_guard_enforces_read_and_zip_path_budgets()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios[0];
        var bytes = File.ReadAllBytes(scenario.Fixture.Path);
        var tiny = new InterimEvaluationPackageValidator(
            new EvaluationPackageLimits(MaximumPackageBytes: 4));
        Assert.Throws<ScenarioValidationException>(() => tiny.Validate(bytes, "candidate"));

        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(archive.CreateEntry("../escape.xml").Open());
            writer.Write("<root/>");
        }
        Assert.Throws<ScenarioValidationException>(() =>
            new InterimEvaluationPackageValidator().Validate(stream.ToArray(), "candidate"));

        var xmlBytes = Zip(("document.xml", "<root>" + new string('x', 128) + "</root>"));
        var tinyXml = new InterimEvaluationPackageValidator(new EvaluationPackageLimits(
            MaximumXmlPartBytes: 32));
        Assert.Throws<ScenarioValidationException>(() => tinyXml.Validate(xmlBytes, "candidate"));

        var expandedBytes = Zip(("payload.bin", new string('x', 256)));
        var tinyExpanded = new InterimEvaluationPackageValidator(new EvaluationPackageLimits(
            MaximumExpandedBytes: 64, MaximumCompressionRatio: 10_000));
        Assert.Throws<ScenarioValidationException>(() =>
            tinyExpanded.Validate(expandedBytes, "candidate"));

        var compressedBytes = Zip(("payload.bin", new string('x', 16_384)));
        var tinyRatio = new InterimEvaluationPackageValidator(new EvaluationPackageLimits(
            MaximumCompressionRatio: 2));
        Assert.Throws<ScenarioValidationException>(() =>
            tinyRatio.Validate(compressedBytes, "candidate"));

        var dtdBytes = Zip(("document.xml", "<!DOCTYPE root [<!ENTITY x 'boom'>]><root>&x;</root>"));
        Assert.Throws<ScenarioValidationException>(() =>
            new InterimEvaluationPackageValidator().Validate(dtdBytes, "candidate"));
    }

    [Fact]
    public void Model_preview_suppresses_active_external_links_and_external_rendering()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios
            .Single(value => value.Id == "defined-term-targeting");
        var baseline = new ScriptedBaselineExecutor().Execute(scenario);
        using var session = new DocxSession(baseline.Output);
        var anchor = session.FindAllByText("Availability Credit").First().Anchor.Id;
        Assert.True(session.AddHyperlink(anchor, new CharSpan(0, 5),
            HyperlinkTarget.External("https://evil.example/track")).Success);
        var candidate = session.Save(false);

        var renderer = new StubArtifactRenderer(available: true);
        var score = new EvaluationScorer(artifactRenderer: renderer).Score(scenario, baseline, candidate,
            ScoreKind.ModelPlanning, Path.Combine(ArtifactRoot, "sanitized-preview"),
            ArtifactRenderMode.TrustedDocuments);

        var html = File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "after-html").Path!);
        Assert.Contains("Content-Security-Policy", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://evil.example", html, StringComparison.Ordinal);
        Assert.Equal("unavailable", Assert.Single(score.Artifacts,
            value => value.Id == "candidate-pdf").Status);
        Assert.Contains("untrusted model candidates", Assert.Single(score.Artifacts,
            value => value.Id == "candidate-pdf").UnavailableReason!, StringComparison.Ordinal);
        Assert.Equal(0, renderer.RenderCalls);
    }

    [Fact]
    public void Renderer_contract_records_deterministic_available_and_unavailable_evidence()
    {
        var scenario = ScenarioLoader.LoadCorpus(CorpusPath).Scenarios
            .Single(value => value.Id == "defined-term-targeting");
        var baseline = new ScriptedBaselineExecutor().Execute(scenario);

        var unavailableRenderer = new StubArtifactRenderer(available: false);
        var unavailable = new EvaluationScorer(artifactRenderer: unavailableRenderer).Score(
            scenario, baseline, baseline.Output, ScoreKind.EngineBaseline,
            Path.Combine(ArtifactRoot, "renderer-unavailable"),
            ArtifactRenderMode.TrustedDocuments);
        Assert.Equal(4, unavailableRenderer.RenderCalls);
        Assert.Equal("unavailable", Assert.Single(unavailable.Artifacts,
            value => value.Id == "candidate-pdf").Status);
        Assert.Contains("stub renderer unavailable", Assert.Single(unavailable.Artifacts,
            value => value.Id == "candidate-pdf").UnavailableReason!, StringComparison.Ordinal);

        var availableRenderer = new StubArtifactRenderer(available: true);
        var available = new EvaluationScorer(artifactRenderer: availableRenderer).Score(
            scenario, baseline, baseline.Output, ScoreKind.EngineBaseline,
            Path.Combine(ArtifactRoot, "renderer-available"),
            ArtifactRenderMode.TrustedDocuments);
        Assert.Equal(4, availableRenderer.RenderCalls);
        Assert.Equal(1, availableRenderer.CompareCalls);
        foreach (var id in new[]
        {
            "input-pdf", "candidate-pdf", "target-pdf", "redline-pdf",
            "before-visual", "candidate-visual", "target-visual", "redline-visual",
            "candidate-target-visual-diff",
        })
        {
            var artifact = Assert.Single(available.Artifacts, value => value.Id == id);
            Assert.Equal("available", artifact.Status);
            Assert.True(File.Exists(artifact.Path), artifact.Path);
        }
        Assert.False(File.Exists(Path.Combine(available.ArtifactDirectory!, "before.docx")));
        Assert.False(File.Exists(Path.Combine(available.ArtifactDirectory!, "target.docx")));
    }

    [Fact]
    public void Requested_model_directory_with_missing_candidate_is_incomplete_and_nonzero()
    {
        var candidateDirectory = Path.Combine(ArtifactRoot, "missing-candidates", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(candidateDirectory);
        var outcome = new LegalEvaluationRunner().Run(new EvaluationRunOptions(
            CorpusPath, "full", "defined-term-targeting", candidateDirectory,
            Path.Combine(ArtifactRoot, "missing-candidate-run"), null));

        Assert.Equal(1, outcome.ExitCode);
        var planning = Assert.IsType<EvaluationScore>(outcome.Results.Single().ModelPlanning);
        Assert.Equal("incomplete", planning.Status);
        Assert.Equal("available", Assert.Single(planning.Artifacts,
            value => value.Id == "evaluation-receipt").Status);
        Assert.DoesNotContain(RepositoryRoot,
            File.ReadAllText(Assert.IsType<string>(outcome.SummaryPath)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Runner_writes_a_portable_summary_even_for_a_fatal_corpus_error()
    {
        var root = Path.Combine(ArtifactRoot, "fatal-summary");
        var missingCorpus = Path.Combine(RepositoryRoot, "eval", "legal", "missing-corpus.json");

        var outcome = new LegalEvaluationRunner().Run(new EvaluationRunOptions(
            missingCorpus, "fast", null, null, root, null));

        Assert.Equal(2, outcome.ExitCode);
        Assert.True(outcome.ArtifactsPublished);
        Assert.NotNull(outcome.FatalError);
        var summaryPath = Assert.IsType<string>(outcome.SummaryPath);
        Assert.True(File.Exists(summaryPath));
        Assert.True(File.Exists(Path.Combine(root, "index.html")));
        var summary = File.ReadAllText(summaryPath);
        Assert.DoesNotContain(RepositoryRoot, summary, StringComparison.Ordinal);
        Assert.Contains("missing-corpus.json", summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Legal_eval_workflow_keeps_every_smoke_evidence_step_runnable_after_failure()
    {
        var workflow = File.ReadAllLines(Path.Combine(
            RepositoryRoot, ".github", "workflows", "legal-eval.yml"));
        foreach (var step in new[]
        {
            "Generate epic 435 smoke workflow",
            "Stage epic 435 smoke fixture",
            "Run epic 435 editing smoke",
            "Reopen and validate epic 435 output",
            "Preserve epic 435 edited document",
            "Upload fast evaluation evidence",
            "Upload full evaluation evidence",
        })
        {
            var start = Array.FindIndex(workflow,
                value => value.Trim() == $"- name: {step}");
            Assert.True(start >= 0, $"workflow step not found: {step}");
            var end = Array.FindIndex(workflow, start + 1,
                value => value.TrimStart().StartsWith("- name: ", StringComparison.Ordinal));
            if (end < 0) end = workflow.Length;
            Assert.Contains(workflow[start..end],
                value => value.Trim() == "if: always()");
        }
    }

    private static void AssertInvalidScenario(
        JsonObject node,
        string directory,
        string fileName,
        LegalCorpus corpus,
        string expectedMessage)
    {
        var path = Path.Combine(directory, fileName);
        File.WriteAllText(path, node.ToJsonString(new() { WriteIndented = true }));
        var exception = Assert.Throws<ScenarioValidationException>(() =>
            ScenarioLoader.LoadScenario(path, corpus.RootDirectory, corpus.Provenance,
                corpus.ExpectedDocumentProvenance));
        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
    }

    private static void AssertInspectableArtifacts(EvaluationScore score)
    {
        Assert.NotNull(score.ArtifactDirectory);
        Assert.True(Directory.Exists(score.ArtifactDirectory));
        foreach (var id in new[]
        {
            "input-docx",
            "candidate-docx",
            "expected-docx",
            "semantic-diff",
            "metrics-json",
            "scenario-summary",
            "before-html",
            "after-html",
            "target-html",
            "redline-docx",
            "target-semantic-diff",
            "evaluation-receipt",
            "artifact-index-markdown",
            "artifact-index-html",
            "artifact-status",
        })
        {
            var artifact = Assert.Single(score.Artifacts, value => value.Id == id);
            Assert.Equal("available", artifact.Status);
            Assert.NotNull(artifact.Path);
            Assert.True(File.Exists(artifact.Path), artifact.Path);
            Assert.False(string.IsNullOrWhiteSpace(artifact.Sha256));
            Assert.False(string.IsNullOrWhiteSpace(artifact.MediaType));
            Assert.True(artifact.SizeBytes > 0);
        }

        Assert.Single(score.Artifacts, value => value.Id == "candidate-pdf");
        Assert.Single(score.Artifacts, value => value.Id == "candidate-visual");

        var receiptRecord = Assert.Single(score.Artifacts,
            value => value.Id == "evaluation-receipt");
        var statusRecord = Assert.Single(score.Artifacts,
            value => value.Id == "artifact-status");
        var markdownIndex = File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "artifact-index-markdown").Path!);
        var htmlIndex = File.ReadAllText(Assert.Single(score.Artifacts,
            value => value.Id == "artifact-index-html").Path!);
        foreach (var fileName in new[] { "evaluation-receipt.json", "artifact-status.json" })
        {
            Assert.Contains(fileName, markdownIndex, StringComparison.Ordinal);
            Assert.Contains(fileName, htmlIndex, StringComparison.Ordinal);
        }
        var receipt = JsonNode.Parse(File.ReadAllText(receiptRecord.Path!))!;
        var receiptIds = receipt["artifacts"]!.AsArray()
            .Select(value => value!["id"]!.GetValue<string>()).ToHashSet(StringComparer.Ordinal);
        Assert.DoesNotContain("evaluation-receipt", receiptIds);
        Assert.DoesNotContain("artifact-status", receiptIds);
        Assert.DoesNotContain("artifact-index-markdown", receiptIds);
        Assert.DoesNotContain("artifact-index-html", receiptIds);
        var fingerprintComponents = new List<string>
        {
            receipt["scenarioId"]!.GetValue<string>(),
            score.Kind.ToString(),
            receipt["status"]!.GetValue<string>(),
        };
        fingerprintComponents.AddRange(receipt["operations"]!.AsArray()
            .Select(value => value!.GetValue<string>()));
        foreach (var artifact in receipt["artifacts"]!.AsArray())
        {
            var value = artifact!;
            var path = value["path"]?.GetValue<string>();
            var size = value["sizeBytes"]?.GetValue<long>();
            var sha256 = value["sha256"]?.GetValue<string>();
            if (path is not null)
            {
                var artifactPath = Path.GetFullPath(Path.Combine(
                    score.ArtifactDirectory!, path.Replace('/', Path.DirectorySeparatorChar)));
                Assert.True(File.Exists(artifactPath), artifactPath);
                Assert.Equal(size, new FileInfo(artifactPath).Length);
                Assert.Equal(sha256, Sha256File(artifactPath));
            }
            fingerprintComponents.Add($"{value["id"]!.GetValue<string>()}:"
                + $"{value["status"]!.GetValue<string>()}:{sha256}:{size}");
        }
        var fingerprintInput = string.Join("\n", fingerprintComponents);
        var expectedRunId = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(fingerprintInput)))
            .ToLowerInvariant();
        Assert.Equal(expectedRunId, receipt["runId"]!.GetValue<string>());
        var status = JsonNode.Parse(File.ReadAllText(statusRecord.Path!))!;
        Assert.Contains(status["artifacts"]!.AsArray(), value =>
            value!["id"]!.GetValue<string>() == "evaluation-receipt");
    }

    private static void AssertDeterministicContainer(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var names = archive.Entries.Select(value => value.FullName).ToList();
        Assert.Equal(names.Order(StringComparer.Ordinal), names);
        Assert.All(archive.Entries, entry =>
            Assert.Equal(new DateTime(2000, 1, 1), entry.LastWriteTime.DateTime));
    }

    private static string Sha256File(string path) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path)))
            .ToLowerInvariant();

    private static void CreateDirectorySymlinkOrSkip(string linkPath, string targetPath)
    {
        try
        {
            Directory.CreateSymbolicLink(linkPath, targetPath);
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException or NotSupportedException)
        {
            throw Xunit.Sdk.SkipException.ForSkip(
                $"symbolic links are unavailable: {exception.Message}");
        }
    }

    private static byte[] Zip(params (string Name, string Content)[] entries)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, content) in entries)
            {
                using var writer = new StreamWriter(archive.CreateEntry(name).Open(),
                    new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                writer.Write(content);
            }
        }
        return stream.ToArray();
    }

    private sealed class StubArtifactRenderer(bool available) : IEvaluationArtifactRenderer
    {
        public int RenderCalls { get; private set; }
        public int CompareCalls { get; private set; }

        public RenderedDocumentEvidence RenderDocument(
            string artifactDirectory,
            string sourceDocxPath,
            string outputPrefix)
        {
            RenderCalls++;
            Assert.True(File.Exists(sourceDocxPath), sourceDocxPath);
            if (!available)
                return new RenderedDocumentEvidence(
                    null, Array.Empty<string>(), "stub renderer unavailable");
            var pdf = Path.Combine(artifactDirectory, outputPrefix + ".pdf");
            var page = Path.Combine(artifactDirectory, outputPrefix + "-page-001.png");
            File.WriteAllBytes(pdf, Encoding.ASCII.GetBytes("%PDF-1.4\n% deterministic stub\n"));
            File.WriteAllBytes(page, new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
            return new RenderedDocumentEvidence(pdf, new[] { page }, null);
        }

        public RenderedVisualDiffEvidence ComparePages(
            string artifactDirectory,
            IReadOnlyList<string> targetPages,
            IReadOnlyList<string> candidatePages)
        {
            CompareCalls++;
            if (!available)
                return new RenderedVisualDiffEvidence(
                    Array.Empty<string>(), "stub renderer unavailable");
            Assert.Single(targetPages);
            Assert.Single(candidatePages);
            var path = Path.Combine(artifactDirectory, "candidate-target-diff-page-001.png");
            File.WriteAllBytes(path, new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
            return new RenderedVisualDiffEvidence(new[] { path }, null);
        }
    }
}
