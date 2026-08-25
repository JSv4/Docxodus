// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using System.Text.Json.Nodes;
using Docxodus.Internal;
using Docxodus.McpServer;
using Docxodus.Verification;

namespace Docxodus.Tests.Eval;

/// <summary>Everything one scenario run produced, ready for its invariants to be checked.</summary>
internal sealed record EvalOutcome
{
    required public string Id { get; init; }
    required public byte[] OpeningBytes { get; init; }
    required public byte[] DeliverableBytes { get; init; }
    /// <summary>Distinct anchors named by the #457 change set from opening to delivered state.</summary>
    required public IReadOnlyList<string> ChangedAnchors { get; init; }
    required public IReadOnlyList<string> ChangedParts { get; init; }
    required public int PartsAdded { get; init; }
    required public int PartsRemoved { get; init; }
    required public string ValidityDecision { get; init; }
    /// <summary>The full #463 verification report, retained for failure artifacts.</summary>
    required public string VerificationJson { get; init; }
    /// <summary>
    /// How many blocks contain each needle the scenario's invariants name, answered by
    /// docxodus_search over every story. Text assertions ask the <em>document</em> this way
    /// rather than substring-matching a markdown projection: the projection renders tables
    /// structurally, so cell text is not present in it as literal prose.
    /// </summary>
    required public IReadOnlyDictionary<string, int> TextMatchCounts { get; init; }
    /// <summary>The markdown-derived text projection. Retained for failure artifacts only.</summary>
    required public string Text { get; init; }
    required public string Html { get; init; }
    required public string SemanticChangesJson { get; init; }
    /// <summary>docxodus_track_changes list on the delivered session — live markup, agent-surface view.</summary>
    required public string RevisionsJson { get; init; }
    /// <summary>docxodus_comment list on the delivered session.</summary>
    required public string CommentsJson { get; init; }
    public RedlineReversibilityProof? Reversibility { get; init; }
    public int GeneratedRevisionCount { get; init; }
}

/// <summary>
/// The deterministic scripted caller for the #466 workflow evaluation suite.
///
/// <para>It drives the <em>agent surface</em> — the MCP tool dispatcher — rather than the typed
/// .NET API, so the engine baseline is measured exactly where an agent meets it. That is what
/// separates engine/tool correctness from model planning quality: a scenario's steps are executed
/// verbatim, with no planning, so a failure here is the engine's, never the model's.</para>
///
/// <para>Fixtures are built through the same step format, so the corpus carries no third-party
/// document bytes and every input is reproducible from source.</para>
/// </summary>
internal static class EvalHarness
{
    /// <summary>The <c>eval/</c> corpus directory, located by walking up from the test binary.</summary>
    public static string CorpusRoot { get; } = LocateCorpusRoot();

    /// <summary>The fast deterministic subset: runs unfiltered on every push.</summary>
    public static IEnumerable<string> ScenarioFiles() =>
        Directory.EnumerateFiles(Path.Combine(CorpusRoot, "scenarios"), "*.json")
            .OrderBy(path => path, StringComparer.Ordinal);

    /// <summary>
    /// The opt-in corpus tier under <c>scenarios/corpus/</c>. Executed only when
    /// <c>DOCXODUS_RUN_EVAL_CORPUS=1</c> (the scheduled eval-corpus workflow sets it);
    /// its scenarios' declarations are validated on every push regardless.
    /// </summary>
    public static IEnumerable<string> CorpusScenarioFiles()
    {
        var directory = Path.Combine(CorpusRoot, "scenarios", "corpus");
        return Directory.Exists(directory)
            ? Directory.EnumerateFiles(directory, "*.json").OrderBy(path => path, StringComparer.Ordinal)
            : Enumerable.Empty<string>();
    }

    public static bool RunCorpusTier { get; } = string.Equals(
        Environment.GetEnvironmentVariable("DOCXODUS_RUN_EVAL_CORPUS"), "1", StringComparison.Ordinal);

    public static JsonElement LoadJson(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement.Clone();
    }

    public static JsonElement LoadScenario(string id)
    {
        var fast = Path.Combine(CorpusRoot, "scenarios", $"{id}.json");
        return LoadJson(File.Exists(fast)
            ? fast
            : Path.Combine(CorpusRoot, "scenarios", "corpus", $"{id}.json"));
    }

    /// <summary>Build a fixture's bytes by running its step script over a blank document.</summary>
    public static byte[] BuildFixture(string name)
    {
        var script = LoadJson(Path.Combine(CorpusRoot, "fixtures", $"{name}.json"));
        using var workspace = new EvalWorkspace();
        var sessionId = workspace.Open(DocxSessionOps.CreateBlankDocx(), trackedChanges: "accept");
        RunSteps(workspace, sessionId, script.GetProperty("steps"));
        return workspace.Save(sessionId);
    }

    /// <summary>The text projection of a standalone package, for fixture-reproducibility checks.</summary>
    public static string TextProjection(byte[] packageBytes)
    {
        using var workspace = new EvalWorkspace();
        var sessionId = workspace.Open(packageBytes, trackedChanges: "accept");
        return ReadStringProperty(workspace.Content(sessionId, "text"), "text");
    }

    /// <summary>Execute one scenario and collect every metric its invariants can assert over.</summary>
    public static EvalOutcome Run(JsonElement scenario)
    {
        var id = scenario.GetProperty("id").GetString()!;
        var openingBytes = BuildFixture(scenario.GetProperty("fixture").GetString()!);
        var trackedChanges = scenario.TryGetProperty("trackedChanges", out var tracked)
            ? tracked.GetString()! : "accept";

        using var workspace = new EvalWorkspace();
        var sessionId = workspace.Open(openingBytes, trackedChanges);
        RunSteps(workspace, sessionId, scenario.GetProperty("steps"));

        var deliverable = workspace.Save(sessionId);
        var semanticChangesJson = workspace.Content(sessionId, "semantic_changes");
        var (changedAnchors, changedParts) = ReadChangedTargets(semanticChangesJson);
        var (added, removed) = ComparePartInventories(openingBytes, deliverable);

        var verificationJson = workspace.Content(sessionId, "verification");
        using var verification = JsonDocument.Parse(verificationJson);
        var matchCounts = CountAssertedText(workspace, sessionId, scenario.GetProperty("invariants"));
        var outcome = new EvalOutcome
        {
            Id = id,
            OpeningBytes = openingBytes,
            DeliverableBytes = deliverable,
            ChangedAnchors = changedAnchors,
            ChangedParts = changedParts,
            PartsAdded = added,
            PartsRemoved = removed,
            ValidityDecision =
                verification.RootElement.GetProperty("decision").GetString() ?? "notEvaluated",
            VerificationJson = verificationJson,
            TextMatchCounts = matchCounts,
            Text = ReadStringProperty(workspace.Content(sessionId, "text"), "text"),
            Html = ReadStringProperty(workspace.Content(sessionId, "html"), "html"),
            SemanticChangesJson = semanticChangesJson,
            RevisionsJson = workspace.Call("docxodus_track_changes", new JsonObject
            {
                ["sessionId"] = sessionId,
                ["action"] = "list",
            }),
            CommentsJson = workspace.Call("docxodus_comment", new JsonObject
            {
                ["sessionId"] = sessionId,
                ["action"] = "list",
            }),
        };

        if (ReversibilityMode(scenario) != "acceptAll")
            return outcome;

        // The intended final is derived, not asserted: accept every revision in the deliverable
        // and prove the deliverable reproduces that on accept and the opening package on reject.
        // RevisionProcessor rather than a second session: a session would stamp its own settings
        // and anchor bookkeeping onto the derived package, and the proof would then report that
        // bookkeeping as a divergence belonging to the redline.
        var intendedFinal = RevisionProcessor
            .AcceptRevisions(new WmlDocument("redline.docx", deliverable))
            .DocumentByteArray;
        var proof = RedlineReversibilityVerifier
            .Prove(openingBytes, intendedFinal, deliverable).Proof;
        return outcome with
        {
            Reversibility = proof,
            GeneratedRevisionCount = proof.RevisionClassifications
                .Count(item => item.Disposition == RedlineRevisionDisposition.Generated),
        };
    }

    public static string ReversibilityMode(JsonElement scenario) =>
        scenario.TryGetProperty("invariants", out var invariants)
        && invariants.TryGetProperty("reversibility", out var reversibility)
        && reversibility.TryGetProperty("mode", out var mode)
            ? mode.GetString() ?? "none"
            : "none";

    /// <summary>
    /// Resolve every needle the invariants name to a match count, over every story, while the
    /// session is still open. Counting up front keeps the assertions pure and puts the numbers
    /// into the failure artifact.
    /// </summary>
    private static IReadOnlyDictionary<string, int> CountAssertedText(
        EvalWorkspace workspace, string sessionId, JsonElement invariants)
    {
        var needles = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var (group, property) in new[]
                 {
                     ("taskCompletion", "textPresent"),
                     ("taskCompletion", "textAbsent"),
                     ("collateral", "textPreserved"),
                 })
        {
            if (invariants.TryGetProperty(group, out var section)
                && section.TryGetProperty(property, out var array))
            {
                foreach (var value in array.EnumerateArray())
                {
                    if (value.GetString() is { Length: > 0 } needle)
                        needles.Add(needle);
                }
            }
        }

        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var needle in needles)
        {
            using var response = JsonDocument.Parse(workspace.Call("docxodus_search", new JsonObject
            {
                ["sessionId"] = sessionId,
                ["mode"] = "text",
                ["query"] = needle,
                ["caseSensitive"] = true,
                ["scope"] = "all",
            }));
            counts[needle] = response.RootElement.GetProperty("matches").GetArrayLength();
        }

        return counts;
    }

    private static void RunSteps(EvalWorkspace workspace, string sessionId, JsonElement steps)
    {
        foreach (var step in steps.EnumerateArray())
        {
            var tool = step.GetProperty("tool").GetString()!;
            var args = JsonNode.Parse(step.GetProperty("args").GetRawText())!.AsObject();
            args["sessionId"] = sessionId;
            if (step.TryGetProperty("target", out var target))
                args[TargetArgumentName(target)] = ResolveAnchor(workspace, sessionId, target);
            if (step.TryGetProperty("targets", out var targets))
            {
                // The key names the argument, so a target that also says `as` is contradictory
                // and must fail the scenario rather than silently prefer one of the two names.
                foreach (var entry in targets.EnumerateObject())
                {
                    if (entry.Value.TryGetProperty("as", out _))
                        throw new InvalidOperationException(
                            $"step '{tool}': targets['{entry.Name}'] must not declare 'as' — "
                            + "the key already names the argument.");
                    args[entry.Name] = ResolveAnchor(workspace, sessionId, entry.Value);
                }
            }

            ThrowOnFailedStep(tool, workspace.Call(tool, args));
        }
    }

    /// <summary>
    /// Edit tools report failure as a result with <c>success: false</c> rather than by throwing.
    /// A failed step must stop the run right here: letting it continue would surface as an
    /// invariant failure misattributed to the engine — or, for a needle that already exists in
    /// the fixture, not surface at all.
    /// </summary>
    private static void ThrowOnFailedStep(string tool, string response)
    {
        using var document = JsonDocument.Parse(response);
        if (document.RootElement.ValueKind == JsonValueKind.Object
            && document.RootElement.TryGetProperty("success", out var success)
            && success.ValueKind == JsonValueKind.False)
            throw new InvalidOperationException($"step '{tool}' failed: {response}");
    }

    private static string TargetArgumentName(JsonElement target) =>
        target.TryGetProperty("as", out var name) ? name.GetString()! : "anchorId";

    /// <summary>
    /// Resolve a target to an anchor id through docxodus_search. A match count below the requested
    /// index fails the scenario outright: fixture drift must surface as an error, never as a
    /// silently different target that then reports a passing precision metric.
    /// </summary>
    private static string ResolveAnchor(EvalWorkspace workspace, string sessionId, JsonElement target)
    {
        var index = target.TryGetProperty("index", out var value) ? value.GetInt32() : 0;
        var request = new JsonObject
        {
            ["sessionId"] = sessionId,
            ["mode"] = target.GetProperty("mode").GetString(),
            ["query"] = target.GetProperty("query").GetString(),
            ["caseSensitive"] =
                !target.TryGetProperty("caseSensitive", out var cased) || cased.GetBoolean(),
        };
        if (target.TryGetProperty("scope", out var scope))
            request["scope"] = scope.GetString();

        using var response = JsonDocument.Parse(workspace.Call("docxodus_search", request));
        var matches = response.RootElement.GetProperty("matches").EnumerateArray().ToList();
        if (matches.Count <= index)
            throw new InvalidOperationException(
                $"target '{target.GetProperty("query").GetString()}' "
                + $"({target.GetProperty("mode").GetString()}) matched {matches.Count} block(s); "
                + $"index {index} was requested. The fixture and the scenario have drifted apart.");

        var match = matches[index];
        // Text/regex matches nest the block under enclosingAnchor; kind/bookmark results are the
        // anchor target itself.
        return (match.TryGetProperty("enclosingAnchor", out var enclosing) ? enclosing : match)
            .GetProperty("id").GetString()!;
    }

    private static (IReadOnlyList<string> Anchors, IReadOnlyList<string> Parts) ReadChangedTargets(
        string semanticChangesJson)
    {
        using var document = JsonDocument.Parse(semanticChangesJson);
        var root = document.RootElement;
        if (!root.TryGetProperty("changes", out var changes))
            return (Array.Empty<string>(), Array.Empty<string>());

        var anchors = new SortedSet<string>(StringComparer.Ordinal);
        var parts = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var change in changes.EnumerateArray())
        {
            // rightAnchor is the delivered-side block; a pure removal only has a left side.
            foreach (var key in new[] { "rightAnchor", "leftAnchor" })
            {
                if (change.TryGetProperty(key, out var anchor)
                    && anchor.ValueKind == JsonValueKind.String
                    && anchor.GetString() is { Length: > 0 } id)
                {
                    anchors.Add(id);
                    break;
                }
            }

            if (change.TryGetProperty("partUri", out var partUri)
                && partUri.ValueKind == JsonValueKind.String
                && partUri.GetString() is { Length: > 0 } uri)
                parts.Add(uri);
        }

        return (anchors.ToList(), parts.ToList());
    }

    private static (int Added, int Removed) ComparePartInventories(byte[] opening, byte[] delivered)
    {
        var before = PartUris(opening);
        var after = PartUris(delivered);
        return (after.Except(before).Count(), before.Except(after).Count());
    }

    /// <summary>The #456 manifest's part-uri inventory, ordered for direct equality checks.</summary>
    public static IReadOnlyList<string> PartUris(byte[] packageBytes)
    {
        using var manifest = JsonDocument.Parse(
            VerificationOps.GeneratePackageManifest(packageBytes));
        return manifest.RootElement.GetProperty("entries").EnumerateArray()
            .Select(entry => entry.GetProperty("uri").GetString() ?? string.Empty)
            .OrderBy(uri => uri, StringComparer.Ordinal)
            .ToList();
    }

    private static string ReadStringProperty(string json, string property)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.GetProperty(property).GetString() ?? string.Empty;
    }

    private static string LocateCorpusRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "eval", "scenario.schema.json");
            if (File.Exists(candidate))
                return Path.Combine(directory.FullName, "eval");
            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            $"eval/ corpus not found above {AppContext.BaseDirectory}");
    }
}

/// <summary>
/// A disposable MCP server rooted at a private temporary directory. Scenarios name locations only
/// through the store, so a scenario physically cannot read or write outside its own workspace.
/// </summary>
internal sealed class EvalWorkspace : IDisposable
{
    private readonly string _root;
    private readonly SessionStore _store;

    public EvalWorkspace()
    {
        _root = Path.Combine(Path.GetTempPath(), $"docxodus-eval-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public string Open(byte[] packageBytes, string trackedChanges)
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, packageBytes);
        using var opened = JsonDocument.Parse(Call("docxodus_open", new JsonObject
        {
            ["path"] = path,
            ["trackedChanges"] = trackedChanges,
        }));
        return opened.RootElement.GetProperty("sessionId").GetString()!;
    }

    public byte[] Save(string sessionId)
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        Call("docxodus_save", new JsonObject { ["sessionId"] = sessionId, ["path"] = path });
        return File.ReadAllBytes(path);
    }

    public string Content(string sessionId, string format) =>
        Call("docxodus_get_content", new JsonObject
        {
            ["sessionId"] = sessionId,
            ["format"] = format,
        });

    public string Call(string tool, JsonObject args)
    {
        using var document = JsonDocument.Parse(args.ToJsonString());
        return Dispatcher.Call(_store, tool, document.RootElement.Clone());
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root))
            Directory.Delete(_root, recursive: true);
    }
}
