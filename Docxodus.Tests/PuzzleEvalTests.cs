#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The puzzle eval (PE0xx): levels that ask whether the agent editing surface can reach a
/// specified document state, scored by <see cref="DocxDiff"/> rather than by judgement.
///
/// <para><b>What these tests are for.</b> They are not testing <see cref="DocxSession"/> — the DS
/// suites do that. They keep the LEVELS honest, which is the only thing a model-facing benchmark
/// needs from CI: that each level's par is actually achievable, that its target is reachable from
/// its start, and that the scoring function cannot be satisfied by doing nothing. A level whose
/// reference solution has quietly stopped solving it is a broken benchmark, and a broken benchmark
/// reports model failures that are really our failures.</para>
///
/// <para><b>Why DocxDiff is the scorer.</b> A solve has to be exact and unarguable, and the
/// comparison engine already answers "are these the same document?" in the only vocabulary that
/// matters — zero revisions between the player's result and the target. It also means the benchmark
/// exercises the differentiator the demo surfaces never touch.</para>
/// </summary>
public class PuzzleEvalTests
{
    // ─── Level pack ──────────────────────────────────────────────────────

    private sealed record LevelParagraph(string Text, string? Style);

    private sealed record LevelStep(
        string Op,
        string Find,
        string? RelativeTo,
        string? Position,
        string? Search,
        string? Replace);

    private sealed record Level(
        string Id,
        string Title,
        int Par,
        string Brief,
        IReadOnlyList<LevelParagraph> Start,
        IReadOnlyList<LevelParagraph> Target,
        IReadOnlyList<LevelStep> Reference);

    /// <summary>
    /// Walks up from the test binary rather than assuming a working directory: the suite runs from
    /// <c>bin/Debug/net10.0</c> under <c>dotnet test</c> and from the repo root under some IDEs.
    /// </summary>
    private static string PuzzlesRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "eval", "puzzles");
            if (Directory.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("eval/puzzles not found above " + AppContext.BaseDirectory);
    }

    public static TheoryData<string> LevelIds()
    {
        var data = new TheoryData<string>();
        foreach (var dir in Directory.EnumerateDirectories(PuzzlesRoot()).OrderBy(d => d, StringComparer.Ordinal))
            if (File.Exists(Path.Combine(dir, "level.json")))
                data.Add(Path.GetFileName(dir));
        return data;
    }

    private static Level LoadLevel(string id)
    {
        var path = Path.Combine(PuzzlesRoot(), id, "level.json");
        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        var root = doc.RootElement;

        static IReadOnlyList<LevelParagraph> Paragraphs(JsonElement side) =>
            side.GetProperty("paragraphs").EnumerateArray()
                .Select(p => new LevelParagraph(
                    p.GetProperty("text").GetString()!,
                    p.TryGetProperty("style", out var s) && s.ValueKind == JsonValueKind.String
                        ? s.GetString()
                        : null))
                .ToList();

        static string? Opt(JsonElement e, string name) =>
            e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

        return new Level(
            root.GetProperty("id").GetString()!,
            root.GetProperty("title").GetString()!,
            root.GetProperty("par").GetInt32(),
            root.GetProperty("brief").GetString()!,
            Paragraphs(root.GetProperty("start")),
            Paragraphs(root.GetProperty("target")),
            root.GetProperty("reference").EnumerateArray()
                .Select(s => new LevelStep(
                    s.GetProperty("op").GetString()!,
                    s.GetProperty("find").GetString()!,
                    Opt(s, "relativeTo"),
                    Opt(s, "position"),
                    Opt(s, "search"),
                    Opt(s, "replace")))
                .ToList());
    }

    // ─── Fixture construction ────────────────────────────────────────────

    /// <summary>
    /// Both sides of a level are built by this one function, so a scoring difference can only come
    /// from the player's edits — never from the two documents having been authored differently.
    /// </summary>
    private static byte[] BuildDocument(IReadOnlyList<LevelParagraph> paragraphs)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            foreach (var p in paragraphs)
            {
                var paragraph = new Paragraph(new Run(new Text(p.Text) { Space = SpaceProcessingModeValues.Preserve }));
                if (p.Style is not null)
                    paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = p.Style });
                body.Append(paragraph);
            }

            main.Document.Save();
        }

        return ms.ToArray();
    }

    // ─── Scoring ─────────────────────────────────────────────────────────

    /// <summary>
    /// The win condition: zero revisions between the player's document and the target. Returning
    /// the revision list rather than a bool so a failing assertion can say WHAT still differs,
    /// which is the difference between a usable benchmark and a red X.
    /// </summary>
    private static IReadOnlyList<DocxDiffRevision> ScoreAgainstTarget(byte[] player, byte[] target) =>
        DocxDiff.GetRevisions(
            new WmlDocument("player.docx", player),
            new WmlDocument("target.docx", target));

    private static string Describe(IReadOnlyList<DocxDiffRevision> revisions) =>
        revisions.Count == 0
            ? "solved"
            : string.Join("; ", revisions.Take(8).Select(r => $"{r.Type}:{Trim(r.Text)}"));

    private static string Trim(string? text) =>
        text is null ? "" : text.Length <= 40 ? text : text[..40] + "…";

    // ─── Content-addressed solution runner ───────────────────────────────

    /// <summary>
    /// Resolve a block the way a player has to: by what it says, not by an id nobody has yet. The
    /// harness deliberately gets no privileged addressing — if a level could only be solved with
    /// ids handed out in advance, it would not be measuring the surface an agent actually faces.
    /// </summary>
    private static string FindAnchor(DocxSession session, string needle)
    {
        // FindAllByText is the same op behind docxodus_search, so the reference solution pays the
        // same discovery cost a player does — including the part where a needle has to be chosen
        // specifically enough to land on one block.
        var hits = session.FindAllByText(needle, null)
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
            .ToList();

        if (hits.Count == 0) throw new InvalidOperationException($"no body block contains '{needle}'");
        return hits[0].Anchor.Id;
    }

    /// <summary>Applies a level's reference solution and returns the number of mutating calls it
    /// took — the value compared against par.</summary>
    private static int RunReference(DocxSession session, Level level)
    {
        int calls = 0;
        foreach (var step in level.Reference)
        {
            var anchor = FindAnchor(session, step.Find);
            switch (step.Op)
            {
                case "moveBlock":
                {
                    var relativeTo = FindAnchor(session, step.RelativeTo!);
                    var position = step.Position == "before" ? Position.Before : Position.After;
                    var result = session.MoveBlock(anchor, relativeTo, position);
                    Assert.True(result.Success, $"{level.Id} step {calls}: {result.Error?.Message}");
                    break;
                }

                case "replaceTextRange":
                {
                    var results = session.ReplaceTextRange(anchor, step.Search!, step.Replace!, null);
                    Assert.All(results, r => Assert.True(r.Success, $"{level.Id} step {calls}: {r.Error?.Message}"));
                    Assert.NotEmpty(results);
                    break;
                }

                default:
                    throw new InvalidOperationException($"unknown reference op '{step.Op}'");
            }

            calls++;
        }

        return calls;
    }

    // ─── The three properties that keep a level honest ───────────────────

    /// <summary>
    /// The level is solvable, and the reference solution solves it. If this fails, the benchmark is
    /// reporting a surface limitation as a model failure.
    /// </summary>
    [Theory]
    [MemberData(nameof(LevelIds))]
    public void PE001_ReferenceSolution_ReachesZeroRevisionsAgainstTheTarget(string levelId)
    {
        var level = LoadLevel(levelId);
        var target = BuildDocument(level.Target);

        using var session = new DocxSession(BuildDocument(level.Start));
        RunReference(session, level);

        var revisions = ScoreAgainstTarget(session.Save(), target);
        Assert.True(revisions.Count == 0, $"{level.Id} unsolved: {Describe(revisions)}");
    }

    /// <summary>
    /// Par is achievable. Par is the score every model run is reported against, so a par nobody can
    /// hit is a benchmark that reports everyone as below average.
    /// </summary>
    [Theory]
    [MemberData(nameof(LevelIds))]
    public void PE002_ReferenceSolution_SolvesWithinPar(string levelId)
    {
        var level = LoadLevel(levelId);
        using var session = new DocxSession(BuildDocument(level.Start));

        var calls = RunReference(session, level);

        Assert.True(calls <= level.Par, $"{level.Id} reference took {calls} calls against par {level.Par}");
    }

    /// <summary>
    /// The starting document does NOT already score as solved. Without this, a level whose target
    /// was built wrong — or a scorer that silently returns nothing — passes PE001 with an empty
    /// solution, and the whole pack reports 100%.
    /// </summary>
    [Theory]
    [MemberData(nameof(LevelIds))]
    public void PE003_StartDocument_DoesNotAlreadyScoreAsSolved(string levelId)
    {
        var level = LoadLevel(levelId);

        var revisions = ScoreAgainstTarget(BuildDocument(level.Start), BuildDocument(level.Target));

        Assert.True(revisions.Count > 0, $"{level.Id} start already equals target — the level asks for nothing");
    }

    /// <summary>
    /// A near-miss must score as unsolved. The scorer is the whole benchmark, so "does it actually
    /// reject a wrong answer" deserves a test of its own rather than being assumed from PE003 —
    /// here, the clauses reordered correctly but the defined term left unconformed.
    /// </summary>
    [Fact]
    public void PE004_Scorer_RejectsAPartialSolve()
    {
        var level = LoadLevel("L01-clause-order");
        var target = BuildDocument(level.Target);

        using var session = new DocxSession(BuildDocument(level.Start));
        foreach (var step in level.Reference.Where(s => s.Op == "moveBlock"))
        {
            var anchor = FindAnchor(session, step.Find);
            var relativeTo = FindAnchor(session, step.RelativeTo!);
            Assert.True(session
                .MoveBlock(anchor, relativeTo, step.Position == "before" ? Position.Before : Position.After)
                .Success);
        }

        var revisions = ScoreAgainstTarget(session.Save(), target);
        Assert.True(revisions.Count > 0, "the scorer accepted a document that still says \"Acme\"");
    }

    /// <summary>
    /// The whole reference solution as ONE batch: same document, same score, one undo step. A level
    /// is a plan, and a plan is the case batching exists for — so the pack asserts the property
    /// rather than leaving it to the DS suites in the abstract.
    /// </summary>
    [Fact]
    public void PE005_ReferenceSolution_AppliedAsOneBatch_ScoresTheSameAndUndoesOnce()
    {
        var level = LoadLevel("L01-clause-order");
        var target = BuildDocument(level.Target);
        var start = BuildDocument(level.Start);

        using var session = new DocxSession(start);
        var batch = session.Batch(new[] { (Func<EditResult>)(() =>
        {
            RunReference(session, level);
            return new EditResult { Success = true };
        }) });

        Assert.True(batch.Success);
        Assert.Empty(ScoreAgainstTarget(session.Save(), target));

        Assert.Equal(1, session.UndoCount);
        Assert.True(session.Undo());
        Assert.Empty(ScoreAgainstTarget(session.Save(), start));
    }
}
