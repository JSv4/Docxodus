// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json.Nodes;
using Docxodus;

namespace LegalEval;

public sealed record BaselineExecution(
    byte[] Input,
    byte[] Output,
    byte[] Expected,
    string SemanticDiffJson,
    bool SemanticDiffSucceeded,
    IReadOnlyList<string> OperationLog,
    bool Succeeded,
    string? Error);

public sealed class BaselineExecutionException : Exception
{
    public BaselineExecutionException(BaselineExecution execution)
        : base(execution.Error ?? "Scripted baseline execution failed") => Execution = execution;

    public BaselineExecution Execution { get; }
}

/// <summary>
/// The deterministic non-model caller.  It uses the same anchor discovery and public session /
/// DocxDiff facades available to tool hosts, so a failing baseline is an engine/tool failure and
/// candidate planning is never scored on top of it.
/// </summary>
public sealed class ScriptedBaselineExecutor
{
    private const string FixedRevisionDate = "2026-01-15T12:00:00Z";
    private readonly IEvaluationPackageValidator _packageValidator;

    public ScriptedBaselineExecutor(IEvaluationPackageValidator? packageValidator = null) =>
        _packageValidator = packageValidator ?? new InterimEvaluationPackageValidator();

    public BaselineExecution Execute(LegalScenario scenario)
    {
        var execution = ExecuteCheckpointed(scenario);
        if (!execution.Succeeded) throw new BaselineExecutionException(execution);
        return execution;
    }

    public BaselineExecution ExecuteCheckpointed(LegalScenario scenario)
    {
        var input = File.ReadAllBytes(scenario.Fixture.Path);
        var expected = File.ReadAllBytes(scenario.ExpectedDocument.Path);
        _packageValidator.Validate(input, "evaluation input");
        _packageValidator.Validate(expected, "pinned expected document");
        var current = input;
        var log = new List<string>();
        string? error = null;

        for (var index = 0; index < scenario.BaselineOperations.Count; index++)
        {
            var operation = scenario.BaselineOperations[index];
            var kind = String(operation, "op");
            log.Add($"begin:{index}:{kind}");
            try
            {
                if (kind == "consolidate")
                {
                    current = ExecuteConsolidate(current, operation, log);
                }
                else
                {
                    using var session = new DocxSession(current, new DocxSessionSettings
                    {
                        CaptureInitialProjection = true,
                        PersistAnchorIds = false,
                        EmitMarkdownPatch = false,
                    });
                    Exception? operationError = null;
                    try
                    {
                        ExecuteSessionOperation(session, operation, log);
                    }
                    catch (Exception exception)
                    {
                        operationError = exception;
                        throw;
                    }
                    finally
                    {
                        // Preserve the latest writable checkpoint even when an operation reports a
                        // post-mutation failure. If saving itself fails, the preceding checkpoint is
                        // still retained by the outer catch.
                        try
                        {
                            current = GeneratedPackageNormalizer.Normalize(session.Save(false));
                        }
                        catch when (operationError is not null)
                        {
                            // The operation's original failure remains the primary diagnostic.
                        }
                    }
                }
                log.Add($"complete:{index}:{kind}");
            }
            catch (Exception exception)
            {
                error = ExceptionDetail(exception);
                log.Add($"failed:{index}:{kind}:{exception.GetType().Name}:{exception.Message}");
                break;
            }
        }

        string semanticDiff;
        var semanticDiffSucceeded = false;
        try
        {
            semanticDiff = DocxDiff.GetEditScriptJson(
                new WmlDocument("input.docx", input),
                new WmlDocument("output.docx", current),
                DiffSettings("Legal Evaluation Artifact"));
            semanticDiffSucceeded = true;
        }
        catch (Exception exception)
        {
            semanticDiff = EvaluationScorer.ErrorEnvelope(
                "semantic-diff", "failed", ExceptionDetail(exception));
            error ??= $"Final semantic diff generation failed: {ExceptionDetail(exception)}";
        }

        return new BaselineExecution(
            input, current, expected, semanticDiff, semanticDiffSucceeded,
            log, error is null, error);
    }

    private static void ExecuteSessionOperation(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var kind = String(operation, "op");
        switch (kind)
        {
            case "replaceText":
                ReplaceText(session, operation, log);
                break;
            case "insertNumberedClause":
                InsertNumberedClause(session, operation, log);
                break;
            case "replaceTableCell":
                ReplaceTableCell(session, operation, log);
                break;
            case "addReviewBundle":
                AddReviewBundle(session, operation, log);
                break;
            case "fillContentControl":
                FillContentControl(session, operation, log);
                break;
            case "failForTest":
                throw new InvalidOperationException(String(operation, "message"));
            default:
                throw new ScenarioValidationException($"unsupported scripted operation '{kind}'");
        }
    }

    private static void ReplaceText(DocxSession session, JsonObject operation, List<string> log)
    {
        var locator = String(operation, "anchorContains");
        var find = String(operation, "find");
        var replacement = String(operation, "replace");
        var anchor = UniqueAnchor(session, locator);
        if (Bool(operation, "tracked", false))
        {
            session.SetTrackedChanges(TrackedChangeMode.RenderInline);
            session.SetRevisionAuthor(String(operation, "author"));
        }
        var replacements = session.ReplaceTextRange(anchor, find, replacement);
        if (replacements.Count != 1)
            throw new InvalidOperationException(
                $"replaceText '{find}' resolved {replacements.Count} matches in anchor located by '{locator}'");
        Ensure(replacements[0], kind: "replaceText");
        log.Add($"replaceText:{locator}:{find}->{replacement}");
    }

    private static void InsertNumberedClause(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var anchor = UniqueAnchor(session, String(operation, "afterAnchorContains"));
        var result = session.InsertParagraph(anchor, Position.After, $"1. {String(operation, "text")}");
        Ensure(result, "insertNumberedClause");
        var created = result.Created.Single(value => value.Kind is "li" or "p").Id;
        Ensure(session.SetParagraphStyle(created, String(operation, "styleId")),
            "insertNumberedClause.style");
        var membership = session.GetListMembership(created);
        if (membership is null)
            throw new InvalidOperationException("insertNumberedClause did not inherit list numbering");
        log.Add($"insertNumberedClause:{created}");
    }

    private static void ReplaceTableCell(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var table = session.FindByKind("tbl").Single().Anchor.Id;
        var resolution = session.ResolveTableCellCoordinate(
            table, Int(operation, "row"), Int(operation, "column"));
        if (!resolution.Success || resolution.Cell is null)
            throw new InvalidOperationException($"replaceTableCell could not resolve cell: {resolution.Error?.Message}");
        Ensure(session.ReplaceCellContent(resolution.Cell.Anchor.Id, String(operation, "text")),
            "replaceTableCell");
        log.Add($"replaceTableCell:{resolution.Cell.Anchor.Id}");
    }

    private static void AddReviewBundle(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var anchor = UniqueAnchor(session, String(operation, "anchorContains"));
        var anchorText = session.Project().AnchorIndex[anchor].TextPreview;
        var spanText = String(operation, "spanText");
        var spanOffset = anchorText.IndexOf(spanText, StringComparison.Ordinal);
        if (spanOffset < 0)
            throw new InvalidOperationException($"review span '{spanText}' is absent");
        var date = DateTime.Parse(FixedRevisionDate, null,
            System.Globalization.DateTimeStyles.AdjustToUniversal
            | System.Globalization.DateTimeStyles.AssumeUniversal);
        var comment = session.AddComment(anchor, new CharSpan(spanOffset, spanText.Length),
            String(operation, "author"), String(operation, "comment"),
            String(operation, "initials"), date);
        Ensure(comment, "addReviewBundle.comment");
        var commentAnchor = comment.Created.Single(value => value.Kind == "cmt").Id;
        Ensure(session.AddCommentReply(commentAnchor, String(operation, "replyAuthor"),
            String(operation, "reply"), String(operation, "replyInitials"), date),
            "addReviewBundle.reply");
        Ensure(session.InsertFootnote(anchor, spanOffset + spanText.Length,
            String(operation, "footnote")), "addReviewBundle.footnote");
        Ensure(session.AddBookmark(String(operation, "bookmark"),
            DocumentRange.In(anchor, new CharSpan(spanOffset, spanText.Length))),
            "addReviewBundle.bookmark");
        var crossReference = session.InsertParagraph(anchor, Position.After,
            $"For the negotiated cap, see [{String(operation, "linkText")}](#{String(operation, "bookmark")}).");
        Ensure(crossReference, "addReviewBundle.crossReference");
        log.Add($"addReviewBundle:{commentAnchor}");
    }

    private static void FillContentControl(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var tag = String(operation, "tag");
        var controls = session.ListContentControls()
            .Where(value => string.Equals(value.Tag, tag, StringComparison.Ordinal)).ToList();
        if (controls.Count != 1)
            throw new InvalidOperationException($"content-control tag '{tag}' resolved {controls.Count} controls");
        Ensure(session.FillContentControlRichText(controls[0].AnchorId, String(operation, "text")),
            "fillContentControl");
        log.Add($"fillContentControl:{tag}");
    }

    private static byte[] ExecuteConsolidate(
        byte[] input, JsonObject operation, List<string> log)
    {
        var reviewersNode = operation["reviewers"] as JsonArray
            ?? throw new ScenarioValidationException("consolidate.reviewers must be an array");
        if (reviewersNode.Count < 2)
            throw new ScenarioValidationException("consolidate requires at least two reviewer documents");
        var reviewers = new List<DocxDiffReviewer>();
        foreach (var reviewerNode in reviewersNode)
        {
            var reviewer = reviewerNode as JsonObject
                ?? throw new ScenarioValidationException("consolidate reviewer must be an object");
            using var session = new DocxSession(input, new DocxSessionSettings
            {
                CaptureInitialProjection = false,
                EmitMarkdownPatch = false,
            });
            ExecuteSessionOperation(session, reviewer, log);
            reviewers.Add(new DocxDiffReviewer
            {
                Author = String(reviewer, "author"),
                Document = new WmlDocument("reviewer.docx",
                    GeneratedPackageNormalizer.Normalize(session.Save(false))),
            });
        }
        var settings = new DocxDiffConsolidateSettings
        {
            Diff = DiffSettings("Legal Evaluation"),
            ConflictResolution = ConflictResolution.FirstReviewerWins,
        };
        // The fixture deliberately has pre-existing revisions; consolidation owns only the accepted
        // shared-base view so its per-reviewer authorship is mechanically attributable.
        settings.Diff.PreAcceptInputRevisions = true;
        var output = DocxDiff.Consolidate(new WmlDocument("base.docx", input), reviewers, settings);
        log.Add($"consolidate:{reviewers.Count}");
        return GeneratedPackageNormalizer.Normalize(output.DocumentByteArray);
    }

    private static DocxDiffSettings DiffSettings(string author) => new()
    {
        AuthorForRevisions = author,
        Deterministic = true,
        DateTimeForRevisions = FixedRevisionDate,
        CompareHeadersFooters = true,
        TrackBlockFormatChanges = true,
    };

    private static string UniqueAnchor(DocxSession session, string text)
    {
        var matches = session.FindAllByText(text)
            .Where(value => value.Anchor.Scope == "body").ToList();
        if (matches.Count != 1)
            throw new InvalidOperationException($"locator '{text}' resolved {matches.Count} body anchors");
        return matches[0].Anchor.Id;
    }

    private static void Ensure(EditResult result, string kind)
    {
        if (!result.Success)
            throw new InvalidOperationException($"{kind} failed: {result.Error?.Code}: {result.Error?.Message}");
    }

    private static string String(JsonObject parent, string name) =>
        parent[name]?.GetValue<string>()
            ?? throw new ScenarioValidationException($"operation property '{name}' must be a string");

    private static int Int(JsonObject parent, string name) =>
        parent[name]?.GetValue<int>()
            ?? throw new ScenarioValidationException($"operation property '{name}' must be an integer");

    private static bool Bool(JsonObject parent, string name, bool fallback) =>
        parent[name]?.GetValue<bool>() ?? fallback;

    private static string ExceptionDetail(Exception exception)
    {
        var detail = $"{exception.GetType().Name}: {exception.Message}";
        var currentRoot = Path.GetFullPath(Directory.GetCurrentDirectory()).TrimEnd(
            Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        return detail.Replace(currentRoot, string.Empty, PathComparison);
    }

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
}
