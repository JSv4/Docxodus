// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json.Nodes;
using Docxodus;
using Docxodus.Verification;

namespace LegalEval;

public sealed record BaselineExecution(
    byte[] Input,
    byte[] Output,
    byte[] Expected,
    string SemanticDiffJson,
    bool SemanticDiffSucceeded,
    IReadOnlyList<string> OperationLog,
    bool Succeeded,
    string? Error,
    PackageManifest? InputManifest = null,
    PackageManifest? OutputManifest = null,
    PackageManifest? ExpectedManifest = null,
    IReadOnlyList<BaselineTransactionTrace>? TransactionTraces = null,
    string? DeliveryReceiptUnavailableReason = null);

public sealed record BaselineTransactionTrace(
    byte[] BeforeBytes,
    byte[] AfterBytes,
    PackageManifest BeforeManifest,
    PackageManifest AfterManifest,
    MutationBatchResult Result,
    DeliveryNormalizedOperation Operation);

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
    private const int MaximumSemanticChanges = 4096;
    private const long MaximumRetainedTraceBytes = 512L * 1024 * 1024;
    private const long MaximumRetainedReviewerBytes = 512L * 1024 * 1024;
    private static readonly DateTime FixedRevisionTimestamp = DateTime.Parse(
        FixedRevisionDate, null,
        System.Globalization.DateTimeStyles.AdjustToUniversal
        | System.Globalization.DateTimeStyles.AssumeUniversal);
    private readonly IEvaluationPackageValidator _packageValidator;

    public ScriptedBaselineExecutor(IEvaluationPackageValidator? packageValidator = null) =>
        _packageValidator = packageValidator ?? new EvaluationPackageValidator();

    public BaselineExecution Execute(LegalScenario scenario)
    {
        var execution = ExecuteCheckpointed(scenario);
        if (!execution.Succeeded) throw new BaselineExecutionException(execution);
        return execution;
    }

    public BaselineExecution ExecuteCheckpointed(LegalScenario scenario)
    {
        var input = ReadBounded(scenario.Fixture.Path, _packageValidator.MaximumPackageBytes,
            "evaluation input");
        var expected = ReadBounded(scenario.ExpectedDocument.Path,
            _packageValidator.MaximumPackageBytes, "pinned expected document");
        RequireSha256(input, scenario.Fixture.SourceSha256, "evaluation input");
        RequireSha256(expected, scenario.ExpectedDocument.SourceSha256,
            "pinned expected document");
        var inputManifest = _packageValidator.Inspect(input, "evaluation input");
        var expectedManifest = _packageValidator.Inspect(expected, "pinned expected document");
        var current = input;
        var outputManifest = inputManifest;
        var log = new List<string>();
        var traces = new List<BaselineTransactionTrace>();
        long retainedTraceBytes = 0;
        string? error = null;
        string? receiptUnavailableReason = null;
        DocxSession? session = null;

        try
        {
            for (var index = 0; index < scenario.BaselineOperations.Count; index++)
            {
                var operation = scenario.BaselineOperations[index];
                var kind = String(operation, "op");
                log.Add($"begin:{index}:{kind}");
                if (kind == "consolidate")
                {
                    if (session is not null || traces.Count != 0)
                        throw new ScenarioValidationException(
                            "consolidate cannot be mixed with session-backed baseline operations");
                    current = ExecuteConsolidate(current, operation, log);
                    outputManifest = _packageValidator.Inspect(current, "scripted baseline output");
                    receiptUnavailableReason =
                        "the consolidate facade does not expose an authoritative ExecuteBatch trace";
                }
                else
                {
                    session ??= new DocxSession(current, new DocxSessionSettings
                    {
                        CaptureInitialProjection = true,
                        PersistAnchorIds = false,
                        EmitMarkdownPatch = false,
                        UtcNowProvider = static () => FixedRevisionTimestamp,
                        DeterministicPackageOutput = true,
                    });
                    var before = session.Save(false);
                    var beforeManifest = _packageValidator.Inspect(
                        before, $"scripted operation {index} input");
                    var normalized = DeliveryNormalizedOperation.Create(
                        "legal_eval", kind, operation.ToJsonString());
                    var result = session.ExecuteBatch(new[]
                    {
                        new MutationBatchStep("legal_eval", kind,
                            value => ExecuteSessionOperation(value, operation, log)),
                    });
                    var after = session.Save(false);
                    var afterManifest = _packageValidator.Inspect(
                        after, $"scripted operation {index} output");
                    retainedTraceBytes = checked(retainedTraceBytes + before.Length + after.Length);
                    if (retainedTraceBytes > MaximumRetainedTraceBytes)
                        throw new ScenarioValidationException(
                            $"baseline transaction evidence exceeds the {MaximumRetainedTraceBytes}-byte limit");
                    traces.Add(new BaselineTransactionTrace(
                        before, after, beforeManifest, afterManifest, result, normalized));
                    current = after;
                    outputManifest = afterManifest;
                    if (!result.Success)
                    {
                        var failure = result.Failure?.Error;
                        throw new InvalidOperationException(
                            $"{kind} failed: {failure?.Code}: {failure?.Message}");
                    }
                }
                log.Add($"complete:{index}:{kind}");
            }
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            error = ExceptionDetail(exception);
            log.Add($"failed:{exception.GetType().Name}:{exception.Message}");
        }
        finally
        {
            session?.Dispose();
        }

        string semanticDiff;
        var semanticDiffSucceeded = false;
        try
        {
            semanticDiff = SemanticDiff.CompareBounded(
                new WmlDocument("input.docx", input),
                new WmlDocument("output.docx", current),
                new SemanticDiffOptions
                {
                    PackageOptions = _packageValidator.ManifestOptions,
                },
                MaximumSemanticChanges).ToCanonicalJson();
            semanticDiffSucceeded = true;
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            semanticDiff = EvaluationScorer.ErrorEnvelope(
                "semantic-diff", "failed", ExceptionDetail(exception));
            error ??= $"Final semantic diff generation failed: {ExceptionDetail(exception)}";
        }

        return new BaselineExecution(
            input, current, expected, semanticDiff, semanticDiffSucceeded,
            log, error is null, error,
            inputManifest, outputManifest, expectedManifest, traces,
            receiptUnavailableReason);
    }

    private static IReadOnlyList<EditResult> ExecuteSessionOperation(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var kind = String(operation, "op");
        switch (kind)
        {
            case "replaceText":
                return ReplaceText(session, operation, log);
            case "insertNumberedClause":
                return InsertNumberedClause(session, operation, log);
            case "replaceTableCell":
                return ReplaceTableCell(session, operation, log);
            case "addReviewBundle":
                return AddReviewBundle(session, operation, log);
            case "fillContentControl":
                return FillContentControl(session, operation, log);
            case "failForTest":
                throw new InvalidOperationException(String(operation, "message"));
            default:
                throw new ScenarioValidationException($"unsupported scripted operation '{kind}'");
        }
    }

    private static IReadOnlyList<EditResult> ReplaceText(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var locator = String(operation, "anchorContains");
        var find = String(operation, "find");
        var replacement = String(operation, "replace");
        var anchor = UniqueAnchor(session, locator);
        var tracked = Bool(operation, "tracked", false);
        // Tracked-change settings are session state. Set both branches explicitly so a tracked
        // replacement cannot leak into a later operation that requested an ordinary edit.
        session.SetTrackedChanges(tracked
            ? TrackedChangeMode.RenderInline
            : TrackedChangeMode.Accept);
        session.SetRevisionAuthor(tracked ? String(operation, "author") : null);
        var replacements = session.ReplaceTextRange(anchor, find, replacement);
        if (replacements.Count != 1)
            throw new InvalidOperationException(
                $"replaceText '{find}' resolved {replacements.Count} matches in anchor located by '{locator}'");
        if (replacements[0].Success)
            log.Add($"replaceText:{locator}:{find}->{replacement}");
        return replacements;
    }

    private static IReadOnlyList<EditResult> InsertNumberedClause(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var anchor = UniqueAnchor(session, String(operation, "afterAnchorContains"));
        var result = session.InsertParagraph(anchor, Position.After, $"1. {String(operation, "text")}");
        if (!result.Success) return new[] { result };
        var created = result.Created.Single(value => value.Kind is "li" or "p").Id;
        var style = session.SetParagraphStyle(created, String(operation, "styleId"));
        if (!style.Success) return new[] { result, style };
        var membership = session.GetListMembership(created);
        if (membership is null)
            return new[]
            {
                result,
                style,
                EditResult.Fail(EditErrorCode.InternalError,
                    "insertNumberedClause did not inherit list numbering"),
            };
        log.Add($"insertNumberedClause:{created}");
        return new[] { result, style };
    }

    private static IReadOnlyList<EditResult> ReplaceTableCell(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var table = session.FindByKind("tbl").Single().Anchor.Id;
        var resolution = session.ResolveTableCellCoordinate(
            table, Int(operation, "row"), Int(operation, "column"));
        if (!resolution.Success || resolution.Cell is null)
            throw new InvalidOperationException($"replaceTableCell could not resolve cell: {resolution.Error?.Message}");
        var result = session.ReplaceCellContent(
            resolution.Cell.Anchor.Id, String(operation, "text"));
        if (result.Success) log.Add($"replaceTableCell:{resolution.Cell.Anchor.Id}");
        return new[] { result };
    }

    private static IReadOnlyList<EditResult> AddReviewBundle(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var anchor = UniqueAnchor(session, String(operation, "anchorContains"));
        var anchorText = session.Project().AnchorIndex[anchor].TextPreview;
        var spanText = String(operation, "spanText");
        var spanOffset = anchorText.IndexOf(spanText, StringComparison.Ordinal);
        if (spanOffset < 0)
            throw new InvalidOperationException($"review span '{spanText}' is absent");
        var date = FixedRevisionTimestamp;
        var comment = session.AddComment(anchor, new CharSpan(spanOffset, spanText.Length),
            String(operation, "author"), String(operation, "comment"),
            String(operation, "initials"), date);
        if (!comment.Success) return new[] { comment };
        var commentAnchor = comment.Created.Single(value => value.Kind == "cmt").Id;
        var reply = session.AddCommentReply(commentAnchor, String(operation, "replyAuthor"),
            String(operation, "reply"), String(operation, "replyInitials"), date);
        if (!reply.Success) return new[] { comment, reply };
        var footnote = session.InsertFootnote(anchor, spanOffset + spanText.Length,
            String(operation, "footnote"));
        if (!footnote.Success) return new[] { comment, reply, footnote };
        var bookmark = session.AddBookmark(String(operation, "bookmark"),
            DocumentRange.In(anchor, new CharSpan(spanOffset, spanText.Length)));
        if (!bookmark.Success) return new[] { comment, reply, footnote, bookmark };
        var crossReference = session.InsertParagraph(anchor, Position.After,
            $"For the negotiated cap, see [{String(operation, "linkText")}](#{String(operation, "bookmark")}).");
        if (!crossReference.Success)
            return new[] { comment, reply, footnote, bookmark, crossReference };
        log.Add($"addReviewBundle:{commentAnchor}");
        return new[] { comment, reply, footnote, bookmark, crossReference };
    }

    private static IReadOnlyList<EditResult> FillContentControl(
        DocxSession session, JsonObject operation, List<string> log)
    {
        var tag = String(operation, "tag");
        var controls = session.ListContentControls()
            .Where(value => string.Equals(value.Tag, tag, StringComparison.Ordinal)).ToList();
        if (controls.Count != 1)
            throw new InvalidOperationException($"content-control tag '{tag}' resolved {controls.Count} controls");
        var result = session.FillContentControlRichText(
            controls[0].AnchorId, String(operation, "text"));
        if (result.Success) log.Add($"fillContentControl:{tag}");
        return new[] { result };
    }

    private static byte[] ExecuteConsolidate(
        byte[] input, JsonObject operation, List<string> log)
    {
        var reviewersNode = operation["reviewers"] as JsonArray
            ?? throw new ScenarioValidationException("consolidate.reviewers must be an array");
        if (reviewersNode.Count < 2)
            throw new ScenarioValidationException("consolidate requires at least two reviewer documents");
        var reviewers = new List<DocxDiffReviewer>();
        long retainedReviewerBytes = 0;
        foreach (var reviewerNode in reviewersNode)
        {
            var reviewer = reviewerNode as JsonObject
                ?? throw new ScenarioValidationException("consolidate reviewer must be an object");
            using var session = new DocxSession(input, new DocxSessionSettings
            {
                CaptureInitialProjection = false,
                EmitMarkdownPatch = false,
                UtcNowProvider = static () => FixedRevisionTimestamp,
                DeterministicPackageOutput = true,
            });
            var results = ExecuteSessionOperation(session, reviewer, log);
            if (results.Any(result => !result.Success))
                throw new InvalidOperationException(
                    $"consolidate reviewer operation failed: {results.First(result => !result.Success).Error?.Message}");
            var reviewerBytes = session.Save(false);
            retainedReviewerBytes = checked(retainedReviewerBytes + reviewerBytes.Length);
            if (retainedReviewerBytes > MaximumRetainedReviewerBytes)
                throw new ScenarioValidationException(
                    $"consolidation reviewer evidence exceeds the {MaximumRetainedReviewerBytes}-byte limit");
            reviewers.Add(new DocxDiffReviewer
            {
                Author = String(reviewer, "author"),
                Document = new WmlDocument("reviewer.docx", reviewerBytes),
            });
        }
        var settings = new DocxDiffConsolidateSettings
        {
            Diff = DiffSettings("Legal Evaluation"),
            ConflictResolution = ConflictResolution.FirstReviewerWins,
        };
        // The selected base carries pre-existing legal review state.  Consolidation must preserve
        // that state while attributing only the two new reviewer deltas to their declared authors.
        settings.Diff.PreAcceptInputRevisions = false;
        settings.Diff.PreserveInputRevisions = true;
        var output = DocxDiff.Consolidate(new WmlDocument("base.docx", input), reviewers, settings);
        log.Add($"consolidate:{reviewers.Count}");
        return output.DocumentByteArray;
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

    private static string String(JsonObject parent, string name) =>
        parent[name]?.GetValue<string>()
            ?? throw new ScenarioValidationException($"operation property '{name}' must be a string");

    private static int Int(JsonObject parent, string name) =>
        parent[name]?.GetValue<int>()
            ?? throw new ScenarioValidationException($"operation property '{name}' must be an integer");

    private static bool Bool(JsonObject parent, string name, bool fallback) =>
        parent[name]?.GetValue<bool>() ?? fallback;

    private static byte[] ReadBounded(string path, long maximumBytes, string label)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 81920, FileOptions.SequentialScan);
        if (stream.Length > maximumBytes)
            throw new ScenarioValidationException(
                $"{label} exceeds the {maximumBytes}-byte package limit");
        var bytes = new byte[checked((int)stream.Length)];
        stream.ReadExactly(bytes);
        if (stream.ReadByte() != -1)
            throw new ScenarioValidationException($"{label} changed while it was being read");
        return bytes;
    }

    private static void RequireSha256(byte[] bytes, string expected, string label)
    {
        var actual = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();
        if (!string.Equals(actual, expected, StringComparison.Ordinal))
            throw new ScenarioValidationException(
                $"{label} changed after corpus validation: expected SHA-256 {expected}, actual {actual}");
    }

    private static string ExceptionDetail(Exception exception)
    {
        var detail = $"{exception.GetType().Name}: {exception.Message}";
        var currentRoot = Path.GetFullPath(Directory.GetCurrentDirectory()).TrimEnd(
            Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        return detail.Replace(currentRoot, string.Empty, PathComparison);
    }

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
