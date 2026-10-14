// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #760: a preview retained for commit is restored byte-for-byte, so the generated ids,
/// timestamps and package hash a caller was shown are exactly what the live document gets.
/// </summary>
[Collection("MCP session registry isolation")]
public class DocxSessionPreviewCommitTests
{
    [Fact]
    public void DS760a_GuardedCommit_RestoresPreviewedGeneratedIdsTimestampsAndHash()
    {
        using var session = Open(new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Preview Author",
        });
        var anchors = BodyParagraphs(session);
        var baseVersion = session.Version;
        var baseHash = session.GetPackageContentHash();

        var preview = session.PreviewBatch(GeneratingSteps(anchors), options: Retain);

        Assert.True(preview.Success, Describe(preview));
        Assert.True(preview.Preview);
        var retention = Assert.IsType<MutationPreviewRetention>(preview.Retention);
        Assert.StartsWith("pv-", retention.PreviewId, StringComparison.Ordinal);
        Assert.Equal(baseVersion, retention.BaseVersion);
        Assert.Equal(baseHash, retention.BasePackageHash);
        Assert.NotNull(preview.PackageHash);
        Assert.NotEqual(baseHash, preview.PackageHash);
        // The preview genuinely exercised generated values: the receipt carries the caveats.
        Assert.Contains(preview.Warnings, warning => warning.Contains("Created anchors", StringComparison.Ordinal));
        Assert.Contains(preview.Warnings, warning => warning.Contains("execution clock", StringComparison.Ordinal));
        // Retention is not a mutation.
        Assert.Equal(baseVersion, session.Version);
        Assert.Equal(baseHash, session.GetPackageContentHash());
        Assert.False(session.Undo());

        var commit = session.CommitPreview(retention.PreviewId);

        Assert.True(commit.Success, Describe(commit));
        Assert.False(commit.Preview);
        Assert.Equal(retention, commit.Retention);
        Assert.Equal(preview.ResultVersion, commit.ResultVersion);
        Assert.Equal(commit.ResultVersion, session.Version);
        Assert.Equal(preview.PackageHash, commit.PackageHash);
        Assert.Equal(preview.PackageHash, session.GetPackageContentHash());
        Assert.Equal(CreatedIds(preview), CreatedIds(commit));
        Assert.Equal(
            DocxSessionJson.SerializeCommentList(preview.CommentChanges.Added),
            DocxSessionJson.SerializeCommentList(session.ListComments()));
        Assert.Equal(
            DocxSessionJson.SerializeRevisionList(preview.RevisionChanges.Added),
            DocxSessionJson.SerializeRevisionList(session.ListRevisions()));
        Assert.All(CreatedIds(commit), id => Assert.Contains(id, session.Project().AnchorIndex.Keys));
        Assert.DoesNotContain(commit.Warnings, warning => warning.Contains("may be generated", StringComparison.Ordinal));
        Assert.DoesNotContain(commit.Warnings, warning => warning.Contains("execution clock", StringComparison.Ordinal));
        Assert.Null(commit.Html);

        // One history step, composable with undo and redo.
        Assert.True(session.Undo());
        Assert.Equal(baseHash, session.GetPackageContentHash());
        Assert.False(session.Undo());
        Assert.True(session.Redo());
        Assert.Equal(preview.PackageHash, session.GetPackageContentHash());
    }

    [Fact]
    public void DS760b_DeterministicCommit_MatchesApplyingTheSameBatchToTheSameBase()
    {
        using var session = Open();
        var anchor = BodyParagraphs(session)[0];
        MutationBatchStep[] Steps() => new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchor, "Deterministic edit.")),
        };

        var preview = session.PreviewBatch(Steps(), MutationBatchMode.BestEffort, Retain);
        var commit = session.CommitPreview(preview.Retention!.PreviewId);
        Assert.True(commit.Success, Describe(commit));
        Assert.Equal(MutationBatchMode.BestEffort, commit.Mode);

        // Undo restores the exact base bytes the shadow was cloned from; applying the same
        // batch there directly is the fresh execution a retained commit stands in for.
        Assert.True(session.Undo());
        var direct = session.ExecuteBatch(Steps(), MutationBatchMode.BestEffort);

        Assert.True(direct.Success, Describe(direct));
        Assert.Equal(preview.PackageHash, commit.PackageHash);
        Assert.Equal(direct.PackageHash, commit.PackageHash);
        Assert.Equal(
            DocxSessionJson.SerializeMutationBatchResult(
                direct with { BaseVersion = commit.BaseVersion, ResultVersion = commit.ResultVersion }),
            DocxSessionJson.SerializeMutationBatchResult(commit with { Retention = null }));
    }

    [Fact]
    public void DS760c_StaleCommit_RefusesWithoutEditingWhenTheBaseMoved()
    {
        using var session = Open(new DocxSessionSettings { PersistAnchorIds = true });
        var anchors = BodyParagraphs(session);
        var preview = session.PreviewBatch(GeneratingSteps(anchors), options: Retain);
        var previewId = preview.Retention!.PreviewId;

        // A configuration change alone is enough to refuse; undoing it re-arms the commit
        // because the session version and package are untouched.
        session.SetRevisionAuthor("Someone Else");
        var settingsStale = session.CommitPreview(previewId);
        Assert.Equal(EditErrorCode.PreviewStale, settingsStale.Failure!.Error.Code);
        Assert.Contains("revision author", settingsStale.Failure.Error.Message, StringComparison.Ordinal);
        session.SetRevisionAuthor(null);

        Assert.True(session.ReplaceText(anchors[1], "An intervening live edit.").Success);
        var version = session.Version;
        var hash = session.GetPackageContentHash();

        var stale = session.CommitPreview(previewId);

        Assert.False(stale.Success);
        Assert.False(stale.RolledBack);
        Assert.Equal(EditErrorCode.PreviewStale, stale.Failure!.Error.Code);
        Assert.Contains($"version {version}", stale.Failure.Error.Message, StringComparison.Ordinal);
        Assert.Equal("commit_preview", stale.Failure.Action);
        Assert.Equal(preview.Retention, stale.Retention);
        Assert.Equal(version, stale.BaseVersion);
        Assert.Equal(version, stale.ResultVersion);
        Assert.Equal(version, session.Version);
        Assert.Equal(hash, session.GetPackageContentHash());
        // Nothing was recorded: the only history entry is the intervening edit.
        Assert.True(session.Undo());
        Assert.False(session.Undo());
        // The version is monotonic, so even the byte-identical base after that undo is a
        // different session state; a caller previews again rather than committing history.
        Assert.Equal(EditErrorCode.PreviewStale, session.CommitPreview(previewId).Failure!.Error.Code);
    }

    [Fact]
    public void DS760d_FailedPreview_IsNotRetainedAndUnknownIdsAreRefused()
    {
        using var session = Open();
        var preview = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText("p:body:missing", "x")),
        }, options: Retain);

        Assert.False(preview.Success);
        Assert.Null(preview.Retention);
        Assert.Contains(preview.Warnings, warning => warning.Contains("not retained", StringComparison.Ordinal));

        var missing = session.CommitPreview("pv-0000");
        Assert.Equal(EditErrorCode.PreviewNotFound, missing.Failure!.Error.Code);
        Assert.Null(missing.Retention);
        Assert.Equal(0, session.Version);
        Assert.Throws<ArgumentException>(() => session.CommitPreview("  "));
    }

    [Fact]
    public void DS760e_CommitConsumesThePreview()
    {
        using var session = Open();
        var anchor = BodyParagraphs(session)[0];
        var preview = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchor, "Once.")),
        }, options: Retain);
        var previewId = preview.Retention!.PreviewId;

        Assert.True(session.CommitPreview(previewId).Success);
        var version = session.Version;

        var again = session.CommitPreview(previewId);

        Assert.Equal(EditErrorCode.PreviewNotFound, again.Failure!.Error.Code);
        Assert.Equal(version, session.Version);
        Assert.Equal(0, session.RetainedPreviews.Count);
    }

    [Fact]
    public void DS760f_RetentionIsBoundedByCountBytesAndTime()
    {
        var now = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
        using var session = Open();
        session.RetainedPreviews = new RetainedPreviews(
            capacity: 1, timeToLive: TimeSpan.FromMinutes(5), utcNow: () => now);
        var anchor = BodyParagraphs(session)[0];
        MutationBatchStep[] Edit(string text) => new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchor, text)),
        };

        var first = session.PreviewBatch(Edit("First."), options: Retain).Retention!;
        Assert.Equal(now.AddMinutes(5), first.ExpiresAt);
        var second = session.PreviewBatch(Edit("Second."), options: Retain).Retention!;
        Assert.Equal(1, session.RetainedPreviews.Count);

        // Capacity: the oldest entry made room for the newest.
        Assert.Equal(EditErrorCode.PreviewNotFound, session.CommitPreview(first.PreviewId).Failure!.Error.Code);

        // Time: the survivor expires without being committed.
        now = now.AddMinutes(6);
        Assert.Equal(EditErrorCode.PreviewNotFound, session.CommitPreview(second.PreviewId).Failure!.Error.Code);
        Assert.Equal(0, session.RetainedPreviews.Count);
        Assert.Equal(0, session.Version);

        // Bytes: a package larger than the whole budget is refused up front, with a warning.
        session.RetainedPreviews = new RetainedPreviews(byteBudget: 16);
        var oversized = session.PreviewBatch(Edit("Third."), options: Retain);
        Assert.True(oversized.Success);
        Assert.Null(oversized.Retention);
        Assert.Contains(oversized.Warnings, warning => warning.Contains("byte budget", StringComparison.Ordinal));
        Assert.Equal(0, session.RetainedPreviews.Count);
    }

    [Fact]
    public void DS760g_HandleFacade_RetainsCommitsAndReplaysUnderATransactionId()
    {
        var handle = DocxSessionOps.OpenSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { PersistAnchorIds = true });
        try
        {
            string anchor;
            using (var projection = JsonDocument.Parse(DocxSessionOps.Project(handle)))
            {
                anchor = projection.RootElement.GetProperty("anchorIndex").EnumerateObject()
                    .First(property => property.Name.StartsWith("p:body:", StringComparison.Ordinal)).Name;
            }

            var previewJson = DocxSessionOps.PreviewBatch(
                handle,
                MutationBatchMode.Atomic,
                shadowHandle => new[]
                {
                    DocxSessionOps.SerializedBatchStep("docx_scalpel", "insert_paragraph",
                        () => DocxSessionOps.InsertParagraph(shadowHandle, anchor, Position.After, "Committed via handle.")),
                },
                Retain);
            using var preview = JsonDocument.Parse(previewJson);
            Assert.True(preview.RootElement.GetProperty("success").GetBoolean());
            var retention = preview.RootElement.GetProperty("retention");
            var previewId = retention.GetProperty("previewId").GetString()!;
            Assert.Equal(0, retention.GetProperty("baseVersion").GetInt64());
            Assert.Equal(DocxSessionOps.GetPackageContentHash(handle), retention.GetProperty("basePackageHash").GetString());
            Assert.EndsWith("Z", retention.GetProperty("expiresAt").GetString()!, StringComparison.Ordinal);
            Assert.Equal(0, DocxSessionOps.GetVersion(handle));

            var request = JsonSerializer.SerializeToElement(new { handle, previewId });
            var first = DocxSessionOps.CommitPreviewTransactional(handle, "tx-commit", request, previewId);
            var replay = DocxSessionOps.CommitPreviewTransactional(handle, "tx-commit", request, previewId);

            Assert.Equal(first, replay);
            using var committed = JsonDocument.Parse(first);
            Assert.True(committed.RootElement.GetProperty("success").GetBoolean());
            Assert.False(committed.RootElement.GetProperty("preview").GetBoolean());
            Assert.Equal("tx-commit", committed.RootElement.GetProperty("transaction").GetProperty("transactionId").GetString());
            Assert.Equal(previewId, committed.RootElement.GetProperty("retention").GetProperty("previewId").GetString());
            Assert.Equal(
                preview.RootElement.GetProperty("packageHash").GetString(),
                committed.RootElement.GetProperty("packageHash").GetString());
            Assert.Equal(1, DocxSessionOps.GetVersion(handle));
            Assert.Equal(
                preview.RootElement.GetProperty("packageHash").GetString(),
                DocxSessionOps.GetPackageContentHash(handle));

            // The same id for a different commit request is a conflict, not a second commit.
            var other = JsonSerializer.SerializeToElement(new { handle, previewId = "pv-other" });
            using var conflict = JsonDocument.Parse(
                DocxSessionOps.CommitPreviewTransactional(handle, "tx-commit", other, "pv-other"));
            Assert.Equal("transaction_conflict",
                conflict.RootElement.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());

            // Without a transaction the consumed preview is simply gone.
            using var consumed = JsonDocument.Parse(DocxSessionOps.CommitPreview(handle, previewId));
            Assert.Equal("preview_not_found",
                consumed.RootElement.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void DS760h_ClientComposedRetention_CommitsABareEnvelopeTheClientCompletes()
    {
        var handle = DocxSessionOps.OpenSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), null);
        try
        {
            string anchor;
            using (var projection = JsonDocument.Parse(DocxSessionOps.Project(handle)))
            {
                anchor = projection.RootElement.GetProperty("anchorIndex").EnumerateObject()
                    .First(property => property.Name.StartsWith("p:body:", StringComparison.Ordinal)).Name;
            }

            // The browser client's flow: run against the shadow handle, then retain it.
            var shadow = SessionRegistry.CloneSessionForPreview(handle);
            string previewId;
            try
            {
                Assert.True(JsonDocument.Parse(DocxSessionOps.ReplaceText(shadow, anchor, "Client composed."))
                    .RootElement.GetProperty("success").GetBoolean());
                using var retained = JsonDocument.Parse(DocxSessionOps.RetainPreview(shadow));
                previewId = retained.RootElement.GetProperty("retention").GetProperty("previewId").GetString()!;
                Assert.Empty(retained.RootElement.GetProperty("warnings").EnumerateArray());
            }
            finally
            {
                SessionRegistry.CloseSession(shadow);
            }
            Assert.Throws<ArgumentException>(() => DocxSessionOps.RetainPreview(shadow));

            using var commit = JsonDocument.Parse(DocxSessionOps.CommitPreview(handle, previewId));
            Assert.True(commit.RootElement.GetProperty("success").GetBoolean());
            Assert.Empty(commit.RootElement.GetProperty("steps").EnumerateArray());
            Assert.Equal(0, commit.RootElement.GetProperty("baseVersion").GetInt64());
            Assert.Equal(1, commit.RootElement.GetProperty("resultVersion").GetInt64());
            Assert.Equal(
                DocxSessionOps.GetPackageContentHash(handle),
                commit.RootElement.GetProperty("packageHash").GetString());
            Assert.Contains("Client composed.", DocxSessionOps.Project(handle), StringComparison.Ordinal);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void DS760i_LiveSession_CannotRetainItself()
    {
        using var session = Open();
        Assert.Throws<InvalidOperationException>(
            () => session.RetainPreview(null, new System.Collections.Generic.List<string>()));
    }

    private static readonly MutationBatchPreviewOptions Retain = new() { Retain = true };

    private static MutationBatchStep[] GeneratingSteps(string[] anchors) => new[]
    {
        new MutationBatchStep("docx_edit", "replace_text",
            s => s.ReplaceText(anchors[0], "Predicted tracked replacement.")),
        new MutationBatchStep("docx_create", "insert_paragraph",
            s => s.InsertParagraph(anchors[0], Position.After, "Predicted new paragraph.")),
        new MutationBatchStep("docx_comment", "add",
            s => s.AddComment(anchors[1], null, "Alice", "Predicted comment stamped by the clock.")),
        new MutationBatchStep("docx_create", "insert_footnote",
            s => s.InsertFootnote(anchors[1], 0, "Predicted footnote.")),
    };

    private static string[] CreatedIds(MutationBatchResult result) =>
        result.Steps.SelectMany(step => step.Results).SelectMany(edit => edit.Created)
            .Select(anchor => anchor.Id).ToArray();

    private static string Describe(MutationBatchResult result) =>
        result.Failure is null
            ? "no failure envelope"
            : $"{result.Failure.Index}:{result.Failure.Action}:{result.Failure.Error.Code}:{result.Failure.Error.Message}";

    private static DocxSession Open(DocxSessionSettings? settings = null) =>
        new(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            settings ?? new DocxSessionSettings { PersistAnchorIds = true });

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:", StringComparison.Ordinal))
            .ToArray();
}
