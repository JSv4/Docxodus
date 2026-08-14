#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Document-version and optimistic-mutation regression coverage (issue #447).</summary>
public class DocxSessionPreconditionTests
{
    [Fact]
    public void DS447_Version_AdvancesOnlyForCommittedMutationUndoAndRedo()
    {
        using var session = Open();
        var first = BodyParagraphs(session)[0];

        Assert.Equal(0, session.Version);
        Assert.False(session.ReplaceText("p:body:missing", "x").Success);
        Assert.Equal(0, session.Version);
        Assert.False(session.Undo());

        Assert.True(session.ReplaceText(first, "Changed").Success);
        Assert.Equal(1, session.Version);

        Assert.Equal(0, session.CompactRuns().RunsRemoved);
        Assert.Equal(1, session.Version);

        Assert.True(session.Undo());
        Assert.Equal(2, session.Version);
        Assert.True(session.Redo());
        Assert.Equal(3, session.Version);
        Assert.False(session.Redo());
        Assert.Equal(3, session.Version);
    }

    [Fact]
    public void DS448_StaleVersion_FailsWithCurrentMetadataWithoutMutationOrHistoryLoss()
    {
        using var session = Open();
        var paragraphs = BodyParagraphs(session);
        var first = paragraphs[0];
        var info = Assert.IsType<AnchorInfo>(session.GetAnchorInfo(first));

        Assert.True(session.ReplaceText(paragraphs[1], "Committed change").Success);
        Assert.Equal(1, session.Version);
        var before = session.Save(persistAnchorIds: true);

        var result = session.ExecuteMutation(
            new MutationPreconditions
            {
                ExpectedVersion = 0,
                AnchorId = first,
                ExpectedContentHash = info.ContentHash,
            },
            s => s.ReplaceText(first, "must not apply"));

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.PreconditionFailed, result.Error?.Code);
        var detail = Assert.IsType<PreconditionFailure>(result.Error?.Precondition);
        Assert.Equal("document_version", detail.Condition);
        Assert.Equal(0L, detail.Expected);
        Assert.Equal(1L, detail.Actual);
        Assert.Equal(1, detail.CurrentVersion);
        Assert.True(detail.CurrentTarget?.Exists);
        Assert.Equal(info.ContentHash, detail.CurrentTarget?.ContentHash);
        Assert.Equal(before, session.Save(persistAnchorIds: true));
        Assert.Equal(1, session.Version);

        // The failed guarded call did not add or consume a history entry: this undoes
        // the one earlier committed mutation.
        Assert.True(session.Undo());
        Assert.Contains("Second paragraph.", session.Project().Markdown);
    }

    [Fact]
    public void DS449_AnchorTextHashAndRangeGuards_ReportExpectedAndActual()
    {
        using var session = Open();
        var first = BodyParagraphs(session)[0];
        var info = Assert.IsType<AnchorInfo>(session.GetAnchorInfo(first));

        Assert.Equal("First paragraph.", info.VisibleText);
        Assert.False(string.IsNullOrWhiteSpace(info.ContentHash));
        Assert.Null(session.EvaluatePreconditions(new MutationPreconditions
        {
            AnchorId = first,
            ExpectedContentHash = info.ContentHash,
            ExpectedText = info.VisibleText,
            ExpectedTextRange = new TextRangePrecondition(0, 5, "First"),
            ExpectedKind = info.Kind,
            ExpectedScope = info.Scope,
        }));

        var error = Assert.IsType<EditError>(session.EvaluatePreconditions(
            new MutationPreconditions
            {
                AnchorId = first,
                ExpectedTextRange = new TextRangePrecondition(6, 9, "different"),
            }));
        Assert.Equal(EditErrorCode.PreconditionFailed, error.Code);
        Assert.Equal("anchor_text_range", error.Precondition?.Condition);
        Assert.Equal("paragraph", error.Precondition?.Actual);
        Assert.Equal(0, session.Version);
        Assert.False(session.Undo());
    }

    [Fact]
    public void DS450_KindChangeAndDeletedAnchor_ReturnCurrentLiveState()
    {
        using var session = Open();
        var paragraphs = BodyParagraphs(session);
        var oldParagraphAnchor = paragraphs[0];
        var deletedAnchor = paragraphs[1];

        Assert.True(session.SetParagraphStyle(oldParagraphAnchor, "Heading1").Success);
        var kindError = Assert.IsType<EditError>(session.EvaluatePreconditions(
            new MutationPreconditions { AnchorId = oldParagraphAnchor, ExpectedKind = "p" }));
        Assert.Equal("anchor_kind", kindError.Precondition?.Condition);
        Assert.Equal("h", kindError.Precondition?.Actual);
        Assert.Equal("h", kindError.Precondition?.CurrentTarget?.Kind);

        Assert.True(session.DeleteBlock(deletedAnchor).Success);
        var deletedError = Assert.IsType<EditError>(session.EvaluatePreconditions(
            new MutationPreconditions { AnchorId = deletedAnchor, ExpectedText = "Second paragraph." }));
        Assert.Equal("anchor_exists", deletedError.Precondition?.Condition);
        Assert.Equal(false, deletedError.Precondition?.Actual);
        Assert.False(deletedError.Precondition?.CurrentTarget?.Exists);
    }

    [Fact]
    public void DS451_ExactMatchCount_IsAtomicAndOneUndoUnit()
    {
        using var session = Open();
        var first = BodyParagraphs(session)[0];
        Assert.True(session.ReplaceText(first, "cat cat").Success);
        Assert.Equal(1, session.Version);
        var beforeFailure = session.Save(persistAnchorIds: true);

        var failed = Assert.Single(session.ReplaceTextRange(first, "cat", "dog",
            new ReplaceOptions { ExpectedMatchCount = 1 }));
        Assert.False(failed.Success);
        Assert.Equal(EditErrorCode.PreconditionFailed, failed.Error?.Code);
        Assert.Equal("match_count", failed.Error?.Precondition?.Condition);
        Assert.Equal(1, failed.Error?.Precondition?.Expected);
        Assert.Equal(2, failed.Error?.Precondition?.Actual);
        Assert.Equal(beforeFailure, session.Save(persistAnchorIds: true));
        Assert.Equal(1, session.Version);

        var replaced = session.ReplaceTextRange(first, "cat", "dog",
            new ReplaceOptions
            {
                Preconditions = new MutationPreconditions
                {
                    ExpectedVersion = 1,
                    ExpectedMatchCount = 2,
                },
            });
        Assert.Equal(2, replaced.Count);
        Assert.All(replaced, r => Assert.True(r.Success));
        Assert.Equal(2, session.Version); // one committed operation, not one per match
        Assert.Contains("dog dog", session.Project().Markdown);

        Assert.True(session.Undo());
        Assert.Equal(3, session.Version);
        Assert.Contains("cat cat", session.Project().Markdown);
    }

    [Fact]
    public void DS452_CrossPartAnchor_GuardsUseExactHeaderTextAndHash()
    {
        using var session = Open();
        var body = BodyParagraphs(session)[0];
        var created = session.SetHeaderText(body, HeaderFooterKind.Default, "Confidential");
        Assert.True(created.Success);
        var header = Assert.Single(created.Created, a => a.Scope.StartsWith("hdr")).Id;
        var info = Assert.IsType<AnchorInfo>(session.GetAnchorInfo(header));

        Assert.Equal("Confidential", info.VisibleText);
        var success = session.ExecuteMutation(
            new MutationPreconditions
            {
                ExpectedVersion = session.Version,
                AnchorId = header,
                ExpectedContentHash = info.ContentHash,
                ExpectedText = "Confidential",
                ExpectedScope = info.Scope,
            },
            s => s.ReplaceText(header, "Privileged"));
        Assert.True(success.Success);

        var staleHash = session.ExecuteMutation(
            new MutationPreconditions { AnchorId = header, ExpectedContentHash = info.ContentHash },
            s => s.ReplaceText(header, "must not apply"));
        Assert.False(staleHash.Success);
        Assert.Equal("anchor_content_hash", staleHash.Error?.Precondition?.Condition);
        Assert.Equal("Privileged", staleHash.Error?.Precondition?.CurrentTarget?.VisibleText);
    }

    [Fact]
    public void DS453_PreconditionFailure_JsonCarriesStructuredCurrentState()
    {
        var error = new EditError(EditErrorCode.PreconditionFailed, "stale", "p:body:u")
        {
            Precondition = new PreconditionFailure(
                "document_version", 4L, 5L, 5L,
                new PreconditionTarget
                {
                    Exists = true,
                    AnchorId = "p:body:u",
                    Kind = "p",
                    Scope = "body",
                    ContentHash = "abc",
                    VisibleText = "current",
                }),
        };

        var json = DocxSessionJson.Serialize(new EditResult { Success = false, Error = error });
        Assert.Contains("\"code\":\"precondition_failed\"", json);
        Assert.Contains("\"currentVersion\":5", json);
        Assert.Contains("\"contentHash\":\"abc\"", json);
        Assert.Contains("\"visibleText\":\"current\"", json);
    }

    private static DocxSession Open() =>
        new(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { PersistAnchorIds = true });

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:"))
            .ToArray();
}
