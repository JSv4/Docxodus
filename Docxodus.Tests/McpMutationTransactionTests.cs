#nullable enable

using System.Collections.Concurrent;
using System.Text.Json;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Issue #449: in-session idempotency and session-wide dispatch ordering.</summary>
[Collection("MCP session registry isolation")]
public sealed class McpMutationTransactionTests : IDisposable
{
    private readonly string _root;
    private readonly string _path;
    private readonly SessionStore _store;

    public McpMutationTransactionTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"mcp-transactions-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _path = Path.Combine(_root, "document.docx");
        File.WriteAllBytes(_path, DocxSession.CreateBlankDocxBytes());
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    [Fact]
    public void MCP449_ResponseLossRetryIsByteExactAndDoesNotDisturbUndoRedo()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        var args = MutationArgs(sessionId, "tx-response-loss", anchor, "inserted exactly once");

        // Simulate transport loss by discarding the first returned string.
        var original = Dispatcher.Call(_store, "docxodus_mutations", J(args));
        var parsed = J(original);
        Assert.True(parsed.GetProperty("success").GetBoolean());
        Assert.False(parsed.GetProperty("rolledBack").GetBoolean());
        Assert.Equal("ok", parsed.GetProperty("status").GetString());
        Assert.Equal(0, parsed.GetProperty("baseVersion").GetInt64());
        Assert.Equal(1, parsed.GetProperty("resultVersion").GetInt64());
        Assert.NotEmpty(parsed.GetProperty("packageHash").GetString()!);
        var step = Assert.Single(parsed.GetProperty("steps").EnumerateArray());
        Assert.Equal("docxodus_edit", step.GetProperty("tool").GetString());
        Assert.Equal("insert_paragraph", step.GetProperty("action").GetString());
        var edit = Assert.Single(step.GetProperty("results").EnumerateArray());
        Assert.NotEmpty(edit.GetProperty("created")[0].GetProperty("id").GetString()!);
        var identity = parsed.GetProperty("transaction");
        Assert.Equal(1, identity.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("tx-response-loss", identity.GetProperty("transactionId").GetString());
        Assert.StartsWith("sha256:", identity.GetProperty("requestFingerprint").GetString());
        var retained = Assert.IsType<MutationTransactionRecord>(
            _store.Get(sessionId).MutationTransactions.GetRecord("tx-response-loss"));
        Assert.Equal(original, retained.SerializedResponse);
        var retainedResult = J(retained.SerializedResponse!);
        Assert.Equal(0, retainedResult.GetProperty("baseVersion").GetInt64());
        Assert.Equal(1, retainedResult.GetProperty("resultVersion").GetInt64());
        Assert.True(retainedResult.GetProperty("success").GetBoolean());
        Assert.False(retainedResult.GetProperty("rolledBack").GetBoolean());
        Assert.NotEmpty(retainedResult.GetProperty("packageHash").GetString()!);
        var retainedStep = Assert.Single(retainedResult.GetProperty("steps").EnumerateArray());
        Assert.Equal("docxodus_edit", retainedStep.GetProperty("tool").GetString());
        Assert.Equal("insert_paragraph", retainedStep.GetProperty("action").GetString());
        Assert.NotEmpty(Assert.Single(retainedStep.GetProperty("results").EnumerateArray())
            .GetProperty("created")[0].GetProperty("id").GetString()!);

        Assert.True(J(Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "undo",
        })))).GetProperty("success").GetBoolean());

        var replay = Dispatcher.Call(_store, "docxodus_mutations", J(args));
        Assert.Equal(original, replay);
        Assert.True(J(Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "redo",
        })))).GetProperty("success").GetBoolean());
        var markdown = GetMarkdown(_store, sessionId);
        Assert.Equal(1, Occurrences(markdown, "inserted exactly once"));
    }

    [Fact]
    public void MCP449_OmittedAtomicModeAndExplicitAtomicModeReplayExactly()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        var omitted = MutationArgs(sessionId, "tx-default-mode", anchor, "default mode");
        var explicitAtomic = omitted.Replace(
            "\"transactionId\":\"tx-default-mode\",",
            "\"mode\":\"atomic\",\"transactionId\":\"tx-default-mode\",");

        var first = Dispatcher.Call(_store, "docxodus_mutations", J(omitted));
        var second = Dispatcher.Call(_store, "docxodus_mutations", J(explicitAtomic));

        Assert.Equal(first, second);
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
    }

    [Fact]
    public void MCP449_SameIdDifferentRequestReturnsTypedConflictWithoutMutation()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        var first = Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(sessionId, "tx-conflict", anchor, "first")));
        var conflict = J(Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(sessionId, "tx-conflict", anchor, "second"))));

        Assert.True(J(first).GetProperty("success").GetBoolean());
        Assert.Equal("transaction_conflict",
            conflict.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
        Assert.DoesNotContain("second", GetMarkdown(_store, sessionId));
    }

    [Fact]
    public void MCP449_AtomicRollbackAndBestEffortPartialResultsAreCached()
    {
        var atomicSession = OpenSession(_store, _path);
        var atomicAnchor = FirstAnchor(_store, atomicSession);
        var atomicArgs = JsonSerializer.Serialize(new
        {
            sessionId = atomicSession,
            transactionId = "tx-atomic-failure",
            mode = "atomic",
            steps = new object[]
            {
                Step("replace_text", atomicAnchor, "speculative"),
                Step("replace_text", "p:body:missing", "fail"),
            },
        });
        var atomic = Dispatcher.Call(_store, "docxodus_mutations", J(atomicArgs));
        Assert.Equal(atomic, Dispatcher.Call(_store, "docxodus_mutations", J(atomicArgs)));
        var atomicResult = J(atomic);
        Assert.Equal("failed", atomicResult.GetProperty("status").GetString());
        Assert.True(atomicResult.GetProperty("rolledBack").GetBoolean());
        Assert.Equal(0, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(atomicSession).Handle));

        var partialSession = OpenSession(_store, _path);
        var partialAnchor = FirstAnchor(_store, partialSession);
        var partialArgs = JsonSerializer.Serialize(new
        {
            sessionId = partialSession,
            transactionId = "tx-partial",
            mode = "best_effort",
            steps = new object[]
            {
                Step("replace_text", partialAnchor, "retained partial"),
                Step("replace_text", "p:body:missing", "fail"),
            },
        });
        var partial = Dispatcher.Call(_store, "docxodus_mutations", J(partialArgs));
        Assert.Equal(partial, Dispatcher.Call(_store, "docxodus_mutations", J(partialArgs)));
        Assert.Equal("partial", J(partial).GetProperty("status").GetString());
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(partialSession).Handle));
        Assert.Contains("retained partial", GetMarkdown(_store, partialSession));
    }

    [Fact]
    public void MCP449_PreconditionAndValidationFailuresReplayBeforeCurrentStateChecks()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        var guardedArgs = JsonSerializer.Serialize(new
        {
            sessionId,
            transactionId = "tx-precondition",
            preconditions = new { expectedVersion = 99 },
            steps = new[] { Step("replace_text", anchor, "must not apply") },
        });
        var failed = Dispatcher.Call(_store, "docxodus_mutations", J(guardedArgs));
        Assert.Equal("precondition_failed", J(failed).GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());

        ReplaceDirect(_store, sessionId, anchor, "later state");
        Assert.Equal(failed, Dispatcher.Call(_store, "docxodus_mutations", J(guardedArgs)));
        Assert.Contains("later state", GetMarkdown(_store, sessionId));

        var invalidArgs = $$"""
            {"sessionId":{{JsonSerializer.Serialize(sessionId)}},"transactionId":"tx-validation","mode":"sideways","steps":[]}
            """;
        var invalid = Dispatcher.Call(_store, "docxodus_mutations", J(invalidArgs));
        Assert.Equal("invalid_batch_step", J(invalid).GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());
        Assert.Equal(invalid, Dispatcher.Call(_store, "docxodus_mutations", J(invalidArgs)));
    }

    [Fact]
    public void MCP449_PreviewDryRunDirectToolsAndStepArgsRejectTransactionIds()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        foreach (var previewProperties in new[]
        {
            "\"mode\":\"preview\"",
            "\"mode\":\"atomic\",\"preview\":true",
        })
        {
            var args = "{\"sessionId\":" + JsonSerializer.Serialize(sessionId)
                + ",\"transactionId\":\"tx-preview\"," + previewProperties
                + ",\"steps\":[{\"tool\":\"docxodus_edit\",\"args\":{\"action\":\"replace_text\""
                + ",\"anchorId\":" + JsonSerializer.Serialize(anchor)
                + ",\"markdown\":\"shadow\"}}]}";
            var rejected = J(Dispatcher.Call(_store, "docxodus_mutations", J(args)));
            Assert.Equal("invalid_transaction", rejected.GetProperty("failure")
                .GetProperty("error").GetProperty("code").GetString());
            Assert.False(rejected.TryGetProperty("transaction", out _));
        }
        Assert.Equal(0, _store.Get(sessionId).MutationTransactions.FullRecordCount);

        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_edit",
            J(JsonSerializer.Serialize(new
            {
                sessionId,
                transactionId = "nested-direct",
                action = "replace_text",
                anchorId = anchor,
                markdown = "no",
            }))));

        var nested = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            steps = new[]
            {
                new
                {
                    tool = "docxodus_edit",
                    args = new
                    {
                        transactionId = "nested-step",
                        action = "replace_text",
                        anchorId = anchor,
                        markdown = "no",
                    },
                },
            },
        }))));
        Assert.Equal("invalid_transaction", nested.GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void MCP449_CanonicalFingerprintNormalizesObjectsWhitespaceAndEscapesOnly()
    {
        var left = J("""
            {
              "sessionId": "session-a",
              "transactionId": "tx-a",
              "unknown": { "z": "\u0061", "a": true },
              "steps": [1, { "right": null, "left": "same" }]
            }
            """);
        var right = J("""{"steps":[1,{"left":"same","right":null}],"unknown":{"a":true,"z":"a"},"mode":"atomic","transactionId":"tx-b","sessionId":"session-b"}""");
        Assert.Equal(
            MutationTransactions.Fingerprint(left),
            MutationTransactions.Fingerprint(right));

        var baseline = MutationTransactions.Fingerprint(J("""{"steps":[],"mode":"atomic","unknown":1}"""));
        Assert.NotEqual(baseline,
            MutationTransactions.Fingerprint(J("""{"steps":[],"mode":"atomic","unknown":1.0}""")));
        Assert.NotEqual(baseline,
            MutationTransactions.Fingerprint(J("""{"steps":[],"mode":"atomic","unknown":"1"}""")));
        Assert.NotEqual(
            MutationTransactions.Fingerprint(J("""{"steps":[],"unknown":"Spelling"}""")),
            MutationTransactions.Fingerprint(J("""{"steps":[],"unknown":"spelling"}""")));
        Assert.NotEqual(baseline,
            MutationTransactions.Fingerprint(J("""{"steps":[],"mode":"atomic","unknown":1,"extra":null}""")));
        Assert.NotEqual(
            MutationTransactions.Fingerprint(J("""{"steps":[1,2]}""")),
            MutationTransactions.Fingerprint(J("""{"steps":[2,1]}""")));
        Assert.NotEqual(
            MutationTransactions.Fingerprint(J("""{"steps":[],"mode":"apply"}""")),
            MutationTransactions.Fingerprint(J("""{"steps":[],"mode":"best_effort"}""")));
        Assert.NotEqual(
            MutationTransactions.Fingerprint(J("""{"steps":[],"preview":false}""")),
            MutationTransactions.Fingerprint(J("""{"steps":[]}""")));
    }

    [Fact]
    public void MCP449_DuplicateKeysAtAnyDepthAreRejectedBeforeReservation()
    {
        var sessionId = OpenSession(_store, _path);
        var duplicate = "{\"sessionId\":" + JsonSerializer.Serialize(sessionId)
            + ",\"transactionId\":\"tx-duplicate\",\"steps\":[{\"tool\":\"docxodus_edit\""
            + ",\"args\":{\"action\":\"replace_text\",\"action\":\"replace_text\"}}]}";

        var error = Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_mutations", J(duplicate)));
        Assert.Contains("duplicate JSON property", error.Message, StringComparison.Ordinal);
        Assert.Equal(0, _store.Get(sessionId).MutationTransactions.FullRecordCount);
        Assert.Null(_store.Get(sessionId).MutationTransactions.GetRecord("tx-duplicate"));
    }

    [Fact]
    public void MCP449_JournalUsesGeneratedMetadataAndBoundedFullThenTombstoneFifos()
    {
        var now = new DateTimeOffset(2026, 8, 14, 12, 0, 0, TimeSpan.Zero);
        var recordNumber = 0;
        var journal = new MutationTransactions(
            fullRecordCapacity: 1,
            tombstoneCapacity: 1,
            utcNow: () => now,
            recordIdFactory: () => $"record-{++recordNumber}");

        var a = AssertReserved(journal.Begin("a", "sha256:a"));
        Assert.Equal("record-1", a.RecordId);
        Assert.Equal(now, a.StartedAt);
        now = now.AddSeconds(1);
        var completedA = journal.Complete(a, "{\"a\":1}");
        Assert.Equal(now, completedA.CompletedAt);

        var b = AssertReserved(journal.Begin("b", "sha256:b"));
        now = now.AddSeconds(1);
        journal.Complete(b, "{\"b\":1}");
        Assert.Equal(1, journal.FullRecordCount);
        Assert.Equal(1, journal.TombstoneCount);
        Assert.Equal(MutationTransactionDecisionKind.ResultEvicted,
            journal.Begin("a", "sha256:a").Kind);
        Assert.Equal(MutationTransactionDecisionKind.Conflict,
            journal.Begin("a", "sha256:different").Kind);

        var c = AssertReserved(journal.Begin("c", "sha256:c"));
        now = now.AddSeconds(1);
        journal.Complete(c, "{\"c\":1}");
        Assert.Null(journal.GetTombstone("a"));
        Assert.Equal(MutationTransactionDecisionKind.Reserved,
            journal.Begin("a", "sha256:fresh-after-both-fifos").Kind);
    }

    [Fact]
    public void MCP449_CompletionClockFailureAfterCommitStillReplaysWithoutApplyingAgain()
    {
        var now = new DateTimeOffset(2026, 8, 14, 13, 0, 0, TimeSpan.Zero);
        var clockCalls = 0;
        var journal = new MutationTransactions(
            utcNow: () => ++clockCalls == 2
                ? throw new InvalidOperationException("completion clock failed")
                : now,
            recordIdFactory: () => "completion-record");
        using var store = new TestSessionStore(
            new LocalFileDocumentStore(_root), () => journal);
        var sessionId = OpenSession(store.Value, _path);
        var args = MutationArgs(
            sessionId,
            "completion",
            FirstAnchor(store.Value, sessionId),
            "committed before clock failure");

        var original = Dispatcher.Call(store.Value, "docxodus_mutations", J(args));
        var replay = Dispatcher.Call(store.Value, "docxodus_mutations", J(args));

        Assert.True(J(original).GetProperty("success").GetBoolean());
        Assert.Equal(original, replay);
        Assert.Equal(1,
            Docxodus.Internal.DocxSessionOps.GetVersion(store.Value.Get(sessionId).Handle));
        Assert.Equal(1, Occurrences(
            GetMarkdown(store.Value, sessionId), "committed before clock failure"));
        var completed = Assert.IsType<MutationTransactionRecord>(
            journal.GetRecord("completion"));
        Assert.Equal(now, completed.StartedAt);
        Assert.Equal(now, completed.CompletedAt);
        Assert.Equal(original, completed.SerializedResponse);
        Assert.Null(journal.GetTombstone("completion"));
    }

    [Fact]
    public void MCP449_EvictionClockFailureStillMovesTheExactIdentityToATombstone()
    {
        var start = new DateTimeOffset(2026, 8, 14, 14, 0, 0, TimeSpan.Zero);
        var clockCalls = 0;
        var recordNumber = 0;
        DateTimeOffset Clock()
        {
            clockCalls++;
            if (clockCalls == 5) throw new InvalidOperationException("eviction clock failed");
            return start.AddSeconds(clockCalls - 1);
        }
        var journal = new MutationTransactions(
            fullRecordCapacity: 1,
            tombstoneCapacity: 2,
            utcNow: Clock,
            recordIdFactory: () => $"eviction-record-{++recordNumber}");
        var first = AssertReserved(journal.Begin("first", "sha256:first"));
        journal.Complete(first, "{\"result\":\"first exact\"}");
        var second = AssertReserved(journal.Begin("second", "sha256:second"));

        var completedSecond = journal.Complete(second, "{\"result\":\"second exact\"}");

        Assert.Equal(1, journal.FullRecordCount);
        Assert.Equal(1, journal.TombstoneCount);
        Assert.Null(journal.GetRecord("first"));
        var tombstone = Assert.IsType<MutationTransactionTombstone>(
            journal.GetTombstone("first"));
        Assert.Equal(first.RecordId, tombstone.RecordId);
        Assert.Equal(first.Identity, tombstone.Identity);
        Assert.Equal(first.StartedAt, tombstone.StartedAt);
        Assert.Equal(start.AddSeconds(1), tombstone.CompletedAt);
        Assert.Equal(tombstone.CompletedAt, tombstone.EvictedAt);
        Assert.Equal(MutationTransactionDecisionKind.ResultEvicted,
            journal.Begin("first", "sha256:first").Kind);
        Assert.Equal(MutationTransactionDecisionKind.Conflict,
            journal.Begin("first", "sha256:different").Kind);
        var replay = journal.Begin("second", "sha256:second");
        Assert.Equal(MutationTransactionDecisionKind.Replay, replay.Kind);
        Assert.Same(completedSecond, replay.Record);
        Assert.Equal("{\"result\":\"second exact\"}", replay.SerializedResponse);
    }

    [Fact]
    public void MCP449_AttachIdentitySerializesTheIdentitySchemaVersion()
    {
        var serialized = MutationTransactions.AttachIdentity(
            "{\"success\":true}\n",
            new MutationTransactionIdentity(7, "tx-version", "sha256:version"));

        Assert.EndsWith("\n", serialized, StringComparison.Ordinal);
        var transaction = J(serialized).GetProperty("transaction");
        Assert.Equal(7, transaction.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("tx-version", transaction.GetProperty("transactionId").GetString());
        Assert.Equal("sha256:version",
            transaction.GetProperty("requestFingerprint").GetString());
    }

    [Fact]
    public void MCP449_DispatcherSerializesEvictedResultAndConflictAndReusesOnlyAfterTombstone()
    {
        using var store = new TestSessionStore(
            new LocalFileDocumentStore(_root),
            () => new MutationTransactions(fullRecordCapacity: 1, tombstoneCapacity: 1));
        var sessionId = OpenSession(store.Value, _path);
        var anchor = FirstAnchor(store.Value, sessionId);

        string Args(string id, string markdown) => JsonSerializer.Serialize(new
        {
            sessionId,
            transactionId = id,
            steps = new[] { Step("replace_text", anchor, markdown) },
        });

        var firstA = Dispatcher.Call(store.Value, "docxodus_mutations", J(Args("a", "a1")));
        Dispatcher.Call(store.Value, "docxodus_mutations", J(Args("b", "b1")));

        var evicted = J(Dispatcher.Call(
            store.Value, "docxodus_mutations", J(Args("a", "a1"))));
        Assert.Equal("transaction_result_evicted", evicted.GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());
        var conflict = J(Dispatcher.Call(
            store.Value, "docxodus_mutations", J(Args("a", "different"))));
        Assert.Equal("transaction_conflict", conflict.GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());
        Assert.Equal(2, Docxodus.Internal.DocxSessionOps.GetVersion(store.Value.Get(sessionId).Handle));

        Dispatcher.Call(store.Value, "docxodus_mutations", J(Args("c", "c1")));
        var reused = Dispatcher.Call(
            store.Value, "docxodus_mutations", J(Args("a", "fresh after tombstone")));
        Assert.NotEqual(firstA, reused);
        Assert.True(J(reused).GetProperty("success").GetBoolean());
        Assert.Equal(4, Docxodus.Internal.DocxSessionOps.GetVersion(store.Value.Get(sessionId).Handle));
    }

    [Fact]
    public void MCP449_UnexpectedPostReservationFailureIsStructuredAndCached()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        var args = MutationArgs(sessionId, "tx-unexpected", anchor, "cannot execute");

        // Invalidate only the lower-level handle while leaving the MCP session registered. This
        // forces an unexpected registry failure after transaction reservation.
        Docxodus.Internal.DocxSessionOps.CloseSession(_store.Get(sessionId).Handle);
        var first = Dispatcher.Call(_store, "docxodus_mutations", J(args));
        var failure = J(first);
        Assert.Equal("internal_error", failure.GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());
        Assert.Equal(first, Dispatcher.Call(_store, "docxodus_mutations", J(args)));
        Assert.Equal(1, _store.Get(sessionId).MutationTransactions.FullRecordCount);
    }

    [Fact]
    public void MCP449_GeneratedRevisionTimestampAndIdentityReplayExactly()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        ReplaceDirect(_store, sessionId, anchor, "revision target");
        Dispatcher.Call(_store, "docxodus_track_changes", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "set_mode",
            mode = "render_inline",
            revisionAuthor = "Reviewer",
        })));
        var args = JsonSerializer.Serialize(new
        {
            sessionId,
            transactionId = "tx-generated-revision",
            steps = new[]
            {
                new
                {
                    tool = "docxodus_edit",
                    args = new
                    {
                        action = "replace_text",
                        anchorId = anchor,
                        markdown = "Generated metadata",
                    },
                },
            },
        });

        var original = Dispatcher.Call(_store, "docxodus_mutations", J(args));
        var revisions = J(original).GetProperty("revisionChanges")
            .GetProperty("added").EnumerateArray().ToArray();
        Assert.NotEmpty(revisions);
        Assert.All(revisions, revision =>
        {
            Assert.NotEmpty(revision.GetProperty("id").GetString()!);
            Assert.False(string.IsNullOrWhiteSpace(revision.GetProperty("date").GetString()));
        });
        Assert.Equal(original, Dispatcher.Call(_store, "docxodus_mutations", J(args)));
    }

    [Fact]
    public void MCP449_TransactionIdsAreSessionScopedAndCloseClearsTheirLifecycle()
    {
        var firstSession = OpenSession(_store, _path);
        var secondSession = OpenSession(_store, _path);
        var firstResult = Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(firstSession, "same-id", FirstAnchor(_store, firstSession), "first session")));
        var secondResult = Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(secondSession, "same-id", FirstAnchor(_store, secondSession), "second session")));
        Assert.True(J(firstResult).GetProperty("success").GetBoolean());
        Assert.True(J(secondResult).GetProperty("success").GetBoolean());

        _store.Close(firstSession);
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(firstSession, "same-id", "p:body:any", "retry after close"))));
        Assert.Contains("second session", GetMarkdown(_store, secondSession));
    }

    [Fact]
    public void MCP449_ConcurrentIdenticalCallsSerializeToOneMutationAndOneExactReplay()
    {
        var sessionId = OpenSession(_store, _path);
        var args = J(MutationArgs(
            sessionId, "tx-concurrent", FirstAnchor(_store, sessionId), "concurrent once"));
        var start = new ManualResetEventSlim(false);
        var results = new ConcurrentBag<string>();
        var calls = Enumerable.Range(0, 8).Select(_ => Task.Run(() =>
        {
            start.Wait();
            results.Add(Dispatcher.Call(_store, "docxodus_mutations", args));
        })).ToArray();

        start.Set();
        Task.WaitAll(calls);
        Assert.Equal(8, results.Count);
        Assert.Single(results.Distinct(StringComparer.Ordinal));
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
        Assert.Equal(1, Occurrences(GetMarkdown(_store, sessionId), "concurrent once"));
    }

    [Fact]
    public void MCP449_SaveCloseAndCloseAllWaitForTheSameSessionDispatchGate()
    {
        var blockingStore = new BlockingDocumentStore(DocxSession.CreateBlankDocxBytes());
        var store = new SessionStore(blockingStore);
        try
        {
            var sessionId = OpenSession(store, "document.docx");
            var save = Task.Run(() => Dispatcher.Call(store, "docxodus_save",
                J(JsonSerializer.Serialize(new { sessionId }))));
            Assert.True(blockingStore.WriteEntered.Wait(TimeSpan.FromSeconds(5)));
            var close = Task.Run(() => Dispatcher.Call(store, "docxodus_close",
                J(JsonSerializer.Serialize(new { sessionId }))));
            Assert.False(close.Wait(TimeSpan.FromMilliseconds(100)));
            blockingStore.ReleaseWrite.Set();
            Assert.True(save.Wait(TimeSpan.FromSeconds(5)));
            Assert.True(close.Wait(TimeSpan.FromSeconds(5)));

            var closeAllSession = OpenSession(store, "document.docx");
            var entered = new ManualResetEventSlim(false);
            var release = new ManualResetEventSlim(false);
            var action = Task.Run(() => store.Dispatch(closeAllSession, () =>
            {
                entered.Set();
                release.Wait();
                return "{}";
            }));
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            var closeAll = Task.Run(store.CloseAll);
            Assert.False(closeAll.Wait(TimeSpan.FromMilliseconds(100)));
            release.Set();
            Assert.True(action.Wait(TimeSpan.FromSeconds(5)));
            Assert.True(closeAll.Wait(TimeSpan.FromSeconds(5)));
            Assert.Throws<McpToolException>(() => store.Get(closeAllSession));
        }
        finally
        {
            blockingStore.ReleaseWrite.Set();
            store.CloseAll();
        }
    }

    [Fact]
    public void MCP449_ReentrantDispatchAndLifecycleCallsFailBeforeLocksWhileExternalCloseWaits()
    {
        using var store = new TestSessionStore(new LocalFileDocumentStore(_root));
        var session = store.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings());
        var otherSession = store.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings());
        using var entered = new ManualResetEventSlim(false);
        using var release = new ManualResetEventSlim(false);
        var errors = new ConcurrentBag<Exception?>();
        var action = Task.Run(() => store.Value.Dispatch(session.Id, () =>
        {
            errors.Add(Record.Exception(() => store.Value.Open(
                DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings())));
            errors.Add(Record.Exception(() => store.Value.Close(session.Id)));
            errors.Add(Record.Exception(store.Value.CloseAll));
            errors.Add(Record.Exception(() =>
                store.Value.Dispatch(session.Id, () => "nested same session")));
            errors.Add(Record.Exception(() =>
                store.Value.Dispatch(otherSession.Id, () => "nested cross session")));
            entered.Set();
            release.Wait();
            return "{}";
        }));
        Task? close = null;
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            Assert.Equal(5, errors.Count);
            Assert.All(errors, error =>
            {
                var typed = Assert.IsType<McpToolException>(error);
                Assert.Contains("session dispatch callback", typed.Message,
                    StringComparison.Ordinal);
            });
            close = Task.Run(() => store.Value.Close(session.Id));
            Assert.False(close.Wait(TimeSpan.FromMilliseconds(100)));

            release.Set();
            Assert.True(action.Wait(TimeSpan.FromSeconds(5)));
            Assert.True(close.Wait(TimeSpan.FromSeconds(5)));
            Assert.Throws<McpToolException>(() => store.Value.Get(session.Id));
        }
        finally
        {
            release.Set();
            action.Wait(TimeSpan.FromSeconds(5));
            close?.Wait(TimeSpan.FromSeconds(5));
        }
    }

    [Fact]
    public void MCP449_DifferentSessionsStillDispatchInParallel()
    {
        using var store = new TestSessionStore(new LocalFileDocumentStore(_root));
        var first = store.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings());
        var second = store.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings());
        using var firstEntered = new ManualResetEventSlim(false);
        using var secondEntered = new ManualResetEventSlim(false);
        using var release = new ManualResetEventSlim(false);
        var firstAction = Task.Run(() => store.Value.Dispatch(first.Id, () =>
        {
            firstEntered.Set();
            release.Wait();
            return "first";
        }));
        var secondAction = Task.Run(() => store.Value.Dispatch(second.Id, () =>
        {
            secondEntered.Set();
            release.Wait();
            return "second";
        }));
        try
        {
            Assert.True(firstEntered.Wait(TimeSpan.FromSeconds(5)));
            Assert.True(secondEntered.Wait(TimeSpan.FromSeconds(5)));
        }
        finally
        {
            release.Set();
        }
        Assert.True(Task.WaitAll(new[] { firstAction, secondAction }, TimeSpan.FromSeconds(5)));
        Assert.Equal("first", firstAction.Result);
        Assert.Equal("second", secondAction.Result);
    }

    [Fact]
    public void MCP449_FlowedChildDispatchRejectsWhileActiveButRunsAfterParentReturns()
    {
        using var parentStore = new TestSessionStore(new LocalFileDocumentStore(_root));
        using var childStore = new TestSessionStore(new LocalFileDocumentStore(_root));
        var parentSession = parentStore.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings());
        var childSession = childStore.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings());
        using var releaseDeferredChild = new ManualResetEventSlim(false);
        Task<Exception?>? concurrentChild = null;
        Task<string>? deferredChild = null;

        var parent = Task.Run(() => parentStore.Value.Dispatch(parentSession.Id, () =>
        {
            concurrentChild = Task.Run(() => Record.Exception(() =>
                childStore.Value.Dispatch(childSession.Id, () => "must reject")));
            if (!concurrentChild.Wait(TimeSpan.FromSeconds(5)))
                throw new TimeoutException("concurrent child did not finish");

            deferredChild = Task.Run(() =>
            {
                releaseDeferredChild.Wait();
                return parentStore.Value.Dispatch(parentSession.Id, () => "deferred child");
            });
            return "parent";
        }));

        try
        {
            Assert.True(parent.Wait(TimeSpan.FromSeconds(5)));
            Assert.Equal("parent", parent.Result);
            var reentrant = Assert.IsType<McpToolException>(concurrentChild!.Result);
            Assert.Contains("session dispatch callback", reentrant.Message,
                StringComparison.Ordinal);

            releaseDeferredChild.Set();
            Assert.True(deferredChild!.Wait(TimeSpan.FromSeconds(5)));
            Assert.Equal("deferred child", deferredChild.Result);
        }
        finally
        {
            releaseDeferredChild.Set();
            deferredChild?.Wait(TimeSpan.FromSeconds(5));
        }
    }

    [Fact]
    public void MCP449_JournalFactoryFailureCannotLeakACoreSession()
    {
        var countBefore = Docxodus.Internal.SessionRegistry.Count;
        using var store = new TestSessionStore(
            new LocalFileDocumentStore(_root),
            () => throw new InvalidOperationException("journal factory failed"));

        var error = Assert.Throws<InvalidOperationException>(() => store.Value.Open(
            DocxSession.CreateBlankDocxBytes(), null, new DocxSessionSettings()));

        Assert.Equal("journal factory failed", error.Message);
        Assert.Equal(countBefore, Docxodus.Internal.SessionRegistry.Count);
    }

    [Fact]
    public void MCP449_TransactionIdRuntimeUsesBlankAndUnicodeScalarContract()
    {
        var sessionId = OpenSession(_store, _path);
        var anchor = FirstAnchor(_store, sessionId);
        foreach (var invalid in new[] { "", " \t\r\n", "\u0085" })
        {
            var error = Assert.Throws<McpToolException>(() => Dispatcher.Call(
                _store,
                "docxodus_mutations",
                J(MutationArgs(sessionId, invalid, anchor, "must not execute"))));
            Assert.Contains("empty or whitespace", error.Message, StringComparison.Ordinal);
        }

        const string byteOrderMark = "\uFEFF";
        var byteOrderMarkResult = J(Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(sessionId, byteOrderMark, anchor, "BOM is not whitespace"))));
        Assert.True(byteOrderMarkResult.GetProperty("success").GetBoolean());
        Assert.NotNull(_store.Get(sessionId).MutationTransactions.GetRecord(byteOrderMark));

        var ascii256 = new string('a', 256);
        var asciiResult = J(Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(sessionId, ascii256, anchor, "ascii boundary"))));
        Assert.True(asciiResult.GetProperty("success").GetBoolean());
        Assert.NotNull(_store.Get(sessionId).MutationTransactions.GetRecord(ascii256));

        var asciiError = Assert.Throws<McpToolException>(() => Dispatcher.Call(
            _store,
            "docxodus_mutations",
            J(MutationArgs(sessionId, new string('a', 257), anchor, "too long"))));
        Assert.Contains("Unicode scalar values", asciiError.Message, StringComparison.Ordinal);

        var emoji256 = string.Concat(Enumerable.Repeat("\U0001F600", 256));
        Assert.Equal(512, emoji256.Length);
        var emojiResult = J(Dispatcher.Call(_store, "docxodus_mutations",
            J(MutationArgs(sessionId, emoji256, anchor, "emoji boundary"))));
        Assert.True(emojiResult.GetProperty("success").GetBoolean());
        Assert.NotNull(_store.Get(sessionId).MutationTransactions.GetRecord(emoji256));

        var emoji257 = emoji256 + "\U0001F600";
        var emojiError = Assert.Throws<McpToolException>(() => Dispatcher.Call(
            _store,
            "docxodus_mutations",
            J(MutationArgs(sessionId, emoji257, anchor, "too many emoji"))));
        Assert.Contains("Unicode scalar values", emojiError.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MCP449_SaveThenRetryPreservesTheSavedMutationWithoutApplyingAgain()
    {
        var sessionId = OpenSession(_store, _path);
        var args = MutationArgs(
            sessionId, "tx-save", FirstAnchor(_store, sessionId), "saved transaction");
        var original = Dispatcher.Call(_store, "docxodus_mutations", J(args));
        Dispatcher.Call(_store, "docxodus_save", J(JsonSerializer.Serialize(new { sessionId })));
        Assert.Equal(original, Dispatcher.Call(_store, "docxodus_mutations", J(args)));
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));

        Dispatcher.Call(_store, "docxodus_close", J(JsonSerializer.Serialize(new { sessionId })));
        var reopened = OpenSession(_store, _path);
        Assert.Contains("saved transaction", GetMarkdown(_store, reopened));
        var reused = Dispatcher.Call(_store, "docxodus_mutations", J(MutationArgs(
            reopened, "tx-save", FirstAnchor(_store, reopened), "fresh identity after reopen")));
        Assert.True(J(reused).GetProperty("success").GetBoolean());
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(reopened).Handle));
        Assert.Contains("fresh identity after reopen", GetMarkdown(_store, reopened));
    }

    [Fact]
    public void MCP449_ToolSchemaDocumentsBoundedApplyingTransactionIdentity()
    {
        var tool = Assert.Single(ToolCatalog.Tools,
            candidate => candidate.Name == "docxodus_mutations");
        var schema = J(tool.InputSchemaJson);
        var transactionId = schema.GetProperty("properties").GetProperty("transactionId");
        Assert.Equal("string", transactionId.GetProperty("type").GetString());
        Assert.Equal(1, transactionId.GetProperty("minLength").GetInt32());
        Assert.Equal(MutationTransactions.MaxTransactionIdLength,
            transactionId.GetProperty("maxLength").GetInt32());
        Assert.Equal(
            @"[^\u0009-\u000D\u0020\u0085\u00A0\u1680\u2000-\u200A\u2028-\u2029\u202F\u205F\u3000]",
            transactionId.GetProperty("pattern").GetString());
        Assert.Contains("APPLYING", transactionId.GetProperty("description").GetString(),
            StringComparison.Ordinal);
        Assert.Contains("non-blank", transactionId.GetProperty("description").GetString(),
            StringComparison.Ordinal);
        Assert.Contains("Unicode scalar values",
            transactionId.GetProperty("description").GetString(), StringComparison.Ordinal);
        Assert.Contains(MutationTransactions.TransactionIdWhiteSpaceDescription,
            transactionId.GetProperty("description").GetString(), StringComparison.Ordinal);
        Assert.Contains("U+FEFF is non-whitespace",
            transactionId.GetProperty("description").GetString(), StringComparison.Ordinal);
        Assert.Equal(128, MutationTransactions.DefaultFullRecordCapacity);
        Assert.Equal(1024, MutationTransactions.DefaultTombstoneCapacity);
        Assert.Equal(32L * 1024 * 1024, MutationTransactions.DefaultResponseByteBudget);
    }

    [Fact]
    public void MCP449_RetainedResponsesAreBoundedByBytesNotOnlyByCount()
    {
        // The count cap is deliberately generous here so only the byte budget can evict.
        var journal = new MutationTransactions(
            fullRecordCapacity: 64, tombstoneCapacity: 8, responseByteBudget: 200);
        string Response(char fill) => "{\"r\":\"" + new string(fill, 40) + "\"}";
        Assert.Equal(96, (long)Response('a').Length * sizeof(char));

        journal.Complete(AssertReserved(journal.Begin("a", "sha256:a")), Response('a'));
        journal.Complete(AssertReserved(journal.Begin("b", "sha256:b")), Response('b'));
        Assert.Equal(192, journal.RetainedResponseBytes);
        Assert.Equal(2, journal.FullRecordCount);
        Assert.Equal(0, journal.TombstoneCount);

        // 288 bytes would exceed the 200-byte budget, so the oldest retained response goes even
        // though the count cap is nowhere near reached.
        journal.Complete(AssertReserved(journal.Begin("c", "sha256:c")), Response('c'));
        Assert.Equal(192, journal.RetainedResponseBytes);
        Assert.Equal(2, journal.FullRecordCount);
        Assert.Equal(1, journal.TombstoneCount);
        Assert.Null(journal.GetRecord("a"));
        Assert.Equal(MutationTransactionDecisionKind.ResultEvicted,
            journal.Begin("a", "sha256:a").Kind);
        Assert.Equal(MutationTransactionDecisionKind.Replay, journal.Begin("c", "sha256:c").Kind);

        // A single response larger than the whole budget evicts itself rather than raising the
        // ceiling: the identity stays bound and answers "evicted", and the bound stays a bound.
        journal.Complete(
            AssertReserved(journal.Begin("huge", "sha256:huge")), new string('x', 200));
        Assert.Equal(0, journal.RetainedResponseBytes);
        Assert.Equal(0, journal.FullRecordCount);
        Assert.Equal(MutationTransactionDecisionKind.ResultEvicted,
            journal.Begin("huge", "sha256:huge").Kind);
    }

    [Fact]
    public void MCP449_RetainedResponseCostIsMeasuredFromRealBatchResults()
    {
        var sessionId = OpenSession(_store, _path);
        var journal = _store.Get(sessionId).MutationTransactions;
        var response = Dispatcher.Call(_store, "docxodus_mutations", J(MutationArgs(
            sessionId, "tx-cost", FirstAnchor(_store, sessionId), "one small step")));

        Assert.Equal((long)response.Length * sizeof(char), journal.RetainedResponseBytes);
        // Anchors the per-session memory cost documented in tools/mcp-server/README.md: even a
        // one-step batch retains kilobytes, because a MutationBatchResult carries every step's
        // results twice plus the markdown patch and the semantic delta sets.
        Assert.InRange(journal.RetainedResponseBytes, 1_024L, 128L * 1024);
    }

    [Fact]
    public void MCP449_IncompleteReservationIsReportedTruthfullyAndIsEvictable()
    {
        var journal = new MutationTransactions(fullRecordCapacity: 4, tombstoneCapacity: 4);
        var reservation = AssertReserved(journal.Begin("stranded", "sha256:original"));

        // The fingerprints match, so reporting a conflict would state the opposite of the truth.
        var identical = journal.Begin("stranded", "sha256:original");
        Assert.Equal(MutationTransactionDecisionKind.Incomplete, identical.Kind);
        Assert.Equal("sha256:original", identical.ExistingIdentity!.RequestFingerprint);

        // A genuinely different request still conflicts, against the ORIGINAL fingerprint.
        var different = journal.Begin("stranded", "sha256:other");
        Assert.Equal(MutationTransactionDecisionKind.Conflict, different.Kind);
        Assert.Equal("sha256:original", different.ExistingIdentity!.RequestFingerprint);

        journal.Abandon(reservation);
        Assert.Null(journal.GetRecord("stranded"));
        var tombstone = Assert.IsType<MutationTransactionTombstone>(
            journal.GetTombstone("stranded"));
        Assert.Null(tombstone.CompletedAt);
        Assert.Equal(MutationTransactionDecisionKind.Incomplete,
            journal.Begin("stranded", "sha256:original").Kind);
        Assert.Equal(MutationTransactionDecisionKind.Conflict,
            journal.Begin("stranded", "sha256:other").Kind);

        journal.Abandon(reservation); // idempotent — no second tombstone, no resurrection
        Assert.Equal(1, journal.TombstoneCount);
        Assert.Throws<InvalidOperationException>(() => journal.Complete(reservation, "{}"));
    }

    [Fact]
    public void MCP449_UncompletedReservationsAreBoundedByTheirOwnFifo()
    {
        var journal = new MutationTransactions(fullRecordCapacity: 2, tombstoneCapacity: 8);
        AssertReserved(journal.Begin("r1", "sha256:r1"));
        AssertReserved(journal.Begin("r2", "sha256:r2"));
        AssertReserved(journal.Begin("r3", "sha256:r3"));

        Assert.Null(journal.GetRecord("r1"));
        Assert.Null(Assert.IsType<MutationTransactionTombstone>(
            journal.GetTombstone("r1")).CompletedAt);
        Assert.Equal(MutationTransactionDecisionKind.Incomplete,
            journal.Begin("r1", "sha256:r1").Kind);
        Assert.NotNull(journal.GetRecord("r2"));
        Assert.NotNull(journal.GetRecord("r3"));
    }

    [Fact]
    public void MCP449_DispatcherReportsAnIncompleteTransactionInsteadOfALyingConflict()
    {
        var sessionId = OpenSession(_store, _path);
        var args = MutationArgs(
            sessionId, "tx-stranded", FirstAnchor(_store, sessionId), "must not execute");

        // Strand a reservation the way only a direct component caller can, then retry identically.
        AssertReserved(_store.Get(sessionId).MutationTransactions.Begin(
            "tx-stranded", MutationTransactions.Fingerprint(J(args))));

        var error = J(Dispatcher.Call(_store, "docxodus_mutations", J(args)))
            .GetProperty("failure").GetProperty("error");
        Assert.Equal("transaction_incomplete", error.GetProperty("code").GetString());
        Assert.DoesNotContain("different request fingerprint",
            error.GetProperty("message").GetString()!, StringComparison.Ordinal);
        Assert.Equal(0, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
        Assert.Equal(0, Occurrences(GetMarkdown(_store, sessionId), "must not execute"));
    }

    private static MutationTransactionRecord AssertReserved(MutationTransactionDecision decision)
    {
        Assert.Equal(MutationTransactionDecisionKind.Reserved, decision.Kind);
        return Assert.IsType<MutationTransactionRecord>(decision.Record);
    }

    private static object Step(string action, string anchorId, string markdown) => new
    {
        tool = "docxodus_edit",
        args = new { action, anchorId, markdown },
    };

    private static string MutationArgs(
        string sessionId,
        string transactionId,
        string anchorId,
        string markdown) => JsonSerializer.Serialize(new
    {
        sessionId,
        transactionId,
        steps = new[]
        {
            new
            {
                tool = "docxodus_edit",
                args = new
                {
                    action = "insert_paragraph",
                    anchorId,
                    position = "after",
                    markdown,
                },
            },
        },
    });

    private static string OpenSession(SessionStore store, string path)
    {
        var opened = J(Dispatcher.Call(store, "docxodus_open",
            J(JsonSerializer.Serialize(new { path }))));
        return opened.GetProperty("sessionId").GetString()!;
    }

    private static string FirstAnchor(SessionStore store, string sessionId)
    {
        var content = J(Dispatcher.Call(store, "docxodus_get_content",
            J(JsonSerializer.Serialize(new { sessionId, format = "markdown" }))));
        return content.GetProperty("anchorIndex").EnumerateObject().First().Name;
    }

    private static string GetMarkdown(SessionStore store, string sessionId) =>
        J(Dispatcher.Call(store, "docxodus_get_content",
            J(JsonSerializer.Serialize(new { sessionId, format = "markdown" }))))
        .GetProperty("markdown").GetString()!;

    private static void ReplaceDirect(
        SessionStore store,
        string sessionId,
        string anchorId,
        string markdown)
    {
        var result = J(Dispatcher.Call(store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "replace_text",
            anchorId,
            markdown,
        }))));
        Assert.True(result.GetProperty("success").GetBoolean());
    }

    private static int Occurrences(string value, string search)
    {
        var count = 0;
        var index = 0;
        while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += search.Length;
        }
        return count;
    }

    private static JsonElement J(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }

    private sealed class BlockingDocumentStore : IDocumentStore
    {
        private readonly byte[] _bytes;

        public BlockingDocumentStore(byte[] bytes) => _bytes = bytes;

        public string Kind => "blocking-test";
        public string RootDescription => "blocking-test";
        public ManualResetEventSlim WriteEntered { get; } = new(false);
        public ManualResetEventSlim ReleaseWrite { get; } = new(false);
        public string Resolve(string location) => location;
        public byte[] Read(string resolvedLocation) => _bytes.ToArray();

        public void Write(string resolvedLocation, byte[] bytes)
        {
            WriteEntered.Set();
            ReleaseWrite.Wait();
        }
    }

    private sealed class TestSessionStore : IDisposable
    {
        public TestSessionStore(
            IDocumentStore documents,
            Func<MutationTransactions>? journalFactory = null) =>
            Value = new SessionStore(documents, journalFactory);

        public SessionStore Value { get; }

        public void Dispose() => Value.CloseAll();
    }
}

[CollectionDefinition("MCP session registry isolation", DisableParallelization = true)]
public sealed class McpSessionRegistryIsolationCollection;
