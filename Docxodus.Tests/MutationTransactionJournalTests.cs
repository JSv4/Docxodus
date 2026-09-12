// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #761: the mutation-transaction journal is one core component per live session that
/// every transport drives — the server-side <see cref="MutationTransactions.Run"/> flow the
/// stdio host uses through <see cref="DocxSessionOps.ExecuteBatchTransactional"/>, and the
/// begin/complete/abandon triplet the browser client uses over the bridge.
/// </summary>
[Collection("MCP session registry isolation")]
public sealed class MutationTransactionJournalTests : IDisposable
{
    private readonly int _handle = DocxSessionOps.OpenSession(
        DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), null);

    public void Dispose() => DocxSessionOps.CloseSession(_handle);

    private string FirstAnchor()
    {
        using var document = JsonDocument.Parse(DocxSessionOps.Project(_handle));
        return document.RootElement.GetProperty("anchorIndex").EnumerateObject()
            .First(entry => entry.Value.GetProperty("scope").GetString() == "body"
                && entry.Value.GetProperty("kind").GetString() == "p").Name;
    }

    private static JsonElement Request(int handle, string markdown) =>
        JsonSerializer.SerializeToElement(new
        {
            handle,
            mode = "atomic",
            steps = new[] { new { operation = "insert_paragraph", args = new { markdown } } },
        });

    private IEnumerable<MutationBatchStep> InsertStep(string anchor, string markdown) => new[]
    {
        DocxSessionOps.SerializedBatchStep("test", "insert_paragraph",
            () => DocxSessionOps.InsertParagraph(_handle, anchor, Position.After, markdown)),
    };

    [Fact]
    public void MTX761a_ServerSideRetry_ReplaysTheRetainedResponseWithoutExecutingAgain()
    {
        var anchor = FirstAnchor();
        var request = Request(_handle, "inserted exactly once");

        var first = DocxSessionOps.ExecuteBatchTransactional(
            _handle, "tx-1", request, MutationBatchMode.Atomic, () => InsertStep(anchor, "inserted exactly once"));
        var version = DocxSessionOps.GetVersion(_handle);
        var retry = DocxSessionOps.ExecuteBatchTransactional(
            _handle, "tx-1", request, MutationBatchMode.Atomic,
            () => throw new InvalidOperationException("a replay must not build steps"));

        Assert.Equal(first, retry);
        Assert.Equal(version, DocxSessionOps.GetVersion(_handle));
        using var response = JsonDocument.Parse(first);
        Assert.True(response.RootElement.GetProperty("success").GetBoolean());
        var transaction = response.RootElement.GetProperty("transaction");
        Assert.Equal("tx-1", transaction.GetProperty("transactionId").GetString());
        Assert.StartsWith("sha256:", transaction.GetProperty("requestFingerprint").GetString());
        Assert.Equal(1, CountParagraphs("inserted exactly once"));
    }

    [Fact]
    public void MTX761b_ReusingAnIdForADifferentRequest_IsAConflictThatChangesNothing()
    {
        var anchor = FirstAnchor();
        DocxSessionOps.ExecuteBatchTransactional(
            _handle, "tx-2", Request(_handle, "first"), MutationBatchMode.Atomic, () => InsertStep(anchor, "first"));
        var version = DocxSessionOps.GetVersion(_handle);

        var conflict = DocxSessionOps.ExecuteBatchTransactional(
            _handle, "tx-2", Request(_handle, "second"), MutationBatchMode.Atomic,
            () => throw new InvalidOperationException("a conflict must not build steps"));

        using var response = JsonDocument.Parse(conflict);
        Assert.False(response.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("transaction_conflict",
            response.RootElement.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        Assert.Equal("tx-2", response.RootElement.GetProperty("transaction").GetProperty("transactionId").GetString());
        Assert.Equal(version, DocxSessionOps.GetVersion(_handle));
        Assert.Equal(0, CountParagraphs("second"));
    }

    [Fact]
    public void MTX761c_FingerprintIgnoresSessionAddressingAndCanonicalizesTheDefaultMode()
    {
        var byHandle = MutationTransactions.Fingerprint(JsonSerializer.SerializeToElement(
            new { handle = 7, transactionId = "x", steps = new[] { new { operation = "undo" } } }));
        var bySession = MutationTransactions.Fingerprint(JsonSerializer.SerializeToElement(
            new { sessionId = "s", mode = "atomic", steps = new[] { new { operation = "undo" } } }));
        var bestEffort = MutationTransactions.Fingerprint(JsonSerializer.SerializeToElement(
            new { mode = "best_effort", steps = new[] { new { operation = "undo" } } }));

        Assert.Equal(byHandle, bySession);
        Assert.NotEqual(byHandle, bestEffort);
        Assert.Throws<ArgumentException>(() =>
            MutationTransactions.Fingerprint(JsonSerializer.SerializeToElement("not an object")));
    }

    [Theory]
    [InlineData("", "empty or whitespace")]
    [InlineData(" \u00a0\u3000", "empty or whitespace")]
    [InlineData("\ufeff", null)]
    [InlineData("ok", null)]
    public void MTX761d_TransactionIdValidation_IsOneRuleForEveryTransport(string id, string? expectedFragment)
    {
        var verdict = MutationTransactions.ValidateTransactionId(id);
        if (expectedFragment is null) Assert.Null(verdict);
        else Assert.Contains(expectedFragment, verdict, StringComparison.Ordinal);
        Assert.Contains("256", MutationTransactions.ValidateTransactionId(new string('a', 257)), StringComparison.Ordinal);
    }

    [Fact]
    public void MTX761e_ClientDrivenFlow_ReservesCompletesAndReplaysTheClientsOwnResult()
    {
        var request = "{\"mode\":\"atomic\",\"what\":\"client batch\"}";

        using (var begin = JsonDocument.Parse(DocxSessionOps.BeginMutationTransaction(_handle, "tx-c", request)))
        {
            Assert.Equal("reserved", begin.RootElement.GetProperty("kind").GetString());
            Assert.Equal(JsonValueKind.Null, begin.RootElement.GetProperty("response").ValueKind);
            Assert.Equal("tx-c", begin.RootElement.GetProperty("transaction").GetProperty("transactionId").GetString());
        }
        // While reserved, the same id reports incomplete: the client owns the outcome.
        using (var again = JsonDocument.Parse(DocxSessionOps.BeginMutationTransaction(_handle, "tx-c", request)))
            Assert.Equal("incomplete", again.RootElement.GetProperty("kind").GetString());

        const string clientResult = "{\"success\":true,\"steps\":[],\"transaction\":{\"transactionId\":\"tx-c\"}}";
        DocxSessionOps.CompleteMutationTransaction(_handle, "tx-c", clientResult);

        using (var replay = JsonDocument.Parse(DocxSessionOps.BeginMutationTransaction(_handle, "tx-c", request)))
        {
            Assert.Equal("replay", replay.RootElement.GetProperty("kind").GetString());
            Assert.Equal(clientResult, replay.RootElement.GetProperty("response").GetString());
        }
        using (var conflict = JsonDocument.Parse(DocxSessionOps.BeginMutationTransaction(
                   _handle, "tx-c", "{\"what\":\"a different batch\"}")))
        {
            Assert.Equal("conflict", conflict.RootElement.GetProperty("kind").GetString());
            using var refusal = JsonDocument.Parse(conflict.RootElement.GetProperty("response").GetString()!);
            Assert.Equal("transaction_conflict",
                refusal.RootElement.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        }
    }

    [Fact]
    public void MTX761f_AbandonedReservation_StaysBoundAsIncompleteAndCompletingItAgainFails()
    {
        const string request = "{\"what\":\"abandoned\"}";
        DocxSessionOps.BeginMutationTransaction(_handle, "tx-a", request);
        DocxSessionOps.AbandonMutationTransaction(_handle, "tx-a");

        using var again = JsonDocument.Parse(DocxSessionOps.BeginMutationTransaction(_handle, "tx-a", request));
        Assert.Equal("incomplete", again.RootElement.GetProperty("kind").GetString());
        Assert.Throws<InvalidOperationException>(() =>
            DocxSessionOps.CompleteMutationTransaction(_handle, "tx-a", "{}"));
        DocxSessionOps.AbandonMutationTransaction(_handle, "never-reserved"); // a no-op, not an error
    }

    [Fact]
    public void MTX761g_JournalsAreOnePerSessionAndCloseClearsThem()
    {
        var other = DocxSessionOps.OpenSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), null);
        try
        {
            Assert.NotSame(SessionRegistry.Transactions(_handle), SessionRegistry.Transactions(other));
            Assert.Same(SessionRegistry.Transactions(_handle), SessionRegistry.Transactions(_handle));
            DocxSessionOps.BeginMutationTransaction(other, "tx-o", "{}");
            var journal = SessionRegistry.Transactions(other);
            Assert.NotNull(journal.GetRecord("tx-o"));
            DocxSessionOps.CloseSession(other);
            Assert.Throws<ArgumentException>(() => SessionRegistry.Transactions(other));
        }
        finally
        {
            DocxSessionOps.CloseSession(other);
        }
    }

    [Fact]
    public void MTX761h_PreviewIsNeverTransactional_AndInvalidIdsAreRefusedBeforeAnyWork()
    {
        Assert.Throws<ArgumentException>(() => DocxSessionOps.ExecuteBatchTransactional(
            _handle, "   ", Request(_handle, "x"), MutationBatchMode.Atomic,
            () => throw new InvalidOperationException("must not build steps")));
        Assert.Throws<ArgumentException>(() =>
            DocxSessionOps.BeginMutationTransaction(_handle, "", "{}"));
    }

    private int CountParagraphs(string text)
    {
        using var document = JsonDocument.Parse(DocxSessionOps.Project(_handle));
        return document.RootElement.GetProperty("markdown").GetString()!
            .Split('\n').Count(line => line.Contains(text, StringComparison.Ordinal));
    }
}
