// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.History;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HistoryClientOpsTests
{
    [Fact]
    public async Task RequestIdsReachSharedCreateRestoreReceiptsAndOldCallsStayCompatible()
    {
        var blobs = new MemoryHistoryBlobStore(); var heads = new MemoryHistoryHeadStore();
        var client = new HistoryClientOps(blobs, heads);
        var bytes = DocxVersionRequestTests.Document("client input");
        var request = new HistoryClientRequest { SchemaVersion = 1, DocumentId = "doc", Operation = "create",
            RequestId = "create-id", Metadata = DocxVersionRequestTests.Metadata("create") };
        async Task<HistoryClientResult> Invoke(HistoryClientRequest value, byte[]? candidate = null) =>
            HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync(HistoryClientJson.Write(value), candidate));
        var first = (await Invoke(request, bytes)).View!;
        var restore = request with { Operation = "restore", RequestId = "restore-id", ExpectedHead = first.Head, VersionId = first.Version.Id };
        var restored = (await Invoke(restore)).View!;
        var legacy = request with { RequestId = null, ExpectedHead = restored.Head };
        var later = (await Invoke(legacy, bytes)).View!;
        client = new HistoryClientOps(blobs, heads);
        Assert.Equal(first.Head, (await Invoke(request, bytes)).View!.Head);
        Assert.Equal(restored.Head, (await Invoke(restore)).View!.Head);
        Assert.Equal(later.Head, (await Call(client, "read")).View!.Head);
        Assert.Equal("create-id", first.State.Requests!.Current!.Id);
        Assert.Equal("restore-id", restored.State.Requests!.Current!.Id);
        Assert.Null(later.State.Requests!.Current); Assert.NotNull(later.State.Requests.Index);
        Assert.Equal("RequestConflict", (await Invoke(request with { ExpectedHead = later.Head }, bytes)).ErrorCode);
        Assert.Equal("InvalidRequest", (await Invoke(request with { RequestId = " " }, bytes)).ErrorCode);
        Assert.DoesNotContain("requestId", HistoryClientJson.Write(legacy));
    }

    [Fact]
    public async Task OmittedOptionalFieldsUseWireDefaultsWithGeneratedMetadata()
    {
        const string request = """
            {"schemaVersion":1,"documentId":"doc","operation":"create",
             "metadata":{"author":"actor","createdAt":"2026-01-01T00:00:00Z"}}
            """;
        var client = new HistoryClientOps(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var parsed = HistoryClientJson.Read<HistoryClientRequest>(request);
        Assert.Equal(25, parsed.Limit);
        Assert.Equal(10_000, parsed.MaxEntriesToScan);
        Assert.Empty(parsed.Metadata!.ApplicationMetadata);
        var created = HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync(request, DocxSession.CreateBlankDocxBytes()));
        Assert.True(created.Success, created.Message);
        var listed = HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync(
            """{"schemaVersion":1,"documentId":"doc","operation":"list"}"""));
        Assert.Single(listed.Page!.Versions);
        var invalid = HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync(
            """{"schemaVersion":1,"documentId":"doc","operation":"list","limit":0}"""));
        Assert.Equal("InvalidRequest", invalid.ErrorCode);
    }

    [Fact]
    public async Task SharedClientBoundaryPublishesExportsReplaysAndRestores()
    {
        var client = new HistoryClientOps(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = DocxSession.CreateBlankDocxBytes();
        var metadata = new DocxVersionMetadata { Author = "actor", CreatedAt = DateTimeOffset.Parse("2026-01-01T12:00:00Z") };
        var first = await Call(client, "create", bytes, metadata: metadata);
        Assert.True(first.Success);
        Assert.Equal(1, first.View!.Head.Revision);
        var exported = await Call(client, "export", version: first.View.Version.Id);
        Assert.Equal(bytes, exported.Bytes);
        var listed = await Call(client, "list");
        Assert.Equal(first.View.Version.Id, Assert.Single(listed.Page!.Versions).Id);
        var got = await Call(client, "get", version: first.View.Version.Id);
        Assert.Equal("actor", got.Version!.Record.Metadata.Author);
        Assert.Equal(bytes, (await Call(client, "materialize", sequence: 0)).Bytes);
        Assert.Equal(bytes, (await Call(client, "replay", sequence: 0)).Bytes);
        Assert.Equal(0, (await Call(client, "resolveTime", cutoff: metadata.CreatedAt)).Sequence);
        var restored = await Call(client, "restore", version: first.View.Version.Id, expected: first.View.Head, metadata: metadata);
        Assert.True(restored.Success);
        Assert.Equal(1, restored.View!.State.Epoch);
        Assert.Equal(1, restored.View.State.Sequence);
        var updates = (await Call(client, "updates", expected: first.View.Head)).Update!;
        Assert.True(updates.Reset);
        Assert.Equal(1, Assert.Single(updates.Entries).Commit.Sequence);
        Assert.Equal(restored.View.Head, updates.View.Head);
        var stale = await Call(client, "create", bytes, metadata: metadata, expected: first.View.Head);
        Assert.False(stale.Success);
        Assert.Equal("StaleHead", stale.ErrorCode);
        Assert.Equal(restored.View.Head, (await Call(client, "read")).View!.Head);
    }

    [Fact]
    public void ClientCountersAreLosslessStringsWhileBlobLengthsRemainNumbers()
    {
        var reference = new HistoryBlobReference(new Docxodus.Verification.VerificationDigest
        { Algorithm = "SHA-256", Value = new string('a', 64) }, 123);
        var head = new HistoryHead(long.MaxValue, reference);
        var json = HistoryClientJson.Write(head);
        Assert.Contains("\"revision\":\"9223372036854775807\"", json);
        Assert.Contains("\"length\":123", json);
        Assert.Equal(head, HistoryClientJson.Read<HistoryHead>(json));
        foreach (var value in new[] { "1", "\"01\"", "\"-1\"", "\"+1\"", "\"9223372036854775808\"" })
            Assert.Throws<System.Text.Json.JsonException>(() => HistoryClientJson.Read<HistoryHead>(
                json.Replace("\"9223372036854775807\"", value, StringComparison.Ordinal)));
    }

    [Fact]
    public async Task InvalidRequestsAndCancellationHaveTypedErrorsWithoutPublishing()
    {
        var client = new HistoryClientOps(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        foreach (var json in new[]
        {
            "{}", "null", "{\"schemaVersion\":1,\"operation\":\"read\",\"operation\":\"create\",\"documentId\":\"doc\"}",
            "{\"schemaVersion\":1,\"operation\":\"read\",\"documentId\":null}",
            "{\"schemaVersion\":1,\"operation\":\"read\",\"documentId\":\"doc\",\"unknown\":true}",
            new string(' ', HistoryClientJson.MaxRequestChars + 1),
        })
        {
            var result = HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync(json));
            Assert.False(result.Success);
            Assert.Equal("InvalidRequest", result.ErrorCode);
        }
        Assert.Equal("InvalidRequest", (await Call(client, "create")).ErrorCode);
        Assert.Equal("InvalidRequest", (await Call(client, "invalid")).ErrorCode);
        var unsupported = HistoryClientJson.Write(new HistoryClientRequest { SchemaVersion = 2, DocumentId = "doc", Operation = "read" });
        Assert.Equal("UnsupportedVersion", HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync(unsupported)).ErrorCode);
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Equal("Canceled", HistoryClientJson.Read<HistoryClientResult>(await client.InvokeAsync("{}", cancellationToken: cancel.Token)).ErrorCode);
        Assert.Null((await Call(client, "read")).View);
    }

    private static async Task<HistoryClientResult> Call(HistoryClientOps client, string operation, byte[]? bytes = null,
        DocxVersionMetadata? metadata = null, HistoryHead? expected = null, HistoryBlobReference? version = null,
        long? sequence = null, DateTimeOffset? cutoff = null) => HistoryClientJson.Read<HistoryClientResult>(
            await client.InvokeAsync(HistoryClientJson.Write(new HistoryClientRequest
            {
                SchemaVersion = 1, DocumentId = "doc", Operation = operation, Metadata = metadata,
                ExpectedHead = expected, VersionId = version, Sequence = sequence, Cutoff = cutoff,
            }), bytes));
}
