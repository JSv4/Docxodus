#nullable enable

using System.Text;
using System.Text.Json.Nodes;
using Docxodus.History;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

public sealed class DocxBackendRecordTests
{
    [Fact]
    public async Task StrictDecisionCodecRejectsUnknownDuplicateMissingAndUnsupportedFields()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var first = await history.CreateVersionAsync("doc", null, Package("abcd"), Metadata("initial"));
        var result = await history.SubmitOperationAsync("doc", Text("a", first.Head, 1, 0, "X"));
        var recordStore = new DocxOperationStore(store);
        var record = await recordStore.LoadDecisionAsync(result.Operation.Id, default);
        Assert.Equal(result.Operation.Record, record);
        Assert.Equal(result.Operation.Id, await recordStore.SaveDecisionAsync(record, default));
        using var stream = await store.OpenReadAsync(result.Operation.Id);
        using var reader = new StreamReader(stream!);
        var original = await reader.ReadToEndAsync();
        await Bad(original.Replace("\"status\":\"accepted\"", "\"status\":\"accepted\",\"status\":\"accepted\"", StringComparison.Ordinal));
        var unknown = JsonNode.Parse(original)!; unknown["record"]!["extra"] = true; await Bad(unknown.ToJsonString());
        foreach (var property in new[] { "appliedText", "contentCommit", "conflict", "proposedSnapshot", "input" })
        {
            var missing = JsonNode.Parse(original)!; missing["record"]!.AsObject().Remove(property); await Bad(missing.ToJsonString());
        }
        var future = JsonNode.Parse(original)!; future["schemaVersion"] = 2;
        await Bad(future.ToJsonString(), PackageChangeError.UnsupportedVersion);
        var badStatus = JsonNode.Parse(original)!; badStatus["record"]!["status"] = "conflict"; await Bad(badStatus.ToJsonString());
        var foreign = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.GetOperationAsync("other", result.Operation.Id));
        Assert.Equal(DocxHistoryError.ForeignDocument, foreign.Code);

        async Task Bad(string json, PackageChangeError code = PackageChangeError.InvalidManifest)
        {
            var reference = await HistoryBlobIO.PutBytesAsync(store, Encoding.UTF8.GetBytes(json), default);
            Assert.Equal(code, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await recordStore.LoadDecisionAsync(reference, default))).Code);
        }
    }

    [Fact]
    public async Task CorruptMissingDecisionInputAndProposalAreNotEmptyHistoryOrSuccessfulRetry()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var first = await history.CreateVersionAsync("doc", null, Package("abcd"), Metadata("initial"));
        var request = Text("a", first.Head, 1, 0, "X");
        var result = await history.SubmitOperationAsync("doc", request);
        foreach (var reference in new[] { result.Operation.Id, result.Operation.Record.Input })
        foreach (var missing in new[] { false, true })
        {
            store.DamagedDigest = reference.Digest.Value; store.Missing = missing;
            Assert.Equal(missing ? PackageChangeError.PayloadMissing : PackageChangeError.PayloadMismatch,
                (await Assert.ThrowsAsync<PackageChangeException>(async () => await history.ReadOperationsSinceAsync("doc", null))).Code);
            await Assert.ThrowsAsync<PackageChangeException>(async () => await history.SubmitOperationAsync("doc", request));
        }
        store.DamagedDigest = result.Operation.Record.ProposedSnapshot.Blob.Digest.Value; store.Missing = true;
        await Assert.ThrowsAsync<PackageChangeException>(async () => await history.ExportOperationProposalAsync("doc", result.Operation.Id));
        store.DamagedDigest = null;
        Assert.Equal(result.View.Head, (await history.ReadAsync("doc"))!.Head);
        Assert.Equal(result.Operation.Id, (await history.SubmitOperationAsync("doc", request)).Operation.Id);
    }

    [Fact]
    public async Task AClaimedPositionMapMustReproduceActualAcceptedEffectsBeforeRebasing()
    {
        var blobs = new MemoryHistoryBlobStore(); var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var first = await history.CreateVersionAsync("doc", null, Package("abcd"), Metadata("initial"));
        var accepted = await history.SubmitOperationAsync("doc", Text("accepted", first.Head, 1, 0, "X"));
        var falseDecision = await new DocxOperationStore(blobs).SaveDecisionAsync(accepted.Operation.Record with
        { AppliedText = accepted.Operation.Record.AppliedText! with { Offset = 2 } }, default);
        var falseState = await new HistoryRecordStore(blobs).SaveStateAsync(accepted.View.State with { Operation = falseDecision });
        var forged = new DocxVersionHistory(blobs, new FixedHead(accepted.View.Head with { State = falseState }));
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await forged.SubmitOperationAsync("doc", Text("pending", first.Head, 3, 0, "Y")))).Code);
        Assert.Equal(accepted.View.Head, (await history.ReadAsync("doc"))!.Head);
    }

    [Fact]
    public async Task ReadSetsAndMetadataNormalizeButChangedCandidateOrResolutionNeverReuseAnId()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var bytes = Package("abcd");
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata("initial"));
        var a = PackageRequest("package", first.Head) with
        {
            ReadParts = new[] { "/word/footnotes.xml", "/word/header1.xml", "/word/header1.xml" },
            Metadata = Metadata("package") with { ApplicationMetadata = new Dictionary<string, string> { ["b"] = "2", ["a"] = "1" } },
        };
        var result = await history.SubmitOperationAsync("doc", a, bytes);
        var reordered = a with
        {
            ReadParts = new[] { "/word/header1.xml", "/word/footnotes.xml" },
            Metadata = a.Metadata with { ApplicationMetadata = new Dictionary<string, string> { ["a"] = "1", ["b"] = "2" } },
        };
        Assert.Equal(result.View.Head, (await history.SubmitOperationAsync("doc", reordered, bytes)).View.Head);
        foreach (var changed in new[] { a with { ReadParts = Array.Empty<string>() }, a with { Resolves = result.Operation.Id } })
            Assert.Equal(DocxHistoryError.RequestConflict, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
                await history.SubmitOperationAsync("doc", changed, bytes))).Code);
        Assert.Equal(DocxHistoryError.RequestConflict, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.SubmitOperationAsync("doc", a, ReplacePartText(bytes, "word/document.xml", "abcd", "new")))).Code);
    }

    private sealed class FixedHead(HistoryHead head) : IHistoryHeadStore
    {
        public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default) => new(head);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => throw new InvalidOperationException("Must not publish an unverified text map.");
    }
}
