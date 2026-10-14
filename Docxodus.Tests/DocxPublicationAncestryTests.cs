#nullable enable

using System.Text;
using System.Text.Json.Nodes;
using Docxodus.History;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxPublicationAncestryTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DecisionOnlyPublicationsInterleaveWithVersionsRestoresAndLegacyReceipts(bool filesystem)
    {
        using var store = new HistoryFaultHarness(filesystem);
        var history = new DocxVersionHistory(store, store);
        var bytes = DocxVersionRequestTests.Document("initial");
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata("initial"));
        var label = await history.CreateVersionAsync("doc", "label", first.Head, bytes, Metadata("label"));
        var decision = await PublishDecision(store, label, "conflict");
        var changed = await history.CreateVersionAsync("doc", "edit", decision.Head,
            DocxVersionRequestTests.Document("changed"), Metadata("edit"));
        var another = await PublishDecision(store, changed, "discard");
        var restored = await history.RestoreVersionAsync("doc", "restore", another.Head, first.Version.Id, Metadata("restore"));
        var last = await history.CreateVersionAsync("doc", restored.Head, bytes, Metadata("legacy label"));
        history = new DocxVersionHistory(store, store);
        foreach (var view in new[] { first, label, decision, changed, another, restored, last })
        {
            var updates = await history.ReadChangesSinceAsync("doc", view.Head);
            Assert.Equal(last.Head, updates.View.Head);
            Assert.Equal(2 - view.State.Sequence, updates.Entries.Count);
            Assert.Equal(view.State.Epoch == 0, updates.Reset);
            Assert.Equal(bytes, await history.ExportVersionAsync("doc", first.Version.Id));
        }
        Assert.Equal(5, (await history.ListVersionsAsync("doc")).Versions.Count);
        Assert.Equal(7, last.Head.Revision);
        Assert.Equal(2, last.State.Sequence);
        Assert.Equal(another.State.Operation, last.State.Operation);
        Assert.Equal(restored.Head, last.State.ParentPublication);
        Assert.Equal(label.Head, (await history.CreateVersionAsync("doc", "label", first.Head, bytes, Metadata("label"))).Head);
        Assert.Equal(2, (await history.ReadChangesSinceAsync("doc", first.Head, 8)).Entries.Count);
        Assert.Equal(DocxHistoryError.TraversalLimit, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.ReadChangesSinceAsync("doc", first.Head, 7))).Code);
        var wire = HistoryClientJson.Write(new HistoryClientResult { View = last });
        Assert.Contains("\"parentPublication\"", wire);
        Assert.Equal(last.State.ParentPublication, HistoryClientJson.Read<HistoryClientResult>(wire).View!.State.ParentPublication);
    }

    [Fact]
    public async Task AncestryRejectsSameVersionForksSkippedParentsAndForgedRevisions()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var first = await history.CreateVersionAsync("doc", null, DocxVersionRequestTests.Document("initial"), Metadata("initial"));
        var a = await PublishDecision(store, first, "a");
        var b = await PublishDecision(store, a, "b");
        var records = new HistoryRecordStore(store);
        async Task Invalid(DocxHistoryStateRecord state, HistoryHead after, long revision = 3)
        {
            var reference = await records.SaveStateAsync(state);
            var forged = new DocxVersionHistory(store, new FixedHead(new HistoryHead(revision, reference)));
            Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
                await forged.ReadChangesSinceAsync("doc", after))).Code);
        }
        var alternateId = await records.SaveStateAsync(a.State with { Operation = first.Version.Id });
        await Invalid(b.State with { ParentPublication = new HistoryHead(2, alternateId) }, a.Head);
        await Invalid(b.State with { ParentPublication = first.Head }, first.Head);
        await Invalid(b.State with { Operation = a.State.Operation }, a.Head);
        await Invalid(b.State, a.Head, revision: 4);
        await Invalid(b.State with { Operation = null, ParentPublication = null }, a.Head);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await history.ReadChangesSinceAsync("doc", first.Head, cancellationToken: canceled.Token));
        Assert.Equal(b.Head, (await history.ReadAsync("doc"))!.Head);
    }

    [Fact]
    public async Task V3RequiresBothFieldsAndOlderCodecsCannotSilentlyDiscardThem()
    {
        using var store = new HistoryFaultHarness();
        var history = new DocxVersionHistory(store, store);
        var first = await history.CreateVersionAsync("doc", null, DocxVersionRequestTests.Document("initial"), Metadata("initial"));
        var decision = await PublishDecision(store, first, "a");
        var records = new HistoryRecordStore(store);
        Assert.Equal(decision.State, await records.LoadStateAsync(decision.Head.State));
        using var input = await store.OpenReadAsync(decision.Head.State);
        var json = (await JsonNode.ParseAsync(input!))!;
        Assert.Equal(3, json["schemaVersion"]!.GetValue<int>());
        foreach (var version in new[] { 1, 2 })
        {
            var old = json.DeepClone(); old["schemaVersion"] = version;
            old["schema"] = "https://docxodus.dev/schemas/history/package-state/v" + version;
            old["record"]!["operation"] = null; old["record"]!["parentPublication"] = null;
            await Bad(old);
        }
        foreach (var field in new[] { "operation", "parentPublication" })
        {
            var missing = json.DeepClone(); missing["record"]!.AsObject().Remove(field); await Bad(missing);
        }
        async Task Bad(JsonNode malformed)
        {
            var reference = await HistoryBlobIO.PutBytesAsync(store, Encoding.UTF8.GetBytes(malformed.ToJsonString()), default);
            Assert.Equal(PackageChangeError.InvalidManifest, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await records.LoadStateAsync(reference))).Code);
        }
    }

    // Opaque decision stubs exercise only the publication contract, not backend reconciliation.
    private static async Task<DocxHistoryView> PublishDecision(HistoryFaultHarness store, DocxHistoryView before, string id)
    {
        var operation = await HistoryBlobIO.PutBytesAsync(store, Encoding.UTF8.GetBytes(id), default);
        var requests = await new HistoryRequestJournalStore(store).AdvanceAsync("doc", before.Head, before.State.Requests,
            new HistoryRequestIdentity(id, operation.Digest));
        var state = before.State with { Operation = operation, ParentPublication = before.Head, Requests = requests };
        var reference = await new HistoryRecordStore(store).SaveStateAsync(state);
        var head = await store.TryAdvanceAsync("doc", before.Head, reference);
        return new DocxHistoryView(head!, state, before.Version);
    }

    private static DocxVersionMetadata Metadata(string label) => DocxVersionRequestTests.Metadata(label);
    private sealed class FixedHead(HistoryHead head) : IHistoryHeadStore
    {
        public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default) => new(head);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => throw new InvalidOperationException("Read only");
    }
}
