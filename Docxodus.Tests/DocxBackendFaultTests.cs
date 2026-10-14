#nullable enable

using Docxodus.History;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

public sealed class DocxBackendFaultTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task EveryDurableWriteAndCASBoundaryRecoversAcceptedAndConflictingOperations(bool filesystem, bool conflicting)
    {
        var bytes = Package("abcd");
        int writes;
        using (var probe = new HistoryFaultHarness(filesystem))
        {
            var prepared = await Setup(probe);
            probe.Arm(HistoryFaultHarness.Fault.None);
            _ = await new DocxVersionHistory(probe, probe).SubmitOperationAsync("doc", prepared.Request);
            writes = probe.Writes;
        }
        Assert.True(writes >= 3);
        var cases = Enumerable.Range(1, writes).SelectMany(at => new[]
        {
            (Fault: HistoryFaultHarness.Fault.BeforeBlob, At: at), (Fault: HistoryFaultHarness.Fault.AfterBlob, At: at),
        }).Concat(new[] { HistoryFaultHarness.Fault.BeforeHead, HistoryFaultHarness.Fault.AfterHead,
            HistoryFaultHarness.Fault.CancelBeforeHead, HistoryFaultHarness.Fault.CancelAfterHead }.Select(f => (Fault: f, At: 1)));
        foreach (var fault in cases)
        {
            using var store = new HistoryFaultHarness(filesystem);
            var prepared = await Setup(store);
            store.Cancellation = new CancellationTokenSource(); store.Arm(fault.Fault, fault.At);
            var action = new DocxVersionHistory(store, store).SubmitOperationAsync("doc", prepared.Request,
                cancellationToken: store.Cancellation.Token).AsTask();
            if (fault.Fault == HistoryFaultHarness.Fault.CancelAfterHead) Assert.NotNull(await action);
            else if (fault.Fault == HistoryFaultHarness.Fault.CancelBeforeHead) await Assert.ThrowsAnyAsync<OperationCanceledException>(() => action);
            else await Assert.ThrowsAsync<IOException>(() => action);
            var committed = fault.Fault is HistoryFaultHarness.Fault.AfterHead or HistoryFaultHarness.Fault.CancelAfterHead;
            var recovered = new DocxVersionHistory(store, store);
            var observed = (await recovered.ReadAsync("doc"))!;
            Assert.Equal(prepared.Before.Head.Revision + (committed ? 1 : 0), observed.Head.Revision);
            if (!committed) Assert.Equal(prepared.Before.Head, observed.Head);
            store.Arm(HistoryFaultHarness.Fault.None);
            var result = await recovered.SubmitOperationAsync("doc", prepared.Request);
            Assert.Equal(committed ? 0 : 1, store.Publications);
            Assert.Equal(prepared.Before.Head.Revision + 1, result.View.Head.Revision);
            Assert.Equal(conflicting ? "conflict" : "accepted", result.Operation.Record.Status);
            Assert.Equal(conflicting ? "aXd" : "aXd!", BodyText(await recovered.ExportVersionAsync("doc", result.View.Version.Id)));
            Assert.Equal(conflicting ? "ab?d" : "abcd!", BodyText(await recovered.ExportOperationProposalAsync("doc", result.Operation.Id)));
            Assert.Equal(2, (await recovered.ReadOperationsSinceAsync("doc", null)).Operations.Count);
            Assert.Equal(conflicting ? 2 : 3, (await recovered.ListVersionsAsync("doc")).Versions.Count);
            var again = await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", prepared.Request);
            Assert.Equal(result.View.Head, again.View.Head); Assert.Equal(result.Operation.Id, again.Operation.Id);
            SameEntries(await recovered.ExportVersionAsync("doc", result.View.Version.Id), await recovered.ReplayAsync("doc", result.View.State.Sequence));
        }

        async Task<(DocxHistoryView Before, DocxOperationRequest Request)> Setup(HistoryFaultHarness store)
        {
            var history = new DocxVersionHistory(store, store);
            var initial = await history.CreateVersionAsync("doc", "initial", null, bytes, Metadata("initial"));
            var winner = await history.SubmitOperationAsync("doc", Text("winner", initial.Head, 1, 2, "X"));
            var request = conflicting ? Text("request", initial.Head, 2, 1, "?") : Text("request", initial.Head, 4, 0, "!");
            return (winner.View, request);
        }
    }

    [Fact]
    public async Task DecisionAuditReadsNoRadixIndexAndBudgetCountsInterveningVersionPublications()
    {
        var blobs = new MemoryHistoryBlobStore(); var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var current = await history.CreateVersionAsync("doc", "initial", null, Package("abcd"), Metadata("initial"));
        var roots = new HashSet<HistoryBlobReference>();
        for (var i = 0; i < 8; i++)
        {
            current = (await history.SubmitOperationAsync("doc", Text("operation-" + i, current.Head, 0, 0, "!"))).View;
            if (current.State.Requests!.Index is { } root) roots.Add(root);
        }
        // A forbidden root catches even ONE restarted receipt lookup, regardless of index depth.
        var audited = new DocxVersionHistory(new ForbiddenIndexReads(blobs, roots), heads);
        Assert.Equal(8, (await audited.ReadOperationsSinceAsync("doc", null, 8)).Operations.Count);
        Assert.Equal(DocxHistoryError.TraversalLimit, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await audited.ReadOperationsSinceAsync("doc", null, 7))).Code);
        var snapshot = await history.ExportVersionAsync("doc", current.Version.Id);
        for (var i = 0; i < 4; i++)
            current = await history.CreateVersionAsync("doc", current.Head, snapshot, Metadata("label"));
        Assert.Equal(DocxHistoryError.TraversalLimit, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.ReadOperationsSinceAsync("doc", null, 8))).Code);
        Assert.Equal(8, (await history.ReadOperationsSinceAsync("doc", null, 12)).Operations.Count);
    }

    private sealed class ForbiddenIndexReads(IHistoryBlobStore inner, HashSet<HistoryBlobReference> roots) : IHistoryBlobStore
    {
        public ValueTask PutAsync(HistoryBlobReference reference, Stream stream, CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("Read-only audit.");
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
        {
            Assert.DoesNotContain(reference, roots);
            return inner.OpenReadAsync(reference, cancellationToken);
        }
    }
}
