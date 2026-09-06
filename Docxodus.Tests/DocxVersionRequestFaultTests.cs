#nullable enable

using Docxodus.History;
using Xunit;
using static Docxodus.Tests.DocxVersionRequestTests;

namespace Docxodus.Tests;

public sealed class DocxVersionRequestFaultTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task EveryImmutableWriteAndHeadBoundaryIsRetryableWithoutDuplicatePublication(bool restore, bool filesystem)
    {
        var initial = Document("initial");
        var candidate = Document("new contents");
        async Task<(DocxVersionHistory History, DocxHistoryView First, DocxHistoryView Current)> Setup(HistoryFaultHarness storage)
        {
            var history = new DocxVersionHistory(storage, storage);
            var first = await history.CreateVersionAsync("doc", "first", null, initial, Metadata("first"));
            var current = first;
            for (var i = 0; i < 6; i++) current = await history.CreateVersionAsync("doc", "label-" + i, current.Head, initial, Metadata("label"));
            return (history, first, current);
        }
        ValueTask<DocxHistoryView> Publish(DocxVersionHistory history, DocxHistoryView first, DocxHistoryView current, CancellationToken cancellationToken = default) => restore
            ? history.RestoreVersionAsync("doc", "tested", current.Head, first.Version.Id, Metadata("tested"), cancellationToken)
            : history.CreateVersionAsync("doc", "tested", current.Head, candidate, Metadata("tested"), cancellationToken);

        int writeCount;
        using (var storage = new HistoryFaultHarness(filesystem))
        {
            var setup = await Setup(storage);
            storage.Arm(HistoryFaultHarness.Fault.None);
            await Publish(setup.History, setup.First, setup.Current);
            writeCount = storage.Writes;
            Assert.True(writeCount >= 5); // Includes index promotion and application state.
        }
        var cases = Enumerable.Range(1, writeCount).SelectMany(at => new[]
        {
            (HistoryFaultHarness.Fault.BeforeBlob, at), (HistoryFaultHarness.Fault.AfterBlob, at),
        }).Concat(new[]
        {
            (HistoryFaultHarness.Fault.BeforeHead, 1), (HistoryFaultHarness.Fault.AfterHead, 1),
            (HistoryFaultHarness.Fault.CancelBeforeHead, 1), (HistoryFaultHarness.Fault.CancelAfterHead, 1),
        });
        foreach (var (fault, at) in cases)
        {
            using var storage = new HistoryFaultHarness(filesystem);
            var (history, first, current) = await Setup(storage);
            storage.Cancellation = new CancellationTokenSource();
            storage.Arm(fault, at);
            if (fault is HistoryFaultHarness.Fault.CancelAfterHead)
                await Publish(history, first, current, storage.Cancellation.Token);
            else if (fault is HistoryFaultHarness.Fault.CancelBeforeHead)
                await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await Publish(history, first, current, storage.Cancellation.Token));
            else
                await Assert.ThrowsAsync<IOException>(async () => await Publish(history, first, current));
            var committed = fault is HistoryFaultHarness.Fault.AfterHead or HistoryFaultHarness.Fault.CancelAfterHead;
            storage.Arm(HistoryFaultHarness.Fault.None);
            history = new DocxVersionHistory(storage, storage); // Discard the producer; no cache may be needed.
            var observed = (await history.ReadAsync("doc"))!;
            Assert.Equal(current.Head.Revision + (committed ? 1 : 0), observed.Head.Revision);
            if (!committed) Assert.Equal(current.Head, observed.Head);
            var retried = await Publish(history, first, current);
            Assert.Equal(current.Head.Revision + 1, retried.Head.Revision);
            Assert.Equal(committed ? 0 : 1, storage.Publications);
            Assert.Equal(retried.Head, (await Publish(history, first, current)).Head);
            Assert.Equal(restore ? initial : candidate, await history.ExportVersionAsync("doc", retried.Version.Id));
            Assert.Equal(first.Head, (await history.CreateVersionAsync("doc", "first", null, initial, Metadata("first"))).Head);
            Assert.Equal(8, (await history.ListVersionsAsync("doc")).Versions.Count);
            Assert.Equal(restore ? 1 : 0, retried.State.Epoch);
        }
    }
}
