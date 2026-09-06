// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.History;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class DocxHistoryUpdateTests
{
    [Fact]
    public async Task UpdatesAreContiguousTimestampedAndDoNotFabricateLabelCommits()
    {
        var blobs = new TrackingBlobs(new MemoryHistoryBlobStore());
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var before = DocxSession.CreateBlankDocxBytes();
        var after = Edited(before);
        var first = await history.CreateVersionAsync("doc", null, before, Metadata(12));
        var joined = await history.ReadChangesSinceAsync("doc", null);
        Assert.True(joined.Reset); Assert.Empty(joined.Entries); Assert.Equal(first.Head, joined.View.Head);
        var named = await history.CreateVersionAsync("doc", first.Head, before, Metadata(11));
        var label = await history.ReadChangesSinceAsync("doc", first.Head);
        Assert.False(label.Reset); Assert.Empty(label.Entries); Assert.Equal(named.Head, label.View.Head);
        var changed = await history.CreateVersionAsync("doc", named.Head, after, Metadata(13));
        var restored = await history.RestoreVersionAsync("doc", changed.Head, first.Version.Id, Metadata(10));
        blobs.Reads.Clear();
        var tail = await history.ReadChangesSinceAsync("doc", first.Head);
        Assert.Equal(first.Head, tail.After);
        Assert.Equal(restored.Head, tail.View.Head);
        Assert.True(tail.Reset);
        Assert.Equal(new long[] { 1, 2 }, tail.Entries.Select(entry => entry.Commit.Sequence));
        Assert.Equal(new[] { 13, 10 }, tail.Entries.Select(entry => entry.Metadata.CreatedAt.Hour));
        Assert.Equal(changed.State.Commit, tail.Entries[0].Id);
        Assert.Equal(restored.State.Commit, tail.Entries[1].Id);
        Assert.Equal(tail.Entries[0].Id, tail.Entries[1].Commit.Parent);
        Assert.DoesNotContain(first.State.Snapshot.Blob, blobs.Reads);
        Assert.DoesNotContain(changed.State.Snapshot.Blob, blobs.Reads);
        Assert.DoesNotContain(tail.Entries[0].Commit.Contribution!, blobs.Reads);
        Assert.Equal(restored.Head, await heads.ReadAsync("doc"));
        var duplicate = await history.ReadChangesSinceAsync("doc", restored.Head);
        Assert.Empty(duplicate.Entries); Assert.False(duplicate.Reset);
        // A join on an established history starts at its latest checkpoint, not at sequence zero.
        Assert.Empty((await history.ReadChangesSinceAsync("doc", null)).Entries);
    }

    [Fact]
    public async Task RewindsSameRevisionForksAndEqualContentVersionBranchesAreRejected()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var bytes = DocxSession.CreateBlankDocxBytes();
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata(12));
        var second = await history.CreateVersionAsync("doc", first.Head, bytes, Metadata(13));
        var rewound = new DocxVersionHistory(blobs, new FixedHead(first.Head));
        await Invalid(async () => await rewound.ReadChangesSinceAsync("doc", second.Head));
        var forkedRevision = new DocxVersionHistory(blobs, new FixedHead(new HistoryHead(second.Head.Revision, first.Head.State)));
        await Invalid(async () => await forkedRevision.ReadChangesSinceAsync("doc", second.Head));
        var branch = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var alternate = await branch.CreateVersionAsync("doc", null, bytes, Metadata(12));
        alternate = await branch.CreateVersionAsync("doc", alternate.Head, bytes, Metadata(13));
        Assert.Equal(first.State.InitialSnapshot, alternate.State.InitialSnapshot);
        await Invalid(async () => await branch.ReadChangesSinceAsync("doc", first.Head));
    }

    [Fact]
    public async Task FabricatedRevisionCannotBePairedWithAGenuineAncestorState()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = DocxSession.CreateBlankDocxBytes();
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata(12));
        var head = first.Head;
        for (var i = 0; i < 3; i++)
            head = (await history.CreateVersionAsync("doc", head, bytes, Metadata(13))).Head;
        Assert.Empty((await history.ReadChangesSinceAsync("doc", first.Head)).Entries);
        await Invalid(async () => await history.ReadChangesSinceAsync("doc", first.Head with { Revision = 2 }));
        await Invalid(async () => await history.ReadChangesSinceAsync("doc", first.Head with { Revision = 3 }));
    }

    [Fact]
    public async Task CommitHashNotJustContentAndSequenceMustExtendTheAcceptedTip()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var bytes = DocxSession.CreateBlankDocxBytes();
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata(12));
        var second = await history.CreateVersionAsync("doc", first.Head, Edited(bytes), Metadata(13));
        var third = await history.RestoreVersionAsync("doc", second.Head, first.Version.Id, Metadata(14));
        var records = new HistoryRecordStore(blobs);
        var original = await records.LoadCommitAsync(second.State.Commit!);
        var fake = await records.SaveCommitAsync(original with { Version = third.Version.Id });
        var falseState = await records.SaveStateAsync(third.State with { Commit = await records.SaveCommitAsync(
            (await records.LoadCommitAsync(third.State.Commit!)) with { Parent = fake }) });
        var forged = new DocxVersionHistory(blobs, new FixedHead(new HistoryHead(third.Head.Revision, falseState)));
        await Invalid(async () => await forged.ReadChangesSinceAsync("doc", second.Head));
    }

    [Fact]
    public async Task BudgetIncludesVersionsAndCommitsAndFailuresDoNotPublish()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var bytes = DocxSession.CreateBlankDocxBytes();
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata(12));
        var second = await history.CreateVersionAsync("doc", first.Head, Edited(bytes), Metadata(13));
        var limit = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReadChangesSinceAsync("doc", first.Head, 1));
        Assert.Equal(DocxHistoryError.TraversalLimit, limit.Code);
        Assert.Single((await history.ReadChangesSinceAsync("doc", first.Head, 2)).Entries);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await history.ReadChangesSinceAsync("doc", first.Head,
            cancellationToken: cancellation.Token));
        var missing = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReadChangesSinceAsync("absent", null));
        Assert.Equal(DocxHistoryError.HistoryUnavailable, missing.Code);
        var foreign = await history.CreateVersionAsync("other", null, bytes, Metadata(12));
        var error = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReadChangesSinceAsync("doc", foreign.Head));
        Assert.Equal(DocxHistoryError.ForeignDocument, error.Code);
        Assert.Equal(second.Head, await heads.ReadAsync("doc"));
    }

    private static async Task Invalid(Func<Task> action) =>
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(action)).Code);
    private static DocxVersionMetadata Metadata(int hour) => new()
    { Author = "actor", CreatedAt = new DateTimeOffset(2026, 1, 1, hour, 0, 0, TimeSpan.Zero) };
    private static byte[] Edited(byte[] bytes)
    {
        using var session = new DocxSession(bytes);
        var anchor = session.Project().AnchorIndex.Keys.First(key => key.StartsWith("p:", StringComparison.Ordinal));
        Assert.True(session.ReplaceText(anchor, "updated live log").Success);
        return session.Save();
    }
    private sealed class FixedHead(HistoryHead head) : IHistoryHeadStore
    {
        public ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default) => new(head);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
            HistoryBlobReference state, CancellationToken cancellationToken = default) => throw new InvalidOperationException("Read only");
    }
    private sealed class TrackingBlobs(IHistoryBlobStore inner) : IHistoryBlobStore
    {
        public List<HistoryBlobReference> Reads { get; } = new();
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default) =>
            inner.PutAsync(reference, content, cancellationToken);
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
        { Reads.Add(reference); return inner.OpenReadAsync(reference, cancellationToken); }
    }
}
