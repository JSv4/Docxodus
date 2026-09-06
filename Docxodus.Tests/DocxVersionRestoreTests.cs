// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.History;
using Xunit;

namespace Docxodus.Tests;

public class DocxVersionRestoreTests
{
    [Fact]
    public async Task RestoreAppendsAnAuditableVersionAndEpochWithoutChangingAnyEarlierBytes()
    {
        var directory = Directory.CreateTempSubdirectory("docxodus-version-restore-").FullName;
        try
        {
            var blobs = new FileHistoryBlobStore(Path.Combine(directory, "blobs"));
            var heads = new FileHistoryHeadStore(Path.Combine(directory, "heads"));
            var history = new DocxVersionHistory(blobs, heads);
            var versions = new List<DocxHistoryView>();
            var expected = new List<byte[]>();
            foreach (var text in new[] { "one", "two", "three" })
            {
                var bytes = Document(text);
                expected.Add(bytes);
                versions.Add(await history.CreateVersionAsync("doc", versions.LastOrDefault()?.Head, bytes, Metadata(text)));
            }
            var restored = await history.RestoreVersionAsync("doc", versions[2].Head, versions[0].Version.Id, Metadata("restore one"));
            Assert.Equal(3, restored.State.Sequence);
            Assert.Equal(1, restored.State.Epoch);
            Assert.Equal(versions[2].Version.Id, restored.Version.Record.Parent);
            Assert.Equal(versions[0].Version.Id, restored.Version.Record.RestoredFrom);
            Assert.Equal(versions[0].State.Snapshot, restored.State.Snapshot);
            Assert.NotEqual(versions[0].Version.Id, restored.Version.Id);
            var commit = await new HistoryRecordStore(blobs).LoadCommitAsync(restored.State.Commit!);
            Assert.Equal("restore", commit.Kind);
            Assert.Null(commit.Contribution);
            Assert.Equal(versions[2].State.Commit, commit.Parent);
            Assert.Equal(versions[2].State.Snapshot, commit.Before);
            Assert.Equal(versions[0].State.Snapshot, commit.After);
            var reopened = new DocxVersionHistory(new FileHistoryBlobStore(Path.Combine(directory, "blobs")),
                new FileHistoryHeadStore(Path.Combine(directory, "heads")));
            Assert.Equal(restored.State, (await reopened.ReadAsync("doc"))!.State);
            Assert.Equal(expected[0], await reopened.ExportVersionAsync("doc", restored.Version.Id));
            for (var i = 0; i < versions.Count; i++)
                Assert.Equal(expected[i], await reopened.ExportVersionAsync("doc", versions[i].Version.Id));
            Assert.Equal(4, (await reopened.ListVersionsAsync("doc")).Versions.Count);
            var edited = await reopened.CreateVersionAsync("doc", restored.Head, Document("after restore"), Metadata("next"));
            Assert.Equal(4, edited.State.Sequence);
            Assert.Equal(1, edited.State.Epoch);
            Assert.Equal(edited.State, (await reopened.ReadAsync("doc"))!.State);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task SameContentRestoreStillRecordsTheExplicitResetAndStalePreviewsFail()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = Document("one");
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata("one"));
        var restored = await history.RestoreVersionAsync("doc", first.Head, first.Version.Id, Metadata("reset"));
        Assert.Equal(1, restored.State.Sequence);
        Assert.Equal(1, restored.State.Epoch);
        Assert.Equal(first.State.Snapshot, restored.State.Snapshot);
        var stale = await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.RestoreVersionAsync("doc", first.Head, first.Version.Id, Metadata("stale")));
        Assert.Equal(DocxHistoryError.StaleHead, stale.Code);
        Assert.Equal(restored.Head, (await history.ReadAsync("doc"))!.Head);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public async Task FailedRestoreMetadataOrHeadWritesNeverPublishAPartialReset(int failAt)
    {
        var blobs = new FaultingBlobs();
        var heads = new FaultingHeads();
        var history = new DocxVersionHistory(blobs, heads);
        var first = await history.CreateVersionAsync("doc", null, Document("one"), Metadata("one"));
        var second = await history.CreateVersionAsync("doc", first.Head, Document("two"), Metadata("two"));
        blobs.Writes = 0;
        blobs.FailAt = failAt;
        heads.Fail = failAt == 4;
        await Assert.ThrowsAsync<IOException>(async () =>
            await history.RestoreVersionAsync("doc", second.Head, first.Version.Id, Metadata("restore")));
        Assert.Equal(second.Head, (await history.ReadAsync("doc"))!.Head);
        Assert.Equal(0, (await history.ReadAsync("doc"))!.State.Epoch);
        Assert.Equal(2, (await history.ListVersionsAsync("doc")).Versions.Count);
        blobs.FailAt = 0;
        heads.Fail = false;
        Assert.Equal(1, (await history.RestoreVersionAsync("doc", second.Head, first.Version.Id, Metadata("retry"))).State.Epoch);
    }

    [Fact]
    public async Task MissingCorruptOrForeignTargetSnapshotsCannotResetTheHead()
    {
        var blobs = new FaultingBlobs();
        var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var first = await history.CreateVersionAsync("doc", null, Document("one"), Metadata("one"));
        var second = await history.CreateVersionAsync("doc", first.Head, Document("two"), Metadata("two"));
        blobs.Blocked = first.State.Snapshot.Blob;
        var missing = await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await history.RestoreVersionAsync("doc", second.Head, first.Version.Id, Metadata("missing")));
        Assert.Equal(PackageChangeError.PayloadMissing, missing.Code);
        blobs.Corrupt = true;
        var corrupt = await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await history.RestoreVersionAsync("doc", second.Head, first.Version.Id, Metadata("corrupt")));
        Assert.Equal(PackageChangeError.PayloadMismatch, corrupt.Code);
        Assert.Equal(second.Head, (await history.ReadAsync("doc"))!.Head);
        blobs.Blocked = null;
        var other = await history.CreateVersionAsync("other", null, Document("other"), Metadata("other"));
        var foreign = await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.RestoreVersionAsync("doc", second.Head, other.Version.Id, Metadata("foreign")));
        Assert.Equal(DocxHistoryError.ForeignDocument, foreign.Code);
    }

    private static DocxVersionMetadata Metadata(string label) => new() { Author = "host", CreatedAt = DateTimeOffset.UnixEpoch, Label = label };
    private static byte[] Document(string text)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        Assert.True(session.ReplaceText(session.Project().AnchorIndex.Keys.First(), text).Success);
        return session.Save();
    }
    private sealed class FaultingBlobs : IHistoryBlobStore
    {
        private readonly MemoryHistoryBlobStore _inner = new();
        internal int Writes { get; set; }
        internal int FailAt { get; set; }
        internal HistoryBlobReference? Blocked { get; set; }
        internal bool Corrupt { get; set; }
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken)
        {
            if (++Writes == FailAt) throw new IOException("Injected blob failure.");
            return _inner.PutAsync(reference, content, cancellationToken);
        }
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken) =>
            reference == Blocked ? ValueTask.FromResult<Stream?>(Corrupt ? new MemoryStream(new byte[reference.Length]) : null)
                : _inner.OpenReadAsync(reference, cancellationToken);
    }
    private sealed class FaultingHeads : IHistoryHeadStore
    {
        private readonly MemoryHistoryHeadStore _inner = new();
        internal bool Fail { get; set; }
        public ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken) => _inner.ReadAsync(documentId, cancellationToken);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected, HistoryBlobReference state, CancellationToken cancellationToken) =>
            Fail ? throw new IOException("Injected head failure.") : _inner.TryAdvanceAsync(documentId, expected, state, cancellationToken);
    }
}
