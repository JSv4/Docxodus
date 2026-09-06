// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DocxVersionHistoryTests
{
    [Fact]
    public async Task FilesystemVersionsReopenWithExactBytesMetadataAndReversibleImport()
    {
        var directory = Directory.CreateTempSubdirectory("docxodus-version-history-").FullName;
        try
        {
            var blobs = new FileHistoryBlobStore(Path.Combine(directory, "blobs"));
            var heads = new FileHistoryHeadStore(Path.Combine(directory, "heads"));
            var history = new DocxVersionHistory(blobs, heads);
            var (before, after) = Documents();
            var first = await history.CreateVersionAsync("doc", null, before, Metadata("first"));
            var named = await history.CreateVersionAsync("doc", first.Head, before, Metadata("named"));
            Assert.Equal(0, named.State.Sequence);
            Assert.Equal(2, named.Head.Revision);
            Assert.Null(named.State.Commit);
            Assert.Equal(first.State.Snapshot, named.State.Snapshot);
            Assert.NotEqual(first.Version.Id, named.Version.Id);
            Assert.Equal(first.Version.Id, named.Version.Record.Parent);
            var edited = await history.CreateVersionAsync("doc", named.Head, after, Metadata("edited"));
            Assert.Equal(1, edited.State.Sequence);
            Assert.Equal(3, edited.Head.Revision);
            Assert.NotNull(edited.State.Commit);
            var reopened = new DocxVersionHistory(new FileHistoryBlobStore(Path.Combine(directory, "blobs")),
                new FileHistoryHeadStore(Path.Combine(directory, "heads")));
            var read = await reopened.ReadAsync("doc");
            Assert.Equal(edited.State, read!.State);
            Assert.Equal(edited.Head, read.Head);
            Assert.Equal("edited", read.Version.Record.Metadata.Label);
            Assert.Equal("matter-123", read.Version.Record.Metadata.ApplicationMetadata["matter"]);
            Assert.Equal(before, await reopened.ExportVersionAsync("doc", first.Version.Id));
            Assert.Equal(after, await reopened.ExportVersionAsync("doc", edited.Version.Id));

            var commit = await new HistoryRecordStore(blobs).LoadCommitAsync(edited.State.Commit!);
            Assert.Equal(edited.Version.Id, commit.Version);
            using var manifest = await blobs.OpenReadAsync(commit.Contribution!);
            using var buffer = new MemoryStream();
            await manifest!.CopyToAsync(buffer);
            var changes = await PackageChangeSetCodec.LoadAsync(buffer.ToArray(), blobs);
            Assert.Equal(edited.State.Snapshot.ContentDigest, PackageManifestGenerator.Generate(changes.Apply(before)).OrderedOpcContentDigest);
            Assert.Equal(first.State.Snapshot.ContentDigest, PackageManifestGenerator.Generate(changes.Invert().Apply(after)).OrderedOpcContentDigest);
            var comparison = await reopened.CompareVersionsAsync("doc", first.Version.Id, edited.Version.Id);
            Assert.NotEmpty(comparison.GetRevisions());
            using var repackBuffer = new MemoryStream();
            repackBuffer.Write(after);
            using (var zip = new ZipArchive(repackBuffer, ZipArchiveMode.Update, leaveOpen: true))
                foreach (var entry in zip.Entries) entry.LastWriteTime = new DateTimeOffset(2002, 3, 4, 5, 6, 8, TimeSpan.Zero);
            var repacked = repackBuffer.ToArray();
            var renamed = await reopened.CreateVersionAsync("doc", edited.Head, repacked, Metadata("named repack"));
            Assert.Equal(edited.State.Sequence, renamed.State.Sequence);
            Assert.Equal(edited.State.Commit, renamed.State.Commit);
            Assert.NotEqual(edited.State.Snapshot.Blob.Digest, renamed.State.Snapshot.Blob.Digest);
            Assert.Equal(renamed.State, (await reopened.ReadAsync("doc"))!.State);
            Assert.Equal(repacked, await reopened.ExportVersionAsync("doc", renamed.Version.Id));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task PaginationRemainsOnItsCapturedChainWhileNewVersionsPublish()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = DocxSession.CreateBlankDocxBytes();
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata("one"));
        var second = await history.CreateVersionAsync("doc", first.Head, bytes, Metadata("two"));
        var page = await history.ListVersionsAsync("doc", limit: 1);
        Assert.Equal(second.Version.Id, Assert.Single(page.Versions).Id);
        await history.CreateVersionAsync("doc", second.Head, bytes, Metadata("three"));
        var rest = await history.ListVersionsAsync("doc", page.Next, limit: 1);
        Assert.Equal(first.Version.Id, Assert.Single(rest.Versions).Id);
        Assert.Null(rest.Next);
        Assert.Equal(3, (await history.ListVersionsAsync("doc")).Versions.Count);
        Assert.Empty((await history.ListVersionsAsync("absent")).Versions);
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await history.ListVersionsAsync("doc", limit: 101));
    }

    [Fact]
    public async Task CompetingPublicationsProduceOneWinnerWithoutOverwritingHistory()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var bytes = DocxSession.CreateBlankDocxBytes();
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata("one"));
        var results = await Task.WhenAll(Enumerable.Range(0, 4).Select(i => Task.Run(async () =>
        {
            try
            {
                await new DocxVersionHistory(blobs, heads).CreateVersionAsync("doc", first.Head, bytes, Metadata($"contender {i}"));
                return true;
            }
            catch (DocxHistoryException error) when (error.Code == DocxHistoryError.StaleHead) { return false; }
        })));
        Assert.Single(results.Where(result => result));
        Assert.Equal(2, (await history.ListVersionsAsync("doc")).Versions.Count);
        Assert.Equal(bytes, await history.ExportVersionAsync("doc", first.Version.Id));
    }

    [Fact]
    public async Task FailureAtEveryBlobWriteOrHeadPublicationLeavesThePreviousHistoryVisible()
    {
        var (before, after) = Documents();
        var successfulBlobs = new FaultingBlobs();
        var successful = new DocxVersionHistory(successfulBlobs, new MemoryHistoryHeadStore());
        var original = await successful.CreateVersionAsync("doc", null, before, Metadata("one"));
        successfulBlobs.Writes = 0;
        await successful.CreateVersionAsync("doc", original.Head, after, Metadata("two"));
        var writeCount = successfulBlobs.Writes;
        Assert.True(writeCount >= 6);
        for (var failAt = 1; failAt <= writeCount + 1; failAt++)
        {
            var blobs = new FaultingBlobs();
            var heads = new FaultingHeads();
            var history = new DocxVersionHistory(blobs, heads);
            var first = await history.CreateVersionAsync("doc", null, before, Metadata("one"));
            blobs.Writes = 0;
            blobs.FailAt = failAt;
            heads.Fail = failAt == writeCount + 1;
            await Assert.ThrowsAsync<IOException>(async () => await history.CreateVersionAsync("doc", first.Head, after, Metadata("two")));
            Assert.Equal(first.Head, (await history.ReadAsync("doc"))!.Head);
            Assert.Equal(first.Version.Id, Assert.Single((await history.ListVersionsAsync("doc")).Versions).Id);
            Assert.Equal(before, await history.ExportVersionAsync("doc", first.Version.Id));
            blobs.FailAt = 0;
            heads.Fail = false;
            var retry = await history.CreateVersionAsync("doc", first.Head, after, Metadata("retry"));
            Assert.Equal(1, retry.State.Sequence);
        }
    }

    [Fact]
    public async Task InputsAreCapturedBeforeAwaitingHeadStorageAndForeignRecordsAreRejected()
    {
        var heads = new PausingHeads();
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), heads);
        var bytes = DocxSession.CreateBlankDocxBytes();
        var expected = bytes.ToArray();
        var values = new Dictionary<string, string> { ["matter"] = "original" };
        var creation = history.CreateVersionAsync("doc", null, bytes, Metadata("one") with { ApplicationMetadata = values }).AsTask();
        await heads.Started.Task;
        Array.Fill(bytes, (byte)0);
        values["matter"] = "mutated";
        heads.Release.SetResult();
        var version = await creation;
        Assert.Equal(expected, await history.ExportVersionAsync("doc", version.Version.Id));
        Assert.Equal("original", version.Version.Record.Metadata.ApplicationMetadata["matter"]);
        var foreign = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.GetVersionAsync("another", version.Version.Id));
        Assert.Equal(DocxHistoryError.ForeignDocument, foreign.Code);
    }

    [Fact]
    public async Task InvalidCrossRecordTipsAreRejectedOnRead()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var (before, after) = Documents();
        var first = await history.CreateVersionAsync("doc", null, before, Metadata("one"));
        var second = await history.CreateVersionAsync("doc", first.Head, after, Metadata("two"));
        var inconsistent = await new HistoryRecordStore(blobs).SaveStateAsync(second.State with { Version = first.Version.Id });
        await heads.TryAdvanceAsync("doc", second.Head, inconsistent);
        var error = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReadAsync("doc"));
        Assert.Equal(DocxHistoryError.InvalidHistory, error.Code);
        var records = new HistoryRecordStore(blobs);
        var foreignVersion = await records.SaveVersionAsync(second.Version.Record with { DocumentId = "other" });
        var commit = await records.LoadCommitAsync(second.State.Commit!);
        var wrongCommit = await records.SaveCommitAsync(commit with { Version = foreignVersion });
        var wrongState = await records.SaveStateAsync(second.State with { Commit = wrongCommit });
        await heads.TryAdvanceAsync("doc", await heads.ReadAsync("doc"), wrongState);
        var foreign = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReadAsync("doc"));
        Assert.Equal(DocxHistoryError.ForeignDocument, foreign.Code);
    }

    [Fact]
    public async Task CancellationAfterSuccessfulCASDoesNotReportACommittedVersionAsFailed()
    {
        using var cancellation = new CancellationTokenSource();
        var heads = new FaultingHeads { CancelAfterSuccess = cancellation };
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), heads);
        var saved = await history.CreateVersionAsync("doc", null, DocxSession.CreateBlankDocxBytes(), Metadata("one"), cancellation.Token);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(saved.Head, (await history.ReadAsync("doc"))!.Head);
    }

    private static DocxVersionMetadata Metadata(string label) => new()
    {
        Author = "Host actor", CreatedAt = DateTimeOffset.UnixEpoch, Label = label, Message = "saved",
        ApplicationMetadata = new Dictionary<string, string> { ["matter"] = "matter-123" },
    };
    private static (byte[] Before, byte[] After) Documents()
    {
        var before = DocxSession.CreateBlankDocxBytes();
        using var session = new DocxSession(before);
        Assert.True(session.ReplaceText(session.Project().AnchorIndex.Keys.First(), "imported version").Success);
        return (before, session.Save());
    }
    private sealed class FaultingBlobs : IHistoryBlobStore
    {
        private readonly MemoryHistoryBlobStore _inner = new();
        internal int Writes { get; set; }
        internal int FailAt { get; set; }
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken)
        {
            if (++Writes == FailAt) throw new IOException("Injected blob failure.");
            return _inner.PutAsync(reference, content, cancellationToken);
        }
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken) => _inner.OpenReadAsync(reference, cancellationToken);
    }
    private sealed class FaultingHeads : IHistoryHeadStore
    {
        private readonly MemoryHistoryHeadStore _inner = new();
        internal bool Fail { get; set; }
        internal CancellationTokenSource? CancelAfterSuccess { get; init; }
        public ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken) => _inner.ReadAsync(documentId, cancellationToken);
        public async ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected, HistoryBlobReference state, CancellationToken cancellationToken)
        {
            if (Fail) throw new IOException("Injected head failure.");
            var result = await _inner.TryAdvanceAsync(documentId, expected, state, cancellationToken);
            if (result is not null) CancelAfterSuccess?.Cancel();
            return result;
        }
    }
    private sealed class PausingHeads : IHistoryHeadStore
    {
        private readonly MemoryHistoryHeadStore _inner = new();
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public async ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken)
        {
            Started.TrySetResult();
            await Release.Task.WaitAsync(cancellationToken);
            return await _inner.ReadAsync(documentId, cancellationToken);
        }
        public ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected, HistoryBlobReference state, CancellationToken cancellationToken) =>
            _inner.TryAdvanceAsync(documentId, expected, state, cancellationToken);
    }
}
