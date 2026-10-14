// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DocxHistoryReconstructionTests
{
    [Fact]
    public async Task EverySequenceMatchesIndependentlyCapturedContentAcrossRepackImportsAndRestore()
    {
        var directory = Directory.CreateTempSubdirectory("docxodus-history-replay-").FullName;
        try
        {
            var blobs = new FileHistoryBlobStore(Path.Combine(directory, "blobs"));
            var heads = new FileHistoryHeadStore(Path.Combine(directory, "heads"));
            var history = new DocxVersionHistory(blobs, heads);
            var initial = DocxSession.CreateBlankDocxBytes();
            var first = await history.CreateVersionAsync("doc", null, initial, Metadata());
            var namedBytes = EditZip(initial, zip => zip.CreateEntry("packaging/"));
            var named = await history.CreateVersionAsync("doc", first.Head, namedBytes, Metadata());
            Assert.Equal(0, named.State.Sequence);
            var changed = EditZip(Document("one", namedBytes), zip => zip.GetEntry("packaging/")?.Delete());
            var second = await history.CreateVersionAsync("doc", named.Head, changed, Metadata());
            var thirdBytes = Document("two", changed);
            var third = await history.CreateVersionAsync("doc", second.Head, thirdBytes, Metadata());
            var restored = await history.RestoreVersionAsync("doc", third.Head, first.Version.Id, Metadata());
            var lastBytes = Document("after restore", initial);
            var last = await history.CreateVersionAsync("doc", restored.Head, lastBytes, Metadata());
            var expected = new[] { initial, changed, thirdBytes, initial, lastBytes };
            var reopened = new DocxVersionHistory(new FileHistoryBlobStore(Path.Combine(directory, "blobs")),
                new FileHistoryHeadStore(Path.Combine(directory, "heads")));
            for (var sequence = 0; sequence < expected.Length; sequence++)
            {
                var materialized = await reopened.MaterializeAsync("doc", sequence);
                var replayed = await reopened.ReplayAsync("doc", sequence);
                Assert.Equal(Digest(expected[sequence]), Digest(materialized));
                Assert.Equal(Digest(expected[sequence]), Digest(replayed));
            }
            Assert.Equal(last.Head, (await reopened.ReadAsync("doc"))!.Head);
            Assert.Equal(lastBytes, await reopened.ExportVersionAsync("doc", last.Version.Id));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task ReplayUsesRecordedEffectsWithoutReadingIntermediateImportSnapshots()
    {
        var blobs = new FilteringBlobs();
        var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var initial = Document("initial");
        var first = await history.CreateVersionAsync("doc", null, initial, Metadata());
        var expected = Document("changed", initial);
        var second = await history.CreateVersionAsync("doc", first.Head, expected, Metadata());
        blobs.Blocked.Add(second.State.Snapshot.Blob);
        var missing = await Assert.ThrowsAsync<PackageChangeException>(async () => await history.MaterializeAsync("doc", 1));
        Assert.Equal(PackageChangeError.PayloadMissing, missing.Code);
        Assert.Equal(Digest(expected), Digest(await history.ReplayAsync("doc", 1)));
        blobs.Blocked.Clear();
        var commit = await new HistoryRecordStore(blobs).LoadCommitAsync(second.State.Commit!);
        blobs.Blocked.Add(commit.Contribution!);
        Assert.Equal(expected, await history.MaterializeAsync("doc", 1));
        var missingEffects = await Assert.ThrowsAsync<PackageChangeException>(async () => await history.ReplayAsync("doc", 1));
        Assert.Equal(PackageChangeError.PayloadMissing, missingEffects.Code);
    }

    [Fact]
    public async Task ScanBudgetsMissingHistoryAndCancellationFailExplicitlyWithoutPublication()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var unavailable = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.MaterializeAsync("absent", 0));
        Assert.Equal(DocxHistoryError.HistoryUnavailable, unavailable.Code);
        DocxHistoryView? current = null;
        for (var i = 0; i < 4; i++)
            current = await history.CreateVersionAsync("doc", current?.Head, Document($"version {i}"), Metadata());
        var limit = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReplayAsync("doc", 1, maxCommitsToScan: 2));
        Assert.Equal(DocxHistoryError.TraversalLimit, limit.Code);
        var materializeLimit = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.MaterializeAsync("doc", 1, maxCommitsToScan: 2));
        Assert.Equal(DocxHistoryError.TraversalLimit, materializeLimit.Code);
        Assert.NotEmpty(await history.ReplayAsync("doc", 1, maxCommitsToScan: 3));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await history.MaterializeAsync("doc", -1));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await history.MaterializeAsync("doc", 4));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await history.ReplayAsync("doc", 1, cancellationToken: cancelled.Token));
        Assert.Equal(current!.Head, (await history.ReadAsync("doc"))!.Head);
    }

    [Theory]
    [InlineData("initial")]
    [InlineData("epoch")]
    [InlineData("parent_version")]
    [InlineData("effects")]
    public async Task HashValidButInconsistentRecordedGraphsAreRejected(string scenario)
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var records = new HistoryRecordStore(blobs);
        var first = await history.CreateVersionAsync("doc", null, Document("initial"), Metadata());
        var second = await history.CreateVersionAsync("doc", first.Head, Document("changed"), Metadata());
        var state = second.State;
        var commit = await records.LoadCommitAsync(state.Commit!);
        if (scenario == "initial") state = state with { InitialSnapshot = second.State.Snapshot };
        if (scenario == "epoch")
        {
            commit = commit with { Epoch = 1 };
            state = state with { Epoch = 1 };
        }
        if (scenario == "parent_version")
        {
            var wrongParent = await records.SaveVersionAsync(first.Version.Record with { Sequence = 1 });
            var wrongVersion = await records.SaveVersionAsync(second.Version.Record with { Parent = wrongParent });
            commit = commit with { Version = wrongVersion };
            state = state with { Version = wrongVersion };
        }
        if (scenario == "effects")
        {
            var unrelated = PackageChangeSet.Create(Document("unrelated"), Document("other"));
            var bytes = await PackageChangeSetCodec.SaveAsync(unrelated, blobs);
            commit = commit with { Contribution = await HistoryBlobIO.PutBytesAsync(blobs, bytes, default) };
        }
        if (scenario != "initial") state = state with { Commit = await records.SaveCommitAsync(commit) };
        var head = await heads.TryAdvanceAsync("doc", second.Head, await records.SaveStateAsync(state));
        var error = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReplayAsync("doc", 1));
        Assert.Equal(DocxHistoryError.InvalidHistory, error.Code);
        Assert.Equal(head, await heads.ReadAsync("doc"));
    }

    private static DocxVersionMetadata Metadata() => new() { Author = "host", CreatedAt = DateTimeOffset.UnixEpoch };
    private static VerificationDigest Digest(byte[] bytes) => PackageManifestGenerator.Generate(bytes).OrderedOpcContentDigest!;
    private static byte[] Document(string text, byte[]? bytes = null)
    {
        using var session = new DocxSession(bytes ?? DocxSession.CreateBlankDocxBytes());
        Assert.True(session.ReplaceText(session.Project().AnchorIndex.Keys.First(), text).Success);
        return session.Save();
    }
    private static byte[] EditZip(byte[] bytes, Action<ZipArchive> edit)
    {
        using var buffer = new MemoryStream();
        buffer.Write(bytes);
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) edit(zip);
        return buffer.ToArray();
    }
    private sealed class FilteringBlobs : IHistoryBlobStore
    {
        private readonly MemoryHistoryBlobStore _inner = new();
        internal HashSet<HistoryBlobReference> Blocked { get; } = new();
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken) => _inner.PutAsync(reference, content, cancellationToken);
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken) =>
            Blocked.Contains(reference) ? ValueTask.FromResult<Stream?>(null) : _inner.OpenReadAsync(reference, cancellationToken);
    }
}
