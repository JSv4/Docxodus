// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Docxodus.History;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxHistoryImportTests
{
    [Theory]
    [InlineData("agreement")]
    [InlineData("charter-collaboration")]
    public async Task RealArchiveImportsReopensContinuesAndRestoresWithoutChangingOriginalFiles(string name)
    {
        var index = await DocxHistoryArchiveArtifactTests.ReadIndexAsync(name);
        var bytes = await Artifact(name + ".docxhistory");
        var root = Directory.CreateTempSubdirectory("history-import-files-").FullName;
        try
        {
            DocxVersionHistory Reopen() => new(new FileHistoryBlobStore(Path.Combine(root, "blobs")),
                new FileHistoryHeadStore(Path.Combine(root, "heads")));
            var imported = await Reopen().ImportHistoryArchiveAsync(bytes);
            Assert.False(imported.AlreadyPresent); Assert.Equal(index.Head, imported.View.Head);
            Assert.True((await Reopen().ImportHistoryArchiveAsync(bytes)).AlreadyPresent);
            var history = Reopen(); var doc = history.Document(index.DocumentId);
            Assert.Equal(index.Head, (await doc.ReadAsync())!.Head);
            foreach (var version in index.Versions)
                Assert.Equal(await Artifact(version.File), await doc.ExportDocxAsync(version.Id));
            if (index.Conflict is { } conflict)
                Assert.Equal(await Artifact(index.ProposalFile!), await doc.ExportOperationProposalAsync(conflict));
            else
            {
                // Original retry receipt still returns publication 1, not the imported/latest head.
                var retry = await doc.CreateVersionAsync("initial", null, await Artifact(index.Versions[0].File),
                    HistoryArchiveFixture.Metadata("Original agreement"));
                Assert.Equal(1, retry.Head.Revision); Assert.Equal(index.Versions[0].Id, retry.Version.Id);
                Assert.Equal(index.Head, (await doc.ReadAsync())!.Head);
            }
            var latest = await doc.ExportDocxAsync();
            var updated = HistoryArchiveFixture.Edit(latest, " Continued after portable import.");
            var saved = await doc.CreateVersionAsync("continued", imported.View.Head, updated, HistoryArchiveFixture.Metadata("Continued"));
            Assert.Equal(index.Head.Revision + 1, saved.Head.Revision);
            var restored = await Reopen().Document(index.DocumentId).RestoreVersionAsync("restore-after-import", saved.Head,
                index.Versions[0].Id, HistoryArchiveFixture.Metadata("Restore first"));
            Assert.Equal(saved.Head.Revision + 1, restored.Head.Revision);
            Assert.Equal(await Artifact(index.Versions[0].File), await Reopen().Document(index.DocumentId).ExportDocxAsync());
            Assert.Equal(DocxHistoryError.ImportConflict, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
                await Reopen().ImportHistoryArchiveAsync(bytes))).Code);
            Assert.Equal(restored.Head, (await doc.ReadAsync())!.Head);
            var outputPath = Path.Combine(root, "continued.docxhistory");
            await using (var output = File.Create(outputPath)) await doc.ExportHistoryArchiveAsync(output);
            using var outputFile = File.OpenRead(outputPath);
            using var reopened = await DocxHistoryArchive.OpenAsync(outputFile);
            Assert.Equal(restored.Head, reopened.View.Head);
            Assert.Equal(updated, await reopened.ExportDocxAsync(saved.Version.Id));
            var compared = await reopened.CompareVersionsAsync(index.Versions[0].Id, saved.Version.Id,
                new DocxDiffSettings { PreAcceptInputRevisions = true, PreserveInputRevisions = true });
            var comparisonPath = Path.Combine(root, "selected-comparison.docx");
            await File.WriteAllBytesAsync(comparisonPath, compared.ToRedline().DocumentByteArray);
            using var comparison = new DocxSession(await File.ReadAllBytesAsync(comparisonPath));
            Assert.NotEmpty(compared.GetRevisions()); Assert.NotNull(comparison);
            Assert.Equal(bytes, await Artifact(name + ".docxhistory"));
        }
        finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task InvalidArchivesUnsupportedStoresAndDifferentHeadsCannotWriteDestinationBlobs()
    {
        var bytes = await Artifact("agreement.docxhistory");
        using var target = new HistoryFaultHarness();
        var history = new DocxVersionHistory(target, target);
        await Assert.ThrowsAsync<PackageChangeException>(async () => await history.ImportHistoryArchiveAsync(bytes[..^1]));
        Assert.Equal(0, target.Writes); Assert.Equal(0, target.Publications);
        var unsupported = new DocxVersionHistory(target, new LegacyHeads(target));
        Assert.Equal(DocxHistoryError.InitializationUnsupported, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await unsupported.ImportHistoryArchiveAsync(bytes))).Code);
        Assert.Equal(0, target.Writes); Assert.Equal(0, target.Publications);
        var index = await DocxHistoryArchiveArtifactTests.ReadIndexAsync("agreement");
        var existing = await history.Document(index.DocumentId).CreateVersionAsync("local", null,
            await Artifact(index.Versions[0].File), HistoryArchiveFixture.Metadata("Local unrelated history"));
        target.Arm(HistoryFaultHarness.Fault.None);
        Assert.Equal(DocxHistoryError.ImportConflict, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.ImportHistoryArchiveAsync(bytes))).Code);
        Assert.Equal(0, target.Writes); Assert.Equal(0, target.Publications);
        Assert.Equal(existing.Head, (await history.ReadAsync(index.DocumentId))!.Head);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EveryBlobFailureLeavesNoHeadAndIdenticalRetryRecovers(bool filesystem)
    {
        // Exhaustive boundaries use the real agreement archive, not a generated empty DOCX.
        var bytes = await Artifact("agreement.docxhistory");
        using var archive = await DocxHistoryArchive.OpenAsync(bytes);
        foreach (var fault in new[] { HistoryFaultHarness.Fault.BeforeBlob, HistoryFaultHarness.Fault.AfterBlob })
            for (var at = 1; at <= archive.Info.BlobCount; at++)
            {
                using var target = new HistoryFaultHarness(filesystem); target.Arm(fault, at);
                var history = new DocxVersionHistory(target, target);
                await Assert.ThrowsAsync<IOException>(async () => await history.ImportHistoryArchiveAsync(bytes));
                Assert.Null(await target.ReadAsync(archive.DocumentId)); Assert.Equal(0, target.Publications);
                target.Arm(HistoryFaultHarness.Fault.None);
                Assert.Equal(archive.View.Head, (await history.ImportHistoryArchiveAsync(bytes)).View.Head);
                Assert.Equal(await archive.ExportDocxAsync(), await history.Document(archive.DocumentId).ExportDocxAsync());
            }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PublicationFailureAndCancellationRespectTheSingleCommitPoint(bool filesystem)
    {
        var bytes = await Artifact("agreement.docxhistory");
        using var archive = await DocxHistoryArchive.OpenAsync(bytes);
        foreach (var fault in new[] { HistoryFaultHarness.Fault.BeforeHead, HistoryFaultHarness.Fault.AfterHead,
            HistoryFaultHarness.Fault.CancelBeforeHead, HistoryFaultHarness.Fault.CancelAfterHead })
        {
            using var target = new HistoryFaultHarness(filesystem) { Cancellation = new() }; target.Arm(fault);
            var history = new DocxVersionHistory(target, target);
            var failure = await Record.ExceptionAsync(async () => await history.ImportHistoryArchiveAsync(bytes,
                cancellationToken: target.Cancellation.Token));
            if (fault == HistoryFaultHarness.Fault.CancelAfterHead) Assert.Null(failure);
            else Assert.NotNull(failure);
            var committed = fault is HistoryFaultHarness.Fault.AfterHead or HistoryFaultHarness.Fault.CancelAfterHead;
            Assert.Equal(committed ? archive.View.Head : null, await target.ReadAsync(archive.DocumentId));
            target.Arm(HistoryFaultHarness.Fault.None);
            var retry = await history.ImportHistoryArchiveAsync(bytes);
            Assert.Equal(committed, retry.AlreadyPresent); Assert.Equal(archive.View.Head, retry.View.Head);
        }
    }

    [Fact]
    public async Task ExactRetryRepairsMissingBlobsButCannotHideCorruptImmutableCollisions()
    {
        var bytes = await Artifact("agreement.docxhistory");
        var root = Directory.CreateTempSubdirectory("history-import-repair-").FullName;
        try
        {
            var blobDirectory = Path.Combine(root, "blobs");
            var history = new DocxVersionHistory(new FileHistoryBlobStore(blobDirectory),
                new FileHistoryHeadStore(Path.Combine(root, "heads")));
            var imported = await history.ImportHistoryArchiveAsync(bytes);
            var snapshot = imported.View.State.Snapshot.Blob;
            var path = Path.Combine(blobDirectory, snapshot.Digest.Value + ".blob");
            var original = await File.ReadAllBytesAsync(path);
            File.Delete(path); // Only this test's private imported copy.
            Assert.True((await history.ImportHistoryArchiveAsync(bytes)).AlreadyPresent);
            Assert.Equal(original, await File.ReadAllBytesAsync(path));
            var corrupt = original.ToArray(); corrupt[0] ^= 1; await File.WriteAllBytesAsync(path, corrupt);
            Assert.Equal(PackageChangeError.PayloadMismatch, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await history.ImportHistoryArchiveAsync(bytes))).Code);
            Assert.Equal(corrupt, await File.ReadAllBytesAsync(path));
            Assert.Equal(imported.View.Head, (await history.ReadAsync(imported.Archive.DocumentId))!.Head);
        }
        finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task SimultaneousIdenticalImportsInitializeOnlyOnce()
    {
        var bytes = await Artifact("agreement.docxhistory");
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var results = await Task.WhenAll(Enumerable.Range(0, 4).Select(_ => Task.Run(async () => await history.ImportHistoryArchiveAsync(bytes))));
        Assert.Single(results, r => !r.AlreadyPresent);
        Assert.All(results, r => Assert.Equal(results[0].View.Head, r.View.Head));
    }

    [Fact]
    public async Task ConcurrentLocalPublicationWinsOverImportAndOwnedInputClosesBeforeInitialization()
    {
        var bytes = await Artifact("agreement.docxhistory"); var index = await DocxHistoryArchiveArtifactTests.ReadIndexAsync("agreement");
        var blobs = new MemoryHistoryBlobStore(); var heads = new MemoryHistoryHeadStore();
        var gate = new GatedInitializer(heads);
        var history = new DocxVersionHistory(blobs, gate);
        var input = new MemoryStream(bytes);
        var pending = history.ImportHistoryArchiveAsync(input, leaveOpen: false).AsTask();
        await gate.Arrived.Task.WaitAsync(TimeSpan.FromSeconds(30));
        Assert.False(input.CanRead);
        var local = await new DocxVersionHistory(blobs, heads).Document(index.DocumentId).CreateVersionAsync("local", null,
            await Artifact(index.Versions[0].File), HistoryArchiveFixture.Metadata("Local"));
        gate.Release.SetResult();
        Assert.Equal(DocxHistoryError.ImportConflict, (await Assert.ThrowsAsync<DocxHistoryException>(() => pending)).Code);
        Assert.Equal(local.Head, await heads.ReadAsync(index.DocumentId));
    }

    private static Task<byte[]> Artifact(string file) => File.ReadAllBytesAsync(Path.Combine(DocxHistoryArchiveArtifactTests.Root, file));

    private sealed class LegacyHeads(IHistoryHeadStore inner) : IHistoryHeadStore
    {
        public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default) => inner.ReadAsync(id, cancellationToken);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => inner.TryAdvanceAsync(id, expected, state, cancellationToken);
    }

    private sealed class GatedInitializer(IHistoryHeadInitializer inner) : IHistoryHeadInitializer
    {
        internal TaskCompletionSource Arrived { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default) => inner.ReadAsync(id, cancellationToken);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => inner.TryAdvanceAsync(id, expected, state, cancellationToken);
        public async ValueTask<HistoryHeadInitializationResult> TryInitializeAsync(string id, HistoryHead head,
            CancellationToken cancellationToken = default)
        {
            Arrived.SetResult(); await Release.Task.WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
            return await inner.TryInitializeAsync(id, head, cancellationToken);
        }
    }
}
