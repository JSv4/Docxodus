// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Text.Json;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DocxSnapshotStoreTests
{
    [Fact]
    public async Task ExactBytesAndIndependentContentIdentitySurviveRepackAndReopen()
    {
        var directory = Directory.CreateTempSubdirectory("docxodus-snapshots-").FullName;
        try
        {
            var bytes = DocxSession.CreateBlankDocxBytes();
            using var buffer = new MemoryStream();
            buffer.Write(bytes);
            using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
                foreach (var entry in zip.Entries) entry.LastWriteTime = new DateTimeOffset(2001, 2, 3, 4, 5, 6, TimeSpan.Zero);
            var repacked = buffer.ToArray();
            var store = new DocxSnapshotStore(new FileHistoryBlobStore(directory));
            var first = await store.CaptureAsync(bytes);
            var second = await store.CaptureAsync(repacked);
            Assert.Equal(first.ContentDigest, second.ContentDigest);
            Assert.NotEqual(first.Blob.Digest, second.Blob.Digest);
            Assert.Equal(first, await store.CaptureAsync(bytes));
            var reference = JsonSerializer.Deserialize<DocxSnapshotReference>(JsonSerializer.Serialize(second))!;
            var reopened = new DocxSnapshotStore(new FileHistoryBlobStore(directory));
            Assert.Equal(bytes, await reopened.ExportAsync(first));
            Assert.Equal(repacked, await reopened.ExportAsync(reference));
            var exported = await reopened.ExportAsync(first);
            exported[0] ^= 1;
            Assert.Equal(bytes, await reopened.ExportAsync(first));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task CaptureOwnsBytesBeforeAwaitingHostStorage()
    {
        var blobs = new RecordingStore { PauseWrites = true };
        var store = new DocxSnapshotStore(blobs);
        var bytes = DocxSession.CreateBlankDocxBytes();
        var expected = bytes.ToArray();
        var capture = store.CaptureAsync(bytes).AsTask();
        await blobs.WriteStarted.Task;
        Array.Fill(bytes, (byte)0);
        blobs.ReleaseWrite.SetResult();
        var reference = await capture;
        Assert.Equal(expected, await store.ExportAsync(reference));
    }

    [Fact]
    public async Task BoundsInvalidPackagesAndInvalidReferencesFailBeforeStorageIO()
    {
        var blobs = new RecordingStore();
        var bytes = DocxSession.CreateBlankDocxBytes();
        var store = new DocxSnapshotStore(blobs);
        var invalid = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.CaptureAsync([1, 2, 3]));
        Assert.Equal(PackageChangeError.InvalidPackage, invalid.Code);
        var limited = new DocxSnapshotStore(blobs, maxSnapshotBytes: bytes.Length - 1);
        var oversized = await Assert.ThrowsAsync<PackageChangeException>(async () => await limited.CaptureAsync(bytes));
        Assert.Equal(PackageChangeError.ResourceLimit, oversized.Code);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await store.CaptureAsync(bytes, cancellation.Token));
        Assert.Equal(0, blobs.Writes);
        var captured = await store.CaptureAsync(bytes);
        await Assert.ThrowsAsync<ArgumentException>(async () => await store.ExportAsync(captured with
        {
            ContentDigest = captured.ContentDigest with { Value = "../invalid" },
        }));
        await Assert.ThrowsAsync<PackageChangeException>(async () => await limited.ExportAsync(captured));
        Assert.Equal(0, blobs.Reads);
    }

    [Fact]
    public async Task MissingBytesAndMismatchedContentIdentityAreDetected()
    {
        var store = new DocxSnapshotStore(new MemoryHistoryBlobStore());
        var reference = await store.CaptureAsync(DocxSession.CreateBlankDocxBytes());
        var missing = new DocxSnapshotStore(new MemoryHistoryBlobStore());
        var absent = await Assert.ThrowsAsync<PackageChangeException>(async () => await missing.ExportAsync(reference));
        Assert.Equal(PackageChangeError.PayloadMissing, absent.Code);
        var mismatch = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.ExportAsync(reference with
        {
            ContentDigest = reference.ContentDigest with { Value = new string('0', 64) },
        }));
        Assert.Equal(PackageChangeError.PayloadMismatch, mismatch.Code);
    }

    [Fact]
    public async Task ComparisonReusesExistingSemanticAndNativeRedlineProducts()
    {
        var before = DocxSession.CreateBlankDocxBytes();
        using var session = new DocxSession(before);
        Assert.True(session.ReplaceText(session.Project().AnchorIndex.Keys.First(), "version two").Success);
        var after = session.Save();
        var store = new DocxSnapshotStore(new MemoryHistoryBlobStore());
        var left = await store.CaptureAsync(before);
        var right = await store.CaptureAsync(after);
        var settings = new DocxDiffSettings { AuthorForRevisions = "History comparison" };
        var comparison = await store.CompareAsync(left, right, settings);
        var expected = DocxDiff.CreateComparison(new WmlDocument("before.docx", before), new WmlDocument("after.docx", after), settings);
        Assert.Equal(expected.GetSemanticChangesJson(), comparison.GetSemanticChangesJson());
        Assert.Equal(expected.ToRedline().DocumentByteArray, comparison.ToRedline().DocumentByteArray);
        Assert.Equal(before, await store.ExportAsync(left));
        Assert.Equal(after, await store.ExportAsync(right));
    }

    private sealed class RecordingStore : IHistoryBlobStore
    {
        private readonly MemoryHistoryBlobStore _inner = new();
        internal bool PauseWrites { get; init; }
        internal int Writes { get; private set; }
        internal int Reads { get; private set; }
        internal TaskCompletionSource WriteStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource ReleaseWrite { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public async ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken)
        {
            Writes++;
            WriteStarted.TrySetResult();
            if (PauseWrites) await ReleaseWrite.Task.WaitAsync(cancellationToken);
            await _inner.PutAsync(reference, content, cancellationToken);
        }

        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken)
        {
            Reads++;
            return _inner.OpenReadAsync(reference, cancellationToken);
        }
    }
}
