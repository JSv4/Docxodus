// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class HistoryBlobStoreTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WritesAreImmutableIdempotentAndCallerStreamsRemainOpen(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        var bytes = new byte[] { 1, 2, 3 };
        var reference = Reference(bytes);
        Assert.Null(await scope.Store.OpenReadAsync(reference));
        using var source = new MemoryStream(bytes);
        await scope.Store.PutAsync(reference, source);
        Assert.True(source.CanRead);
        source.Position = 0;
        await scope.Store.PutAsync(reference, source);
        bytes[0] = 99;
        using var read = await scope.Store.OpenReadAsync(reference);
        Assert.NotNull(read);
        Assert.False(read!.CanWrite);
        if (read is MemoryStream memory) Assert.Throws<UnauthorizedAccessException>(() => memory.GetBuffer());
        Assert.Equal(new byte[] { 1, 2, 3 }, await ReadAll(read));
        await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await scope.Store.OpenReadAsync(reference with { Length = 2 }));
        Assert.Empty(Directory.GetFiles(scope.Directory, "*.tmp"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InvalidOrCancelledWritesNeverPublishAndCannotOverwrite(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        var bytes = new byte[] { 1, 2, 3 };
        var reference = Reference(bytes);
        foreach (var invalid in new[] { new byte[] { 1, 2 }, new byte[] { 1, 2, 3, 4 }, new byte[] { 3, 2, 1 } })
        {
            using var content = new MemoryStream(invalid);
            var error = await Assert.ThrowsAsync<PackageChangeException>(async () => await scope.Store.PutAsync(reference, content));
            Assert.Equal(PackageChangeError.PayloadMismatch, error.Code);
            Assert.Null(await scope.Store.OpenReadAsync(reference));
            Assert.True(content.CanRead);
        }
        using var cancellation = new CancellationTokenSource();
        using var cancelled = new CancellingStream(bytes, cancellation);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await scope.Store.PutAsync(reference, cancelled, cancellation.Token));
        Assert.Null(await scope.Store.OpenReadAsync(reference));
        using var good = new MemoryStream(bytes);
        await scope.Store.PutAsync(reference, good);
        using var bad = new MemoryStream(new byte[] { 9, 9, 9 });
        await Assert.ThrowsAsync<PackageChangeException>(async () => await scope.Store.PutAsync(reference, bad));
        using var read = await scope.Store.OpenReadAsync(reference);
        Assert.Equal(bytes, await ReadAll(read!));
        Assert.Empty(Directory.GetFiles(scope.Directory, "*.tmp"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MalformedReferencesAndLimitsFailBeforeReadingOrCreatingFiles(bool filesystem)
    {
        using var scope = new StoreScope(filesystem, maxBlobBytes: 4);
        var reference = Reference(new byte[] { 1 });
        foreach (var invalid in new[]
        {
            reference with { Digest = reference.Digest with { Value = "../escape" } },
            reference with { Digest = reference.Digest with { Value = new string('A', 64) } },
            reference with { Digest = reference.Digest with { Algorithm = "MD5" } },
            reference with { Length = -1 },
        })
        {
            using var content = new MemoryStream(new byte[] { 1 });
            await Assert.ThrowsAsync<ArgumentException>(async () => await scope.Store.PutAsync(invalid, content));
            await Assert.ThrowsAsync<ArgumentException>(async () => await scope.Store.OpenReadAsync(invalid));
            Assert.Equal(0, content.Position);
        }
        using var oversized = new MemoryStream(new byte[5]);
        var limit = await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await scope.Store.PutAsync(Reference(new byte[5]), oversized));
        Assert.Equal(PackageChangeError.ResourceLimit, limit.Code);
        Assert.Equal(0, oversized.Position);
        Assert.Empty(Directory.GetFiles(scope.Directory));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task IndependentConcurrentWritesAndEmptyBlobsAreSupported(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        var bytes = Enumerable.Range(0, 100_000).Select(i => (byte)i).ToArray();
        var reference = Reference(bytes);
        await Task.WhenAll(Enumerable.Range(0, 12).Select(_ => Task.Run(async () =>
        {
            IHistoryBlobStore store = filesystem ? new FileHistoryBlobStore(scope.Directory) : scope.Store;
            using var stream = new MemoryStream(bytes);
            await store.PutAsync(reference, stream);
        })));
        using var read = await scope.Store.OpenReadAsync(reference);
        Assert.Equal(bytes, await ReadAll(read!));
        using var empty = new MemoryStream();
        await scope.Store.PutAsync(Reference([]), empty);
        using var emptyRead = await scope.Store.OpenReadAsync(Reference([]));
        Assert.Empty(await ReadAll(emptyRead!));
        if (filesystem) Assert.Equal(2, Directory.GetFiles(scope.Directory).Length);
    }

    [Fact]
    public async Task FilesystemCodecReopensAndCorruptionIsNeitherHiddenNorRepaired()
    {
        using var scope = new StoreScope(filesystem: true);
        var before = DocxSession.CreateBlankDocxBytes();
        using var session = new DocxSession(before);
        Assert.True(session.ReplaceText(session.Project().AnchorIndex.Keys.First(), "persisted").Success);
        var after = session.Save();
        var changes = PackageChangeSet.Create(before, after);
        var manifest = await PackageChangeSetCodec.SaveAsync(changes, scope.Store);
        var reopened = new FileHistoryBlobStore(scope.Directory);
        var loaded = await PackageChangeSetCodec.LoadAsync(manifest, reopened);
        Assert.Equal(changes.AfterDigest, PackageManifestGenerator.Generate(loaded.Apply(before)).OrderedOpcContentDigest);
        Assert.Equal(changes.BeforeDigest, PackageManifestGenerator.Generate(loaded.Invert().Apply(after)).OrderedOpcContentDigest);

        var digest = changes.PayloadDigests.First();
        var good = changes.GetPayload(digest);
        var corrupt = good.ToArray();
        corrupt[0] ^= 1;
        var path = Path.Combine(scope.Directory, digest.Value + ".blob");
        await File.WriteAllBytesAsync(path, corrupt);
        var error = await Assert.ThrowsAsync<PackageChangeException>(async () => await PackageChangeSetCodec.LoadAsync(manifest, reopened));
        Assert.Equal(PackageChangeError.PayloadMismatch, error.Code);
        using var content = new MemoryStream(good);
        await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await reopened.PutAsync(new HistoryBlobReference(digest, good.Length), content));
        Assert.Equal(corrupt, await File.ReadAllBytesAsync(path));
        Assert.Empty(Directory.GetFiles(scope.Directory, "*.tmp"));
    }

    private static HistoryBlobReference Reference(byte[] bytes) => new(new VerificationDigest
    {
        Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    }, bytes.Length);

    private static async Task<byte[]> ReadAll(Stream stream)
    {
        using var buffer = new MemoryStream();
        await stream.CopyToAsync(buffer);
        return buffer.ToArray();
    }

    private sealed class StoreScope : IDisposable
    {
        internal string Directory { get; } = System.IO.Directory.CreateTempSubdirectory("docxodus-history-blobs-").FullName;
        internal IHistoryBlobStore Store { get; }

        internal StoreScope(bool filesystem, int maxBlobBytes = 256 * 1024 * 1024) =>
            Store = filesystem ? new FileHistoryBlobStore(Directory, maxBlobBytes) : new MemoryHistoryBlobStore(maxBlobBytes);

        public void Dispose() => System.IO.Directory.Delete(Directory, recursive: true);
    }

    private sealed class CancellingStream(byte[] bytes, CancellationTokenSource cancellation) : MemoryStream(bytes)
    {
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            var result = base.ReadAsync(buffer, cancellationToken);
            cancellation.Cancel();
            return result;
        }
    }
}
