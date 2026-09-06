// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class HistoryHeadStoreTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CompareAndSwapIsAtomicAndRevisionPreventsABA(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        Assert.Null(await scope.Store.ReadAsync("doc"));
        var results = await Task.WhenAll(Enumerable.Range(0, 12).Select(_ => Task.Run(async () =>
        {
            var store = filesystem ? new FileHistoryHeadStore(scope.Directory) : scope.Store;
            return await store.TryAdvanceAsync("doc", null, State('a'));
        })));
        var first = Assert.Single(results.Where(head => head is not null))!;
        Assert.Equal(1, first.Revision);
        Assert.Equal(first, await scope.Store.ReadAsync("doc"));
        var second = await scope.Store.TryAdvanceAsync("doc", first, State('b'));
        Assert.NotNull(second);
        var third = await scope.Store.TryAdvanceAsync("doc", second, State('a'));
        Assert.NotNull(third);
        Assert.Equal(3, third!.Revision);
        Assert.Null(await scope.Store.TryAdvanceAsync("doc", first, State('c')));
        Assert.Equal(third, await scope.Store.ReadAsync("doc"));
        var same = await scope.Store.TryAdvanceAsync("doc", third, State('a'));
        Assert.Equal(4, same!.Revision);
        if (filesystem) Assert.Equal(same, await new FileHistoryHeadStore(scope.Directory).ReadAsync("doc"));
        Assert.Empty(Directory.GetFiles(scope.Directory, "*.tmp"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DocumentKeysAreIsolatedAndNeverTreatedAsPaths(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        foreach (var id in new[] { "../elsewhere", "/absolute", "document", "DOCUMENT", "合同 📝" })
        {
            Assert.Null(await scope.Store.ReadAsync(id));
            Assert.Equal(1, (await scope.Store.TryAdvanceAsync(id, null, State('a')))!.Revision);
        }
        if (filesystem)
        {
            Assert.Equal(5, Directory.GetFiles(scope.Directory, "*.head").Length);
            Assert.All(Directory.GetFiles(scope.Directory), path => Assert.Equal(69, Path.GetFileName(path).Length));
        }
        await Assert.ThrowsAnyAsync<ArgumentException>(async () => await scope.Store.ReadAsync("\ud800"));
        await Assert.ThrowsAnyAsync<ArgumentException>(async () => await scope.Store.ReadAsync(new string('x', 1025)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InvalidExpectationsReferencesAndCancellationCannotAdvance(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await scope.Store.TryAdvanceAsync("doc", null, State('a'), cancellation.Token));
        await Assert.ThrowsAnyAsync<ArgumentException>(async () =>
            await scope.Store.TryAdvanceAsync("doc", new HistoryHead(0, State('a')), State('b')));
        await Assert.ThrowsAnyAsync<ArgumentException>(async () =>
            await scope.Store.TryAdvanceAsync("doc", null, State('a') with { Length = -1 }));
        await Assert.ThrowsAsync<OverflowException>(async () =>
            await scope.Store.TryAdvanceAsync("doc", new HistoryHead(long.MaxValue, State('a')), State('b')));
        Assert.Null(await scope.Store.ReadAsync("doc"));
        Assert.Empty(Directory.GetFiles(scope.Directory));
    }

    [Fact]
    public async Task FilesystemLockWaitIsCancellableAndLockFileCanBeReusedAfterRelease()
    {
        using var scope = new StoreScope(filesystem: true);
        var path = Path.Combine(scope.Directory, HistoryHeadCodec.Key("doc") + ".lock");
        using (var held = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            using var cancellation = new CancellationTokenSource();
            var pending = scope.Store.TryAdvanceAsync("doc", null, State('a'), cancellation.Token).AsTask();
            Assert.False(pending.IsCompleted);
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => pending);
            Assert.Null(await scope.Store.ReadAsync("doc"));
        }
        Assert.NotNull(await scope.Store.TryAdvanceAsync("doc", null, State('a')));
    }

    [Theory]
    [InlineData("truncated")]
    [InlineData("duplicate")]
    [InlineData("version")]
    [InlineData("oversized")]
    [InlineData("invalid_digest")]
    public async Task CorruptFilesystemHeadsAreRejectedNotOverwritten(string scenario)
    {
        using var scope = new StoreScope(filesystem: true);
        var head = (await scope.Store.TryAdvanceAsync("doc", null, State('a')))!;
        var path = Path.Combine(scope.Directory, HistoryHeadCodec.Key("doc") + ".head");
        var original = Encoding.UTF8.GetString(HistoryHeadCodec.Encode(head));
        var corrupt = scenario switch
        {
            "truncated" => original[..^1],
            "duplicate" => original.Replace("\"schemaVersion\":1", "\"schemaVersion\":1,\"schemaVersion\":1"),
            "version" => original.Replace("\"schemaVersion\":1", "\"schemaVersion\":2"),
            "oversized" => new string(' ', 1025),
            _ => original.Replace(new string('a', 64), "bad"),
        };
        await File.WriteAllTextAsync(path, corrupt);
        await Assert.ThrowsAsync<InvalidDataException>(async () => await scope.Store.ReadAsync("doc"));
        await Assert.ThrowsAsync<InvalidDataException>(async () => await scope.Store.TryAdvanceAsync("doc", head, State('b')));
        Assert.Equal(corrupt, await File.ReadAllTextAsync(path));
    }

    private static HistoryBlobReference State(char digit) => new(new VerificationDigest
    {
        Algorithm = "SHA-256", Value = new string(digit, 64),
    }, 100);

    private sealed class StoreScope : IDisposable
    {
        internal string Directory { get; } = System.IO.Directory.CreateTempSubdirectory("docxodus-history-heads-").FullName;
        internal IHistoryHeadStore Store { get; }
        internal StoreScope(bool filesystem) => Store = filesystem ? new FileHistoryHeadStore(Directory) : new MemoryHistoryHeadStore();
        public void Dispose() => System.IO.Directory.Delete(Directory, recursive: true);
    }
}
