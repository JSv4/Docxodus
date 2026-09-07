// Copyright (c) John Scrudato IV. All rights reserved.
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
    public async Task ExactInitializationIsAbsentOnlyAndSharesTheOrdinaryPublicationLock(bool filesystem)
    {
        using var scope = new StoreScope(filesystem);
        var initializer = (IHistoryHeadInitializer)scope.Store;
        var imported = new HistoryHead(87, State('a'));
        var results = await Task.WhenAll(Enumerable.Range(0, 12).Select(_ => Task.Run(async () =>
            await (filesystem ? new FileHistoryHeadStore(scope.Directory) : initializer).TryInitializeAsync("doc", imported))));
        Assert.Single(results, r => r.Initialized);
        Assert.All(results, r => Assert.Equal(imported, r.Head));
        var other = await initializer.TryInitializeAsync("doc", new HistoryHead(1000, State('b')));
        Assert.False(other.Initialized); Assert.Equal(imported, other.Head);
        var advanced = await scope.Store.TryAdvanceAsync("doc", imported, State('b'));
        Assert.Equal(88, advanced!.Revision);
        Assert.Equal(advanced, (await initializer.TryInitializeAsync("doc", imported)).Head);
        Assert.Equal(advanced, await scope.Store.ReadAsync("doc"));
        await Assert.ThrowsAnyAsync<ArgumentException>(async () =>
            await initializer.TryInitializeAsync("other", imported with { Revision = 0 }));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await initializer.TryInitializeAsync("other", imported, new CancellationToken(true)));
        Assert.Null(await scope.Store.ReadAsync("other"));
        Assert.Empty(Directory.GetFiles(scope.Directory, "*.tmp"));

        // Ordinary absent CAS and imported revision compete on the same key.
        var raced = await Task.WhenAll(Task.Run(async () =>
            (HistoryHead?)(await initializer.TryInitializeAsync("race", imported)).Head), Task.Run(async () =>
            await scope.Store.TryAdvanceAsync("race", null, State('c'))));
        var winner = await scope.Store.ReadAsync("race");
        Assert.True(winner == imported || winner == new HistoryHead(1, State('c')));
        Assert.All(raced.Where(r => r is not null), r => Assert.Equal(winner, r));
    }

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
