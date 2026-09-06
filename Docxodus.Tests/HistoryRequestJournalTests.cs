#nullable enable

using System.Security.Cryptography;
using System.Text;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public sealed class HistoryRequestJournalTests
{
    [Theory]
    [InlineData(731)]
    [InlineData(982451653)]
    public async Task CompressedIndexFindsEveryOriginalResultAcrossRandomInsertionAndRestart(int seed)
    {
        var random = new Random(seed);
        var blobs = new CountingBlobs();
        var store = new HistoryRequestJournalStore(blobs);
        HistoryRequestJournal? journal = null;
        HistoryHead? head = null;
        var receipts = new List<(HistoryRequestIdentity Request, HistoryHead Head)>();
        var ids = Enumerable.Range(0, 1500).OrderBy(_ => random.Next()).ToArray();
        foreach (var id in ids)
        {
            var request = Identity("replica/合同/" + id);
            Assert.Null(await store.FindAsync("doc", head, journal, request));
            journal = await store.AdvanceAsync("doc", head, journal, request);
            head = Head(receipts.Count + 1);
            receipts.Add((request, head));
            var old = receipts[random.Next(receipts.Count)];
            Assert.Equal(old.Head, await new HistoryRequestJournalStore(blobs).FindAsync("doc", head, journal, old.Request));
        }
        foreach (var (request, original) in receipts.OrderBy(_ => random.Next()))
        {
            blobs.Reads = 0;
            Assert.Equal(original, await store.FindAsync("doc", head, journal, request));
            Assert.InRange(blobs.Reads, 0, 257); // Hash width is the hard bound, not history size.
        }
        Assert.Null(await store.FindAsync("doc", head, journal, Identity("absent")));
        Assert.True(blobs.MaximumLength <= HistoryRequestJournalStore.MaxNodeBytes);
        var before = blobs.Writes;
        foreach (var receipt in receipts.Take(20))
            await store.FindAsync("doc", head, journal, receipt.Request);
        Assert.Equal(before, blobs.Writes); // Queries never publish/repair indexes.
    }

    [Fact]
    public async Task LegacyPublicationsPromoteReceiptsAndConflictsDoNotBecomeMisses()
    {
        var store = new HistoryRequestJournalStore(new MemoryHistoryBlobStore());
        var request = Identity("r1");
        var first = await store.AdvanceAsync("doc", null, null, request);
        var second = await store.AdvanceAsync("doc", Head(1), first, null);
        Assert.Null(second!.Current);
        Assert.Equal(Head(1), await store.FindAsync("doc", Head(2), second, request));
        foreach (var (head, journal) in new[] { (Head(1), first), (Head(2), second) })
        {
            var changed = request with { Fingerprint = Digest("changed") };
            var error = await Assert.ThrowsAsync<DocxHistoryException>(async () => await store.FindAsync("doc", head, journal, changed));
            Assert.Equal(DocxHistoryError.RequestConflict, error.Code);
            await Assert.ThrowsAsync<DocxHistoryException>(async () => await store.FindAsync("foreign", head, journal, request));
            await Assert.ThrowsAsync<DocxHistoryException>(async () => await store.FindAsync("doc", Head(999), journal, request));
        }
    }

    [Fact]
    public async Task MissingCorruptAndCanceledIndexLookupsFailClosed()
    {
        var blobs = new CountingBlobs();
        var store = new HistoryRequestJournalStore(blobs);
        var first = await store.AdvanceAsync("doc", null, null, Identity("r1"));
        var second = await store.AdvanceAsync("doc", Head(1), first, Identity("r2"));
        blobs.Missing = true;
        Assert.Equal(PackageChangeError.PayloadMissing, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await store.FindAsync("doc", Head(2), second, Identity("r1")))).Code);
        blobs.Missing = false; blobs.Corrupt = true;
        Assert.Equal(PackageChangeError.PayloadMismatch, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await store.FindAsync("doc", Head(2), second, Identity("r1")))).Code);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await store.FindAsync("doc", Head(2), second, Identity("r2"), canceled.Token));
    }

    internal static HistoryRequestIdentity Identity(string id) => new(id, Digest("input:" + id));
    internal static VerificationDigest Digest(string text) => new()
    { Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant() };
    internal static HistoryHead Head(long revision) => new(revision, new HistoryBlobReference(Digest("state:" + revision), 100));

    private sealed class CountingBlobs : IHistoryBlobStore
    {
        private readonly MemoryHistoryBlobStore _inner = new();
        internal int Reads, Writes, MaximumLength;
        internal bool Missing, Corrupt;
        public async ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default)
        { Writes++; MaximumLength = Math.Max(MaximumLength, reference.Length); await _inner.PutAsync(reference, content, cancellationToken); }
        public async ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
        {
            Reads++;
            if (Missing) return null;
            if (!Corrupt) return await _inner.OpenReadAsync(reference, cancellationToken);
            return new MemoryStream(new byte[reference.Length]);
        }
    }
}
