#nullable enable

using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using Docxodus.History;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxVersionRequestTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CreateAndRestoreRetriesReturnOriginalHeadsAfterReopenAndLaterLegacyWrites(bool file)
    {
        var root = Directory.CreateTempSubdirectory("version-request-").FullName;
        try
        {
            IHistoryBlobStore blobs = file ? new FileHistoryBlobStore(Path.Combine(root, "blobs")) : new MemoryHistoryBlobStore();
            IHistoryHeadStore heads = file ? new FileHistoryHeadStore(Path.Combine(root, "heads")) : new MemoryHistoryHeadStore();
            var history = new DocxVersionHistory(blobs, heads);
            var firstBytes = Document("initial");
            var secondBytes = Document("revised");
            var first = await history.CreateVersionAsync("doc", "r1", null, firstBytes, Metadata("one"));
            var second = await history.CreateVersionAsync("doc", "r2", first.Head, secondBytes, Metadata("two"));
            var restored = await history.RestoreVersionAsync("doc", "r3", second.Head, first.Version.Id, Metadata("restore"));
            var legacy = await history.CreateVersionAsync("doc", restored.Head, secondBytes, Metadata("legacy"));
            history = file ? new DocxVersionHistory(new FileHistoryBlobStore(Path.Combine(root, "blobs")),
                new FileHistoryHeadStore(Path.Combine(root, "heads"))) : new DocxVersionHistory(blobs, heads);
            Assert.Equal(first.Head, (await history.CreateVersionAsync("doc", "r1", null, firstBytes, Metadata("one"))).Head);
            Assert.Equal(second.Head, (await history.CreateVersionAsync("doc", "r2", first.Head, secondBytes, Metadata("two"))).Head);
            Assert.Equal(restored.Head, (await history.RestoreVersionAsync("doc", "r3", second.Head, first.Version.Id, Metadata("restore"))).Head);
            Assert.Equal(legacy.Head, (await history.ReadAsync("doc"))!.Head);
            Assert.Equal(4, (await history.ListVersionsAsync("doc")).Versions.Count);
            Assert.Equal(firstBytes, await history.ExportVersionAsync("doc", first.Version.Id));
            Assert.Equal(secondBytes, await history.ExportVersionAsync("doc", second.Version.Id));
            Assert.Equal(1, restored.State.Epoch);
            Assert.Equal(3, (await history.ReadChangesSinceAsync("doc", first.Head)).Entries.Count);
        }
        finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task IdReuseBindsExactBytesExpectedHeadOperationTargetAndNormalizedMetadata()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = Document("initial");
        var metadata = Metadata("named") with { ApplicationMetadata = new Dictionary<string, string> { ["b"] = "2", ["a"] = "1" } };
        var first = await history.CreateVersionAsync("doc", "r1", null, bytes, metadata);
        var reordered = metadata with { ApplicationMetadata = new Dictionary<string, string> { ["a"] = "1", ["b"] = "2" } };
        Assert.Equal(first.Head, (await history.CreateVersionAsync("doc", "r1", null, bytes, reordered)).Head);
        async Task Conflict(Func<ValueTask<DocxHistoryView>> action) => Assert.Equal(DocxHistoryError.RequestConflict,
            (await Assert.ThrowsAsync<DocxHistoryException>(async () => await action())).Code);
        await Conflict(() => history.CreateVersionAsync("doc", "r1", first.Head, bytes, metadata));
        await Conflict(() => history.CreateVersionAsync("doc", "r1", null, Document("changed"), metadata));
        await Conflict(() => history.CreateVersionAsync("doc", "r1", null, bytes, metadata with { Author = "other" }));
        await Conflict(() => history.CreateVersionAsync("doc", "r1", null, bytes, metadata with { CreatedAt = metadata.CreatedAt.AddTicks(1) }));
        await Conflict(() => history.RestoreVersionAsync("doc", "r1", first.Head, first.Version.Id, metadata));
        // Intentional second version of identical bytes is distinct; request IDs aren't blob hashes.
        var second = await history.CreateVersionAsync("doc", "r2", first.Head, bytes, metadata);
        Assert.NotEqual(first.Version.Id, second.Version.Id);
        Assert.Equal(first.State.Snapshot, second.State.Snapshot);
        Assert.Equal(0, second.State.Sequence);
        Assert.Equal(1, (await history.CreateVersionAsync("another-document", "r1", null, bytes, metadata)).Head.Revision);
    }

    [Fact]
    public async Task ConcurrentIdenticalRetriesHaveOnePublicationAndOneResult()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var bytes = Document("shared");
        var results = await Task.WhenAll(Enumerable.Range(0, 16).Select(_ => Task.Run(async () =>
            await new DocxVersionHistory(blobs, heads).CreateVersionAsync("doc", "one-request", null, bytes, Metadata("one")))));
        Assert.All(results, result => Assert.Equal(results[0].Head, result.Head));
        Assert.Equal(1, (await heads.ReadAsync("doc"))!.Revision);
        var view = await new DocxVersionHistory(blobs, heads).ReadAsync("doc");
        Assert.Equal(results[0].Version.Id, view!.Version.Id);
    }

    [Fact]
    public async Task LostAcknowledgementAfterCommittedCASIsRecoveredAfterOtherPublications()
    {
        var blobs = new MemoryHistoryBlobStore();
        var backing = new MemoryHistoryHeadStore();
        var heads = new LostAckHeads(backing) { LoseNext = true };
        var history = new DocxVersionHistory(blobs, heads);
        var bytes = Document("saved");
        await Assert.ThrowsAsync<IOException>(async () => await history.CreateVersionAsync("doc", "r1", null, bytes, Metadata("one")));
        var original = (await history.ReadAsync("doc"))!;
        var later = await history.CreateVersionAsync("doc", original.Head, bytes, Metadata("later"));
        var restarted = new DocxVersionHistory(blobs, backing);
        Assert.Equal(original.Head, (await restarted.CreateVersionAsync("doc", "r1", null, bytes, Metadata("one"))).Head);
        Assert.Equal(later.Head, (await restarted.ReadAsync("doc"))!.Head);
        heads.LoseNext = true;
        await Assert.ThrowsAsync<IOException>(async () => await history.RestoreVersionAsync("doc", "restore", later.Head, original.Version.Id, Metadata("restore")));
        var reset = (await restarted.ReadAsync("doc"))!;
        Assert.Equal(reset.Head, (await restarted.RestoreVersionAsync("doc", "restore", later.Head, original.Version.Id, Metadata("restore"))).Head);
        Assert.Equal(1, reset.State.Epoch);
    }

    [Fact]
    public async Task LegacyV1StateIsPreservedUntilFirstIdempotentPublicationAndV2RejectsForgedRevision()
    {
        var blobs = new MemoryHistoryBlobStore();
        var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var bytes = Document("legacy");
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata("one"));
        using (var stream = await blobs.OpenReadAsync(first.Head.State))
        using (var reader = new StreamReader(stream!))
        {
            var json = await reader.ReadToEndAsync();
            Assert.Contains("\"schemaVersion\":1", json);
            Assert.DoesNotContain("requests", json);
            var malformed = JsonNode.Parse(json)!;
            malformed["record"]!["requests"] = null;
            var bad = await HistoryBlobIO.PutBytesAsync(blobs, Encoding.UTF8.GetBytes(malformed.ToJsonString()), default);
            Assert.Equal(PackageChangeError.InvalidManifest, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await new HistoryRecordStore(blobs).LoadStateAsync(bad))).Code);
        }
        var second = await history.CreateVersionAsync("doc", "r2", first.Head, bytes, Metadata("two"));
        using (var stream = await blobs.OpenReadAsync(second.Head.State))
        using (var reader = new StreamReader(stream!)) Assert.Contains("\"schemaVersion\":2", await reader.ReadToEndAsync());
        var invalid = await heads.TryAdvanceAsync("doc", second.Head, second.Head.State);
        Assert.Equal(DocxHistoryError.InvalidHistory,
            (await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ReadAsync("doc"))).Code);
        Assert.NotNull(invalid);
    }

    [Fact]
    public async Task InvalidRequestIdsNeverFallBackToNonIdempotentPublication()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = Document("unchanged");
        foreach (var id in new[] { null, "", " ", "\ud800", new string('x', 1025) })
            await Assert.ThrowsAnyAsync<ArgumentException>(async () => await history.CreateVersionAsync("doc", id!, null, bytes, Metadata("invalid")));
        Assert.Null(await history.ReadAsync("doc"));
    }

    [Fact]
    public async Task RequestIdsTooLargeToIndexAreRefusedBeforeTheyCanStrandLaterPublications()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = Document("bounded");
        // Each ID is within its own limit, but the pair cannot fit the receipt index node that the
        // NEXT publication would have to promote it into. The refusal must happen here: accepting
        // it would publish a head no later publication of this document could ever extend.
        var documentId = new string('\u5408', 1024);
        Assert.Equal(PackageChangeError.ResourceLimit, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await history.CreateVersionAsync(documentId, new string('\u540c', 1024), null, bytes, Metadata("oversized")))).Code);
        Assert.Null(await history.ReadAsync(documentId));

        // The same document keeps working with a request ID that fits, including across the
        // publication that promotes the first receipt and a later retry of it.
        var first = await history.CreateVersionAsync(documentId, "matter-42", null, bytes, Metadata("first"));
        var second = await history.CreateVersionAsync(documentId, "matter-43", first.Head, Document("changed"), Metadata("second"));
        Assert.Equal(first.Head, (await history.CreateVersionAsync(documentId, "matter-42", null, bytes, Metadata("first"))).Head);
        Assert.Equal(second.Head, (await history.ReadAsync(documentId))!.Head);
    }

    internal static DocxVersionMetadata Metadata(string label) => new()
    { Author = "actor", CreatedAt = DateTimeOffset.Parse("2026-01-01T12:00:00Z"), Label = label };

    internal static byte[] Document(string text)
    {
        var bytes = DocxSession.CreateBlankDocxBytes();
        using var stream = new MemoryStream(); stream.Write(bytes);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var part = zip.GetEntry("word/document.xml")!;
            using var content = part.Open();
            var xml = Encoding.UTF8.GetBytes("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>"
                + System.Security.SecurityElement.Escape(text) + "</w:t></w:r></w:p></w:body></w:document>");
            content.SetLength(0); content.Write(xml);
        }
        return stream.ToArray();
    }

    private sealed class LostAckHeads(IHistoryHeadStore inner) : IHistoryHeadStore
    {
        internal bool LoseNext;
        public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default) => inner.ReadAsync(id, cancellationToken);
        public async ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state, CancellationToken cancellationToken = default)
        {
            var result = await inner.TryAdvanceAsync(id, expected, state, cancellationToken);
            if (result is not null && LoseNext) { LoseNext = false; throw new IOException("Acknowledgement lost after durable publication."); }
            return result;
        }
    }
}
