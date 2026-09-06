// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class PackageChangeSetCodecTests
{
    [Fact]
    public async Task PersistedManifestAndBlobsReplayBothDirectionsWithoutTheOriginalChangeSet()
    {
        var (before, after, original) = Example();
        var store = new TestBlobStore();
        var manifest = await PackageChangeSetCodec.SaveAsync(original, store);
        Assert.Equal(original.PayloadDigests.Count, store.Values.Count);
        Assert.DoesNotContain("base64", Encoding.UTF8.GetString(manifest));

        var restored = await PackageChangeSetCodec.LoadAsync(manifest, store);
        Assert.Equal(manifest, PackageChangeSetCodec.Encode(restored));
        Assert.Equal(original.Changes, restored.Changes);
        Assert.Equal(PackageManifestGenerator.Generate(after).OrderedOpcContentDigest,
            PackageManifestGenerator.Generate(restored.Apply(before)).OrderedOpcContentDigest);
        Assert.Equal(PackageManifestGenerator.Generate(before).OrderedOpcContentDigest,
            PackageManifestGenerator.Generate(restored.Invert().Apply(after)).OrderedOpcContentDigest);
        Assert.All(store.ReadStreams, stream => Assert.False(stream.CanRead));
    }

    [Theory]
    [InlineData("unknown", PackageChangeError.InvalidManifest)]
    [InlineData("duplicate", PackageChangeError.InvalidManifest)]
    [InlineData("version", PackageChangeError.UnsupportedVersion)]
    [InlineData("path", PackageChangeError.InvalidManifest)]
    [InlineData("missing_blob_ref", PackageChangeError.InvalidManifest)]
    [InlineData("duplicate_change", PackageChangeError.InvalidManifest)]
    [InlineData("case_duplicate_change", PackageChangeError.InvalidManifest)]
    [InlineData("negative_length", PackageChangeError.InvalidManifest)]
    [InlineData("invalid_digest", PackageChangeError.InvalidManifest)]
    [InlineData("no_op_different_roots", PackageChangeError.InvalidManifest)]
    public async Task InvalidManifestFailsBeforeAnyBlobReads(string scenario, PackageChangeError expected)
    {
        var (_, _, original) = Example();
        var manifest = PackageChangeSetCodec.Encode(original);
        var json = JsonNode.Parse(manifest)!.AsObject();
        switch (scenario)
        {
            case "unknown": json["unexpected"] = true; break;
            case "version": json["schemaVersion"] = 2; break;
            case "path": json["changes"]![0]!["afterEntryName"] = "../escape.xml"; break;
            case "missing_blob_ref": json["payloads"]!.AsArray().RemoveAt(0); break;
            case "duplicate_change": json["changes"]!.AsArray().Add(json["changes"]![0]!.DeepClone()); break;
            case "case_duplicate_change":
                var duplicate = json["changes"]![0]!.DeepClone();
                foreach (var property in new[] { "uri", "beforeEntryName", "afterEntryName" })
                    duplicate[property] = duplicate[property]!.GetValue<string>().ToUpperInvariant();
                json["changes"]!.AsArray().Add(duplicate);
                break;
            case "negative_length": json["payloads"]![0]!["length"] = -1; break;
            case "invalid_digest": json["beforeDigest"] = "not-a-digest"; break;
            case "no_op_different_roots": json["changes"] = new JsonArray(); json["payloads"] = new JsonArray(); break;
        }
        var text = json.ToJsonString();
        if (scenario == "duplicate") text = text.Replace("\"schemaVersion\":1", "\"schemaVersion\":1,\"schemaVersion\":1");
        var store = new TestBlobStore();
        var error = await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await PackageChangeSetCodec.LoadAsync(Encoding.UTF8.GetBytes(text), store));
        Assert.Equal(expected, error.Code);
        Assert.Equal(0, store.ReadCount);
    }

    [Theory]
    [InlineData("missing", PackageChangeError.PayloadMissing)]
    [InlineData("truncated", PackageChangeError.PayloadMismatch)]
    [InlineData("long", PackageChangeError.PayloadMismatch)]
    [InlineData("corrupt", PackageChangeError.PayloadMismatch)]
    public async Task BlobCorruptionIsDetected(string scenario, PackageChangeError expected)
    {
        var (_, _, changes) = Example();
        var store = new TestBlobStore();
        var manifest = await PackageChangeSetCodec.SaveAsync(changes, store);
        var key = store.Values.Keys.First();
        if (scenario == "missing") store.Values.Remove(key);
        if (scenario == "truncated") store.Values[key] = store.Values[key][..^1];
        if (scenario == "long") store.Values[key] = store.Values[key].Append((byte)0).ToArray();
        if (scenario == "corrupt") store.Values[key][0] ^= 1;

        var error = await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await PackageChangeSetCodec.LoadAsync(manifest, store));
        Assert.Equal(expected, error.Code);
        Assert.All(store.ReadStreams, stream => Assert.False(stream.CanRead));
    }

    [Fact]
    public async Task AllMetadataBudgetsAreEnforcedBeforeStorageReadsOrWrites()
    {
        var (_, _, changes) = Example();
        var manifest = PackageChangeSetCodec.Encode(changes);
        foreach (var limits in new[]
        {
            new PackageChangeLimits { MaxManifestBytes = manifest.Length - 1 },
            new PackageChangeLimits { MaxEntryNameLength = 1 },
            new PackageChangeLimits { MaxPayloadBytes = 1 },
            new PackageChangeLimits { MaxTotalPayloadBytes = 1 },
        })
        {
            var store = new TestBlobStore();
            var read = await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await PackageChangeSetCodec.LoadAsync(manifest, store, limits));
            var write = await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await PackageChangeSetCodec.SaveAsync(changes, store, limits));
            Assert.Equal(PackageChangeError.ResourceLimit, read.Code);
            Assert.Equal(PackageChangeError.ResourceLimit, write.Code);
            Assert.Equal(0, store.ReadCount);
            Assert.Empty(store.Values);
        }
    }

    [Fact]
    public async Task AVerifiedButIncorrectBeforePayloadCannotBeApplied()
    {
        var (before, _, changes) = Example();
        var store = new TestBlobStore();
        var manifest = JsonNode.Parse(await PackageChangeSetCodec.SaveAsync(changes, store))!;
        var oldDigest = manifest["changes"]![0]!["beforeDigest"]!.GetValue<string>();
        var bytes = Encoding.UTF8.GetBytes("a different valid blob");
        var digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        store.Values[digest] = bytes;
        manifest["changes"]![0]!["beforeDigest"] = digest;
        var reference = manifest["payloads"]!.AsArray().Single(p => p!["digest"]!.GetValue<string>() == oldDigest)!;
        reference["digest"] = digest;
        reference["length"] = bytes.Length;

        var decoded = await PackageChangeSetCodec.LoadAsync(Encoding.UTF8.GetBytes(manifest.ToJsonString()), store);
        var error = Assert.Throws<PackageChangeException>(() => decoded.Apply(before));
        Assert.Equal(PackageChangeError.BaseMismatch, error.Code);
    }

    [Fact]
    public async Task CancellationAndFailedStorageNeverReturnAPublishedManifest()
    {
        var (_, _, changes) = Example();
        var store = new TestBlobStore { FailWrites = true };
        await Assert.ThrowsAsync<IOException>(async () => await PackageChangeSetCodec.SaveAsync(changes, store));
        store.FailWrites = false;
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await PackageChangeSetCodec.SaveAsync(changes, store, cancellationToken: cancelled.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await PackageChangeSetCodec.LoadAsync(PackageChangeSetCodec.Encode(changes), store,
                cancellationToken: cancelled.Token));
        Assert.Empty(store.Values);
        Assert.Equal(0, store.ReadCount);
    }

    [Fact]
    public async Task NoOpManifestNeedsNoBlobs()
    {
        var bytes = DocxSession.CreateBlankDocxBytes();
        var store = new TestBlobStore();
        var manifest = await PackageChangeSetCodec.SaveAsync(PackageChangeSet.Create(bytes, bytes), store);
        var restored = await PackageChangeSetCodec.LoadAsync(manifest, store);
        Assert.Equal(bytes, restored.Apply(bytes));
        Assert.Empty(store.Values);
        Assert.Equal(0, store.ReadCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DirectoryArtifactsAndCaseOnlyRenamesRoundTrip(bool rename)
    {
        var before = DocxSession.CreateBlankDocxBytes();
        if (rename) before = EditEntries(before, archive =>
        {
            using var writer = new StreamWriter(archive.CreateEntry("customXml/item.xml").Open());
            writer.Write("<data/>");
        });
        var after = EditEntries(before, archive =>
        {
            if (!rename) archive.CreateEntry("empty/");
            else
            {
                var original = archive.GetEntry("customXml/item.xml")!;
                using (var source = original.Open())
                using (var target = archive.CreateEntry("customXml/ITEM.xml").Open()) source.CopyTo(target);
                original.Delete();
            }
        });
        var changes = PackageChangeSet.Create(before, after);
        Assert.NotEmpty(changes.Changes);
        if (!rename) Assert.Equal(changes.BeforeDigest, changes.AfterDigest);
        var store = new TestBlobStore();
        var manifest = await PackageChangeSetCodec.SaveAsync(changes, store);
        var restored = await PackageChangeSetCodec.LoadAsync(manifest, store);
        Assert.Equal(manifest, PackageChangeSetCodec.Encode(restored));
        Assert.Equal(EntryNames(after), EntryNames(restored.Apply(before)));
        Assert.Equal(EntryNames(before), EntryNames(restored.Invert().Apply(after)));
    }

    private static byte[] EditEntries(byte[] bytes, Action<ZipArchive> edit)
    {
        using var buffer = new MemoryStream();
        buffer.Write(bytes);
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true)) edit(archive);
        return buffer.ToArray();
    }

    private static string[] EntryNames(byte[] bytes)
    {
        using var buffer = new MemoryStream(bytes);
        using var archive = new ZipArchive(buffer, ZipArchiveMode.Read);
        return archive.Entries.Select(e => e.FullName).Order(StringComparer.Ordinal).ToArray();
    }

    private static (byte[] Before, byte[] After, PackageChangeSet Changes) Example()
    {
        var before = DocxSession.CreateBlankDocxBytes();
        using var session = new DocxSession(before);
        var anchor = session.Project().AnchorIndex.Keys.First();
        Assert.True(session.ReplaceText(anchor, "durable change").Success);
        var after = session.Save();
        return (before, after, PackageChangeSet.Create(before, after));
    }

    private sealed class TestBlobStore : IHistoryBlobStore
    {
        internal Dictionary<string, byte[]> Values { get; } = new();
        internal List<Stream> ReadStreams { get; } = new();
        internal int ReadCount { get; private set; }
        internal bool FailWrites { get; set; }

        public async ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken)
        {
            if (FailWrites) throw new IOException("Storage unavailable.");
            using var buffer = new MemoryStream();
            await content.CopyToAsync(buffer, cancellationToken);
            Values[reference.Digest.Value] = buffer.ToArray();
        }

        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken)
        {
            ReadCount++;
            Stream? stream = Values.TryGetValue(reference.Digest.Value, out var bytes) ? new FragmentedStream(bytes) : null;
            if (stream is not null) ReadStreams.Add(stream);
            return ValueTask.FromResult(stream);
        }
    }

    private sealed class FragmentedStream(byte[] bytes) : MemoryStream(bytes, writable: false)
    {
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            base.ReadAsync(buffer[..Math.Min(buffer.Length, 3)], cancellationToken);
    }
}
