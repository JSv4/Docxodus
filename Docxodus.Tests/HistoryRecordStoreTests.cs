// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class HistoryRecordStoreTests
{
    [Fact]
    public async Task AllRecordKindsRoundTripWithoutReadingTheirReferencedGraphs()
    {
        var blobs = new MemoryHistoryBlobStore();
        var store = new HistoryRecordStore(blobs);
        var version = Version();
        var versionId = await store.SaveVersionAsync(version);
        var loadedVersion = await store.LoadVersionAsync(versionId);
        Assert.Equal(version.DocumentId, loadedVersion.DocumentId);
        Assert.Equal(version.Metadata.ApplicationMetadata, loadedVersion.Metadata.ApplicationMetadata);
        Assert.Equal(versionId, await store.SaveVersionAsync(loadedVersion));
        var commit = new PackageHistoryCommitRecord
        {
            After = Snapshot('b'), Before = Snapshot('a'), Contribution = Ref('c'), DocumentId = "doc", Epoch = 0,
            Kind = "import", Parent = null, Sequence = 1, Version = versionId,
        };
        var commitId = await store.SaveCommitAsync(commit);
        Assert.Equal(commit, await store.LoadCommitAsync(commitId));
        var state = new DocxHistoryStateRecord
        {
            Commit = commitId, DocumentId = "doc", Epoch = 0, InitialSnapshot = Snapshot('a'), Sequence = 1,
            Snapshot = Snapshot('b'), Version = versionId,
        };
        var stateId = await store.SaveStateAsync(state);
        Assert.Equal(state, await store.LoadStateAsync(stateId));
        var restored = commit with { Kind = "restore", Contribution = null, Epoch = 1, After = Snapshot('a') };
        Assert.Equal(restored, await store.LoadCommitAsync(await store.SaveCommitAsync(restored)));
        var wrongKind = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.LoadStateAsync(versionId));
        Assert.Equal(PackageChangeError.InvalidManifest, wrongKind.Code);
    }

    [Fact]
    public async Task CanonicalMetadataOrderingAndImmutableLoadedValues()
    {
        var store = new HistoryRecordStore(new MemoryHistoryBlobStore());
        var metadata = new Dictionary<string, string> { ["z"] = "last", ["a"] = "first" };
        var record = Version() with { Metadata = Version().Metadata with { ApplicationMetadata = metadata } };
        var first = await store.SaveVersionAsync(record);
        var reordered = record with { Metadata = record.Metadata with
        {
            ApplicationMetadata = new Dictionary<string, string> { ["a"] = "first", ["z"] = "last" },
        } };
        Assert.Equal(first, await store.SaveVersionAsync(reordered));
        metadata["z"] = "changed";
        var loaded = await store.LoadVersionAsync(first);
        Assert.Equal("last", loaded.Metadata.ApplicationMetadata["z"]);
        Assert.Throws<NotSupportedException>(() => ((IDictionary<string, string>)loaded.Metadata.ApplicationMetadata).Add("bad", "bad"));
        Assert.NotEqual(first, await store.SaveVersionAsync(record with { Nonce = Guid.NewGuid() }));
    }

    [Theory]
    [InlineData("version", PackageChangeError.UnsupportedVersion)]
    [InlineData("unknown", PackageChangeError.InvalidManifest)]
    [InlineData("duplicate", PackageChangeError.InvalidManifest)]
    [InlineData("missing", PackageChangeError.InvalidManifest)]
    [InlineData("missing_length", PackageChangeError.InvalidManifest)]
    [InlineData("null_record", PackageChangeError.InvalidManifest)]
    [InlineData("null_values", PackageChangeError.InvalidManifest)]
    [InlineData("invalid_reference", PackageChangeError.InvalidManifest)]
    [InlineData("negative_sequence", PackageChangeError.InvalidManifest)]
    [InlineData("empty_nonce", PackageChangeError.InvalidManifest)]
    public async Task MalformedRecordsFailExplicitly(string scenario, PackageChangeError expected)
    {
        var blobs = new MemoryHistoryBlobStore();
        var store = new HistoryRecordStore(blobs);
        var id = await store.SaveVersionAsync(Version());
        using var original = await blobs.OpenReadAsync(id);
        var json = (await JsonNode.ParseAsync(original!))!;
        switch (scenario)
        {
            case "version": json["schemaVersion"] = 2; json["record"] = "future shape"; break;
            case "unknown": json["record"]!["unexpected"] = true; break;
            case "missing": json["record"]!.AsObject().Remove("sequence"); break;
            case "missing_length": json["record"]!["snapshot"]!["blob"]!.AsObject().Remove("length"); break;
            case "null_record": json["record"] = null; break;
            case "null_values": json["record"]!["metadata"]!["applicationMetadata"]!["null"] = null; break;
            case "invalid_reference": json["record"]!["snapshot"]!["blob"]!["digest"]!["value"] = "../invalid"; break;
            case "negative_sequence": json["record"]!["sequence"] = -1; break;
            case "empty_nonce": json["record"]!["nonce"] = Guid.Empty.ToString(); break;
        }
        var text = json.ToJsonString();
        if (scenario == "duplicate") text = text.Replace("\"sequence\":0", "\"sequence\":0,\"sequence\":0");
        var corrupt = await Put(blobs, Encoding.UTF8.GetBytes(text));
        var error = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.LoadVersionAsync(corrupt));
        Assert.Equal(expected, error.Code);
    }

    [Fact]
    public async Task BudgetsAndCancellationRejectWithoutStorageAccess()
    {
        var blobs = new RejectingStore();
        var store = new HistoryRecordStore(blobs, maxRecordBytes: 100);
        var tooLarge = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.SaveVersionAsync(Version()));
        Assert.Equal(PackageChangeError.ResourceLimit, tooLarge.Code);
        var reference = Ref('a') with { Length = 101 };
        var readLimit = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.LoadVersionAsync(reference));
        Assert.Equal(PackageChangeError.ResourceLimit, readLimit.Code);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await store.SaveVersionAsync(Version(), cancelled.Token));
        var oversizedText = Version() with { Metadata = Version().Metadata with { Message = new string('x', 16385) } };
        var fieldLimit = await Assert.ThrowsAsync<PackageChangeException>(async () => await store.SaveVersionAsync(oversizedText));
        Assert.Equal(PackageChangeError.ResourceLimit, fieldLimit.Code);
    }

    [Fact]
    public async Task StructuralPositionAndKindInvariantsAreEnforced()
    {
        var store = new HistoryRecordStore(new RejectingStore());
        var commit = new PackageHistoryCommitRecord
        {
            After = Snapshot('a'), Before = Snapshot('a'), Contribution = Ref('c'), DocumentId = "doc", Epoch = 0,
            Kind = "import", Parent = null, Sequence = 1, Version = Ref('d'),
        };
        foreach (var invalid in new[] { commit, commit with { Kind = "restore" }, commit with { Sequence = 2 },
            commit with { Kind = "restore", Contribution = null } })
            await Assert.ThrowsAsync<PackageChangeException>(async () => await store.SaveCommitAsync(invalid));
        await Assert.ThrowsAsync<PackageChangeException>(async () => await store.SaveVersionAsync(Version() with { RestoredFrom = Ref('a') }));
    }

    private static DocxVersionRecord Version() => new()
    {
        DocumentId = "doc", Metadata = new DocxVersionMetadata { Author = "Author", CreatedAt = DateTimeOffset.UnixEpoch },
        Nonce = Guid.Parse("2156d2d0-71d1-4301-8c59-0a19d0e78444"), Parent = null, RestoredFrom = null,
        Sequence = 0, Snapshot = Snapshot('a'),
    };
    private static HistoryBlobReference Ref(char digit) => new(new VerificationDigest { Algorithm = "SHA-256", Value = new string(digit, 64) }, 10);
    private static DocxSnapshotReference Snapshot(char digit) => new(Ref(digit), Ref(digit).Digest);
    private static async Task<HistoryBlobReference> Put(IHistoryBlobStore blobs, byte[] bytes)
    {
        var reference = new HistoryBlobReference(new VerificationDigest
        {
            Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
        }, bytes.Length);
        using var content = new MemoryStream(bytes);
        await blobs.PutAsync(reference, content);
        return reference;
    }
    private sealed class RejectingStore : IHistoryBlobStore
    {
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken) =>
            throw new InvalidOperationException("Unexpected write.");
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken) =>
            throw new InvalidOperationException("Unexpected read.");
    }
}
