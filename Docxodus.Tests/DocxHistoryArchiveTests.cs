// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using System.Buffers.Binary;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Docxodus.History;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxHistoryArchiveTests
{
    [Fact]
    public async Task RealHistoryFileReopensIndependentlyAndExportsExactVersionFiles()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var directory = Directory.CreateTempSubdirectory("docxodus-portable-history-").FullName;
        try
        {
            var path = Path.Combine(directory, "Charter.docxhistory");
            DocxHistoryArchiveInfo info;
            await using (var output = File.Create(path))
                info = await fixture.History.ExportHistoryArchiveAsync(HistoryArchiveFixture.Id, output);
            Assert.Equal(fixture.Latest.Head, info.Head);
            Assert.True(new FileInfo(path).Length > 1000);
            using var input = File.OpenRead(path);
            using (var archive = await DocxHistoryArchive.OpenAsync(input))
            {
                Assert.Equal(info, archive.Info);
                Assert.Equal(fixture.Latest.Head, archive.View.Head);
                Assert.Equal(fixture.Original, await archive.ExportDocxAsync(fixture.Initial.Version.Id));
                var versions = await archive.ListVersionsAsync();
                Assert.Equal(6, versions.Versions.Count);
                foreach (var version in versions.Versions)
                {
                    var bytes = await archive.ExportDocxAsync(version.Id);
                    var versionPath = Path.Combine(directory, $"version-{version.Id.Digest.Value}.docx");
                    await File.WriteAllBytesAsync(versionPath, bytes);
                    Assert.Equal(await fixture.History.ExportVersionAsync(HistoryArchiveFixture.Id, version.Id),
                        await File.ReadAllBytesAsync(versionPath));
                    using var reopened = new DocxSession(await File.ReadAllBytesAsync(versionPath));
                    Assert.NotNull(reopened);
                }
                Assert.Equal(fixture.Proposal, await archive.ExportOperationProposalAsync(fixture.Conflict.Operation.Id));
                using var exportedAgain = new MemoryStream();
                Assert.Equal(info, await archive.ExportHistoryArchiveAsync(exportedAgain));
                // Container compression may vary by source-stream chunking; every entry stays exact.
                using var firstZip = new ZipArchive(new MemoryStream(await File.ReadAllBytesAsync(path)), ZipArchiveMode.Read);
                using var secondZip = new ZipArchive(new MemoryStream(exportedAgain.ToArray()), ZipArchiveMode.Read);
                Assert.Equal(firstZip.Entries.Select(e => e.FullName), secondZip.Entries.Select(e => e.FullName));
                foreach (var entry in firstZip.Entries) Assert.Equal(Read(entry), Read(secondZip.GetEntry(entry.FullName)!));
            }
            Assert.True(input.CanRead); // Default stream ownership is caller-owned.
            using var bytesArchive = await DocxHistoryArchive.OpenAsync(await File.ReadAllBytesAsync(path));
            Assert.Equal(fixture.Latest.Head, (await bytesArchive.ReadAsync())!.Head);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task OwnedInputCopiesBeforeAwaitAndConcurrentReadsDoNotShareZipPosition()
    {
        var (bytes, _) = await SmallArchive();
        var saved = bytes.ToArray();
        var opening = DocxHistoryArchive.OpenAsync(bytes).AsTask(); Array.Fill(bytes, (byte)0);
        using var archive = await opening;
        var reads = await Task.WhenAll(Enumerable.Range(0, 12).Select(_ => archive.ExportDocxAsync().AsTask()));
        Assert.All(reads, read => Assert.Equal(reads[0], read));
        using var input = new MemoryStream(saved);
        var owned = await DocxHistoryArchive.OpenAsync(input, leaveOpen: false); owned.Dispose();
        Assert.False(input.CanRead);
        await Assert.ThrowsAsync<ObjectDisposedException>(async () => await owned.ExportDocxAsync());
    }

    [Theory]
    [InlineData("duplicate-zip")]
    [InlineData("unknown-zip")]
    [InlineData("path-zip")]
    [InlineData("uppercase-zip")]
    [InlineData("duplicate-json")]
    [InlineData("unknown-json")]
    [InlineData("numeric-revision")]
    [InlineData("noncanonical-revision")]
    [InlineData("unsorted-inventory")]
    [InlineData("extra-blob")]
    [InlineData("missing-blob")]
    [InlineData("wrong-blob-length")]
    [InlineData("corrupt-blob")]
    public async Task HostileHistoryFilesFailClosed(string scenario)
    {
        var (bytes, _) = await SmallArchive();
        using var original = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entries = original.Entries.Select(e => (e.FullName, Bytes: Read(e))).ToList();
        var manifest = Encoding.UTF8.GetString(entries[0].Bytes);
        var parsed = JsonNode.Parse(manifest)!.AsObject();
        switch (scenario)
        {
            case "duplicate-zip": entries.Add(entries[1]); break;
            case "unknown-zip": entries.Add(("private.txt", new byte[] { 42 })); break;
            case "path-zip": entries.Add(("../private.txt", new byte[] { 42 })); break;
            case "uppercase-zip": entries[1] = (entries[1].FullName.ToUpperInvariant(), entries[1].Bytes); break;
            case "duplicate-json": manifest = manifest.Replace("\"schemaVersion\":1", "\"schemaVersion\":1,\"schemaVersion\":1"); break;
            case "unknown-json": parsed["extra"] = true; manifest = parsed.ToJsonString(); break;
            case "numeric-revision": parsed["head"]!["revision"] = 1; manifest = parsed.ToJsonString(); break;
            case "noncanonical-revision": parsed["head"]!["revision"] = "01"; manifest = parsed.ToJsonString(); break;
            case "unsorted-inventory":
                parsed["blobs"]!.AsArray().Add(parsed["blobs"]![0]!.DeepClone()); manifest = parsed.ToJsonString(); break;
            case "extra-blob":
                var extra = DocxOperationStore.Reference(new byte[] { 42 });
                entries.Add((HistoryArchiveManifest.BlobName(extra), new byte[] { 42 }));
                parsed["blobs"]!.AsArray().Add(new JsonObject { ["sha256"] = extra.Digest.Value, ["length"] = 1 });
                parsed["blobs"] = new JsonArray(parsed["blobs"]!.AsArray().Select(n => n!.DeepClone())
                    .OrderBy(n => n["sha256"]!.GetValue<string>(), StringComparer.Ordinal).ToArray());
                manifest = parsed.ToJsonString(); break;
            case "missing-blob": entries.RemoveAt(1); break;
            case "wrong-blob-length": entries[1] = (entries[1].FullName, new byte[1]); break;
            case "corrupt-blob": entries[1].Bytes[0] ^= 0xff; break;
        }
        entries[0] = (HistoryArchiveManifest.EntryName, Encoding.UTF8.GetBytes(manifest));
        using var malformed = new MemoryStream(Pack(entries));
        var error = await Record.ExceptionAsync(async () => { using var _ = await DocxHistoryArchive.OpenAsync(malformed); });
        Assert.True(error is PackageChangeException or DocxHistoryException, error?.ToString() ?? "Archive unexpectedly opened.");
        Assert.True(malformed.CanRead);
    }

    [Fact]
    public async Task ArchiveByteAndEntryBudgetsPrecedeZipEntryAllocationOrDecompression()
    {
        var (bytes, _) = await SmallArchive();
        Assert.Equal(PackageChangeError.ResourceLimit, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await DocxHistoryArchive.OpenAsync(bytes, new DocxHistoryArchiveLimits { MaxArchiveBytes = bytes.Length - 1 }))).Code);
        Assert.Equal(PackageChangeError.ResourceLimit, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await DocxHistoryArchive.OpenAsync(bytes, new DocxHistoryArchiveLimits { MaxBlobs = 1 }))).Code);
        using var counted = new CountedInput(bytes);
        await Assert.ThrowsAsync<PackageChangeException>(async () => await DocxHistoryArchive.OpenAsync(counted,
            limits: new DocxHistoryArchiveLimits { MaxArchiveBytes = bytes.Length - 1 }));
        Assert.Equal(0, counted.BytesRead);
        await Assert.ThrowsAsync<PackageChangeException>(async () => await DocxHistoryArchive.OpenAsync(counted,
            limits: new DocxHistoryArchiveLimits { MaxBlobs = 1 }));
        Assert.Equal(22, counted.BytesRead); // Only the footer: no central directory or entry data.
        var malformed = bytes[..^1];
        await Assert.ThrowsAsync<PackageChangeException>(async () => await DocxHistoryArchive.OpenAsync(malformed));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await DocxHistoryArchive.OpenAsync(bytes, cancellationToken: new CancellationToken(true)));
        var (history, _) = await SmallHistory();
        using var output = new MemoryStream();
        Assert.Equal(PackageChangeError.ResourceLimit, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await history.ExportHistoryArchiveAsync("doc", output, new DocxHistoryArchiveLimits { MaxArchiveBytes = 32 }))).Code);
        Assert.True(output.CanWrite); Assert.True(output.Length <= 32);
        await Assert.ThrowsAsync<ArgumentException>(async () => await history.ExportHistoryArchiveAsync("doc", output));
    }

    [Fact]
    public async Task CompressedSnapshotCapDoesNotRejectLargerUncompressedEffectPayloads()
    {
        var before = DocxSession.CreateBlankDocxBytes(); using var buffer = new MemoryStream(); buffer.Write(before);
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!; XDocument document;
            using (var input = entry.Open()) document = XDocument.Load(input);
            var paragraph = new XElement(W.p, new XElement(W.r, new XElement(W.t,
                string.Join(" ", Enumerable.Range(0, 20_000).Select(i => $"Clause-{i:D6}")))));
            var body = document.Descendants(W.body).Single();
            if (body.Element(W.sectPr) is { } section) section.AddBeforeSelf(paragraph); else body.Add(paragraph);
            entry.Delete(); using var output = zip.CreateEntry("word/document.xml").Open(); document.Save(output);
        }
        var after = buffer.ToArray(); var cap = Math.Max(before.Length, after.Length) + 16;
        var blobs = new MemoryHistoryBlobStore();
        var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore(), maxSnapshotBytes: cap);
        var initial = await history.CreateVersionAsync("doc", null, before, HistoryArchiveFixture.Metadata("initial"));
        var edited = await history.CreateVersionAsync("doc", initial.Head, after, HistoryArchiveFixture.Metadata("long clauses"));
        var commit = await new HistoryRecordStore(blobs).LoadCommitAsync(edited.State.Commit!);
        var manifest = await HistoryBlobIO.ReadBytesAsync(blobs, commit.Contribution!, 1024 * 1024, default);
        Assert.Contains(PackageChangeSetCodec.ReadInventory(manifest, new PackageChangeLimits()).Payloads, p => p.Length > cap);
        using var archiveFile = new MemoryStream();
        await history.ExportHistoryArchiveAsync("doc", archiveFile);
        using var reopened = await DocxHistoryArchive.OpenAsync(archiveFile.ToArray());
        Assert.Equal(after, await reopened.ExportDocxAsync());
    }

    [Fact]
    public async Task HashValidButFalseContributionDeltaIsRejectedByTheFileReader()
    {
        var blobs = new MemoryHistoryBlobStore(); var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var before = File.ReadAllBytes("../../../../TestFiles/NVCA-Model-COI.docx");
        var initial = await history.CreateVersionAsync("doc", null, before, HistoryArchiveFixture.Metadata("initial"));
        var latest = await history.CreateVersionAsync("doc", initial.Head, HistoryArchiveFixture.Edit(before, " changed"),
            HistoryArchiveFixture.Metadata("edited"));
        var records = new HistoryRecordStore(blobs); var commit = await records.LoadCommitAsync(latest.State.Commit!);
        var bytes = await HistoryBlobIO.ReadBytesAsync(blobs, commit.Contribution!, 1024 * 1024, default);
        var json = JsonNode.Parse(bytes)!;
        var change = json["changes"]![0]!;
        change["uri"] = "/word/forged.xml"; change["beforeEntryName"] = "word/forged.xml"; change["afterEntryName"] = "word/forged.xml";
        var falseEffectBytes = Encoding.UTF8.GetBytes(json.ToJsonString());
        _ = PackageChangeSetCodec.ReadInventory(falseEffectBytes, new PackageChangeLimits()); // The codec itself accepts this manifest.
        var falseEffect = await HistoryBlobIO.PutBytesAsync(blobs, falseEffectBytes, default);
        var falseCommit = await records.SaveCommitAsync(commit with { Contribution = falseEffect });
        var falseState = await records.SaveStateAsync(latest.State with { Commit = falseCommit });
        var graph = await HistoryArchiveGraph.LoadAsync("doc", latest.Head, blobs);
        var inventory = graph.Inventory.Except(new[] { latest.Head.State, latest.State.Commit!, commit.Contribution! })
            .Concat(new[] { falseState, falseCommit, falseEffect }).OrderBy(r => r.Digest.Value, StringComparer.Ordinal).ToArray();
        var manifest = new HistoryArchiveManifest("doc", latest.Head with { State = falseState }, inventory);
        var entries = new List<(string Name, byte[] Bytes)> { (HistoryArchiveManifest.EntryName, manifest.Encode(new())) };
        foreach (var reference in inventory) entries.Add((HistoryArchiveManifest.BlobName(reference),
            await HistoryBlobIO.ReadBytesAsync(blobs, reference, 256 * 1024 * 1024, default)));
        var error = await Assert.ThrowsAsync<DocxHistoryException>(async () => await DocxHistoryArchive.OpenAsync(Pack(entries)));
        Assert.Equal(DocxHistoryError.InvalidHistory, error.Code);
        Assert.Contains("exact endpoint entry delta", error.Message);
    }

    [Fact]
    public async Task ExportCapturesOneHeadWhileAnotherWriterPublishesAndSupportsForwardOnlyOutput()
    {
        var blobs = new MemoryHistoryBlobStore(); var heads = new MemoryHistoryHeadStore();
        var history = new DocxVersionHistory(blobs, heads);
        var first = await history.CreateVersionAsync("doc", null, DocxSession.CreateBlankDocxBytes(), HistoryArchiveFixture.Metadata("first"));
        var paused = new PausedHead(heads); var exporting = new DocxVersionHistory(blobs, paused);
        using var buffer = new MemoryStream(); using var output = new ForwardOnly(buffer);
        var pending = exporting.ExportHistoryArchiveAsync("doc", output).AsTask();
        await paused.Captured.Task;
        var next = await history.CreateVersionAsync("doc", first.Head, DocxSession.CreateBlankDocxBytes(), HistoryArchiveFixture.Metadata("next"));
        paused.Release.SetResult();
        Assert.Equal(first.Head, (await pending).Head); Assert.Equal(1, paused.Reads);
        using var archive = await DocxHistoryArchive.OpenAsync(buffer.ToArray());
        Assert.Equal(first.Head, archive.View.Head);
        Assert.Equal(first.Version.Id, Assert.Single((await archive.ListVersionsAsync()).Versions).Id);
        Assert.DoesNotContain(next.Version.Id, archive.Inventory);
        using var input = new ForwardOnly(new MemoryStream(buffer.ToArray()));
        await Assert.ThrowsAsync<ArgumentException>(async () => await DocxHistoryArchive.OpenAsync(input));
    }

    [Fact]
    public async Task Zip64FooterIsSupportedAndItsHugeCountIsRejectedBeforeAllocation()
    {
        var (bytes, view) = await SmallArchive();
        var wide = Zip64(bytes);
        using var archive = await DocxHistoryArchive.OpenAsync(wide); Assert.Equal(view.Head, archive.View.Head);
        BinaryPrimitives.WriteUInt64LittleEndian(wide.AsSpan(bytes.Length - 22 + 24, 8), ulong.MaxValue);
        BinaryPrimitives.WriteUInt64LittleEndian(wide.AsSpan(bytes.Length - 22 + 32, 8), ulong.MaxValue);
        Assert.Equal(PackageChangeError.ResourceLimit, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await DocxHistoryArchive.OpenAsync(wide))).Code);
    }

    [Fact]
    public async Task CallerMutationAfterOpenCannotBypassBlobIntegrityChecks()
    {
        var (bytes, _) = await SmallArchive(); using var input = new MemoryStream(bytes, writable: true);
        using var archive = await DocxHistoryArchive.OpenAsync(input);
        Array.Fill(bytes, (byte)0);
        Assert.NotNull(await Record.ExceptionAsync(async () => await archive.ExportDocxAsync()));
    }

    private static byte[] Zip64(byte[] bytes)
    {
        var footer = bytes.AsSpan(bytes.Length - 22, 22).ToArray();
        var count = BinaryPrimitives.ReadUInt16LittleEndian(footer.AsSpan(10, 2));
        var size = BinaryPrimitives.ReadUInt32LittleEndian(footer.AsSpan(12, 4));
        var offset = BinaryPrimitives.ReadUInt32LittleEndian(footer.AsSpan(16, 4));
        using var output = new MemoryStream(); output.Write(bytes.AsSpan(0, bytes.Length - 22));
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        writer.Write(0x06064b50u); writer.Write(44ul); writer.Write((ushort)45); writer.Write((ushort)45);
        writer.Write(0u); writer.Write(0u); writer.Write((ulong)count); writer.Write((ulong)count);
        writer.Write((ulong)size); writer.Write((ulong)offset);
        writer.Write(0x07064b50u); writer.Write(0u); writer.Write((ulong)bytes.Length - 22); writer.Write(1u);
        BinaryPrimitives.WriteUInt16LittleEndian(footer.AsSpan(8, 2), ushort.MaxValue);
        BinaryPrimitives.WriteUInt16LittleEndian(footer.AsSpan(10, 2), ushort.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(footer.AsSpan(12, 4), uint.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(footer.AsSpan(16, 4), uint.MaxValue);
        output.Write(footer); return output.ToArray();
    }

    private sealed class PausedHead(IHistoryHeadStore inner) : IHistoryHeadStore
    {
        internal TaskCompletionSource Captured { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal int Reads;
        public async ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default)
        {
            Reads++; var captured = await inner.ReadAsync(documentId, cancellationToken);
            Captured.SetResult(); await Release.Task.WaitAsync(cancellationToken); return captured;
        }
        public ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }
    private sealed class ForwardOnly(Stream inner) : Stream
    {
        public override bool CanRead => inner.CanRead;
        public override bool CanWrite => inner.CanWrite;
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }

    private sealed class CountedInput(byte[] bytes) : MemoryStream(bytes, writable: false)
    {
        internal long BytesRead;
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = base.Read(buffer.Span); BytesRead += read; return ValueTask.FromResult(read);
        }
    }

    private static async Task<(DocxVersionHistory History, DocxHistoryView View)> SmallHistory()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var view = await history.CreateVersionAsync("doc", "initial", null, DocxSession.CreateBlankDocxBytes(), HistoryArchiveFixture.Metadata("initial"));
        return (history, view);
    }
    private static async Task<(byte[] Bytes, DocxHistoryView View)> SmallArchive()
    {
        var (history, view) = await SmallHistory(); using var output = new MemoryStream();
        await history.ExportHistoryArchiveAsync("doc", output); return (output.ToArray(), view);
    }
    private static byte[] Read(ZipArchiveEntry entry)
    { using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray(); }
    private static byte[] Pack(IEnumerable<(string Name, byte[] Bytes)> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var entry in entries)
            { using var content = zip.CreateEntry(entry.Name).Open(); content.Write(entry.Bytes); }
        return output.ToArray();
    }
}
