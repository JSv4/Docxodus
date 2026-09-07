// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Security.Cryptography;
using System.Text.Json;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Checked-in, independently reopenable files, not merely transient in-memory fixtures.</summary>
public sealed class DocxHistoryArchiveArtifactTests
{
    internal static readonly string Root = Path.GetFullPath("../../../../TestFiles/HistoryArchive");
    private static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true };

    [Fact]
    public async Task CheckedInLegalAndBackendHistoryFilesPreserveExactSnapshotsAndArbitraryComparison()
    {
        // Explicit fixture maintenance only. Ordinary test runs NEVER write into TestFiles.
        if (Environment.GetEnvironmentVariable("DOCXODUS_REBUILD_HISTORY_FIXTURES") == "1") await GenerateAsync();
        foreach (var name in new[] { "agreement", "charter-collaboration" })
        {
            var index = await ReadIndexAsync(name);
            var archivePath = Path.Combine(Root, name + ".docxhistory");
            Assert.Equal(index.ArchiveSha256, Digest(await File.ReadAllBytesAsync(archivePath)));
            using var file = File.OpenRead(archivePath);
            using var archive = await DocxHistoryArchive.OpenAsync(file);
            Assert.Equal(index.DocumentId, archive.DocumentId); Assert.Equal(index.Head, archive.View.Head);
            var versions = (await archive.ListVersionsAsync(limit: 100)).Versions;
            Assert.Equal(index.Versions.Select(v => v.Id).Reverse(), versions.Select(v => v.Id));
            foreach (var version in index.Versions)
            {
                var expected = await File.ReadAllBytesAsync(Path.Combine(Root, version.File));
                Assert.Equal(version.Sha256, Digest(expected));
                Assert.Equal(expected, await archive.ExportDocxAsync(version.Id));
                var stored = await archive.GetVersionAsync(version.Id);
                Assert.Equal(stored.Record.Snapshot.ContentDigest, PackageManifestGenerator.Generate(
                    await archive.ReplayAsync(stored.Record.Sequence)).OrderedOpcContentDigest);
                // Every exported version must be a document a session can actually project, not just bytes.
                using var docx = new DocxSession(expected);
                Assert.NotEmpty(docx.Project().Markdown);
            }
            Assert.Equal(await File.ReadAllBytesAsync(Path.Combine(Root, index.Versions[^1].File)), await archive.ExportDocxAsync());
            // Select a non-latest pair from immutable version IDs. No pairwise diffs are stored in the archive.
            var before = index.Versions[0].Id; var after = index.Versions.First(v => v.Sha256 != index.Versions[0].Sha256).Id;
            var comparison = await archive.CompareVersionsAsync(before, after, ComparisonSettings());
            Assert.NotEmpty(comparison.GetRevisions());
            // The committed comparison is a golden redline, not an unchecked by-product: comparing the
            // same immutable pair out of the archive must reproduce that file exactly. Regenerate with
            // DOCXODUS_REBUILD_HISTORY_FIXTURES=1 when a deliberate DocxDiff change moves the output.
            var committed = await File.ReadAllBytesAsync(Path.Combine(Root, index.ComparisonFile));
            Assert.Equal(index.ComparisonSha256, Digest(committed));
            Assert.Equal(committed, comparison.ToRedline().DocumentByteArray);
            using (var redline = new DocxSession(committed))
                Assert.NotEmpty(redline.ListRevisions());
            var compared = new WmlDocument(index.ComparisonFile, committed);
            Assert.Empty(DocxDiff.GetRevisions(new WmlDocument("before.docx", await archive.ExportDocxAsync(before)),
                RevisionProcessor.RejectRevisions(compared)));
            Assert.Empty(DocxDiff.GetRevisions(new WmlDocument("after.docx", await archive.ExportDocxAsync(after)),
                RevisionProcessor.AcceptRevisions(compared)));
            if (index.Conflict is { } conflict)
            {
                Assert.Equal("conflict", (await archive.GetOperationAsync(conflict)).Record.Status);
                Assert.Equal(await File.ReadAllBytesAsync(Path.Combine(Root, index.ProposalFile!)),
                    await archive.ExportOperationProposalAsync(conflict));
                Assert.Equal(3, (await archive.ReadOperationsSinceAsync(null)).Operations.Count);
            }
        }
    }

    internal static async Task<GoldenIndex> ReadIndexAsync(string name) => JsonSerializer.Deserialize<GoldenIndex>(
        await File.ReadAllTextAsync(Path.Combine(Root, name + ".json")), Json)!;

    private static async Task GenerateAsync()
    {
        Directory.CreateDirectory(Root);
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        const string id = "matter-services-agreement";
        var original = await File.ReadAllBytesAsync("../../../../TestFiles/VP/VP004-Legal-Contract.docx");
        var initial = await history.CreateVersionAsync(id, "initial", null, original, HistoryArchiveFixture.Metadata("Original agreement"));
        var revised = await history.CreateVersionAsync(id, "revised", initial.Head,
            HistoryArchiveFixture.Edit(original, " — counsel's revised draft"), HistoryArchiveFixture.Metadata("Counsel revision"));
        var named = await history.CreateVersionAsync(id, "approved", revised.Head,
            await history.ExportVersionAsync(id, revised.Version.Id), HistoryArchiveFixture.Metadata("Approved for review"));
        await history.RestoreVersionAsync(id, "restored", named.Head, initial.Version.Id, HistoryArchiveFixture.Metadata("Restore original"));
        await WriteAsync("agreement", history, id, null);
        var fixture = await HistoryArchiveFixture.CreateAsync();
        await WriteAsync("charter-collaboration", fixture.History, HistoryArchiveFixture.Id, fixture.Conflict.Operation.Id);
    }

    private static async Task WriteAsync(string name, DocxVersionHistory history, string id, HistoryBlobReference? conflict)
    {
        var path = Path.Combine(Root, name + ".docxhistory");
        DocxHistoryArchiveInfo info;
        await using (var file = File.Create(path)) info = await history.ExportHistoryArchiveAsync(id, file);
        var versions = new List<GoldenVersion>();
        foreach (var version in (await history.ListVersionsAsync(id, limit: 100)).Versions.Reverse())
        {
            var file = $"{name}-v{versions.Count + 1}.docx";
            var bytes = await history.ExportVersionAsync(id, version.Id);
            await File.WriteAllBytesAsync(Path.Combine(Root, file), bytes);
            versions.Add(new(version.Id, file, Digest(bytes)));
        }
        string? proposal = null;
        if (conflict is not null)
        {
            proposal = name + "-conflicting-proposal.docx";
            await File.WriteAllBytesAsync(Path.Combine(Root, proposal), await history.ExportOperationProposalAsync(id, conflict));
        }
        var comparisonFile = name + "-comparison.docx";
        using (var archive = await DocxHistoryArchive.OpenAsync(await File.ReadAllBytesAsync(path)))
        {
            var after = versions.First(v => v.Sha256 != versions[0].Sha256);
            var comparison = await archive.CompareVersionsAsync(versions[0].Id, after.Id, ComparisonSettings());
            await File.WriteAllBytesAsync(Path.Combine(Root, comparisonFile), comparison.ToRedline().DocumentByteArray);
        }
        var index = new GoldenIndex(id, info.Head, Digest(await File.ReadAllBytesAsync(path)), versions.ToArray(), conflict, proposal,
            comparisonFile, Digest(await File.ReadAllBytesAsync(Path.Combine(Root, comparisonFile))));
        await File.WriteAllTextAsync(Path.Combine(Root, name + ".json"), JsonSerializer.Serialize(index, Json) + "\n");
    }

    private static DocxDiffSettings ComparisonSettings() => new() { PreAcceptInputRevisions = true, PreserveInputRevisions = true };
    private static string Digest(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    internal sealed record GoldenIndex(string DocumentId, HistoryHead Head, string ArchiveSha256,
        GoldenVersion[] Versions, HistoryBlobReference? Conflict, string? ProposalFile,
        string ComparisonFile, string ComparisonSha256);
    internal sealed record GoldenVersion(HistoryBlobReference Id, string File, string Sha256);
}
