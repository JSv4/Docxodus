// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public sealed class HistoryArchiveGraphTests
{
    [Fact]
    public async Task RealLegalDocumentClosureAloneReopensSnapshotsEffectsRestoresAndConflictingProposals()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var source = new CountingBlobs(fixture.Blobs);
        var graph = await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head, source);
        Assert.Equal(fixture.Latest.Head, graph.View.Head);
        Assert.DoesNotContain(graph.Inventory, r => r == fixture.Unrelated.Version.Id || r == fixture.Sentinel);
        Assert.DoesNotContain(fixture.Unrelated.Version.Id, source.Reads.Keys);
        Assert.DoesNotContain(fixture.Sentinel, source.Reads.Keys);
        Assert.Equal(graph.Inventory.Count, graph.Inventory.Select(r => r.Digest).Distinct().Count());
        var isolated = new MemoryHistoryBlobStore();
        foreach (var reference in graph.Inventory)
        {
            using var stream = await fixture.Blobs.OpenReadAsync(reference);
            await isolated.PutAsync(reference, stream!);
        }
        var reopened = new DocxVersionHistory(isolated, new PinnedHead(fixture.Latest.Head));
        Assert.Equal(fixture.Latest.Head, (await reopened.ReadAsync(HistoryArchiveFixture.Id))!.Head);
        var originalPage = await fixture.History.ListVersionsAsync(HistoryArchiveFixture.Id);
        var importedPage = await reopened.ListVersionsAsync(HistoryArchiveFixture.Id);
        Assert.Equal(originalPage.Versions.Select(v => v.Id), importedPage.Versions.Select(v => v.Id));
        foreach (var version in importedPage.Versions)
        {
            Assert.Equal(await fixture.History.ExportVersionAsync(HistoryArchiveFixture.Id, version.Id),
                await reopened.ExportVersionAsync(HistoryArchiveFixture.Id, version.Id));
            Assert.Equal(version.Record.Snapshot.ContentDigest, PackageManifestGenerator.Generate(
                await reopened.ReplayAsync(HistoryArchiveFixture.Id, version.Record.Sequence)).OrderedOpcContentDigest);
        }
        var operations = await reopened.ReadOperationsSinceAsync(HistoryArchiveFixture.Id, null);
        Assert.Equal(new[] { "accepted", "conflict", "accepted" }, operations.Operations.Select(o => o.Record.Status));
        Assert.Equal(fixture.Conflict.Operation.Id, operations.Operations[1].Id);
        Assert.Equal(fixture.Proposal, await reopened.ExportOperationProposalAsync(HistoryArchiveFixture.Id, fixture.Conflict.Operation.Id));
        var firstRequest = fixture.Initial.State.Requests!.Current!;
        Assert.Equal(fixture.Initial.Head, await new HistoryRequestJournalStore(isolated).FindAsync(
            HistoryArchiveFixture.Id, fixture.Latest.Head, fixture.Latest.State.Requests, firstRequest));
        Assert.Equal(fixture.Original, await reopened.ExportVersionAsync(HistoryArchiveFixture.Id, fixture.Initial.Version.Id));
    }

    [Fact]
    public async Task TraversalUsesPinnedHeadEvenWhenSourceHasNewerPublications()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var captured = fixture.Latest;
        var newer = await fixture.History.CreateVersionAsync(HistoryArchiveFixture.Id, "after-capture", captured.Head,
            fixture.Original, HistoryArchiveFixture.Metadata("later"));
        var graph = await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, captured.Head, fixture.Blobs);
        Assert.Equal(captured.Head, graph.View.Head);
        Assert.DoesNotContain(newer.Head.State, graph.Inventory);
        Assert.DoesNotContain(newer.Version.Id, graph.Inventory);
    }

    [Fact]
    public async Task AllExplicitGraphBudgetsAreEnforced()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        foreach (var limits in new[]
        {
            new DocxHistoryArchiveLimits { MaxBlobs = 1 },
            new DocxHistoryArchiveLimits { MaxTotalBlobBytes = 1 },
            new DocxHistoryArchiveLimits { MaxMetadataBytes = 1 },
            new DocxHistoryArchiveLimits { MaxValidationBytes = 1 },
            new DocxHistoryArchiveLimits { MaxExpandedBytes = 1 },
            new DocxHistoryArchiveLimits { MaxBlobBytes = 1 },
            new DocxHistoryArchiveLimits { MaxSnapshotBytes = 1 },
            new DocxHistoryArchiveLimits { MaxRecordBytes = 1 },
            new DocxHistoryArchiveLimits { MaxManifestBytes = 1 },
        })
            Assert.Equal(PackageChangeError.ResourceLimit, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head, fixture.Blobs, limits))).Code);
        Assert.Equal(DocxHistoryError.TraversalLimit, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head, fixture.Blobs,
                new DocxHistoryArchiveLimits { MaxEdges = 1 }))).Code);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await HistoryArchiveGraph.LoadAsync(
            HistoryArchiveFixture.Id, fixture.Latest.Head, fixture.Blobs, cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public async Task MissingCorruptAndForeignReachableDataFailEvenAwayFromTheLatestSnapshot()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var graph = await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head, fixture.Blobs);
        // Includes old snapshots, both branches of the receipt index, original request heads,
        // proposed work, package effects and payloads. Every retained byte is actually inspected.
        foreach (var reference in graph.Inventory)
        {
            var missing = new ReplacedBlob(fixture.Blobs, reference, null);
            Assert.Equal(PackageChangeError.PayloadMissing, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
                await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head, missing))).Code);
        }
        var corrupt = new ReplacedBlob(fixture.Blobs, fixture.Initial.State.Snapshot.Blob, new byte[] { 0 });
        Assert.Equal(PackageChangeError.PayloadMismatch, (await Assert.ThrowsAsync<PackageChangeException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head, corrupt))).Code);
        Assert.Equal(DocxHistoryError.ForeignDocument, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync("another-document", fixture.Latest.Head, fixture.Blobs))).Code);
    }

    [Fact]
    public async Task StateReferenceCannotHideASecondInvalidHeadRevision()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var falseHead = fixture.Latest.Head with { Revision = fixture.Latest.Head.Revision + 1 };
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, falseHead, fixture.Blobs))).Code);
    }

    [Fact]
    public async Task SameContentForkCommitCannotSubstituteForThePublishedVersion()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync(); var records = new HistoryRecordStore(fixture.Blobs);
        var commit = await records.LoadCommitAsync(fixture.Latest.State.Commit!);
        var version = await records.LoadVersionAsync(commit.Version);
        var fork = await records.SaveVersionAsync(version with { Nonce = Guid.NewGuid() });
        var forkCommit = await records.SaveCommitAsync(commit with { Version = fork });
        var falseState = await records.SaveStateAsync(fixture.Latest.State with { Commit = forkCommit });
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id,
                fixture.Latest.Head with { State = falseState }, fixture.Blobs))).Code);
    }

    [Fact]
    public async Task FutureReceiptCannotBeHiddenInAnOlderJournal()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var later = await fixture.History.CreateVersionAsync(HistoryArchiveFixture.Id, "later", fixture.Latest.Head,
            fixture.Original, HistoryArchiveFixture.Metadata("later"));
        var falseState = await new HistoryRecordStore(fixture.Blobs).SaveStateAsync(fixture.Latest.State with
        {
            Requests = fixture.Latest.State.Requests! with { Index = later.State.Requests!.Index },
        });
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id,
                fixture.Latest.Head with { State = falseState }, fixture.Blobs))).Code);
    }

    [Fact]
    public async Task OneBlobMayBeBothAnExactSnapshotAndAnOpaqueContributionPayload()
    {
        var original = DocxSession.CreateBlankDocxBytes();
        using var buffer = new MemoryStream(); buffer.Write(original);
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            var typesEntry = zip.GetEntry("[Content_Types].xml")!; XDocument types;
            using (var input = typesEntry.Open()) types = XDocument.Load(input);
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Default", new XAttribute("Extension", "docx"),
                new XAttribute("ContentType", "application/octet-stream")));
            typesEntry.Delete();
            using (var output = zip.CreateEntry("[Content_Types].xml").Open()) types.Save(output);
            using var entry = zip.CreateEntry("custom/attached.docx").Open(); entry.Write(original);
        }
        var blobs = new MemoryHistoryBlobStore(); var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var first = await history.CreateVersionAsync("doc", null, original, HistoryArchiveFixture.Metadata("original"));
        var next = await history.CreateVersionAsync("doc", first.Head, buffer.ToArray(), HistoryArchiveFixture.Metadata("attachment"));
        var graph = await HistoryArchiveGraph.LoadAsync("doc", next.Head, blobs);
        Assert.Single(graph.Inventory, r => r == first.State.Snapshot.Blob);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TextProposalAndAcceptedMapMustRepresentTheOriginalIntent(bool missingMap)
    {
        var blobs = new MemoryHistoryBlobStore(); var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var bytes = File.ReadAllBytes("../../../../TestFiles/NVCA-Model-COI.docx");
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var input = zip.GetEntry("word/document.xml")!.Open();
        var texts = XDocument.Load(input).Descendants(W.t).ToArray();
        var ordinal = Array.FindIndex(texts, t => t.Value.Length >= 4);
        Assert.True(ordinal >= 0);
        var first = await history.CreateVersionAsync("doc", "initial", null, bytes, HistoryArchiveFixture.Metadata("initial"));
        var request = new DocxOperationRequest
        {
            RequestId = "text-a", Base = first.Head, Kind = "text", Metadata = HistoryArchiveFixture.Metadata("a"),
            Text = new DocxTextSplice("/word/document.xml", ordinal, 0, 1, "X"),
        };
        var accepted = await history.SubmitOperationAsync("doc", request);
        var conflict = await history.SubmitOperationAsync("doc", request with
        { RequestId = "text-b", Text = request.Text with { Insert = "Y" } });
        Assert.Equal("conflict", conflict.Operation.Record.Status);
        await HistoryArchiveGraph.LoadAsync("doc", conflict.View.Head, blobs);
        var selected = missingMap ? accepted : conflict;
        var wrong = await new DocxOperationStore(blobs).SaveDecisionAsync(missingMap
            ? selected.Operation.Record with { AppliedText = null }
            : selected.Operation.Record with { ProposedSnapshot = first.State.Snapshot }, default);
        var state = await new HistoryRecordStore(blobs).SaveStateAsync(selected.View.State with { Operation = wrong });
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync("doc", selected.View.Head with { State = state }, blobs))).Code);
    }

    [Fact]
    public async Task ANewPublicationCannotForgetRetryReceiptsRetainedByItsParent()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var state = await new HistoryRecordStore(fixture.Blobs).SaveStateAsync(fixture.Latest.State with
        { Requests = fixture.Latest.State.Requests! with { Index = null } });
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, fixture.Latest.Head with { State = state }, fixture.Blobs))).Code);
    }

    [Fact]
    public async Task AConflictCannotBeAcceptedAsResolvedTwice()
    {
        var fixture = await HistoryArchiveFixture.CreateAsync();
        var again = await fixture.History.SubmitOperationAsync(HistoryArchiveFixture.Id, new DocxOperationRequest
        {
            RequestId = "resolve-again", Base = fixture.Latest.Head, Kind = "discard",
            Resolves = fixture.Conflict.Operation.Id, Metadata = HistoryArchiveFixture.Metadata("Already resolved"),
        });
        Assert.Equal("AlreadyResolved", again.Operation.Record.Conflict);
        var wrong = await new DocxOperationStore(fixture.Blobs).SaveDecisionAsync(again.Operation.Record with
        { Status = "accepted", Conflict = null }, default);
        var state = await new HistoryRecordStore(fixture.Blobs).SaveStateAsync(again.View.State with { Operation = wrong });
        Assert.Equal(DocxHistoryError.InvalidHistory, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await HistoryArchiveGraph.LoadAsync(HistoryArchiveFixture.Id, again.View.Head with { State = state }, fixture.Blobs))).Code);
    }

    [Fact]
    public async Task PersistentReceiptDAGIsReadOncePerNodeNotOncePerHistoricalRoot()
    {
        var blobs = new MemoryHistoryBlobStore(); var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var bytes = DocxSession.CreateBlankDocxBytes(); DocxHistoryView? current = null;
        for (var i = 0; i < 96; i++) current = await history.CreateVersionAsync("doc", $"request-{i}", current?.Head,
            bytes, HistoryArchiveFixture.Metadata($"version-{i}"));
        var counting = new CountingBlobs(blobs);
        var graph = await HistoryArchiveGraph.LoadAsync("doc", current!.Head, counting);
        Assert.Equal(graph.Inventory.Count, counting.Reads.Count);
        Assert.All(counting.Reads.Values, count => Assert.Equal(1, count));
    }

    [Fact]
    public async Task RestoreMayRetainASameDocumentVersionFromADifferentInitialBranch()
    {
        var blobs = new MemoryHistoryBlobStore();
        var left = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var right = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var bytes = DocxSession.CreateBlankDocxBytes();
        var a = await left.CreateVersionAsync("doc", null, bytes, HistoryArchiveFixture.Metadata("a"));
        var source = File.ReadAllBytes("../../../../TestFiles/VP/VP004-Legal-Contract.docx");
        var b = await right.CreateVersionAsync("doc", null, source, HistoryArchiveFixture.Metadata("b"));
        var restored = await left.RestoreVersionAsync("doc", a.Head, b.Version.Id, HistoryArchiveFixture.Metadata("restore b"));
        var graph = await HistoryArchiveGraph.LoadAsync("doc", restored.Head, blobs);
        Assert.Contains(b.Version.Id, graph.Inventory);
        Assert.Contains(b.State.Snapshot.Blob, graph.Inventory);
        Assert.DoesNotContain(b.Head.State, graph.Inventory);
        Assert.Equal(a.State.InitialSnapshot, graph.View.State.InitialSnapshot);
        Assert.Equal(b.Version.Id, graph.View.Version.Record.RestoredFrom);
    }

    internal sealed class PinnedHead(HistoryHead head) : IHistoryHeadStore
    {
        public ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default) => ValueTask.FromResult<HistoryHead?>(head);
        public ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }
    private sealed class CountingBlobs(IHistoryBlobStore inner) : IHistoryBlobStore
    {
        internal Dictionary<HistoryBlobReference, int> Reads { get; } = new();
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
        { Reads[reference] = Reads.GetValueOrDefault(reference) + 1; return inner.OpenReadAsync(reference, cancellationToken); }
    }
    private sealed class ReplacedBlob(IHistoryBlobStore inner, HistoryBlobReference target, byte[]? replacement) : IHistoryBlobStore
    {
        public ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default) => reference == target
            ? ValueTask.FromResult<Stream?>(replacement is null ? null : new MemoryStream(replacement, false))
            : inner.OpenReadAsync(reference, cancellationToken);
    }
}

internal sealed record HistoryArchiveFixture(MemoryHistoryBlobStore Blobs, DocxVersionHistory History,
    byte[] Original, byte[] Proposal, DocxHistoryView Initial, DocxHistoryView Latest,
    DocxOperationResult Conflict, DocxHistoryView Unrelated, HistoryBlobReference Sentinel)
{
    internal const string Id = "matter-nvca-charter";
    internal static DocxVersionMetadata Metadata(string label) => new()
    {
        Author = "Archive test counsel", CreatedAt = new DateTimeOffset(2026, 8, 15, 10, 0, 0, TimeSpan.FromHours(-5)),
        Label = label, ApplicationMetadata = new Dictionary<string, string> { ["matter"] = "NVCA charter" },
    };
    internal static async Task<HistoryArchiveFixture> CreateAsync()
    {
        var original = File.ReadAllBytes("../../../../TestFiles/NVCA-Model-COI.docx");
        var blobs = new MemoryHistoryBlobStore(); var history = new DocxVersionHistory(blobs, new MemoryHistoryHeadStore());
        var initial = await history.CreateVersionAsync(Id, "initial", null, original, Metadata("Original charter"));
        var named = await history.CreateVersionAsync(Id, "named", initial.Head, original, Metadata("Partner review"));
        var edited = await history.CreateVersionAsync(Id, "edited", named.Head, Edit(original, " Counsel revision"), Metadata("Revised charter"));
        var restored = await history.RestoreVersionAsync(Id, "restore", edited.Head, initial.Version.Id, Metadata("Restore original"));
        var request = new DocxOperationRequest { RequestId = "counsel-a", Base = restored.Head, Kind = "package", Metadata = Metadata("Counsel A") };
        await history.SubmitOperationAsync(Id, request, Edit(original, " Counsel A"));
        var proposal = Edit(original, " Counsel B");
        var conflict = await history.SubmitOperationAsync(Id, request with { RequestId = "counsel-b", Metadata = Metadata("Counsel B") }, proposal);
        Assert.Equal("conflict", conflict.Operation.Record.Status);
        var discarded = await history.SubmitOperationAsync(Id, new DocxOperationRequest
        {
            RequestId = "discard-b", Base = conflict.View.Head, Kind = "discard", Resolves = conflict.Operation.Id,
            Metadata = Metadata("Keep counsel A"),
        });
        var latest = await history.CreateVersionAsync(Id, "final-label", discarded.View.Head,
            await history.ExportVersionAsync(Id, discarded.View.Version.Id), Metadata("Approved"));
        var unrelated = await history.CreateVersionAsync("unrelated-client", null, Edit(original, " Confidential unrelated client"), Metadata("PRIVATE"));
        var sentinel = await HistoryBlobIO.PutBytesAsync(blobs, Encoding.UTF8.GetBytes("Unreferenced private data"), default);
        return new(blobs, history, original, proposal, initial, latest, conflict, unrelated, sentinel);
    }
    internal static byte[] Edit(byte[] original, string suffix)
    {
        using var output = new MemoryStream(); output.Write(original);
        using (var zip = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            XDocument document;
            using (var source = entry.Open()) document = XDocument.Load(source, LoadOptions.PreserveWhitespace);
            var text = document.Descendants(W.t).First(t => !string.IsNullOrWhiteSpace(t.Value));
            text.Value += suffix;
            entry.Delete();
            using var destination = zip.CreateEntry("word/document.xml").Open(); document.Save(destination, SaveOptions.DisableFormatting);
        }
        return output.ToArray();
    }
}
