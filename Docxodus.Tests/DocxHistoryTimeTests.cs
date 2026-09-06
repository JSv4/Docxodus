// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DocxHistoryTimeTests
{
    [Fact]
    public async Task LookupUsesCommitTimesWithSequenceTieBreaksAndIgnoresBackdatedLabels()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var initial = Document("initial");
        var first = await history.CreateVersionAsync("doc", null, initial, Metadata(0));
        var oneBytes = Document("one");
        var one = await history.CreateVersionAsync("doc", first.Head, oneBytes, Metadata(10));
        var label = await history.CreateVersionAsync("doc", one.Head, oneBytes, Metadata(-10));
        Assert.Equal(0, await history.ResolveSequenceAtTimeAsync("doc", Time(5)));
        var two = await history.CreateVersionAsync("doc", label.Head, Document("two"), Metadata(10));
        Assert.Equal(2, await history.ResolveSequenceAtTimeAsync("doc", Time(10)));
        var threeBytes = Document("three");
        var three = await history.CreateVersionAsync("doc", two.Head, threeBytes, Metadata(5));
        Assert.Equal(3, await history.ResolveSequenceAtTimeAsync("doc", Time(5)));
        Assert.Equal(3, await history.ResolveSequenceAtTimeAsync("doc", Time(20)));
        var lateLabel = await history.CreateVersionAsync("doc", three.Head, threeBytes, Metadata(-20));
        Assert.Equal(0, await history.ResolveSequenceAtTimeAsync("doc", Time(4)));
        var absent = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ResolveSequenceAtTimeAsync("doc", Time(-1)));
        Assert.Equal(DocxHistoryError.HistoryUnavailable, absent.Code);
        var restored = await history.RestoreVersionAsync("doc", lateLabel.Head, first.Version.Id, Metadata(20));
        Assert.Equal(3, await history.ResolveSequenceAtTimeAsync("doc", Time(19)));
        var sequence = await history.ResolveSequenceAtTimeAsync("doc", Time(20));
        Assert.Equal(4, sequence);
        Assert.Equal(PackageManifestGenerator.Generate(initial).OrderedOpcContentDigest,
            PackageManifestGenerator.Generate(await history.MaterializeAsync("doc", sequence)).OrderedOpcContentDigest);
        Assert.Equal(restored.Head, (await history.ReadAsync("doc"))!.Head);
    }

    [Fact]
    public async Task InitialTimeComesFromTheInitialVersionNotLaterMetadataOnlyVersions()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var bytes = Document("initial");
        var first = await history.CreateVersionAsync("doc", null, bytes, Metadata(10));
        var named = await history.CreateVersionAsync("doc", first.Head, bytes, Metadata(0));
        var unavailable = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ResolveSequenceAtTimeAsync("doc", Time(5)));
        Assert.Equal(DocxHistoryError.HistoryUnavailable, unavailable.Code);
        Assert.Equal(0, await history.ResolveSequenceAtTimeAsync("doc", Time(10)));
        Assert.Equal(named.Head, (await history.ReadAsync("doc"))!.Head);
    }

    [Fact]
    public async Task BudgetIncludesInitialVersionAncestryAndCancellationIsExplicit()
    {
        var history = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        DocxHistoryView? current = null;
        for (var i = 0; i < 4; i++)
            current = await history.CreateVersionAsync("doc", current?.Head, Document($"v{i}"), Metadata(i * 10));
        Assert.Equal(3, await history.ResolveSequenceAtTimeAsync("doc", Time(30), maxEntriesToScan: 1));
        var limit = await Assert.ThrowsAsync<DocxHistoryException>(async () =>
            await history.ResolveSequenceAtTimeAsync("doc", Time(0), maxEntriesToScan: 3));
        Assert.Equal(DocxHistoryError.TraversalLimit, limit.Code);
        Assert.Equal(0, await history.ResolveSequenceAtTimeAsync("doc", Time(0), maxEntriesToScan: 4));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await history.ResolveSequenceAtTimeAsync("doc", Time(0), cancellationToken: cancellation.Token));
        var missing = await Assert.ThrowsAsync<DocxHistoryException>(async () => await history.ResolveSequenceAtTimeAsync("absent", Time(0)));
        Assert.Equal(DocxHistoryError.HistoryUnavailable, missing.Code);
        Assert.Equal(current!.Head, (await history.ReadAsync("doc"))!.Head);
    }

    private static DateTimeOffset Time(int seconds) => DateTimeOffset.UnixEpoch.AddSeconds(seconds);
    private static DocxVersionMetadata Metadata(int seconds) => new() { Author = "host", CreatedAt = Time(seconds) };
    private static byte[] Document(string text)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        Assert.True(session.ReplaceText(session.Project().AnchorIndex.Keys.First(), text).Success);
        return session.Save();
    }
}
