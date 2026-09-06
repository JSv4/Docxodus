#nullable enable

using System.IO.Compression;
using Docxodus.History;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;
using static Docxodus.Tests.DocxVersionRequestTests;

namespace Docxodus.Tests;

/// <summary>
/// Stateful oracle independent of history records, package digests, and replay implementation.
/// Override DOCXODUS_HISTORY_FUZZ_SEEDS/STEPS for the extended deterministic corpus. Every failure
/// reports its complete seed/command trace. Each seed checks all exact exports and all sequences.
/// </summary>
public sealed class DocxVersionModelFuzzTests(ITestOutputHelper output)
{
    public static IEnumerable<object[]> Seeds => Enumerable.Range(0, Setting("SEEDS", 16, 1024))
        .Select(index => new object[] { 60493 + index * 7919 });
    private sealed record Command(string? Id, HistoryHead? Expected, byte[] Bytes, DocxStoredVersion? Restore,
        DocxVersionMetadata Metadata);
    private sealed record ExpectedVersion(DocxHistoryView View, byte[] Bytes, DateTimeOffset Time, bool ContentBoundary);

    [Theory]
    [MemberData(nameof(Seeds))]
    public async Task VersionStreamsMatchIndependentModelUnderRetriesRacesFaultsAndRestart(int seed)
    {
        var steps = Setting("STEPS", 96, 10_000);
        var random = new Random(seed);
        var filesystem = seed % 5 == 0;
        using var storage = new HistoryFaultHarness(filesystem);
        var history = new DocxVersionHistory(storage, storage);
        var trace = new List<string>();
        var versions = new List<ExpectedVersion>();
        var requests = new List<(Command Command, DocxHistoryView Result)>();
        var pool = Enumerable.Range(0, 6).Select(i => Package(seed, i)).ToArray();
        var counts = new Dictionary<string, int>();
        void Count(string name) => counts[name] = counts.GetValueOrDefault(name) + 1;
        ExpectedVersion? Latest() => versions.LastOrDefault();
        DocxVersionMetadata Meta(int step) => Metadata("seed " + seed + " step " + step) with
        {
            Author = "actor-" + random.Next(3),
            CreatedAt = DateTimeOffset.UnixEpoch.AddMinutes(random.Next(-50, 51)).AddTicks(random.Next(7)),
            ApplicationMetadata = new Dictionary<string, string> { ["seed"] = seed.ToString(), ["step"] = step.ToString(), ["unicode"] = "合同 📝" },
        };
        try
        {
            for (var step = 0; step < steps; step++)
            {
                history = new DocxVersionHistory(storage, storage); // No producer instance survives a step.
                var action = versions.Count == 0 ? 0 : random.Next(12);
                trace.Add($"{step}: action={action}, head={Latest()?.View.Head.Revision ?? 0}");
                if (action <= 5)
                {
                    var restore = action == 5 ? versions[random.Next(versions.Count)] : null;
                    var bytes = restore?.Bytes ?? (action == 1 ? Latest()!.Bytes
                        : action == 2 ? Repack(Latest()!.Bytes, step) : pool[random.Next(pool.Length)]);
                    var requestId = action == 4 ? null : $"{seed}/{step}";
                    var command = new Command(requestId, Latest()?.View.Head, bytes, restore?.View.Version, Meta(step));
                    var lostAck = requestId is not null && step % 17 == 0;
                    if (lostAck) storage.Arm(HistoryFaultHarness.Fault.AfterHead);
                    DocxHistoryView result;
                    if (lostAck)
                    {
                        await Assert.ThrowsAsync<IOException>(async () => await Invoke(history, command));
                        storage.Arm(HistoryFaultHarness.Fault.None);
                        result = await Invoke(new DocxVersionHistory(storage, storage), command);
                        Assert.Equal(0, storage.Publications);
                        Count("lost-ack-recovery");
                    }
                    else result = await Invoke(history, command);
                    await Accept(command, result);
                    Count(restore is not null ? "restore" : action == 2 ? "repack" : requestId is null ? "legacy-save" : "save");
                }
                else if (action == 6 && requests.Count > 0)
                {
                    var old = requests[random.Next(requests.Count)];
                    var head = Latest()!.View.Head;
                    Assert.Equal(old.Result.Head, (await Invoke(history, old.Command)).Head);
                    Assert.Equal(head, (await history.ReadAsync("doc"))!.Head);
                    Count("old-request-retry");
                }
                else if (action == 7 && requests.Count > 0)
                {
                    var old = requests[random.Next(requests.Count)];
                    var collision = old.Command with { Metadata = old.Command.Metadata with { Message = "different logical request" } };
                    Assert.Equal(DocxHistoryError.RequestConflict,
                        (await Assert.ThrowsAsync<DocxHistoryException>(async () => await Invoke(history, collision))).Code);
                    Count("id-reuse-refusal");
                }
                else if (action == 8 && versions.Count > 1)
                {
                    var stale = new Command($"stale/{seed}/{step}", versions[random.Next(versions.Count - 1)].View.Head,
                        pool[random.Next(pool.Length)], null, Meta(step));
                    Assert.Equal(DocxHistoryError.StaleHead,
                        (await Assert.ThrowsAsync<DocxHistoryException>(async () => await Invoke(history, stale))).Code);
                    // A request never committed/bound may be corrected and submitted intentionally.
                    var corrected = stale with { Expected = Latest()!.View.Head };
                    await Accept(corrected, await Invoke(history, corrected));
                    Count("stale-then-corrected");
                }
                else if (action == 9)
                {
                    var target = versions[random.Next(versions.Count)];
                    storage.DamagedDigest = target.View.State.Snapshot.Blob.Digest.Value;
                    storage.Missing = random.Next(2) == 0;
                    try
                    {
                        Assert.Equal(storage.Missing ? PackageChangeError.PayloadMissing : PackageChangeError.PayloadMismatch,
                            (await Assert.ThrowsAsync<PackageChangeException>(async () => await history.ExportVersionAsync("doc", target.View.Version.Id))).Code);
                    }
                    finally { storage.DamagedDigest = null; }
                    Count("damaged-storage-refusal");
                }
                else if (action == 10)
                {
                    var sameRequest = random.Next(2) == 0;
                    var first = new Command($"race/{seed}/{step}/0", Latest()!.View.Head, pool[random.Next(pool.Length)], null, Meta(step));
                    var second = sameRequest ? first : first with { Id = $"race/{seed}/{step}/1", Bytes = pool[random.Next(pool.Length)] };
                    var firstWinner = random.Next(2) == 0;
                    trace.Add($"  race: sameRequest={sameRequest}, firstWinner={firstWinner}");
                    async Task<DocxHistoryView?> Compete(Command command)
                    {
                        try { return await Invoke(new DocxVersionHistory(storage, storage), command); }
                        catch (DocxHistoryException error) when (error.Code == DocxHistoryError.StaleHead) { return null; }
                    }
                    var attempts = storage.Publications;
                    storage.HoldCompetingPublications(firstWinner ? "left" : "right");
                    DocxHistoryView?[] results;
                    try
                    {
                        results = await Task.WhenAll(
                            Task.Run(() => storage.AsContenderAsync("left", () => Compete(first))),
                            Task.Run(() => storage.AsContenderAsync("right", () => Compete(second))));
                    }
                    finally { storage.ReleaseCompetingPublications(); }
                    Assert.Equal(2, storage.Publications - attempts); // Both must reach the exact same CAS expectation.
                    if (sameRequest)
                    {
                        Assert.NotNull(results[0]); Assert.NotNull(results[1]);
                        Assert.Equal(results[0]!.Head, results[1]!.Head);
                        await Accept(first, results[0]!);
                    }
                    else
                    {
                        Assert.Single(results.Where(x => x is not null));
                        Assert.Equal(firstWinner, results[0] is not null);
                        await Accept(results[0] is not null ? first : second, results[0] ?? results[1]!);
                    }
                    Count(sameRequest ? "identical-race" : "competing-race");
                }
                else
                {
                    var selected = versions[random.Next(versions.Count)];
                    Assert.Equal(selected.Bytes, await history.ExportVersionAsync("doc", selected.View.Version.Id));
                    Count("read");
                }
                Assert.Equal(Latest()!.View.Head, (await history.ReadAsync("doc"))!.Head);
            }

            // Immutable pagination and exact snapshots are checked against independently retained inputs.
            var listed = new List<DocxStoredVersion>();
            HistoryBlobReference? cursor = null;
            do
            {
                var page = await history.ListVersionsAsync("doc", cursor, 7);
                listed.AddRange(page.Versions); cursor = page.Next;
            } while (cursor is not null);
            Assert.Equal(versions.Select(v => v.View.Version.Id).Reverse(), listed.Select(v => v.Id));
            foreach (var version in versions)
                Assert.Equal(version.Bytes, await history.ExportVersionAsync("doc", version.View.Version.Id));

            var boundaries = versions.Where(v => v.ContentBoundary).ToArray();
            foreach (var boundary in boundaries)
            {
                var sequence = boundary.View.State.Sequence;
                // ZIP-entry comparison does not call the production manifest/digest/replay owner.
                SameContent(boundary.Bytes, await history.MaterializeAsync("doc", sequence));
                SameContent(boundary.Bytes, await history.ReplayAsync("doc", sequence));
            }
            foreach (var cutoff in boundaries.Select(v => v.Time).Append(DateTimeOffset.UnixEpoch.AddDays(-1)))
            {
                var expected = boundaries.Where(v => v.Time <= cutoff).LastOrDefault();
                if (expected is null)
                    Assert.Equal(DocxHistoryError.HistoryUnavailable, (await Assert.ThrowsAsync<DocxHistoryException>(async () =>
                        await history.ResolveSequenceAtTimeAsync("doc", cutoff))).Code);
                else Assert.Equal(expected.View.State.Sequence, await history.ResolveSequenceAtTimeAsync("doc", cutoff));
            }
            foreach (var (command, result) in requests)
                Assert.Equal(result.Head, (await Invoke(new DocxVersionHistory(storage, storage), command)).Head);
            output.WriteLine($"seed={seed}; steps={steps}; filesystem={filesystem}; versions={versions.Count}; replayed-sequences={boundaries.Length}; "
                + string.Join(", ", counts.OrderBy(p => p.Key).Select(p => p.Key + "=" + p.Value)));
        }
        catch (Exception error)
        { throw new XunitException($"History fuzz failure seed={seed}, steps={steps}, filesystem={filesystem}\n{string.Join("\n", trace)}\n{error}"); }

        async Task Accept(Command command, DocxHistoryView result)
        {
            var previous = Latest();
            var changed = previous is not null && !Equivalent(previous.Bytes, command.Bytes);
            var boundary = previous is null || command.Restore is not null || changed;
            var sequence = (previous?.View.State.Sequence ?? 0) + (previous is not null && boundary ? 1 : 0);
            var epoch = (previous?.View.State.Epoch ?? 0) + (command.Restore is not null ? 1 : 0);
            Assert.Equal((previous?.View.Head.Revision ?? 0) + 1, result.Head.Revision);
            Assert.Equal(sequence, result.State.Sequence); Assert.Equal(epoch, result.State.Epoch);
            Assert.Equal(previous?.View.Version.Id, result.Version.Record.Parent);
            Assert.Equal(command.Restore?.Id, result.Version.Record.RestoredFrom);
            Assert.Equal(command.Metadata.Author, result.Version.Record.Metadata.Author);
            Assert.Equal(command.Metadata.CreatedAt, result.Version.Record.Metadata.CreatedAt);
            Assert.Equal(command.Metadata.Label, result.Version.Record.Metadata.Label);
            Assert.Equal(command.Metadata.ApplicationMetadata, result.Version.Record.Metadata.ApplicationMetadata);
            Assert.Equal(command.Id, result.State.Requests?.Current?.Id);
            Assert.Equal(command.Bytes, await history.ExportVersionAsync("doc", result.Version.Id));
            versions.Add(new ExpectedVersion(result, command.Bytes.ToArray(), command.Metadata.CreatedAt, boundary));
            if (command.Id is not null) requests.Add((command, result));
        }
    }

    private static ValueTask<DocxHistoryView> Invoke(DocxVersionHistory history, Command command) => command.Restore is not null
        ? command.Id is null ? history.RestoreVersionAsync("doc", command.Expected!, command.Restore.Id, command.Metadata)
            : history.RestoreVersionAsync("doc", command.Id, command.Expected!, command.Restore.Id, command.Metadata)
        : command.Id is null ? history.CreateVersionAsync("doc", command.Expected, command.Bytes, command.Metadata)
            : history.CreateVersionAsync("doc", command.Id, command.Expected, command.Bytes, command.Metadata);

    private static byte[] Package(int seed, int variant)
    {
        using var stream = new MemoryStream(); stream.Write(Document($"seed {seed} variant {variant}: café 📝 合同"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!.AddCustomXmlPart("application/octet-stream");
            var opaque = new byte[257]; new Random(seed ^ variant).NextBytes(opaque);
            using var payload = new MemoryStream(opaque); part.FeedData(payload);
        }
        return stream.ToArray();
    }

    private static byte[] Repack(byte[] bytes, int step)
    {
        using var stream = new MemoryStream(); stream.Write(bytes);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
            foreach (var entry in zip.Entries) entry.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero).AddMinutes(step);
        return stream.ToArray();
    }

    private static SortedDictionary<string, byte[]> Entries(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var entries = new SortedDictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries.Where(e => !e.FullName.EndsWith('/')))
        {
            using var input = entry.Open(); using var content = new MemoryStream(); input.CopyTo(content);
            entries.Add(entry.FullName, content.ToArray());
        }
        return entries;
    }
    private static bool Equivalent(byte[] left, byte[] right)
    {
        var a = Entries(left); var b = Entries(right);
        return a.Count == b.Count && a.All(p => b.TryGetValue(p.Key, out var bytes) && p.Value.AsSpan().SequenceEqual(bytes));
    }
    private static void SameContent(byte[] expected, byte[] actual) => Assert.True(Equivalent(expected, actual), "Independent uncompressed ZIP-entry oracle differs.");
    private static int Setting(string name, int fallback, int maximum)
    {
        var text = Environment.GetEnvironmentVariable("DOCXODUS_HISTORY_FUZZ_" + name);
        if (text is null) return fallback;
        return int.TryParse(text, out var value) && value > 0 && value <= maximum ? value
            : throw new ArgumentException($"Invalid history fuzz {name}; expected 1..{maximum}.");
    }
}
