#nullable enable

using System.Xml.Linq;
using System.Text;
using Docxodus.History;
using Xunit;
using Xunit.Abstractions;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

public sealed class DocxBackendModelFuzzTests(ITestOutputHelper output)
{
    public static IEnumerable<object[]> Seeds() => Enumerable.Range(0, Setting("SEEDS", 8, 256)).Select(i => new object[] { 190739 + i * 7919 });

    [Theory]
    [MemberData(nameof(Seeds))]
    public async Task ReorderedBackendStreamsMatchIndependentObservedCharacterModel(int seed)
    {
        var rounds = Setting("ROUNDS", 8, 1024); var random = new Random(seed);
        using var store = new HistoryFaultHarness(seed % 4 == 0);
        var history = new DocxVersionHistory(store, store);
        DocxHistoryView? current = null;
        var trace = new List<string>();
        var requests = new List<(DocxOperationRequest Request, DocxOperationResult Result)>();
        var versions = new List<(DocxHistoryView View, byte[] Bytes)>();
        var counts = new SortedDictionary<string, int>();
        int identities = 0;
        try
        {
            for (var round = 0; round < rounds; round++)
            {
                var template = $"abcdefghij café 📝 [{seed}:{round}]";
                var original = Package(template);
                current = await history.CreateVersionAsync("doc", "round-" + round, current?.Head, original, Metadata("round-" + round));
                versions.Add((current, original));
                var baseline = current;
                var observed = template.EnumerateRunes().Select(r => new Atom(identities++, r.ToString())).ToList();
                var model = new List<Atom>(observed);
                var commands = new List<DocxOperationRequest>
                {
                    Text($"{round}:a", baseline.Head, 2, 0, "A" + round),
                    Text($"{round}:b", baseline.Head, 2, 0, "B" + round),
                    Text($"{round}:delete", baseline.Head, 6, 2, ""),
                    Text($"{round}:replace", baseline.Head, 7, 1, "R" + round),
                    Text($"{round}:end", baseline.Head, template.Length, 0, "終"),
                    Text($"{round}:noop", baseline.Head, 4, 0, ""),
                }.Select(r => r with { Metadata = r.Metadata with
                { Author = "actor-" + random.Next(3), CreatedAt = DateTimeOffset.UnixEpoch.AddSeconds(random.Next(12)) } }).ToList();
                if (random.Next(2) == 0)
                {
                    var rightFirst = random.Next(2) == 0;
                    trace.Add($"round={round} race rightFirst={rightFirst}");
                    store.Arm(HistoryFaultHarness.Fault.None); store.HoldCompetingPublications(rightFirst ? "right" : "left");
                    DocxOperationResult[] raced;
                    try
                    {
                        raced = await Task.WhenAll(
                            Task.Run(() => store.AsContenderAsync("left", async () => await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", commands[0]))),
                            Task.Run(() => store.AsContenderAsync("right", async () => await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", commands[1]))));
                    }
                    finally { store.ReleaseCompetingPublications(); }
                    Assert.Equal(3, store.Publications); Count("forced-races");
                    foreach (var i in rightFirst ? new[] { 1, 0 } : new[] { 0, 1 }) await Check(commands[i], raced[i]);
                    commands.RemoveRange(0, 2);
                }
                // Fisher-Yates schedules genuine original-base requests, not successively recaptured edits.
                for (var i = commands.Count - 1; i > 0; i--)
                { var other = random.Next(i + 1); (commands[i], commands[other]) = (commands[other], commands[i]); }
                foreach (var request in commands)
                {
                    trace.Add($"round={round} submit {request.RequestId}");
                    history = new DocxVersionHistory(store, store);
                    if (random.Next(5) == 0)
                    {
                        store.Arm(HistoryFaultHarness.Fault.AfterHead);
                        await Assert.ThrowsAsync<IOException>(async () => await history.SubmitOperationAsync("doc", request));
                        store.Arm(HistoryFaultHarness.Fault.None); Count("lost-ack");
                    }
                    await Check(request, await history.SubmitOperationAsync("doc", request));
                    if (random.Next(3) == 0)
                    {
                        var old = requests[random.Next(requests.Count)];
                        Assert.Equal(old.Result.View.Head, (await new DocxVersionHistory(store, store).SubmitOperationAsync("doc", old.Request)).View.Head);
                        Count("old-retries");
                    }
                }
                var conflict = requests.Last(r => r.Request.Base == baseline.Head && r.Result.Operation.Record.Status == "conflict");
                var discard = new DocxOperationRequest { RequestId = round + ":discard", Base = current!.Head,
                    Kind = "discard", Metadata = Metadata("discard"), Resolves = conflict.Result.Operation.Id };
                var discarded = await history.SubmitOperationAsync("doc", discard);
                Assert.Equal("accepted", discarded.Operation.Record.Status);
                Assert.Equal(current.Version.Id, discarded.View.Version.Id);
                Assert.Equal(current.Head.Revision + 1, discarded.View.Head.Revision);
                current = discarded.View; requests.Add((discard, discarded)); Count("discarded-conflicts");
                var tail = await history.ReadOperationsSinceAsync("doc", baseline.Head);
                Assert.Equal(7, tail.Operations.Count);
                Assert.Equal(Enumerable.Range(1, 7).Select(i => baseline.Head.Revision + i), tail.Operations.Select(o => o.Record.Revision));
                Assert.Equal(current.Head, tail.View.Head);

                async Task Check(DocxOperationRequest request, DocxOperationResult result)
                {
                    var beforeText = string.Concat(model.Select(a => a.Text));
                    var splice = request.Text!;
                    // Independent oracle: delete only still-present observed character identities,
                    // insert before the original right-hand character. It uses no production maps,
                    // resolved offsets, package digests, or replay to calculate expected text.
                    var deleting = splice.DeleteCount == 0 ? new List<Atom>() : observed.Skip(splice.Offset).Take(splice.DeleteCount).ToList();
                    var overlaps = deleting.Any(atom => !model.Contains(atom));
                    if (!overlaps)
                    {
                        var at = deleting.Count > 0 ? model.IndexOf(deleting[0])
                            : splice.Offset == template.Length ? model.Count : model.IndexOf(observed[splice.Offset]);
                        Assert.True(at >= 0);
                        if (deleting.Count > 0) Assert.Equal(deleting, model.Skip(at).Take(deleting.Count));
                        model.RemoveAll(deleting.Contains);
                        model.InsertRange(at, splice.Insert.EnumerateRunes().Select(r => new Atom(identities++, r.ToString())));
                    }
                    var expected = string.Concat(model.Select(a => a.Text));
                    Assert.Equal(overlaps ? "conflict" : "accepted", result.Operation.Record.Status);
                    Assert.Equal(overlaps ? "OverlappingText" : null, result.Operation.Record.Conflict);
                    Assert.Equal(current!.Head.Revision + 1, result.View.Head.Revision);
                    var changed = beforeText != expected;
                    Assert.Equal(current.State.Sequence + (changed ? 1 : 0), result.View.State.Sequence);
                    if (!changed) Assert.Equal(current.Version.Id, result.View.Version.Id);
                    else Assert.Equal(current.Version.Id, result.View.Version.Record.Parent);
                    var exact = await history.ExportVersionAsync("doc", result.View.Version.Id);
                    Assert.Equal(expected, BodyText(exact)); Untouched(original, exact, "word/document.xml");
                    SameNonTargetBody(original, exact); NoNewValidationErrors(original, exact);
                    var proposal = template[..splice.Offset] + splice.Insert + template[(splice.Offset + splice.DeleteCount)..];
                    Assert.Equal(proposal, BodyText(await history.ExportOperationProposalAsync("doc", result.Operation.Id)));
                    if (changed) versions.Add((result.View, exact));
                    current = result.View; requests.Add((request, result)); Count(overlaps ? "conflicts" : changed ? "accepted-changes" : "accepted-noops");
                }
            }
            foreach (var version in versions)
            {
                Assert.Equal(version.Bytes, await history.ExportVersionAsync("doc", version.View.Version.Id));
                SameEntries(version.Bytes, await history.ReplayAsync("doc", version.View.State.Sequence));
                Count("replayed-versions");
            }
            foreach (var request in requests)
                Assert.Equal(request.Result.View.Head, (await history.SubmitOperationAsync("doc", request.Request)).View.Head);
            Count("end-retries", requests.Count);
            output.WriteLine($"seed={seed}; filesystem={seed % 4 == 0}; rounds={rounds}; " + string.Join(", ", counts.Select(p => p.Key + "=" + p.Value)));
        }
        catch (Exception error) { throw new Exception($"Backend seed={seed}; rounds={rounds}; trace:\n{string.Join("\n", trace)}", error); }
        void Count(string name, int amount = 1) => counts[name] = counts.GetValueOrDefault(name) + amount;
    }

    private sealed record Atom(int Id, string Text);
    private static void SameNonTargetBody(byte[] before, byte[] after)
    {
        static XElement Normalize(byte[] bytes)
        {
            var root = XElement.Parse(Encoding.UTF8.GetString(Entries(bytes)["word/document.xml"]));
            var target = root.Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).First();
            target.Value = "TEXT"; target.Attribute(XNamespace.Xml + "space")?.Remove();
            return root;
        }
        Assert.True(XNode.DeepEquals(Normalize(before), Normalize(after)), "Non-target body/review/relationship topology changed.");
    }
    private static int Setting(string name, int fallback, int maximum)
    {
        var value = Environment.GetEnvironmentVariable("DOCXODUS_BACKEND_FUZZ_" + name);
        if (value is null) return fallback;
        return int.TryParse(value, out var parsed) && parsed > 0 && parsed <= maximum ? parsed
            : throw new ArgumentException($"Invalid backend fuzz {name}; expected 1..{maximum}.");
    }
}
