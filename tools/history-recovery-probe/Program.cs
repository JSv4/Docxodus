// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using Docxodus;
using Docxodus.History;

// Test-only subprocess. The parent supplies a private existing test directory. No server,
// network protocol, product transport, or process-local recovery state is involved.
if (args.Length != 4 || args[0] is not ("seed" or "crash" or "recover")
    || args[2] is not ("version" or "restore" or "text" or "conflict") || args[3] is not ("before" or "after"))
    throw new ArgumentException("Usage: HistoryRecoveryProbe seed|crash|recover EXISTING_TEST_ROOT version|restore|text|conflict before|after");
if (JsonSerializer.IsReflectionEnabledByDefault) throw new Exception("This probe must run with reflection serialization disabled.");
var root = Path.GetFullPath(args[1]);
if (!Directory.Exists(root)) throw new ArgumentException("Parent must create the private test directory.");
var kind = args[2]; var boundary = args[3];
var blobs = new FileHistoryBlobStore(Path.Combine(root, "blobs"));
IHistoryHeadStore heads = new FileHistoryHeadStore(Path.Combine(root, "heads"));
var history = new DocxVersionHistory(blobs, heads);
var scenarioPath = Path.Combine(root, "scenario.json");
if (args[0] == "seed")
{
    var first = await history.CreateVersionAsync("doc", "initial", null, Document("abcd"), Metadata("initial"));
    var winner = await history.SubmitOperationAsync("doc", Text("winner", first.Head, 1, 2, "X"));
    var scenario = new Scenario(first, winner.View, Document("version after crash"));
    await File.WriteAllBytesAsync(scenarioPath, JsonSerializer.SerializeToUtf8Bytes(scenario, ProbeJson.Default.Scenario));
    Console.WriteLine("seeded reflection-disabled history");
    return;
}
var saved = JsonSerializer.Deserialize(await File.ReadAllBytesAsync(scenarioPath), ProbeJson.Default.Scenario)
    ?? throw new Exception("Missing scenario.");
if (args[0] == "crash")
{
    history = new DocxVersionHistory(blobs, new CrashHeads(heads, boundary));
    _ = await Submit(history);
    throw new Exception("Crash boundary unexpectedly returned.");
}

var observed = (await history.ReadAsync("doc"))!;
Check(observed.Head.Revision == saved.Seed.Head.Revision + (boundary == "after" ? 1 : 0), "Crash visibility differs.");
if (boundary == "before") Check(observed.Head == saved.Seed.Head, "Pre-commit crash changed the head.");
var recovered = await Submit(history);
Check(recovered.Head.Revision == saved.Seed.Head.Revision + 1, "Request did not publish exactly once.");
if (boundary == "after") Check(recovered.Head == observed.Head, "Retry did not recover exact original publication.");
var exact = await history.ExportVersionAsync("doc", recovered.Version.Id);
var expectedText = kind switch { "version" => "version after crash", "restore" => "abcd", "text" => "aXd!", _ => "aXd" };
Check(Body(exact) == expectedText, "Recovered document is incorrect.");
Check(recovered.State.Epoch == (kind == "restore" ? 1 : 0), "Restore epoch differs.");
var later = await history.CreateVersionAsync("doc", recovered.Head, exact, Metadata("later legacy label"));
Check((await Submit(new DocxVersionHistory(blobs, heads))).Head == recovered.Head, "Old retry returned a new/latest publication.");
Check((await history.ReadAsync("doc"))!.Head == later.Head, "Retry moved current head.");
Check((await history.ListVersionsAsync("doc")).Versions.Count == (kind == "conflict" ? 3 : 4), "Duplicate or missing version.");
var decisions = await history.ReadOperationsSinceAsync("doc", null);
Check(decisions.Operations.Count == (kind is "text" or "conflict" ? 2 : 1), "Duplicate or missing operation outcome.");
if (kind is "text" or "conflict")
{
    var operation = decisions.Operations[^1];
    Check(operation.Record.Status == (kind == "conflict" ? "conflict" : "accepted"), "Recovered decision differs.");
    Check(Body(await history.ExportOperationProposalAsync("doc", operation.Id)) == (kind == "conflict" ? "ab?d" : "abcd!"),
        "Original contender was lost.");
}
var replay = await history.ReplayAsync("doc", recovered.State.Sequence);
var expectedEntries = Entries(exact); var replayEntries = Entries(replay);
Check(expectedEntries.Count == replayEntries.Count && expectedEntries.All(p => replayEntries.TryGetValue(p.Key, out var b)
    && p.Value.AsSpan().SequenceEqual(b)), "Recorded-effect replay differs from independently exported package entries.");
// Round-trip newly generated output graphs in the reflection-disabled process too.
var roundTrip = JsonSerializer.Deserialize(JsonSerializer.SerializeToUtf8Bytes(decisions, ProbeJson.Default.DocxOperationUpdate),
    ProbeJson.Default.DocxOperationUpdate)!;
Check(roundTrip.View.Head == decisions.View.Head && roundTrip.Operations.Count == decisions.Operations.Count, "Generated graph round-trip differs.");
Console.WriteLine($"recovered {kind} {boundary}: exact original result, immutable versions, decisions, proposal, replay; reflection disabled");

async Task<DocxHistoryView> Submit(DocxVersionHistory service) => kind switch
{
    "version" => await service.CreateVersionAsync("doc", "request-version", saved.Seed.Head, saved.Candidate, Metadata("request")),
    "restore" => await service.RestoreVersionAsync("doc", "request-restore", saved.Seed.Head, saved.First.Version.Id, Metadata("request")),
    "text" => (await service.SubmitOperationAsync("doc", Text("request-text", saved.First.Head, 4, 0, "!"))).View,
    _ => (await service.SubmitOperationAsync("doc", Text("request-conflict", saved.First.Head, 2, 1, "?"))).View,
};
static DocxVersionMetadata Metadata(string label) => new()
{ Author = "process-probe", CreatedAt = DateTimeOffset.Parse("2026-01-01T12:00:00Z"), Label = label };
static DocxOperationRequest Text(string id, HistoryHead baseline, int offset, int delete, string insert) => new()
{ RequestId = id, Base = baseline, Kind = "text", Metadata = Metadata(id), Text = new DocxTextSplice("/word/document.xml", 0, offset, delete, insert) };
static byte[] Document(string text)
{
    using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
    var anchor = session.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:", StringComparison.Ordinal));
    Check(session.ReplaceText(anchor, text).Success, "Fixture edit failed.");
    return session.Save();
}
static string Body(byte[] bytes) => XDocument.Parse(Encoding.UTF8.GetString(Entries(bytes)["word/document.xml"]))
    .Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).First().Value;
static Dictionary<string, byte[]> Entries(byte[] bytes)
{
    using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
    return zip.Entries.ToDictionary(e => e.FullName, e =>
    { using var stream = e.Open(); using var buffer = new MemoryStream(); stream.CopyTo(buffer); return buffer.ToArray(); });
}
static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }

internal sealed record Scenario(DocxHistoryView First, DocxHistoryView Seed, byte[] Candidate);

internal sealed class CrashHeads(IHistoryHeadStore inner, string boundary) : IHistoryHeadStore
{
    public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default) => inner.ReadAsync(id, cancellationToken);
    public async ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state,
        CancellationToken cancellationToken = default)
    {
        if (boundary == "before") await KillAsync();
        var result = await inner.TryAdvanceAsync(id, expected, state, cancellationToken);
        if (result is null) throw new Exception("Unexpected competing publication in process probe.");
        await KillAsync();
        return result;
    }
    private async Task KillAsync()
    {
        Console.WriteLine("terminating-at-" + boundary + "-cas"); Console.Out.Flush();
        Process.GetCurrentProcess().Kill();
        await Task.Delay(Timeout.Infinite); // Never report success or dispose the producer normally.
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true)]
[JsonSerializable(typeof(Scenario))]
[JsonSerializable(typeof(DocxOperationUpdate))]
internal partial class ProbeJson : JsonSerializerContext { }
