#nullable enable

using System.Text.Json;
using Docxodus.History;
using Docxodus.Internal;
using Docxodus.McpServer;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public sealed class McpHistoryTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "mcp-history-" + Guid.NewGuid().ToString("N"));
    private readonly SessionStore _store;
    private readonly DocSession _session;
    private readonly byte[] _source = DocxSession.CreateBlankDocxBytes();
    private const string Metadata = "\"metadata\":{\"author\":\"actor\",\"createdAt\":\"2026-01-01T12:00:00Z\"}";

    public McpHistoryTests()
    {
        Directory.CreateDirectory(_root);
        var documents = new LocalFileDocumentStore(_root);
        _store = new SessionStore(documents, history: HistoryTool.Configure(Path.Combine(_root, "history")));
        var path = documents.Resolve("document.docx");
        documents.Write(path, _source);
        _session = _store.Open(_source, path, new DocxSessionSettings());
    }

    [Fact]
    public void ImportRejectsUnknownMetadataBeforeCopyingItsValue()
    {
        var args = J("{\"action\":\"importArchive\",\"archiveB64\":\"AA==\",\"unknown\":" + Q(new string('x', 2 * 1024 * 1024)) + "}");
        var error = Assert.Throws<McpToolException>(() => HistoryTool.Execute(_store, _session, args));
        Assert.Contains("accepts only", error.Message);
        Assert.Null(Result(Call("read")).View);
    }

    [Fact]
    public async Task ScopedOperationReadsAndComparisonExposePreservedWorkWithoutFilesystemWrites()
    {
        var history = new DocxVersionHistory(new FileHistoryBlobStore(Path.Combine(_root, "history", "blobs")),
            new FileHistoryHeadStore(Path.Combine(_root, "history", "heads")));
        var original = await File.ReadAllBytesAsync(Path.Combine(DocxHistoryArchiveArtifactTests.Root, "agreement-v1.docx"));
        var revised = await File.ReadAllBytesAsync(Path.Combine(DocxHistoryArchiveArtifactTests.Root, "agreement-v2.docx"));
        var first = await history.CreateVersionAsync(_session.Location!, null, original, HistoryArchiveFixture.Metadata("Original"));
        var winner = await history.SubmitOperationAsync(_session.Location!, new DocxOperationRequest { RequestId = "winner", Base = first.Head,
            Kind = "package", Metadata = HistoryArchiveFixture.Metadata("Winner") }, revised);
        var proposal = HistoryArchiveFixture.Edit(original, " Conflicting counsel draft");
        var conflict = await history.SubmitOperationAsync(_session.Location!, new DocxOperationRequest { RequestId = "conflict", Base = first.Head,
            Kind = "package", Metadata = HistoryArchiveFixture.Metadata("Conflict") }, proposal);
        Assert.Equal("conflict", conflict.Operation.Record.Status);
        var files = Directory.GetFiles(_root, "*", SearchOption.AllDirectories).OrderBy(p => p).ToArray();
        Assert.Equal(2, Result(Call("operations")).OperationUpdate!.Operations.Count);
        var args = "\"operationId\":" + HistoryClientJson.Write(conflict.Operation.Id);
        Assert.Equal("conflict", Result(Call("getOperation", args)).Operation!.Record.Status);
        Assert.Equal(proposal, Call("exportOperationProposal", args).GetProperty("bytes").GetBytesFromBase64());
        var comparison = Call("compare", "\"beforeVersionId\":" + HistoryClientJson.Write(first.Version.Id)
            + ",\"afterVersionId\":" + HistoryClientJson.Write(winner.View.Version.Id)).GetProperty("bytes").GetBytesFromBase64();
        using var session = new DocxSession(comparison); Assert.NotNull(session);
        Assert.Equal(files, Directory.GetFiles(_root, "*", SearchOption.AllDirectories).OrderBy(p => p));
        Assert.Equal(_source, await File.ReadAllBytesAsync(_session.Location!));
    }

    [Fact]
    public async Task ArchiveImportIsBoundToSessionIdentityAndNeverReplacesEditorOrSource()
    {
        var foreign = await File.ReadAllBytesAsync(Path.Combine(DocxHistoryArchiveArtifactTests.Root, "agreement.docxhistory"));
        Assert.Equal("ForeignDocument", Result(Call("importArchive", "\"archiveB64\":" + Q(Convert.ToBase64String(foreign)), success: false)).ErrorCode);
        Assert.Null(Result(Call("read")).View);
        var sourceHistory = new DocxVersionHistory(new MemoryHistoryBlobStore(), new MemoryHistoryHeadStore());
        var initial = await sourceHistory.CreateVersionAsync(_session.Location!, null, _source,
            HistoryArchiveFixture.Metadata("Scoped"));
        var latest = await sourceHistory.CreateVersionAsync(_session.Location!, initial.Head, _source,
            HistoryArchiveFixture.Metadata("Label"));
        using var buffer = new MemoryStream(); await sourceHistory.ExportHistoryArchiveAsync(_session.Location!, buffer);
        ReplaceText("Unsaved draft must survive import"); var before = DocxSessionOps.Project(_session.Handle);
        var args = "\"archiveB64\":" + Q(Convert.ToBase64String(buffer.ToArray()));
        var imported = Result(Call("importArchive", args)).Import!;
        Assert.Equal(latest.Head, imported.View.Head); Assert.False(imported.AlreadyPresent);
        Assert.True(Result(Call("importArchive", args)).Import!.AlreadyPresent);
        Assert.Equal(_source, Result(Call("exportDocx")).Bytes);
        using var reopened = await DocxHistoryArchive.OpenAsync(Result(Call("exportArchive")).Bytes!);
        Assert.Equal(latest.Head, reopened.View.Head);
        Assert.Equal(before, DocxSessionOps.Project(_session.Handle));
        Assert.Equal(_source, await File.ReadAllBytesAsync(_session.Location!));
    }

    [Fact]
    public void RestoreRequestIdsRetryOriginalResultButCreateCannotRecaptureRetryBytes()
    {
        Assert.Throws<McpToolException>(() => Call("create", Metadata + ",\"requestId\":\"unsafe-recapture\""));
        Assert.Null(Result(Call("read")).View);
        var first = Result(Call("create", Metadata)).View!;
        var args = Metadata + ",\"requestId\":\"restore-id\",\"expectedHead\":" + HistoryClientJson.Write(first.Head)
            + ",\"versionId\":" + HistoryClientJson.Write(first.Version.Id);
        var restored = Result(Call("restore", args)).View!;
        var later = Result(Call("create", Metadata + ",\"expectedHead\":" + HistoryClientJson.Write(restored.Head))).View!;
        Assert.Equal(restored.Head, Result(Call("restore", args)).View!.Head);
        Assert.Equal(later.Head, Result(Call("read")).View!.Head);
        Assert.Equal("restore-id", restored.State.Requests!.Current!.Id);
    }

    [Fact]
    public void PublishReadRenderReplayRestoreLeavesLocalWorkAndSourceUntouched()
    {
        var first = Result(Call("create", Metadata)).View!;
        var firstBytes = Call("export", "\"versionId\":" + HistoryClientJson.Write(first.Version.Id))
            .GetProperty("bytes").GetBytesFromBase64();
        Assert.Equal(first.Head, Result(Call("read")).View!.Head);
        Assert.Equal(first.Version.Id, Assert.Single(Result(Call("list")).Page!.Versions).Id);
        Assert.Equal("actor", Result(Call("get", "\"versionId\":" + HistoryClientJson.Write(first.Version.Id)))
            .Version!.Record.Metadata.Author);
        Assert.Equal("0", Call("resolveTime", "\"cutoff\":\"2026-01-01T12:00:00Z\"").GetProperty("sequence").GetString());
        Assert.Equal(firstBytes, Call("materialize", "\"sequence\":\"0\"").GetProperty("bytes").GetBytesFromBase64());
        Assert.Equal(firstBytes, Call("replay", "\"sequence\":\"0\"").GetProperty("bytes").GetBytesFromBase64());

        ReplaceText("New local draft");
        var second = Result(Call("create", Metadata + ",\"expectedHead\":" + HistoryClientJson.Write(first.Head))).View!;
        Assert.Equal(1, second.State.Sequence);
        var update = Result(Call("updates", "\"expectedHead\":" + HistoryClientJson.Write(first.Head))).Update!;
        Assert.False(update.Reset);
        Assert.Equal(1, Assert.Single(update.Entries).Commit.Sequence);
        Assert.Empty(Result(Call("updates", "\"expectedHead\":" + HistoryClientJson.Write(second.Head))).Update!.Entries);
        var html = Call("render", "\"sequence\":\"1\"");
        Assert.Equal("1", html.GetProperty("sequence").GetString());
        Assert.Contains("New local draft", html.GetProperty("html").GetString());
        Assert.Equal("1", Call("render", "\"cutoff\":\"2026-01-01T12:00:00Z\"").GetProperty("sequence").GetString());
        Assert.DoesNotContain("New local draft", Call("render", "\"sequence\":\"0\"").GetProperty("html").GetString());

        var beforeRestore = DocxSessionOps.Project(_session.Handle);
        var restored = Result(Call("restore", Metadata + ",\"expectedHead\":" + HistoryClientJson.Write(second.Head)
            + ",\"versionId\":" + HistoryClientJson.Write(first.Version.Id))).View!;
        Assert.Equal(2, restored.State.Sequence);
        Assert.Equal(1, restored.State.Epoch);
        Assert.Equal(firstBytes, Call("materialize", "\"sequence\":\"2\"").GetProperty("bytes").GetBytesFromBase64());
        Assert.Equal(beforeRestore, DocxSessionOps.Project(_session.Handle));
        Assert.Equal(_source, File.ReadAllBytes(_session.Location!));
        Assert.True(Result(Call("updates", "\"expectedHead\":" + HistoryClientJson.Write(second.Head))).Update!.Reset);
        var stale = Result(Call("create", Metadata + ",\"expectedHead\":" + HistoryClientJson.Write(second.Head), success: false));
        Assert.Equal("StaleHead", stale.ErrorCode);
        Assert.Equal(restored.Head, Result(Call("read")).View!.Head);
    }

    [Fact]
    public void FileHistoryReopensButSavingACopyDoesNotRetargetExistingSession()
    {
        var first = Result(Call("create", Metadata)).View!;
        Dispatcher.Call(_store, "docxodus_save", J("{\"sessionId\":" + Q(_session.Id) + ",\"path\":\"copy.docx\"}"));
        Assert.Equal(first.Head, Result(Call("read")).View!.Head);
        var reopenedStore = new SessionStore(_store.Documents, history: HistoryTool.Configure(Path.Combine(_root, "history")));
        try
        {
            var reopened = reopenedStore.Open(_source, _session.Location, new DocxSessionSettings());
            var result = Dispatcher.Call(reopenedStore, "docxodus_history", J("{\"sessionId\":" + Q(reopened.Id) + ",\"action\":\"read\"}"));
            Assert.Equal(first.Head, HistoryClientJson.Read<HistoryClientResult>(result).View!.Head);
            var copy = reopenedStore.Open(_source, _store.Documents.Resolve("copy.docx"), new DocxSessionSettings());
            result = Dispatcher.Call(reopenedStore, "docxodus_history", J("{\"sessionId\":" + Q(copy.Id) + ",\"action\":\"read\"}"));
            Assert.Null(HistoryClientJson.Read<HistoryClientResult>(result).View);
        }
        finally { reopenedStore.CloseAll(); }
    }

    [Fact]
    public void SessionCapabilityAndDocumentIdentityCannotBeOverridden()
    {
        var first = Result(Call("create", Metadata)).View!;
        var other = _store.Open(_source, _store.Documents.Resolve("other.docx"), new DocxSessionSettings());
        var foreign = Dispatcher.Call(_store, "docxodus_history", J("{\"sessionId\":" + Q(other.Id)
            + ",\"action\":\"export\",\"versionId\":" + HistoryClientJson.Write(first.Version.Id) + "}"));
        Assert.Equal("ForeignDocument", HistoryClientJson.Read<HistoryClientResult>(foreign).ErrorCode);
        foreach (var field in new[] { "documentId", "operation", "schemaVersion" })
            Assert.Throws<McpToolException>(() => Call("read", Q(field) + ":\"override\""));
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_history",
            J("{\"sessionId\":\"unknown\",\"action\":\"read\"}")));
        var noOrigin = _store.Open(_source, null, new DocxSessionSettings());
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_history",
            J("{\"sessionId\":" + Q(noOrigin.Id) + ",\"action\":\"read\"}")));
        _store.Close(_session.Id);
        Assert.Throws<McpToolException>(() => Call("read"));
    }

    [Fact]
    public void DisabledRelativeRootAndMalformedRequestsFailWithoutPublication()
    {
        Assert.Null(HistoryTool.Configure(null));
        Assert.Null(HistoryTool.Configure(" "));
        Assert.Throws<McpToolException>(() => HistoryTool.Configure("relative"));
        var disabled = new SessionStore(_store.Documents);
        try
        {
            var session = disabled.Open(_source, _session.Location, new DocxSessionSettings());
            var failure = Assert.Throws<McpToolException>(() => Dispatcher.Call(disabled, "docxodus_history",
                J("{\"sessionId\":" + Q(session.Id) + ",\"action\":\"read\"}")));
            Assert.Contains(HistoryTool.RootVariable, failure.Message);
        }
        finally { disabled.CloseAll(); }
        Assert.Throws<McpToolException>(() => Call("read", "\"action\":\"create\""));
        Assert.Throws<McpToolException>(() => Call("read", "\"sessionId\":" + Q(_session.Id)));
        Assert.Throws<JsonException>(() => Call("read", "\"root\":\"/tmp/not-authorized\""));
        Assert.Throws<JsonException>(() => Call("materialize", "\"sequence\":0"));
        Assert.Throws<McpToolException>(() => Call("render"));
        Assert.Throws<McpToolException>(() => Call("render", "\"sequence\":\"0\",\"cutoff\":\"2026-01-01T12:00:00Z\""));
        Assert.Equal("InvalidRequest", Result(Call("create", success: false)).ErrorCode);
        Assert.Null(Result(Call("read")).View);
        Call("create", Metadata);
        Assert.Equal("HistoryUnavailable", Result(Call("render", "\"cutoff\":\"2025-01-01T12:00:00Z\"", success: false)).ErrorCode);
    }

    [Fact]
    public void RenderingLargeCheckpointDoesNotApplyMetadataLimitToBinaryResult()
    {
        using var stream = new MemoryStream();
        stream.Write(_source);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!.AddImagePart(ImagePartType.Png);
            var opaque = new byte[450_000];
            new Random(42).NextBytes(opaque);
            using var payload = new MemoryStream(opaque);
            part.FeedData(payload); // Opaque, unreferenced asset; renderer need not decode it.
        }
        var large = _store.Open(stream.ToArray(), _store.Documents.Resolve("large.docx"), new DocxSessionSettings());
        var request = "{\"sessionId\":" + Q(large.Id) + ",\"action\":\"create\"," + Metadata + "}";
        Assert.True(HistoryClientJson.Read<HistoryClientResult>(Dispatcher.Call(_store, "docxodus_history", J(request))).Success);
        request = "{\"sessionId\":" + Q(large.Id) + ",\"action\":\"materialize\",\"sequence\":\"0\"}";
        Assert.True(Dispatcher.Call(_store, "docxodus_history", J(request)).Length > HistoryClientJson.MaxRequestChars);
        request = "{\"sessionId\":" + Q(large.Id) + ",\"action\":\"render\",\"sequence\":\"0\"}";
        var result = J(Dispatcher.Call(_store, "docxodus_history", J(request)));
        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Contains("html", result.GetProperty("html").GetString());
    }

    private JsonElement Call(string action, string? extra = null, bool success = true)
    {
        var args = "{\"sessionId\":" + Q(_session.Id) + ",\"action\":" + Q(action)
            + (extra is null ? "" : "," + extra) + "}";
        var result = J(Dispatcher.Call(_store, "docxodus_history", J(args)));
        Assert.Equal(success, result.GetProperty("success").GetBoolean());
        return result;
    }

    private void ReplaceText(string text)
    {
        var blocks = J(Dispatcher.Call(_store, "docxodus_get_content",
            J("{\"sessionId\":" + Q(_session.Id) + ",\"format\":\"blocks\"}"))).GetProperty("blocks");
        var anchor = blocks.EnumerateObject().First().Name;
        var result = Dispatcher.Call(_store, "docxodus_edit", J("{\"sessionId\":" + Q(_session.Id)
            + ",\"action\":\"replace_text\",\"anchorId\":" + Q(anchor) + ",\"markdown\":" + Q(text) + "}"));
        Assert.True(J(result).GetProperty("success").GetBoolean());
    }

    private static HistoryClientResult Result(JsonElement value) => HistoryClientJson.Read<HistoryClientResult>(value.GetRawText());
    private static string Q(string value) => JsonSerializer.Serialize(value);
    private static JsonElement J(string json) { using var doc = JsonDocument.Parse(json); return doc.RootElement.Clone(); }
    public void Dispose() { _store.CloseAll(); Directory.Delete(_root, recursive: true); }
}
