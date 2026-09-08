#nullable enable

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using Docxodus.History;
using Docxodus.Internal;

namespace DocxodusWasm;

/// <summary>WASM history binding; the host supplies storage, never a network transport here.</summary>
[SupportedOSPlatform("browser")]
public static partial class HistoryBridge
{
    private sealed class Client(HistoryClientOps ops)
    {
        public HistoryClientOps Ops { get; } = ops;
        public int Active { get; set; }
    }
    private static readonly Dictionary<int, Client> Clients = new();
    private static int _nextHandle;

    [JSExport]
    public static int Open(int adapterId)
    {
        var storage = new JsStorage(adapterId);
        var handle = checked(++_nextHandle);
        Clients.Add(handle, new Client(new HistoryClientOps(storage, storage)));
        return handle;
    }

    [JSExport]
    public static int OpenWithInitialization(int adapterId)
    {
        var storage = new InitializableJsStorage(adapterId);
        var handle = checked(++_nextHandle);
        Clients.Add(handle, new Client(new HistoryClientOps(storage, storage)));
        return handle;
    }

    [JSExport]
    public static async Task<string> OpenArchive(byte[] bytes)
    {
        HistoryClientOps? ops = null;
        try
        {
            ops = await HistoryClientOps.OpenArchiveAsync(bytes);
            var handle = checked(++_nextHandle);
            var result = HistoryClientJson.Write(new HistoryClientResult { Handle = handle, Archive = ops.ArchiveInfo });
            Clients.Add(handle, new Client(ops)); ops = null;
            return result;
        }
        catch (Exception error) when (HistoryClientOps.IsClientError(error)) { return HistoryClientOps.Failure(error); }
        finally { ops?.Dispose(); }
    }

    [JSExport]
    public static void Close(int handle)
    {
        if (!Clients.TryGetValue(handle, out var client)) return;
        if (client.Active != 0) throw new InvalidOperationException("History client has pending calls; await them before closing.");
        client.Ops.Dispose();
        Clients.Remove(handle);
    }

    [JSExport]
    public static async Task<string> Invoke(int handle, string requestJson, byte[] bytes)
    {
        if (!Clients.TryGetValue(handle, out var client)) throw new ArgumentException("Unknown history client handle.");
        client.Active++;
        try
        {
            if (HistoryClientJson.Read<HistoryClientRequest>(requestJson).Operation == "compare") ComparisonEngine.EnsureWarm();
            return await client.Ops.InvokeAsync(requestJson, bytes);
        }
        catch (Exception error) when (HistoryClientOps.IsClientError(error)) { return HistoryClientOps.Failure(error); }
        finally { client.Active--; }
    }

    [JSImport("readBlob", "docxodus.history")]
    private static partial Task<string> ReadBlob(int adapterId, string referenceJson);
    [JSImport("putBlob", "docxodus.history")]
    private static partial Task PutBlob(int adapterId, string referenceJson, byte[] bytes);
    [JSImport("readHead", "docxodus.history")]
    private static partial Task<string> ReadHead(int adapterId, string documentId);
    [JSImport("advanceHead", "docxodus.history")]
    private static partial Task<string> AdvanceHead(int adapterId, string documentId, string expectedJson, string stateJson);

    [JSImport("initializeHead", "docxodus.history")]
    private static partial Task<string> InitializeHead(int adapterId, string documentId, string headJson);

    private class JsStorage(int adapterId) : IHistoryBlobStore, IHistoryHeadStore
    {
        protected int AdapterId { get; } = adapterId;
        public async ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var encoded = await ReadBlob(AdapterId, HistoryClientJson.Write(reference));
            if (encoded == "") return null;
            // Reject an oversized adapter reply before allocating its decoded bytes. Core verifies
            // the exact length and hash before trusting any payload. Empty blobs encode as "=".
            if (encoded == "=") return new MemoryStream(Array.Empty<byte>(), writable: false);
            if ((long)encoded.Length > ((long)reference.Length + 2) / 3 * 4)
                throw new IOException("History adapter returned an oversized blob.");
            return new MemoryStream(Convert.FromBase64String(encoded), writable: false);
        }

        public async ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var buffer = new MemoryStream(reference.Length);
            await HistoryBlobIO.CopyVerifiedAsync(reference, content, buffer, cancellationToken);
            await PutBlob(AdapterId, HistoryClientJson.Write(reference), buffer.ToArray());
        }

        public async ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var json = await ReadHead(AdapterId, documentId);
            return json == "null" ? null : HistoryClientJson.Read<HistoryHead>(json);
        }

        public async ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
            HistoryBlobReference state, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var json = await AdvanceHead(AdapterId, documentId, HistoryClientJson.Write(expected), HistoryClientJson.Write(state));
            return json == "null" ? null : HistoryClientJson.Read<HistoryHead>(json);
        }
    }

    private sealed class InitializableJsStorage(int adapterId) : JsStorage(adapterId), IHistoryHeadInitializer
    {
        public async ValueTask<HistoryHeadInitializationResult> TryInitializeAsync(string documentId, HistoryHead head,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return HistoryClientJson.Read<HistoryHeadInitializationResult>(
                await InitializeHead(AdapterId, documentId, HistoryClientJson.Write(head)));
        }
    }
}
