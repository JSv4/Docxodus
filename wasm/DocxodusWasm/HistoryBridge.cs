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
    public static void Close(int handle)
    {
        if (!Clients.TryGetValue(handle, out var client)) return;
        if (client.Active != 0) throw new InvalidOperationException("History client has pending calls; await them before closing.");
        Clients.Remove(handle);
    }

    [JSExport]
    public static async Task<string> Invoke(int handle, string requestJson, byte[] bytes)
    {
        if (!Clients.TryGetValue(handle, out var client)) throw new ArgumentException("Unknown history client handle.");
        client.Active++;
        try { return await client.Ops.InvokeAsync(requestJson, bytes); }
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

    private sealed class JsStorage(int adapterId) : IHistoryBlobStore, IHistoryHeadStore
    {
        public async ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var encoded = await ReadBlob(adapterId, HistoryClientJson.Write(reference));
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
            await content.CopyToAsync(buffer, cancellationToken);
            await PutBlob(adapterId, HistoryClientJson.Write(reference), buffer.ToArray());
        }

        public async ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var json = await ReadHead(adapterId, documentId);
            return json == "null" ? null : HistoryClientJson.Read<HistoryHead>(json);
        }

        public async ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
            HistoryBlobReference state, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var json = await AdvanceHead(adapterId, documentId, HistoryClientJson.Write(expected), HistoryClientJson.Write(state));
            return json == "null" ? null : HistoryClientJson.Read<HistoryHead>(json);
        }
    }
}
