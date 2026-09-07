#nullable enable

using System.Text.Json;
using Docxodus.History;
using Docxodus.Internal;

namespace Docxodus.PyHost;

/// <summary>History over existing local stdio dispatch; no new transit or network layer.</summary>
internal static class HistoryBindings
{
    private static readonly Dictionary<int, HistoryClientOps> Clients = new();
    private static int _nextHandle;

    public static string Open(JsonElement args)
    {
        var root = args.TryGetProperty("root", out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString() : null;
        IHistoryBlobStore blobs;
        IHistoryHeadStore heads;
        if (root is null)
        {
            blobs = new MemoryHistoryBlobStore(); heads = new MemoryHistoryHeadStore();
        }
        else
        {
            if (!Path.IsPathFullyQualified(root)) throw new ArgumentException("History storage root must be an absolute host-owned path.");
            blobs = new FileHistoryBlobStore(Path.Combine(root, "blobs"));
            heads = new FileHistoryHeadStore(Path.Combine(root, "heads"));
        }
        var handle = checked(++_nextHandle);
        Clients.Add(handle, new HistoryClientOps(blobs, heads));
        return handle.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    public static string Close(JsonElement args)
    {
        var handle = args.GetProperty("handle").GetInt32();
        if (Clients.TryGetValue(handle, out var client)) client.Dispose();
        Clients.Remove(handle);
        return "null";
    }

    public static string OpenArchive(JsonElement args)
    {
        HistoryClientOps? client = null;
        try
        {
            client = HistoryClientOps.OpenArchiveAsync(Decode(args.GetProperty("docxB64"), HistoryClientOps.MaxArchiveBytes))
                .GetAwaiter().GetResult();
            var handle = checked(++_nextHandle);
            var json = HistoryClientJson.Write(new HistoryClientResult { Handle = handle, Archive = client.ArchiveInfo });
            Clients.Add(handle, client); client = null; return json;
        }
        catch (Exception error) when (HistoryClientOps.IsClientError(error)) { return HistoryClientOps.Failure(error); }
        finally { client?.Dispose(); }
    }

    public static string Invoke(JsonElement args)
    {
        if (!Clients.TryGetValue(args.GetProperty("handle").GetInt32(), out var client))
            throw new ArgumentException("Unknown history client handle.");
        byte[]? bytes = null;
        if (args.TryGetProperty("docxB64", out var data))
        {
            var importing = args.GetProperty("request").GetProperty("operation").GetString() == "importArchive";
            bytes = Decode(data, importing ? HistoryClientOps.MaxArchiveBytes : 256 * 1024 * 1024);
        }
        return client.InvokeAsync(args.GetProperty("request").GetRawText(), bytes).GetAwaiter().GetResult();
    }

    private static byte[] Decode(JsonElement value, int limit)
    {
        var encoded = value.GetString() ?? throw new ArgumentException("Binary base64 must not be null.");
        if ((long)encoded.Length > ((long)limit + 2) / 3 * 4)
            throw new PackageChangeException(PackageChangeError.ResourceLimit, "Binary input exceeds its byte limit.");
        return Convert.FromBase64String(encoded);
    }

    public static void CloseAll() { foreach (var client in Clients.Values) client.Dispose(); Clients.Clear(); }
}
