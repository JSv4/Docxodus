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
        Clients.Remove(args.GetProperty("handle").GetInt32());
        return "null";
    }

    public static string Invoke(JsonElement args)
    {
        if (!Clients.TryGetValue(args.GetProperty("handle").GetInt32(), out var client))
            throw new ArgumentException("Unknown history client handle.");
        byte[]? bytes = null;
        if (args.TryGetProperty("docxB64", out var data))
        {
            var encoded = data.GetString() ?? throw new ArgumentException("DOCX base64 must not be null.");
            if ((long)encoded.Length > ((256L * 1024 * 1024 + 2) / 3) * 4)
                throw new ArgumentException("DOCX input exceeds the byte limit.");
            bytes = Convert.FromBase64String(encoded);
        }
        return client.InvokeAsync(args.GetProperty("request").GetRawText(), bytes).GetAwaiter().GetResult();
    }

    public static void CloseAll() => Clients.Clear();
}
