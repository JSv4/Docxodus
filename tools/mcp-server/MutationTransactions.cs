#nullable enable

using System.Buffers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>The stable, versioned identity attached to a transaction-aware batch result.</summary>
internal sealed record MutationTransactionIdentity(
    int SchemaVersion,
    string TransactionId,
    string RequestFingerprint);

/// <summary>
/// One retained transaction result. Record ids and timestamps are server bookkeeping for tests,
/// diagnostics, and a future durable implementation; only <see cref="Identity"/> is placed on
/// the MCP response in this in-session epic.
/// </summary>
internal sealed record MutationTransactionRecord(
    string RecordId,
    MutationTransactionIdentity Identity,
    DateTimeOffset StartedAt,
    DateTimeOffset? CompletedAt,
    string? SerializedResponse);

/// <summary>
/// An identity that is still bound but whose response is gone. <see cref="CompletedAt"/> is the
/// discriminator: a value means a terminal response existed and was evicted; null means the
/// reservation never recorded one, so the mutation's outcome is unknown.
/// </summary>
internal sealed record MutationTransactionTombstone(
    string RecordId,
    MutationTransactionIdentity Identity,
    DateTimeOffset StartedAt,
    DateTimeOffset? CompletedAt,
    DateTimeOffset EvictedAt);

internal enum MutationTransactionDecisionKind
{
    Reserved,
    Replay,
    Conflict,
    ResultEvicted,

    /// <summary>The id is bound to this exact request, but no terminal response was ever
    /// recorded for it, so whether the mutation applied is unknown.</summary>
    Incomplete,
}

internal sealed record MutationTransactionDecision(
    MutationTransactionDecisionKind Kind,
    MutationTransactionRecord? Record = null,
    MutationTransactionIdentity? ExistingIdentity = null,
    string? SerializedResponse = null);

/// <summary>
/// Bounded, per-session transaction-id registry. Full responses and response-less tombstones use
/// independent FIFO limits. A tombstone keeps an evicted id bound to its original fingerprint for
/// a further window, preventing a recently forgotten retry from becoming a fresh mutation.
/// Retained responses are additionally bounded by a byte budget, because a batch result is
/// unbounded in size while a count is not a memory bound.
/// </summary>
internal sealed class MutationTransactions
{
    public const int SchemaVersion = 1;
    public const int DefaultFullRecordCapacity = 128;
    public const int DefaultTombstoneCapacity = 1024;
    public const int MaxTransactionIdLength = 256;

    /// <summary>
    /// Ceiling on the UTF-16 payload of all retained responses for one session. Chosen so the
    /// 128-response count cap stays the binding constraint for ordinary batches and this budget
    /// only bites on unusually large ones; see <c>tools/mcp-server/README.md</c> for the measured
    /// per-response cost this is sized against.
    /// </summary>
    public const long DefaultResponseByteBudget = 32L * 1024 * 1024;

    // Stable Unicode White_Space definition for transaction ids. Runtime validation, the schema
    // regex, and its prose are all derived from this one table so their blank-string semantics
    // cannot drift with .NET or ECMAScript whitespace classifications.
    private static readonly (int Start, int End)[] TransactionIdWhiteSpaceRanges =
    {
        (0x0009, 0x000D),
        (0x0020, 0x0020),
        (0x0085, 0x0085),
        (0x00A0, 0x00A0),
        (0x1680, 0x1680),
        (0x2000, 0x200A),
        (0x2028, 0x2029),
        (0x202F, 0x202F),
        (0x205F, 0x205F),
        (0x3000, 0x3000),
    };

    internal static string TransactionIdNonBlankPattern { get; } =
        BuildTransactionIdNonBlankPattern();

    internal static string TransactionIdWhiteSpaceDescription { get; } =
        string.Join(", ", TransactionIdWhiteSpaceRanges.Select(static range =>
            range.Start == range.End
                ? "U+" + ScalarHex(range.Start)
                : "U+" + ScalarHex(range.Start) + "-U+" + ScalarHex(range.End)));

    internal static string TransactionIdSchemaDescription { get; } =
        "Optional non-blank caller identity for an APPLYING batch only, limited to "
        + MaxTransactionIdLength.ToString(System.Globalization.CultureInfo.InvariantCulture)
        + " Unicode "
        + "scalar values. Blank means composed only of exactly these Unicode White_Space code "
        + "points: "
        + TransactionIdWhiteSpaceDescription
        + "; U+FEFF is non-whitespace. The first terminal response is retained in this open "
        + "session; an identical retry returns that exact serialized response without executing "
        + "or rechecking preconditions. Reusing the id for a different canonical request returns "
        + "transaction_conflict. Preview/dry-run rejects this field.";

    private readonly int _fullRecordCapacity;
    private readonly int _tombstoneCapacity;
    private readonly long _responseByteBudget;
    private readonly Func<DateTimeOffset> _utcNow;
    private readonly Func<string> _recordIdFactory;
    private readonly Dictionary<string, MutationTransactionRecord> _records =
        new(StringComparer.Ordinal);
    private readonly Queue<string> _completedFifo = new();
    private readonly Dictionary<string, MutationTransactionTombstone> _tombstones =
        new(StringComparer.Ordinal);
    private readonly Queue<string> _tombstoneFifo = new();

    // Reservations are tracked by reference, not id: an id can be reserved again after its
    // tombstone expires, and a stale queue entry must never prune that later live reservation.
    private readonly Queue<MutationTransactionRecord> _reservationFifo = new();
    private long _retainedResponseBytes;

    public MutationTransactions(
        int fullRecordCapacity = DefaultFullRecordCapacity,
        int tombstoneCapacity = DefaultTombstoneCapacity,
        Func<DateTimeOffset>? utcNow = null,
        Func<string>? recordIdFactory = null,
        long responseByteBudget = DefaultResponseByteBudget)
    {
        if (fullRecordCapacity < 1)
            throw new ArgumentOutOfRangeException(nameof(fullRecordCapacity));
        if (tombstoneCapacity < 0)
            throw new ArgumentOutOfRangeException(nameof(tombstoneCapacity));
        if (responseByteBudget < 1)
            throw new ArgumentOutOfRangeException(nameof(responseByteBudget));
        _fullRecordCapacity = fullRecordCapacity;
        _tombstoneCapacity = tombstoneCapacity;
        _responseByteBudget = responseByteBudget;
        _utcNow = utcNow ?? (() => DateTimeOffset.UtcNow);
        _recordIdFactory = recordIdFactory ?? NewRecordId;
    }

    internal int FullRecordCount
    {
        get { lock (_records) return _completedFifo.Count; }
    }

    internal int TombstoneCount
    {
        get { lock (_records) return _tombstones.Count; }
    }

    /// <summary>UTF-16 payload of every retained response — the quantity the byte budget caps.</summary>
    internal long RetainedResponseBytes
    {
        get { lock (_records) return _retainedResponseBytes; }
    }

    private static long ResponseByteCost(string? serializedResponse) =>
        serializedResponse is null ? 0L : (long)serializedResponse.Length * sizeof(char);

    internal MutationTransactionRecord? GetRecord(string transactionId)
    {
        lock (_records)
            return _records.TryGetValue(transactionId, out var record) ? record : null;
    }

    internal MutationTransactionTombstone? GetTombstone(string transactionId)
    {
        lock (_records)
            return _tombstones.TryGetValue(transactionId, out var tombstone) ? tombstone : null;
    }

    internal static bool IsBlankTransactionId(string transactionId)
    {
        foreach (var rune in transactionId.EnumerateRunes())
        {
            if (!TransactionIdWhiteSpaceRanges.Any(range =>
                    rune.Value >= range.Start && rune.Value <= range.End))
                return false;
        }
        return true;
    }

    private static string BuildTransactionIdNonBlankPattern()
    {
        var pattern = new StringBuilder("[^");
        foreach (var range in TransactionIdWhiteSpaceRanges)
        {
            pattern.Append("\\u").Append(ScalarHex(range.Start));
            if (range.Start != range.End)
                pattern.Append("-\\u").Append(ScalarHex(range.End));
        }
        return pattern.Append(']').ToString();
    }

    private static string ScalarHex(int value) =>
        value.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>Reserve a new identity, or resolve it to replay/conflict/expired/incomplete
    /// deterministically.</summary>
    public MutationTransactionDecision Begin(string transactionId, string requestFingerprint)
    {
        var requested = new MutationTransactionIdentity(
            SchemaVersion, transactionId, requestFingerprint);
        lock (_records)
        {
            if (_records.TryGetValue(transactionId, out var record))
            {
                if (!string.Equals(record.Identity.RequestFingerprint, requestFingerprint,
                        StringComparison.Ordinal))
                    return new MutationTransactionDecision(
                        MutationTransactionDecisionKind.Conflict,
                        ExistingIdentity: record.Identity);
                if (record.SerializedResponse is not null)
                    return new MutationTransactionDecision(
                        MutationTransactionDecisionKind.Replay,
                        record,
                        record.Identity,
                        record.SerializedResponse);

                // The id is bound to this exact request but never recorded a terminal response.
                // Per-session dispatch serialization plus Abandon mean this cannot occur through
                // Dispatcher; it is reachable when the component is driven directly. Reporting a
                // conflict here would be false — the fingerprints match — so it gets its own kind.
                return new MutationTransactionDecision(
                    MutationTransactionDecisionKind.Incomplete,
                    ExistingIdentity: record.Identity);
            }

            if (_tombstones.TryGetValue(transactionId, out var tombstone))
            {
                if (!string.Equals(tombstone.Identity.RequestFingerprint, requestFingerprint,
                        StringComparison.Ordinal))
                    return new MutationTransactionDecision(
                        MutationTransactionDecisionKind.Conflict,
                        ExistingIdentity: tombstone.Identity);
                return new MutationTransactionDecision(
                    tombstone.CompletedAt is null
                        ? MutationTransactionDecisionKind.Incomplete
                        : MutationTransactionDecisionKind.ResultEvicted,
                    ExistingIdentity: tombstone.Identity);
            }

            var reserved = new MutationTransactionRecord(
                _recordIdFactory(), requested, _utcNow(), null, null);
            _records.Add(transactionId, reserved);
            _reservationFifo.Enqueue(reserved);
            EvictStaleReservations();
            return new MutationTransactionDecision(
                MutationTransactionDecisionKind.Reserved, reserved, requested);
        }
    }

    /// <summary>
    /// Release a reservation that will never record a terminal response, keeping the identity
    /// bound as an outcome-unknown tombstone rather than stranding it in the live record map.
    /// Idempotent, and a no-op once the reservation has completed.
    /// </summary>
    public void Abandon(MutationTransactionRecord reservation)
    {
        ArgumentNullException.ThrowIfNull(reservation);
        lock (_records)
            RetireReservation(reservation);
    }

    private bool RetireReservation(MutationTransactionRecord reservation)
    {
        var id = reservation.Identity.TransactionId;
        if (!_records.TryGetValue(id, out var current)
            || !ReferenceEquals(current, reservation)
            || current.SerializedResponse is not null)
            return false;
        _records.Remove(id);
        Entomb(current);
        return true;
    }

    /// <summary>Atomically retain the exact response and apply FIFO eviction.</summary>
    public MutationTransactionRecord Complete(
        MutationTransactionRecord reservation,
        string serializedResponse)
    {
        ArgumentNullException.ThrowIfNull(reservation);
        ArgumentNullException.ThrowIfNull(serializedResponse);
        lock (_records)
        {
            if (!_records.TryGetValue(reservation.Identity.TransactionId, out var current)
                || !ReferenceEquals(current, reservation)
                || current.SerializedResponse is not null)
                throw new InvalidOperationException("mutation transaction reservation is no longer active");

            var completed = current with
            {
                // The document mutation has already committed when Complete is called. Clock
                // injection is diagnostic metadata and must never strand that committed result
                // behind an active reservation that cannot replay.
                CompletedAt = UtcNowOr(current.StartedAt),
                SerializedResponse = serializedResponse,
            };
            _records[completed.Identity.TransactionId] = completed;
            _completedFifo.Enqueue(completed.Identity.TransactionId);
            _retainedResponseBytes += ResponseByteCost(serializedResponse);
            EvictCompletedRecords();
            return completed;
        }
    }

    /// <summary>
    /// Bound retained responses by BOTH the count cap and the byte budget, oldest first. A single
    /// response larger than the whole budget evicts itself rather than raising the ceiling: an
    /// identical retry then answers <c>transaction_result_evicted</c>, which is a safe answer,
    /// whereas an unbounded retained response is not a bound at all.
    /// </summary>
    private void EvictCompletedRecords()
    {
        while (_completedFifo.Count > 0
            && (_completedFifo.Count > _fullRecordCapacity
                || _retainedResponseBytes > _responseByteBudget))
        {
            var id = _completedFifo.Dequeue();
            if (!_records.Remove(id, out var evicted)) continue;
            // Decrement on every path a retained response leaves _records, or the running total
            // drifts upward and the budget silently becomes permanent.
            _retainedResponseBytes -= ResponseByteCost(evicted.SerializedResponse);
            Entomb(evicted);
        }

        TrimTombstones();
    }

    /// <summary>
    /// Bound reservations that were never completed or abandoned. Only reachable through direct
    /// component use; the Dispatcher always completes or abandons. Identity stays bound as an
    /// outcome-unknown tombstone so a stale retry cannot silently become a fresh mutation.
    /// </summary>
    private void EvictStaleReservations()
    {
        while (_reservationFifo.Count > _fullRecordCapacity)
            RetireReservation(_reservationFifo.Dequeue());
        TrimTombstones();
    }

    private void Entomb(MutationTransactionRecord evicted)
    {
        if (_tombstoneCapacity == 0) return;
        var id = evicted.Identity.TransactionId;
        _tombstones[id] = new MutationTransactionTombstone(
            evicted.RecordId,
            evicted.Identity,
            evicted.StartedAt,
            evicted.CompletedAt,
            // Eviction must move the identity from a full record to a tombstone as one
            // logical operation even when an injected diagnostics clock fails.
            UtcNowOr(evicted.CompletedAt ?? evicted.StartedAt));
        _tombstoneFifo.Enqueue(id);
    }

    private void TrimTombstones()
    {
        while (_tombstoneFifo.Count > _tombstoneCapacity)
        {
            var id = _tombstoneFifo.Dequeue();
            _tombstones.Remove(id);
        }
    }

    private DateTimeOffset UtcNowOr(DateTimeOffset fallback)
    {
        try
        {
            return _utcNow();
        }
        catch (Exception)
        {
            return fallback;
        }
    }

    /// <summary>
    /// SHA-256 over a deterministic JSON rendering. Root session/transaction identity is excluded;
    /// objects are sorted, arrays and scalar spelling are retained, and numeric tokens are copied
    /// verbatim. Parsing already normalizes insignificant whitespace and equivalent string escapes.
    /// </summary>
    public static string Fingerprint(JsonElement request)
    {
        if (request.ValueKind != JsonValueKind.Object)
            throw new McpToolException("docxodus_mutations arguments must be an object");

        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        }))
        {
            WriteCanonical(writer, request, isRoot: true, "$arguments");
        }
        var hash = SHA256.HashData(buffer.WrittenSpan);
        return "sha256:" + Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static void WriteCanonical(
        Utf8JsonWriter writer,
        JsonElement value,
        bool isRoot,
        string path)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var properties = value.EnumerateObject().ToArray();
                var names = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in properties)
                {
                    if (!names.Add(property.Name))
                        throw new McpToolException(
                            $"duplicate JSON property {JsonSerializer.Serialize(property.Name)} at {path}");
                }
                Array.Sort(properties, static (left, right) =>
                    StringComparer.Ordinal.Compare(left.Name, right.Name));

                writer.WriteStartObject();
                var synthesizeAtomicMode = isRoot && !names.Contains("mode");
                foreach (var property in properties)
                {
                    if (isRoot && property.Name is "sessionId" or "transactionId") continue;
                    if (synthesizeAtomicMode
                        && StringComparer.Ordinal.Compare("mode", property.Name) < 0)
                    {
                        writer.WriteString("mode", "atomic");
                        synthesizeAtomicMode = false;
                    }
                    writer.WritePropertyName(property.Name);
                    WriteCanonical(writer, property.Value, isRoot: false,
                        path + "." + property.Name);
                }
                if (synthesizeAtomicMode)
                    writer.WriteString("mode", "atomic");
                writer.WriteEndObject();
                break;
            }
            case JsonValueKind.Array:
            {
                writer.WriteStartArray();
                var index = 0;
                foreach (var item in value.EnumerateArray())
                {
                    WriteCanonical(writer, item, isRoot: false, $"{path}[{index}]");
                    index++;
                }
                writer.WriteEndArray();
                break;
            }
            case JsonValueKind.String:
                writer.WriteStringValue(value.GetString());
                break;
            case JsonValueKind.Number:
                writer.WriteRawValue(value.GetRawText(), skipInputValidation: true);
                break;
            case JsonValueKind.True:
                writer.WriteBooleanValue(true);
                break;
            case JsonValueKind.False:
                writer.WriteBooleanValue(false);
                break;
            case JsonValueKind.Null:
                writer.WriteNullValue();
                break;
            default:
                throw new McpToolException($"unsupported JSON value at {path}");
        }
    }

    /// <summary>Add the versioned identity to the already-serialized core batch result.</summary>
    public static string AttachIdentity(
        string serializedBatchResult,
        MutationTransactionIdentity identity)
    {
        var end = serializedBatchResult.Length - 1;
        while (end >= 0 && char.IsWhiteSpace(serializedBatchResult[end])) end--;
        if (end < 1 || serializedBatchResult[0] != '{' || serializedBatchResult[end] != '}')
            throw new InvalidOperationException("mutation batch result was not a JSON object");

        var suffix = serializedBatchResult[(end + 1)..];
        return serializedBatchResult[..end]
            + ",\"transaction\":{\"schemaVersion\":"
            + identity.SchemaVersion.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + ",\"transactionId\":"
            + JsonRpcIo.JsonString(identity.TransactionId)
            + ",\"requestFingerprint\":"
            + JsonRpcIo.JsonString(identity.RequestFingerprint)
            + "}}"
            + suffix;
    }

    /// <summary>Serialize a transport-level terminal outcome through MutationBatchResult.</summary>
    public static string SerializeFailure(
        MutationBatchMode mode,
        bool preview,
        long version,
        EditErrorCode code,
        string message,
        string action,
        bool rolledBack = false)
        => SerializeFailure(
            mode, preview, version, new EditError(code, message), action, rolledBack);

    public static string SerializeFailure(
        MutationBatchMode mode,
        bool preview,
        long version,
        EditError error,
        string action,
        bool rolledBack = false)
    {
        var edit = new EditResult { Success = false, Error = error };
        var step = new MutationBatchStepResult(
            0, "docxodus_mutations", action, new[] { edit }, rolledBack);
        return DocxSessionJson.SerializeMutationBatchResult(new MutationBatchResult
        {
            Mode = mode,
            Preview = preview,
            Success = false,
            RolledBack = rolledBack,
            BaseVersion = version,
            ResultVersion = version,
            Steps = new[] { step },
            Failure = new MutationBatchFailure(
                step.Index, step.Tool, step.Action, error, rolledBack),
        });
    }

    private static string NewRecordId() =>
        "mtx_" + Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();
}
