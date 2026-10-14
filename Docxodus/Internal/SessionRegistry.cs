#nullable enable

using System;
using System.Collections.Concurrent;
using System.Threading;

namespace Docxodus.Internal;

/// <summary>
/// Process-wide pool of <see cref="DocxSession"/> instances keyed by an integer handle.
/// Shared between the WASM JSExport bridge (<c>DocxSessionBridge</c>) and the stdio
/// NDJSON host (<c>docxodus-pyhost</c>) so both transports speak the same handle protocol.
/// Sessions live until <see cref="CloseSession"/> is called or the host process exits.
/// </summary>
internal static class SessionRegistry
{
    private static readonly ConcurrentDictionary<int, DocxSession> _sessions = new();
    private static readonly ConcurrentDictionary<int, MutationTransactions> _transactions = new();
    private static int _nextId;

    public static int OpenSession(byte[] bytes, DocxSessionSettings? settings)
    {
        var session = new DocxSession(bytes, settings);
        return Register(session);
    }

    /// <summary>
    /// Register a complete isolated clone of an existing session. The returned handle owns only
    /// the shadow; callers must close it in a finally block. No live history/cache/config object is
    /// shared with the clone.
    /// </summary>
    public static int CloneSessionForPreview(int handle)
    {
        var shadow = Get(handle).CreateShadowSession();
        try
        {
            return Register(shadow);
        }
        catch
        {
            shadow.Dispose();
            throw;
        }
    }

    private static int Register(DocxSession session)
    {
        var id = Interlocked.Increment(ref _nextId);
        _sessions[id] = session;
        return id;
    }

    public static void CloseSession(int handle)
    {
        // Closing clears the journal: a reopened document starts a new transaction-identity
        // namespace even when it opens the same saved bytes.
        _transactions.TryRemove(handle, out _);
        if (_sessions.TryRemove(handle, out var s)) s.Dispose();
    }

    /// <summary>
    /// The mutation-transaction journal of one live session — one per handle regardless of
    /// which transport drives it, created on first use, discarded with the session.
    /// </summary>
    public static MutationTransactions Transactions(int handle)
    {
        _ = Get(handle);
        return _transactions.GetOrAdd(handle, _ => new MutationTransactions());
    }

    /// <summary>Install a caller-configured journal (retention limits, clocks) for a live session,
    /// replacing the default one. Hosts that own their own session wrapper call this on open.</summary>
    public static void AttachTransactions(int handle, MutationTransactions journal)
    {
        ArgumentNullException.ThrowIfNull(journal);
        _ = Get(handle);
        _transactions[handle] = journal;
    }

    public static DocxSession Get(int handle)
    {
        if (!_sessions.TryGetValue(handle, out var s))
            throw new ArgumentException($"unknown session handle: {handle}");
        return s;
    }

    public static int Count => _sessions.Count;

    public static void DisposeAll()
    {
        _transactions.Clear();
        foreach (var kv in _sessions)
        {
            if (_sessions.TryRemove(kv.Key, out var s)) s.Dispose();
        }
    }
}
