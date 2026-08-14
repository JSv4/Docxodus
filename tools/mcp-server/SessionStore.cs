#nullable enable

using System.Collections.Concurrent;
using System.Security.Cryptography;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>
/// One open document. Wraps the raw <see cref="DocxSessionOps"/> integer handle with the
/// store location the stdio server needs so <c>docxodus_save</c> can default to "write back".
/// </summary>
internal sealed class DocSession
{
    required public string Id { get; init; }
    public int Handle { get; init; }

    /// <summary>
    /// Serializes every action addressed to this session. The core session has its own mutation
    /// lock, but MCP dispatch also has to order reads, saves, closes, and transaction-id replay
    /// decisions with mutations so a retry cannot race a different session action.
    /// </summary>
    internal object DispatchGate { get; } = new();

    /// <summary>In-session mutation transaction identities and their exact serialized results.</summary>
    internal MutationTransactions MutationTransactions { get; init; } = new();

    /// <summary>False after close has won the dispatch race for this session.</summary>
    internal volatile bool Active = true;

    /// <summary>Store-resolved location this session was opened from — already checked to be in
    /// scope, so a save back to it needs no re-validation. Null only if a session was opened from
    /// bytes with no origin.</summary>
    public string? Location { get; set; }
}

/// <summary>
/// External session-id → <see cref="DocSession"/> registry for the MCP tool surface.
/// Deliberately separate from <see cref="Docxodus.Internal.SessionRegistry"/> (which only
/// knows integer handles) so callers use unguessable capability ids and retain document-store
/// location metadata without exposing process-local handles.
/// </summary>
internal sealed class SessionStore
{
    private sealed class DispatchFrame
    {
        public volatile bool Active = true;
    }

    // Static scope rejects cross-store reentry too. ExecutionContext copies the frame reference,
    // while Active is shared so deferred children cease being reentrant when the parent returns.
    private static readonly AsyncLocal<DispatchFrame?> CurrentDispatch = new();

    private readonly ConcurrentDictionary<string, DocSession> _sessions = new();
    private readonly object _lifecycleGate = new();
    private readonly System.Func<MutationTransactions> _mutationTransactionsFactory;

    /// <param name="documents">Backing document store. Defaults to a local store rooted at the
    /// process's current directory, which is only appropriate for tests — the server proper
    /// passes the environment-configured store from <see cref="DocumentStores.FromEnvironment"/>.</param>
    public SessionStore(
        IDocumentStore? documents = null,
        System.Func<MutationTransactions>? mutationTransactionsFactory = null)
    {
        Documents = documents
            ?? new LocalFileDocumentStore(System.IO.Directory.GetCurrentDirectory());
        _mutationTransactionsFactory = mutationTransactionsFactory
            ?? (() => new MutationTransactions());
    }

    /// <summary>Where this server's documents are read from and written to. Every session in the
    /// process shares it, and it is already rooted at the configured scope.</summary>
    public IDocumentStore Documents { get; }

    public DocSession Open(byte[] bytes, string? location, DocxSessionSettings settings)
    {
        RejectReentrantLifecycle("open");
        // Construct fallible per-session collaborators before allocating a core handle. If the
        // factory fails, there is nothing in either registry to clean up.
        var mutationTransactions = _mutationTransactionsFactory();
        lock (_lifecycleGate)
        {
            var handle = DocxSessionOps.OpenSession(bytes, settings);
            try
            {
                var session = new DocSession
                {
                    Id = NewSessionId(),
                    Handle = handle,
                    Location = location,
                    MutationTransactions = mutationTransactions,
                };
                _sessions[session.Id] = session;
                return session;
            }
            catch
            {
                DocxSessionOps.CloseSession(handle);
                throw;
            }
        }
    }

    /// <summary>
    /// Unguessable session id. These are capabilities, not just names: holding one is what lets a
    /// caller act on that document, so a sequential counter would let anything able to make tool
    /// calls address a session it was never given (including one belonging to another concurrent
    /// caller, in any future non-stdio transport where that is possible).
    /// </summary>
    private static string NewSessionId() =>
        "s_" + System.Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();

    public DocSession Get(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId, out var session) || !session.Active)
            throw new McpToolException($"unknown session_id: {sessionId}");
        return session;
    }

    /// <summary>
    /// Run one complete session-bound dispatch while holding the session's synchronous gate.
    /// The active check after taking the gate closes the lookup/close race: a caller that found
    /// the session before close removed it still cannot enter the disposed core handle.
    /// Callbacks are non-reentrant so lifecycle operations retain one lifecycle-to-session lock order.
    /// Callback-spawned work must preserve <see cref="System.Threading.ExecutionContext"/> flow;
    /// deliberately suppressed or unsafe flow is unsupported by this synchronous dispatch contract.
    /// </summary>
    public string Dispatch(string sessionId, System.Func<string> action)
    {
        if (CurrentDispatch.Value?.Active == true)
            throw new McpToolException(
                "cannot nest a session dispatch within a session dispatch callback");
        if (!_sessions.TryGetValue(sessionId, out var session))
            throw new McpToolException($"unknown session_id: {sessionId}");
        lock (session.DispatchGate)
        {
            if (!session.Active
                || !_sessions.TryGetValue(sessionId, out var current)
                || !ReferenceEquals(session, current))
                throw new McpToolException($"unknown session_id: {sessionId}");
            var priorFrame = CurrentDispatch.Value;
            var frame = new DispatchFrame();
            CurrentDispatch.Value = frame;
            try
            {
                return action();
            }
            finally
            {
                frame.Active = false;
                CurrentDispatch.Value = priorFrame;
            }
        }
    }

    public void Close(string sessionId)
    {
        RejectReentrantLifecycle("close");
        lock (_lifecycleGate)
        {
            if (!_sessions.TryGetValue(sessionId, out var session)) return;
            lock (session.DispatchGate)
            {
                if (!session.Active) return;
                session.Active = false;
                _sessions.TryRemove(sessionId, out _);
                DocxSessionOps.CloseSession(session.Handle);
            }
        }
    }

    public void CloseAll()
    {
        RejectReentrantLifecycle("close all sessions");
        lock (_lifecycleGate)
        {
            foreach (var kv in _sessions)
            {
                lock (kv.Value.DispatchGate)
                {
                    if (!kv.Value.Active) continue;
                    kv.Value.Active = false;
                    _sessions.TryRemove(kv.Key, out _);
                    DocxSessionOps.CloseSession(kv.Value.Handle);
                }
            }
        }
    }

    private void RejectReentrantLifecycle(string operation)
    {
        if (CurrentDispatch.Value?.Active == true)
            throw new McpToolException(
                $"cannot {operation} from within a session dispatch callback");
    }
}

/// <summary>Business-level tool failure — reported as an MCP tool result with <c>isError: true</c>,
/// never as a JSON-RPC protocol error (those are reserved for transport-level problems).</summary>
internal sealed class McpToolException : System.Exception
{
    public McpToolException(string message) : base(message) { }
}
