#nullable enable

using System.Collections.Concurrent;
using System.Security.Cryptography;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>
/// One open document. Wraps the raw <see cref="DocxSessionOps"/> integer handle with the
/// bookkeeping a stdio tool server needs that the handle alone doesn't carry: the store location
/// it was opened from (so <c>docxodus_save</c> can default to "write back"), and the settings
/// it was opened with (so a whole-document transform — see <see cref="Rebind"/> — can reopen
/// an equivalent session instead of losing tracked-change/undo configuration).
/// </summary>
internal sealed class DocSession
{
    required public string Id { get; init; }
    public int Handle { get; set; }

    /// <summary>Store-resolved location this session was opened from — already checked to be in
    /// scope, so a save back to it needs no re-validation. Null only if a session was opened from
    /// bytes with no origin.</summary>
    public string? Location { get; set; }

    public DocxSessionSettings Settings { get; set; } = new();
}

/// <summary>
/// External session-id → <see cref="DocSession"/> registry for the MCP tool surface.
/// Deliberately separate from <see cref="Docxodus.Internal.SessionRegistry"/> (which only
/// knows integer handles): tools like <c>docxodus_track_changes</c>'s accept-all/reject-all
/// need to swap the underlying handle for a fresh one built from transformed bytes while the
/// caller keeps addressing the same session id — see <see cref="Rebind"/>.
/// </summary>
internal sealed class SessionStore
{
    private readonly ConcurrentDictionary<string, DocSession> _sessions = new();

    /// <param name="documents">Backing document store. Defaults to a local store rooted at the
    /// process's current directory, which is only appropriate for tests — the server proper
    /// passes the environment-configured store from <see cref="DocumentStores.FromEnvironment"/>.</param>
    public SessionStore(IDocumentStore? documents = null) =>
        Documents = documents ?? new LocalFileDocumentStore(System.IO.Directory.GetCurrentDirectory());

    /// <summary>Where this server's documents are read from and written to. Every session in the
    /// process shares it, and it is already rooted at the configured scope.</summary>
    public IDocumentStore Documents { get; }

    public DocSession Open(byte[] bytes, string? location, DocxSessionSettings settings)
    {
        var handle = DocxSessionOps.OpenSession(bytes, settings);
        var session = new DocSession
        {
            Id = NewSessionId(),
            Handle = handle,
            Location = location,
            Settings = settings,
        };
        _sessions[session.Id] = session;
        return session;
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
        if (!_sessions.TryGetValue(sessionId, out var session))
            throw new McpToolException($"unknown session_id: {sessionId}");
        return session;
    }

    public void Close(string sessionId)
    {
        if (_sessions.TryRemove(sessionId, out var session))
            DocxSessionOps.CloseSession(session.Handle);
    }

    /// <summary>
    /// Replace a session's underlying document with <paramref name="newBytes"/> in place: opens
    /// a fresh handle (with the session's original settings, unchanged) and closes the old one,
    /// without changing the external session id the caller addresses it by. Used by
    /// whole-document byte transforms — currently only "accept every tracked change" / "reject
    /// every tracked change" — that have no in-place session mutation of their own.
    /// <paramref name="newBytes"/> must have been produced from a Unid-preserving save (see
    /// <see cref="DocxSessionOps.SaveWithAnchorIds"/>) so the reopened session resolves the same
    /// anchor ids the caller already has cached; the session's OWN
    /// <see cref="DocxSessionSettings.PersistAnchorIds"/> is passed through unchanged (not
    /// forced true) so a later <c>docxodus_save</c> to disk still strips that bookkeeping the
    /// way a save-to-disk should.
    /// </summary>
    public void Rebind(DocSession session, byte[] newBytes)
    {
        var newHandle = DocxSessionOps.OpenSession(newBytes, session.Settings);
        var oldHandle = session.Handle;
        session.Handle = newHandle;
        DocxSessionOps.CloseSession(oldHandle);
    }

    public void CloseAll()
    {
        foreach (var kv in _sessions)
            DocxSessionOps.CloseSession(kv.Value.Handle);
        _sessions.Clear();
    }
}

/// <summary>Business-level tool failure — reported as an MCP tool result with <c>isError: true</c>,
/// never as a JSON-RPC protocol error (those are reserved for transport-level problems).</summary>
internal sealed class McpToolException : System.Exception
{
    public McpToolException(string message) : base(message) { }
}
