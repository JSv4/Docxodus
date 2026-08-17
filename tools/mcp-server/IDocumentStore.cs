#nullable enable

namespace Docxodus.McpServer;

/// <summary>
/// Where this server reads and writes documents. One store instance is created at startup from
/// configuration (see <see cref="DocumentStores"/>) and shared by every session in the process.
///
/// <para>The interface is deliberately three methods over opaque <em>location</em> strings rather
/// than a filesystem-shaped API, so a backend that isn't a filesystem — object storage, a content
/// repository — can implement it without pretending to have directories. A backend supplies:
/// <see cref="Resolve"/> (turn a caller-supplied location into the canonical, in-scope identifier
/// this store will accept, or reject it), plus byte-level <see cref="Read"/>/<see cref="Write"/>
/// of an already-resolved location.</para>
///
/// <para><b>Isolation is the store's job, not the caller's.</b> Every location a tool call names
/// passes through <see cref="Resolve"/> first, and a store is constructed already rooted at its
/// scope — so a session physically cannot name something outside it. This is why the check lives
/// here rather than as an <c>if</c> in the dispatcher: there is no code path that reads or writes
/// without having resolved first, so the guarantee doesn't depend on remembering to check.</para>
/// </summary>
internal interface IDocumentStore
{
    /// <summary>Backend identifier for diagnostics and error messages, e.g. <c>"file"</c>.</summary>
    string Kind { get; }

    /// <summary>Human-readable description of this store's scope root, for diagnostics. Appears in
    /// the "outside this store's scope" rejection message so an operator can see what the root
    /// actually resolved to.</summary>
    string RootDescription { get; }

    /// <summary>
    /// Canonicalize <paramref name="location"/> and verify it falls inside this store's scope,
    /// returning the identifier <see cref="Read"/>/<see cref="Write"/> accept. The returned value
    /// is what a session records so a later save with no explicit location writes back to the same
    /// place.
    /// </summary>
    /// <exception cref="McpToolException">The location is malformed, or resolves outside this
    /// store's scope.</exception>
    string Resolve(string location);

    /// <summary>Read the bytes at an already-<see cref="Resolve"/>d location.</summary>
    /// <exception cref="McpToolException">The location does not exist or cannot be read.</exception>
    byte[] Read(string resolvedLocation);

    /// <summary>
    /// Read with a caller-owned byte ceiling. Backends should reject from metadata/stream length
    /// before materializing content; this default preserves compatibility for non-file backends.
    /// </summary>
    byte[] Read(string resolvedLocation, long maximumBytes)
    {
        if (maximumBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        var bytes = Read(resolvedLocation);
        if (bytes.LongLength > maximumBytes)
            throw new McpToolException(
                $"document exceeds the {maximumBytes}-byte read limit");
        return bytes;
    }

    /// <summary>Write bytes to an already-<see cref="Resolve"/>d location, creating or replacing it.</summary>
    /// <exception cref="McpToolException">The location cannot be written.</exception>
    void Write(string resolvedLocation, byte[] bytes);
}
