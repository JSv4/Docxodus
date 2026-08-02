#nullable enable

using System;
using System.IO;

namespace Docxodus.McpServer;

/// <summary>
/// Single owner of "which <see cref="IDocumentStore"/> is this process using, and rooted where" —
/// the same single-owner-facade shape the core library uses for its transports
/// (<c>DocxSessionOps</c>, <c>DocxDiffOps</c>). Configuration is read once at startup from the
/// environment; adding a backend means adding an <see cref="IDocumentStore"/> implementation and
/// one case in <see cref="FromEnvironment"/>, with no change to the dispatcher or the tool schemas.
///
/// <para><b>The backend and its root are operator configuration, never tool arguments.</b> If the
/// agent could name a backend or a root per call, the scope would be chosen by the thing the scope
/// exists to contain. So the only storage input a tool call carries is a location <em>within</em>
/// the configured store.</para>
///
/// <list type="table">
///   <listheader><term>Variable</term><description>Meaning</description></listheader>
///   <item>
///     <term><c>DOCXODUS_STORAGE_BACKEND</c></term>
///     <description>Backend id. Only <c>file</c> is implemented; defaults to <c>file</c>.</description>
///   </item>
///   <item>
///     <term><c>DOCXODUS_STORAGE_ROOT</c></term>
///     <description>Directory every location must resolve inside. Defaults to the user's home
///     directory, which keeps ordinary local use working (an agent can open
///     <c>~/Downloads/contract.docx</c>) while still excluding the rest of the machine. Set it
///     narrower to confine the server; set it to the filesystem root to opt out of confinement
///     entirely.</description>
///   </item>
///   <item>
///     <term><c>DOCXODUS_STORAGE_SCOPE</c></term>
///     <description>Optional scope segment appended to the root, so the effective root becomes
///     <c>{root}/{scope}</c> and concurrent server processes carrying different scopes cannot see
///     each other's documents. Supplied by whoever launches the server — not by the agent, which
///     never learns it. Because the scope is just a stable path segment, passing the same value
///     next session reaches the same documents; persistence is the filesystem's.</description>
///   </item>
/// </list>
/// </summary>
internal static class DocumentStores
{
    public const string BackendVariable = "DOCXODUS_STORAGE_BACKEND";
    public const string RootVariable = "DOCXODUS_STORAGE_ROOT";
    public const string ScopeVariable = "DOCXODUS_STORAGE_SCOPE";

    public static IDocumentStore FromEnvironment() => Create(
        Environment.GetEnvironmentVariable(BackendVariable),
        Environment.GetEnvironmentVariable(RootVariable),
        Environment.GetEnvironmentVariable(ScopeVariable));

    /// <summary>Testable core of <see cref="FromEnvironment"/>: the same resolution against
    /// explicitly supplied values rather than the ambient environment.</summary>
    public static IDocumentStore Create(string? backend, string? root, string? scope)
    {
        var kind = string.IsNullOrWhiteSpace(backend) ? "file" : backend.Trim().ToLowerInvariant();
        if (kind != "file")
            throw new McpToolException(
                $"unsupported {BackendVariable} '{backend}'. This build supports: file");

        var basePath = string.IsNullOrWhiteSpace(root) ? DefaultRoot() : root.Trim();
        return new LocalFileDocumentStore(ApplyScope(basePath, scope));
    }

    /// <summary>
    /// Append the scope segment, rejecting anything that could climb out of the base — the scope
    /// comes from the launching process rather than the agent, but a misconfigured value should
    /// fail loudly rather than silently widen the root.
    /// </summary>
    internal static string ApplyScope(string basePath, string? scope)
    {
        if (string.IsNullOrWhiteSpace(scope)) return basePath;

        var trimmed = scope.Trim();
        if (Path.IsPathRooted(trimmed)
            || trimmed.Contains("..", StringComparison.Ordinal)
            || trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new McpToolException(
                $"invalid {ScopeVariable} '{scope}': must be a single relative path segment " +
                "(no separators, no '..', no absolute path)");
        }

        return Path.Combine(basePath, trimmed);
    }

    private static string DefaultRoot()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return string.IsNullOrWhiteSpace(home) ? Directory.GetCurrentDirectory() : home;
    }
}
