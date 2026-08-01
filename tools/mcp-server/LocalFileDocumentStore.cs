#nullable enable

using System;
using System.IO;

namespace Docxodus.McpServer;

/// <summary>
/// Local-filesystem <see cref="IDocumentStore"/>, rooted at a directory that every location must
/// resolve inside. A relative location resolves under the root; an absolute one is accepted only
/// if it canonicalizes to something the root contains. That combination is what lets the default
/// configuration stay useful for ordinary local work (root <c>$HOME</c> → an agent can open
/// <c>~/Downloads/contract.docx</c> by its natural absolute path) while a narrower root
/// (<c>/var/docxodus/tenant-42</c>) confines the same unmodified tool surface completely.
///
/// <para><b>Containment is enforced against symlink-resolved paths.</b> Lexical normalization
/// alone (<see cref="Path.GetFullPath"/>) collapses <c>..</c> but does not follow links, so a
/// symlink inside the root pointing outside it would otherwise pass — the classic escape. Both
/// the root and each candidate are resolved through their link chains before comparison. For a
/// location that doesn't exist yet (the save case), the deepest existing ancestor is resolved and
/// the not-yet-existing segments are appended, so a link partway up the path is still followed.</para>
/// </summary>
internal sealed class LocalFileDocumentStore : IDocumentStore
{
    private readonly string _root;

    /// <param name="root">Directory every location must resolve inside. Created if absent.</param>
    public LocalFileDocumentStore(string root)
    {
        if (string.IsNullOrWhiteSpace(root))
            throw new ArgumentException("root must be a non-empty path", nameof(root));

        var full = Path.GetFullPath(root);
        try
        {
            Directory.CreateDirectory(full);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new McpToolException($"storage root '{full}' could not be created: {ex.Message}");
        }

        // Resolve the root's own link chain once, so comparisons are like-for-like (on macOS,
        // for instance, /tmp is a symlink to /private/tmp — an unresolved root would reject
        // every path under it).
        _root = Canonicalize(full);
    }

    public string Kind => "file";

    public string RootDescription => _root;

    public string Resolve(string location)
    {
        if (string.IsNullOrWhiteSpace(location))
            throw new McpToolException("location must be a non-empty path");

        var combined = Path.IsPathRooted(location) ? location : Path.Combine(_root, location);
        var candidate = Canonicalize(combined);

        if (!IsInsideRoot(candidate))
            throw new McpToolException(
                $"'{location}' resolves outside this server's document scope ({_root}). " +
                "Use a path inside that directory, or start the server with a wider " +
                "DOCXODUS_STORAGE_ROOT if this location should be reachable.");

        return candidate;
    }

    public byte[] Read(string resolvedLocation)
    {
        try
        {
            return File.ReadAllBytes(resolvedLocation);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            throw new McpToolException($"could not read '{resolvedLocation}': {ex.Message}");
        }
    }

    public void Write(string resolvedLocation, byte[] bytes)
    {
        try
        {
            // Safe to create: resolvedLocation is already known to be inside the root, so this
            // can only ever create directories within scope.
            var dir = Path.GetDirectoryName(resolvedLocation);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            File.WriteAllBytes(resolvedLocation, bytes);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            throw new McpToolException($"could not write '{resolvedLocation}': {ex.Message}");
        }
    }

    /// <summary>Bound on link-following recursion; also breaks symlink cycles.</summary>
    private const int MaxLinkDepth = 64;

    /// <summary>
    /// Full path with <c>..</c>/<c>.</c> collapsed AND symlinks followed — <b>including links in
    /// intermediate components</b>, which is the case that matters: an attacker plants
    /// <c>{root}/link → /elsewhere</c> and names <c>{root}/link/secret.docx</c>, whose own leaf is
    /// not a link at all. Resolving only the leaf (or only the deepest existing ancestor) misses
    /// that entirely, so each component is resolved from the volume root down.
    /// </summary>
    private static string Canonicalize(string path)
    {
        try
        {
            return CanonicalizeCore(Path.GetFullPath(path), 0);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Unreadable somewhere along the path: fall back to the lexical form, which is still
            // containment-checked — an unresolvable path cannot be used to escape, only rejected.
            return Path.GetFullPath(path);
        }
    }

    private static string CanonicalizeCore(string path, int depth)
    {
        if (depth > MaxLinkDepth) return path;

        var parent = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(parent) || parent == path)
            return path; // volume root — nothing above it to resolve

        var here = Path.Combine(CanonicalizeCore(parent, depth + 1), Path.GetFileName(path));

        var target = ReadLinkTarget(here);
        if (target is null) return here; // not a link (or doesn't exist yet)

        // A link target may be relative to the directory holding the link.
        var absolute = Path.IsPathRooted(target)
            ? target
            : Path.Combine(Path.GetDirectoryName(here) ?? string.Empty, target);
        return CanonicalizeCore(Path.GetFullPath(absolute), depth + 1);
    }

    /// <summary>
    /// The link target of <paramref name="path"/>, or null if it is not a link. Reads the link
    /// itself rather than following it, so a <em>dangling</em> link is still detected — otherwise
    /// a link to a not-yet-existing directory would look like an ordinary missing path and a write
    /// through it would land outside the root.
    /// </summary>
    private static string? ReadLinkTarget(string path)
    {
        try
        {
            if (new DirectoryInfo(path).LinkTarget is { } dirTarget) return dirTarget;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            // fall through to the file probe
        }

        try
        {
            return new FileInfo(path).LinkTarget;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return null;
        }
    }

    /// <summary>
    /// True when <paramref name="candidate"/> is the root or sits beneath it. Uses
    /// <see cref="Path.GetRelativePath"/> rather than a string prefix test so that segment
    /// boundaries are respected (<c>/srv/base-2</c> is not inside <c>/srv/base</c>) and so the
    /// comparison follows the platform's own case rules.
    /// </summary>
    private bool IsInsideRoot(string candidate)
    {
        var relative = Path.GetRelativePath(_root, candidate);

        if (relative == ".") return true;                    // the root itself
        if (Path.IsPathRooted(relative)) return false;       // no shared ancestor (other volume)
        if (relative == "..") return false;
        if (relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)) return false;
        if (Path.AltDirectorySeparatorChar != Path.DirectorySeparatorChar
            && relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal)) return false;

        return true;
    }
}
