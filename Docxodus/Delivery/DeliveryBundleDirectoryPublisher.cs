// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.ObjectModel;

namespace Docxodus.Delivery;

/// <summary>
/// Publishes a verified delivery bundle as a new directory. Publication never replaces an
/// existing path: all bytes are written to an owned sibling directory, the manifest is written
/// last, and a same-filesystem directory rename is the only commit point.
/// </summary>
public static class DeliveryBundleDirectoryPublisher
{
    private const string StageMarkerName = ".docxodus-delivery-stage";
    private const int StageCreationAttempts = 32;

    /// <summary>
    /// Publish all available artifacts and the canonical manifest to a new directory. The
    /// manifest is independently parsed and verified from staged filesystem bytes before commit.
    /// </summary>
    public static string Publish(
        DeliveryBundle bundle,
        string targetDirectory,
        DeliveryBundleVerificationLimits? verificationLimits = null)
    {
        return Publish(bundle, targetDirectory, verificationLimits, faultInjector: null);
    }

    internal static string Publish(
        DeliveryBundle bundle,
        string targetDirectory,
        DeliveryBundleVerificationLimits? verificationLimits,
        IDeliveryBundleDirectoryPublisherFaultInjector? faultInjector)
    {
        ArgumentNullException.ThrowIfNull(bundle);
        var declaredPaths = bundle.Manifest.Payload.Artifacts
            .Select(artifact => ValidateRelativePath(artifact.RelativePath, nameof(bundle)))
            .Append(DeliveryBundle.ManifestFileName)
            .ToArray();
        ValidatePathSet(declaredPaths, nameof(bundle));
        if (declaredPaths.Any(path => path.Equals(StageMarkerName,
                StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException(
                $"Delivery bundle declares the reserved publication path '{StageMarkerName}'.");

        var availableArtifacts = bundle.Manifest.Payload.Artifacts
            .Where(artifact => artifact.Availability == DeliveryArtifactAvailability.Available)
            .ToArray();
        var bundleBytes = bundle.ArtifactBytes;
        var expectedIds = availableArtifacts.Select(artifact => artifact.ArtifactId)
            .ToHashSet(StringComparer.Ordinal);
        if (!expectedIds.SetEquals(bundleBytes.Keys))
            throw new InvalidDataException(
                "Delivery bundle bytes do not match its available artifact declarations.");

        var source = new DeliveryBundleDirectoryPublicationSource
        {
            Artifacts = availableArtifacts.Select(artifact =>
                new DeliveryBundleDirectoryPublicationArtifact(
                    artifact.RelativePath,
                    bundleBytes[artifact.ArtifactId])).ToArray(),
            ManifestRelativePath = DeliveryBundle.ManifestFileName,
            CanonicalManifestBytes = bundle.ManifestBytes,
            VerifyStagedBytes = stagedFiles =>
            {
                if (!stagedFiles.TryGetValue(DeliveryBundle.ManifestFileName,
                        out var stagedManifest))
                    throw new InvalidDataException("The staged delivery manifest is missing.");
                var stagedArtifacts = availableArtifacts.ToDictionary(
                    artifact => artifact.ArtifactId,
                    artifact => stagedFiles.TryGetValue(artifact.RelativePath, out var bytes)
                        ? bytes
                        : throw new InvalidDataException(
                            $"Staged delivery artifact is missing: '{artifact.RelativePath}'."),
                    StringComparer.Ordinal);
                var verification = DeliveryBundleVerifier.VerifyJson(
                    stagedManifest,
                    stagedArtifacts,
                    verificationLimits);
                if (!verification.IsValid)
                    throw new InvalidDataException(
                        $"Staged delivery bundle verification failed: {verification.Findings[0]}");
            },
        };
        return Publish(source, targetDirectory, faultInjector);
    }

    /// <summary>
    /// Internal model boundary used while the delivery bundle builder owns the public model.
    /// A public <c>DeliveryBundle</c> overload can project into this source without duplicating any
    /// filesystem or verification behavior.
    /// </summary>
    internal static string Publish(
        DeliveryBundleDirectoryPublicationSource bundle,
        string targetDirectory,
        IDeliveryBundleDirectoryPublisherFaultInjector? faultInjector = null)
    {
        ArgumentNullException.ThrowIfNull(bundle);

        var target = ValidateTarget(targetDirectory);
        var parent = Path.GetDirectoryName(target)!;
        var snapshot = SnapshotAndValidate(bundle);
        var injector = faultInjector ?? NoDeliveryBundleDirectoryPublisherFaultInjector.Instance;

        string? stage = null;
        var stageOwned = false;
        try
        {
            (stage, stageOwned) = CreateOwnedStage(parent, Path.GetFileName(target));

            for (var index = 0; index < snapshot.Artifacts.Count; index++)
            {
                var artifact = snapshot.Artifacts[index];
                injector.OnCheckpoint(new DeliveryBundleDirectoryPublisherFaultContext(
                    DeliveryBundleDirectoryPublisherCheckpoint.BeforeArtifactWrite,
                    target,
                    stage,
                    artifact.RelativePath,
                    index));
                WriteNewFile(stage, artifact.RelativePath, artifact.Bytes);
            }

            injector.OnCheckpoint(new DeliveryBundleDirectoryPublisherFaultContext(
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeManifestWrite,
                target,
                stage,
                snapshot.ManifestRelativePath,
                snapshot.Artifacts.Count));
            WriteNewFile(stage, snapshot.ManifestRelativePath, snapshot.ManifestBytes);

            injector.OnCheckpoint(new DeliveryBundleDirectoryPublisherFaultContext(
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification,
                target,
                stage,
                null,
                null));

            var stagedBytes = ReadExpectedStage(stage, snapshot);
            bundle.VerifyStagedBytes(stagedBytes);

            injector.OnCheckpoint(new DeliveryBundleDirectoryPublisherFaultContext(
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit,
                target,
                stage,
                null,
                null));

            // Verification is over fresh reads from the stage. Re-read immediately before the
            // commit so a fault hook or concurrent process cannot change verified bytes in place.
            var preCommitBytes = ReadExpectedStage(stage, snapshot);
            EnsureByteEquality(stagedBytes, preCommitBytes);
            EnsureTargetStillAvailable(target, parent);

            File.Delete(Path.Combine(stage, StageMarkerName));
            Directory.Move(stage, target);
            stageOwned = false;
            return target;
        }
        catch
        {
            if (stageOwned && stage is not null)
                TryDeleteOwnedStage(stage);
            throw;
        }
    }

    private static string ValidateTarget(string targetDirectory)
    {
        if (string.IsNullOrWhiteSpace(targetDirectory))
            throw new ArgumentException("An explicit target directory is required.", nameof(targetDirectory));
        if (!Path.IsPathFullyQualified(targetDirectory))
            throw new ArgumentException("The target directory must be a fully qualified path.", nameof(targetDirectory));

        var target = Path.TrimEndingDirectorySeparator(Path.GetFullPath(targetDirectory));
        var parent = Path.GetDirectoryName(target);
        if (string.IsNullOrEmpty(parent) || string.Equals(target, Path.GetPathRoot(target),
                StringComparison.Ordinal))
            throw new ArgumentException("The target directory must have a parent.", nameof(targetDirectory));
        if (!Directory.Exists(parent))
            throw new DirectoryNotFoundException($"Target parent directory does not exist: '{parent}'.");

        EnsureNoSymlinks(parent, includeLeaf: true);
        EnsureTargetDoesNotExist(target);
        return target;
    }

    private static void EnsureTargetStillAvailable(string target, string parent)
    {
        if (!Directory.Exists(parent))
            throw new DirectoryNotFoundException($"Target parent directory no longer exists: '{parent}'.");
        EnsureNoSymlinks(parent, includeLeaf: true);
        EnsureTargetDoesNotExist(target);
    }

    private static void EnsureTargetDoesNotExist(string target)
    {
        if (PathEntryExists(target))
            throw new IOException($"Delivery target already exists: '{target}'.");
    }

    private static DeliveryBundleDirectoryPublicationSnapshot SnapshotAndValidate(
        DeliveryBundleDirectoryPublicationSource source)
    {
        if (source.Artifacts is null)
            throw new ArgumentException("Bundle artifacts are required.", nameof(source));
        if (source.VerifyStagedBytes is null)
            throw new ArgumentException("A staged-byte verifier is required.", nameof(source));

        var artifacts = new List<DeliveryBundleDirectoryPublicationArtifactSnapshot>(
            source.Artifacts.Count);
        foreach (var artifact in source.Artifacts)
        {
            if (artifact is null)
                throw new ArgumentException("Bundle artifacts cannot contain null entries.", nameof(source));
            artifacts.Add(new DeliveryBundleDirectoryPublicationArtifactSnapshot(
                ValidateRelativePath(artifact.RelativePath, nameof(source)),
                artifact.Bytes.ToArray()));
        }

        var manifestPath = ValidateRelativePath(source.ManifestRelativePath, nameof(source));
        var allPaths = artifacts.Select(item => item.RelativePath).Append(manifestPath).ToArray();
        ValidatePathSet(allPaths, nameof(source));
        if (allPaths.Any(path => path.Equals(StageMarkerName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException(
                $"Bundle path is reserved by atomic publication: '{StageMarkerName}'.",
                nameof(source));

        return new DeliveryBundleDirectoryPublicationSnapshot(
            artifacts.AsReadOnly(),
            manifestPath,
            source.CanonicalManifestBytes.ToArray());
    }

    private static string ValidateRelativePath(string relativePath, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
            throw new ArgumentException("Artifact paths must be non-empty canonical relative paths.", parameterName);
        if (Path.IsPathRooted(relativePath)
            || relativePath.StartsWith("/", StringComparison.Ordinal)
            || relativePath.StartsWith("\\", StringComparison.Ordinal)
            || (relativePath.Length >= 2 && char.IsAsciiLetter(relativePath[0])
                && relativePath[1] == ':'))
            throw new ArgumentException($"Artifact path must be relative: '{relativePath}'.", parameterName);
        if (relativePath.Contains('\\', StringComparison.Ordinal))
            throw new ArgumentException(
                $"Artifact path must use canonical '/' separators: '{relativePath}'.", parameterName);

        var segments = relativePath.Split('/');
        if (segments.Any(segment => segment.Length == 0 || segment is "." or ".."))
            throw new ArgumentException(
                $"Artifact path contains an empty, current, or parent segment: '{relativePath}'.",
                parameterName);

        foreach (var segment in segments)
            ValidatePathSegment(segment, relativePath, parameterName);

        // This containment check is deliberately in addition to lexical segment checks. It keeps
        // the invariant explicit if platform path rules expand in the future.
        var probeRoot = Path.Combine(Path.GetTempPath(), "docxodus-delivery-path-probe");
        var combined = Path.GetFullPath(Path.Combine(probeRoot,
            relativePath.Replace('/', Path.DirectorySeparatorChar)));
        if (!IsStrictDescendant(probeRoot, combined))
            throw new ArgumentException($"Artifact path escapes its bundle: '{relativePath}'.", parameterName);

        return string.Join('/', segments);
    }

    private static void ValidatePathSegment(
        string segment,
        string relativePath,
        string parameterName)
    {
        if (segment.EndsWith(' ') || segment.EndsWith('.'))
            throw new ArgumentException(
                $"Artifact path has a non-portable trailing character: '{relativePath}'.",
                parameterName);
        if (segment.Any(character => character < ' '
                || character is '\0' or '<' or '>' or ':' or '"' or '|' or '?' or '*'))
            throw new ArgumentException(
                $"Artifact path contains a non-portable filename character: '{relativePath}'.",
                parameterName);

        var stem = segment.Split('.', 2)[0];
        if (stem.Equals("CON", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("PRN", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("AUX", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("NUL", StringComparison.OrdinalIgnoreCase)
            || (stem.Length == 4
                && (stem.StartsWith("COM", StringComparison.OrdinalIgnoreCase)
                    || stem.StartsWith("LPT", StringComparison.OrdinalIgnoreCase))
                && stem[3] is >= '1' and <= '9'))
            throw new ArgumentException(
                $"Artifact path contains a reserved filename: '{relativePath}'.", parameterName);
    }

    private static void ValidatePathSet(IReadOnlyList<string> paths, string parameterName)
    {
        var exact = new HashSet<string>(StringComparer.Ordinal);
        var folded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in paths)
        {
            if (!exact.Add(path))
                throw new ArgumentException($"Duplicate bundle path: '{path}'.", parameterName);
            if (!folded.Add(path))
                throw new ArgumentException($"Case-colliding bundle path: '{path}'.", parameterName);
        }

        var ordered = paths.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
        for (var index = 0; index < ordered.Length; index++)
        {
            for (var other = index + 1; other < ordered.Length; other++)
            {
                if (ordered[other].StartsWith(ordered[index] + '/',
                    StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        $"Bundle path '{ordered[index]}' conflicts with descendant '{ordered[other]}'.",
                        parameterName);
            }
        }
    }

    private static (string Path, bool Owned) CreateOwnedStage(string parent, string targetName)
    {
        for (var attempt = 0; attempt < StageCreationAttempts; attempt++)
        {
            var stage = Path.Combine(parent,
                $".{targetName}.docxodus-stage-{Guid.NewGuid():N}");
            if (PathEntryExists(stage))
                continue;

            Directory.CreateDirectory(stage);
            try
            {
                using var marker = new FileStream(
                    Path.Combine(stage, StageMarkerName),
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None,
                    bufferSize: 1,
                    FileOptions.WriteThrough);
                marker.WriteByte(1);
                marker.Flush(flushToDisk: true);
                return (stage, true);
            }
            catch
            {
                TryDeleteOwnedStage(stage);
                throw;
            }
        }

        throw new IOException("Could not allocate an owned delivery staging directory.");
    }

    private static void WriteNewFile(string stage, string relativePath, byte[] bytes)
    {
        var destination = ResolveStagePath(stage, relativePath);
        CreateSafeParentDirectories(stage, Path.GetDirectoryName(destination)!);

        if (PathEntryExists(destination))
            throw new IOException($"Staged bundle path already exists: '{relativePath}'.");

        using var stream = new FileStream(
            destination,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 64 * 1024,
            FileOptions.WriteThrough);
        stream.Write(bytes);
        stream.Flush(flushToDisk: true);
    }

    private static void CreateSafeParentDirectories(string stage, string destinationParent)
    {
        var relative = Path.GetRelativePath(stage, destinationParent);
        if (relative == ".")
            return;

        var current = stage;
        foreach (var segment in relative.Split(Path.DirectorySeparatorChar))
        {
            current = Path.Combine(current, segment);
            if (PathEntryExists(current))
            {
                if (!Directory.Exists(current) || IsSymlink(current))
                    throw new IOException($"Staged bundle path traverses a symbolic link: '{current}'.");
                continue;
            }

            Directory.CreateDirectory(current);
            if (IsSymlink(current))
                throw new IOException($"Staged bundle path became a symbolic link: '{current}'.");
        }
    }

    private static IReadOnlyDictionary<string, byte[]> ReadExpectedStage(
        string stage,
        DeliveryBundleDirectoryPublicationSnapshot snapshot)
    {
        if (!Directory.Exists(stage) || IsSymlink(stage))
            throw new IOException("The owned delivery staging directory is missing or unsafe.");

        var expectedPaths = snapshot.Artifacts.Select(item => item.RelativePath)
            .Append(snapshot.ManifestRelativePath)
            .ToHashSet(StringComparer.Ordinal);
        var actualPaths = EnumerateSafeStageFiles(stage);
        if (!actualPaths.SetEquals(expectedPaths))
        {
            var unexpected = actualPaths.Except(expectedPaths, StringComparer.Ordinal)
                .OrderBy(path => path, StringComparer.Ordinal).FirstOrDefault();
            var missing = expectedPaths.Except(actualPaths, StringComparer.Ordinal)
                .OrderBy(path => path, StringComparer.Ordinal).FirstOrDefault();
            throw new IOException(unexpected is not null
                ? $"Unexpected file appeared in the delivery stage: '{unexpected}'."
                : $"Expected file is missing from the delivery stage: '{missing}'.");
        }

        var bytes = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var relativePath in expectedPaths.OrderBy(path => path, StringComparer.Ordinal))
        {
            var fullPath = ResolveStagePath(stage, relativePath);
            EnsureNoSymlinks(fullPath, includeLeaf: true, stopAt: stage);
            bytes.Add(relativePath, File.ReadAllBytes(fullPath));
        }

        return new ReadOnlyDictionary<string, byte[]>(bytes);
    }

    private static HashSet<string> EnumerateSafeStageFiles(string stage)
    {
        var files = new HashSet<string>(StringComparer.Ordinal);
        var caseFolded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Stack<string>();
        pending.Push(stage);

        while (pending.Count > 0)
        {
            var directory = pending.Pop();
            foreach (var entry in new DirectoryInfo(directory).EnumerateFileSystemInfos())
            {
                if (entry.FullName == Path.Combine(stage, StageMarkerName))
                    continue;
                if (IsSymlink(entry.FullName))
                    throw new IOException($"Symbolic links are not allowed in a delivery stage: '{entry.FullName}'.");

                if ((entry.Attributes & FileAttributes.Directory) != 0)
                {
                    pending.Push(entry.FullName);
                    continue;
                }

                var relative = Path.GetRelativePath(stage, entry.FullName)
                    .Replace(Path.DirectorySeparatorChar, '/');
                if (!files.Add(relative) || !caseFolded.Add(relative))
                    throw new IOException($"Duplicate or case-colliding staged path: '{relative}'.");
            }
        }

        return files;
    }

    private static void EnsureByteEquality(
        IReadOnlyDictionary<string, byte[]> verified,
        IReadOnlyDictionary<string, byte[]> preCommit)
    {
        if (verified.Count != preCommit.Count)
            throw new IOException("The staged delivery bundle changed after verification.");

        foreach (var item in verified)
        {
            if (!preCommit.TryGetValue(item.Key, out var bytes)
                || !item.Value.AsSpan().SequenceEqual(bytes))
                throw new IOException(
                    $"Staged delivery artifact changed after verification: '{item.Key}'.");
        }
    }

    private static string ResolveStagePath(string stage, string relativePath)
    {
        var destination = Path.GetFullPath(Path.Combine(stage,
            relativePath.Replace('/', Path.DirectorySeparatorChar)));
        if (!IsStrictDescendant(stage, destination))
            throw new IOException($"Bundle path escapes the owned delivery stage: '{relativePath}'.");
        return destination;
    }

    private static bool IsStrictDescendant(string parent, string candidate)
    {
        var relative = Path.GetRelativePath(parent, candidate);
        if (relative == "." || Path.IsPathRooted(relative) || relative == "..")
            return false;
        if (relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            return false;
        if (Path.AltDirectorySeparatorChar != Path.DirectorySeparatorChar
            && relative.StartsWith(".." + Path.AltDirectorySeparatorChar,
                StringComparison.Ordinal))
            return false;
        return true;
    }

    private static void EnsureNoSymlinks(
        string path,
        bool includeLeaf,
        string? stopAt = null)
    {
        var full = Path.GetFullPath(path);
        var stop = stopAt is null ? Path.GetPathRoot(full)! : Path.GetFullPath(stopAt);
        if (!string.Equals(stop, Path.GetPathRoot(full), StringComparison.Ordinal)
            && !string.Equals(full, stop, StringComparison.Ordinal)
            && !IsStrictDescendant(stop, full))
            throw new IOException($"Path is outside the expected root: '{path}'.");

        var relative = Path.GetRelativePath(stop, full);
        var segments = relative == "."
            ? Array.Empty<string>()
            : relative.Split(Path.DirectorySeparatorChar);
        var current = stop;
        for (var index = 0; index < segments.Length; index++)
        {
            current = Path.Combine(current, segments[index]);
            if (!includeLeaf && index == segments.Length - 1)
                break;
            if (PathEntryExists(current) && IsSymlink(current))
                throw new IOException($"Symbolic-link path components are not allowed: '{current}'.");
        }
    }

    private static bool PathEntryExists(string path)
    {
        if (File.Exists(path) || Directory.Exists(path))
            return true;

        try
        {
            return new FileInfo(path).LinkTarget is not null
                || new DirectoryInfo(path).LinkTarget is not null;
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException
            or System.Security.SecurityException)
        {
            return false;
        }
    }

    private static bool IsSymlink(string path)
    {
        try
        {
            var info = Directory.Exists(path)
                ? (FileSystemInfo)new DirectoryInfo(path)
                : new FileInfo(path);
            return info.LinkTarget is not null
                || (info.Attributes & FileAttributes.ReparsePoint) != 0;
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException
            or System.Security.SecurityException)
        {
            throw new IOException($"Could not inspect delivery path '{path}'.", exception);
        }
    }

    private static void TryDeleteOwnedStage(string stage)
    {
        try
        {
            if (Directory.Exists(stage) && !IsSymlink(stage))
                Directory.Delete(stage, recursive: true);
            else if (PathEntryExists(stage))
                File.Delete(stage);
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException
            or System.Security.SecurityException)
        {
            // Best-effort cleanup must never replace the operation's original failure. The random
            // stage name and in-memory ownership bit ensure no unrelated directory is selected.
        }
    }

    private sealed record DeliveryBundleDirectoryPublicationArtifactSnapshot(
        string RelativePath,
        byte[] Bytes);

    private sealed record DeliveryBundleDirectoryPublicationSnapshot(
        ReadOnlyCollection<DeliveryBundleDirectoryPublicationArtifactSnapshot> Artifacts,
        string ManifestRelativePath,
        byte[] ManifestBytes);
}

/// <summary>Minimal internal adapter between the filesystem publisher and bundle model owner.</summary>
internal sealed class DeliveryBundleDirectoryPublicationSource
{
    required internal IReadOnlyList<DeliveryBundleDirectoryPublicationArtifact> Artifacts { get; init; }
    required internal string ManifestRelativePath { get; init; }
    required internal ReadOnlyMemory<byte> CanonicalManifestBytes { get; init; }
    required internal Action<IReadOnlyDictionary<string, byte[]>> VerifyStagedBytes { get; init; }
}

/// <summary>One available artifact projected from a delivery bundle for directory publication.</summary>
internal sealed record DeliveryBundleDirectoryPublicationArtifact(
    string RelativePath,
    ReadOnlyMemory<byte> Bytes);

internal enum DeliveryBundleDirectoryPublisherCheckpoint
{
    BeforeArtifactWrite,
    BeforeManifestWrite,
    BeforeVerification,
    BeforeCommit,
}

internal sealed record DeliveryBundleDirectoryPublisherFaultContext(
    DeliveryBundleDirectoryPublisherCheckpoint Checkpoint,
    string TargetDirectory,
    string StageDirectory,
    string? RelativePath,
    int? ArtifactIndex);

internal interface IDeliveryBundleDirectoryPublisherFaultInjector
{
    void OnCheckpoint(DeliveryBundleDirectoryPublisherFaultContext context);
}

internal sealed class NoDeliveryBundleDirectoryPublisherFaultInjector
    : IDeliveryBundleDirectoryPublisherFaultInjector
{
    internal static readonly NoDeliveryBundleDirectoryPublisherFaultInjector Instance = new();

    private NoDeliveryBundleDirectoryPublisherFaultInjector()
    {
    }

    public void OnCheckpoint(DeliveryBundleDirectoryPublisherFaultContext context)
    {
    }
}
