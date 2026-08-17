// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace Docxodus.Delivery;

/// <summary>Explicit publication policy for non-delivery diagnostic bundles.</summary>
public sealed record DeliveryBundleDirectoryPublicationOptions
{
    /// <summary>
    /// Permit an incomplete or failed manifest to be published for diagnostics. The target still
    /// uses the same verified, no-replace atomic commit, but callers must not present it as a
    /// successful delivery.
    /// </summary>
    public bool AllowDiagnosticBundle { get; init; }
}

/// <summary>
/// Publishes a verified delivery bundle as a new directory. Publication never replaces an
/// existing path: all bytes are written to an owned sibling directory, the manifest is written
/// last, and a same-filesystem directory rename is the only commit point.
/// </summary>
public static class DeliveryBundleDirectoryPublisher
{
    private const string StageMarkerName = ".docxodus-delivery-stage";
    private const int StageCreationAttempts = 32;
    private const int AtFileDescriptorCurrentWorkingDirectory = -100;
    private const uint RenameNoReplace = 1;
    private const uint RenameExclusive = 0x00000004;
    private const int LockExclusive = 2;
    private const int LockNonBlocking = 4;
    private const UnixFileMode PrivateDirectoryMode = UnixFileMode.UserRead
        | UnixFileMode.UserWrite | UnixFileMode.UserExecute;
    private const UnixFileMode PrivateFileMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
    private static readonly object[] PublicationLeaseStripes =
        Enumerable.Range(0, 64).Select(_ => new object()).ToArray();

    /// <summary>
    /// Publish all available artifacts and the canonical manifest to a new directory. The
    /// manifest is independently parsed and verified from staged filesystem bytes before commit.
    /// </summary>
    public static string Publish(
        DeliveryBundle bundle,
        string targetDirectory,
        DeliveryBundleVerificationLimits? verificationLimits = null,
        DeliveryBundleDirectoryPublicationOptions? publicationOptions = null)
    {
        ArgumentNullException.ThrowIfNull(bundle);
        publicationOptions ??= new DeliveryBundleDirectoryPublicationOptions();
        if (bundle.Manifest.Payload.Status != DeliveryBundleStatus.Complete
            && !publicationOptions.AllowDiagnosticBundle)
            throw new DeliveryBundleException(
                "diagnostic_bundle_publication_not_enabled",
                $"Bundle status '{bundle.Manifest.Payload.Status}' requires explicit diagnostic publication.");
        return Publish(bundle, targetDirectory, verificationLimits, faultInjector: null);
    }

    internal static string Publish(
        DeliveryBundle bundle,
        string targetDirectory,
        DeliveryBundleVerificationLimits? verificationLimits,
        IDeliveryBundleDirectoryPublisherFaultInjector? faultInjector)
    {
        ArgumentNullException.ThrowIfNull(bundle);
        verificationLimits ??= new DeliveryBundleVerificationLimits();
        verificationLimits.Validate();
        if (bundle.Manifest.Payload.Artifacts.Count > verificationLimits.MaxArtifacts)
            throw new InvalidDataException("Delivery artifact count exceeds the publication resource limit.");
        if (bundle.Manifest.Payload.Relationships.Count > verificationLimits.MaxRelationships)
            throw new InvalidDataException("Delivery relationship count exceeds the publication resource limit.");
        var bundleBytes = bundle.OwnedArtifactBytes;
        var preflight = DeliveryBundleVerifier.Verify(
            bundle.Manifest, bundleBytes, verificationLimits);
        if (!preflight.IsValid)
            throw new InvalidDataException(
                $"Delivery bundle exceeds publication limits or is invalid: {preflight.Findings[0]}");

        var manifestBytes = bundle.ManifestBytes;
        if (manifestBytes.LongLength > verificationLimits.MaxManifestBytes)
            throw new InvalidDataException("Delivery manifest exceeds the publication resource limit.");
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
        long totalBytes = 0;
        foreach (var artifact in availableArtifacts)
        {
            if (!bundleBytes.TryGetValue(artifact.ArtifactId, out var bytes))
                continue;
            if (bytes.LongLength > verificationLimits.MaxArtifactBytes)
                throw new InvalidDataException(
                    $"Delivery artifact '{artifact.ArtifactId}' exceeds the publication resource limit.");
            if (totalBytes > verificationLimits.MaxTotalArtifactBytes - bytes.LongLength)
                throw new InvalidDataException(
                    "Delivery artifacts exceed the total publication resource limit.");
            totalBytes += bytes.LongLength;
        }
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
            CanonicalManifestBytes = manifestBytes,
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
        using var publicationLease = AcquirePublicationLease(target, parent);
        EnsureTargetStillAvailable(target, parent);

        string? stage = null;
        var stageOwned = false;
        byte[]? stageMarkerToken = null;
        FileStream? stageMarkerHandle = null;
        try
        {
            (stage, stageMarkerToken, stageMarkerHandle) = CreateOwnedStage(
                parent, Path.GetFileName(target));
            stageOwned = true;

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

            {
                var stagedBytes = ReadExpectedStage(stage, snapshot, stageMarkerToken);
                bundle.VerifyStagedBytes(stagedBytes);
            }

            injector.OnCheckpoint(new DeliveryBundleDirectoryPublisherFaultContext(
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit,
                target,
                stage,
                null,
                null));

            // Verification is over fresh reads from the stage. Re-read immediately before the
            // commit so a fault hook or concurrent process cannot change verified bytes in place.
            EnsureStageMatchesSnapshot(stage, snapshot, stageMarkerToken);
            EnsureTargetStillAvailable(target, parent);

            injector.OnCheckpoint(new DeliveryBundleDirectoryPublisherFaultContext(
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit,
                target,
                stage,
                null,
                null));

            // Keep the authenticated ownership marker present across every callback. The only
            // marker-free interval is the non-callback sequence immediately before rename.
            EnsureStageMatchesSnapshot(stage, snapshot, stageMarkerToken);
            EnsureTargetStillAvailable(target, parent);
            stageMarkerHandle.Dispose();
            stageMarkerHandle = null;
            File.Delete(Path.Combine(stage, StageMarkerName));
            MoveDirectoryNoReplace(stage, target);
            stageMarkerToken = null;
            stageOwned = false;
            return target;
        }
        catch (Exception publicationFailure)
        {
            if (stageOwned && stage is not null)
            {
                var cleanupFailure = TryDeleteOwnedStage(
                    stage, stageMarkerToken, stageMarkerHandle, snapshot);
                if (cleanupFailure is not null)
                {
                    throw new IOException(
                        $"Delivery publication failed and its owned staging directory could not "
                        + $"be removed safely: '{stage}'.",
                        new AggregateException(publicationFailure, cleanupFailure));
                }
            }
            throw;
        }
        finally
        {
            stageMarkerHandle?.Dispose();
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

    private static IDisposable AcquirePublicationLease(string target, string parent)
    {
        var targetDigest = Convert.ToHexString(
                SHA256.HashData(Encoding.UTF8.GetBytes(target)))
            .ToLowerInvariant()[..24];
        var path = Path.Combine(
            parent, $".docxodus-delivery-publish-{targetDigest}.lock");
        var stripe = PublicationLeaseStripes[
            Convert.ToInt32(targetDigest[..2], 16) % PublicationLeaseStripes.Length];
        Monitor.Enter(stripe);
        FileStream? lease = null;
        var ownershipTransferred = false;
        try
        {
            if (PathEntryExists(path) && IsSymlink(path))
                throw new IOException($"Delivery publication lease path is unsafe: '{path}'.");

            lease = OpenPrivateFile(
                path, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                FileShare.ReadWrite, FileOptions.WriteThrough, bufferSize: 32);
            if (IsSymlink(path))
                throw new IOException($"Delivery publication lease path is unsafe: '{path}'.");
            AcquireOperatingSystemLease(lease);
            ownershipTransferred = true;
            return new PublicationLease(lease, stripe);
        }
        catch (IOException exception)
        {
            throw new IOException(
                $"Another publisher already owns the delivery target commit lease: '{target}'.",
                exception);
        }
        finally
        {
            if (!ownershipTransferred)
            {
                lease?.Dispose();
                Monitor.Exit(stripe);
            }
        }
    }

    private static void AcquireOperatingSystemLease(FileStream lease)
    {
        if (!OperatingSystem.IsMacOS())
        {
            lease.Lock(0, 1);
            return;
        }

        // FileStream.Lock is unsupported on macOS. flock provides the same process-crash-safe,
        // close-released advisory lease without changing the persistent sidecar's contents.
        var descriptor = lease.SafeFileHandle.DangerousGetHandle().ToInt32();
        if (LockFileMac(descriptor, LockExclusive | LockNonBlocking) != 0)
            throw new IOException(
                "The delivery publication lease is already held.",
                new Win32Exception(Marshal.GetLastPInvokeError()));
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
                && IsDosDeviceDigit(stem[3]))
            || stem.Equals("CONIN$", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("CONOUT$", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("CLOCK$", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Artifact path contains a reserved filename: '{relativePath}'.", parameterName);
    }

    private static bool IsDosDeviceDigit(char value) => value is >= '1' and <= '9'
        or '\u00B9' or '\u00B2' or '\u00B3';

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

    private static (string Path, byte[] MarkerToken, FileStream MarkerHandle) CreateOwnedStage(
        string parent,
        string targetName)
    {
        for (var attempt = 0; attempt < StageCreationAttempts; attempt++)
        {
            var stage = Path.Combine(parent,
                $".{targetName}.docxodus-stage-{Guid.NewGuid():N}");
            if (!TryCreatePrivateDirectory(stage))
                continue;
            var markerToken = RandomNumberGenerator.GetBytes(32);
            FileStream? marker = null;
            try
            {
                marker = CreateStageMarker(stage, markerToken);
                return (stage, markerToken, marker);
            }
            catch
            {
                marker?.Dispose();
                TryDeleteFreshStage(stage);
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

        using var stream = OpenPrivateFile(
            destination, FileMode.CreateNew, FileAccess.Write, FileShare.None,
            FileOptions.WriteThrough, 64 * 1024);
        stream.Write(bytes);
        stream.Flush(flushToDisk: true);
    }

    private static FileStream CreateStageMarker(string stage, byte[] markerToken)
    {
        var marker = OpenPrivateFile(
            Path.Combine(stage, StageMarkerName), FileMode.CreateNew,
            FileAccess.ReadWrite, FileShare.Read | FileShare.Delete,
            FileOptions.WriteThrough, markerToken.Length);
        try
        {
            marker.Write(markerToken);
            marker.Flush(flushToDisk: true);
            return marker;
        }
        catch
        {
            marker.Dispose();
            throw;
        }
    }

    private static void TryDeleteFreshStage(string stage)
    {
        try
        {
            if (!Directory.Exists(stage) || IsSymlink(stage))
                return;
            var entries = new DirectoryInfo(stage).EnumerateFileSystemInfos().ToArray();
            if (entries.Any(entry =>
                    !string.Equals(entry.Name, StageMarkerName, StringComparison.Ordinal)
                    || IsSymlink(entry.FullName)
                    || (entry.Attributes & FileAttributes.Directory) != 0))
                return;
            Directory.Delete(stage, recursive: true);
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException
            or System.Security.SecurityException)
        {
            // A setup failure remains primary. A path no longer shaped like the just-created
            // empty stage is deliberately preserved rather than deleted recursively.
        }
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

            CreatePrivateDirectory(current);
            if (IsSymlink(current))
                throw new IOException($"Staged bundle path became a symbolic link: '{current}'.");
        }
    }

    private static IReadOnlyDictionary<string, byte[]> ReadExpectedStage(
        string stage,
        DeliveryBundleDirectoryPublicationSnapshot snapshot,
        byte[] markerToken)
    {
        if (!Directory.Exists(stage) || IsSymlink(stage))
            throw new IOException("The owned delivery staging directory is missing or unsafe.");

        var expectedBytes = snapshot.Artifacts.ToDictionary(
            item => item.RelativePath, item => item.Bytes, StringComparer.Ordinal);
        expectedBytes.Add(snapshot.ManifestRelativePath, snapshot.ManifestBytes);
        var expectedPaths = expectedBytes.Keys.ToHashSet(StringComparer.Ordinal);
        var markerPath = Path.Combine(stage, StageMarkerName);
        if (!File.Exists(markerPath) || IsSymlink(markerPath)
            || !MarkerMatches(markerPath, markerToken))
            throw new IOException("The owned delivery staging marker is missing or invalid.");
        var actualPaths = EnumerateSafeStageFiles(
            stage, expectedPaths, allowStageMarker: true);
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
            bytes.Add(relativePath, ReadExactFile(
                fullPath, relativePath, expectedBytes[relativePath].Length));
        }
        if (!MarkerMatches(markerPath, markerToken))
            throw new IOException("The owned delivery staging marker changed during inspection.");

        return new ReadOnlyDictionary<string, byte[]>(bytes);
    }

    private static byte[] ReadExactFile(
        string fullPath,
        string relativePath,
        int expectedLength)
    {
        var metadataLength = new FileInfo(fullPath).Length;
        if (metadataLength != expectedLength)
            throw UnexpectedStagedLength(relativePath, expectedLength, metadataLength);
        using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            bufferSize: 64 * 1024,
            FileOptions.SequentialScan);
        if (stream.Length != expectedLength)
            throw UnexpectedStagedLength(relativePath, expectedLength, stream.Length);

        var bytes = GC.AllocateUninitializedArray<byte>(expectedLength);
        try
        {
            stream.ReadExactly(bytes);
        }
        catch (EndOfStreamException exception)
        {
            throw UnexpectedStagedLength(
                relativePath, expectedLength, stream.Position, exception);
        }
        if (stream.ReadByte() != -1 || stream.Length != expectedLength)
            throw UnexpectedStagedLength(relativePath, expectedLength, stream.Length);
        return bytes;
    }

    private static IOException UnexpectedStagedLength(
        string relativePath,
        long expectedLength,
        long actualLength,
        Exception? innerException = null) => new(
        $"Staged delivery artifact changed after verification or initial write: "
        + $"'{relativePath}' has length {actualLength}, expected {expectedLength}.",
        innerException);

    private static HashSet<string> EnumerateSafeStageFiles(
        string stage,
        IReadOnlySet<string> expectedPaths,
        bool allowStageMarker)
    {
        var files = new HashSet<string>(StringComparer.Ordinal);
        var caseFolded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Stack<string>();
        pending.Push(stage);
        var maximumEntries = checked((expectedPaths.Count * 2) + 16);
        var inspectedEntries = 0;

        while (pending.Count > 0)
        {
            var directory = pending.Pop();
            foreach (var entry in new DirectoryInfo(directory).EnumerateFileSystemInfos())
            {
                if (++inspectedEntries > maximumEntries)
                    throw new IOException("Delivery stage entry count exceeds its declared shape.");
                if (entry.FullName == Path.Combine(stage, StageMarkerName)
                    && allowStageMarker)
                    continue;
                if (IsSymlink(entry.FullName))
                    throw new IOException($"Symbolic links are not allowed in a delivery stage: '{entry.FullName}'.");

                var relative = Path.GetRelativePath(stage, entry.FullName)
                    .Replace(Path.DirectorySeparatorChar, '/');

                if ((entry.Attributes & FileAttributes.Directory) != 0)
                {
                    if (!expectedPaths.Any(expected => expected.StartsWith(
                            relative + '/', StringComparison.Ordinal)))
                        throw new IOException(
                            $"Unexpected directory appeared in the delivery stage: '{relative}'.");
                    pending.Push(entry.FullName);
                    continue;
                }

                if (!files.Add(relative) || !caseFolded.Add(relative))
                    throw new IOException($"Duplicate or case-colliding staged path: '{relative}'.");
            }
        }

        return files;
    }

    private static void EnsureStageMatchesSnapshot(
        string stage,
        DeliveryBundleDirectoryPublicationSnapshot snapshot,
        byte[] markerToken)
    {
        if (!Directory.Exists(stage) || IsSymlink(stage))
            throw new IOException("The owned delivery staging directory is missing or unsafe.");
        var expected = snapshot.Artifacts.ToDictionary(
            item => item.RelativePath, item => item.Bytes, StringComparer.Ordinal);
        expected.Add(snapshot.ManifestRelativePath, snapshot.ManifestBytes);
        var markerPath = Path.Combine(stage, StageMarkerName);
        if (!File.Exists(markerPath) || IsSymlink(markerPath)
            || !MarkerMatches(markerPath, markerToken))
            throw new IOException("The owned delivery staging marker is missing or invalid.");
        var actualPaths = EnumerateSafeStageFiles(
            stage, expected.Keys.ToHashSet(StringComparer.Ordinal), allowStageMarker: true);
        if (!actualPaths.SetEquals(expected.Keys))
            throw new IOException("The staged delivery bundle changed after verification.");

        foreach (var item in expected.OrderBy(item => item.Key, StringComparer.Ordinal))
        {
            var fullPath = ResolveStagePath(stage, item.Key);
            EnsureNoSymlinks(fullPath, includeLeaf: true, stopAt: stage);
            if (!FileMatches(fullPath, item.Key, item.Value))
                throw new IOException(
                    $"Staged delivery artifact changed after verification: '{item.Key}'.");
        }
        if (!MarkerMatches(markerPath, markerToken))
            throw new IOException("The owned delivery staging marker changed during inspection.");
    }

    private static bool FileMatches(string fullPath, string relativePath, byte[] expected)
    {
        var metadataLength = new FileInfo(fullPath).Length;
        if (metadataLength != expected.LongLength)
            throw UnexpectedStagedLength(relativePath, expected.LongLength, metadataLength);
        using var stream = new FileStream(
            fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete,
            bufferSize: 64 * 1024, FileOptions.SequentialScan);
        if (stream.Length != expected.LongLength)
            throw UnexpectedStagedLength(relativePath, expected.LongLength, stream.Length);
        var buffer = new byte[Math.Min(64 * 1024, Math.Max(expected.Length, 1))];
        var offset = 0;
        while (offset < expected.Length)
        {
            var count = Math.Min(buffer.Length, expected.Length - offset);
            stream.ReadExactly(buffer.AsSpan(0, count));
            if (!buffer.AsSpan(0, count).SequenceEqual(expected.AsSpan(offset, count)))
                return false;
            offset += count;
        }
        return stream.ReadByte() == -1 && stream.Length == expected.LongLength;
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

    private static Exception? TryDeleteOwnedStage(
        string stage,
        byte[]? markerToken,
        FileStream? markerHandle,
        DeliveryBundleDirectoryPublicationSnapshot snapshot)
    {
        try
        {
            if (!PathEntryExists(stage))
                return null;
            if (!Directory.Exists(stage) || IsSymlink(stage))
                throw new IOException(
                    $"The owned delivery stage is no longer a safe directory: '{stage}'.");
            if (markerToken is null)
                throw new IOException(
                    $"The delivery stage no longer has in-memory ownership evidence: '{stage}'.");
            if (markerHandle is null || markerHandle.SafeFileHandle.IsClosed)
                throw new IOException(
                    $"The delivery stage no longer has an open ownership handle: '{stage}'.");
            var markerPath = Path.Combine(stage, StageMarkerName);
            var challenge = RandomNumberGenerator.GetBytes(markerToken.Length);
            markerHandle.Position = 0;
            markerHandle.SetLength(0);
            markerHandle.Write(challenge);
            markerHandle.Flush(flushToDisk: true);
            if (!File.Exists(markerPath) || IsSymlink(markerPath)
                || !MarkerMatches(markerPath, challenge))
                throw new IOException(
                    $"The owned delivery stage marker is missing or invalid: '{markerPath}'.");

            EnsureCleanupTreeMatches(stage, snapshot, challenge);

            Directory.Delete(stage, recursive: true);
            return null;
        }
        catch (Exception exception) when (exception is IOException
            or UnauthorizedAccessException
            or System.Security.SecurityException)
        {
            return exception;
        }
    }

    private static void EnsureCleanupTreeMatches(
        string stage,
        DeliveryBundleDirectoryPublicationSnapshot snapshot,
        byte[] markerToken)
    {
        var expectedBytes = snapshot.Artifacts.ToDictionary(
            item => item.RelativePath, item => item.Bytes, StringComparer.Ordinal);
        expectedBytes.Add(snapshot.ManifestRelativePath, snapshot.ManifestBytes);
        var actualPaths = EnumerateSafeStageFiles(
            stage, expectedBytes.Keys.ToHashSet(StringComparer.Ordinal),
            allowStageMarker: true);
        foreach (var relativePath in actualPaths)
        {
            if (!expectedBytes.TryGetValue(relativePath, out var expected))
                throw new IOException(
                    $"Unexpected file appeared in the delivery stage: '{relativePath}'.");
            var fullPath = ResolveStagePath(stage, relativePath);
            EnsureNoSymlinks(fullPath, includeLeaf: true, stopAt: stage);
            var actual = ReadExactFile(fullPath, relativePath, expected.Length);
            if (!actual.AsSpan().SequenceEqual(expected))
                throw new IOException(
                    $"Staged delivery artifact changed before cleanup: '{relativePath}'.");
        }

        if (!MarkerMatches(Path.Combine(stage, StageMarkerName), markerToken))
            throw new IOException("The owned delivery staging marker changed during cleanup.");
    }

    private static bool MarkerMatches(string markerPath, byte[] expected)
    {
        using var stream = new FileStream(
            markerPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            bufferSize: expected.Length,
            FileOptions.SequentialScan);
        if (stream.Length != expected.Length)
            return false;
        var actual = GC.AllocateUninitializedArray<byte>(expected.Length);
        try
        {
            stream.ReadExactly(actual);
        }
        catch (EndOfStreamException)
        {
            return false;
        }
        return stream.ReadByte() == -1
            && stream.Length == expected.Length
            && CryptographicOperations.FixedTimeEquals(actual, expected);
    }

    private static void CreatePrivateDirectory(string path)
    {
        if (OperatingSystem.IsWindows())
            Directory.CreateDirectory(path);
        else
            Directory.CreateDirectory(path, PrivateDirectoryMode);
    }

    private static bool TryCreatePrivateDirectory(string path)
    {
        if (OperatingSystem.IsWindows())
        {
            if (CreateDirectoryWindows(path, IntPtr.Zero))
                return true;
            const int alreadyExists = 183;
            var error = Marshal.GetLastPInvokeError();
            if (error == alreadyExists)
                return false;
            throw new IOException(
                $"Could not create private delivery staging directory '{path}'.",
                new Win32Exception(error));
        }

        if (OperatingSystem.IsLinux() || OperatingSystem.IsMacOS())
        {
            if (CreateDirectoryUnix(path, Convert.ToUInt32(PrivateDirectoryMode)) == 0)
                return true;
            const int alreadyExists = 17;
            var error = Marshal.GetLastPInvokeError();
            if (error == alreadyExists)
                return false;
            throw new IOException(
                $"Could not create private delivery staging directory '{path}'.",
                new Win32Exception(error));
        }

        if (PathEntryExists(path))
            return false;
        CreatePrivateDirectory(path);
        return true;
    }

    private static FileStream OpenPrivateFile(
        string path,
        FileMode mode,
        FileAccess access,
        FileShare share,
        FileOptions options,
        int bufferSize)
    {
        var streamOptions = new FileStreamOptions
        {
            Mode = mode,
            Access = access,
            Share = share,
            Options = options,
            BufferSize = bufferSize,
        };
        if (!OperatingSystem.IsWindows())
            streamOptions.UnixCreateMode = PrivateFileMode;
        return new FileStream(path, streamOptions);
    }

    private static void MoveDirectoryNoReplace(string source, string destination)
    {
        int result;
        if (OperatingSystem.IsLinux())
            result = RenameAt2(AtFileDescriptorCurrentWorkingDirectory, source,
                AtFileDescriptorCurrentWorkingDirectory, destination, RenameNoReplace);
        else if (OperatingSystem.IsMacOS())
            result = RenameExclusiveMac(source, destination, RenameExclusive);
        else
        {
            // MoveFile on Windows fails atomically when the destination exists. Other supported
            // platforms retain Directory.Move's no-replace contract.
            Directory.Move(source, destination);
            return;
        }

        if (result != 0)
            throw new IOException(
                $"Could not atomically publish delivery directory '{destination}'.",
                new Win32Exception(Marshal.GetLastPInvokeError()));
    }

    [DllImport("libc", EntryPoint = "renameat2", SetLastError = true)]
    private static extern int RenameAt2(
        int oldDirectoryFileDescriptor,
        string oldPath,
        int newDirectoryFileDescriptor,
        string newPath,
        uint flags);

    [DllImport("libSystem.B.dylib", EntryPoint = "renamex_np", SetLastError = true)]
    private static extern int RenameExclusiveMac(string oldPath, string newPath, uint flags);

    [DllImport("libSystem.B.dylib", EntryPoint = "flock", SetLastError = true)]
    private static extern int LockFileMac(int fileDescriptor, int operation);

    [DllImport("libc", EntryPoint = "mkdir", SetLastError = true)]
    private static extern int CreateDirectoryUnix(string path, uint mode);

    [DllImport("kernel32.dll", EntryPoint = "CreateDirectoryW", CharSet = CharSet.Unicode,
        SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateDirectoryWindows(
        string path,
        IntPtr securityAttributes);

    private sealed class PublicationLease : IDisposable
    {
        private FileStream? _stream;
        private object? _stripe;

        internal PublicationLease(FileStream stream, object stripe)
        {
            _stream = stream;
            _stripe = stripe;
        }

        public void Dispose()
        {
            var stream = Interlocked.Exchange(ref _stream, null);
            var stripe = Interlocked.Exchange(ref _stripe, null);
            try
            {
                stream?.Dispose();
            }
            finally
            {
                if (stripe is not null)
                    Monitor.Exit(stripe);
            }
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
    BeforeDirectoryCommit,
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
