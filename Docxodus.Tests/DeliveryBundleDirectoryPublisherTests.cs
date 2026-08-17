// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Security.Cryptography;
using Docxodus.Delivery;
using Xunit;

namespace Docxodus.Tests;

public sealed class DeliveryBundleDirectoryPublisherTests
{
    private static readonly byte[] WorkingBytes = { 0, 1, 2, 3, 255 };
    private static readonly byte[] ReportBytes = Encoding.UTF8.GetBytes("{\"valid\":true}\n");
    private static readonly byte[] ManifestBytes = Encoding.UTF8.GetBytes(
        "{\"schema\":\"https://docxodus.dev/schemas/delivery/bundle/v1\"}\n");

    [Fact]
    public void Publish_WritesExactBytesAndManifestLast_ThenVerifiesFreshStage()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        IReadOnlyDictionary<string, byte[]>? verified = null;
        var trace = new List<(DeliveryBundleDirectoryPublisherCheckpoint Point, string? Path)>();
        var source = Source(files => verified = Clone(files));
        var injector = new CallbackFaultInjector(context =>
            trace.Add((context.Checkpoint, context.RelativePath)));

        var published = DeliveryBundleDirectoryPublisher.Publish(source, target, injector);

        Assert.Equal(Path.GetFullPath(target), published);
        Assert.Equal(WorkingBytes, File.ReadAllBytes(Path.Combine(target, "documents", "working.docx")));
        Assert.Equal(ReportBytes, File.ReadAllBytes(Path.Combine(target, "reports", "validation.json")));
        Assert.Equal(ManifestBytes, File.ReadAllBytes(Path.Combine(target, "manifest.json")));
        Assert.False(File.Exists(Path.Combine(target, ".docxodus-delivery-stage")));
        Assert.NotNull(verified);
        Assert.Equal(WorkingBytes, verified!["documents/working.docx"]);
        Assert.Equal(ReportBytes, verified["reports/validation.json"]);
        Assert.Equal(ManifestBytes, verified["manifest.json"]);

        Assert.Equal(new[]
        {
            (DeliveryBundleDirectoryPublisherCheckpoint.BeforeArtifactWrite,
                (string?)"documents/working.docx"),
            (DeliveryBundleDirectoryPublisherCheckpoint.BeforeArtifactWrite,
                (string?)"reports/validation.json"),
            (DeliveryBundleDirectoryPublisherCheckpoint.BeforeManifestWrite,
                (string?)"manifest.json"),
            (DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification, (string?)null),
            (DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit, (string?)null),
            (DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit, (string?)null),
        }, trace);
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_PublicBundleOverload_UsesCanonicalManifestAndDeclaredPaths()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var bundle = ModelBundle();

        DeliveryBundleDirectoryPublisher.Publish(bundle, target);

        Assert.Equal(bundle.GetArtifactBytes("working"),
            File.ReadAllBytes(Path.Combine(target, "documents", "working.docx")));
        Assert.Equal(bundle.ManifestBytes,
            File.ReadAllBytes(Path.Combine(target, DeliveryBundle.ManifestFileName)));
        Assert.True(DeliveryBundleVerifier.VerifyJson(
            File.ReadAllBytes(Path.Combine(target, DeliveryBundle.ManifestFileName)),
            new Dictionary<string, byte[]>
            {
                ["working"] = File.ReadAllBytes(
                    Path.Combine(target, "documents", "working.docx")),
            }).IsValid);
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_RejectsBundleResourceLimitsBeforeCreatingAStage()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var checkpointReached = false;
        var injector = new CallbackFaultInjector(_ => checkpointReached = true);

        Assert.Throws<InvalidDataException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(
                ModelBundle(),
                target,
                new DeliveryBundleVerificationLimits { MaxArtifactBytes = 1 },
                injector));

        Assert.False(checkpointReached);
        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);

        checkpointReached = false;
        Assert.Throws<InvalidDataException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(
                ModelBundle(),
                Path.Combine(temporary.Path, "strict-metadata"),
                new DeliveryBundleVerificationLimits { MaxStringLength = 3 },
                injector));
        Assert.False(checkpointReached);
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_PublicBundleOverload_RequiresExplicitDiagnosticPublication()
    {
        using var temporary = new TemporaryDirectory();
        var rejectedTarget = Path.Combine(temporary.Path, "rejected");
        var diagnosticTarget = Path.Combine(temporary.Path, "diagnostic");
        var bundle = DiagnosticModelBundle();

        var exception = Assert.Throws<DeliveryBundleException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(bundle, rejectedTarget));

        Assert.Equal("diagnostic_bundle_publication_not_enabled", exception.Code);
        Assert.False(PathEntryExists(rejectedTarget));

        DeliveryBundleDirectoryPublisher.Publish(
            bundle,
            diagnosticTarget,
            publicationOptions: new DeliveryBundleDirectoryPublicationOptions
            {
                AllowDiagnosticBundle = true,
            });

        var verification = DeliveryBundleVerifier.VerifyJson(
            File.ReadAllBytes(Path.Combine(diagnosticTarget, DeliveryBundle.ManifestFileName)));
        Assert.True(verification.IsValid,
            string.Join(Environment.NewLine, verification.Findings));
        AssertNoStageDirectories(temporary.Path);
    }

    [Theory]
    [InlineData("documents/working.docx")]
    [InlineData("reports/validation.json")]
    [InlineData("manifest.json")]
    public void Publish_WhenAnyWriteFaults_CleansOwnedStageAndLeavesNoTarget(string path)
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.RelativePath == path)
                throw new InjectedPublicationException(path);
        });

        Assert.Throws<InjectedPublicationException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_WhenVerificationFails_CleansOwnedStageAndLeavesNoTarget()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var source = Source(_ => throw new InjectedPublicationException("verify"));

        Assert.Throws<InjectedPublicationException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(source, target));

        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);
    }

    [Theory]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification)]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)]
    public void Publish_WhenPreCommitCheckpointFaults_CleansStageAndLeavesNoTarget(
        int checkpointValue)
    {
        var checkpoint = (DeliveryBundleDirectoryPublisherCheckpoint)checkpointValue;
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint == checkpoint)
                throw new InjectedPublicationException(checkpoint.ToString());
        });

        Assert.Throws<InjectedPublicationException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_WhenFinalCommitCheckpointFaults_CleansAuthenticatedStage()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        string? stage = null;
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint !=
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)
                return;
            stage = context.StageDirectory;
            Assert.True(File.Exists(Path.Combine(
                context.StageDirectory, ".docxodus-delivery-stage")));
            throw new InjectedPublicationException("final commit checkpoint");
        });

        Assert.Throws<InjectedPublicationException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.NotNull(stage);
        Assert.False(PathEntryExists(stage!));
        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_CopiedMarkerCannotAuthorizeDeletingAReplacementDirectory()
    {
        if (OperatingSystem.IsWindows())
            return; // Windows correctly prevents moving a directory containing an open marker.

        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        string? replacement = null;
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint != DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)
                return;
            var markerPath = Path.Combine(
                context.StageDirectory, ".docxodus-delivery-stage");
            var copiedMarker = File.ReadAllBytes(markerPath);
            Directory.Move(context.StageDirectory, context.StageDirectory + "-displaced");
            Directory.CreateDirectory(context.StageDirectory);
            File.WriteAllBytes(markerPath, copiedMarker);
            replacement = Path.Combine(context.StageDirectory, "foreign.txt");
            File.WriteAllText(replacement, "belongs to another owner");
            throw new InjectedPublicationException("foreign replacement");
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.NotNull(replacement);
        Assert.Equal("belongs to another owner", File.ReadAllText(replacement!));
        Assert.False(PathEntryExists(target));
    }

    [Fact]
    public void Publish_RejectsUnexpectedDirectoryAtFinalCheckpoint()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        string? stage = null;
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint !=
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)
                return;
            stage = context.StageDirectory;
            Directory.CreateDirectory(Path.Combine(context.StageDirectory, "unexpected-empty"));
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.NotNull(stage);
        Assert.True(Directory.Exists(Path.Combine(stage!, "unexpected-empty")));
        Assert.False(PathEntryExists(target));
    }

    [Fact]
    public void Publish_UnlockedPersistentLeaseFromPriorFailureDoesNotBlockRetry()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint == DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification)
                throw new InjectedPublicationException("leave persistent lease");
        });

        Assert.Throws<InjectedPublicationException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        DeliveryBundleDirectoryPublisher.Publish(Source(), target);
        Assert.True(Directory.Exists(target));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_DoesNotMutatePreExistingLeaseFile()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var digest = Convert.ToHexString(SHA256.HashData(
                Encoding.UTF8.GetBytes(Path.GetFullPath(target))))
            .ToLowerInvariant()[..24];
        var lease = Path.Combine(
            temporary.Path, $".docxodus-delivery-publish-{digest}.lock");
        File.WriteAllText(lease, "must remain unchanged");

        DeliveryBundleDirectoryPublisher.Publish(Source(), target);

        Assert.Equal("must remain unchanged", File.ReadAllText(lease));
    }

    [Fact]
    public void Publish_CreatesPrivateStageDirectoriesAndFilesOnUnix()
    {
        if (OperatingSystem.IsWindows())
            return;

        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");

        DeliveryBundleDirectoryPublisher.Publish(Source(), target);

        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
            File.GetUnixFileMode(target));
        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite,
            File.GetUnixFileMode(Path.Combine(target, "documents", "working.docx")));
    }

    [Fact]
    public void Publish_WhenOwnedStageIsReplaced_ReportsCleanupFailureAndPreservesForeignPath()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        string? replacedStage = null;
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint !=
                DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)
                return;
            replacedStage = context.StageDirectory;
            Directory.Delete(context.StageDirectory, recursive: true);
            File.WriteAllText(context.StageDirectory, "foreign replacement");
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.IsType<AggregateException>(exception.InnerException);
        Assert.NotNull(replacedStage);
        Assert.Equal("foreign replacement", File.ReadAllText(replacedStage!));
        Assert.False(PathEntryExists(target));
    }

    [Fact]
    public void Publish_ExistingTargetIsUntouched()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        Directory.CreateDirectory(target);
        var sentinel = Path.Combine(target, "belongs-to-caller.txt");
        File.WriteAllText(sentinel, "keep me");

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target));

        Assert.Equal("keep me", File.ReadAllText(sentinel));
        Assert.Single(Directory.EnumerateFileSystemEntries(target));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public async Task Publish_ConcurrentCreate_HasExactlyOneWinnerAndNoStageLeak()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");

        var attempts = Enumerable.Range(0, 2).Select(_ => Task.Run(() =>
        {
            try
            {
                DeliveryBundleDirectoryPublisher.Publish(Source(), target);
                return (Exception?)null;
            }
            catch (Exception exception)
            {
                return exception;
            }
        })).ToArray();

        var results = await Task.WhenAll(attempts);

        Assert.Single(results, result => result is null);
        Assert.Single(results, result => result is IOException);
        Assert.Equal(WorkingBytes,
            File.ReadAllBytes(Path.Combine(target, "documents", "working.docx")));
        AssertNoStageDirectories(temporary.Path);
    }

    [Theory]
    [InlineData("../escape.bin")]
    [InlineData("a/../../escape.bin")]
    [InlineData("/rooted.bin")]
    [InlineData("C:/drive-rooted.bin")]
    [InlineData("a\\windows-separator.bin")]
    [InlineData("a/./current.bin")]
    [InlineData("a//empty.bin")]
    [InlineData("a/trailing/")]
    [InlineData("a/NUL.txt")]
    [InlineData("a/COM¹.txt")]
    [InlineData("a/LPT².log")]
    [InlineData("a/CONIN$")]
    [InlineData("a/CONOUT$.txt")]
    public void Publish_RejectsTraversalRootAndNonCanonicalArtifactPaths(string maliciousPath)
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var source = Source(artifacts:
        [
            new DeliveryBundleDirectoryPublicationArtifact(
                maliciousPath, new byte[] { 1 }),
        ]);

        Assert.Throws<ArgumentException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(source, target));

        Assert.False(PathEntryExists(target));
        Assert.False(File.Exists(Path.Combine(temporary.Path, "escape.bin")));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_RejectsDuplicateCaseCollisionAndFileDirectoryCollision()
    {
        using var temporary = new TemporaryDirectory();

        AssertInvalidPaths(temporary.Path, "same.bin", "same.bin");
        AssertInvalidPaths(temporary.Path, "Case.bin", "case.bin");
        AssertInvalidPaths(temporary.Path, "parent", "parent/child.bin");
    }

    [Fact]
    public void Publish_RejectsArtifactCollisionWithManifest()
    {
        using var temporary = new TemporaryDirectory();
        var source = Source(artifacts:
        [
            new DeliveryBundleDirectoryPublicationArtifact(
                "MANIFEST.JSON", new byte[] { 1 }),
        ]);

        Assert.Throws<ArgumentException>(() => DeliveryBundleDirectoryPublisher.Publish(
            source, Path.Combine(temporary.Path, "delivery")));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_RequiresFullyQualifiedNewTargetWithExistingParent()
    {
        using var temporary = new TemporaryDirectory();

        Assert.Throws<ArgumentException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), "relative-delivery"));
        Assert.Throws<ArgumentException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), "  "));
        Assert.Throws<DirectoryNotFoundException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(),
                Path.Combine(temporary.Path, "missing", "delivery")));
    }

    [Fact]
    public void Publish_RejectsDanglingTargetSymlinkWithoutTouchingItsDestination()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var outside = Path.Combine(temporary.Path, "outside");
        File.CreateSymbolicLink(target, outside);

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target));

        Assert.False(PathEntryExists(outside));
        Assert.NotNull(new FileInfo(target).LinkTarget);
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_RejectsSymlinkedTargetParent()
    {
        using var temporary = new TemporaryDirectory();
        var outside = Path.Combine(temporary.Path, "outside");
        Directory.CreateDirectory(outside);
        var linkedParent = Path.Combine(temporary.Path, "linked-parent");
        Directory.CreateSymbolicLink(linkedParent, outside);
        var target = Path.Combine(linkedParent, "delivery");

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target));

        Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
        AssertNoStageDirectories(temporary.Path);
    }

    [Fact]
    public void Publish_RejectsSymlinkInjectedIntoOwnedStageAndDoesNotWriteOutside()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var outside = Path.Combine(temporary.Path, "outside");
        Directory.CreateDirectory(outside);
        var injected = false;
        var injector = new CallbackFaultInjector(context =>
        {
            if (injected
                || context.Checkpoint != DeliveryBundleDirectoryPublisherCheckpoint.BeforeArtifactWrite
                || context.RelativePath != "documents/working.docx")
                return;

            Directory.CreateSymbolicLink(
                Path.Combine(context.StageDirectory, "documents"), outside);
            injected = true;
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
        Assert.False(PathEntryExists(target));
    }

    [Fact]
    public void Publish_DetectsMutationAfterVerificationBeforeRename()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint == DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)
                File.WriteAllBytes(
                    Path.Combine(context.StageDirectory, "documents", "working.docx"),
                    new byte[] { 99 });
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.Contains("changed after verification", exception.ToString(),
            StringComparison.Ordinal);
        Assert.False(PathEntryExists(target));
    }

    [Fact]
    public void Publish_RejectsExpandedStageFileBeforeAllocatingItsReplacementSize()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint != DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification)
                return;
            using var stream = new FileStream(
                Path.Combine(context.StageDirectory, "documents", "working.docx"),
                FileMode.Append,
                FileAccess.Write,
                FileShare.None);
            stream.Write(new byte[] { 9, 9, 9 });
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.Contains("has length", exception.ToString(), StringComparison.Ordinal);
        Assert.Contains("expected", exception.ToString(), StringComparison.Ordinal);
        Assert.False(PathEntryExists(target));
    }

    [Fact]
    public void Publish_TargetCreatedBeforeRenameIsUntouched()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint != DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)
                return;
            Directory.CreateDirectory(context.TargetDirectory);
            File.WriteAllText(Path.Combine(context.TargetDirectory, "winner.txt"), "other publisher");
        });

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Equal("other publisher", File.ReadAllText(Path.Combine(target, "winner.txt")));
        AssertNoStageDirectories(temporary.Path);
    }

    private static DeliveryBundleDirectoryPublicationSource Source(
        Action<IReadOnlyDictionary<string, byte[]>>? verifier = null,
        IReadOnlyList<DeliveryBundleDirectoryPublicationArtifact>? artifacts = null) => new()
    {
        Artifacts = artifacts ??
        [
            new DeliveryBundleDirectoryPublicationArtifact(
                "documents/working.docx", WorkingBytes),
            new DeliveryBundleDirectoryPublicationArtifact(
                "reports/validation.json", ReportBytes),
        ],
        ManifestRelativePath = "manifest.json",
        CanonicalManifestBytes = ManifestBytes,
        VerifyStagedBytes = verifier ?? (_ => { }),
    };

    private static DeliveryBundle ModelBundle()
    {
        var request = new DeliveryBundleRequest(
            new DeliveryDocumentSnapshot("baseline", 1, new byte[] { 1 }),
            new DeliveryDocumentSnapshot("working", 2, new byte[] { 2 }),
            new DeliveryDocumentSnapshot("final", 3, new byte[] { 3 }),
            new DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = DeliveryRevisionPolicy.Preserve,
                GeneratedRevisions = DeliveryRevisionPolicy.Accept,
            },
            new[]
            {
                new DeliveryArtifactRequest
                {
                    ArtifactId = "working",
                    Kind = DeliveryArtifactKind.WorkingDocx,
                    Requiredness = DeliveryArtifactRequiredness.Required,
                },
            });
        return DeliveryBundle.Create(request, new[]
        {
            DeliveryBundleArtifactInput.Available(
                "working",
                DeliveryArtifactKind.WorkingDocx,
                "documents/working.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                WorkingBytes),
        });
    }

    private static DeliveryBundle DiagnosticModelBundle()
    {
        var request = new DeliveryBundleRequest(
            new DeliveryDocumentSnapshot("baseline", 1, new byte[] { 1 }),
            new DeliveryDocumentSnapshot("working", 2, new byte[] { 2 }),
            new DeliveryDocumentSnapshot("final", 3, new byte[] { 3 }),
            new DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = DeliveryRevisionPolicy.Preserve,
                GeneratedRevisions = DeliveryRevisionPolicy.Accept,
            },
            new[]
            {
                new DeliveryArtifactRequest
                {
                    ArtifactId = "working",
                    Kind = DeliveryArtifactKind.WorkingDocx,
                    Requiredness = DeliveryArtifactRequiredness.Required,
                },
            });
        return DeliveryBundle.Create(request, new[]
        {
            DeliveryBundleArtifactInput.Unavailable(
                "working",
                DeliveryArtifactKind.WorkingDocx,
                "documents/working.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "Intentionally unavailable for diagnostic publication coverage."),
        });
    }

    private static IReadOnlyDictionary<string, byte[]> Clone(
        IReadOnlyDictionary<string, byte[]> files) => files.ToDictionary(
            item => item.Key,
            item => item.Value.ToArray(),
            StringComparer.Ordinal);

    private static void AssertInvalidPaths(string root, params string[] paths)
    {
        var target = Path.Combine(root, $"delivery-{Guid.NewGuid():N}");
        var source = Source(artifacts: paths.Select(path =>
            new DeliveryBundleDirectoryPublicationArtifact(path, new byte[] { 1 })).ToArray());

        Assert.Throws<ArgumentException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(source, target));
        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(root);
    }

    private static void AssertNoStageDirectories(string parent)
    {
        Assert.DoesNotContain(Directory.EnumerateFileSystemEntries(parent), path =>
            Path.GetFileName(path).Contains(".docxodus-stage-", StringComparison.Ordinal));
    }

    private static bool PathEntryExists(string path)
    {
        if (File.Exists(path) || Directory.Exists(path))
            return true;
        return new FileInfo(path).LinkTarget is not null
            || new DirectoryInfo(path).LinkTarget is not null;
    }

    private sealed class CallbackFaultInjector : IDeliveryBundleDirectoryPublisherFaultInjector
    {
        private readonly Action<DeliveryBundleDirectoryPublisherFaultContext> _callback;

        internal CallbackFaultInjector(
            Action<DeliveryBundleDirectoryPublisherFaultContext> callback) =>
            _callback = callback;

        public void OnCheckpoint(DeliveryBundleDirectoryPublisherFaultContext context) =>
            _callback(context);
    }

    private sealed class InjectedPublicationException : Exception
    {
        internal InjectedPublicationException(string point)
            : base($"Injected failure at {point}.")
        {
        }
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        internal TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                $"docxodus-directory-publisher-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
    }
}
