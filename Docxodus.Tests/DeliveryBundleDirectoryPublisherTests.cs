// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
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
        using var ready = new Barrier(2);
        var injector = new CallbackFaultInjector(context =>
        {
            if (context.Checkpoint == DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)
                Assert.True(ready.SignalAndWait(TimeSpan.FromSeconds(15)));
        });

        var attempts = Enumerable.Range(0, 2).Select(_ => Task.Run(() =>
        {
            try
            {
                DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector);
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

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);
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

        Assert.Contains("changed after verification", exception.Message, StringComparison.Ordinal);
        Assert.False(PathEntryExists(target));
        AssertNoStageDirectories(temporary.Path);
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
