// Copyright (c) John Scrudato IV. All rights reserved.
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

    [MacOSFact]
    public void Publish_MacOSTempAlias_PreservesLexicalPathAndPublishesVerifiedBundle()
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        Assert.StartsWith("/var/", target, StringComparison.Ordinal);
        var canonicalTarget = "/private" + target;
        Assert.NotEqual(Path.GetFullPath(target), Path.GetFullPath(canonicalTarget));
        var bundle = ModelBundle();

        var failure = Record.Exception(() =>
        {
            var published = DeliveryBundleDirectoryPublisher.Publish(bundle, target);
            Assert.Equal(Path.GetFullPath(target), published);
        });

        Assert.Null(failure);
        Assert.Equal(bundle.GetArtifactBytes("working"),
            File.ReadAllBytes(Path.Combine(canonicalTarget, "documents", "working.docx")));
        var manifest = File.ReadAllBytes(Path.Combine(target, DeliveryBundle.ManifestFileName));
        Assert.Equal(bundle.ManifestBytes, manifest);
        Assert.True(DeliveryBundleVerifier.VerifyJson(manifest, new Dictionary<string, byte[]>
        {
            ["working"] = File.ReadAllBytes(Path.Combine(target, "documents", "working.docx")),
        }).IsValid);
        Assert.Equal(new[] { target }, Directory.GetFileSystemEntries(temporary.Path));
        AssertNoStageDirectories(temporary.Path);
    }

    [UnixTheory]
    [InlineData("/var", "/private/var")]
    [InlineData("/var", "private/var")]
    [InlineData("/tmp", "/private/tmp")]
    [InlineData("/tmp", "private/tmp")]
    public void EnsureNoSymlinks_AcceptsExactMacOSTempAliases(string alias, string linkTarget)
    {
        var tree = new PolicyTree();
        tree.Entries[alias] = (FileAttributes.Directory | FileAttributes.ReparsePoint, linkTarget);

        var failure = Record.Exception(() => tree.Validate(alias + "/delivery", isMacOS: true));

        Assert.Null(failure);
        Assert.Contains("/private", tree.Reads);
        Assert.Contains("/private" + alias, tree.Reads);
        Assert.Contains("/private" + alias + "/delivery", tree.Reads);
    }

    [UnixTheory]
    [InlineData("arbitrary-root-alias")]
    [InlineData("etc-alias")]
    [InlineData("wrong-target")]
    [InlineData("relative-target-elsewhere")]
    [InlineData("relative-target-hidden-link")]
    [InlineData("chained-link")]
    [InlineData("target-parent-link")]
    [InlineData("nested-link")]
    [InlineData("dangling-target")]
    [InlineData("missing-target-parent")]
    [InlineData("inaccessible-target-parent")]
    [InlineData("unexpected-reparse-point")]
    [InlineData("target-is-file")]
    [InlineData("non-macos-var")]
    [InlineData("non-macos-tmp")]
    [InlineData("stage-boundary")]
    public void EnsureNoSymlinks_AcceptsSafeAliasButRefusesUnsafePaths(string scenario)
    {
        var tree = new PolicyTree();
        // A blanket rejection must not satisfy the security matrix.
        Assert.Null(Record.Exception(() => tree.Validate("/var/delivery", isMacOS: true)));
        tree.Reads.Clear();
        var path = "/var/delivery";
        var isMacOS = true;
        string? stopAt = null;
        switch (scenario)
        {
            case "arbitrary-root-alias":
            case "etc-alias":
                path = scenario == "etc-alias" ? "/etc/delivery" : "/alias/delivery";
                tree.Entries[Path.GetDirectoryName(path)!] =
                    (FileAttributes.Directory | FileAttributes.ReparsePoint, "/private/var");
                break;
            case "wrong-target":
            case "relative-target-elsewhere":
                tree.Entries["/var"] = (FileAttributes.Directory | FileAttributes.ReparsePoint,
                    scenario == "wrong-target" ? "/outside" : "../outside");
                break;
            case "relative-target-hidden-link":
                tree.Entries["/var"] =
                    (FileAttributes.Directory | FileAttributes.ReparsePoint, "private/redirect/../var");
                tree.Entries["/private/redirect"] =
                    (FileAttributes.Directory | FileAttributes.ReparsePoint, "/outside");
                break;
            case "chained-link":
                tree.Entries["/private/var"] =
                    (FileAttributes.Directory | FileAttributes.ReparsePoint, "/outside");
                break;
            case "target-parent-link":
                tree.Entries["/private"] =
                    (FileAttributes.Directory | FileAttributes.ReparsePoint, "/outside");
                break;
            case "nested-link":
                tree.Entries["/private/var/delivery"] =
                    (FileAttributes.Directory | FileAttributes.ReparsePoint, "/outside");
                break;
            case "dangling-target":
                tree.Entries.Remove("/private/var");
                break;
            case "missing-target-parent":
                tree.Entries.Remove("/private");
                break;
            case "inaccessible-target-parent":
                tree.Inaccessible = "/private";
                break;
            case "unexpected-reparse-point":
                tree.Entries["/var"] = (FileAttributes.Directory | FileAttributes.ReparsePoint, null);
                break;
            case "target-is-file":
                tree.Entries["/private/var"] = (FileAttributes.Normal, null);
                break;
            case "non-macos-var":
            case "non-macos-tmp":
                isMacOS = false;
                path = scenario == "non-macos-var" ? "/var/delivery" : "/tmp/delivery";
                break;
            case "stage-boundary":
                stopAt = "/";
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(scenario));
        }

        var rejection = Record.Exception(() => tree.Validate(path, isMacOS, stopAt));
        Assert.True(rejection is IOException or UnauthorizedAccessException, rejection?.ToString());
        Assert.DoesNotContain("/outside", tree.Reads);
    }

    [Theory]
    [InlineData("/tmp/delivery", "/private/tmp")]
    [InlineData("/var/delivery", "/private/var")]
    public void EnsureNoSymlinks_AliasTargetSwappedForALinkBetweenReadsIsRefused(string path, string resolved)
    {
        var tree = new PolicyTree();
        tree.Entries["/outside"] = (FileAttributes.Directory, null);
        tree.Entries["/outside/delivery"] = (FileAttributes.Directory, null);
        tree.AfterFirstRead[resolved] = (FileAttributes.Directory | FileAttributes.ReparsePoint, "/outside");

        Assert.Throws<IOException>(() => tree.Validate(path, isMacOS: true));
    }

    [UnixTheory]
    [InlineData(true)]
    [InlineData(false)]
    public void EnsureNoSymlinks_OrdinaryPathsAndLeafBoundaryNeedNoAlias(bool isMacOS)
    {
        var tree = new PolicyTree();
        tree.Validate("/private/var/delivery", isMacOS);
        tree.Validate("/private/var/missing", isMacOS, includeLeaf: false);
        Assert.Throws<IOException>(() => tree.Validate("/outside", isMacOS, "/private"));
    }

    [Theory]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification)]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)]
    public void Publish_TempAliasRejectsLinkToEmptyDirectoryAtEveryCommitCheckpoint(int checkpointValue)
    {
        // An empty link target adds no unexpected file, so only the symlink rejection can stop it.
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var outside = Path.Combine(temporary.Path, "outside");
        Directory.CreateDirectory(outside);
        var injector = new CallbackFaultInjector(context =>
        {
            if ((int)context.Checkpoint == checkpointValue)
                Directory.CreateSymbolicLink(Path.Combine(context.StageDirectory, "injected"), outside);
        });

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.False(PathEntryExists(target));
        Assert.Empty(Directory.GetFileSystemEntries(outside));
    }

    [Theory]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeVerification)]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeCommit)]
    [InlineData((int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)]
    public void Publish_TempAliasRejectsStageLinkAtEveryCommitCheckpoint(int checkpointValue)
    {
        using var temporary = new TemporaryDirectory();
        var target = Path.Combine(temporary.Path, "delivery");
        var outside = Path.Combine(temporary.Path, "outside");
        Directory.CreateDirectory(outside);
        var sentinel = Path.Combine(outside, "keep.txt");
        File.WriteAllText(sentinel, "attacker destination");
        string? stage = null;
        var injector = new CallbackFaultInjector(context =>
        {
            if ((int)context.Checkpoint != checkpointValue)
                return;
            stage = context.StageDirectory;
            Directory.CreateSymbolicLink(Path.Combine(stage, "injected"), outside);
        });

        Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.NotNull(stage);
        Assert.False(PathEntryExists(target));
        Assert.Equal(new[] { sentinel }, Directory.GetFileSystemEntries(outside));
        Assert.Equal("attacker destination", File.ReadAllText(sentinel));
        if (checkpointValue == (int)DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)
        {
            Assert.True(Directory.Exists(stage));
            Assert.NotNull(new DirectoryInfo(Path.Combine(stage!, "injected")).LinkTarget);
        }
        else
        {
            AssertNoStageDirectories(temporary.Path);
        }
    }

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
    public void Publish_WhenMarkerFreeCommitWindowFaults_PreservesAmbiguousStage()
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
            throw new InjectedPublicationException("marker-free commit window");
        });

        var exception = Assert.Throws<IOException>(() =>
            DeliveryBundleDirectoryPublisher.Publish(Source(), target, injector));

        Assert.Contains("could not be removed safely", exception.Message,
            StringComparison.Ordinal);
        Assert.IsType<AggregateException>(exception.InnerException);
        Assert.NotNull(stage);
        Assert.True(Directory.Exists(stage));
        Assert.False(File.Exists(Path.Combine(stage!, ".docxodus-delivery-stage")));
        Assert.False(PathEntryExists(target));
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
        Assert.DoesNotContain(Directory.EnumerateFileSystemEntries(parent), path =>
            Path.GetFileName(path).StartsWith(
                ".docxodus-delivery-publish-", StringComparison.Ordinal));
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

    private sealed class PolicyTree
    {
        internal Dictionary<string, (FileAttributes Attributes, string? LinkTarget)> Entries { get; } = new()
        {
            ["/var"] = (FileAttributes.Directory | FileAttributes.ReparsePoint, "private/var"),
            ["/tmp"] = (FileAttributes.Directory | FileAttributes.ReparsePoint, "private/tmp"),
            ["/private"] = (FileAttributes.Directory, null),
            ["/private/var"] = (FileAttributes.Directory, null),
            ["/private/tmp"] = (FileAttributes.Directory, null),
            ["/private/var/delivery"] = (FileAttributes.Directory, null),
            ["/private/tmp/delivery"] = (FileAttributes.Directory, null),
        };

        internal List<string> Reads { get; } = new();
        internal string? Inaccessible { get; set; }

        // An entry swapped in after the first read of its path: a link raced in between two reads.
        internal Dictionary<string, (FileAttributes Attributes, string? LinkTarget)> AfterFirstRead { get; } = new();

        internal void Validate(string path, bool isMacOS, string? stopAt = null, bool includeLeaf = true) =>
            DeliveryBundleDirectoryPublisher.EnsureNoSymlinks(path, includeLeaf, stopAt, isMacOS, Read);

        private (FileAttributes Attributes, string? LinkTarget) Read(string path)
        {
            Reads.Add(path);
            if (path == Inaccessible)
                throw new UnauthorizedAccessException(path);
            if (Reads.Count(read => read == path) > 1 && AfterFirstRead.TryGetValue(path, out var swapped))
                return swapped;
            return Entries.TryGetValue(path, out var entry)
                ? entry
                : throw new DirectoryNotFoundException(path);
        }
    }

    private sealed class MacOSFactAttribute : FactAttribute
    {
        public MacOSFactAttribute()
        {
            if (!OperatingSystem.IsMacOS())
                Skip = "Requires the real macOS /var temporary-directory alias.";
        }
    }

    private sealed class UnixTheoryAttribute : TheoryAttribute
    {
        public UnixTheoryAttribute()
        {
            if (OperatingSystem.IsWindows())
                Skip = "Exercises Unix system-alias paths; runs on macOS and Linux.";
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
