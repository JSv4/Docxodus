// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using Docxodus.Delivery;
using Xunit;

namespace Docxodus.Tests;

public sealed class DeliveryBundleMacTempTests
{
    private static readonly byte[] Payload = Encoding.UTF8.GetBytes("publication path regression");
    private static readonly byte[] Manifest = Encoding.UTF8.GetBytes("{\"fixture\":\"filesystem transaction\"}");

    [MacTemporaryPathTheory]
    [InlineData("/tmp")]
    [InlineData("/var")]
    public void Publish_UsesPhysicalPathsThroughoutButReturnsTheRequestedPath(string alias)
    {
        using var area = new MacTemporaryArea(alias);
        var requested = Path.Combine(area.AliasPath, "bundle");
        var physical = Path.Combine(area.PhysicalPath, "bundle");
        var visited = new HashSet<DeliveryBundleDirectoryPublisherCheckpoint>();
        var observer = new PublicationObserver(context =>
        {
            Assert.Equal(physical, context.TargetDirectory);
            Assert.Equal(area.PhysicalPath, Path.GetDirectoryName(context.StageDirectory));
            visited.Add(context.Checkpoint);
        });

        Assert.Equal(requested, DeliveryBundleDirectoryPublisher.Publish(NewSource(), requested, observer));

        Assert.Equal(Enum.GetValues<DeliveryBundleDirectoryPublisherCheckpoint>().Length, visited.Count);
        AssertPublished(physical);
        Assert.Equal(new[] { physical }, Directory.GetFileSystemEntries(area.PhysicalPath));
    }

    [MacTemporaryPathTheory]
    [InlineData("/tmp", true)]
    [InlineData("/tmp", false)]
    [InlineData("/var", true)]
    [InlineData("/var", false)]
    public void Publish_BothSpellingsCompeteForOneLeaseEvenAfterTheStageMarkerIsRemoved(
        string alias, bool aliasFirst)
    {
        using var area = new MacTemporaryArea(alias);
        var aliasTarget = Path.Combine(area.AliasPath, "bundle");
        var physicalTarget = Path.Combine(area.PhysicalPath, "bundle");
        var first = aliasFirst ? aliasTarget : physicalTarget;
        var second = aliasFirst ? physicalTarget : aliasTarget;
        var attempted = false;
        var observer = new PublicationObserver(context =>
        {
            if (context.Checkpoint != DeliveryBundleDirectoryPublisherCheckpoint.BeforeDirectoryCommit)
                return;

            attempted = true;
            Assert.False(Directory.Exists(physicalTarget));
            Assert.False(File.Exists(Path.Combine(context.StageDirectory, ".docxodus-delivery-stage")));
            var failure = Assert.Throws<IOException>(() =>
                DeliveryBundleDirectoryPublisher.Publish(NewSource(), second));
            Assert.Contains("commit lease", failure.Message, StringComparison.Ordinal);
        });

        Assert.Equal(first, DeliveryBundleDirectoryPublisher.Publish(NewSource(), first, observer));

        Assert.True(attempted);
        AssertPublished(physicalTarget);
        Assert.Equal(new[] { physicalTarget }, Directory.GetFileSystemEntries(area.PhysicalPath));
    }

    [MacTemporaryPathTheory]
    [InlineData("/tmp", false)]
    [InlineData("/tmp", true)]
    [InlineData("/var", false)]
    [InlineData("/var", true)]
    public void Publish_SystemAliasDoesNotPermitAnotherLinkInTheDestinationParent(
        string alias, bool dangling)
    {
        using var area = new MacTemporaryArea(alias);
        // Prove this alias can publish before asking for rejection of the additional link.
        DeliveryBundleDirectoryPublisher.Publish(NewSource(), Path.Combine(area.AliasPath, "control"));
        var outside = Path.Combine(area.PhysicalPath, "outside");
        if (!dangling)
            Directory.CreateDirectory(outside);
        var redirect = Path.Combine(area.PhysicalPath, "redirect");
        Directory.CreateSymbolicLink(redirect, outside);

        Assert.ThrowsAny<IOException>(() => DeliveryBundleDirectoryPublisher.Publish(
            NewSource(), Path.Combine(area.AliasPath, "redirect", "bundle")));

        Assert.Equal(outside, new DirectoryInfo(redirect).LinkTarget);
        Assert.False(Directory.Exists(Path.Combine(outside, "bundle")));
        if (dangling)
            Assert.False(Directory.Exists(outside));
        else
            Assert.Empty(Directory.GetFileSystemEntries(outside));
        Assert.DoesNotContain(Directory.GetFileSystemEntries(area.PhysicalPath), entry =>
            Path.GetFileName(entry).StartsWith(".", StringComparison.Ordinal));
    }

    private static DeliveryBundleDirectoryPublicationSource NewSource() => new()
    {
        Artifacts =
        [
            new DeliveryBundleDirectoryPublicationArtifact("data/content.bin", Payload),
        ],
        ManifestRelativePath = "manifest.json",
        CanonicalManifestBytes = Manifest,
        VerifyStagedBytes = files =>
        {
            Assert.Equal(2, files.Count);
            Assert.Equal(Payload, files["data/content.bin"]);
            Assert.Equal(Manifest, files["manifest.json"]);
        },
    };

    private static void AssertPublished(string target)
    {
        Assert.Equal(Payload, File.ReadAllBytes(Path.Combine(target, "data", "content.bin")));
        Assert.Equal(Manifest, File.ReadAllBytes(Path.Combine(target, "manifest.json")));
        Assert.Equal(2, Directory.GetFiles(target, "*", SearchOption.AllDirectories).Length);
    }

    private sealed class PublicationObserver(Action<DeliveryBundleDirectoryPublisherFaultContext> observe)
        : IDeliveryBundleDirectoryPublisherFaultInjector
    {
        public void OnCheckpoint(DeliveryBundleDirectoryPublisherFaultContext context) => observe(context);
    }

    private sealed class MacTemporaryPathTheoryAttribute : TheoryAttribute
    {
        public MacTemporaryPathTheoryAttribute()
        {
            if (!OperatingSystem.IsMacOS())
                Skip = "Uses macOS system temporary-directory aliases on the real filesystem.";
        }
    }

    private sealed class MacTemporaryArea : IDisposable
    {
        internal MacTemporaryArea(string alias)
        {
            var systemDestination = "/private" + alias;
            var actualLink = new DirectoryInfo(alias).LinkTarget;
            Assert.True(actualLink == systemDestination || actualLink == systemDestination[1..],
                $"Expected the standard system alias at {alias}; found {actualLink ?? "no link"}.");
            var parent = alias == "/var" ? "/private/var/tmp" : "/private/tmp";
            PhysicalPath = Path.Combine(parent, "docxodus-mac-path-" + Guid.NewGuid().ToString("N"));
            AliasPath = PhysicalPath["/private".Length..];
            Directory.CreateDirectory(PhysicalPath);
        }

        internal string PhysicalPath { get; }
        internal string AliasPath { get; }

        public void Dispose() => Directory.Delete(PhysicalPath, recursive: true);
    }
}
