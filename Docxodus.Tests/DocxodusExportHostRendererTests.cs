// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Delivery;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxodusExportHostRendererTests
{
    [Fact]
    public async Task DEH001_RequestFrameLimitIsCheckedBeforeBase64Materialization()
    {
        var renderer = Renderer(maximumFrameBytes: 64);
        var request = Request("html");

        var exception = await Assert.ThrowsAsync<InvalidDataException>(async () =>
            await renderer.RenderBatchAsync(new[] { request }));

        Assert.Contains("frame limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DEH002_BatchRequestCountIsBoundedBeforeSnapshotOrProcessLaunch()
    {
        var renderer = Renderer(maximumFrameBytes: 1024 * 1024);
        var request = Request("same-instance-is-never-enumerated");
        var requests = Enumerable.Repeat(request, 1_025).ToArray();

        var exception = await Assert.ThrowsAsync<InvalidDataException>(async () =>
            await renderer.RenderBatchAsync(requests));

        Assert.Contains("batch-request limit", exception.Message, StringComparison.Ordinal);
    }

    private static DocxodusExportHostRenderer Renderer(int maximumFrameBytes)
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("The current process path is unavailable.");
        return new DocxodusExportHostRenderer(new DocxodusExportHostRendererOptions
        {
            NodeExecutablePath = processPath,
            HostScriptPath = typeof(DocxodusExportHostRendererTests).Assembly.Location,
            MaximumFrameBytes = maximumFrameBytes,
        });
    }

    private static DeliveryRenderRequest Request(string artifactId) => new(
        artifactId,
        DeliveryArtifactKind.StandaloneHtml,
        DeliveryReviewProfile.Final,
        DeliveryCommentProfile.Hidden,
        new DeliveryDocumentSnapshot(
            "final", 1, DocxSessionTests.BuildDS001_SimpleTwoParagraphs()));
}
