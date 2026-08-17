// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using Docxodus.Delivery;
using Xunit;
using DeliveryCliApp = Docxodus.DeliveryCli.DeliveryCli;

namespace Docxodus.Tests;

public sealed class DeliveryCliTests : IDisposable
{
    private readonly string _root;

    public DeliveryCliTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"docxodus-delivery-cli-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
    }

    [Fact]
    public async Task RunAsync_PublishesAtomicBundleThatReopensAndVerifies()
    {
        var baselinePath = Path.Combine(_root, "baseline.docx");
        var workingPath = Path.Combine(_root, "working.docx");
        var outputDirectory = Path.Combine(_root, "delivery");
        var baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        File.WriteAllBytes(baselinePath, baseline);
        using (var session = new DocxSession(baseline, new DocxSessionSettings
               {
                   PersistAnchorIds = false,
                   EmitMarkdownPatch = false,
                   CaptureInitialProjection = false,
               }))
        {
            var anchor = session.Project().AnchorIndex.Keys.First(value =>
                value.StartsWith("p:body:", StringComparison.Ordinal));
            Assert.True(session.ReplaceText(anchor, "CLI delivery edit.").Success);
            File.WriteAllBytes(workingPath, session.Save(persistAnchorIds: false));
        }
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var exitCode = await DeliveryCliApp.RunAsync(new[]
        {
            baselinePath,
            workingPath,
            outputDirectory,
            "--baseline-version=0",
            "--working-version=1",
            "--final-version=1",
            "--final-name=final",
            "--pre-existing=preserve",
            "--generated=accept",
            "--artifact=team:final:final-docx:required",
            "--artifact=semantic:semantic-delta:required",
            "--artifact=package:package-delta:required",
            "--artifact=validation:validation-report:required",
        }, stdout, stderr);

        Assert.Equal(0, exitCode);
        Assert.Equal(string.Empty, stderr.ToString());
        using var result = JsonDocument.Parse(stdout.ToString());
        Assert.Equal("complete", result.RootElement.GetProperty("status").GetString());
        Assert.True(result.RootElement.GetProperty("verified").GetBoolean());
        Assert.Equal(Path.GetFullPath(outputDirectory),
            result.RootElement.GetProperty("outputDirectory").GetString());

        var manifestPath = Path.Combine(outputDirectory, DeliveryBundle.ManifestFileName);
        var manifestBytes = File.ReadAllBytes(manifestPath);
        using var manifest = JsonDocument.Parse(manifestBytes);
        var available = manifest.RootElement.GetProperty("payload").GetProperty("artifacts")
            .EnumerateArray()
            .Where(value => value.GetProperty("availability").GetString() == "available")
            .ToDictionary(
                value => value.GetProperty("artifactId").GetString()!,
                value => File.ReadAllBytes(Path.Combine(
                    outputDirectory,
                    value.GetProperty("relativePath").GetString()!
                        .Replace('/', Path.DirectorySeparatorChar))),
                StringComparer.Ordinal);
        var verification = DeliveryBundleVerifier.VerifyJson(manifestBytes, available);
        Assert.True(verification.IsValid,
            string.Join(Environment.NewLine, verification.Findings));
    }

    [Fact]
    public async Task RunAsync_RequiresExplicitPoliciesAndLeavesNoOutput()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var outputDirectory = Path.Combine(_root, "not-created");

        var exitCode = await DeliveryCliApp.RunAsync(new[]
        {
            "baseline.docx",
            "working.docx",
            outputDirectory,
            "--baseline-version=0",
            "--final-version=1",
            "--final-name=final",
            "--artifact=final:final-docx:required",
        }, stdout, stderr);

        Assert.Equal(2, exitCode);
        Assert.Contains("Both --pre-existing and --generated", stderr.ToString(),
            StringComparison.Ordinal);
        Assert.False(Directory.Exists(outputDirectory));
    }

    [Fact]
    public async Task RunAsync_RejectsInvalidPdfProfileBeforeReadingInputs()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var exitCode = await DeliveryCliApp.RunAsync(new[]
        {
            Path.Combine(_root, "missing-baseline.docx"),
            Path.Combine(_root, "missing-working.docx"),
            Path.Combine(_root, "not-created"),
            "--baseline-version=0",
            "--final-version=1",
            "--final-name=final",
            "--pre-existing=preserve",
            "--generated=accept",
            "--artifact=pdf:final-pdf:required:original:hidden",
        }, stdout, stderr);

        Assert.Equal(2, exitCode);
        Assert.Contains("requires the final review profile", stderr.ToString(),
            StringComparison.Ordinal);
        Assert.DoesNotContain("file not found", stderr.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RunAsync_RejectsSparseOversizeInputBeforeAllocatingIt()
    {
        var baselinePath = Path.Combine(_root, "oversize.docx");
        using (var stream = new FileStream(baselinePath, FileMode.CreateNew, FileAccess.Write))
            stream.SetLength(DeliveryArtifactRequestRules.MaximumInputPackageBytes + 1);
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();

        var exitCode = await DeliveryCliApp.RunAsync(new[]
        {
            baselinePath,
            Path.Combine(_root, "working.docx"),
            Path.Combine(_root, "not-created"),
            "--baseline-version=0",
            "--final-version=1",
            "--final-name=final",
            "--pre-existing=preserve",
            "--generated=accept",
            "--artifact=final:final-docx:required",
        }, stdout, stderr);

        Assert.Equal(1, exitCode);
        Assert.Contains("input budget", stderr.ToString(), StringComparison.Ordinal);
    }

    public void Dispose()
    {
        if (Directory.Exists(_root))
            Directory.Delete(_root, recursive: true);
    }
}
