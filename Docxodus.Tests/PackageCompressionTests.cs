#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// Regression coverage for issue #331. The real multi-part fixture is intentionally used instead
/// of a generated document: its Word-authored ZIP entries carry the misleading "superfast"
/// deflate hint that causes .NET update-mode saves to lose compression.
/// </summary>
public sealed class PackageCompressionTests
{
    private const string RegressionFixture = "HC031-Complicated-Document.docx";
    private readonly ITestOutputHelper _output;

    public PackageCompressionTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void PKG331_SessionSave_MultiPartPackageDoesNotLoseCompression()
    {
        var input = LoadFixture(RegressionFixture);
        using var session = new DocxSession(input);

        var output = session.Save();
        var deltas = PartDeltas(input, output).ToList();
        WritePackageReport(input, output, deltas);

        var mainDocument = Assert.Single(deltas, x => x.Name == "word/document.xml");
        Assert.True(mainDocument.AfterUncompressed > mainDocument.BeforeUncompressed,
            "The fixture must retain the small XML-growth precondition from issue #331.");
        Assert.True(mainDocument.AfterCompressed <= mainDocument.BeforeCompressed,
            $"The main document lost compression:\n{FormatDeltas(deltas)}");
        Assert.True(TotalUncompressed(output) - TotalUncompressed(input) < 2_048,
            $"Unexpected uncompressed package growth:\n{FormatDeltas(deltas)}");
        Assert.True(output.Length <= input.Length * 0.90,
            $"Expected at least a 10% package-size reduction; {input.Length:n0} -> {output.Length:n0}." +
            $"\n{FormatDeltas(deltas)}");

        using var verifyStream = new MemoryStream(output);
        using var verified = WordprocessingDocument.Open(verifyStream, isEditable: false);
        Assert.NotNull(verified.MainDocumentPart);
    }

    [Fact]
    public void PKG332_Normalization_PreservesEveryEntryPayloadAndOpcStructure()
    {
        var input = LoadFixture(RegressionFixture);
        var output = ZipPackageOutputNormalizer.Normalize(input);
        var beforeParts = ReadParts(input);
        var afterParts = ReadParts(output);

        Assert.Equal(beforeParts.Keys.Order(), afterParts.Keys.Order());
        Assert.Contains("[Content_Types].xml", afterParts.Keys);
        Assert.Contains("_rels/.rels", afterParts.Keys);
        Assert.Contains("word/_rels/document.xml.rels", afterParts.Keys);
        foreach (var (name, payload) in beforeParts)
        {
            Assert.True(payload.AsSpan().SequenceEqual(afterParts[name]),
                $"Normalization changed the uncompressed payload of '{name}'.");
        }
    }

    [Fact]
    public void PKG333_SessionSave_KeepsAlreadyStoredMediaUncompressed()
    {
        var input = LoadFixture("DB007-Notes.docx");
        using var session = new DocxSession(input);

        var output = session.Save();
        using var archive = new ZipArchive(new MemoryStream(output), ZipArchiveMode.Read);
        var firstImage = archive.GetEntry("word/media/image1.tmp");
        var secondImage = archive.GetEntry("word/media/image2.tmp");

        Assert.NotNull(firstImage);
        Assert.NotNull(secondImage);
        Assert.Equal(firstImage.Length, firstImage.CompressedLength);
        Assert.Equal(secondImage.Length, secondImage.CompressedLength);
    }

    [Fact]
    public void PKG334_MemoryStreamDocument_OutputUsesTheSameCompressionPolicy()
    {
        var input = LoadFixture(RegressionFixture);
        using var streamDocument = new OpenXmlMemoryStreamDocument(
            new WmlDocument(RegressionFixture, input));

        using (var wordDocument = streamDocument.GetWordprocessingDocument())
        {
            wordDocument.MainDocumentPart!.Document.Save();
        }

        var output = streamDocument.GetModifiedWmlDocument().DocumentByteArray;
        var deltas = PartDeltas(input, output).ToList();
        WritePackageReport(input, output, deltas);

        Assert.True(output.Length <= input.Length * 0.90,
            $"The shared stream output path did not apply the package compression policy.\n" +
            FormatDeltas(deltas));
    }

    private static byte[] LoadFixture(string name) =>
        File.ReadAllBytes(Path.Combine("../../../../TestFiles", name));

    private static Dictionary<string, byte[]> ReadParts(byte[] packageBytes)
    {
        using var archive = new ZipArchive(new MemoryStream(packageBytes), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(
            x => x.FullName,
            x =>
            {
                using var source = x.Open();
                using var payload = new MemoryStream();
                source.CopyTo(payload);
                return payload.ToArray();
            });
    }

    private static long TotalUncompressed(byte[] packageBytes)
    {
        using var archive = new ZipArchive(new MemoryStream(packageBytes), ZipArchiveMode.Read);
        return archive.Entries.Sum(x => x.Length);
    }

    private static IEnumerable<PartDelta> PartDeltas(byte[] before, byte[] after)
    {
        using var beforeZip = new ZipArchive(new MemoryStream(before), ZipArchiveMode.Read);
        using var afterZip = new ZipArchive(new MemoryStream(after), ZipArchiveMode.Read);
        var beforeEntries = beforeZip.Entries.ToDictionary(x => x.FullName);
        var afterEntries = afterZip.Entries.ToDictionary(x => x.FullName);

        foreach (var name in beforeEntries.Keys.Union(afterEntries.Keys).Order())
        {
            beforeEntries.TryGetValue(name, out var oldEntry);
            afterEntries.TryGetValue(name, out var newEntry);
            yield return new PartDelta(
                name,
                oldEntry?.CompressedLength ?? 0,
                newEntry?.CompressedLength ?? 0,
                oldEntry?.Length ?? 0,
                newEntry?.Length ?? 0);
        }
    }

    private void WritePackageReport(byte[] before, byte[] after, IReadOnlyCollection<PartDelta> deltas)
    {
        _output.WriteLine(
            $"Package: compressed {before.Length:n0} -> {after.Length:n0} " +
            $"({after.Length - before.Length:+#,0;-#,0;0}); uncompressed " +
            $"{TotalUncompressed(before):n0} -> {TotalUncompressed(after):n0}");
        _output.WriteLine(FormatDeltas(deltas));
    }

    private static string FormatDeltas(IEnumerable<PartDelta> deltas) => string.Join(
        Environment.NewLine,
        deltas.OrderByDescending(x => Math.Abs(x.CompressedDelta)).Select(x =>
            $"{x.Name}: compressed {x.BeforeCompressed:n0} -> {x.AfterCompressed:n0} " +
            $"({x.CompressedDelta:+#,0;-#,0;0}); uncompressed " +
            $"{x.BeforeUncompressed:n0} -> {x.AfterUncompressed:n0} " +
            $"({x.UncompressedDelta:+#,0;-#,0;0})"));

    private sealed record PartDelta(
        string Name,
        long BeforeCompressed,
        long AfterCompressed,
        long BeforeUncompressed,
        long AfterUncompressed)
    {
        public long CompressedDelta => AfterCompressed - BeforeCompressed;

        public long UncompressedDelta => AfterUncompressed - BeforeUncompressed;
    }
}
