// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Text;
using Docxodus.History;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class PackageChangeSetTests
{
    [Fact]
    public void ReplayAndInvertPreserveEveryEntryIncludingRelationshipsAndOpaquePayloads()
    {
        var before = Entries(DocxSession.CreateBlankDocxBytes());
        before["customXml/old.bin"] = new byte[] { 0, 1, 2, 255 };
        DeclareBinary(before);
        var after = before.ToDictionary(pair => pair.Key, pair => pair.Value);
        after.Remove("customXml/old.bin");
        after["customXml/new.bin"] = new byte[] { 255, 0, 9, 1 };
        after["customXml/opaque.xml"] = Utf8("<?xml version=\"1.0\"?><opaque>  <x/> \n</opaque>");
        after["word/_rels/document.xml.rels"] = Utf8("""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="opaque" Type="urn:test:opaque" Target="../customXml/new.bin"/>
              <Relationship Id="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test" TargetMode="External"/>
            </Relationships>
            """);
        after["[Content_Types].xml"] = Utf8(Encoding.UTF8.GetString(after["[Content_Types].xml"])
            .Replace("</Types>", "<Override PartName=\"/customXml/opaque.xml\" ContentType=\"application/xml\"/></Types>"));
        var beforeBytes = Package(before);
        var afterBytes = Package(after);
        var change = PackageChangeSet.Create(beforeBytes, afterBytes);

        Assert.Contains(change.Changes, e => e.Uri == "/customXml/old.bin" && e.AfterDigest is null);
        Assert.Contains(change.Changes, e => e.Uri == "/customXml/new.bin" && e.BeforeDigest is null);
        Assert.Contains(change.Changes, e => e.Uri == "/[Content_Types].xml");
        Assert.Contains(change.Changes, e => e.Uri == "/word/_rels/document.xml.rels");
        AssertEntries(after, change.Apply(beforeBytes));
        AssertEntries(before, change.Invert().Apply(afterBytes));
        Assert.Equal(change.BeforeDigest, change.Invert().AfterDigest);
        AssertEntries(after, change.Invert().Invert().Apply(beforeBytes));
        AssertEntries(before, beforeBytes);
        AssertEntries(after, afterBytes);
    }

    [Fact]
    public void RepackedBaseIsAcceptedAndContainerOnlyChangesAreAnExactNoOp()
    {
        var before = Entries(DocxSession.CreateBlankDocxBytes());
        var after = before.ToDictionary(pair => pair.Key, pair => pair.Value);
        after["word/document.xml"] = Utf8(Encoding.UTF8.GetString(after["word/document.xml"])
            .Replace("<w:p", "\n<w:p", StringComparison.Ordinal));
        var original = Package(before);
        var repacked = Package(before.Reverse(), CompressionLevel.Optimal, 2025);
        Assert.NotEqual(original, repacked);

        var delta = PackageChangeSet.Create(original, Package(after));
        AssertEntries(after, delta.Apply(repacked));
        var noOp = PackageChangeSet.Create(original, repacked);
        Assert.Empty(noOp.Changes);
        Assert.Empty(noOp.PayloadDigests);
        Assert.Equal(0, noOp.RetainedPayloadBytes);
        var output = noOp.Apply(repacked);
        Assert.Equal(repacked, output);
        Assert.NotSame(repacked, output);
    }

    [Fact]
    public void ChangedPayloadsAreDeduplicatedAndCannotBeMutatedByCallers()
    {
        var before = Entries(DocxSession.CreateBlankDocxBytes());
        DeclareBinary(before);
        before["assets/large.bin"] = Enumerable.Range(0, 100_000).Select(i => (byte)i).ToArray();
        var after = before.ToDictionary(pair => pair.Key, pair => pair.Value);
        after["assets/a.bin"] = new byte[] { 1, 2, 3 };
        after["assets/b.bin"] = new byte[] { 1, 2, 3 };
        var beforeBytes = Package(before);
        var afterBytes = Package(after);
        var expected = afterBytes.ToArray();
        var delta = PackageChangeSet.Create(beforeBytes, afterBytes);

        Assert.Equal(2, delta.Changes.Count);
        var digest = Assert.Single(delta.PayloadDigests);
        Assert.Equal(3, delta.RetainedPayloadBytes);
        Assert.DoesNotContain(delta.Changes, c => c.Uri == "/assets/large.bin");
        delta.GetPayload(digest)[0] = 255;
        Array.Fill(afterBytes, (byte)0);
        Assert.Equal(new byte[] { 1, 2, 3 }, delta.GetPayload(digest));
        AssertEntries(Entries(expected), delta.Apply(beforeBytes));
        Assert.Throws<NotSupportedException>(() =>
            ((IList<PackageEntryChange>)delta.Changes).Clear());
    }

    [Fact]
    public void EmptyEntryIsDistinctFromMissingEntryAndUnicodeNamesRoundTrip()
    {
        var before = Entries(DocxSession.CreateBlankDocxBytes());
        DeclareBinary(before);
        var after = before.ToDictionary(pair => pair.Key, pair => pair.Value);
        after["assets/%C3%A9.bin"] = Array.Empty<byte>();
        var source = Package(before);
        var target = Package(after);
        var delta = PackageChangeSet.Create(source, target);
        var entry = Assert.Single(delta.Changes);

        Assert.Null(entry.BeforeDigest);
        Assert.NotNull(entry.AfterDigest);
        Assert.Equal("assets/%C3%A9.bin", entry.AfterEntryName);
        Assert.Empty(delta.GetPayload(entry.AfterDigest));
        AssertEntries(after, delta.Apply(source));
        AssertEntries(before, delta.Invert().Apply(target));
    }

    [Fact]
    public void StaleBaseIsRejectedEvenWhenOnlyAnUnrelatedEntryChanged()
    {
        var before = Entries(DocxSession.CreateBlankDocxBytes());
        DeclareBinary(before);
        var after = before.ToDictionary(pair => pair.Key, pair => pair.Value);
        after["assets/a.bin"] = new byte[] { 1 };
        var delta = PackageChangeSet.Create(Package(before), Package(after));
        before["assets/unrelated.bin"] = new byte[] { 2 };
        var stale = Package(before);
        var original = stale.ToArray();

        var error = Assert.Throws<PackageChangeException>(() => delta.Apply(stale));
        Assert.Equal(PackageChangeError.BaseMismatch, error.Code);
        Assert.Equal(original, stale);
    }

    [Theory]
    [InlineData(2d)]
    [InlineData(1_000d)]
    public void ReplayRemainsReadableWithinTheSameExpansionLimits(double maximumRatio)
    {
        var before = Entries(DocxSession.CreateBlankDocxBytes());
        DeclareBinary(before);
        before["assets/unchanged.bin"] = new byte[1_000_000];
        var after = before.ToDictionary(pair => pair.Key, pair => pair.Value);
        after["assets/changed.bin"] = new byte[1_000_000];
        var source = Package(before);
        var target = Package(after);
        var options = new PackageManifestOptions { MaxCompressionRatio = maximumRatio };
        var delta = PackageChangeSet.Create(source, target, options);

        var output = delta.Apply(source, options);
        Assert.True(PackageManifestGenerator.Generate(output, options).IsValid);
        AssertEntries(after, output);
        AssertEntries(before, delta.Invert().Apply(output, options));
    }

    [Fact]
    public void InvalidAmbiguousAndOverBudgetPackagesAreRejected()
    {
        var valid = DocxSession.CreateBlankDocxBytes();
        var duplicate = Entries(valid).ToList();
        duplicate.Add(duplicate.Single(p => p.Key == "word/document.xml"));
        var unsafeEntries = Entries(valid);
        unsafeEntries["../escape.xml"] = Utf8("<x/>");
        foreach (var invalid in new[] { Utf8("not a ZIP"), Package(duplicate), Package(unsafeEntries) })
        {
            Assert.Equal(PackageChangeError.InvalidPackage,
                Assert.Throws<PackageChangeException>(() => PackageChangeSet.Create(valid, invalid)).Code);
        }
        Assert.Equal(PackageChangeError.InvalidPackage,
            Assert.Throws<PackageChangeException>(() => PackageChangeSet.Create(valid, valid,
                new PackageManifestOptions { MaxEntryCount = 1 })).Code);
        var delta = PackageChangeSet.Create(valid, valid);
        Assert.Equal(PackageChangeError.InvalidPackage,
            Assert.Throws<PackageChangeException>(() => delta.Apply(valid,
                new PackageManifestOptions { MaxTotalUncompressedBytes = 1 })).Code);
    }

    [Fact]
    public void RealDocumentTrackedEditReplaysWithoutRetainingUnchangedStoriesOrAssets()
    {
        var before = File.ReadAllBytes("../../../../TestFiles/NVCA-Model-COI.docx");
        using var session = new DocxSession(before, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Package replay test",
        });
        var match = session.Grep(@"\[specify percentage\]").First();
        var result = session.ReplaceMatch(match, "a majority");
        Assert.True(result.Success, result.Error?.Message);
        var after = session.Save();
        var delta = PackageChangeSet.Create(before, after);

        Assert.Contains(delta.Changes, c => c.Uri == "/word/document.xml");
        Assert.DoesNotContain(delta.Changes, c => c.Uri.Contains("/media/", StringComparison.Ordinal)
            || c.Uri.Contains("header", StringComparison.Ordinal)
            || c.Uri.Contains("footer", StringComparison.Ordinal)
            || c.Uri.EndsWith("footnotes.xml", StringComparison.Ordinal));
        var replayed = delta.Apply(before);
        AssertEntries(Entries(after), replayed);
        AssertEntries(Entries(before), delta.Invert().Apply(after));
        using var reopened = new DocxSession(replayed);
        Assert.Contains(reopened.ListRevisions(), r => r.Author == "Package replay test");
    }

    private static Dictionary<string, byte[]> Entries(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var input = entry.Open();
            using var buffer = new MemoryStream();
            input.CopyTo(buffer);
            return buffer.ToArray();
        });
    }

    private static byte[] Package(IEnumerable<KeyValuePair<string, byte[]>> entries,
        CompressionLevel compression = CompressionLevel.NoCompression, int year = 2024)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, bytes) in entries)
            {
                var entry = archive.CreateEntry(name, compression);
                entry.LastWriteTime = new DateTimeOffset(year, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var output = entry.Open();
                output.Write(bytes);
            }
        }
        return buffer.ToArray();
    }

    private static void DeclareBinary(Dictionary<string, byte[]> entries) =>
        entries["[Content_Types].xml"] = Utf8(Encoding.UTF8.GetString(entries["[Content_Types].xml"])
            .Replace("</Types>", "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/></Types>"));

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);

    private static void AssertEntries(Dictionary<string, byte[]> expected, byte[] actual)
    {
        var entries = Entries(actual);
        Assert.Equal(expected.Keys.OrderBy(k => k), entries.Keys.OrderBy(k => k));
        foreach (var (name, bytes) in expected) Assert.Equal(bytes, entries[name]);
        Assert.True(PackageManifestGenerator.Generate(actual).IsValid);
    }
}
