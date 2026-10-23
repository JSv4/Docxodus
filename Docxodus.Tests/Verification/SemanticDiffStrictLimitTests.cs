using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

public class SemanticDiffStrictLimitTests
{
    public static IEnumerable<object[]> LimitCases()
    {
        foreach (var strict in new[] { false, true })
        foreach (var supplement in new[] { false, true })
        foreach (var bytes in new[] { false, true })
        foreach (var limit in new[] { "xml", "entry", "total", "ratio" })
            yield return new object[] { strict, supplement, bytes, limit };
    }

    [Theory]
    [MemberData(nameof(LimitCases))]
    public void Limits_apply_to_the_supplied_package_not_its_transitional_conversion(
        bool strict, bool supplement, bool bytes, string limit)
    {
        var document = Document(strict);
        var options = new SemanticDiffOptions
        {
            IncludePackageChanges = supplement,
            PackageOptions = AtLimit(document.DocumentByteArray, limit),
        };
        Assert.True(PackageManifestGenerator.Generate(document.DocumentByteArray, options.PackageOptions).IsValid);

        var result = bytes
            ? SemanticDiff.Compare(document.DocumentByteArray, document.DocumentByteArray, options)
            : SemanticDiff.Compare(document, document, options);

        Assert.Empty(result.Changes);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Over_limit_strict_input_is_refused_on_either_side_even_without_the_supplement(
        bool supplement, bool oversizedRight)
    {
        var document = Document(strict: true);
        var smaller = DocxDiffStrictConformanceTests.AsStrict(IrTestDocuments.Create("Small"));
        foreach (var limit in new[] { "xml", "entry", "total", "ratio" })
        {
            var options = new SemanticDiffOptions
            {
                IncludePackageChanges = supplement,
                PackageOptions = AtLimit(smaller.DocumentByteArray, limit),
            };
            var error = Assert.Throws<InvalidDataException>(() => SemanticDiff.Compare(
                oversizedRight ? smaller.DocumentByteArray : document.DocumentByteArray,
                oversizedRight ? document.DocumentByteArray : smaller.DocumentByteArray,
                options));
            Assert.Contains(oversizedRight ? "right package" : "left package", error.Message);
        }
    }

    [Fact]
    public void Strict_and_transitional_edits_keep_the_same_semantic_evidence()
    {
        var before = IrTestDocuments.Create("Before", "Unchanged");
        var after = IrTestDocuments.Create("After", "Unchanged");
        var strictBefore = DocxDiffStrictConformanceTests.AsStrict(before);
        var strictAfter = DocxDiffStrictConformanceTests.AsStrict(after);
        var expected = SemanticDiff.Compare(before, after).ToCanonicalJson();

        Assert.Equal(expected, SemanticDiff.Compare(strictBefore, strictAfter).ToCanonicalJson());
        Assert.Equal(expected, SemanticDiff.Compare(strictBefore, after).ToCanonicalJson());
        Assert.Equal(expected, SemanticDiff.Compare(before, strictAfter).ToCanonicalJson());
    }

    private static WmlDocument Document(bool strict)
    {
        // The document part is larger than content types and relationships, so conversion grows
        // both the largest XML/entry and the aggregate beyond their exact raw-byte boundaries.
        var document = IrTestDocuments.Create(string.Join(" ", Enumerable.Repeat("Contract clause", 100)));
        using (var stream = new MemoryStream(document.DocumentByteArray))
        using (var package = WordprocessingDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator().Validate(package));
        return strict ? DocxDiffStrictConformanceTests.AsStrict(document) : document;
    }

    private static PackageManifestOptions AtLimit(byte[] bytes, string limit)
    {
        using var stream = new MemoryStream(bytes);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var largest = zip.Entries.Max(entry => entry.Length);
        return limit switch
        {
            "xml" => new PackageManifestOptions { MaxXmlPartBytes = largest },
            "entry" => new PackageManifestOptions { MaxEntryUncompressedBytes = largest },
            "total" => new PackageManifestOptions { MaxTotalUncompressedBytes = zip.Entries.Sum(entry => entry.Length) },
            "ratio" => new PackageManifestOptions { MaxCompressionRatio = zip.Entries.Max(entry => (double)entry.Length / entry.CompressedLength) },
            _ => throw new ArgumentOutOfRangeException(nameof(limit)),
        };
    }
}
