#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Compares two SEPARATELY PRODUCED DOCX packages for equivalence, part by part.
///
/// <para>Raw <see cref="DocxodusDocument.DocumentByteArray"/> equality cannot express this. A DOCX is a
/// ZIP, and every entry carries the wall-clock time at which it was written, so two packages built from
/// identical content a moment apart differ in those header bytes — and once a stored timestamp changes,
/// the surrounding compressed stream shifts with it. The two calls under test run milliseconds apart and
/// so USUALLY land in the same timestamp granule; on a loaded CI runner they sometimes straddle one, and
/// the assertion fails on a single byte with no relation to the behavior being tested.</para>
///
/// <para>Comparing the part set and each part's bytes keeps the whole claim — same parts, same content,
/// same relationships — while dropping only the container metadata that no test means to assert. Verified
/// against the failure mode directly: two packages produced four seconds apart differ in raw bytes and
/// agree on every part.</para>
/// </summary>
internal static class PackageEquivalence
{
    /// <summary>Asserts that <paramref name="actual"/> holds exactly the parts of <paramref name="expected"/>, byte for byte.</summary>
    public static void AssertSamePackage(WmlDocument expected, WmlDocument actual)
    {
        Dictionary<string, byte[]> expectedParts = ReadParts(expected);
        Dictionary<string, byte[]> actualParts = ReadParts(actual);

        Assert.Equal(
            expectedParts.Keys.OrderBy(uri => uri, StringComparer.Ordinal),
            actualParts.Keys.OrderBy(uri => uri, StringComparer.Ordinal));

        foreach (string uri in expectedParts.Keys.OrderBy(uri => uri, StringComparer.Ordinal))
        {
            byte[] expectedPart = expectedParts[uri];
            byte[] actualPart = actualParts[uri];
            if (expectedPart.AsSpan().SequenceEqual(actualPart))
                continue;

            Assert.Fail($"package part {uri} differs: {DescribeDifference(expectedPart, actualPart)}");
        }
    }

    private static Dictionary<string, byte[]> ReadParts(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray.ToArray());
        using var wordDocument = WordprocessingDocument.Open(stream, false);

        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var part in wordDocument.GetPackage().GetParts())
        {
            using var partStream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var buffer = new MemoryStream();
            partStream.CopyTo(buffer);
            parts[part.Uri.ToString()] = buffer.ToArray();
        }
        return parts;
    }

    /// <summary>The first differing offset plus a short window either side, mirroring xUnit's byte-collection report.</summary>
    private static string DescribeDifference(byte[] expected, byte[] actual)
    {
        int shared = Math.Min(expected.Length, actual.Length);
        int position = 0;
        while (position < shared && expected[position] == actual[position])
            position++;

        if (position == shared)
            return $"lengths {expected.Length} and {actual.Length}, identical through byte {shared}";

        return $"first difference at byte {position} " +
            $"(expected {Window(expected, position)}, actual {Window(actual, position)}); " +
            $"lengths {expected.Length} and {actual.Length}";
    }

    private static string Window(byte[] bytes, int position)
    {
        int start = Math.Max(0, position - 2);
        int end = Math.Min(bytes.Length, position + 3);
        return "[" + string.Join(", ", bytes[start..end]) + "]";
    }
}
