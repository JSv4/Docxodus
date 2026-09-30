// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// Regression gate for epic #835: a redline from <see cref="DocxCompare.Compare"/> must not contain an
/// Open XML SDK validator error (Office 2019) whose id neither input has.
///
/// <para><b>Sample.</b> Cross-pairs of <c>TestFiles/**/*.docx</c> fixtures, chosen by a stable hash of their
/// relative paths: a fixture joins the pool when <c>hash(seed|path)</c> falls under
/// <see cref="PoolPerMille"/>, and an ordered pair of pool fixtures is compared when
/// <c>hash(seed|left|right)</c> falls under <see cref="PairsPerMille"/>. Adding a fixture therefore only adds
/// pairs that involve it; it never reshuffles the pairs already sampled. Fixtures that do not open as
/// WordprocessingML are skipped. <see cref="RevisionPairs"/> adds every pair the corpus names as two versions of
/// one document, and <see cref="PinnedPairs"/> adds corpus pairs that reproduce a violation class neither
/// reaches.</para>
///
/// <para><b>Check.</b> For each pair, every output error is keyed by its <see cref="ValidationErrorInfo.Id"/> plus
/// the qualified name of the element it reports, and is a violation when neither input has an error with the same
/// key (so an input's duplicate bookmark id does not excuse an output's duplicate shape id). An output the SDK
/// cannot open, and a comparison that throws, are violations too.</para>
///
/// <para><b>Ratchet.</b> The violations must equal <see cref="AllowList"/> exactly: a violation the list does
/// not name fails the test, and so does a listed entry that no longer occurs, so the list can only shrink.
/// Epic #835 is done when it is empty. Deleting or renaming a fixture changes the sample; an allow-listed
/// pair that disappears that way is reported as stale.</para>
/// </summary>
public class DocxCompareOutputValidityGateTests
{
    private const string Seed = "835";

    /// <summary>Fixtures in the pool per thousand — about 140 of the current corpus.</summary>
    private const ulong PoolPerMille = 200;

    /// <summary>Ordered pool pairs compared per thousand — about 290 pairs.</summary>
    private const ulong PairsPerMille = 15;

    /// <summary>Corpus pairs that reproduce a violation class the seeded sample does not reach.</summary>
    private static readonly (string Left, string Right)[] PinnedPairs =
    [
        ("HW002-Table06.docx", "HW002-Table13.docx"), // duplicate theme / fontTable relationships
        ("HW002-Table06.docx", "HC005-TaskPlanTemplate.docx"), // duplicate styles / settings relationships
    ];

    /// <summary>
    /// Known residue, <c>"left | right | key"</c>. Every entry must still occur; remove it once fixed. The
    /// comment names the sub-issue of epic #835 that owns it.
    /// </summary>
    private static readonly string[] AllowList =
    [
        "CU004-Chart-Cached-Data-04.docx | DB010-FrontMatter.docx | Sem_UniqueAttributeValue@wp:docPr", // #860 duplicate drawing wp:docPr id
        "HC014-RTL-Table-01.docx | WC/WC014-SmartArt-Before.docx | Sem_UniqueAttributeValue@wp:docPr", // #860 duplicate drawing wp:docPr id
        "LIR015-en-US-cardinalText.docx | Blank-altChunk.docx | Sem_AttributeValueDataTypeDetailed@w:numberingChange", // #861 w:numberingChange/@w:original over 15 chars
        "LIR015-en-US-cardinalText.docx | LIR013-en-US-001.docx | Sem_AttributeValueDataTypeDetailed@w:numberingChange", // #861 w:numberingChange/@w:original over 15 chars
        "LIR015-en-US-cardinalText.docx | RC/RC005-Before.docx | Sem_AttributeValueDataTypeDetailed@w:numberingChange", // #861 w:numberingChange/@w:original over 15 chars
        "LIR019-en-US-lowerLetter.docx | DA009-InvalidXPath.docx | Sem_AttributeValueDataTypeDetailed@w:numberingChange", // #861 w:numberingChange/@w:original over 15 chars
        "LIR019-en-US-lowerLetter.docx | RP/RP044-MERGEFORMAT-Field-Code-Rejected.docx | Sem_AttributeValueDataTypeDetailed@w:numberingChange", // #861 w:numberingChange/@w:original over 15 chars
        "LIR019-en-US-lowerLetter.docx | RP/RP048-Deleted-Inserted-Para-Mark-Rejected.docx | Sem_AttributeValueDataTypeDetailed@w:numberingChange", // #861 w:numberingChange/@w:original over 15 chars
        "WC/WC014-SmartArt-Before.docx | CU004-Chart-Cached-Data-04.docx | Sem_UniqueAttributeValue@wp:docPr", // #860 duplicate drawing wp:docPr id
        "WC/WC037-Textbox-Before.docx | WC/WC037-Textbox-After1.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC037-Textbox-Before.docx | WC/WC037-Textbox-After1.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC044-Text-Box.docx | WC/WC044-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC044-Text-Box.docx | WC/WC044-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC045-Text-Box.docx | WC/WC045-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC045-Text-Box.docx | WC/WC045-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC046-Two-Text-Box.docx | WC/WC046-Two-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC046-Two-Text-Box.docx | WC/WC046-Two-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC047-Two-Text-Box.docx | WC/WC047-Two-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC047-Two-Text-Box.docx | WC/WC047-Two-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC048-Text-Box-in-Cell.docx | WC/WC048-Text-Box-in-Cell-Mod.docx | Sem_UniqueAttributeValue@v:group", // #860 duplicate VML shape id
        "WC/WC048-Text-Box-in-Cell.docx | WC/WC048-Text-Box-in-Cell-Mod.docx | Sem_UniqueAttributeValue@v:rect", // #860 duplicate VML shape id
        "WC/WC048-Text-Box-in-Cell.docx | WC/WC048-Text-Box-in-Cell-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC048-Text-Box-in-Cell.docx | WC/WC048-Text-Box-in-Cell-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC049-Text-Box-in-Cell.docx | WC/WC049-Text-Box-in-Cell-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC049-Text-Box-in-Cell.docx | WC/WC049-Text-Box-in-Cell-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC050-Table-in-Text-Box.docx | WC/WC050-Table-in-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC050-Table-in-Text-Box.docx | WC/WC050-Table-in-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC051-Table-in-Text-Box.docx | WC/WC051-Table-in-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC051-Table-in-Text-Box.docx | WC/WC051-Table-in-Text-Box-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC065-Textbox.docx | WC/WC065-Textbox-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC065-Textbox.docx | WC/WC065-Textbox-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC067-Textbox-Image.docx | WC/WC067-Textbox-Image-Mod.docx | Sem_UniqueAttributeValue@v:shape", // #860 duplicate VML shape id
        "WC/WC067-Textbox-Image.docx | WC/WC067-Textbox-Image-Mod.docx | Sem_UniqueAttributeValue@v:shapetype", // #860 duplicate VML shapetype id
        "WC/WC067-Textbox-Image.docx | WC/WC067-Textbox-Image-Mod.docx | Sem_UniqueAttributeValue@wp:docPr", // #860 duplicate drawing wp:docPr id
    ];

    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static readonly ConcurrentDictionary<string, CorpusInput?> Inputs = new(StringComparer.Ordinal);

    private static readonly ParallelOptions Parallelism = new() { MaxDegreeOfParallelism = Math.Max(1, Environment.ProcessorCount) };

    private readonly ITestOutputHelper _output;

    public DocxCompareOutputValidityGateTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void Compare_OutputIntroducesNoValidatorErrorAbsentFromBothInputs()
    {
        var stopwatch = Stopwatch.StartNew();
        TimeSpan cpuAtStart = Process.GetCurrentProcess().TotalProcessorTime;
        List<(string Left, string Right)> pairs = SampledPairs();
        Assert.True(pairs.Count >= 200, $"The seeded sample shrank to {pairs.Count} pairs; check the TestFiles glob.");

        var found = new ConcurrentDictionary<string, string>(StringComparer.Ordinal);
        Parallel.ForEach(pairs, Parallelism, pair =>
        {
            foreach (var (key, sample) in IntroducedViolations(pair.Left, pair.Right))
                found.TryAdd($"{pair.Left} | {pair.Right} | {key}", sample);
        });

        _output.WriteLine(
            $"{pairs.Count} pairs, {found.Count} introduced violations, {stopwatch.Elapsed.TotalSeconds:F1}s wall, " +
            $"{(Process.GetCurrentProcess().TotalProcessorTime - cpuAtStart).TotalSeconds:F1}s CPU");
        AssertMatchesAllowList(found);
    }

    /// <summary>The allow-list names each entry once, well formed, and only for pairs the gate compares.</summary>
    [Fact]
    public void AllowList_EntriesAreUniqueAndCompared()
    {
        Assert.Empty(AllowList.GroupBy(entry => entry, StringComparer.Ordinal).Where(g => g.Count() > 1).Select(g => g.Key));
        Assert.DoesNotContain(AllowList, entry => entry.Split(" | ").Length != 3);

        var compared = CandidatePairs().Select(pair => $"{pair.Left} | {pair.Right}").ToHashSet(StringComparer.Ordinal);
        Assert.DoesNotContain(AllowList, entry => !compared.Contains(entry[..entry.LastIndexOf(" | ", StringComparison.Ordinal)]));
    }

    private static void AssertMatchesAllowList(IReadOnlyDictionary<string, string> found)
    {
        var allowed = AllowList.ToHashSet(StringComparer.Ordinal);
        var added = found.Keys.Where(key => !allowed.Contains(key)).OrderBy(key => key, StringComparer.Ordinal).ToList();
        var gone = allowed.Where(key => !found.ContainsKey(key)).OrderBy(key => key, StringComparer.Ordinal).ToList();
        if (added.Count == 0 && gone.Count == 0)
            return;

        var message = new StringBuilder();
        if (added.Count > 0)
        {
            message.AppendLine($"{added.Count} compare output violation(s) absent from both inputs. Fix the cause; allow-list only residue a");
            message.AppendLine("sub-issue of #835 tracks, by adding to AllowList:");
            foreach (var key in added)
                message.AppendLine($"    \"{key}\", // {found[key]}");
        }

        if (gone.Count > 0)
        {
            message.AppendLine($"{gone.Count} allow-listed violation(s) no longer occur. The ratchet only shrinks; remove from AllowList:");
            foreach (var key in gone)
                message.AppendLine($"    \"{key}\",");
        }

        Assert.Fail(message.ToString());
    }

    /// <summary>The compared pairs whose fixtures both open as WordprocessingML documents.</summary>
    private static List<(string Left, string Right)> SampledPairs()
    {
        List<(string Left, string Right)> candidates = CandidatePairs();
        Parallel.ForEach(candidates.SelectMany(pair => new[] { pair.Left, pair.Right }).Distinct(StringComparer.Ordinal), Parallelism, path => Input(path));
        return candidates.Where(pair => Input(pair.Left) is not null && Input(pair.Right) is not null).ToList();
    }

    /// <summary>The seeded sample plus the pinned pairs, in ordinal order.</summary>
    private static List<(string Left, string Right)> CandidatePairs()
    {
        List<string> corpus = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
            .Where(file => !file.Name.StartsWith("~$", StringComparison.Ordinal))
            .Select(file => Path.GetRelativePath(TestFilesDir.FullName, file.FullName).Replace('\\', '/'))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToList();
        List<string> pool = corpus.Where(path => StableHash($"{Seed}|{path}") % 1000UL < PoolPerMille).ToList();

        var pairs = new SortedSet<(string Left, string Right)>(PinnedPairs.Concat(RevisionPairs(corpus)), PairComparer.Instance);
        foreach (var left in pool)
        {
            foreach (var right in pool)
            {
                if (left != right && StableHash($"{Seed}|{left}|{right}") % 1000UL < PairsPerMille)
                    pairs.Add((left, right));
            }
        }

        return pairs.ToList();
    }

    /// <summary>
    /// Every fixture pair the corpus names as two versions of one document — <c>X</c> and <c>X-Mod</c>,
    /// <c>X-Before</c> and each <c>X-After*</c> — compared in authoring order. These are the comparisons users
    /// run, and they reach classes (text boxes from both versions in one part) that unrelated cross-pairs miss.
    /// </summary>
    private static IEnumerable<(string Left, string Right)> RevisionPairs(IReadOnlyCollection<string> corpus)
    {
        var present = corpus.ToHashSet(StringComparer.Ordinal);
        foreach (var path in corpus)
        {
            if (path.EndsWith("-Mod.docx", StringComparison.Ordinal) && present.Contains(path[..^"-Mod.docx".Length] + ".docx"))
                yield return (path[..^"-Mod.docx".Length] + ".docx", path);

            int before = path.LastIndexOf("-Before", StringComparison.Ordinal);
            if (before < 0)
                continue;
            string stem = path[..before];
            foreach (var after in corpus.Where(other => other.StartsWith(stem + "-After", StringComparison.Ordinal)))
                yield return (path, after);
        }
    }

    /// <summary>FNV-1a over UTF-8 with a final avalanche; independent of the runtime's string hashing.</summary>
    private static ulong StableHash(string text)
    {
        ulong hash = 14695981039346656037UL;
        foreach (byte b in Encoding.UTF8.GetBytes(text))
        {
            hash ^= b;
            hash *= 1099511628211UL;
        }

        hash ^= hash >> 33;
        hash *= 0xff51afd7ed558ccdUL;
        hash ^= hash >> 33;
        return hash;
    }

    /// <summary>The output violations of one pair whose error id neither input has, one sample per key.</summary>
    private static IEnumerable<(string Key, string Sample)> IntroducedViolations(string left, string right)
    {
        CorpusInput original = Input(left)!;
        CorpusInput revised = Input(right)!;

        WmlDocument output;
        try
        {
            output = DocxCompare.Compare(original.Document, revised.Document);
        }
        catch (Exception ex)
        {
            return new[] { ("CompareThrew", $"{ex.GetType().Name}: {ex.Message}") };
        }

        IReadOnlyList<ValidatorError> errors;
        try
        {
            errors = Validate(output.DocumentByteArray);
        }
        catch (Exception ex)
        {
            return new[] { ("PackageOpenFailure", $"{ex.GetType().Name}: {ex.Message}") };
        }

        return errors
            .Where(error => !original.ErrorKeys.Contains(error.Key) && !revised.ErrorKeys.Contains(error.Key))
            .GroupBy(error => error.Key, StringComparer.Ordinal)
            .Select(group => (group.Key, group.First().Sample));
    }

    private static IReadOnlyList<ValidatorError> Validate(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        using var document = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(document)
            .Select(error => new ValidatorError(
                error.Id,
                $"{error.Id}@{(error.Node is null ? "(package)" : $"{error.Node.Prefix}:{error.Node.LocalName}")}",
                error.Description.Length <= 160 ? error.Description : error.Description[..160] + "…"))
            .ToList();
    }

    /// <summary>A fixture and its validator error ids, or null when it is not an openable WordprocessingML document.</summary>
    private static CorpusInput? Input(string path) => Inputs.GetOrAdd(path, static relative =>
    {
        try
        {
            byte[] bytes = File.ReadAllBytes(Path.Combine(TestFilesDir.FullName, relative));
            using (var stream = new MemoryStream())
            {
                stream.Write(bytes, 0, bytes.Length);
                using var document = WordprocessingDocument.Open(stream, false);
                if (document.MainDocumentPart?.Document?.Body is null)
                    return null;
            }

            return new CorpusInput(new WmlDocument(relative, bytes), Validate(bytes).Select(error => error.Key).ToHashSet(StringComparer.Ordinal));
        }
        catch (Exception)
        {
            return null;
        }
    });

    private sealed record CorpusInput(WmlDocument Document, HashSet<string> ErrorKeys);

    private sealed record ValidatorError(string Id, string Key, string Sample);

    private sealed class PairComparer : IComparer<(string Left, string Right)>
    {
        public static readonly PairComparer Instance = new();

        public int Compare((string Left, string Right) x, (string Left, string Right) y)
        {
            int left = string.CompareOrdinal(x.Left, y.Left);
            return left != 0 ? left : string.CompareOrdinal(x.Right, y.Right);
        }
    }
}
