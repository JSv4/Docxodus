// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

[CollectionDefinition("Long-paragraph token diff memory", DisableParallelization = true)]
public sealed class LongParagraphTokenDiffCollection
{
}

/// <summary>
/// A single very long paragraph diffs in bounded memory and time (issue #964). The word-level anchor
/// pass used to fill a full n·m match table for any paragraph pair: about 1 GB for a 16k-word paragraph,
/// and a 20k-word pair needed 1.6 GB — fatal in a WASM heap. These tests run alone so the allocation
/// counter sees only the diff under test; the bounds are several times what the capped path uses and
/// far below what the uncapped table needs.
/// </summary>
[Collection("Long-paragraph token diff memory")]
public class IrTokenDifferLongParagraphTests
{
    private const int WordCount = 20_000;
    private const long MemoryBudgetBytes = 256L * 1024 * 1024;
    private static readonly TimeSpan TimeBudget = TimeSpan.FromSeconds(30);
    private static readonly IrRunFormat Plain = new() { Bold = false, UnmodeledDigest = default };

    [Fact]
    public void LongParagraph_with_a_few_edits_diffs_in_bounded_memory_and_keeps_the_shared_words()
    {
        var leftWords = RandomWords(seed: 964, WordCount);
        var rightWords = (string[])leftWords.Clone();
        int[] edited = { 17, 4_000, 9_999, 15_123, 19_990 };
        foreach (var i in edited)
            rightWords[i] = "edited" + i;

        var (diff, allocated, elapsed) = Measure(Tokens(leftWords), Tokens(rightWords));

        Assert.True(allocated < MemoryBudgetBytes, $"diff allocated {allocated / (1024 * 1024)} MB");
        Assert.True(elapsed < TimeBudget, $"diff took {elapsed}");

        // The fallback alignment must still find the shared text: only the five edited words change.
        var deletedWords = diff.Ops.Where(o => o.Kind == IrTokenOpKind.Delete).Sum(o => o.LeftLength);
        var insertedWords = diff.Ops.Where(o => o.Kind == IrTokenOpKind.Insert).Sum(o => o.RightLength);
        Assert.Equal(edited.Length, deletedWords);
        Assert.Equal(edited.Length, insertedWords);
    }

    [Fact]
    public void LongParagraph_pair_that_is_mostly_different_diffs_in_bounded_memory()
    {
        var left = Tokens(RandomWords(seed: 1, WordCount));
        var right = Tokens(RandomWords(seed: 2, WordCount));

        var (_, allocated, elapsed) = Measure(left, right);

        Assert.True(allocated < MemoryBudgetBytes, $"diff allocated {allocated / (1024 * 1024)} MB");
        Assert.True(elapsed < TimeBudget, $"diff took {elapsed}");
    }

    [Fact]
    public void LongParagraph_compare_through_DocxDiff_reports_only_the_edited_word()
    {
        var leftWords = RandomWords(seed: 7, WordCount);
        var rightWords = (string[])leftWords.Clone();
        rightWords[12_345] = "rewritten";
        var left = IrTestDocuments.Create(string.Join(" ", leftWords));
        var right = IrTestDocuments.Create(string.Join(" ", rightWords));

        long before = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        var revisions = DocxDiff.GetRevisions(left, right);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - before;

        Assert.True(allocated < 2 * MemoryBudgetBytes, $"compare allocated {allocated / (1024 * 1024)} MB");
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(60), $"compare took {stopwatch.Elapsed}");

        Assert.Contains(revisions, r => r.Text != null && r.Text.Contains("rewritten", StringComparison.Ordinal));
        Assert.DoesNotContain(revisions, r => r.Text != null && r.Text.Length > 100);
    }

    [Fact]
    public void LongParagraph_built_from_a_few_repeated_words_diffs_in_bounded_memory()
    {
        // No key is unique, or rare enough to split on: the bounded path must still terminate quickly.
        string[] vocabulary = { "the", "of", "and" };
        var random = new Random(5);
        var leftWords = Enumerable.Range(0, WordCount).Select(_ => vocabulary[random.Next(3)]).ToArray();
        var rightWords = Enumerable.Range(0, WordCount).Select(_ => vocabulary[random.Next(3)]).ToArray();

        var (_, allocated, elapsed) = Measure(Tokens(leftWords), Tokens(rightWords));

        Assert.True(allocated < MemoryBudgetBytes, $"diff allocated {allocated / (1024 * 1024)} MB");
        Assert.True(elapsed < TimeBudget, $"diff took {elapsed}");
    }

    [Fact]
    public void LongParagraph_diff_is_deterministic()
    {
        var left = Tokens(RandomWords(seed: 3, WordCount));
        var right = Tokens(RandomWords(seed: 4, WordCount));

        var first = IrTokenDiffer.Diff(left, right, new IrDiffSettings());
        var second = IrTokenDiffer.Diff(left, right, new IrDiffSettings());

        Assert.Equal(first.Ops.ToList(), second.Ops.ToList());
    }

    private static (IrTokenDiff Diff, long Allocated, TimeSpan Elapsed) Measure(
        List<IrDiffToken> left, List<IrDiffToken> right)
    {
        long before = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        var diff = IrTokenDiffer.Diff(left, right, new IrDiffSettings());
        stopwatch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        IrTokenDiffAsserts.AssertInvariants(left, right, diff);
        return (diff, allocated, stopwatch.Elapsed);
    }

    /// <summary>Deterministic prose-like words: a log-uniform rank over a 5,000-word vocabulary, so a few
    /// words are very common and most are rare — the repetition profile of real text.</summary>
    private static string[] RandomWords(int seed, int count)
    {
        var random = new Random(seed);
        var words = new string[count];
        for (int i = 0; i < count; i++)
        {
            int rank = (int)Math.Exp(random.NextDouble() * Math.Log(5_000));
            words[i] = WordFor(rank);
        }
        return words;
    }

    private static string WordFor(int rank)
    {
        var sb = new StringBuilder();
        do
        {
            sb.Append((char)('a' + (rank % 26)));
            rank /= 26;
        }
        while (rank > 0);
        return sb.ToString();
    }

    private static List<IrDiffToken> Tokens(string[] words)
    {
        var tokens = new List<IrDiffToken>(words.Length * 2);
        for (int i = 0; i < words.Length; i++)
        {
            if (i > 0)
                tokens.Add(new IrDiffToken(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain));
            tokens.Add(new IrDiffToken(IrDiffTokenKind.Word, words[i], words[i], 0, words[i].Length, Plain));
        }
        return tokens;
    }
}
