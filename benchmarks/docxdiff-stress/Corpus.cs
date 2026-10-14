// Corpus-wide differential: run every document in a directory through the DocxDiff products and
// digest the results, so two builds can be compared byte for byte.
//
// The eight generated variants of one reference document answer "did this change the diff of a
// heavyweight legal form". They do NOT answer "did this change the diff of anything else", and the
// reference document happens to carry no tracked revisions, no tables and no drawings — three of the
// shapes the engine's riskiest code paths are keyed on. This mode closes that: point it at
// TestFiles/ and it exercises every document the repository has.
//
// Failures are digested too. An exception is a result like any other, and a change in WHICH
// exception a malformed document produces is exactly the kind of regression worth catching.
//
// The unit of comparison is an Observation, not a digest string: a product call has more
// observable effects than its return value, and the #616 regression lived on one of the others
// (see Observation.cs). Adding a channel is one field on that record rather than a new sink[...]
// family per mode.

using System.Collections.Concurrent;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Docxodus;

namespace Docxodus.Stress;

internal static class Corpus
{
    public static int Run(string root, string outPath, string? checkPath, int? limit, int threads)
    {
        var files = Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
            .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                     || f.EndsWith(".docm", StringComparison.OrdinalIgnoreCase)
                     || f.EndsWith(".dotx", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();
        if (limit is > 0) files = files.Take(limit.Value).ToList();

        Console.WriteLine($"corpus  : {root}");
        Console.WriteLine($"documents: {files.Count:N0}");
        Console.WriteLine($"threads : {threads}");
        Console.WriteLine();

        var digests = new ConcurrentDictionary<string, Observation>(StringComparer.Ordinal);
        var done = 0;
        var errors = 0;

        Parallel.ForEach(files.Select((f, i) => (Path: f, Index: i)),
            new ParallelOptions { MaxDegreeOfParallelism = threads }, item =>
        {
            var (path, index) = item;
            var name = Path.GetRelativePath(root, path).Replace('\\', '/');
            byte[] bytes;
            try { bytes = File.ReadAllBytes(path); }
            catch (Exception ex) { digests[$"{name}#read"] = Observation.NotRun($"READ-FAIL {ex.GetType().Name}"); return; }

            // Every document is compared against a deterministically edited copy of itself, and
            // against itself unchanged (the identical-bytes fast paths).
            byte[] edited;
            try { edited = CorpusVariant.Edit(bytes); }
            catch (Exception ex) { digests[$"{name}#variant"] = Observation.NotRun($"VARIANT-FAIL {ex.GetType().Name}"); edited = bytes; }

            // A second, differently-offset edit of the same document: the N-way case needs two
            // reviewers who actually disagree, or the merger never has a competitor to order.
            byte[] editedAlt;
            try { editedAlt = CorpusVariant.Edit(bytes, offset: 2); }
            catch (Exception ex) { digests[$"{name}#variant-alt"] = Observation.NotRun($"VARIANT-FAIL {ex.GetType().Name}"); editedAlt = bytes; }

            var left = new WmlDocument("left.docx", bytes);
            var right = new WmlDocument("right.docx", edited);
            var same = new WmlDocument("same.docx", bytes);
            var alt = new WmlDocument("alt.docx", editedAlt);

            // Default settings over the edited pair: the ordinary comparison path.
            Record(digests, name, "edited", left, right, new DocxDiffSettings());

            // The same document on both sides: the byte-identical shortcuts, including the one
            // GetRevisions gained. Cheap, and it pins the shortcut against the full pipeline.
            Record(digests, name, "identical", left, same, new DocxDiffSettings());

            // The revision-view transform is only reached by a document that actually carries
            // revision markup, and it is the path IrReader.Read was restructured around. Force both
            // input-revision policies so the pre-accept flatten and the preserve renderer both run.
            Record(digests, name, "preaccept", left, right,
                new DocxDiffSettings { PreAcceptInputRevisions = true });
            Record(digests, name, "preserve", left, right,
                new DocxDiffSettings { PreserveInputRevisions = true });

            // A fifth comparison with ONE non-default setting, rotated per document so the whole
            // DocxDiffSettings surface gets exercised across the corpus for one extra comparison per
            // document rather than a combinatorial explosion (issue #624, item 3).
            var rotated = SettingsRotation.For(index);
            Record(digests, name, $"rot-{rotated.Name}", left, right, rotated.Settings);

            // N-way. The consolidate path has its own reader fan-out, its own merger and its own
            // markup renderer, and NONE of the three products above touches any of them. Two
            // reviewers off the same base is the smallest shape that exercises reviewer ordering and
            // conflict competitor order.
            RecordConsolidate(digests, name, "consolidate", left, right, alt,
                new DocxDiffConsolidateSettings());

            // And the N-way rotation (issue #632): the same one-non-default-setting argument as the
            // fifth pairwise comparison, applied to the composed settings surface — the pairwise
            // variations wrapped in a consolidate settings object, plus the conflict-resolution
            // policies that exist only there.
            var consolidateRotation = SettingsRotation.ForConsolidate(index);
            RecordConsolidate(digests, name, $"consolidate-rot-{consolidateRotation.Name}",
                left, right, alt, consolidateRotation.Settings);

            var n = Interlocked.Increment(ref done);
            if (n % 50 == 0) Console.WriteLine($"  {n,5}/{files.Count} ...");
        });

        foreach (var kv in digests)
            if (kv.Value.Result.Contains("FAIL", StringComparison.Ordinal)) errors++;

        var sorted = new SortedDictionary<string, Observation>(digests, StringComparer.Ordinal);
        File.WriteAllText(outPath, JsonSerializer.Serialize(sorted, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine();
        Console.WriteLine($"{sorted.Count:N0} digests over {files.Count:N0} documents written to {outPath}");
        Console.WriteLine($"({errors:N0} of them record a thrown exception rather than a result — that is expected for");
        Console.WriteLine(" malformed or unsupported fixtures, and a CHANGE in them is still a regression.)");
        ReportChannelCoverage(sorted);

        if (checkPath is null) return 0;

        var expected = JsonSerializer.Deserialize<SortedDictionary<string, Observation>>(File.ReadAllText(checkPath))!;
        var mismatches = 0;
        foreach (var (k, want) in expected)
        {
            if (!sorted.TryGetValue(k, out var got)) { Console.WriteLine($"[parity] MISSING  {k}"); mismatches++; }
            else if (got.DifferencesFrom(want) is { Count: > 0 } diffs)
            {
                Console.WriteLine($"[parity] MISMATCH {k}");
                foreach (var d in diffs) Console.WriteLine($"           {d}");
                mismatches++;
            }
        }

        foreach (var k in sorted.Keys.Where(k => !expected.ContainsKey(k)))
        {
            Console.WriteLine($"[parity] EXTRA    {k}");
            mismatches++;
        }

        Console.WriteLine();
        Console.WriteLine(mismatches == 0
            ? $"[parity] OK - all {expected.Count:N0} digests identical across {files.Count:N0} documents"
            : $"[parity] {mismatches:N0} MISMATCH(ES) of {expected.Count:N0}");
        return mismatches == 0 ? 0 : 2;
    }

    // ─── One product call, every channel ─────────────────────────────────
    //
    // Each product is run ONCE with a warning-collecting settings clone, with both inputs hashed
    // around the call. The pre-flight report used to need its own second run of every product;
    // capturing it from the call under observation is both cheaper and stricter, because it is the
    // same call that produced the result.

    private static void Record(
        ConcurrentDictionary<string, Observation> sink, string name, string mode,
        WmlDocument left, WmlDocument right, DocxDiffSettings settings)
    {
        var products = new (string Key, Func<string> Run)[]
        {
            ("redline", () => StableRedlineDigest(DocxDiff.Compare(left, right, settings).DocumentByteArray)),
            ("revisions", () => RevisionsDigest(DocxDiff.GetRevisions(left, right, settings))),
            ("editscript", () => Sha(Encoding.UTF8.GetBytes(DocxDiff.GetEditScriptJson(left, right, settings)))),
        };

        // The pre-flight report is captured from the SAME call that produces the result, rather than
        // from a second run of every product: cheaper, and stricter, because a product that stops
        // running the pre-flight records "none" against the very call whose result it produced.
        var reported = new List<string>();
        var callerCallback = settings.OnCompatibilityWarning;
        settings.OnCompatibilityWarning = report =>
        {
            foreach (var w in report.Warnings.OrderBy(w => w.Feature.Id, StringComparer.Ordinal))
                reported.Add(w.Feature.Id);
            callerCallback?.Invoke(report);
        };

        var observed = new Dictionary<string, (string Result, string Warnings, string Mutation)>(
            StringComparer.Ordinal);
        foreach (var (key, run) in products)
        {
            reported.Clear();
            var beforeLeft = Sha(left.DocumentByteArray);
            var beforeRight = Sha(right.DocumentByteArray);
            var result = Try(run);
            var mutation = MutationOf(beforeLeft, beforeRight, left, right);
            observed[key] = (result, reported.Count == 0 ? "none" : string.Join(",", reported), mutation);
        }

        // Order independence. Since #616 the statics delegate to a memoized DocxDiffComparison that
        // shares ONE provenance-bearing IR snapshot across every product, so asking for the products
        // in a different order traverses different shared state. Ask a single comparison for all
        // three in REVERSE order and require each to match what the static produced — which also
        // pins CreateComparison(l, r, s) against DocxDiff.Compare(l, r, s), a class the corpus
        // otherwise never reaches because it only ever calls the statics.
        //
        // Each product is wrapped separately rather than the block as a whole, so a settings mode
        // that makes every product throw (ThrowOnCompatibilityWarning) compares one recorded failure
        // against the other and reads as stable, which it is.
        DocxDiffComparison? comparison = null;
        string? createFailure = null;
        try { comparison = DocxDiff.CreateComparison(left, right, settings); }
        catch (Exception ex) { createFailure = $"FAIL {ex.GetType().Name}: {Truncate(ex.Message)}"; }

        string ViaComparison(Func<string> run) => createFailure ?? Try(run);
        var reorderedScript = ViaComparison(() => Sha(Encoding.UTF8.GetBytes(comparison!.GetEditScriptJson())));
        var reorderedRevisions = ViaComparison(() => RevisionsDigest(comparison!.GetRevisions()));
        var reorderedRedline = ViaComparison(() =>
            StableRedlineDigest(comparison!.ToRedline().DocumentByteArray));

        string Variance(string key, string reordered) =>
            observed[key].Result == reordered ? "stable" : $"reordered {reordered}";

        Store(sink, name, mode, "redline", observed["redline"], Variance("redline", reorderedRedline));
        Store(sink, name, mode, "revisions", observed["revisions"], Variance("revisions", reorderedRevisions));
        Store(sink, name, mode, "editscript", observed["editscript"], Variance("editscript", reorderedScript));
    }

    private static void Store(
        ConcurrentDictionary<string, Observation> sink, string name, string mode, string product,
        (string Result, string Warnings, string Mutation) observed, string variance) =>
        sink[$"{name}#{mode}/{product}"] = new Observation
        {
            Result = observed.Result,
            Warnings = observed.Warnings,
            InputMutation = observed.Mutation,
            OrderVariance = variance,
        };

    // Every consolidate product, over two reviewers off one base. Reviewer order is significant to
    // conflict reporting, so the digest carries the authors and the per-revision order as emitted.
    //
    // This is a structural mirror of Record. It used to record Warnings and OrderVariance as "n/a"
    // on two premises that were both false by the time they were written down (issue #632): the
    // consolidate settings type COMPOSES DocxDiffSettings rather than inheriting it, so
    // settings.Diff carries the compatibility subscription like any other; and since #617 the four
    // statics each delegate to a single-use DocxDiffConsolidation, whose caller-held form shares
    // one memoized read/merge across every product in whatever order the caller asks. The first
    // false premise is what let the #629 gap — three of the four N-way entry points never ran the
    // pre-flight — hide behind thousands of rows that looked as though the question had been asked
    // and answered.
    private static void RecordConsolidate(
        ConcurrentDictionary<string, Observation> sink, string name, string mode,
        WmlDocument baseDoc, WmlDocument reviewerA, WmlDocument reviewerB,
        DocxDiffConsolidateSettings settings)
    {
        var reviewers = new List<DocxDiffReviewer>
        {
            new() { Author = "Reviewer A", Document = reviewerA },
            new() { Author = "Reviewer B", Document = reviewerB },
        };

        var products = new (string Key, Func<string> Run)[]
        {
            ("redline", () =>
                StableRedlineDigest(DocxDiff.Consolidate(baseDoc, reviewers, settings).DocumentByteArray)),
            ("revisions", () =>
                ConsolidatedRevisionsDigest(DocxDiff.GetConsolidatedRevisions(baseDoc, reviewers, settings))),
            ("conflicts", () => ConflictsDigest(DocxDiff.GetConflicts(baseDoc, reviewers, settings))),
            ("editscript", () =>
                Sha(Encoding.UTF8.GetBytes(DocxDiff.GetConsolidatedEditScriptJson(baseDoc, reviewers, settings)))),
        };

        var reported = new List<string>();
        var callerCallback = settings.Diff.OnCompatibilityWarning;
        settings.Diff.OnCompatibilityWarning = report =>
        {
            foreach (var w in report.Warnings.OrderBy(w => w.Feature.Id, StringComparer.Ordinal))
                reported.Add(w.Feature.Id);
            callerCallback?.Invoke(report);
        };

        var observed = new Dictionary<string, (string Result, string Warnings, string Mutation)>(
            StringComparer.Ordinal);
        foreach (var (key, run) in products)
        {
            reported.Clear();
            var beforeBase = Sha(baseDoc.DocumentByteArray);
            var beforeA = Sha(reviewerA.DocumentByteArray);
            var beforeB = Sha(reviewerB.DocumentByteArray);
            var result = Try(run);
            var moved = new List<string>();
            if (Sha(baseDoc.DocumentByteArray) != beforeBase) moved.Add("base");
            if (Sha(reviewerA.DocumentByteArray) != beforeA) moved.Add("reviewerA");
            if (Sha(reviewerB.DocumentByteArray) != beforeB) moved.Add("reviewerB");
            observed[key] = (result, reported.Count == 0 ? "none" : string.Join(",", reported),
                moved.Count == 0 ? "clean" : "MUTATED " + string.Join("+", moved));
        }

        // Order independence, N-way: ask a single caller-held consolidation for all four products in
        // REVERSE order and require each to match what its static produced — which also pins
        // CreateConsolidation against the statics, a class corpus mode otherwise never reaches.
        DocxDiffConsolidation? consolidation = null;
        string? createFailure = null;
        try { consolidation = DocxDiff.CreateConsolidation(baseDoc, reviewers, settings); }
        catch (Exception ex) { createFailure = $"FAIL {ex.GetType().Name}: {Truncate(ex.Message)}"; }

        string ViaConsolidation(Func<string> run) => createFailure ?? Try(run);
        var reorderedScript = ViaConsolidation(() =>
            Sha(Encoding.UTF8.GetBytes(consolidation!.GetConsolidatedEditScriptJson())));
        var reorderedConflicts = ViaConsolidation(() => ConflictsDigest(consolidation!.GetConflicts()));
        var reorderedRevisions = ViaConsolidation(() =>
            ConsolidatedRevisionsDigest(consolidation!.GetConsolidatedRevisions()));
        var reorderedRedline = ViaConsolidation(() =>
            StableRedlineDigest(consolidation!.Consolidate().DocumentByteArray));

        string Variance(string key, string reordered) =>
            observed[key].Result == reordered ? "stable" : $"reordered {reordered}";

        Store(sink, name, mode, "redline", observed["redline"], Variance("redline", reorderedRedline));
        Store(sink, name, mode, "revisions", observed["revisions"], Variance("revisions", reorderedRevisions));
        Store(sink, name, mode, "conflicts", observed["conflicts"], Variance("conflicts", reorderedConflicts));
        Store(sink, name, mode, "editscript", observed["editscript"], Variance("editscript", reorderedScript));
    }

    private static string ConsolidatedRevisionsDigest(IReadOnlyList<DocxDiffConsolidatedRevision> revs)
    {
        var sb = new StringBuilder().Append(revs.Count).Append('\n');
        foreach (var r in revs)
            sb.Append(r.Type).Append('|').Append(r.Author).Append('|').Append(r.Text).Append('|')
              .Append(r.LeftAnchor).Append('|').Append(r.RightAnchor).Append('|')
              .Append(r.ConflictId).Append('\n');
        return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
    }

    private static string ConflictsDigest(IReadOnlyList<DocxDiffConflict> conflicts)
    {
        var sb = new StringBuilder().Append(conflicts.Count).Append('\n');
        foreach (var c in conflicts)
            sb.Append(c.Id).Append('|').Append(c.BaseAnchor).Append('|')
              .Append(c.TokenStart).Append('-').Append(c.TokenEnd).Append('|')
              .Append(string.Join(",", c.Competitors.Select(x => x.Author + "=" + x.ResultText))).Append('\n');
        return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
    }

    /// <summary><c>clean</c>, or which side's bytes the call moved. <c>IrReader.Read</c> promises
    /// "the caller's DocumentByteArray is left byte-for-byte unchanged" and <c>PreAccept</c> promises
    /// "the input is untouched"; nothing verified either, and sharing one parsed snapshot across
    /// stages is exactly the direction that ends in mutating a caller's bytes.</summary>
    private static string MutationOf(string beforeLeft, string beforeRight, WmlDocument left, WmlDocument right)
    {
        var moved = new List<string>();
        if (Sha(left.DocumentByteArray) != beforeLeft) moved.Add("left");
        if (Sha(right.DocumentByteArray) != beforeRight) moved.Add("right");
        return moved.Count == 0 ? "clean" : "MUTATED " + string.Join("+", moved);
    }

    private static string RevisionsDigest(IReadOnlyList<DocxDiffRevision> revs)
    {
        var sb = new StringBuilder().Append(revs.Count).Append('\n');
        foreach (var r in revs)
            sb.Append(r.Type).Append('|').Append(r.Author).Append('|').Append(r.Text).Append('|')
              .Append(r.LeftAnchor).Append('|').Append(r.RightAnchor).Append('|')
              .Append(r.MoveGroupId).Append('|').Append(r.IsMoveSource).Append('\n');
        return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
    }

    /// <summary>How many observations record a non-default value on each channel. A channel that is
    /// never anything but its default across the whole corpus is coverage nobody should count.</summary>
    private static void ReportChannelCoverage(SortedDictionary<string, Observation> observations)
    {
        int warnings = 0, mutations = 0, variances = 0, warningsObservable = 0, varianceObservable = 0;
        foreach (var o in observations.Values)
        {
            if (o.Warnings is not ("n/a" or "none")) warnings++;
            if (o.Warnings != "n/a") warningsObservable++;
            if (o.InputMutation.StartsWith("MUTATED", StringComparison.Ordinal)) mutations++;
            if (o.OrderVariance != "n/a") varianceObservable++;
            if (o.OrderVariance is not ("n/a" or "stable")) variances++;
        }

        Console.WriteLine();
        Console.WriteLine("channel coverage (non-default of observable):");
        Console.WriteLine($"  warnings       {warnings,6:N0} / {warningsObservable,6:N0}");
        Console.WriteLine($"  inputMutation  {mutations,6:N0} / {observations.Count,6:N0}");
        Console.WriteLine($"  orderVariance  {variances,6:N0} / {varianceObservable,6:N0}");
    }

    // A thrown exception is a recorded outcome, not a crash: the type and message are digested so a
    // build that throws something different, or stops throwing, shows up as a mismatch.
    private static string Try(Func<string> f)
    {
        try { return f(); }
        catch (Exception ex) { return $"FAIL {ex.GetType().Name}: {Truncate(ex.Message)}"; }
    }

    private static string Truncate(string s)
    {
        s = s.ReplaceLineEndings(" ").Trim();
        return s.Length <= 160 ? s : s[..160];
    }

    private static string Sha(byte[] b) => Convert.ToHexString(SHA256.HashData(b));

    // The redline package is digested entry by entry, exactly as produced. Generated part names used
    // to be folded to a placeholder first, because #621 meant a redline that imported or created a
    // part emitted a fresh GUID per run and origin/main disagreed with itself. #623 fixed that, so
    // the fold now only makes the harness less sensitive -- and its "[PR][0-9a-f]{32}" pattern would
    // also fold a legitimate token of that shape out of XML content.
    private static string StableRedlineDigest(byte[] package)
    {
        var entries = new SortedDictionary<string, string>(StringComparer.Ordinal);
        using (var ms = new MemoryStream(package))
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
        {
            foreach (var entry in zip.Entries)
            {
                using var stream = entry.Open();
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                entries[entry.FullName] = Sha(buffer.ToArray());
            }
        }

        var sb = new StringBuilder();
        foreach (var (k, v) in entries) sb.Append(k).Append('=').Append(v).Append('\n');
        return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
    }
}
