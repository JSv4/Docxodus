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

        var digests = new ConcurrentDictionary<string, string>(StringComparer.Ordinal);
        var done = 0;
        var errors = 0;

        Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = threads }, path =>
        {
            var name = Path.GetRelativePath(root, path).Replace('\\', '/');
            byte[] bytes;
            try { bytes = File.ReadAllBytes(path); }
            catch (Exception ex) { digests[$"{name}#read"] = $"READ-FAIL {ex.GetType().Name}"; return; }

            // Every document is compared against a deterministically edited copy of itself, and
            // against itself unchanged (the identical-bytes fast paths).
            byte[] edited;
            try { edited = CorpusVariant.Edit(bytes); }
            catch (Exception ex) { digests[$"{name}#variant"] = $"VARIANT-FAIL {ex.GetType().Name}"; edited = bytes; }

            // A second, differently-offset edit of the same document: the N-way case needs two
            // reviewers who actually disagree, or the merger never has a competitor to order.
            byte[] editedAlt;
            try { editedAlt = CorpusVariant.Edit(bytes, offset: 2); }
            catch (Exception ex) { digests[$"{name}#variant-alt"] = $"VARIANT-FAIL {ex.GetType().Name}"; editedAlt = bytes; }

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

            // The compatibility pre-flight is a SECOND observable output, and it is invisible to
            // everything above: with it disengaged, a product that never runs it and a product that
            // runs it and finds nothing produce the same digest. Digest the report itself, so a
            // shortcut that skips the pre-flight rather than the work shows up as a mismatch — on the
            // identical pair as well, where a shortcut is most tempting and least observable.
            RecordPreflight(digests, name, "edited", left, right);
            RecordPreflight(digests, name, "identical", left, same);

            // N-way. The consolidate path has its own reader fan-out, its own merger and its own
            // markup renderer, and NONE of the three products above touches any of them. Two
            // reviewers off the same base is the smallest shape that exercises reviewer ordering and
            // conflict competitor order.
            RecordConsolidate(digests, name, left, right, alt);

            var n = Interlocked.Increment(ref done);
            if (n % 50 == 0) Console.WriteLine($"  {n,5}/{files.Count} ...");
        });

        foreach (var kv in digests)
            if (kv.Value.Contains("FAIL", StringComparison.Ordinal)) errors++;

        var sorted = new SortedDictionary<string, string>(digests, StringComparer.Ordinal);
        File.WriteAllText(outPath, JsonSerializer.Serialize(sorted, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine();
        Console.WriteLine($"{sorted.Count:N0} digests over {files.Count:N0} documents written to {outPath}");
        Console.WriteLine($"({errors:N0} of them record a thrown exception rather than a result — that is expected for");
        Console.WriteLine(" malformed or unsupported fixtures, and a CHANGE in them is still a regression.)");

        if (checkPath is null) return 0;

        var expected = JsonSerializer.Deserialize<SortedDictionary<string, string>>(File.ReadAllText(checkPath))!;
        var mismatches = 0;
        foreach (var (k, want) in expected)
        {
            if (!sorted.TryGetValue(k, out var got)) { Console.WriteLine($"[parity] MISSING  {k}"); mismatches++; }
            else if (got != want)
            {
                Console.WriteLine($"[parity] MISMATCH {k}");
                Console.WriteLine($"           expected {want}");
                Console.WriteLine($"           actual   {got}");
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

    private static void Record(
        ConcurrentDictionary<string, string> sink, string name, string mode,
        WmlDocument left, WmlDocument right, DocxDiffSettings settings)
    {
        sink[$"{name}#{mode}/redline"] = Try(() => StableRedlineDigest(DocxDiff.Compare(left, right, settings).DocumentByteArray));
        sink[$"{name}#{mode}/revisions"] = Try(() =>
        {
            var revs = DocxDiff.GetRevisions(left, right, settings);
            var sb = new StringBuilder().Append(revs.Count).Append('\n');
            foreach (var r in revs)
                sb.Append(r.Type).Append('|').Append(r.Author).Append('|').Append(r.Text).Append('|')
                  .Append(r.LeftAnchor).Append('|').Append(r.RightAnchor).Append('|')
                  .Append(r.MoveGroupId).Append('|').Append(r.IsMoveSource).Append('\n');
            return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
        });
        sink[$"{name}#{mode}/editscript"] = Try(() => Sha(Encoding.UTF8.GetBytes(DocxDiff.GetEditScriptJson(left, right, settings))));
    }

    // Every consolidate product, over two reviewers off one base. Reviewer order is significant to
    // conflict reporting, so the digest carries the authors and the per-revision order as emitted.
    private static void RecordConsolidate(
        ConcurrentDictionary<string, string> sink, string name,
        WmlDocument baseDoc, WmlDocument reviewerA, WmlDocument reviewerB)
    {
        var reviewers = new List<DocxDiffReviewer>
        {
            new() { Author = "Reviewer A", Document = reviewerA },
            new() { Author = "Reviewer B", Document = reviewerB },
        };
        var settings = new DocxDiffConsolidateSettings();

        sink[$"{name}#consolidate/redline"] = Try(() =>
            StableRedlineDigest(DocxDiff.Consolidate(baseDoc, reviewers, settings).DocumentByteArray));

        sink[$"{name}#consolidate/revisions"] = Try(() =>
        {
            var revs = DocxDiff.GetConsolidatedRevisions(baseDoc, reviewers, settings);
            var sb = new StringBuilder().Append(revs.Count).Append('\n');
            foreach (var r in revs)
                sb.Append(r.Type).Append('|').Append(r.Author).Append('|').Append(r.Text).Append('|')
                  .Append(r.LeftAnchor).Append('|').Append(r.RightAnchor).Append('|')
                  .Append(r.ConflictId).Append('\n');
            return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
        });

        sink[$"{name}#consolidate/conflicts"] = Try(() =>
        {
            var conflicts = DocxDiff.GetConflicts(baseDoc, reviewers, settings);
            var sb = new StringBuilder().Append(conflicts.Count).Append('\n');
            foreach (var c in conflicts)
                sb.Append(c.Id).Append('|').Append(c.BaseAnchor).Append('|')
                  .Append(c.TokenStart).Append('-').Append(c.TokenEnd).Append('|')
                  .Append(string.Join(",", c.Competitors.Select(x => x.Author + "=" + x.ResultText))).Append('\n');
            return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
        });

        sink[$"{name}#consolidate/editscript"] = Try(() =>
            Sha(Encoding.UTF8.GetBytes(DocxDiff.GetConsolidatedEditScriptJson(baseDoc, reviewers, settings))));
    }

    // The compatibility report each product hands back through the OnCompatibilityWarning callback,
    // digested per product. A product that stops running the pre-flight records "none" where it used
    // to record a report — which no output digest can tell you, because the output was already right.
    private static void RecordPreflight(
        ConcurrentDictionary<string, string> sink, string name, string mode,
        WmlDocument left, WmlDocument right)
    {
        foreach (var (product, run) in PreflightProducts(left, right))
            sink[$"{name}#{mode}/preflight-{product}"] = Try(() =>
            {
                var seen = new List<string>();
                var settings = new DocxDiffSettings
                {
                    OnCompatibilityWarning = report =>
                    {
                        foreach (var w in report.Warnings.OrderBy(w => w.Feature.Id, StringComparer.Ordinal))
                            seen.Add(w.Feature.Id);
                    },
                };
                run(settings);
                return seen.Count == 0 ? "none" : string.Join(",", seen);
            });
    }

    private static IEnumerable<(string Product, Action<DocxDiffSettings> Run)> PreflightProducts(
        WmlDocument left, WmlDocument right)
    {
        yield return ("redline", s => DocxDiff.Compare(left, right, s));
        yield return ("revisions", s => DocxDiff.GetRevisions(left, right, s));
        yield return ("editscript", s => DocxDiff.GetEditScriptJson(left, right, s));
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

    // Media and diagram parts imported into a redline are named "P" + a fresh GUID, and the
    // relationships pointing at them get ids of "R" + a fresh GUID. Two runs of the SAME comparison
    // on the SAME build therefore emit packages that differ in those part names, in the relationship
    // ids and targets, and in the r:embed/r:id references in the story parts — while the bytes they
    // hold are identical. That is a real defect, but a PRE-EXISTING one:
    // origin/main disagrees with itself on exactly the documents this affects. Digesting raw package
    // bytes would therefore report 161 differences per run and drown any genuine regression on the
    // 54 media-bearing documents in the corpus.
    //
    // So the redline is digested rename-invariantly: every generated part name is folded to a
    // placeholder, in entry names AND inside XML content, and entries are hashed in canonical-name
    // order. Content still has to match exactly — this hides only the naming churn, so a genuine
    // change to any part's bytes, to the set of parts, or to which content sits under which
    // canonical name still shows up as a mismatch.
    // "P<32 hex>" for a generated part name, "R<32 hex>" for a generated relationship id.
    private static readonly System.Text.RegularExpressions.Regex GeneratedPartName =
        new("[PR][0-9a-f]{32}", System.Text.RegularExpressions.RegexOptions.Compiled);

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
                var content = buffer.ToArray();

                var canonicalName = GeneratedPartName.Replace(entry.FullName, "@");
                var isXml = entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                         || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
                var digest = isXml
                    ? Sha(Encoding.UTF8.GetBytes(GeneratedPartName.Replace(Encoding.UTF8.GetString(content), "@")))
                    : Sha(content);

                // Two parts can fold to one canonical name (P<guid1>.jpeg and P<guid2>.jpeg). Their
                // content digests are combined in sorted order so the fold stays order-independent.
                entries[canonicalName] = entries.TryGetValue(canonicalName, out var existing)
                    ? string.Join(",", new[] { existing, digest }.OrderBy(x => x, StringComparer.Ordinal))
                    : digest;
            }
        }

        var sb = new StringBuilder();
        foreach (var (k, v) in entries) sb.Append(k).Append('=').Append(v).Append('\n');
        return Sha(Encoding.UTF8.GetBytes(sb.ToString()));
    }
}
