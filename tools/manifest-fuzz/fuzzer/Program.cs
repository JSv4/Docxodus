#nullable enable
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Docxodus.Verification;

// Fuzz harness for PackageManifestGenerator.Generate().
// Oracle: never throws, deterministic, non-mutating, canonical JSON always parses.
// Feedback signal: manifest "features" (packageKind + finding codes + digest shape) drive
// corpus growth, since each finding code corresponds to a distinct parser branch.

string seedsDir = "", outDir = "";
int workerId = 0, seconds = 60;
int rngSeed = 12345;
for (int a = 0; a < args.Length - 1; a++)
{
    switch (args[a])
    {
        case "--seeds": seedsDir = args[++a]; break;
        case "--out": outDir = args[++a]; break;
        case "--id": workerId = int.Parse(args[++a]); break;
        case "--seconds": seconds = int.Parse(args[++a]); break;
        case "--rngseed": rngSeed = int.Parse(args[++a]); break;
    }
}
string corpusDir = Path.Combine(outDir, "corpus");
string crashDir = Path.Combine(outDir, "crashes");
string hangDir = Path.Combine(outDir, "hangs");
string bugDir = Path.Combine(outDir, "bugs");
Directory.CreateDirectory(corpusDir);
Directory.CreateDirectory(crashDir);
Directory.CreateDirectory(hangDir);
Directory.CreateDirectory(bugDir);

const int MaxInput = 256 * 1024;
var rng = new Random(rngSeed);
var corpus = new List<byte[]>();
var loadedNames = new HashSet<string>();
var features = new HashSet<string>();
var crashKeys = new HashSet<string>();
long execs = 0, crashes = 0, bugs = 0, added = 0;
int corpusCounter = 0;

foreach (var f in Directory.GetFiles(seedsDir))
{
    var b = File.ReadAllBytes(f);
    if (b.Length <= MaxInput) corpus.Add(b);
}
if (corpus.Count == 0) { Console.Error.WriteLine("no seeds"); return 2; }

var tiny = new PackageManifestOptions
{
    MaxEntryCount = 4,
    MaxTotalUncompressedBytes = 4096,
    MaxXmlPartBytes = 512,
    MaxCompressionRatio = 3,
    MaxUriLength = 24,
};

byte[][] dict =
{
    new byte[] { 0x50, 0x4B, 0x03, 0x04 },
    new byte[] { 0x50, 0x4B, 0x01, 0x02 },
    new byte[] { 0x50, 0x4B, 0x05, 0x06 },
    new byte[] { 0x50, 0x4B, 0x06, 0x06 },
    new byte[] { 0x50, 0x4B, 0x06, 0x07 },
    new byte[] { 0x50, 0x4B, 0x07, 0x08 },
    new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 },
    Encoding.ASCII.GetBytes("[Content_Types].xml"),
    Encoding.ASCII.GetBytes("_rels/.rels"),
    Encoding.ASCII.GetBytes("word/document.xml"),
    Encoding.ASCII.GetBytes(".rels"),
    Encoding.ASCII.GetBytes("%2e"),
    Encoding.ASCII.GetBytes("%2f"),
    Encoding.ASCII.GetBytes("%25"),
    Encoding.ASCII.GetBytes("%C3%A9"),
    Encoding.ASCII.GetBytes("../"),
    Encoding.ASCII.GetBytes("..\\"),
    Encoding.ASCII.GetBytes("xml:space=\"preserve\""),
    Encoding.ASCII.GetBytes("TargetMode=\"External\""),
    Encoding.ASCII.GetBytes("<!DOCTYPE"),
    Encoding.ASCII.GetBytes("<?xml"),
    Encoding.ASCII.GetBytes("<Relationship "),
    Encoding.ASCII.GetBytes("<Override PartName=\""),
    Encoding.ASCII.GetBytes("<Default Extension=\""),
    Encoding.ASCII.GetBytes("EncryptedPackage"),
    Encoding.ASCII.GetBytes("EncryptionInfo"),
    Encoding.Unicode.GetBytes("EncryptedPackage"),
    Encoding.Unicode.GetBytes("Root Entry"),
    Encoding.ASCII.GetBytes("application/vnd.openxmlformats-"),
    Encoding.ASCII.GetBytes("ContentType=\""),
    Encoding.ASCII.GetBytes("w:ins"),
    Encoding.ASCII.GetBytes("w:del"),
    Encoding.ASCII.GetBytes("r:embed"),
    Encoding.ASCII.GetBytes("xsi:type"),
    Encoding.ASCII.GetBytes("mc:Ignorable"),
    Encoding.ASCII.GetBytes("urn:"),
    Encoding.ASCII.GetBytes("file:///"),
};
ulong[] interesting = { 0, 1, 2, 0x7F, 0x80, 0xFF, 0x100, 0x1FF, 0x200, 0xFFF, 0x1000,
    0xFFFF, 0x10000, 0x7FFFFFFF, 0xFFFFFFFD, 0xFFFFFFFE, 0xFFFFFFFF, 0x100000000,
    0x7FFFFFFFFFFFFFFF, 0xFFFFFFFFFFFFFFFF };

// hang watchdog
long heartbeat = Stopwatch.GetTimestamp();
byte[]? inflight = null;
long execsSnapshot = 0;
var watchdog = new Thread(() =>
{
    while (true)
    {
        Thread.Sleep(2000);
        var startedAt = Interlocked.Read(ref heartbeat);
        if (startedAt != 0 && Stopwatch.GetElapsedTime(startedAt).TotalSeconds > 20)
        {
            var snap = inflight;
            if (snap != null)
                File.WriteAllBytes(Path.Combine(hangDir, $"hang-w{workerId}-{Interlocked.Read(ref execsSnapshot)}.bin"), snap);
            Console.Error.WriteLine("HANG detected - input dumped");
            Environment.FailFast("fuzz hang");
        }
    }
}) { IsBackground = true };
watchdog.Start();

var sw = Stopwatch.StartNew();
var lastStats = TimeSpan.Zero;
var lastImport = TimeSpan.Zero;

byte[] Mutate(byte[] seed)
{
    var buf = new List<byte>(seed);
    int rounds = 1 << rng.Next(0, 5);
    for (int i = 0; i < rounds; i++)
    {
        if (buf.Count == 0) { buf.AddRange(dict[rng.Next(dict.Length)]); continue; }
        int off = rng.Next(buf.Count);
        // bias 1 in 4 mutations toward the tail where ZIP metadata lives
        if (rng.Next(4) == 0 && buf.Count > 64) off = buf.Count - 1 - rng.Next(64);
        switch (rng.Next(11))
        {
            case 0: buf[off] ^= (byte)(1 << rng.Next(8)); break;
            case 1: buf[off] = (byte)rng.Next(256); break;
            case 2: buf[off] = (byte)(buf[off] + rng.Next(-16, 17)); break;
            case 3:
            {
                var v = interesting[rng.Next(interesting.Length)];
                int width = rng.Next(3) switch { 0 => 2, 1 => 4, _ => 8 };
                for (int k = 0; k < width && off + k < buf.Count; k++)
                    buf[off + k] = (byte)(v >> (8 * k));
                break;
            }
            case 4:
            {
                int len = 1 + rng.Next(Math.Max(1, buf.Count / 4));
                len = Math.Min(len, buf.Count - off);
                buf.RemoveRange(off, len);
                break;
            }
            case 5:
            {
                int len = 1 + rng.Next(Math.Max(1, Math.Min(buf.Count, 4096)));
                int src = rng.Next(Math.Max(1, buf.Count - len));
                var chunk = buf.GetRange(src, Math.Min(len, buf.Count - src));
                buf.InsertRange(rng.Next(buf.Count + 1), chunk);
                break;
            }
            case 6:
            {
                var other = corpus[rng.Next(corpus.Count)];
                if (other.Length == 0) break;
                int len = 1 + rng.Next(Math.Min(other.Length, 4096));
                int src = rng.Next(Math.Max(1, other.Length - len));
                var chunk = new byte[Math.Min(len, other.Length - src)];
                Array.Copy(other, src, chunk, 0, chunk.Length);
                if (rng.Next(2) == 0) buf.InsertRange(rng.Next(buf.Count + 1), chunk);
                else
                    for (int k = 0; k < chunk.Length && off + k < buf.Count; k++)
                        buf[off + k] = chunk[k];
                break;
            }
            case 7: buf.InsertRange(rng.Next(buf.Count + 1), dict[rng.Next(dict.Length)]); break;
            case 8:
            {
                var tok = dict[rng.Next(dict.Length)];
                for (int k = 0; k < tok.Length && off + k < buf.Count; k++)
                    buf[off + k] = tok[k];
                break;
            }
            case 9:
                if (rng.Next(2) == 0) buf.RemoveRange(off, buf.Count - off);
                else
                {
                    int n = 1 + rng.Next(256);
                    for (int k = 0; k < n; k++)
                        buf.Add(rng.Next(3) == 0 ? (byte)rng.Next(256) : (byte)0);
                }
                break;
            case 10:
            {
                // zero/0xFF a whole aligned 4-byte field
                int f4 = (off / 4) * 4;
                byte v = rng.Next(2) == 0 ? (byte)0 : (byte)0xFF;
                for (int k = 0; k < 4 && f4 + k < buf.Count; k++) buf[f4 + k] = v;
                break;
            }
        }
        if (buf.Count > MaxInput) buf.RemoveRange(MaxInput, buf.Count - MaxInput);
    }
    return buf.ToArray();
}

string Feature(PackageManifest m)
{
    var codes = string.Join(",", m.Findings.Select(f => f.Code).Distinct().OrderBy(c => c, StringComparer.Ordinal));
    return $"{m.PackageKind}|{m.IsValid}|{(m.OrderedOpcContentDigest is null ? 0 : 1)}{(m.NormalizedSemanticDigest is null ? 0 : 1)}|{codes}";
}

void RecordCrash(byte[] input, Exception ex, string profile)
{
    crashes++;
    var frame = ex.StackTrace?.Split('\n').FirstOrDefault(l => l.Contains("Docxodus"))?.Trim() ?? "?";
    var key = $"{ex.GetType().FullName}|{frame}";
    var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(key)))[..12];
    var dir = Path.Combine(crashDir, hash);
    if (crashKeys.Add(key) && !Directory.Exists(dir))
    {
        Directory.CreateDirectory(dir);
        File.WriteAllBytes(Path.Combine(dir, "input.bin"), input);
        File.WriteAllText(Path.Combine(dir, "info.txt"),
            $"profile: {profile}\ninput bytes: {input.Length}\nkey: {key}\n\n{ex}");
        Console.WriteLine($"[w{workerId}] NEW CRASH {hash}: {key}");
    }
}

void RecordBug(byte[] input, string kind, string detail)
{
    bugs++;
    var hash = Convert.ToHexString(SHA256.HashData(input))[..12];
    var dir = Path.Combine(bugDir, $"{kind}-{hash}");
    if (!Directory.Exists(dir))
    {
        Directory.CreateDirectory(dir);
        File.WriteAllBytes(Path.Combine(dir, "input.bin"), input);
        File.WriteAllText(Path.Combine(dir, "info.txt"), detail);
        Console.WriteLine($"[w{workerId}] NEW BUG {kind}-{hash}");
    }
}

void DeepCheck(byte[] input, PackageManifestOptions? profile, PackageManifest first)
{
    var copy = (byte[])input.Clone();
    byte[] j1, j2;
    try
    {
        j1 = first.ToJsonBytes();
        var again = PackageManifestGenerator.Generate(input, profile);
        j2 = again.ToJsonBytes();
    }
    catch (Exception ex)
    {
        RecordCrash(input, ex, profile is null ? "default-deep" : "tiny-deep");
        return;
    }
    if (!j1.AsSpan().SequenceEqual(j2))
        RecordBug(input, "nondeterminism", $"profile: {(profile is null ? "default" : "tiny")}\nfirst:\n{Encoding.UTF8.GetString(j1)}\n\nsecond:\n{Encoding.UTF8.GetString(j2)}");
    if (!input.AsSpan().SequenceEqual(copy))
        RecordBug(input, "input-mutated", "Generate() modified the caller's byte array");
    try { using var _ = JsonDocument.Parse(j1); }
    catch (Exception ex) { RecordBug(input, "bad-json", ex.ToString()); }
    try { first.ToJson(indented: true); }
    catch (Exception ex) { RecordBug(input, "indented-throw", ex.ToString()); }
}

while (sw.Elapsed.TotalSeconds < seconds)
{
    var seed = corpus[rng.Next(corpus.Count)];
    var input = Mutate(seed);
    var profile = (execs & 1) == 0 ? null : tiny;
    string profileName = profile is null ? "default" : "tiny";

    inflight = input;
    Interlocked.Exchange(ref execsSnapshot, execs);
    Interlocked.Exchange(ref heartbeat, Stopwatch.GetTimestamp());
    PackageManifest manifest;
    try { manifest = PackageManifestGenerator.Generate(input, profile); }
    catch (Exception ex)
    {
        Interlocked.Exchange(ref heartbeat, 0);
        RecordCrash(input, ex, profileName);
        execs++;
        continue;
    }
    Interlocked.Exchange(ref heartbeat, 0);
    execs++;

    var feat = Feature(manifest);
    bool isNew = features.Add(feat);
    if (isNew)
    {
        if (corpus.Count >= 8192) corpus[rng.Next(corpus.Count)] = input;
        else corpus.Add(input);
        added++;
        if (corpusCounter < 2000)
        {
            var name = $"w{workerId}-{corpusCounter++}.bin";
            File.WriteAllBytes(Path.Combine(corpusDir, name), input);
            loadedNames.Add(name);
        }
        DeepCheck(input, profile, manifest);
    }
    else if ((execs & 127) == 0)
    {
        DeepCheck(input, profile, manifest);
    }

    if (sw.Elapsed - lastImport > TimeSpan.FromSeconds(15) && corpus.Count < 8192)
    {
        lastImport = sw.Elapsed;
        foreach (var f in Directory.GetFiles(corpusDir))
        {
            var name = Path.GetFileName(f);
            if (loadedNames.Add(name))
            {
                try
                {
                    var b = File.ReadAllBytes(f);
                    if (b.Length <= MaxInput) corpus.Add(b);
                }
                catch (IOException) { loadedNames.Remove(name); }
            }
        }
    }
    if (sw.Elapsed - lastStats > TimeSpan.FromSeconds(10))
    {
        lastStats = sw.Elapsed;
        Console.WriteLine($"[w{workerId}] {sw.Elapsed:hh\\:mm\\:ss} execs={execs} ({execs / Math.Max(1, sw.Elapsed.TotalSeconds):F0}/s) corpus={corpus.Count} features={features.Count} crashes={crashes} bugs={bugs}");
    }
}
Console.WriteLine($"[w{workerId}] DONE execs={execs} corpus={corpus.Count} features={features.Count} crashes={crashes} bugs={bugs} elapsed={sw.Elapsed}");
return 0;
