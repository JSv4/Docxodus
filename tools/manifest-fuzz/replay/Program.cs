using System.Text;
using System.Text.Json;
using Docxodus.Verification;

// Replay every discovered input through the full oracle under both option profiles.
var tiny = new PackageManifestOptions
{
    MaxEntryCount = 4, MaxTotalUncompressedBytes = 4096, MaxXmlPartBytes = 512,
    MaxCompressionRatio = 3, MaxUriLength = 24,
};
var profiles = new (string Name, PackageManifestOptions? Opts)[] { ("default", null), ("tiny", tiny) };
long files = 0, failures = 0;
foreach (var dir in args)
{
    if (!Directory.Exists(dir)) continue;
    foreach (var f in Directory.EnumerateFiles(dir))
    {
        if (Path.GetFileName(f) == "README.txt") continue;
        byte[] input;
        try { input = File.ReadAllBytes(f); } catch (IOException) { continue; }
        files++;
        var copy = (byte[])input.Clone();
        foreach (var (name, opts) in profiles)
        {
            try
            {
                var m1 = PackageManifestGenerator.Generate(input, opts);
                var j1 = m1.ToJsonBytes();
                var j2 = PackageManifestGenerator.Generate(input, opts).ToJsonBytes();
                if (!j1.AsSpan().SequenceEqual(j2))
                { failures++; Console.WriteLine($"NONDETERMINISM {name} {f}"); }
                if (!input.AsSpan().SequenceEqual(copy))
                { failures++; Console.WriteLine($"INPUT-MUTATED {name} {f}"); }
                using var _ = JsonDocument.Parse(j1);
                m1.ToJson(indented: true);
            }
            catch (Exception ex)
            { failures++; Console.WriteLine($"THROW {name} {f}: {ex.GetType().Name}: {ex.Message}"); }
        }
        if (files % 5000 == 0) Console.WriteLine($"... {files} files");
    }
}
Console.WriteLine($"REPLAY DONE files={files} failures={failures}");
return failures == 0 ? 0 : 1;
