using Docxodus.Verification;
using SharpFuzz;

// AFL++ persistent-mode harness. The manifest contract says Generate() never throws for
// arbitrary bytes with valid options, so any escaped exception is a genuine crash signal.
Fuzzer.Run(stream =>
{
    using var ms = new MemoryStream();
    stream.CopyTo(ms);
    PackageManifestGenerator.Generate(ms.ToArray());
});
