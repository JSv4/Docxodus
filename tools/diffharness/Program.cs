#nullable enable
using DiffHarness;

if (args.Length == 0)
{
    Console.Error.WriteLine(
        "usage:\n" +
        "  diffharness gen <base.docx> <outRootDir>            generate the scenario corpus\n" +
        "  diffharness diff <left.docx> <right.docx> <outDir> [author]   run DocxDiff + round-trip\n");
    return 2;
}

switch (args[0])
{
    case "gen":
    {
        if (args.Length < 3) { Console.Error.WriteLine("gen needs <base.docx> <outRootDir>"); return 2; }
        var manifest = Scenarios.Generate(args[1], args[2]);
        Console.WriteLine($"generated {manifest} scenarios into {args[2]}");
        return 0;
    }
    case "diff":
    {
        if (args.Length < 4) { Console.Error.WriteLine("diff needs <left> <right> <outDir> [author]"); return 2; }
        var author = args.Length > 4 ? args[4] : "Docxodus";
        var report = DiffRunner.Run(args[1], args[2], args[3], author);
        Console.WriteLine(
            $"content-clean:{report.ContentClean}  " +
            $"body(acc/rej):{report.AcceptBodyEqualsRight}/{report.RejectBodyEqualsLeft}  " +
            $"notes(acc/rej):{report.AcceptNotesEqualRight}/{report.RejectNotesEqualLeft}  " +
            $"hdrftr(acc/rej):{report.AcceptHdrFtrSetEqualsRight}/{report.RejectHdrFtrSetEqualsLeft}  " +
            $"revs(fine/compat):{report.RevisionCountFine}/{report.RevisionCountCompat}  " +
            $"hdrftrParts ours/orig:{report.HdrFtrPartsOurs}/{report.HdrFtrPartsOriginal}");
        if (report.AcceptBodyFirstDiff is { } a) Console.WriteLine("  ACCEPT BODY DIFF: " + a);
        if (report.RejectBodyFirstDiff is { } rj) Console.WriteLine("  REJECT BODY DIFF: " + rj);
        return report.ContentClean ? 0 : 1;
    }
    case "diffall":
    {
        if (args.Length < 2) { Console.Error.WriteLine("diffall needs <corpusDir> [author]"); return 2; }
        var author = args.Length > 2 ? args[2] : "Docxodus";
        var corpus = args[1];
        var manifestPath = Path.Combine(corpus, "manifest.json");
        using var mdoc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(manifestPath));
        int fails = 0, total = 0;
        Console.WriteLine($"{"scenario",-26} {"body",9} {"notes",9} {"hdrftr",9} {"fine",5} {"hf#",7}");
        foreach (var el in mdoc.RootElement.EnumerateArray())
        {
            var id = el.GetProperty("id").GetString()!;
            var dir = Path.Combine(corpus, id);
            var r = DiffRunner.Run(Path.Combine(dir, "left.docx"), Path.Combine(dir, "right.docx"), dir, author);
            total++;
            if (!r.ContentClean) fails++;
            string Pair(bool a, bool b) => $"{(a ? "Y" : "n")}/{(b ? "Y" : "n")}";
            Console.WriteLine(
                $"{id,-26} {Pair(r.AcceptBodyEqualsRight, r.RejectBodyEqualsLeft),9} " +
                $"{Pair(r.AcceptNotesEqualRight, r.RejectNotesEqualLeft),9} " +
                $"{Pair(r.AcceptHdrFtrSetEqualsRight, r.RejectHdrFtrSetEqualsLeft),9} " +
                $"{r.RevisionCountFine,5} {r.HdrFtrPartsOurs + "/" + r.HdrFtrPartsOriginal,7}" +
                $"{(r.ContentClean ? "" : "  <-- FAIL")}");
        }
        Console.WriteLine($"\n{total - fails}/{total} content-clean; {fails} failed. " +
            "(hf# = header/footer part count ours/original; >orig = duplicate-part bloat)");
        return fails == 0 ? 0 : 1;
    }
    default:
        Console.Error.WriteLine($"unknown command '{args[0]}'");
        return 2;
}
