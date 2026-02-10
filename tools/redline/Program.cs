// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.IO;
using Docxodus;

namespace Redline;

class Program
{
    const string Version = "1.1.0";

    static int Main(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help")
        {
            PrintUsage();
            return 0;
        }

        if (args[0] is "-v" or "--version")
        {
            Console.WriteLine($"redline {Version}");
            return 0;
        }

        // Partition args into positional and flags
        var positional = new System.Collections.Generic.List<string>();
        var flags = new System.Collections.Generic.List<string>();

        foreach (var arg in args)
        {
            if (arg.StartsWith("--"))
                flags.Add(arg);
            else
                positional.Add(arg);
        }

        // Detect legacy 4-arg format: redline <author> <original> <modified> <output>
        // vs new format: redline <original> <modified> <output> [--flags...]
        string authorTag;
        string originalFilePath;
        string modifiedFilePath;
        string outputFilePath;

        if (positional.Count == 4 && flags.Count == 0)
        {
            // Legacy format: redline <author> <original> <modified> <output>
            authorTag = positional[0];
            originalFilePath = positional[1];
            modifiedFilePath = positional[2];
            outputFilePath = positional[3];
        }
        else if (positional.Count == 3)
        {
            // New format: redline <original> <modified> <output> [--flags...]
            originalFilePath = positional[0];
            modifiedFilePath = positional[1];
            outputFilePath = positional[2];
            authorTag = "Redline";
        }
        else
        {
            Console.Error.WriteLine("Error: Invalid number of arguments.");
            Console.Error.WriteLine();
            PrintUsage();
            return 1;
        }

        if (!File.Exists(originalFilePath))
        {
            Console.Error.WriteLine($"Error: Original file not found: {originalFilePath}");
            return 1;
        }

        if (!File.Exists(modifiedFilePath))
        {
            Console.Error.WriteLine($"Error: Modified file not found: {modifiedFilePath}");
            return 1;
        }

        // Build settings from flags
        var settings = new WmlComparerSettings
        {
            DetailThreshold = 0
        };

        foreach (var flag in flags)
        {
            if (flag.StartsWith("--author="))
            {
                authorTag = flag["--author=".Length..];
            }
            else if (flag.StartsWith("--detail-threshold="))
            {
                if (!double.TryParse(flag["--detail-threshold=".Length..], NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
                {
                    Console.Error.WriteLine($"Error: Invalid value for --detail-threshold: {flag["--detail-threshold=".Length..]}");
                    return 1;
                }
                settings.DetailThreshold = val;
            }
            else if (flag == "--case-insensitive")
            {
                settings.CaseInsensitive = true;
            }
            else if (flag == "--detect-moves")
            {
                settings.DetectMoves = true;
            }
            else if (flag == "--simplify-move-markup")
            {
                settings.SimplifyMoveMarkup = true;
            }
            else if (flag.StartsWith("--move-similarity-threshold="))
            {
                if (!double.TryParse(flag["--move-similarity-threshold=".Length..], NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
                {
                    Console.Error.WriteLine($"Error: Invalid value for --move-similarity-threshold: {flag["--move-similarity-threshold=".Length..]}");
                    return 1;
                }
                settings.MoveSimilarityThreshold = val;
            }
            else if (flag.StartsWith("--move-minimum-word-count="))
            {
                if (!int.TryParse(flag["--move-minimum-word-count=".Length..], out var val))
                {
                    Console.Error.WriteLine($"Error: Invalid value for --move-minimum-word-count: {flag["--move-minimum-word-count=".Length..]}");
                    return 1;
                }
                settings.MoveMinimumWordCount = val;
            }
            else if (flag == "--no-detect-format-changes")
            {
                settings.DetectFormatChanges = false;
            }
            else if (flag == "--no-conflate-spaces")
            {
                settings.ConflateBreakingAndNonbreakingSpaces = false;
            }
            else if (flag.StartsWith("--date-time="))
            {
                settings.DateTimeForRevisions = flag["--date-time=".Length..];
            }
            else
            {
                Console.Error.WriteLine($"Error: Unknown flag: {flag}");
                Console.Error.WriteLine();
                PrintUsage();
                return 1;
            }
        }

        settings.AuthorForRevisions = authorTag;

        try
        {
            var originalBytes = File.ReadAllBytes(originalFilePath);
            var modifiedBytes = File.ReadAllBytes(modifiedFilePath);
            var originalDocument = new WmlDocument(originalFilePath, originalBytes);
            var modifiedDocument = new WmlDocument(modifiedFilePath, modifiedBytes);

            Console.WriteLine($"Comparing documents...");
            Console.WriteLine($"  Original: {originalFilePath}");
            Console.WriteLine($"  Modified: {modifiedFilePath}");

            var result = WmlComparer.Compare(originalDocument, modifiedDocument, settings);
            var revisions = WmlComparer.GetRevisions(result, settings);

            File.WriteAllBytes(outputFilePath, result.DocumentByteArray);

            Console.WriteLine();
            Console.WriteLine($"Redline complete: {revisions.Count} revision(s) found");
            Console.WriteLine($"  Output: {outputFilePath}");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            if (Environment.GetEnvironmentVariable("REDLINE_DEBUG") == "1")
            {
                Console.Error.WriteLine();
                Console.Error.WriteLine("Stack trace:");
                Console.Error.WriteLine(ex.StackTrace);
            }
            return 1;
        }
    }

    static void PrintUsage()
    {
        Console.WriteLine($"redline {Version} - Compare Word documents and generate redline diffs");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  redline <original.docx> <modified.docx> <output.docx> [options]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  original.docx    Path to the original document");
        Console.WriteLine("  modified.docx    Path to the modified document");
        Console.WriteLine("  output.docx      Path for the output redline document");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --author=<name>                   Author name for tracked changes (default: Redline)");
        Console.WriteLine("  --detail-threshold=<0.0-1.0>      Comparison granularity (lower = more detailed, default: 0)");
        Console.WriteLine("  --case-insensitive                 Ignore case differences");
        Console.WriteLine("  --detect-moves                     Enable move detection");
        Console.WriteLine("  --simplify-move-markup             Convert moves to del/ins for Word compatibility");
        Console.WriteLine("  --move-similarity-threshold=<val>  Jaccard threshold for move matching (default: 0.8)");
        Console.WriteLine("  --move-minimum-word-count=<val>    Min words for move detection (default: 3)");
        Console.WriteLine("  --no-detect-format-changes         Disable formatting change detection");
        Console.WriteLine("  --no-conflate-spaces               Distinguish breaking/non-breaking spaces");
        Console.WriteLine("  --date-time=<ISO8601>              Custom timestamp for revisions");
        Console.WriteLine("  -h, --help                         Show this help message");
        Console.WriteLine("  -v, --version                      Show version information");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  redline contract-v1.docx contract-v2.docx redline.docx");
        Console.WriteLine("  redline draft.docx final.docx changes.docx --author=\"Legal Review\"");
        Console.WriteLine("  redline old.docx new.docx diff.docx --detect-moves --simplify-move-markup");
        Console.WriteLine("  redline old.docx new.docx diff.docx --detail-threshold=0.5 --case-insensitive");
        Console.WriteLine();
        Console.WriteLine("Environment Variables:");
        Console.WriteLine("  REDLINE_DEBUG=1  Show detailed error information");
    }
}
