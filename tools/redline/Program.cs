// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using OpenXmlPowerTools;

namespace Redline;

class Program
{
    const string Version = "1.0.0";

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

        if (args.Length < 3 || args.Length > 4)
        {
            Console.Error.WriteLine("Error: Invalid number of arguments.");
            Console.Error.WriteLine();
            PrintUsage();
            return 1;
        }

        // Parse arguments: redline <original> <modified> <output> [--author=<name>]
        // Or legacy format: redline <author> <original> <modified> <output>
        string authorTag;
        string originalFilePath;
        string modifiedFilePath;
        string outputFilePath;

        if (args.Length == 4 && !args[3].StartsWith("--"))
        {
            // Legacy format: redline <author> <original> <modified> <output>
            authorTag = args[0];
            originalFilePath = args[1];
            modifiedFilePath = args[2];
            outputFilePath = args[3];
        }
        else
        {
            // New format: redline <original> <modified> <output> [--author=<name>]
            originalFilePath = args[0];
            modifiedFilePath = args[1];
            outputFilePath = args[2];
            authorTag = "Redline";

            if (args.Length == 4 && args[3].StartsWith("--author="))
            {
                authorTag = args[3]["--author=".Length..];
            }
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

        try
        {
            var originalBytes = File.ReadAllBytes(originalFilePath);
            var modifiedBytes = File.ReadAllBytes(modifiedFilePath);
            var originalDocument = new WmlDocument(originalFilePath, originalBytes);
            var modifiedDocument = new WmlDocument(modifiedFilePath, modifiedBytes);

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = authorTag,
                DetailThreshold = 0
            };

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
        Console.WriteLine("  redline <original.docx> <modified.docx> <output.docx> [--author=<name>]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  original.docx    Path to the original document");
        Console.WriteLine("  modified.docx    Path to the modified document");
        Console.WriteLine("  output.docx      Path for the output redline document");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --author=<name>  Author name for tracked changes (default: Redline)");
        Console.WriteLine("  -h, --help       Show this help message");
        Console.WriteLine("  -v, --version    Show version information");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  redline contract-v1.docx contract-v2.docx redline.docx");
        Console.WriteLine("  redline draft.docx final.docx changes.docx --author=\"Legal Review\"");
        Console.WriteLine();
        Console.WriteLine("Environment Variables:");
        Console.WriteLine("  REDLINE_DEBUG=1  Show detailed error information");
    }
}
