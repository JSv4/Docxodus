// Issue #96 Stress Test
// Creates complex documents with dozens of moves and hundreds of changes
// to thoroughly validate the ID uniqueness fix

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using Docxodus;

class StressTest
{
    static readonly Random Rng = new Random(42); // Fixed seed for reproducibility
    static XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static void Main()
    {
        Console.WriteLine("╔════════════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║              ISSUE #96 STRESS TEST                                 ║");
        Console.WriteLine("║     Dozens of moves + Hundreds of changes                          ║");
        Console.WriteLine("╚════════════════════════════════════════════════════════════════════╝\n");

        // Test 1: Large document with many moves
        Console.WriteLine("═══════════════════════════════════════════════════════════════════════");
        Console.WriteLine("TEST 1: 50 paragraphs, ~15 moves, ~30 other changes");
        Console.WriteLine("═══════════════════════════════════════════════════════════════════════\n");
        RunStressTest("StressTest1", 50, 15, 30);

        // Test 2: Even larger with more chaos
        Console.WriteLine("\n═══════════════════════════════════════════════════════════════════════");
        Console.WriteLine("TEST 2: 100 paragraphs, ~25 moves, ~50 other changes");
        Console.WriteLine("═══════════════════════════════════════════════════════════════════════\n");
        RunStressTest("StressTest2", 100, 25, 50);

        // Test 3: Maximum chaos
        Console.WriteLine("\n═══════════════════════════════════════════════════════════════════════");
        Console.WriteLine("TEST 3: 200 paragraphs, ~40 moves, ~100 other changes");
        Console.WriteLine("═══════════════════════════════════════════════════════════════════════\n");
        RunStressTest("StressTest3", 200, 40, 100);

        Console.WriteLine("\n════════════════════════════════════════════════════════════════════════");
        Console.WriteLine("                    ALL STRESS TESTS COMPLETE");
        Console.WriteLine("════════════════════════════════════════════════════════════════════════\n");
    }

    static void RunStressTest(string name, int paragraphCount, int moveCount, int changeCount)
    {
        Console.WriteLine($"Generating {paragraphCount} paragraphs...");

        // Generate original document with numbered paragraphs containing unique content
        var originalParagraphs = Enumerable.Range(1, paragraphCount)
            .Select(i => GenerateParagraph(i))
            .ToList();

        // Create modified version with moves and changes
        var modifiedParagraphs = new List<string>(originalParagraphs);

        // Track what we're doing for reporting
        var moves = new List<(int from, int to, string text)>();
        var deletions = new List<int>();
        var insertions = new List<int>();
        var modifications = new List<int>();

        // Apply moves: pick random paragraphs and move them to new positions
        Console.WriteLine($"Applying ~{moveCount} moves...");
        var availableForMove = Enumerable.Range(0, modifiedParagraphs.Count).ToList();
        for (int i = 0; i < moveCount && availableForMove.Count > 2; i++)
        {
            int fromIdx = availableForMove[Rng.Next(availableForMove.Count)];
            availableForMove.Remove(fromIdx);

            var para = modifiedParagraphs[fromIdx];
            modifiedParagraphs.RemoveAt(fromIdx);

            // Adjust available indices after removal
            availableForMove = availableForMove.Select(x => x > fromIdx ? x - 1 : x).ToList();

            int toIdx = Rng.Next(modifiedParagraphs.Count + 1);
            modifiedParagraphs.Insert(toIdx, para);

            // Adjust available indices after insertion
            availableForMove = availableForMove.Select(x => x >= toIdx ? x + 1 : x).ToList();

            moves.Add((fromIdx, toIdx, para.Substring(0, Math.Min(40, para.Length))));
        }

        // Apply deletions
        int deleteCount = changeCount / 3;
        Console.WriteLine($"Applying ~{deleteCount} deletions...");
        for (int i = 0; i < deleteCount && modifiedParagraphs.Count > paragraphCount / 2; i++)
        {
            int idx = Rng.Next(modifiedParagraphs.Count);
            deletions.Add(idx);
            modifiedParagraphs.RemoveAt(idx);
        }

        // Apply insertions
        int insertCount = changeCount / 3;
        Console.WriteLine($"Applying ~{insertCount} insertions...");
        for (int i = 0; i < insertCount; i++)
        {
            int idx = Rng.Next(modifiedParagraphs.Count + 1);
            insertions.Add(idx);
            modifiedParagraphs.Insert(idx, $"[NEW-{i + 1}] This is a newly inserted paragraph with enough words to be meaningful. " +
                $"It contains various content including technical terms, legal jargon, and general prose. " +
                $"The purpose is to test the comparison engine with substantial insertions. Reference: INS-{Guid.NewGuid():N}");
        }

        // Apply modifications (change words in existing paragraphs)
        int modifyCount = changeCount / 3;
        Console.WriteLine($"Applying ~{modifyCount} modifications...");
        for (int i = 0; i < modifyCount && modifiedParagraphs.Count > 0; i++)
        {
            int idx = Rng.Next(modifiedParagraphs.Count);
            var para = modifiedParagraphs[idx];
            // Modify by replacing some words
            para = para.Replace("paragraph", "section")
                       .Replace("content", "material")
                       .Replace("document", "file");
            if (!para.Contains("[MODIFIED]"))
            {
                para = "[MODIFIED] " + para;
            }
            modifiedParagraphs[idx] = para;
            modifications.Add(idx);
        }

        Console.WriteLine($"\nCreating documents...");
        Console.WriteLine($"  Original: {originalParagraphs.Count} paragraphs");
        Console.WriteLine($"  Modified: {modifiedParagraphs.Count} paragraphs");

        var doc1 = CreateDocument(originalParagraphs);
        var doc2 = CreateDocument(modifiedParagraphs);

        Console.WriteLine($"\nComparing with DetectMoves=true...");

        var settings = new WmlComparerSettings
        {
            DetectMoves = true,
            SimplifyMoveMarkup = false,
            MoveSimilarityThreshold = 0.75,  // Slightly lower threshold to catch more moves
            MoveMinimumWordCount = 5,
            AuthorForRevisions = "StressTest"
        };

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        var compared = WmlComparer.Compare(doc1, doc2, settings);
        stopwatch.Stop();

        Console.WriteLine($"  Comparison completed in {stopwatch.ElapsedMilliseconds}ms");

        // Save output
        var outputPath = $"{name}_Output.docx";
        File.WriteAllBytes(outputPath, compared.DocumentByteArray);
        Console.WriteLine($"\n📄 Output: {Path.GetFullPath(outputPath)}");

        // Analyze results
        Console.WriteLine($"\n--- ANALYSIS ---\n");

        using var stream = new MemoryStream(compared.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(stream, false);
        var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

        // Count elements
        var stats = new Dictionary<string, int>
        {
            ["moveFrom"] = mainXDoc.Descendants(W + "moveFrom").Count(),
            ["moveTo"] = mainXDoc.Descendants(W + "moveTo").Count(),
            ["moveFromRangeStart"] = mainXDoc.Descendants(W + "moveFromRangeStart").Count(),
            ["moveFromRangeEnd"] = mainXDoc.Descendants(W + "moveFromRangeEnd").Count(),
            ["moveToRangeStart"] = mainXDoc.Descendants(W + "moveToRangeStart").Count(),
            ["moveToRangeEnd"] = mainXDoc.Descendants(W + "moveToRangeEnd").Count(),
            ["del"] = mainXDoc.Descendants(W + "del").Count(),
            ["ins"] = mainXDoc.Descendants(W + "ins").Count(),
            ["rPrChange"] = mainXDoc.Descendants(W + "rPrChange").Count(),
        };

        Console.WriteLine("Element Counts:");
        Console.WriteLine($"  Move elements:");
        Console.WriteLine($"    moveFrom:          {stats["moveFrom"],4}");
        Console.WriteLine($"    moveTo:            {stats["moveTo"],4}");
        Console.WriteLine($"    moveFromRangeStart:{stats["moveFromRangeStart"],4}");
        Console.WriteLine($"    moveToRangeStart:  {stats["moveToRangeStart"],4}");
        Console.WriteLine($"  Revision elements:");
        Console.WriteLine($"    del:               {stats["del"],4}");
        Console.WriteLine($"    ins:               {stats["ins"],4}");
        Console.WriteLine($"    rPrChange:         {stats["rPrChange"],4}");

        // Collect all revision IDs
        var revisionElements = new[] { "ins", "del", "moveFrom", "moveTo", "rPrChange" };
        var allIds = new List<(string Id, string Type)>();

        foreach (var elemName in revisionElements)
        {
            foreach (var elem in mainXDoc.Descendants(W + elemName))
            {
                var id = elem.Attribute(W + "id")?.Value;
                if (id != null)
                {
                    allIds.Add((id, elemName));
                }
            }
        }

        Console.WriteLine($"\n  Total revision elements with IDs: {allIds.Count}");

        // Check for duplicates - THE CRITICAL TEST
        var duplicates = allIds.GroupBy(x => x.Id)
            .Where(g => g.Count() > 1)
            .ToList();

        if (duplicates.Count > 0)
        {
            Console.WriteLine($"\n  ❌ DUPLICATE IDs FOUND - ISSUE #96 BUG!");
            foreach (var dup in duplicates.Take(10))
            {
                Console.WriteLine($"     ID={dup.Key}: {string.Join(", ", dup.Select(x => x.Type))}");
            }
            if (duplicates.Count > 10)
            {
                Console.WriteLine($"     ... and {duplicates.Count - 10} more duplicates");
            }
        }
        else
        {
            Console.WriteLine($"\n  ✅ All {allIds.Count} revision IDs are UNIQUE!");
        }

        // Check move name pairing
        var moveFromNames = mainXDoc.Descendants(W + "moveFromRangeStart")
            .Select(e => e.Attribute(W + "name")?.Value)
            .Where(n => !string.IsNullOrEmpty(n))
            .ToList();

        var moveToNames = mainXDoc.Descendants(W + "moveToRangeStart")
            .Select(e => e.Attribute(W + "name")?.Value)
            .Where(n => !string.IsNullOrEmpty(n))
            .ToList();

        var uniqueMoveNames = moveFromNames.Union(moveToNames).Distinct().ToList();
        Console.WriteLine($"\n  Move names: {uniqueMoveNames.Count} unique ({string.Join(", ", uniqueMoveNames.Take(10))}{(uniqueMoveNames.Count > 10 ? "..." : "")})");

        // Validate all names are paired
        var unpairedFrom = moveFromNames.Except(moveToNames).ToList();
        var unpairedTo = moveToNames.Except(moveFromNames).ToList();

        if (unpairedFrom.Any() || unpairedTo.Any())
        {
            Console.WriteLine($"  ⚠️  Unpaired move names found:");
            if (unpairedFrom.Any()) Console.WriteLine($"      From without To: {string.Join(", ", unpairedFrom)}");
            if (unpairedTo.Any()) Console.WriteLine($"      To without From: {string.Join(", ", unpairedTo)}");
        }
        else if (uniqueMoveNames.Count > 0)
        {
            Console.WriteLine($"  ✅ All move names properly paired!");
        }

        // OpenXML validation
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(wDoc).ToList();

        if (errors.Count == 0)
        {
            Console.WriteLine($"\n  ✅ OpenXML validation PASSED");
        }
        else
        {
            Console.WriteLine($"\n  ⚠️  OpenXML validation: {errors.Count} issues");
            foreach (var error in errors.Take(5))
            {
                Console.WriteLine($"      {error.ErrorType}: {error.Description.Substring(0, Math.Min(80, error.Description.Length))}...");
            }
        }

        // Final verdict
        Console.WriteLine($"\n--- VERDICT ---");
        bool passed = duplicates.Count == 0;
        if (passed)
        {
            Console.WriteLine($"  🎉 {name} PASSED - No duplicate IDs with {stats["moveFrom"]} moves and {stats["del"] + stats["ins"]} del/ins");
        }
        else
        {
            Console.WriteLine($"  💥 {name} FAILED - Found {duplicates.Count} duplicate IDs");
        }
    }

    static string GenerateParagraph(int index)
    {
        var templates = new[]
        {
            "Paragraph {0}: This document section contains important information about the project requirements and specifications. Reference ID: {1}",
            "Section {0}: The following content describes the technical implementation details for the proposed system architecture. Doc: {1}",
            "Item {0}: According to the agreement dated herein, the parties shall comply with all terms and conditions specified. Contract: {1}",
            "Clause {0}: The licensee agrees to use the software only for purposes permitted under this license agreement. License: {1}",
            "Article {0}: This paragraph establishes the fundamental principles governing the relationship between the entities. Ref: {1}",
            "Point {0}: The data processing activities shall be conducted in accordance with applicable privacy regulations. GDPR: {1}",
            "Note {0}: All modifications to this document must be tracked and approved by the designated review committee. Rev: {1}",
            "Entry {0}: The financial statements have been prepared in accordance with generally accepted accounting principles. GAAP: {1}",
            "Record {0}: This memorandum summarizes the key decisions made during the executive committee meeting. Minutes: {1}",
            "Statement {0}: The undersigned hereby certifies that all information provided is true and accurate. Cert: {1}",
            "Provision {0}: Notwithstanding the foregoing, the obligations set forth herein shall survive termination. Legal: {1}",
            "Stipulation {0}: The contractor shall deliver all work products by the specified deadline. Deadline: {1}",
            "Requirement {0}: The system shall support concurrent users and maintain response times under load. Perf: {1}",
            "Specification {0}: All API endpoints must implement proper authentication and authorization. Security: {1}",
            "Definition {0}: For purposes of this agreement, the following terms shall have the meanings ascribed. Terms: {1}",
        };

        var template = templates[index % templates.Length];
        return string.Format(template, index, $"DOC-{index:D4}-{Guid.NewGuid().ToString().Substring(0, 8).ToUpper()}");
    }

    static WmlDocument CreateDocument(List<string> paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    paragraphs.Select(text =>
                        new Paragraph(
                            new Run(
                                new Text(text)
                            )
                        )
                    )
                )
            );

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(
                        new RunPropertiesBaseStyle(
                            new RunFonts { Ascii = "Calibri" },
                            new FontSize { Val = "22" }
                        )
                    ),
                    new ParagraphPropertiesDefault()
                )
            );

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();

            doc.Save();
        }

        stream.Position = 0;
        return new WmlDocument("test.docx", stream.ToArray());
    }
}
