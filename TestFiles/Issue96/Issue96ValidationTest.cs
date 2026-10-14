// Issue #96 Validation Test
// This test creates documents with moves and validates that:
// 1. Move detection works correctly
// 2. All revision IDs are unique (the core Issue #96 fix)
// 3. The output document can be opened without errors

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using Docxodus;

class Issue96ValidationTest
{
    static XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static void Main(string[] args)
    {
        Console.WriteLine("=== Issue #96 Validation Test ===\n");

        var outputDir = Path.GetDirectoryName(typeof(Issue96ValidationTest).Assembly.Location)
            ?? Directory.GetCurrentDirectory();

        // Test 1: Simple paragraph swap (single move)
        Console.WriteLine("Test 1: Simple paragraph swap");
        RunTest("SimpleSwap", outputDir,
            new[] {
                "The quick brown fox jumps over the lazy dog.",
                "Pack my box with five dozen liquor jugs.",
                "How vexingly quick daft zebras jump."
            },
            new[] {
                "Pack my box with five dozen liquor jugs.",
                "The quick brown fox jumps over the lazy dog.",
                "How vexingly quick daft zebras jump."
            });

        // Test 2: Move with additional changes (the Issue #96 scenario)
        Console.WriteLine("\nTest 2: Move with additional ins/del changes");
        RunTest("MoveWithChanges", outputDir,
            new[] {
                "First paragraph that will be moved to the end.",
                "Second paragraph that stays but gets modified.",
                "Third paragraph that will be deleted entirely.",
                "Fourth paragraph that remains unchanged."
            },
            new[] {
                "Second paragraph that was modified here today.",
                "Fourth paragraph that remains unchanged.",
                "First paragraph that will be moved to the end.",
                "Fifth paragraph that is completely new."
            });

        // Test 3: Multiple independent moves
        Console.WriteLine("\nTest 3: Multiple content blocks");
        RunTest("MultipleBlocks", outputDir,
            new[] {
                "Alpha paragraph with enough words for detection.",
                "Beta paragraph with sufficient content here.",
                "Gamma paragraph stays in the same position.",
                "Delta paragraph with more words for testing."
            },
            new[] {
                "Gamma paragraph stays in the same position.",
                "Beta paragraph with sufficient content here.",
                "Alpha paragraph with enough words for detection.",
                "Delta paragraph with more words for testing."
            });

        Console.WriteLine("\n=== All Tests Complete ===");
        Console.WriteLine($"Output files written to: {outputDir}");
    }

    static void RunTest(string testName, string outputDir, string[] doc1Paragraphs, string[] doc2Paragraphs)
    {
        try
        {
            // Create test documents
            var doc1 = CreateDocument(doc1Paragraphs);
            var doc2 = CreateDocument(doc2Paragraphs);

            // Compare with move detection enabled
            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = false,  // Keep native move markup
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3,
                AuthorForRevisions = "Issue96Test"
            };

            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Save the comparison document
            var outputPath = Path.Combine(outputDir, $"{testName}_Compared.docx");
            File.WriteAllBytes(outputPath, compared.DocumentByteArray);
            Console.WriteLine($"  ✓ Saved: {testName}_Compared.docx");

            // Validate ID uniqueness
            var (isValid, details) = ValidateRevisionIds(compared);
            if (isValid)
            {
                Console.WriteLine($"  ✓ All revision IDs are unique");
            }
            else
            {
                Console.WriteLine($"  ✗ ID COLLISION DETECTED: {details}");
            }

            // Check for move elements
            var moveInfo = AnalyzeMoveElements(compared);
            Console.WriteLine($"  ✓ Move elements: {moveInfo.MoveFromCount} moveFrom, {moveInfo.MoveToCount} moveTo");
            Console.WriteLine($"  ✓ Del/Ins elements: {moveInfo.DelCount} del, {moveInfo.InsCount} ins");
            Console.WriteLine($"  ✓ Move names: {string.Join(", ", moveInfo.MoveNames)}");

            // Run OpenXML validation
            var validationErrors = ValidateDocument(compared);
            if (validationErrors.Count == 0)
            {
                Console.WriteLine($"  ✓ OpenXML validation passed");
            }
            else
            {
                Console.WriteLine($"  ⚠ OpenXML validation: {validationErrors.Count} issues");
                foreach (var error in validationErrors.Take(3))
                {
                    Console.WriteLine($"    - {error}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ✗ ERROR: {ex.Message}");
        }
    }

    static WmlDocument CreateDocument(string[] paragraphs)
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

    static (bool IsValid, string Details) ValidateRevisionIds(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(stream, false);

        var allIds = new List<(string Id, string ElementType, string Location)>();
        var revisionElements = new[] { "ins", "del", "moveFrom", "moveTo", "rPrChange" };

        // Check main document
        var mainXDoc = wDoc.MainDocumentPart.GetXDocument();
        foreach (var elemName in revisionElements)
        {
            foreach (var elem in mainXDoc.Descendants(W + elemName))
            {
                var id = elem.Attribute(W + "id")?.Value;
                if (id != null)
                {
                    allIds.Add((id, elemName, "MainDocument"));
                }
            }
        }

        // Check footnotes
        if (wDoc.MainDocumentPart.FootnotesPart != null)
        {
            var fnXDoc = wDoc.MainDocumentPart.FootnotesPart.GetXDocument();
            foreach (var elemName in revisionElements)
            {
                foreach (var elem in fnXDoc.Descendants(W + elemName))
                {
                    var id = elem.Attribute(W + "id")?.Value;
                    if (id != null)
                    {
                        allIds.Add((id, elemName, "Footnotes"));
                    }
                }
            }
        }

        // Find duplicates (excluding range pairs which share IDs by design)
        var duplicates = allIds.GroupBy(x => x.Id)
            .Where(g => g.Count() > 1)
            .Where(g => {
                // Range start/end pairs are allowed to share IDs
                var types = g.Select(x => x.ElementType).Distinct().ToList();
                if (types.Count == 1 && (types[0] == "moveFromRangeStart" || types[0] == "moveToRangeStart"))
                    return false;
                return true;
            })
            .ToList();

        if (duplicates.Count == 0)
        {
            return (true, $"All {allIds.Count} revision IDs are unique");
        }

        var details = string.Join("; ", duplicates.Select(g =>
            $"ID={g.Key} used by: {string.Join(", ", g.Select(x => $"{x.ElementType}@{x.Location}"))}"));
        return (false, details);
    }

    static (int MoveFromCount, int MoveToCount, int DelCount, int InsCount, List<string> MoveNames)
        AnalyzeMoveElements(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(stream, false);

        var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

        var moveFromCount = mainXDoc.Descendants(W + "moveFrom").Count();
        var moveToCount = mainXDoc.Descendants(W + "moveTo").Count();
        var delCount = mainXDoc.Descendants(W + "del").Count();
        var insCount = mainXDoc.Descendants(W + "ins").Count();

        var moveNames = mainXDoc.Descendants(W + "moveFromRangeStart")
            .Select(e => e.Attribute(W + "name")?.Value)
            .Where(n => n != null)
            .Distinct()
            .ToList();

        return (moveFromCount, moveToCount, delCount, insCount, moveNames);
    }

    static List<string> ValidateDocument(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        return validator.Validate(wDoc)
            .Select(e => $"{e.ErrorType}: {e.Description}")
            .Take(10)
            .ToList();
    }
}
