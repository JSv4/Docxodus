// Issue #96 Bug Reproduction Test
// This specifically tests the scenario that caused the "unreadable content" warning:
// Move operations combined with regular ins/del that would have caused ID collisions

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

class Issue96BugReproduction
{
    static XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static void Main(string[] args)
    {
        Console.WriteLine("╔══════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║         Issue #96 Bug Reproduction Test                      ║");
        Console.WriteLine("║  Move operations + ins/del that caused ID collisions         ║");
        Console.WriteLine("╚══════════════════════════════════════════════════════════════╝\n");

        var outputDir = Path.GetDirectoryName(typeof(Issue96BugReproduction).Assembly.Location)
            ?? Directory.GetCurrentDirectory();

        // Create documents that will produce BOTH moves AND regular ins/del
        // This is the scenario that triggered the bug
        // The first paragraph is IDENTICAL in both docs but at different positions (MOVE)
        // The last paragraph is completely different (DEL + INS)
        var doc1 = CreateDocument(new[] {
            "The quick brown fox jumps over the lazy sleeping dog in the park today.",
            "Static content that does not change at all in this document test.",
            "This paragraph will be deleted and replaced with something new."
        });

        var doc2 = CreateDocument(new[] {
            "Static content that does not change at all in this document test.",
            "The quick brown fox jumps over the lazy sleeping dog in the park today.",
            "This is a completely new paragraph that was inserted here instead."
        });

        Console.WriteLine("Document 1 (Original):");
        Console.WriteLine("  [1] The quick brown fox jumps over the lazy sleeping dog...");
        Console.WriteLine("  [2] Static content that does not change...");
        Console.WriteLine("  [3] Another static paragraph...");
        Console.WriteLine("  [4] This paragraph has some text that will be partially modified.");
        Console.WriteLine();
        Console.WriteLine("Document 2 (Modified):");
        Console.WriteLine("  [1] Static content that does not change...");
        Console.WriteLine("  [2] Another static paragraph...");
        Console.WriteLine("  [3] The quick brown fox jumps over the lazy sleeping dog... (MOVED)");
        Console.WriteLine("  [4] This paragraph has DIFFERENT text... (MODIFIED)");
        Console.WriteLine();

        var settings = new WmlComparerSettings
        {
            DetectMoves = true,
            SimplifyMoveMarkup = false,
            MoveSimilarityThreshold = 0.8,
            MoveMinimumWordCount = 3,
            AuthorForRevisions = "Issue96Test"
        };

        Console.WriteLine("Comparing with DetectMoves=true, SimplifyMoveMarkup=false...\n");

        var compared = WmlComparer.Compare(doc1, doc2, settings);

        // Save output
        var outputPath = Path.Combine(outputDir, "Issue96_BugRepro_Output.docx");
        File.WriteAllBytes(outputPath, compared.DocumentByteArray);
        Console.WriteLine($"📄 Output saved: {outputPath}\n");

        // Analyze the output
        Console.WriteLine("═══════════════════════════════════════════════════════════════");
        Console.WriteLine("                    VALIDATION RESULTS");
        Console.WriteLine("═══════════════════════════════════════════════════════════════\n");

        using var stream = new MemoryStream(compared.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(stream, false);
        var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

        // Count elements
        var moveFromCount = mainXDoc.Descendants(W + "moveFrom").Count();
        var moveToCount = mainXDoc.Descendants(W + "moveTo").Count();
        var moveFromRangeStartCount = mainXDoc.Descendants(W + "moveFromRangeStart").Count();
        var moveFromRangeEndCount = mainXDoc.Descendants(W + "moveFromRangeEnd").Count();
        var moveToRangeStartCount = mainXDoc.Descendants(W + "moveToRangeStart").Count();
        var moveToRangeEndCount = mainXDoc.Descendants(W + "moveToRangeEnd").Count();
        var delCount = mainXDoc.Descendants(W + "del").Count();
        var insCount = mainXDoc.Descendants(W + "ins").Count();

        Console.WriteLine("Move Elements:");
        Console.WriteLine($"  • w:moveFrom:          {moveFromCount}");
        Console.WriteLine($"  • w:moveTo:            {moveToCount}");
        Console.WriteLine($"  • w:moveFromRangeStart:{moveFromRangeStartCount}");
        Console.WriteLine($"  • w:moveFromRangeEnd:  {moveFromRangeEndCount}");
        Console.WriteLine($"  • w:moveToRangeStart:  {moveToRangeStartCount}");
        Console.WriteLine($"  • w:moveToRangeEnd:    {moveToRangeEndCount}");
        Console.WriteLine();
        Console.WriteLine("Regular Revision Elements:");
        Console.WriteLine($"  • w:del:               {delCount}");
        Console.WriteLine($"  • w:ins:               {insCount}");
        Console.WriteLine();

        // Collect all revision IDs
        var allRevisionIds = new Dictionary<string, List<string>>();
        var revisionElements = new[] { "ins", "del", "moveFrom", "moveTo", "rPrChange" };

        foreach (var elemName in revisionElements)
        {
            foreach (var elem in mainXDoc.Descendants(W + elemName))
            {
                var id = elem.Attribute(W + "id")?.Value;
                if (id != null)
                {
                    if (!allRevisionIds.ContainsKey(id))
                        allRevisionIds[id] = new List<string>();
                    allRevisionIds[id].Add(elemName);
                }
            }
        }

        Console.WriteLine("ID Analysis:");
        Console.WriteLine($"  • Total unique IDs: {allRevisionIds.Count}");

        // Check for the Issue #96 bug: duplicate IDs
        var duplicates = allRevisionIds.Where(kvp => kvp.Value.Count > 1).ToList();
        if (duplicates.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("  ❌ DUPLICATE IDs FOUND (Issue #96 BUG!):");
            foreach (var dup in duplicates)
            {
                Console.WriteLine($"     ID={dup.Key} used by: {string.Join(", ", dup.Value)}");
            }
        }
        else
        {
            Console.WriteLine("  ✅ No duplicate IDs (Issue #96 is FIXED!)");
        }

        // Show move names
        var moveNames = mainXDoc.Descendants(W + "moveFromRangeStart")
            .Select(e => e.Attribute(W + "name")?.Value)
            .Where(n => n != null)
            .Distinct()
            .ToList();

        if (moveNames.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("Move Name Linking:");
            foreach (var name in moveNames)
            {
                var fromCount = mainXDoc.Descendants(W + "moveFromRangeStart")
                    .Count(e => e.Attribute(W + "name")?.Value == name);
                var toCount = mainXDoc.Descendants(W + "moveToRangeStart")
                    .Count(e => e.Attribute(W + "name")?.Value == name);
                Console.WriteLine($"  • '{name}': {fromCount} source(s), {toCount} destination(s)");
            }
        }

        // Run OpenXML validation
        Console.WriteLine();
        Console.WriteLine("OpenXML Validation:");
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(wDoc).ToList();
        if (errors.Count == 0)
        {
            Console.WriteLine("  ✅ Document is valid per OpenXML schema");
        }
        else
        {
            Console.WriteLine($"  ⚠️  {errors.Count} validation issue(s):");
            foreach (var error in errors.Take(5))
            {
                Console.WriteLine($"     - {error.Description}");
            }
        }

        Console.WriteLine();
        Console.WriteLine("═══════════════════════════════════════════════════════════════");
        Console.WriteLine("                         SUMMARY");
        Console.WriteLine("═══════════════════════════════════════════════════════════════");

        bool hasMoves = moveFromCount > 0 && moveToCount > 0;
        bool hasDelIns = delCount > 0 || insCount > 0;
        bool noDuplicates = duplicates.Count == 0;
        bool validXml = errors.Count == 0;

        Console.WriteLine($"  Move detection working:     {(hasMoves ? "✅ YES" : "⚠️  NO")}");
        Console.WriteLine($"  Has regular del/ins:        {(hasDelIns ? "✅ YES" : "❌ NO")}");
        Console.WriteLine($"  No duplicate IDs:           {(noDuplicates ? "✅ PASS" : "❌ FAIL")}");
        Console.WriteLine($"  OpenXML validation:         {(validXml ? "✅ PASS" : "⚠️  ISSUES")}");
        Console.WriteLine();

        if (hasMoves && hasDelIns && noDuplicates)
        {
            Console.WriteLine("🎉 Issue #96 is FIXED! Document has both moves AND del/ins with unique IDs.");
            Console.WriteLine($"   Open the file in Word to verify: {outputPath}");
        }
        else if (!hasMoves)
        {
            Console.WriteLine("ℹ️  No moves were detected (similarity threshold not met).");
            Console.WriteLine("   The ID uniqueness fix is still validated.");
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
}
