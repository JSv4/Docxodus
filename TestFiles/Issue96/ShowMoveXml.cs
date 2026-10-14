// Simple test to show move markup is working correctly
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;

class ShowMoveXml
{
    static void Main()
    {
        Console.WriteLine("Creating test documents...\n");

        // Create doc1: A, B, C
        var doc1 = CreateDoc(
            "This is paragraph A with enough words for move detection.",
            "This is paragraph B with sufficient content here.",
            "This is paragraph C with more words added."
        );

        // Create doc2: B, A, C (A moved after B)
        var doc2 = CreateDoc(
            "This is paragraph B with sufficient content here.",
            "This is paragraph A with enough words for move detection.",
            "This is paragraph C with more words added."
        );

        Console.WriteLine("Doc1: [A] [B] [C]");
        Console.WriteLine("Doc2: [B] [A] [C]  (A moved after B)\n");

        var settings = new WmlComparerSettings
        {
            DetectMoves = true,
            SimplifyMoveMarkup = false,
            MoveSimilarityThreshold = 0.8,
            MoveMinimumWordCount = 3
        };

        var compared = WmlComparer.Compare(doc1, doc2, settings);

        // Extract and display the body XML
        using var stream = new MemoryStream(compared.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(stream, false);

        var body = wDoc.MainDocumentPart.Document.Body;
        var bodyXml = XElement.Parse(body.OuterXml);

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        Console.WriteLine("=== MOVE ELEMENTS FOUND ===\n");

        // Show moveFromRangeStart elements
        var moveFromStarts = bodyXml.Descendants(w + "moveFromRangeStart").ToList();
        Console.WriteLine($"moveFromRangeStart: {moveFromStarts.Count}");
        foreach (var e in moveFromStarts)
        {
            Console.WriteLine($"  id={e.Attribute(w + "id")?.Value}, name={e.Attribute(w + "name")?.Value}");
        }

        // Show moveFrom elements
        var moveFroms = bodyXml.Descendants(w + "moveFrom").ToList();
        Console.WriteLine($"\nmoveFrom: {moveFroms.Count}");
        foreach (var e in moveFroms)
        {
            var text = string.Join("", e.Descendants(w + "t").Select(t => t.Value));
            Console.WriteLine($"  id={e.Attribute(w + "id")?.Value}, text=\"{Truncate(text, 50)}\"");
        }

        // Show moveToRangeStart elements
        var moveToStarts = bodyXml.Descendants(w + "moveToRangeStart").ToList();
        Console.WriteLine($"\nmoveToRangeStart: {moveToStarts.Count}");
        foreach (var e in moveToStarts)
        {
            Console.WriteLine($"  id={e.Attribute(w + "id")?.Value}, name={e.Attribute(w + "name")?.Value}");
        }

        // Show moveTo elements
        var moveTos = bodyXml.Descendants(w + "moveTo").ToList();
        Console.WriteLine($"\nmoveTo: {moveTos.Count}");
        foreach (var e in moveTos)
        {
            var text = string.Join("", e.Descendants(w + "t").Select(t => t.Value));
            Console.WriteLine($"  id={e.Attribute(w + "id")?.Value}, text=\"{Truncate(text, 50)}\"");
        }

        Console.WriteLine("\n=== DEL/INS ELEMENTS ===\n");

        var dels = bodyXml.Descendants(w + "del").ToList();
        var inss = bodyXml.Descendants(w + "ins").ToList();
        Console.WriteLine($"del: {dels.Count}");
        Console.WriteLine($"ins: {inss.Count}");

        Console.WriteLine("\n=== ALL REVISION IDs ===\n");

        var allIds = new[] { "del", "ins", "moveFrom", "moveTo" }
            .SelectMany(name => bodyXml.Descendants(w + name)
                .Select(e => new { Type = name, Id = e.Attribute(w + "id")?.Value }))
            .Where(x => x.Id != null)
            .OrderBy(x => int.Parse(x.Id))
            .ToList();

        foreach (var item in allIds)
        {
            Console.WriteLine($"  {item.Type,-10} id={item.Id}");
        }

        // Check for duplicates
        var duplicates = allIds.GroupBy(x => x.Id).Where(g => g.Count() > 1).ToList();
        Console.WriteLine();
        if (duplicates.Any())
        {
            Console.WriteLine("❌ DUPLICATE IDs FOUND (BUG!):");
            foreach (var dup in duplicates)
            {
                Console.WriteLine($"   ID {dup.Key}: {string.Join(", ", dup.Select(x => x.Type))}");
            }
        }
        else
        {
            Console.WriteLine("✅ All IDs are unique - Issue #96 is FIXED!");
        }

        // Save output file
        var outputPath = "Issue96_MoveDemo.docx";
        File.WriteAllBytes(outputPath, compared.DocumentByteArray);
        Console.WriteLine($"\n📄 Output saved: {Path.GetFullPath(outputPath)}");
    }

    static WmlDocument CreateDoc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(t => new Paragraph(new Run(new Text(t))))));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();
            doc.Save();
        }
        stream.Position = 0;
        return new WmlDocument("test.docx", stream.ToArray());
    }

    static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
