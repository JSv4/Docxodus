// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus;

namespace Docx2OC;

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
            Console.WriteLine($"docx2oc {Version}");
            return 0;
        }

        if (args.Length < 1 || args.Length > 2)
        {
            Console.Error.WriteLine("Error: Invalid number of arguments.");
            Console.Error.WriteLine();
            PrintUsage();
            return 1;
        }

        // Parse arguments: docx2oc <input.docx> [output.json]
        var inputFilePath = args[0];
        string outputFilePath;

        if (args.Length == 2)
        {
            outputFilePath = args[1];
        }
        else
        {
            // Default: same filename with .oc extension
            outputFilePath = Path.ChangeExtension(inputFilePath, ".oc");
        }

        if (!File.Exists(inputFilePath))
        {
            Console.Error.WriteLine($"Error: Input file not found: {inputFilePath}");
            return 1;
        }

        if (!inputFilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine($"Error: Input file must be a .docx file: {inputFilePath}");
            return 1;
        }

        try
        {
            Console.WriteLine($"Exporting to OpenContracts format...");
            Console.WriteLine($"  Input:  {inputFilePath}");

            var documentBytes = File.ReadAllBytes(inputFilePath);
            var wmlDocument = new WmlDocument(inputFilePath, documentBytes);

            var export = OpenContractExporter.Export(wmlDocument);

            var json = JsonSerializer.Serialize(export, ExportJsonContext.Default.OpenContractDocExport);
            File.WriteAllText(outputFilePath, json);

            Console.WriteLine($"  Output: {outputFilePath}");
            Console.WriteLine();
            Console.WriteLine($"Export complete:");
            Console.WriteLine($"  Title:       {export.Title}");
            Console.WriteLine($"  Pages:       {export.PageCount}");
            Console.WriteLine($"  Content:     {export.Content.Length:N0} characters");
            Console.WriteLine($"  Annotations: {export.LabelledText.Count}");
            Console.WriteLine($"  PAWLS Pages: {export.PawlsFileContent.Count}");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            if (Environment.GetEnvironmentVariable("DOCX2OC_DEBUG") == "1")
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
        Console.WriteLine($"docx2oc {Version} - Export Word documents to OpenContracts format");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  docx2oc <input.docx> [output.json]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  input.docx     Path to the input Word document");
        Console.WriteLine("  output.json    Path for the output JSON file (optional)");
        Console.WriteLine("                 Default: same name as input with .oc extension");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -h, --help     Show this help message");
        Console.WriteLine("  -v, --version  Show version information");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  docx2oc contract.docx                  # Creates contract.oc");
        Console.WriteLine("  docx2oc contract.docx export.json      # Creates export.json");
        Console.WriteLine("  docx2oc document.docx analysis.oc      # Creates analysis.oc");
        Console.WriteLine();
        Console.WriteLine("Output Format:");
        Console.WriteLine("  The output is a JSON file containing:");
        Console.WriteLine("  - title: Document title");
        Console.WriteLine("  - content: Full extracted text");
        Console.WriteLine("  - pageCount: Estimated page count");
        Console.WriteLine("  - pawlsFileContent: Token positions for each page");
        Console.WriteLine("  - labelledText: Structural annotations (sections, paragraphs, tables)");
        Console.WriteLine("  - relationships: Annotation relationships");
        Console.WriteLine();
        Console.WriteLine("Environment Variables:");
        Console.WriteLine("  DOCX2OC_DEBUG=1  Show detailed error information");
    }
}

// Source-generated serializer for the export (issue #903). Reflection-based System.Text.Json is
// disabled in trimmed and NativeAOT builds, so the export's shape is declared here at compile time.
// The options match the reflection-based ones this replaced: indented, camelCase, nulls omitted.
// OpenContractsAnnotation.AnnotationJson is typed `object`; the runtime types the exporter can put
// there are registered so the generated resolver can serialize them polymorphically.
[JsonSourceGenerationOptions(
    WriteIndented = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(OpenContractDocExport))]
[JsonSerializable(typeof(TextSpan))]
[JsonSerializable(typeof(Dictionary<string, OpenContractsSinglePageAnnotation>))]
internal sealed partial class ExportJsonContext : JsonSerializerContext
{
}
