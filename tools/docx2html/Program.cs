// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using SkiaSharp;

namespace Docx2Html;

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
            Console.WriteLine($"docx2html {Version}");
            return 0;
        }

        // Parse arguments
        string? inputPath = null;
        string? outputPath = null;
        string? pageTitle = null;
        string cssPrefix = "pt-";
        bool embedImages = true;
        bool inlineStyles = false;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].StartsWith("--"))
            {
                if (args[i].StartsWith("--title="))
                {
                    pageTitle = args[i]["--title=".Length..];
                }
                else if (args[i].StartsWith("--css-prefix="))
                {
                    cssPrefix = args[i]["--css-prefix=".Length..];
                }
                else if (args[i] == "--inline-styles")
                {
                    inlineStyles = true;
                }
                else if (args[i] == "--extract-images")
                {
                    embedImages = false;
                }
                else if (args[i].StartsWith("-"))
                {
                    Console.Error.WriteLine($"Error: Unknown option: {args[i]}");
                    return 1;
                }
            }
            else if (inputPath == null)
            {
                inputPath = args[i];
            }
            else if (outputPath == null)
            {
                outputPath = args[i];
            }
            else
            {
                Console.Error.WriteLine($"Error: Unexpected argument: {args[i]}");
                return 1;
            }
        }

        if (inputPath == null)
        {
            Console.Error.WriteLine("Error: Input file is required.");
            Console.Error.WriteLine();
            PrintUsage();
            return 1;
        }

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Error: Input file not found: {inputPath}");
            return 1;
        }

        // Default output path
        if (outputPath == null)
        {
            outputPath = Path.ChangeExtension(inputPath, ".html");
        }

        // Default page title
        if (pageTitle == null)
        {
            pageTitle = Path.GetFileNameWithoutExtension(inputPath);
        }

        try
        {
            Console.WriteLine($"Converting document...");
            Console.WriteLine($"  Input:  {inputPath}");
            Console.WriteLine($"  Output: {outputPath}");

            var imageDirectoryName = Path.Combine(
                Path.GetDirectoryName(outputPath) ?? ".",
                Path.GetFileNameWithoutExtension(outputPath) + "_files"
            );
            int imageCounter = 0;

            byte[] byteArray = File.ReadAllBytes(inputPath);
            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);

            using var wDoc = WordprocessingDocument.Open(memoryStream, true);

            // Try to get title from document properties
            var docTitle = wDoc.CoreFilePropertiesPart?.GetXDocument()
                .Descendants(DC.title)
                .FirstOrDefault()?.Value;
            if (!string.IsNullOrEmpty(docTitle))
            {
                pageTitle = docTitle;
            }

            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = pageTitle,
                FabricateCssClasses = !inlineStyles,
                CssClassPrefix = cssPrefix,
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                ImageHandler = imageInfo =>
                {
                    if (imageInfo.ContentType == null)
                        return null;

                    string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    SKEncodedImageFormat? imageFormat = extension switch
                    {
                        "png" => SKEncodedImageFormat.Png,
                        "gif" => SKEncodedImageFormat.Gif,
                        "bmp" => SKEncodedImageFormat.Png, // Convert BMP to PNG
                        "jpeg" => SKEncodedImageFormat.Jpeg,
                        "tiff" => SKEncodedImageFormat.Png, // Convert TIFF to PNG
                        "x-wmf" => SKEncodedImageFormat.Png, // Convert WMF to PNG
                        _ => null
                    };

                    if (imageFormat == null)
                        return null;

                    // Update extension for converted formats
                    if (extension is "bmp" or "tiff" or "x-wmf")
                        extension = "png";

                    ++imageCounter;

                    if (embedImages)
                    {
                        // Embed as base64 data URI
                        try
                        {
                            string base64;
                            string mimeType;

                            if (imageInfo.ImageBytes != null)
                            {
                                // Re-encode if format changed, otherwise use original
                                if (imageInfo.ContentType.Contains(extension))
                                {
                                    base64 = Convert.ToBase64String(imageInfo.ImageBytes);
                                    mimeType = imageInfo.ContentType;
                                }
                                else
                                {
                                    using var bitmap = SKBitmap.Decode(imageInfo.ImageBytes);
                                    if (bitmap == null) return null;
                                    using var image = SKImage.FromBitmap(bitmap);
                                    using var data = image.Encode(imageFormat.Value, 90);
                                    base64 = Convert.ToBase64String(data.ToArray());
                                    mimeType = $"image/{extension}";
                                }
                            }
                            else if (imageInfo.Bitmap != null)
                            {
                                using var image = SKImage.FromBitmap(imageInfo.Bitmap);
                                using var data = image.Encode(imageFormat.Value, 90);
                                base64 = Convert.ToBase64String(data.ToArray());
                                mimeType = $"image/{extension}";
                            }
                            else
                            {
                                return null;
                            }

                            return new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, $"data:{mimeType};base64,{base64}"),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null
                                    ? new XAttribute(NoNamespace.alt, imageInfo.AltText)
                                    : null);
                        }
                        catch
                        {
                            return null;
                        }
                    }
                    else
                    {
                        // Save to external file
                        try
                        {
                            if (!Directory.Exists(imageDirectoryName))
                                Directory.CreateDirectory(imageDirectoryName);

                            string imageFileName = Path.Combine(imageDirectoryName,
                                $"image{imageCounter}.{extension}");

                            imageInfo.SaveImage(imageFileName, imageFormat.Value);

                            // Use relative path in HTML
                            string relativePath = Path.GetFileName(imageDirectoryName) + "/" +
                                Path.GetFileName(imageFileName);

                            return new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, relativePath),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null
                                    ? new XAttribute(NoNamespace.alt, imageInfo.AltText)
                                    : null);
                        }
                        catch
                        {
                            return null;
                        }
                    }
                }
            };

            XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            File.WriteAllText(outputPath, htmlString, Encoding.UTF8);

            Console.WriteLine();
            Console.WriteLine($"Conversion complete!");
            if (!embedImages && imageCounter > 0)
            {
                Console.WriteLine($"  Images: {imageCounter} image(s) saved to {imageDirectoryName}");
            }
            else if (imageCounter > 0)
            {
                Console.WriteLine($"  Images: {imageCounter} image(s) embedded as base64");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            if (Environment.GetEnvironmentVariable("DOCX2HTML_DEBUG") == "1")
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
        Console.WriteLine($"docx2html {Version} - Convert Word documents to HTML");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  docx2html <input.docx> [output.html] [options]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  input.docx       Path to the input Word document");
        Console.WriteLine("  output.html      Path for the output HTML file (default: input with .html extension)");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --title=<text>       Page title (default: document title or filename)");
        Console.WriteLine("  --css-prefix=<text>  CSS class prefix (default: pt-)");
        Console.WriteLine("  --inline-styles      Use inline styles instead of CSS classes");
        Console.WriteLine("  --extract-images     Save images to separate files instead of embedding");
        Console.WriteLine("  -h, --help           Show this help message");
        Console.WriteLine("  -v, --version        Show version information");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  docx2html document.docx");
        Console.WriteLine("  docx2html document.docx output.html");
        Console.WriteLine("  docx2html document.docx --title=\"My Document\"");
        Console.WriteLine("  docx2html document.docx --extract-images --inline-styles");
        Console.WriteLine();
        Console.WriteLine("Environment Variables:");
        Console.WriteLine("  DOCX2HTML_DEBUG=1    Show detailed error information");
    }
}
