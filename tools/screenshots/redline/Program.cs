// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;

namespace RedlineScreenshot;

internal static class Program
{
    private const string QualifiedKeyHolder = "A “Qualified Key Holder” is a Key Holder";
    private const string Interpretation = "Any reference in this Agreement to “vote” or “voting”";
    private const string Affiliate = "“Affiliate” means, with respect to any specified Person";
    private const string Sanctions = "“Sanctions” means applicable laws and regulations";
    private const string Series = "“Series [___] Preferred Stock” means";
    private const string Shares = "“Shares” shall mean and include";

    private static readonly string[] ExpectedDefinitionMarkers =
    {
        "(a)", "(b)", "(c)", "(d)", "(e)", "(f)", "(g)", "(g)", "(h)",
        "(i)", "(j)", "(k)", "(l)", "(m)", "(n)", "(o)", "(p)",
    };

    private static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine(
                "usage: redline-screenshot <NVCA-Model-VA-10-1-2025.docx> <output-directory>");
            return 1;
        }

        var sourcePath = Path.GetFullPath(args[0]);
        var outputDirectory = Path.GetFullPath(args[1]);
        Directory.CreateDirectory(outputDirectory);

        var sourceBytes = File.ReadAllBytes(sourcePath);
        var structurallyEdited = ApplyStructuralEdits(sourceBytes);
        var modifiedBytes = ApplyTextEdits(structurallyEdited);

        var modifiedPath = Path.Combine(outputDirectory, "modified.docx");
        File.WriteAllBytes(modifiedPath, modifiedBytes);

        var redline = DocxDiff.Compare(
            new WmlDocument(sourcePath, sourceBytes),
            new WmlDocument(modifiedPath, modifiedBytes),
            new DocxDiffSettings { AuthorForRevisions = "Company Counsel" });

        var redlinePath = Path.Combine(outputDirectory, "redline.docx");
        File.WriteAllBytes(redlinePath, redline.DocumentByteArray);

        var html = RenderRedline(redline.DocumentByteArray);
        AssertDefinitionNumbering(html);

        var htmlPath = Path.Combine(outputDirectory, "redline.html");
        File.WriteAllText(
            htmlPath,
            html.ToString(SaveOptions.DisableFormatting),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        Console.WriteLine($"modified: {modifiedPath}");
        Console.WriteLine($"redline:  {redlinePath}");
        Console.WriteLine($"html:     {htmlPath}");
        Console.WriteLine($"markers:  {string.Join(", ", ExpectedDefinitionMarkers)}");
        return 0;
    }

    private static byte[] ApplyStructuralEdits(byte[] sourceBytes)
    {
        using var stream = ExpandableStream(sourceBytes);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = document.MainDocumentPart
                ?? throw new InvalidOperationException("The voting agreement has no main document part.");
            var mainDocument = mainPart.Document
                ?? throw new InvalidOperationException("The voting agreement has no main document.");
            var body = mainDocument.Body
                ?? throw new InvalidOperationException("The voting agreement has no document body.");

            var qualifiedKeyHolder = FindParagraph(body, QualifiedKeyHolder);
            var interpretation = FindParagraph(body, Interpretation);
            var affiliate = FindParagraph(body, Affiliate);
            var sanctions = FindParagraph(body, Sanctions);

            // The accepted edit round shown in the README: move the interpretation definition
            // to the start, delete Qualified Key Holder, and insert Sanctions Authority. Moving
            // the existing paragraph (rather than cloning it) gives DocxDiff a true relocation.
            interpretation.Remove();
            affiliate.InsertBeforeSelf(interpretation);
            qualifiedKeyHolder.Remove();
            sanctions.InsertAfterSelf(BuildSanctionsAuthorityParagraph(sanctions));

            mainDocument.Save();
        }

        return stream.ToArray();
    }

    private static byte[] ApplyTextEdits(byte[] structurallyEdited)
    {
        using var session = new DocxSession(structurallyEdited, new DocxSessionSettings
        {
            CaptureInitialProjection = false,
            EmitMarkdownPatch = false,
        });

        var series = session.FindByText(Series)
            ?? throw new InvalidOperationException("Could not find the Series definition.");
        var seriesResults = session.ReplaceTextRange(series.Anchor.Id, "[___]", "A");
        if (seriesResults.Count != 2 || seriesResults.Any(result => !result.Success))
            throw new InvalidOperationException("Expected to fill both Series placeholders.");

        var shares = session.FindByText(Shares)
            ?? throw new InvalidOperationException("Could not find the Shares definition.");
        var sharesResults = session.ReplaceTextRange(
            shares.Anchor.Id,
            "shall mean and include",
            "means");
        if (sharesResults.Count != 1 || !sharesResults[0].Success)
            throw new InvalidOperationException("Expected to tighten the Shares definition once.");

        return session.Save();
    }

    private static Paragraph BuildSanctionsAuthorityParagraph(Paragraph sanctions)
    {
        var paragraph = new Paragraph();
        if (sanctions.ParagraphProperties is not null)
        {
            paragraph.Append(
                (ParagraphProperties)sanctions.ParagraphProperties.CloneNode(deep: true));
        }

        paragraph.Append(
            new Run(
                new RunProperties(new Bold()),
                new Text("“Sanctions Authority”")),
            new Run(
                new Text(
                    " means the United States (including OFAC and the U.S. Department of State), " +
                    "the United Nations Security Council, the European Union, and His Majesty’s " +
                    "Treasury of the United Kingdom.")));
        return paragraph;
    }

    private static XElement RenderRedline(byte[] redlineBytes)
    {
        using var stream = ExpandableStream(redlineBytes);
        using var document = WordprocessingDocument.Open(stream, true);
        return WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings
        {
            PageTitle = "NVCA voting agreement — redline",
            FabricateCssClasses = false,
            RenderTrackedChanges = true,
            RenderMoveOperations = true,
            IncludeRevisionMetadata = true,
            ShowDeletedContent = true,
        });
    }

    private static void AssertDefinitionNumbering(XElement html)
    {
        var xhtml = (XNamespace)"http://www.w3.org/1999/xhtml";
        var markerPattern = new Regex("^\\([a-p]\\)$", RegexOptions.CultureInvariant);
        var actual = html
            .Descendants(xhtml + "span")
            .Where(element => (string?)element.Attribute("data-list-marker") == "true")
            .Where(element => !element
                .Descendants(xhtml + "span")
                .Any(descendant => (string?)descendant.Attribute("data-list-marker") == "true"))
            .Select(element => element.Value)
            .Where(value => markerPattern.IsMatch(value))
            .Take(ExpectedDefinitionMarkers.Length)
            .ToArray();

        if (!actual.SequenceEqual(ExpectedDefinitionMarkers))
        {
            throw new InvalidOperationException(
                "Tracked-change list numbering regressed. " +
                $"Expected [{string.Join(", ", ExpectedDefinitionMarkers)}], " +
                $"got [{string.Join(", ", actual)}].");
        }
    }

    private static Paragraph FindParagraph(Body body, string text)
    {
        return body
            .Descendants<Paragraph>()
            .Single(paragraph => paragraph.InnerText.Contains(text, StringComparison.Ordinal));
    }

    private static MemoryStream ExpandableStream(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }
}
