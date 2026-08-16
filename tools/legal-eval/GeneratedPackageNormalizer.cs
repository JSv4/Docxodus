// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;

namespace LegalEval;

/// <summary>
/// Makes runner-authored packages byte reproducible without canonicalizing caller-owned model
/// candidates. Relationship ids are fixed through the SDK before the shared package-output policy
/// orders and timestamps the ZIP container.
/// </summary>
internal static class GeneratedPackageNormalizer
{
    private const string FixedRevisionDate = "2026-01-15T12:00:00Z";
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static byte[] Normalize(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var main = package.MainDocumentPart!;
            SetRelationshipId(package, main, "rIdMainDocument");
            SetRelationshipId(main, main.StyleDefinitionsPart, "rIdStyles");
            SetRelationshipId(main, main.NumberingDefinitionsPart, "rIdNumbering");
            SetRelationshipId(main, main.DocumentSettingsPart, "rIdSettings");
            SetRelationshipId(main, main.FontTablePart, "rIdFontTable");
            SetRelationshipId(main, main.ThemePart, "rIdTheme");
            SetRelationshipId(main, main.FootnotesPart, "rIdFootnotes");
            SetRelationshipId(main, main.EndnotesPart, "rIdEndnotes");
            SetRelationshipId(main, main.WordprocessingCommentsPart, "rIdComments");
            SetRelationshipId(main, main.WordprocessingCommentsExPart, "rIdCommentsExtended");
            SetRelationshipId(main, main.WordprocessingCommentsIdsPart, "rIdCommentsIds");

            var parts = main.Parts.Select(value => value.OpenXmlPart)
                .Append(main).Distinct().ToList();
            var dateName = XName.Get("date",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            foreach (var part in parts)
            {
                if (!part.ContentType.EndsWith("+xml", StringComparison.Ordinal)
                    && !part.ContentType.EndsWith("/xml", StringComparison.Ordinal))
                    continue;
                try
                {
                    var xml = part.GetXDocument();
                    foreach (var date in xml.Descendants().Attributes(dateName))
                        date.Value = FixedRevisionDate;
                    part.PutXDocument();
                }
                catch (XmlException)
                {
                    // An opaque XML extension part is outside this fixture's modeled surface.
                }
            }
        }
        return ZipPackageOutputNormalizer.NormalizeDeterministic(stream.ToArray());
    }

    /// <summary>
    /// Restore revision-bearing baseline paragraphs that a composite comparison flattened even
    /// though their accepted text is untouched.  A paragraph is copied only when its accepted text
    /// has one unambiguous, revision-free peer in the result.
    /// </summary>
    internal static byte[] RestoreUnchangedReviewParagraphs(byte[] baseline, byte[] result)
    {
        XDocument baselineXml;
        using (var baselineStream = new MemoryStream(baseline))
        using (var baselinePackage = WordprocessingDocument.Open(baselineStream, false))
            baselineXml = new XDocument(baselinePackage.MainDocumentPart!.GetXDocument());

        using var output = new MemoryStream();
        output.Write(result);
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, true))
        {
            var main = package.MainDocumentPart!;
            var resultXml = main.GetXDocument();
            var resultParagraphs = resultXml.Descendants(W + "p").ToList();
            var changed = false;
            foreach (var paragraph in baselineXml.Descendants(W + "p")
                .Where(ContainsInlineRevision))
            {
                var acceptedText = string.Concat(
                    paragraph.Descendants(W + "t").Select(value => value.Value));
                var matches = resultParagraphs.Where(value =>
                        !ContainsInlineRevision(value)
                        && string.Equals(string.Concat(value.Descendants(W + "t")
                                .Select(text => text.Value)),
                            acceptedText, StringComparison.Ordinal))
                    .ToList();
                if (matches.Count != 1) continue;
                var replacement = new XElement(paragraph);
                var offset = resultParagraphs.IndexOf(matches[0]);
                matches[0].ReplaceWith(replacement);
                resultParagraphs[offset] = replacement;
                changed = true;
            }
            var baselineParagraphs = baselineXml.Descendants(W + "p").ToList();
            if (baselineParagraphs.Count == resultParagraphs.Count)
            {
                for (var index = 0; index < baselineParagraphs.Count; index++)
                {
                    var baselineProperties = baselineParagraphs[index].Element(W + "pPr");
                    if (baselineProperties is null
                        || resultParagraphs[index].Element(W + "pPr") is not null)
                        continue;
                    resultParagraphs[index].AddFirst(new XElement(baselineProperties));
                    changed = true;
                }
            }
            if (changed) main.PutXDocument();
        }
        return ZipPackageOutputNormalizer.NormalizeDeterministic(output.ToArray());
    }

    private static bool ContainsInlineRevision(XElement paragraph) =>
        paragraph.Descendants().Any(value => value.Name == W + "ins"
            || value.Name == W + "del"
            || value.Name == W + "moveFrom"
            || value.Name == W + "moveTo");

    private static void SetRelationshipId(
        OpenXmlPartContainer parent, OpenXmlPart? part, string relationshipId)
    {
        if (part is not null
            && !string.Equals(parent.GetIdOfPart(part), relationshipId, StringComparison.Ordinal))
            parent.ChangeIdOfPart(part, relationshipId);
    }
}
