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

    private static void SetRelationshipId(
        OpenXmlPartContainer parent, OpenXmlPart? part, string relationshipId)
    {
        if (part is not null
            && !string.Equals(parent.GetIdOfPart(part), relationshipId, StringComparison.Ordinal))
            parent.ChangeIdOfPart(part, relationshipId);
    }
}
