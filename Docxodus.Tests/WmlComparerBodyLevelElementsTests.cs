// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Regression tests for issue #128 — body-level non-paragraph elements
/// (<see cref="W.bookmarkStart"/>/<see cref="W.bookmarkEnd"/>,
/// <see cref="W.permStart"/>/<see cref="W.permEnd"/>, <see cref="W.proofErr"/>)
/// that appear as direct children of <see cref="W.body"/> rather than nested
/// inside <see cref="W.p"/>.
///
/// PR #124 patched <c>FindIndexOfNextParaMark</c> for one sibling site; this
/// file covers the producer-side filter (<c>ElementsToThrowAway</c>) and the
/// other three consumer sites flagged in #128
/// (<c>FindCommonAtBeginningAndEnd</c>, <c>SplitAtParagraphMark</c>,
/// <c>DoLcsAlgorithm</c>).
///
/// Fixtures are built programmatically (a few KB each) instead of the 4 MB
/// binary pair used by <see cref="WmlComparerBodyLevelBookmarkTests"/>, so the
/// tests pin specific behavior — <see cref="WmlComparer.Compare"/> must
/// succeed — rather than just asserting "no <see cref="System.NullReferenceException"/>".
/// </summary>
public class WmlComparerBodyLevelElementsTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void WC_BodyBookmark_001_BookmarkAtBodyLevel_ComparesWithoutException()
    {
        var before = BuildDocxWithBodyLevelElement(
            new XElement(W + "bookmarkStart",
                new XAttribute(W + "id", "1"),
                new XAttribute(W + "name", "MARK1")),
            new XElement(W + "bookmarkEnd",
                new XAttribute(W + "id", "1")),
            paragraphTexts: new[] { "First paragraph.", "Second paragraph." });
        var after = BuildDocxWithBodyLevelElement(
            bodyLevelBefore: null,
            bodyLevelAfter: null,
            paragraphTexts: new[] { "First paragraph.", "Second paragraph modified." });

        var settings = new WmlComparerSettings();
        var result = WmlComparer.Compare(new WmlDocument("before.docx", before),
                                         new WmlDocument("after.docx", after),
                                         settings);

        Assert.NotNull(result);
        Assert.NotNull(result.DocumentByteArray);
    }

    [Fact]
    public void WC_BodyBookmark_002_BookmarkAtBodyLevel_ReverseDirection()
    {
        var before = BuildDocxWithBodyLevelElement(
            bodyLevelBefore: null,
            bodyLevelAfter: null,
            paragraphTexts: new[] { "First paragraph.", "Second paragraph." });
        var after = BuildDocxWithBodyLevelElement(
            new XElement(W + "bookmarkStart",
                new XAttribute(W + "id", "1"),
                new XAttribute(W + "name", "MARK1")),
            new XElement(W + "bookmarkEnd",
                new XAttribute(W + "id", "1")),
            paragraphTexts: new[] { "First paragraph.", "Second paragraph modified." });

        var settings = new WmlComparerSettings();
        var result = WmlComparer.Compare(new WmlDocument("before.docx", before),
                                         new WmlDocument("after.docx", after),
                                         settings);

        Assert.NotNull(result);
        Assert.NotNull(result.DocumentByteArray);
    }

    [Fact]
    public void WC_BodyPerm_001_PermStartEndAtBodyLevel_ComparesWithoutException()
    {
        var before = BuildDocxWithBodyLevelElement(
            new XElement(W + "permStart",
                new XAttribute(W + "id", "1"),
                new XAttribute(W + "edGrp", "everyone")),
            new XElement(W + "permEnd",
                new XAttribute(W + "id", "1")),
            paragraphTexts: new[] { "Para A.", "Para B." });
        var after = BuildDocxWithBodyLevelElement(
            bodyLevelBefore: null,
            bodyLevelAfter: null,
            paragraphTexts: new[] { "Para A.", "Para B modified." });

        var settings = new WmlComparerSettings();
        var result = WmlComparer.Compare(new WmlDocument("before.docx", before),
                                         new WmlDocument("after.docx", after),
                                         settings);

        Assert.NotNull(result);
        Assert.NotNull(result.DocumentByteArray);
    }

    [Fact]
    public void WC_BodyProofErr_001_ProofErrAtBodyLevel_ComparesWithoutException()
    {
        var before = BuildDocxWithBodyLevelElement(
            new XElement(W + "proofErr",
                new XAttribute(W + "type", "spellStart")),
            new XElement(W + "proofErr",
                new XAttribute(W + "type", "spellEnd")),
            paragraphTexts: new[] { "Para A.", "Para B." });
        var after = BuildDocxWithBodyLevelElement(
            bodyLevelBefore: null,
            bodyLevelAfter: null,
            paragraphTexts: new[] { "Para A.", "Para B modified." });

        var settings = new WmlComparerSettings();
        var result = WmlComparer.Compare(new WmlDocument("before.docx", before),
                                         new WmlDocument("after.docx", after),
                                         settings);

        Assert.NotNull(result);
        Assert.NotNull(result.DocumentByteArray);
    }

    [Fact]
    public void WC_BodyBookmark_003_BothSidesHaveBodyLevelBookmark()
    {
        // Both before and after have body-level bookmarks; ensures the producer
        // filters them symmetrically and the consumer doesn't trip on either side.
        var before = BuildDocxWithBodyLevelElement(
            new XElement(W + "bookmarkStart",
                new XAttribute(W + "id", "1"),
                new XAttribute(W + "name", "MARK_BEFORE")),
            new XElement(W + "bookmarkEnd",
                new XAttribute(W + "id", "1")),
            paragraphTexts: new[] { "Heading.", "Body text." });
        var after = BuildDocxWithBodyLevelElement(
            new XElement(W + "bookmarkStart",
                new XAttribute(W + "id", "2"),
                new XAttribute(W + "name", "MARK_AFTER")),
            new XElement(W + "bookmarkEnd",
                new XAttribute(W + "id", "2")),
            paragraphTexts: new[] { "Heading.", "Body text revised." });

        var settings = new WmlComparerSettings();
        var result = WmlComparer.Compare(new WmlDocument("before.docx", before),
                                         new WmlDocument("after.docx", after),
                                         settings);

        Assert.NotNull(result);
        Assert.NotNull(result.DocumentByteArray);
    }

    [Fact]
    public void WC_FinalSectionHeaderFooterReferences_BothSidesPreservedInRedline()
    {
        var before = IrTestDocuments.WithHeaderAndFooter(
            "Header logo placeholder",
            "Footer placeholder",
            "Body before.");
        var after = IrTestDocuments.WithHeaderAndFooter(
            "Header logo placeholder",
            "Footer placeholder",
            "Body after.");

        var settings = new WmlComparerSettings();
        var result = WmlComparer.Compare(before, after, settings);

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var sectPr = wDoc.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!.Element(W + "sectPr")!;

        Assert.NotEmpty(sectPr.Elements(W + "headerReference"));
        Assert.NotEmpty(sectPr.Elements(W + "footerReference"));
    }

    /// <summary>
    /// Build a minimal DOCX whose body contains: paragraph, optional body-level
    /// marker pair, more paragraphs. Marker pair is inserted between the first
    /// and second paragraph when supplied.
    /// </summary>
    private static byte[] BuildDocxWithBodyLevelElement(
        XElement? bodyLevelBefore,
        XElement? bodyLevelAfter,
        string[] paragraphTexts)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var body = new XElement(W + "body");
            // First paragraph
            body.Add(BuildParagraph(paragraphTexts[0]));
            // Body-level markers between para 1 and the rest
            if (bodyLevelBefore != null)
                body.Add(bodyLevelBefore);
            if (bodyLevelAfter != null)
                body.Add(bodyLevelAfter);
            // Remaining paragraphs
            for (int i = 1; i < paragraphTexts.Length; i++)
                body.Add(BuildParagraph(paragraphTexts[i]));

            var docXml = new XElement(W + "document",
                new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                body);
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    private static XElement BuildParagraph(string text)
    {
        return new XElement(W + "p",
            new XElement(W + "r",
                new XElement(W + "t",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    text)));
    }
}
