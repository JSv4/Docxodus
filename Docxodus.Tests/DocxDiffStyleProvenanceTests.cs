#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Style-definition provenance of <see cref="DocxDiff.Compare"/> output. The result's styles part keeps the
/// ORIGINAL (left) document-level defaults, while each shared,
/// presentation-eligible style whose effective formatting differs has its CURRENT payload updated to the
/// RIGHT effective formatting (docDefaults + basedOn chain + own definition resolved), with the left's
/// effective payload archived in a tracked <c>w:rPrChange</c>/<c>w:pPrChange</c> inside the definition.
/// This makes otherwise invisible docDefaults changes reversible without copying an untrackable package part.
/// Right-only styles are copied; left-only styles survive for deleted content.
/// </summary>
public class DocxDiffStyleProvenanceTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument Doc(
        string ddFont, string? normalFont, string text,
        bool includeGlossary = false, bool includeNumbering = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var paragraph = new Paragraph(new Run(new Text(text)));
            if (includeNumbering)
                paragraph.PrependChild(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })));
            mainPart.Document = new Document(new Body(paragraph));
            var normal = normalFont is null
                ? new Style(new StyleName { Val = "Normal" })
                : new Style(new StyleName { Val = "Normal" },
                    new StyleRunProperties(new RunFonts { Ascii = normalFont, HighAnsi = normalFont }));
            normal.Type = StyleValues.Paragraph;
            normal.StyleId = "Normal";
            normal.Default = true;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = ddFont, HighAnsi = ddFont }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()),
                normal);
            if (includeNumbering)
            {
                var numbering = mainPart.AddNewPart<NumberingDefinitionsPart>();
                using var writer = new StreamWriter(numbering.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write($"<w:numbering xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/>" +
                    "<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>" +
                    "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>");
            }
            if (includeGlossary)
            {
                var glossary = mainPart.AddNewPart<GlossaryDocumentPart>();
                using var writer = new StreamWriter(glossary.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write($"<w:glossaryDocument xmlns:w=\"{W.NamespaceName}\"><w:docParts/></w:glossaryDocument>");
            }
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("doc.docx", stream.ToArray());
    }

    private static WmlDocument DocWithStyledBreak(string defaultFont)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(
                new RunProperties(new RunStyle { Val = "BreakStyle" }), new Break()))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:docDefaults><w:rPrDefault><w:rPr>" +
                    $"<w:rFonts w:ascii=\"{defaultFont}\" w:hAnsi=\"{defaultFont}\"/>" +
                    "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
                    "<w:style w:type=\"character\" w:styleId=\"BreakStyle\"><w:name w:val=\"BreakStyle\"/></w:style>" +
                    "</w:styles>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("styled-break.docx", stream.ToArray());
    }

    private static WmlDocument DocWithInputStyleRevision(string defaultFont)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Shared line.")))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:docDefaults><w:rPrDefault><w:rPr>" +
                    $"<w:rFonts w:ascii=\"{defaultFont}\" w:hAnsi=\"{defaultFont}\"/>" +
                    "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/>" +
                    "<w:rPr><w:rFonts w:ascii=\"Legacy\" w:hAnsi=\"Legacy\"/>" +
                    "<w:rPrChange w:id=\"4\" w:author=\"Input Reviewer\" w:date=\"2001-01-01T00:00:00Z\">" +
                    "<w:rPr><w:rFonts w:ascii=\"Archived\" w:hAnsi=\"Archived\"/></w:rPr></w:rPrChange>" +
                    "</w:rPr></w:style></w:styles>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("input-style-revision.docx", stream.ToArray());
    }

    private static WmlDocument DocWithLiteralDefaults(string font, string size, string line, string text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:docDefaults><w:rPrDefault><w:rPr>" +
                    $"<w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\"/><w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>" +
                    "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" +
                    $"<w:spacing w:line=\"{line}\" w:lineRule=\"auto\"/>" +
                    "</w:pPr></w:pPrDefault></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
                    "<w:name w:val=\"Normal\"/></w:style></w:styles>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("literal-defaults.docx", stream.ToArray());
    }

    private static WmlDocument DocWithThemeDefault(string size)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("Shared line.")))));
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:docDefaults><w:rPrDefault><w:rPr>" +
                    $"<w:rFonts w:asciiTheme=\"minorHAnsi\"/><w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>" +
                    "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
                    "<w:name w:val=\"Normal\"/></w:style></w:styles>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("theme-default.docx", stream.ToArray());
    }

    private static WmlDocument DocWithHeaderOnlyStyle(string defaultFont)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "HeaderText" }),
                new Run(new Text("Stable header."))));
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "BodyText" }),
                    new Run(new Text("Stable body."))),
                new SectionProperties(new HeaderReference
                {
                    Id = main.GetIdOfPart(header),
                    Type = HeaderFooterValues.Default,
                })));

            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:docDefaults><w:rPrDefault><w:rPr>" +
                    $"<w:rFonts w:ascii=\"{defaultFont}\" w:hAnsi=\"{defaultFont}\"/>" +
                    "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"BodyText\"><w:name w:val=\"BodyText\"/></w:style>" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"HeaderText\"><w:name w:val=\"HeaderText\"/></w:style>" +
                    "</w:styles>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("header-style.docx", stream.ToArray());
    }

    private static XDocument StylesOf(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.StyleDefinitionsPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }

    private static XElement StyleOf(XDocument styles, string id) =>
        styles.Root!.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == id);

    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    private static List<XElement> BodyParagraphProperties(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        var body = XDocument.Parse(reader.ReadToEnd()).Root!.Element(W + "body")!;
        return body.Elements(W + "p")
            .Select(p => new XElement(p.Element(W + "pPr") ?? new XElement(W + "pPr")))
            .ToList();
    }

    private static List<XElement> BodyDirectParagraphProperties(WmlDocument doc) =>
        BodyParagraphProperties(doc).Select(p =>
        {
            p.Element(W + "pStyle")?.Remove();
            return p;
        }).ToList();

    private static List<XElement> BodyRunProperties(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        var body = XDocument.Parse(reader.ReadToEnd()).Root!.Element(W + "body")!;
        return body.Elements(W + "p")
            .SelectMany(p => p.Descendants(W + "r"))
            .Select(r => new XElement(r.Element(W + "rPr") ?? new XElement(W + "rPr")))
            .ToList();
    }

    private static void AssertDirectParagraphPropertiesEqual(WmlDocument expected, WmlDocument actual)
    {
        var expectedProperties = BodyDirectParagraphProperties(expected);
        var actualProperties = BodyDirectParagraphProperties(actual);
        Assert.Equal(expectedProperties.Count, actualProperties.Count);
        for (int i = 0; i < expectedProperties.Count; i++)
            Assert.True(XNode.DeepEquals(expectedProperties[i], actualProperties[i]),
                $"paragraph {i} direct properties differ\nexpected: {expectedProperties[i]}\nactual: {actualProperties[i]}");
    }

    private static void AssertRunPropertiesEqual(WmlDocument expected, WmlDocument actual)
    {
        var expectedProperties = BodyRunProperties(expected);
        var actualProperties = BodyRunProperties(actual);
        Assert.Equal(expectedProperties.Count, actualProperties.Count);
        for (int i = 0; i < expectedProperties.Count; i++)
            Assert.True(XNode.DeepEquals(expectedProperties[i], actualProperties[i]),
                $"run {i} direct properties differ\nexpected: {expectedProperties[i]}\nactual: {actualProperties[i]}");
    }

    private static void AssertStylesSchemaValid(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var styles = wdoc.MainDocumentPart!.StyleDefinitionsPart!;
        foreach (var version in new[] { FileFormatVersions.Office2010, FileFormatVersions.Office2019 })
        {
            var errors = new OpenXmlValidator(version).Validate(styles)
                .Select(error => $"{error.Id}@{error.Path?.XPath}: {error.Description}")
                .ToList();
            Assert.True(errors.Count == 0,
                $"styles.xml schema errors for {version}:\n{string.Join("\n", errors)}");
        }
    }

    private static void AddImplicitHeader(MainDocumentPart main, Body body)
    {
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text("Retained implicit header."))));
        body.Append(new SectionProperties(new HeaderReference
        {
            Id = main.GetIdOfPart(header),
            Type = HeaderFooterValues.Default,
        }));
    }

    /// <summary>
    /// Synthetic reduction of word_tolerated_misplaced_pgsz → word_tolerated_misplaced_uipriority:
    /// the left package has named LibreOffice-style paragraphs but no docDefaults/default paragraph
    /// style, while the right is a total rewrite with a mix of implicit-default and named styles.
    /// </summary>
    private static WmlDocument StylelessLeft(
        bool leaveOneParagraphImplicit = false,
        bool includeSharedSpine = false,
        string? legacyBasedOn = null,
        string? conflictingCharacterStyleId = null,
        bool includeGlossary = false,
        bool includeImplicitHeader = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Paragraph LeftParagraph(string text, bool implicitStyle = false) => implicitStyle
                ? new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                    new Run(new RunProperties(new Bold()), new Text(text)))
                : new Paragraph(new ParagraphProperties(
                    new ParagraphStyleId { Val = "Legacy" },
                    new Justification { Val = JustificationValues.Right }),
                    new Run(new RunProperties(new Bold()), new Text(text)));

            var body = new Body();
            if (includeSharedSpine)
                body.Append(LeftParagraph("Shared stable spine."));
            // Keep the replacement vocabulary deliberately disjoint from ImportedStyleRight so the
            // aligner emits DeleteBlock + InsertBlock rather than a modified-pair operation.
            body.Append(LeftParagraph("Saffron zeppelin orchard.", leaveOneParagraphImplicit));
            body.Append(LeftParagraph("Cobalt marmot thimble."));
            if (includeImplicitHeader)
                AddImplicitHeader(main, body);
            main.Document = new Document(body);

            // Deliberately no w:docDefaults and no paragraph w:default="1".
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                var basedOn = legacyBasedOn is null ? "" : $"<w:basedOn w:val=\"{legacyBasedOn}\"/>";
                var conflictingStyle = conflictingCharacterStyleId is null
                    ? ""
                    : $"<w:style w:type=\"character\" w:styleId=\"{conflictingCharacterStyleId}\">" +
                      $"<w:name w:val=\"{conflictingCharacterStyleId}\"/></w:style>";
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"Legacy\"><w:name w:val=\"Legacy\"/>" + basedOn +
                    "<w:pPr><w:jc w:val=\"right\"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>" +
                    conflictingStyle + "</w:styles>");
            }
            if (includeGlossary)
            {
                var glossary = main.AddNewPart<GlossaryDocumentPart>();
                using var writer = new StreamWriter(glossary.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write($"<w:glossaryDocument xmlns:w=\"{W.NamespaceName}\"/>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("left.docx", stream.ToArray());
    }

    private static WmlDocument ImportedStyleRight(
        bool includeSharedSpine = false,
        string? orderedBasedOn = null,
        bool useExplicitStylesOnly = false,
        bool normalHasInputPropertyRevisions = false,
        bool includeImplicitHeader = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            if (includeSharedSpine)
                body.Append(new Paragraph(new ParagraphProperties(
                    new ParagraphStyleId { Val = "Legacy" },
                    new Justification { Val = JustificationValues.Right }),
                    new Run(new RunProperties(new Bold()), new Text("Shared stable spine."))));
            // No pStyle: this must pull the copied right default paragraph style into the guarded set.
            if (!useExplicitStylesOnly)
                body.Append(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                    new Run(new RunProperties(new Italic()), new Text("Vermilion kettle lattice."))));
            body.Append(new Paragraph(new ParagraphProperties(
                new ParagraphStyleId { Val = "Imported" },
                new Indentation { Left = "360" }),
                new Run(new RunProperties(new Underline { Val = UnderlineValues.Single }),
                    new Text("Umber sparrow fjord."))));
            // Keep a third inserted paragraph so Imported is not the Word-shaped insert/delete seam;
            // this style deliberately has ind + jc but no spacing, exercising CT_PPr child order.
            body.Append(new Paragraph(new ParagraphProperties(
                new ParagraphStyleId { Val = "Ordered" },
                new Justification { Val = JustificationValues.Right }),
                new Run(new RunProperties(new SmallCaps()), new Text("Azure nickel quasar."))));
            // The final inserted paragraph becomes the Word-shaped insert/delete seam. Its direct
            // alignment matches the left tail, keeping the fixture focused on the style definition.
            body.Append(new Paragraph(useExplicitStylesOnly
                ? new ParagraphProperties(new ParagraphStyleId { Val = "Ordered" },
                    new Justification { Val = JustificationValues.Right })
                : new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                new Run(new RunProperties(new Caps()), new Text("Ivory lantern tundra."))));
            if (includeImplicitHeader)
                AddImplicitHeader(main, body);
            main.Document = new Document(body);

            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using (var writer = new StreamWriter(styles.GetStream(FileMode.Create, FileAccess.Write)))
            {
                const string InputPPrChange =
                    "<w:pPrChange w:id=\"91\" w:author=\"Input Reviewer\" w:date=\"2001-01-01T00:00:00Z\">" +
                    "<w:pPr><w:jc w:val=\"left\"/></w:pPr></w:pPrChange>";
                const string InputRPrChange =
                    "<w:rPrChange w:id=\"92\" w:author=\"Input Reviewer\" w:date=\"2001-01-01T00:00:00Z\">" +
                    "<w:rPr><w:i/></w:rPr></w:rPrChange>";
                var normalPPr = normalHasInputPropertyRevisions
                    ? "<w:pPr>" + InputPPrChange + "</w:pPr>"
                    : "";
                var normalRPrChange = normalHasInputPropertyRevisions ? InputRPrChange : "";
                var orderedBasedOnMarkup = orderedBasedOn is null
                    ? ""
                    : $"<w:basedOn w:val=\"{orderedBasedOn}\"/>";
                writer.Write($"<w:styles xmlns:w=\"{W.NamespaceName}\">" +
                    "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/>" +
                    "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>" +
                    "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
                    "<w:name w:val=\"Normal\"/>" + normalPPr +
                    "<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/>" +
                    "<w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/>" + normalRPrChange + "</w:rPr></w:style>" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"Legacy\"><w:name w:val=\"Legacy\"/>" +
                    "<w:pPr><w:jc w:val=\"right\"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"Imported\"><w:name w:val=\"Imported\"/>" +
                    "<w:pPr><w:spacing w:before=\"120\" w:after=\"60\"/></w:pPr>" +
                    "<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/>" +
                    "</w:rPr></w:style>" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"Ordered\"><w:name w:val=\"Ordered\"/>" + orderedBasedOnMarkup +
                    "<w:pPr><w:ind w:left=\"480\"/><w:jc w:val=\"right\"/></w:pPr>" +
                    "<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:lang w:val=\"en-US\"/></w:rPr></w:style>" +
                    "<w:style w:type=\"paragraph\" w:styleId=\"Unused\"><w:name w:val=\"Unused\"/>" +
                    "<w:pPr><w:spacing w:after=\"720\"/></w:pPr></w:style></w:styles>");
            }
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("right.docx", stream.ToArray());
    }

    [Fact]
    public void Output_KeepsLeftDocDefaults()
    {
        var left = Doc("Courier New", null, "Shared line.");
        var right = Doc("Arial", null, "Shared line.");

        var result = DocxDiff.Compare(left, right);

        var dd = StylesOf(result).Root!
            .Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr")
            ?.Element(W + "rFonts");
        Assert.Equal("Courier New", (string?)dd?.Attribute(W + "ascii"));
    }

    [Fact]
    public void SharedStyleWithEqualDefinitions_ProjectsEffectiveRunFormatting_WhenDocDefaultsDiffer()
    {
        // Both Normals are formatting-empty. The style-level projection keeps the output docDefaults
        // left-owned, then makes Normal's current rPr right-effective and archives the left-effective rPr.
        var left = Doc("Courier New", null, "Shared line.");
        var right = Doc("Arial", null, "Shared line.");

        var result = DocxDiff.Compare(left, right);

        var normal = StyleOf(StylesOf(result), "Normal");
        var rPr = normal.Element(W + "rPr");
        Assert.Equal("Arial", (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("Courier New", (string?)rPr?.Element(W + "rPrChange")?.Element(W + "rPr")
            ?.Element(W + "rFonts")?.Attribute(W + "ascii"));
    }

    [Fact]
    public void DocDefaultsOnly_SharedNormalProjectsRightPresentation_AndRoundTrips()
    {
        var left = DocWithLiteralDefaults("Courier New", "22", "240", "Shared line.");
        var right = DocWithLiteralDefaults("Arial", "28", "360", "Shared line.");

        var result = DocxDiff.Compare(left, right);
        AssertStylesSchemaValid(result);

        // The package-level defaults deliberately remain left-owned: revisions cannot switch a package part.
        var outputDefaults = StylesOf(result).Root!.Element(W + "docDefaults")!;
        Assert.Equal("Courier New", (string?)outputDefaults.Element(W + "rPrDefault")?.Element(W + "rPr")
            ?.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("240", (string?)outputDefaults.Element(W + "pPrDefault")?.Element(W + "pPr")
            ?.Element(W + "spacing")?.Attribute(W + "line"));

        var normal = StyleOf(StylesOf(result), "Normal");
        var currentPPr = normal.Element(W + "pPr")!;
        var currentRPr = normal.Element(W + "rPr")!;
        Assert.Equal("360", (string?)currentPPr.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.Equal("240", (string?)currentPPr.Element(W + "pPrChange")?.Element(W + "pPr")
            ?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.Equal("Arial", (string?)currentRPr.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("28", (string?)currentRPr.Element(W + "sz")?.Attribute(W + "val"));
        Assert.Equal("Courier New", (string?)currentRPr.Element(W + "rPrChange")?.Element(W + "rPr")
            ?.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("22", (string?)currentRPr.Element(W + "rPrChange")?.Element(W + "rPr")
            ?.Element(W + "sz")?.Attribute(W + "val"));

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));

        var acceptedNormal = StyleOf(StylesOf(accepted), "Normal");
        var rejectedNormal = StyleOf(StylesOf(rejected), "Normal");
        Assert.Equal("360", (string?)acceptedNormal.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.Equal("Arial", (string?)acceptedNormal.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("28", (string?)acceptedNormal.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"));
        Assert.Equal("240", (string?)rejectedNormal.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.Equal("Courier New", (string?)rejectedNormal.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("22", (string?)rejectedNormal.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"));
    }

    [Fact]
    public void DocDefaultsProjection_DeclinesThemeReferences_WithoutCopyingPackagePresentation()
    {
        var left = DocWithThemeDefault("22");
        var right = DocWithThemeDefault("28");

        var result = DocxDiff.Compare(left, right);

        var normal = StyleOf(StylesOf(result), "Normal");
        Assert.Null(normal.Element(W + "rPr"));
        var defaults = StylesOf(result).Root!.Element(W + "docDefaults")!;
        Assert.Equal("22", (string?)defaults.Element(W + "rPrDefault")?.Element(W + "rPr")
            ?.Element(W + "sz")?.Attribute(W + "val"));
    }

    [Fact]
    public void DocDefaultsProjection_TracksAStyleUsedOnlyByAnUnchangedHeader()
    {
        var left = DocWithHeaderOnlyStyle("Courier New");
        var right = DocWithHeaderOnlyStyle("Arial");

        var result = DocxDiff.Compare(left, right);
        var headerStyle = StyleOf(StylesOf(result), "HeaderText");

        Assert.Equal("Arial", (string?)headerStyle.Element(W + "rPr")?.Element(W + "rFonts")
            ?.Attribute(W + "ascii"));
        Assert.Equal("Courier New", (string?)headerStyle.Element(W + "rPr")?.Element(W + "rPrChange")
            ?.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii"));
    }

    [Fact]
    public void DocDefaultsProjection_TracksCharacterStyleOnBreakOnlyRun()
    {
        // IrBreak has no IrFormat, but its source w:r can still have an active w:rStyle. The projection
        // must find that style reference so the accepted break paragraph does not retain left defaults.
        var left = DocWithStyledBreak("Courier New");
        var right = DocWithStyledBreak("Arial");

        var result = DocxDiff.Compare(left, right);
        var breakStyle = StyleOf(StylesOf(result), "BreakStyle");

        Assert.Equal("Arial", (string?)breakStyle.Element(W + "rPr")?.Element(W + "rFonts")
            ?.Attribute(W + "ascii"));
        Assert.Equal("Courier New", (string?)breakStyle.Element(W + "rPr")?.Element(W + "rPrChange")
            ?.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii"));
    }

    [Fact]
    public void DocDefaultsProjection_PreservesInputStylePropertyRevisions()
    {
        // Equal raw payloads intentionally ignore rPrChange, so the projection must decline rather than
        // replacing the entire style rPr and silently dropping existing review history.
        var left = DocWithInputStyleRevision("Courier New");
        var right = DocWithInputStyleRevision("Arial");

        var result = DocxDiff.Compare(left, right);
        var normalRPr = StyleOf(StylesOf(result), "Normal").Element(W + "rPr")!;

        Assert.Equal("Legacy", (string?)normalRPr.Element(W + "rFonts")?.Attribute(W + "ascii"));
        Assert.Equal("Input Reviewer", (string?)normalRPr.Element(W + "rPrChange")?.Attribute(W + "author"));
        Assert.Equal("Archived", (string?)normalRPr.Element(W + "rPrChange")?.Element(W + "rPr")
            ?.Element(W + "rFonts")?.Attribute(W + "ascii"));
    }

    [Fact]
    public void DocDefaultsProjection_DeclinesListsWhoseLabelsCanInheritDefaults()
    {
        var left = Doc("Courier New", null, "One", includeNumbering: true);
        var right = Doc("Arial", null, "One", includeNumbering: true);

        var result = DocxDiff.Compare(left, right);

        Assert.Null(StyleOf(StylesOf(result), "Normal").Element(W + "rPr"));
    }

    [Fact]
    public void DocDefaultsProjection_DeclinesWhenLeftOwnsAGlossary()
    {
        var left = Doc("Courier New", null, "Shared line.", includeGlossary: true);
        var right = Doc("Arial", null, "Shared line.");

        var result = DocxDiff.Compare(left, right);

        Assert.Null(StyleOf(StylesOf(result), "Normal").Element(W + "rPr"));
    }

    [Fact]
    public void SharedStyleWithDifferingDefinition_UpdatesToRightEffective_AndTracksOldPayload()
    {
        var left = Doc("Courier New", "Consolas", "Shared line.");
        var right = Doc("Calibri", "Arial", "Shared line.");

        var result = DocxDiff.Compare(left, right);

        var normal = StyleOf(StylesOf(result), "Normal");
        var rPr = normal.Element(W + "rPr");
        Assert.NotNull(rPr);
        // Current payload = right's EFFECTIVE formatting (its own def wins over its docDefaults).
        Assert.Equal("Arial", (string?)rPr!.Element(W + "rFonts")?.Attribute(W + "ascii"));
        // Old payload archived in a tracked rPrChange, carrying the left's effective fonts.
        var change = rPr.Element(W + "rPrChange");
        Assert.NotNull(change);
        Assert.Equal("Consolas",
            (string?)change!.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii"));
    }

    [Fact]
    public void RightOnlyStyle_IsCopied_AndLeftOnlyStyleSurvives()
    {
        static WmlDocument WithExtraStyle(string ddFont, string extraId, string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = extraId }),
                    new Run(new Text(text)))));
                mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                    new DocDefaults(
                        new RunPropertiesDefault(new RunPropertiesBaseStyle(
                            new RunFonts { Ascii = ddFont, HighAnsi = ddFont }, new FontSize { Val = "22" })),
                        new ParagraphPropertiesDefault()),
                    new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                    new Style(new StyleName { Val = extraId }, new StyleRunProperties(new Italic()))
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = extraId,
                    });
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("doc.docx", stream.ToArray());
        }

        var left = WithExtraStyle("Courier New", "LeftOnly", "Old text entirely.");
        var right = WithExtraStyle("Arial", "RightOnly", "Completely new words.");

        var result = DocxDiff.Compare(left, right);

        var styles = StylesOf(result);
        Assert.Contains(styles.Root!.Elements(W + "style"), s => (string?)s.Attribute(W + "styleId") == "LeftOnly");
        Assert.Contains(styles.Root!.Elements(W + "style"), s => (string?)s.Attribute(W + "styleId") == "RightOnly");

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void FullReplacement_StylelessLeft_NormalizesOnlyUsedInsertedParagraphStyles_AndRoundTripsDirectFormatting()
    {
        var left = StylelessLeft();
        var right = ImportedStyleRight();

        var result = DocxDiff.Compare(left, right);
        var styles = StylesOf(result);
        AssertStylesSchemaValid(result);
        var normal = StyleOf(styles, "Normal");
        var imported = StyleOf(styles, "Imported");
        var unused = StyleOf(styles, "Unused");

        // The implicit right paragraphs need the copied default style; named inserted paragraphs get
        // the same compact line metric. Unused right-only styles remain raw-copied.
        Assert.NotNull(normal.Attribute(W + "default"));
        Assert.Equal("240", (string?)normal.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.NotNull(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
        Assert.Equal("60", (string?)imported.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "after"));
        Assert.Equal("240", (string?)imported.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.NotNull(imported.Element(W + "rPr")?.Element(W + "rPrChange"));
        Assert.Null(unused.Element(W + "pPr")?.Element(W + "pPrChange"));
        Assert.Null(unused.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));

        // Capture the archived style payloads before processors mutate the package.  This is the
        // critical semantic test: RevisionProcessor accepts a style-property revision by keeping
        // current pPr/rPr, and rejects it by replacing the entire property element with this inner
        // payload (not by merely removing the marker).
        var normalPriorPPr = new XElement(normal.Element(W + "pPr")!.Element(W + "pPrChange")!.Element(W + "pPr")!);
        var normalPriorRPr = new XElement(normal.Element(W + "rPr")!.Element(W + "rPrChange")!.Element(W + "rPr")!);
        var importedPriorPPr = new XElement(imported.Element(W + "pPr")!.Element(W + "pPrChange")!.Element(W + "pPr")!);
        var importedPriorRPr = new XElement(imported.Element(W + "rPr")!.Element(W + "rPrChange")!.Element(W + "rPr")!);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
        AssertDirectParagraphPropertiesEqual(right, accepted);
        AssertDirectParagraphPropertiesEqual(left, rejected);
        AssertRunPropertiesEqual(right, accepted);
        AssertRunPropertiesEqual(left, rejected);

        var acceptedStyles = StylesOf(accepted);
        var rejectedStyles = StylesOf(rejected);
        var acceptedNormal = StyleOf(acceptedStyles, "Normal");
        var acceptedImported = StyleOf(acceptedStyles, "Imported");
        var rejectedNormal = StyleOf(rejectedStyles, "Normal");
        var rejectedImported = StyleOf(rejectedStyles, "Imported");
        Assert.Null(acceptedNormal.Descendants(W + "pPrChange").FirstOrDefault());
        Assert.Null(acceptedNormal.Descendants(W + "rPrChange").FirstOrDefault());
        Assert.Null(acceptedImported.Descendants(W + "pPrChange").FirstOrDefault());
        Assert.Null(acceptedImported.Descendants(W + "rPrChange").FirstOrDefault());
        Assert.Equal("240", (string?)acceptedImported.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.True(XNode.DeepEquals(normalPriorPPr, rejectedNormal.Element(W + "pPr")));
        Assert.True(XNode.DeepEquals(normalPriorRPr, rejectedNormal.Element(W + "rPr")));
        Assert.True(XNode.DeepEquals(importedPriorPPr, rejectedImported.Element(W + "pPr")));
        Assert.True(XNode.DeepEquals(importedPriorRPr, rejectedImported.Element(W + "rPr")));
        Assert.Equal("Imported", (string?)BodyParagraphProperties(accepted)[1]
            .Element(W + "pStyle")?.Attribute(W + "val"));
    }

    [Fact]
    public void ImportedStyleNormalization_DeclinesWhenLeftHasImplicitParagraphOrBodySpine()
    {
        // Either condition makes a right default style observable after reject (implicit left p), or
        // proves the comparison is not a full-body replacement (equal spine), so the general raw-copy
        // behavior deliberately remains in force.
        foreach (var pair in new[]
        {
            (Left: StylelessLeft(leaveOneParagraphImplicit: true), Right: ImportedStyleRight()),
            (Left: StylelessLeft(includeSharedSpine: true), Right: ImportedStyleRight(includeSharedSpine: true)),
        })
        {
            var normal = StyleOf(StylesOf(DocxDiff.Compare(pair.Left, pair.Right)), "Normal");
            Assert.Null(normal.Attribute(W + "default"));
            Assert.Null(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
        }
    }

    [Fact]
    public void ImportedStyleNormalization_DeclinesWhenLeftPStyleBasedOnIsMissing()
    {
        // Legacy is explicit on every left paragraph, but its missing Normal parent would resolve
        // only after the right default style is copied. The guarded projection must not make that
        // foreign style observable after rejecting the body replacement.
        var left = StylelessLeft(legacyBasedOn: "Normal");
        var normal = StyleOf(StylesOf(DocxDiff.Compare(left, ImportedStyleRight())), "Normal");

        Assert.Null(normal.Attribute(W + "default"));
        Assert.Null(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    [Fact]
    public void ImportedStyleNormalization_DeclinesWhenUsedRightStyleChainIsUnresolved()
    {
        // Ordered is used by inserted paragraphs. Its malformed right basedOn must not be made
        // partially observable by normalizing other styles in the package.
        var normal = StyleOf(StylesOf(DocxDiff.Compare(
            StylelessLeft(), ImportedStyleRight(orderedBasedOn: "Missing"))), "Normal");

        Assert.Null(normal.Attribute(W + "default"));
        Assert.Null(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    [Fact]
    public void ImportedStyleNormalization_DeclinesWhenUsedRightStyleIdCollidesWithLeftCharacterStyle()
    {
        // styleId is package-global even though the general merger locates existing styles by
        // (type, styleId). Do not retain a normalized right Imported paragraph style alongside this
        // unrelated LEFT character style with the same id.
        var styles = StylesOf(DocxDiff.Compare(
            StylelessLeft(conflictingCharacterStyleId: "Imported"), ImportedStyleRight()));
        var normal = StyleOf(styles, "Normal");

        Assert.Null(normal.Attribute(W + "default"));
        Assert.Null(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    [Fact]
    public void ImportedStyleNormalization_DeclinesWhenLeftHasGlossaryPart()
    {
        // The renderer preserves glossary building blocks untouched. Until their style reachability
        // is modeled, they disqualify the special projection rather than silently acquiring the
        // copied right default after a body rejection.
        var styles = StylesOf(DocxDiff.Compare(StylelessLeft(includeGlossary: true), ImportedStyleRight()));
        var normal = StyleOf(styles, "Normal");

        Assert.Null(normal.Attribute(W + "default"));
        Assert.Null(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    [Fact]
    public void ImportedStyleNormalization_DeclinesWhenRetainedHeaderHasImplicitParagraph()
    {
        // Both headers are byte-for-byte equivalent, so this is not a header edit. The unchanged
        // LEFT header still has an implicit paragraph that would observe a right default after a
        // rejected body replacement, and must therefore block the special projection.
        var styles = StylesOf(DocxDiff.Compare(
            StylelessLeft(includeImplicitHeader: true),
            ImportedStyleRight(includeImplicitHeader: true)));
        var normal = StyleOf(styles, "Normal");

        Assert.Null(normal.Attribute(W + "default"));
        Assert.Null(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    [Fact]
    public void FullReplacement_NormalizesAndRetainsDefaultBasedOnAncestor()
    {
        // No paragraph is implicitly styled here: Normal is reachable only through Ordered's
        // right-side basedOn chain. It still needs both normalization and w:default retention.
        var styles = StylesOf(DocxDiff.Compare(
            StylelessLeft(), ImportedStyleRight(orderedBasedOn: "Normal", useExplicitStylesOnly: true)));
        var normal = StyleOf(styles, "Normal");
        var ordered = StyleOf(styles, "Ordered");

        Assert.Equal("Normal", (string?)ordered.Element(W + "basedOn")?.Attribute(W + "val"));
        Assert.NotNull(normal.Attribute(W + "default"));
        Assert.Equal("240", (string?)normal.Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "line"));
        Assert.NotNull(normal.Element(W + "pPr")?.Element(W + "pPrChange"));
    }

    [Fact]
    public void FullReplacement_PreservesInputPropertyRevisionsOnUsedDefaultStyle()
    {
        var result = DocxDiff.Compare(
            StylelessLeft(),
            ImportedStyleRight(normalHasInputPropertyRevisions: true),
            new DocxDiffSettings { PreserveInputRevisions = true });
        var normal = StyleOf(StylesOf(result), "Normal");
        var pPr = normal.Element(W + "pPr")!;
        var rPr = normal.Element(W + "rPr")!;

        // RawStylePayload removes property-change markers. A used right default that already has
        // them therefore stays raw-copied, including its default role, instead of being rewritten
        // with this comparison's synthetic style revision.
        Assert.NotNull(normal.Attribute(W + "default"));
        Assert.Null(pPr.Element(W + "spacing"));
        Assert.Equal("Input Reviewer", (string?)pPr.Element(W + "pPrChange")?.Attribute(W + "author"));
        Assert.Equal("Input Reviewer", (string?)rPr.Element(W + "rPrChange")?.Attribute(W + "author"));
        Assert.Single(pPr.Elements(W + "pPrChange"));
        Assert.Single(rPr.Elements(W + "rPrChange"));
    }

    [Fact]
    public void FullReplacement_InsertedStyleSpacingPrecedesIndAndJc()
    {
        var styles = StylesOf(DocxDiff.Compare(StylelessLeft(), ImportedStyleRight()));
        var orderedPPr = StyleOf(styles, "Ordered").Element(W + "pPr")!;

        // CT_PPrBase order is ... spacing, ind, ..., jc, ..., pPrChange. A previous implementation
        // appended synthesized spacing after ind/jc, which is schema-invalid even though it renders.
        Assert.Equal(
            new[] { "spacing", "ind", "jc", "pPrChange" },
            orderedPPr.Elements().Select(e => e.Name.LocalName).ToArray());
    }

    [Fact]
    public void FullReplacement_InsertedStyleKernPrecedesLangWhenStyleHasNoSize()
    {
        var styles = StylesOf(DocxDiff.Compare(StylelessLeft(), ImportedStyleRight()));
        var orderedRPr = StyleOf(styles, "Ordered").Element(W + "rPr")!;
        var names = orderedRPr.Elements().Select(e => e.Name.LocalName).ToList();
        var rFontsIndex = names.IndexOf("rFonts");
        var kernIndex = names.IndexOf("kern");
        var langIndex = names.IndexOf("lang");

        // Ordered deliberately has rFonts + lang but no size in the right source. CT_RPr requires
        // synthesized kern between them; appending after lang makes styles.xml schema-invalid.
        Assert.True(rFontsIndex >= 0 && kernIndex >= 0 && langIndex >= 0);
        Assert.True(rFontsIndex < kernIndex && kernIndex < langIndex);
    }

}
