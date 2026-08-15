// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace LegalEval;

/// <summary>
/// Builds the reviewable, repository-authored fixture recipe into a DOCX. The JSON recipe keeps
/// the legal text reviewable alongside the pinned binary oracle and provides an explicit
/// reproducibility check. The generated package is intentionally small but carries the structures
/// that real agreement workflows most often damage (numbering, tables, comments, revisions,
/// controls, notes, fields, running content, signatures, and two sections).
/// </summary>
public static class LegalFixtureFactory
{
    private static readonly DateTime FixedDate =
        new(2026, 1, 15, 12, 0, 0, DateTimeKind.Utc);

    public static byte[] Build(string recipePath)
    {
        var recipe = JsonNode.Parse(File.ReadAllText(recipePath)) as JsonObject
            ?? throw new ScenarioValidationException($"{recipePath}: fixture root must be an object");
        if (String(recipe, "schemaVersion") != "1.0"
            || String(recipe, "fixtureType") != "legal-services-agreement-v1")
            throw new ScenarioValidationException(
                $"{recipePath}: unsupported fixture recipe version or fixtureType");
        var document = Object(recipe, "document");

        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            main.Document = new W.Document();
            main.AddNewPart<StyleDefinitionsPart>("rIdStyles").Styles = Styles();
            main.AddNewPart<NumberingDefinitionsPart>("rIdNumbering").Numbering = Numbering();
            main.AddNewPart<DocumentSettingsPart>("rIdSettings").Settings = new W.Settings(
                new W.UpdateFieldsOnOpen { Val = true });

            var header = main.AddNewPart<HeaderPart>("rIdHeaderDefault");
            header.Header = new W.Header(Paragraph(String(document, "headerText"), "Header"));
            var footer = main.AddNewPart<FooterPart>("rIdFooterDefault");
            footer.Footer = new W.Footer(new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Footer" }),
                new W.Run(new W.Text(String(document, "footerText") + " | Page ")),
                new W.SimpleField(new W.Run(new W.Text("1"))) { Instruction = " PAGE " }));

            var body = new W.Body();
            body.Append(Paragraph(String(document, "title"), "AgreementTitle"));
            body.Append(Paragraph(String(document, "intro"), "AgreementBody"));
            body.Append(Paragraph("1. DEFINITIONS", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "definedTermClause"), 1));
            body.Append(NumberedParagraph(String(document, "definedTermUse"), 9));
            body.Append(Paragraph(String(document, "definedTermDecoy"), "AgreementBody"));
            body.Append(Paragraph("2. SERVICES", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "servicesClause"), 2));
            body.Append(Paragraph("3. FEES AND PAYMENT", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "feesClause"), 3));
            body.Append(EconomicsTable(Array(document, "economicsTable")));
            body.Append(Paragraph("4. CONFIDENTIALITY", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "confidentialityClause"), 4));
            body.Append(Paragraph(String(document, "crossReferenceClause"), "AgreementBody"));
            body.Append(Paragraph("5. NOTICES", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "noticeClause"), 5));
            body.Append(Paragraph(String(document, "noticeDecoy"), "AgreementBody"));
            body.Append(Paragraph("6. LIMITATION OF LIABILITY", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "liabilityClause"), 6));
            body.Append(Paragraph("7. GOVERNING LAW", "AgreementHeading"));
            body.Append(NumberedParagraph(String(document, "governingLawClause"), 7));
            body.Append(Paragraph("8. CLIENT INFORMATION", "AgreementHeading"));
            body.Append(ClientNameControl(Object(document, "contentControl")));

            // A paragraph-level section break makes all following signature material section 2.
            var signatureHeading = Paragraph("SIGNATURES", "AgreementHeading");
            signatureHeading.ParagraphProperties!.Append(SectionProperties(main, firstSection: true));
            body.Append(signatureHeading);
            foreach (var signature in Array(document, "signatureLines"))
                body.Append(Paragraph(signature!.GetValue<string>(), "SignatureLine"));
            body.Append(Paragraph("Exhibit A — Service Levels", "AgreementHeading"));
            body.Append(Paragraph(String(document, "exhibitText"), "AgreementBody"));
            body.Append(SectionProperties(main, firstSection: false));
            main.Document.Append(body);
            main.Document.Save();
        }

        var clean = stream.ToArray();
        using var session = new DocxSession(clean, new DocxSessionSettings
        {
            CaptureInitialProjection = true,
            PersistAnchorIds = false,
        });

        var confidentiality = RequireAnchor(session, "Provider shall keep Client Confidential Information");
        var confidentialityText = session.Project().AnchorIndex[confidentiality].TextPreview;
        var confidentialityWord = confidentialityText.IndexOf("Confidential", StringComparison.Ordinal);
        Ensure(session.AddBookmark("Confidentiality",
            DocumentRange.In(confidentiality, new CharSpan(confidentialityWord, "Confidentiality".Length))),
            "add fixture bookmark");

        var crossReference = RequireAnchor(session, "See Section 4 for confidentiality obligations");
        var crossReferenceText = session.Project().AnchorIndex[crossReference].TextPreview;
        var sectionFour = crossReferenceText.IndexOf("Section 4", StringComparison.Ordinal);
        Ensure(session.AddHyperlink(crossReference, new CharSpan(sectionFour, "Section 4".Length),
            HyperlinkTarget.Internal("Confidentiality")), "add fixture cross-reference");

        var services = RequireAnchor(session, "commercially reasonable efforts");
        Ensure(session.InsertFootnote(services,
            session.Project().AnchorIndex[services].TextPreview.IndexOf("Provider", StringComparison.Ordinal)
                + "Provider".Length,
            "Service levels are measured monthly in the Client's primary production environment."),
            "add fixture footnote");

        var liability = RequireAnchor(session, "aggregate liability");
        Ensure(session.AddComment(liability, null, "Deal Team",
            "Confirm the negotiated liability cap before execution.", "DT", FixedDate),
            "add fixture comment");

        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        session.SetRevisionAuthor("Prior Counsel");
        var servicesText = session.Project().AnchorIndex[services].TextPreview;
        var phrase = "commercially reasonable efforts";
        var phraseOffset = servicesText.IndexOf(phrase, StringComparison.Ordinal);
        Ensure(session.ReplaceTextAtSpan(services, phraseOffset, phrase.Length, "reasonable efforts"),
            "add fixture pre-existing revision");

        var bytes = session.Save(false);
        return GeneratedPackageNormalizer.Normalize(bytes);
    }

    private static W.Styles Styles()
    {
        var styles = new W.Styles();
        styles.Append(Style("Normal", "Normal", isDefault: true));
        styles.Append(Style("AgreementBody", "Agreement Body"));
        styles.Append(Style("AgreementTitle", "Agreement Title", bold: true, size: "32"));
        styles.Append(Style("AgreementHeading", "Agreement Heading", bold: true, size: "24"));
        styles.Append(Style("AgreementClause", "Agreement Clause"));
        styles.Append(Style("SignatureLine", "Signature Line"));
        styles.Append(Style("Header", "Header"));
        styles.Append(Style("Footer", "Footer"));
        return styles;
    }

    private static W.Style Style(
        string id, string name, bool isDefault = false, bool bold = false, string? size = null)
    {
        var style = new W.Style(new W.StyleName { Val = name })
        {
            Type = W.StyleValues.Paragraph,
            StyleId = id,
            Default = isDefault,
        };
        if (bold || size is not null)
        {
            var properties = new W.StyleRunProperties();
            if (bold) properties.Append(new W.Bold());
            if (size is not null) properties.Append(new W.FontSize { Val = size });
            style.Append(properties);
        }
        return style;
    }

    private static W.Numbering Numbering()
    {
        var abstractNumber = new W.AbstractNum(
            new W.MultiLevelType { Val = W.MultiLevelValues.Multilevel },
            NumberingLevel(0, "%1.", "0"),
            NumberingLevel(1, "%1.%2", "720"))
        { AbstractNumberId = 17 };
        return new W.Numbering(
            abstractNumber,
            new W.NumberingInstance(new W.AbstractNumId { Val = 17 }) { NumberID = 42 });
    }

    private static W.Level NumberingLevel(int index, string text, string left) =>
        new(
            new W.StartNumberingValue { Val = 1 },
            new W.NumberingFormat { Val = W.NumberFormatValues.Decimal },
            new W.LevelText { Val = text },
            new W.PreviousParagraphProperties(new W.Indentation { Left = left, Hanging = "360" }))
        { LevelIndex = index };

    private static W.Paragraph Paragraph(string text, string style) =>
        new(
            new W.ParagraphProperties(new W.ParagraphStyleId { Val = style }),
            new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static W.Paragraph NumberedParagraph(string text, int ordinal) =>
        new(
            new W.ParagraphProperties(
                new W.ParagraphStyleId { Val = "AgreementClause" },
                new W.NumberingProperties(
                    new W.NumberingLevelReference { Val = 0 },
                    new W.NumberingId { Val = 42 }),
                new W.ParagraphMarkRunProperties(new W.Vanish())),
            new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }),
            new W.BookmarkStart { Name = $"Clause_{ordinal}", Id = (100 + ordinal).ToString() },
            new W.BookmarkEnd { Id = (100 + ordinal).ToString() });

    private static W.Table EconomicsTable(JsonArray rows)
    {
        var table = new W.Table(new W.TableProperties(
            new W.TableStyle { Val = "TableGrid" },
            new W.TableWidth { Type = W.TableWidthUnitValues.Pct, Width = "5000" }));
        table.Append(new W.TableGrid(
            new W.GridColumn { Width = "2400" },
            new W.GridColumn { Width = "2400" },
            new W.GridColumn { Width = "2400" }));
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var values = rows[rowIndex] as JsonArray
                ?? throw new ScenarioValidationException("economicsTable rows must be arrays");
            var row = new W.TableRow();
            if (rowIndex == 0)
                row.AppendChild(new W.TableRowProperties(new W.TableHeader()));
            foreach (var value in values)
            {
                var paragraph = Paragraph(value?.GetValue<string>() ?? string.Empty, "AgreementBody");
                if (rowIndex == 0)
                    paragraph.Descendants<W.Run>().First().RunProperties = new W.RunProperties(new W.Bold());
                row.Append(new W.TableCell(
                    new W.TableCellProperties(new W.TableCellWidth
                        { Type = W.TableWidthUnitValues.Dxa, Width = "2400" }),
                    paragraph));
            }
            table.Append(row);
        }
        return table;
    }

    private static W.SdtBlock ClientNameControl(JsonObject control) =>
        new(
            new W.SdtProperties(
                new W.SdtAlias { Val = String(control, "alias") },
                new W.Tag { Val = String(control, "tag") },
                new W.SdtId { Val = Int(control, "id") },
                new W.SdtPlaceholder(new W.DocPartReference { Val = "DefaultPlaceholder_1081868562" }),
                new W.ShowingPlaceholder()),
            new W.SdtContentBlock(Paragraph(String(control, "placeholder"), "AgreementBody")));

    private static W.SectionProperties SectionProperties(MainDocumentPart main, bool firstSection)
    {
        var section = new W.SectionProperties();
        section.Append(
            new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = "rIdHeaderDefault" },
            new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = "rIdFooterDefault" });
        if (firstSection)
            section.Append(new W.SectionType { Val = W.SectionMarkValues.NextPage });
        section.Append(
            new W.PageSize { Width = 12240, Height = 15840 },
            new W.PageMargin
            {
                Top = 1080,
                Right = 1080,
                Bottom = 1080,
                Left = 1080,
                Header = 360,
                Footer = 360,
                Gutter = 0,
            });
        return section;
    }

    private static string RequireAnchor(DocxSession session, string text) =>
        session.FindAllByText(text).Single().Anchor.Id;

    private static void Ensure(EditResult result, string operation)
    {
        if (!result.Success)
            throw new InvalidOperationException($"Fixture operation '{operation}' failed: {result.Error?.Message}");
    }

    private static JsonObject Object(JsonObject parent, string name) =>
        parent[name] as JsonObject
            ?? throw new ScenarioValidationException($"fixture property '{name}' must be an object");

    private static JsonArray Array(JsonObject parent, string name) =>
        parent[name] as JsonArray
            ?? throw new ScenarioValidationException($"fixture property '{name}' must be an array");

    private static string String(JsonObject parent, string name) =>
        parent[name]?.GetValue<string>()
            ?? throw new ScenarioValidationException($"fixture property '{name}' must be a string");

    private static int Int(JsonObject parent, string name) =>
        parent[name]?.GetValue<int>()
            ?? throw new ScenarioValidationException($"fixture property '{name}' must be an integer");
}
