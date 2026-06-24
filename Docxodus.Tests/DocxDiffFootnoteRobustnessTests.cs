#nullable enable
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Footnote id↔reference↔text robustness beyond the scenario corpus: cases where a note's <c>w:id</c> does NOT
/// equal its body-reference ordinal, so the renumber pass cannot rely on a positional coincidence to keep a
/// reference-deleted-but-definition-preserved note resolvable. The synthetic scenario base uses ids {1,2} in
/// reference order, which masks this; the real NVCA contract (111 footnotes, gaps) does not.
/// </summary>
public class DocxDiffFootnoteRobustnessTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A doc whose footnotes carry NON-sequential ids (1 and 5), referenced in body order [1, 5].</summary>
    private static byte[] BuildDocWithGappedFootnoteIds()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">First footnote text.</w:t></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"5\"><w:p><w:r><w:t xml:space=\"preserve\">Fifth footnote text with a gapped id.</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Alpha paragraph references one.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Beta paragraph references five.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"5\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>");
        }
        return ms.ToArray();
    }

    /// <summary>Right = left with the footnote-5-referencing paragraph wholly rewritten (its w:footnoteReference
    /// dropped). The fn5 DEFINITION lingers in the part (unreferenced) — the reconcile case.</summary>
    private static byte[] RewriteFiveReferencingParagraph(byte[] left)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var p = body.Elements<Paragraph>().First(x =>
                string.Concat(x.Descendants<Text>().Select(t => t.Text)).Contains("Beta paragraph"));
            foreach (var r in p.Elements<Run>().ToList()) r.Remove();
            p.AppendChild(new Run(new Text("Beta paragraph has been wholly rewritten.")
                { Space = SpaceProcessingModeValues.Preserve }));
        }
        return ms.ToArray();
    }

    /// <summary>A doc with the THREE reserved boilerplate footnotes Word can emit — separator (-1),
    /// continuationSeparator (0), AND continuationNotice with a POSITIVE id (1, as the real NVCA contract uses) —
    /// plus real footnotes referenced in the body.</summary>
    private static byte[] BuildDocWithPositiveIdReservedFootnote()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                // continuationNotice with POSITIVE id 1 — the reserved note the NVCA contract carries.
                "<w:footnote w:type=\"continuationNotice\" w:id=\"1\"><w:p/></w:footnote>" +
                "<w:footnote w:id=\"2\"><w:p><w:r><w:t xml:space=\"preserve\">First real footnote.</w:t></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"3\"><w:p><w:r><w:t xml:space=\"preserve\">Second real footnote.</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Alpha references the first.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Beta references the second.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"3\"/></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>");
        }
        return ms.ToArray();
    }

    private static byte[] ReplaceWord(byte[] left, string find, string repl)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var t = doc.MainDocumentPart!.Document!.Body!.Descendants<Text>().First(x => x.Text.Contains(find));
            t.Text = t.Text.Replace(find, repl);
            t.Space = SpaceProcessingModeValues.Preserve;
        }
        return ms.ToArray();
    }

    [Fact]
    public void ReservedContinuationNoticeAtPositiveId_DoesNotCollideWithRenumberedRealFootnote()
    {
        // Editing a body paragraph (NOT a footnote) on a doc whose continuationNotice boilerplate occupies id 1
        // must not make a renumbered real footnote ALSO land on id 1. Reproduces the NVCA-scale duplicate that
        // every edit (even format/body-only) tripped; the synthetic corpus (reserved ids only -1/0) masks it.
        var left = new WmlDocument("left.docx", BuildDocWithPositiveIdReservedFootnote());
        var right = new WmlDocument("right.docx", ReplaceWord(left.DocumentByteArray, "Alpha", "Gamma"));

        var result = DocxDiff.Compare(left, right);

        var ids = FootnoteIds(result);
        Assert.Equal(ids.Distinct().Count(), ids.Count);

        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0, $"schema errors: {string.Join(" | ", newErrors.Take(5))}");

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(ReferencedFootnoteTexts(right), ReferencedFootnoteTexts(accepted));
        Assert.Equal(ReferencedFootnoteTexts(left), ReferencedFootnoteTexts(rejected));
    }

    [Fact]
    public void DeletingReferenceToGappedIdFootnote_KeepsStructureResolvable()
    {
        var left = new WmlDocument("left.docx", BuildDocWithGappedFootnoteIds());
        var right = new WmlDocument("right.docx", RewriteFiveReferencingParagraph(left.DocumentByteArray));

        var result = DocxDiff.Compare(left, right);

        // Unique footnote definition ids.
        var ids = FootnoteIds(result);
        Assert.Equal(ids.Distinct().Count(), ids.Count);

        // No NEW schema errors vs the left.
        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0, $"schema errors: {string.Join(" | ", newErrors.Take(5))}");

        // id → ref → text round-trip: accept ≡ right (only fn1 referenced), reject ≡ left (fn1 + fn5).
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(ReferencedFootnoteTexts(right), ReferencedFootnoteTexts(accepted));
        Assert.Equal(ReferencedFootnoteTexts(left), ReferencedFootnoteTexts(rejected));
    }

    // ---- helpers (mirrors DocxDiffScenarioTests, scoped to footnotes) --------------------------

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    private static List<long> FootnoteIds(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var footnotes = w.MainDocumentPart?.FootnotesPart?.Footnotes;
        return footnotes is null
            ? new List<long>()
            : footnotes.Elements<Footnote>().Where(f => f.Id is not null).Select(f => f.Id!.Value).ToList();
    }

    private static List<string> ReferencedFootnoteTexts(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var defText = new Dictionary<long, string>();
        foreach (var f in main.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>())
            if (f.Id is not null)
                defText[f.Id.Value] = string.Concat(f.Descendants<Text>().Select(t => t.Text));
        var seq = new List<string>();
        foreach (var r in main.Document?.Body?.Descendants<FootnoteReference>() ?? Enumerable.Empty<FootnoteReference>())
        {
            var id = r.Id?.Value;
            seq.Add(id is not null && defText.TryGetValue(id.Value, out var t) ? t : "(unresolved)");
        }
        return seq;
    }

    private static HashSet<string> SchemaErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        return validator.Validate(w).Select(e => $"{e.Id}@{e.Path?.XPath}: {e.Description}").ToHashSet();
    }
}
