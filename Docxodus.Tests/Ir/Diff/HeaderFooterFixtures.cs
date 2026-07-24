#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using WordType = DocumentFormat.OpenXml.WordprocessingDocumentType;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Programmatic header/footer-bearing document builders + story-text projections for the
/// header/footer scope-diff campaign (spec: 2026-07-03-docxdiff-header-footer-diff-design.md).
/// A document is described as N sections, each with body paragraph texts and explicit
/// header/footer story references (kind → part id); parts are declared once and may be
/// referenced by several sections (the inheritance/multi-reference shapes).
/// </summary>
internal static class HeaderFooterFixtures
{
    public const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public const string Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace Wn = Wns;

    /// <summary>One section: its body paragraph texts and its explicit story references
    /// (kind "default"/"first"/"even" → part id declared in the Build call).</summary>
    public sealed record Section(
        string[] Paras,
        (string Kind, string PartId)[]? Headers = null,
        (string Kind, string PartId)[]? Footers = null);

    /// <summary>
    /// Build a document with <paramref name="sections"/> in order. <paramref name="headerParts"/> /
    /// <paramref name="footerParts"/> map a part id (also used as the relationship id) to the part's
    /// paragraph texts. Parts are created in dictionary insertion order — pass ordered dictionaries
    /// when part-enumeration order matters. A part never referenced by any section is still created.
    /// </summary>
    public static WmlDocument Build(
        IReadOnlyList<Section> sections,
        IReadOnlyDictionary<string, string[]>? headerParts = null,
        IReadOnlyDictionary<string, string[]>? footerParts = null,
        bool titlePg = false, bool evenAndOddHeaders = false)
    {
        headerParts ??= new Dictionary<string, string[]>();
        footerParts ??= new Dictionary<string, string[]>();
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            var settings = main.AddNewPart<DocumentSettingsPart>();
            WritePart(settings,
                $"<w:settings xmlns:w=\"{Wns}\">" +
                (evenAndOddHeaders ? "<w:evenAndOddHeaders/>" : "") +
                "</w:settings>");

            foreach (var (id, paras) in headerParts)
                WritePart(main.AddNewPart<HeaderPart>(id),
                    $"<w:hdr xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\">{Paras(paras)}</w:hdr>");
            foreach (var (id, paras) in footerParts)
                WritePart(main.AddNewPart<FooterPart>(id),
                    $"<w:ftr xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\">{Paras(paras)}</w:ftr>");

            var body = new StringBuilder();
            for (int s = 0; s < sections.Count; s++)
            {
                var sect = sections[s];
                bool last = s == sections.Count - 1;
                // Non-last sections carry their sectPr inside the LAST paragraph's pPr; the last
                // section's sectPr is the trailing body-level element (the OOXML section model).
                for (int p = 0; p < sect.Paras.Length; p++)
                {
                    bool carrySectPr = !last && p == sect.Paras.Length - 1;
                    body.Append("<w:p>");
                    if (carrySectPr)
                        body.Append("<w:pPr>").Append(SectPr(sect, titlePg)).Append("</w:pPr>");
                    body.Append($"<w:r><w:t xml:space=\"preserve\">{sect.Paras[p]}</w:t></w:r></w:p>");
                }
                if (last)
                    body.Append(SectPr(sect, titlePg));
            }

            WritePart(main,
                $"<w:document xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\"><w:body>{body}</w:body></w:document>");
        }
        return new WmlDocument("hf.docx", ms.ToArray());
    }

    private static string SectPr(Section sect, bool titlePg)
    {
        var sb = new StringBuilder("<w:sectPr>");
        foreach (var (kind, partId) in sect.Headers ?? Array.Empty<(string, string)>())
            sb.Append($"<w:headerReference w:type=\"{kind}\" r:id=\"{partId}\"/>");
        foreach (var (kind, partId) in sect.Footers ?? Array.Empty<(string, string)>())
            sb.Append($"<w:footerReference w:type=\"{kind}\" r:id=\"{partId}\"/>");
        sb.Append("<w:pgSz w:w=\"12240\" w:h=\"15840\"/>");
        if (titlePg) sb.Append("<w:titlePg/>");
        sb.Append("</w:sectPr>");
        return sb.ToString();
    }

    /// <summary>Each entry is a paragraph text — or, when it starts with <c>&lt;w:</c>, raw block XML
    /// (a <c>w:tbl</c>, a pre-built <c>w:p</c>) emitted verbatim.</summary>
    private static string Paras(string[] texts) =>
        string.Concat(texts.Select(t => t.StartsWith("<w:", StringComparison.Ordinal)
            ? t
            : $"<w:p><w:r><w:t xml:space=\"preserve\">{t}</w:t></w:r></w:p>"));

    private static void WritePart(OpenXmlPart part, string xml)
    {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        using var w = new StreamWriter(s);
        w.Write(xml);
    }

    /// <summary>Convenience: a single-section doc with one default header and/or footer.</summary>
    public static WmlDocument Simple(string[] bodyParas, string[]? headerParas = null, string[]? footerParas = null)
    {
        var headers = new List<(string, string)>();
        var footers = new List<(string, string)>();
        var hp = new Dictionary<string, string[]>();
        var fp = new Dictionary<string, string[]>();
        if (headerParas != null) { hp["rIdH1"] = headerParas; headers.Add(("default", "rIdH1")); }
        if (footerParas != null) { fp["rIdF1"] = footerParas; footers.Add(("default", "rIdF1")); }
        return Build(new[] { new Section(bodyParas, headers.ToArray(), footers.ToArray()) }, hp, fp);
    }

    /// <summary>Per-part visible text (<c>w:t</c> concatenated), header parts then footer parts, in
    /// part-enumeration order — the story-scope analogue of <c>Docs.PlainText</c>.</summary>
    public static List<string> StoryTexts(WmlDocument d)
    {
        var result = new List<string>();
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        foreach (var h in main.HeaderParts) result.Add(PartText(h));
        foreach (var f in main.FooterParts) result.Add(PartText(f));
        return result;
    }

    /// <summary>Visible text of the effective story a given section+kind resolves to, or null when no
    /// reference exists in that section or any earlier section. <paramref name="isHeader"/> selects header vs
    /// footer references; <paramref name="sectionIndex"/> is the document-order sectPr ordinal.</summary>
    public static string? ReferencedStoryText(WmlDocument d, bool isHeader, int sectionIndex, string kind)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        XDocument mainXd;
        using (var s = main.GetStream(FileMode.Open, FileAccess.Read)) mainXd = XDocument.Load(s);
        var sectPrs = mainXd.Descendants(Wn + "sectPr").ToList();
        if (sectionIndex >= sectPrs.Count) return null;
        XNamespace r = Rns;
        var refName = Wn + (isHeader ? "headerReference" : "footerReference");
        var reference = sectPrs.Take(sectionIndex + 1).Reverse()
            .SelectMany(sectPr => sectPr.Elements(refName))
            .FirstOrDefault(e => ((string?)e.Attribute(Wn + "type") ?? "default") == kind);
        if (reference is null) return null;
        var relId = (string?)reference.Attribute(r + "id");
        if (relId is null) return null;
        var part = main.GetPartById(relId);
        return PartText(part);
    }

    private static string PartText(OpenXmlPart part)
    {
        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
        var xd = XDocument.Load(s);
        return string.Concat(xd.Descendants(Wn + "t").Select(t => t.Value));
    }

    /// <summary>A minimal valid 1×1 transparent PNG (for image-in-story fixtures).</summary>
    public static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    /// <summary>A <c>w:p</c> carrying an inline picture whose blip references <paramref name="relId"/>.</summary>
    public static string ImageParagraphXml(string relId) =>
        "<w:p><w:r><w:drawing>" +
        "<wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
        "<wp:extent cx=\"95250\" cy=\"95250\"/><wp:docPr id=\"1\" name=\"Logo\"/>" +
        "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Logo\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        $"<pic:blipFill><a:blip r:embed=\"{relId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"95250\" cy=\"95250\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    /// <summary>Add a 1×1 PNG <see cref="ImagePart"/> under the FIRST header part with relationship id
    /// <paramref name="relId"/>, so a story block built with <see cref="ImageParagraphXml"/> resolves.</summary>
    public static WmlDocument WithImageInFirstHeaderPart(WmlDocument d, string relId)
        => WithImageInHeaderPart(d, headerIndex: 0, relId: relId);

    /// <summary>Add a 1×1 PNG under one header part in part-enumeration order. Useful for a shared-story
    /// reassignment regression where the primary and cloned stories each own different image relationships.</summary>
    public static WmlDocument WithImageInHeaderPart(WmlDocument d, int headerIndex, string relId)
    {
        using var ms = new MemoryStream();
        ms.Write(d.DocumentByteArray, 0, d.DocumentByteArray.Length);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var headerPart = doc.MainDocumentPart!.HeaderParts.ElementAt(headerIndex);
            var imagePart = headerPart.AddNewPart<ImagePart>("image/png", relId);
            using var s = imagePart.GetStream(FileMode.Create, FileAccess.Write);
            s.Write(OnePixelPng, 0, OnePixelPng.Length);
        }
        return new WmlDocument(d.FileName, ms.ToArray());
    }

    /// <summary>Raw XML of every header part then footer part, in part-enumeration order.</summary>
    public static List<string> StoryPartsXml(WmlDocument d)
    {
        var result = new List<string>();
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        foreach (var p in main.HeaderParts.Cast<OpenXmlPart>().Concat(main.FooterParts))
        {
            using var s = p.GetStream(FileMode.Open, FileAccess.Read);
            result.Add(XDocument.Load(s).ToString(SaveOptions.DisableFormatting));
        }
        return result;
    }
}
