#nullable enable

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Builders for the small programmatic DOCX fixtures the <see cref="Docxodus.Ir.IrReader"/>
/// tests exercise. Each fixture includes the parts CLAUDE.md flags as required for a
/// well-formed package built from scratch (StyleDefinitionsPart, DocumentSettingsPart).
/// </summary>
internal static class IrTestDocuments
{
    internal const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// A document whose body holds one simple text paragraph per supplied string.
    /// </summary>
    internal static WmlDocument Create(params string[] paragraphTexts)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            foreach (var text in paragraphTexts)
                body.Append(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));

            main.Document.Save();
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose <c>w:body</c> inner XML is exactly <paramref name="bodyInnerXml"/> — the
    /// raw OOXML between <c>&lt;w:body&gt;</c> and <c>&lt;/w:body&gt;</c>. Lets a test express any
    /// body shape (tables, breaks, opaque elements, sectPr, revisions) directly.
    /// </summary>
    internal static WmlDocument FromBodyXml(string bodyInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
            using (var partStream = main.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(documentXml);
            }
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// Like <see cref="FromBodyXml"/>, but the <c>w:document</c> root also declares the DrawingML / VML /
    /// markup-compatibility namespaces (<c>a</c>, <c>wp</c>, <c>wps</c>, <c>v</c>, <c>mc</c>, <c>r</c>) so
    /// a test can express a realistic <c>w:drawing</c>/<c>wps:txbx</c> or <c>w:pict</c>/<c>v:textbox</c>
    /// textbox shape (including an <c>mc:AlternateContent</c> Choice/Fallback pair) in the body fragment.
    /// </summary>
    internal static WmlDocument FromBodyXmlWithDrawingNamespaces(string bodyInnerXml)
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string WpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const string Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        const string V = "urn:schemas-microsoft-com:vml";
        const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var documentXml =
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:a=\"{A}\" xmlns:wp=\"{WpNs}\" " +
                $"xmlns:wps=\"{Wps}\" xmlns:v=\"{V}\" xmlns:mc=\"{Mc}\">" +
                $"<w:body>{bodyInnerXml}</w:body></w:document>";
            WritePartXml(main, documentXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// Like <see cref="FromBodyXml"/>, but also wires up external hyperlink relationships on the
    /// main document part. Each entry in <paramref name="hyperlinkRels"/> maps a relationship id
    /// (the <c>r:id</c> a <c>w:hyperlink</c> references) to its external target URI, added as an
    /// external relationship so <c>part.HyperlinkRelationships</c> resolves it.
    /// </summary>
    internal static WmlDocument FromBodyXmlWithHyperlinks(
        string bodyInnerXml,
        params (string RelId, string Uri)[] hyperlinkRels)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
            using (var partStream = main.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(documentXml);
            }

            foreach (var (relId, uri) in hyperlinkRels)
                main.AddHyperlinkRelationship(new System.Uri(uri, System.UriKind.RelativeOrAbsolute), true, relId);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose <c>w:body</c> inner XML is <paramref name="bodyInnerXml"/> and whose
    /// <c>w:styles</c> inner XML (the content between <c>&lt;w:styles&gt;</c> and
    /// <c>&lt;/w:styles&gt;</c>) is <paramref name="stylesInnerXml"/>. Lets a test wire up a style
    /// chain (e.g. a style carrying <c>w:numPr</c>, optionally via <c>w:basedOn</c>) that a body
    /// paragraph references by <c>w:pStyle</c>.
    /// </summary>
    internal static WmlDocument FromBodyAndStylesXml(string bodyInnerXml, string stylesInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var stylesXml = $"<w:styles xmlns:w=\"{W}\">{stylesInnerXml}</w:styles>";
            using (var stylesStream = stylesPart.GetStream(FileMode.Create, FileAccess.Write))
            using (var stylesWriter = new StreamWriter(stylesStream))
            {
                stylesWriter.Write(stylesXml);
            }

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
            using (var partStream = main.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(documentXml);
            }
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose <c>w:body</c> inner XML is <paramref name="bodyInnerXml"/> and that also wires
    /// up a single <see cref="HeaderPart"/> whose <c>w:hdr</c> inner XML is
    /// <paramref name="headerInnerXml"/>, referenced from the body's trailing <c>w:sectPr</c> via a
    /// default <c>w:headerReference</c>. Lets a test place revision markup in a header part only.
    /// </summary>
    internal static WmlDocument FromBodyAndHeaderXml(string bodyInnerXml, string headerInnerXml)
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            const string headerRelId = "rIdHdr1";
            var headerPart = main.AddNewPart<HeaderPart>(headerRelId);
            WritePartXml(headerPart, $"<w:hdr xmlns:w=\"{W}\">{headerInnerXml}</w:hdr>");

            var documentXml =
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\"><w:body>{bodyInnerXml}" +
                $"<w:sectPr><w:headerReference w:type=\"default\" r:id=\"{headerRelId}\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";
            WritePartXml(main, documentXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose <c>w:body</c>, <c>w:styles</c>, and <c>w:numbering</c> inner XML are each
    /// supplied directly, plus an optional theme. A null <paramref name="numberingInnerXml"/> omits
    /// the <see cref="NumberingDefinitionsPart"/> entirely; a null
    /// <paramref name="themeFontSchemeInnerXml"/> omits the <see cref="ThemePart"/>. When a theme is
    /// supplied, <paramref name="themeFontSchemeInnerXml"/> is the inner XML of <c>a:fontScheme</c>
    /// (e.g. <c>&lt;a:majorFont&gt;&lt;a:latin typeface="Calibri Light"/&gt;&lt;/a:majorFont&gt;…</c>).
    /// </summary>
    internal static WmlDocument FromParts(
        string bodyInnerXml,
        string stylesInnerXml = "",
        string? numberingInnerXml = null,
        string? themeFontSchemeInnerXml = null)
    {
        const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            WritePartXml(stylesPart, $"<w:styles xmlns:w=\"{W}\">{stylesInnerXml}</w:styles>");

            if (numberingInnerXml is not null)
            {
                var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
                WritePartXml(numberingPart, $"<w:numbering xmlns:w=\"{W}\">{numberingInnerXml}</w:numbering>");
            }

            if (themeFontSchemeInnerXml is not null)
            {
                var themePart = main.AddNewPart<ThemePart>();
                WritePartXml(themePart,
                    $"<a:theme xmlns:a=\"{A}\" name=\"t\"><a:themeElements>" +
                    $"<a:fontScheme name=\"fs\">{themeFontSchemeInnerXml}</a:fontScheme>" +
                    "</a:themeElements></a:theme>");
            }

            WritePartXml(main, $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>");
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }

    /// <summary>
    /// A document whose <c>w:body</c> inner XML is <paramref name="bodyInnerXml"/>, with one or more
    /// <see cref="ImagePart"/>s added to the main document part. Each entry in
    /// <paramref name="imageParts"/> maps a relationship id (the <c>r:embed</c> an <c>a:blip</c>
    /// references) to the raw image bytes stored in that part. The body XML must reference the same
    /// rel ids via <c>a:blip r:embed</c>. Namespaces for <c>r:</c> / <c>a:</c> / <c>wp:</c> are not
    /// declared automatically — callers should include them on the <c>w:document</c> root if needed;
    /// this builder declares <c>w:</c> only, so tests pass a fully namespaced body fragment.
    /// </summary>
    internal static WmlDocument FromBodyXmlWithImageParts(
        string bodyInnerXml,
        params (string RelId, byte[] Bytes)[] imageParts)
    {
        const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string WpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            foreach (var (relId, bytes) in imageParts)
            {
                var imagePart = main.AddNewPart<ImagePart>("image/png", relId);
                using var imgStream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
                imgStream.Write(bytes, 0, bytes.Length);
            }

            var documentXml =
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:a=\"{A}\" xmlns:wp=\"{WpNs}\">" +
                $"<w:body>{bodyInnerXml}</w:body></w:document>";
            using (var partStream = main.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(documentXml);
            }
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// Minimal valid PNG bytes (an 8-byte signature plus a stub) — enough for an
    /// <see cref="ImagePart"/> to store and for the reader to hash. Not a renderable image; the IR
    /// only ever hashes the bytes, never decodes them.
    /// </summary>
    internal static byte[] TinyPng { get; } =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR length + type
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
    };

    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// A document with one header part and one footer part, each holding a single text paragraph, and
    /// a body section that references both (default type). Used for header/footer scope parity tests.
    /// </summary>
    internal static WmlDocument WithHeaderAndFooter(
        string headerText, string footerText, string bodyText = "Body paragraph")
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var headerPart = main.AddNewPart<HeaderPart>();
            var headerId = main.GetIdOfPart(headerPart);
            WritePartXml(headerPart,
                $"<w:hdr xmlns:w=\"{W}\"><w:p><w:r><w:t xml:space=\"preserve\">{headerText}</w:t></w:r></w:p></w:hdr>");

            var footerPart = main.AddNewPart<FooterPart>();
            var footerId = main.GetIdOfPart(footerPart);
            WritePartXml(footerPart,
                $"<w:ftr xmlns:w=\"{W}\"><w:p><w:r><w:t xml:space=\"preserve\">{footerText}</w:t></w:r></w:p></w:ftr>");

            var bodyXml =
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{RNs}\"><w:body>" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{bodyText}</w:t></w:r></w:p>" +
                "<w:sectPr>" +
                $"<w:headerReference w:type=\"default\" r:id=\"{headerId}\"/>" +
                $"<w:footerReference w:type=\"default\" r:id=\"{footerId}\"/>" +
                "</w:sectPr></w:body></w:document>";
            WritePartXml(main, bodyXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// Like <see cref="FromBodyXml"/>, but also wires up a <see cref="FootnotesPart"/> whose real note
    /// (id=1) holds <paramref name="footnoteText"/> in a single paragraph (plus the two Word-reserved
    /// boilerplate notes), and appends a footnote-reference run to the body so the note is referenced. Used
    /// by the diff fuzzer's footnote-edit mutation: the body XML is supplied verbatim, the footnote ref is
    /// appended in its own trailing paragraph so it does not perturb the body paragraphs the fuzzer mutates.
    /// </summary>
    internal static WmlDocument FromBodyXmlWithFootnote(string bodyInnerXml, string footnoteText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fnPart = main.AddNewPart<FootnotesPart>();
            WritePartXml(fnPart,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                $"<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">{Escape(footnoteText)}</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}" +
                "<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "</w:body></w:document>";
            WritePartXml(main, documentXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    private static string Escape(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    /// <summary>
    /// Like <see cref="FromBodyXmlWithFootnote"/> but the real footnote (id=1) paragraph's INNER XML is
    /// supplied directly (e.g. <c>&lt;w:pPr&gt;&lt;w:jc w:val="center"/&gt;&lt;/w:pPr&gt;&lt;w:r&gt;…</c>),
    /// so a test can vary the footnote paragraph's <c>w:pPr</c> across two documents.
    /// </summary>
    internal static WmlDocument FromBodyXmlWithFootnoteParagraph(string bodyInnerXml, string footnotePInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fnPart = main.AddNewPart<FootnotesPart>();
            WritePartXml(fnPart,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                $"<w:footnote w:id=\"1\"><w:p>{footnotePInnerXml}</w:p></w:footnote>" +
                "</w:footnotes>");

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}" +
                "<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "</w:body></w:document>";
            WritePartXml(main, documentXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document with a footnotes part and an endnotes part. Each part carries the two Word-reserved
    /// boilerplate notes (separator id=-1, continuationSeparator id=0) plus one real note (id=1) whose
    /// single paragraph holds the supplied text. The body references the footnote/endnote via runs.
    /// </summary>
    internal static WmlDocument WithFootnoteAndEndnote(string footnoteText, string endnoteText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var fnPart = main.AddNewPart<FootnotesPart>();
            WritePartXml(fnPart,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                $"<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">{footnoteText}</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            var enPart = main.AddNewPart<EndnotesPart>();
            WritePartXml(enPart,
                $"<w:endnotes xmlns:w=\"{W}\">" +
                "<w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>" +
                "<w:endnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>" +
                $"<w:endnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">{endnoteText}</w:t></w:r></w:p></w:endnote>" +
                "</w:endnotes>");

            var bodyXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>" +
                "<w:p><w:r><w:t>Body.</w:t></w:r>" +
                "<w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r>" +
                "<w:r><w:endnoteReference w:id=\"1\"/></w:r></w:p>" +
                "</w:body></w:document>";
            WritePartXml(main, bodyXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document with a comments part holding one comment (author/initials/date) and a body whose
    /// supplied inner XML wires up the <c>w:commentRangeStart</c>/<c>End</c>/<c>w:commentReference</c>
    /// plumbing. The body inner XML is inserted verbatim, so a test can place comment ranges exactly.
    /// </summary>
    internal static WmlDocument WithComment(
        string author, string initials, string date, string commentText, string bodyInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            WritePartXml(commentsPart,
                $"<w:comments xmlns:w=\"{W}\">" +
                $"<w:comment w:id=\"0\" w:author=\"{author}\" w:initials=\"{initials}\" w:date=\"{date}\">" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{commentText}</w:t></w:r></w:p>" +
                "</w:comment></w:comments>");

            var bodyXml = $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
            WritePartXml(main, bodyXml);
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }
}
