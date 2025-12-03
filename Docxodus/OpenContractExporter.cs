#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

#region Export Result Classes

/// <summary>
/// The complete OpenContracts document export format.
/// Compatible with the OpenContracts ecosystem for document analysis.
/// </summary>
public class OpenContractDocExport
{
    /// <summary>
    /// Document title (from core properties or filename).
    /// </summary>
    public string Title { get; set; } = "";

    /// <summary>
    /// Complete document text content. ALL text from the document must be included.
    /// </summary>
    public string Content { get; set; } = "";

    /// <summary>
    /// Optional document description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Estimated page count of the document.
    /// </summary>
    public int PageCount { get; set; }

    /// <summary>
    /// PAWLS-format page layout information with token positions.
    /// </summary>
    public List<PawlsPage> PawlsFileContent { get; set; } = new();

    /// <summary>
    /// Document-level labels (categories applied to the whole document).
    /// </summary>
    public List<string> DocLabels { get; set; } = new();

    /// <summary>
    /// Annotations/labeled text spans in the document.
    /// </summary>
    public List<OpenContractsAnnotation> LabelledText { get; set; } = new();

    /// <summary>
    /// Relationships between annotations.
    /// </summary>
    public List<OpenContractsRelationship>? Relationships { get; set; }
}

/// <summary>
/// PAWLS page containing page boundary and token information.
/// </summary>
public class PawlsPage
{
    /// <summary>
    /// Page boundary information (dimensions and index).
    /// </summary>
    public PawlsPageBoundary Page { get; set; } = new();

    /// <summary>
    /// Tokens on this page with position information.
    /// </summary>
    public List<PawlsToken> Tokens { get; set; } = new();
}

/// <summary>
/// Page boundary information for PAWLS format.
/// </summary>
public class PawlsPageBoundary
{
    /// <summary>
    /// Page width in points.
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// Page height in points.
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// Zero-based page index.
    /// </summary>
    public int Index { get; set; }
}

/// <summary>
/// Token with position information for PAWLS format.
/// </summary>
public class PawlsToken
{
    /// <summary>
    /// X coordinate (left edge) in points.
    /// </summary>
    public double X { get; set; }

    /// <summary>
    /// Y coordinate (top edge) in points.
    /// </summary>
    public double Y { get; set; }

    /// <summary>
    /// Token width in points.
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// Token height in points.
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// The text content of this token.
    /// </summary>
    public string Text { get; set; } = "";
}

/// <summary>
/// OpenContracts annotation format.
/// </summary>
public class OpenContractsAnnotation
{
    /// <summary>
    /// Unique annotation identifier.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Label/category for this annotation.
    /// </summary>
    public string AnnotationLabel { get; set; } = "";

    /// <summary>
    /// The raw text content of the annotation.
    /// </summary>
    public string RawText { get; set; } = "";

    /// <summary>
    /// Starting page number (0-indexed).
    /// </summary>
    public int Page { get; set; }

    /// <summary>
    /// Position data for the annotation. Can be either:
    /// - A dictionary of page indices to single-page annotation data
    /// - A TextSpan with start/end character offsets
    /// </summary>
    public object? AnnotationJson { get; set; }

    /// <summary>
    /// Parent annotation ID for hierarchical annotations.
    /// </summary>
    public string? ParentId { get; set; }

    /// <summary>
    /// Type of annotation (e.g., "text", "structural").
    /// </summary>
    public string? AnnotationType { get; set; }

    /// <summary>
    /// Whether this is a structural element (section, heading, etc.).
    /// </summary>
    public bool Structural { get; set; }
}

/// <summary>
/// Per-page annotation position data.
/// </summary>
public class OpenContractsSinglePageAnnotation
{
    /// <summary>
    /// Bounding box for the annotation on this page.
    /// </summary>
    public BoundingBox Bounds { get; set; } = new();

    /// <summary>
    /// Token indices that make up this annotation on this page.
    /// </summary>
    public List<TokenId> TokensJsons { get; set; } = new();

    /// <summary>
    /// Raw text content on this page.
    /// </summary>
    public string RawText { get; set; } = "";
}

/// <summary>
/// Bounding box coordinates.
/// </summary>
public class BoundingBox
{
    /// <summary>
    /// Top edge coordinate.
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// Bottom edge coordinate.
    /// </summary>
    public double Bottom { get; set; }

    /// <summary>
    /// Left edge coordinate.
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// Right edge coordinate.
    /// </summary>
    public double Right { get; set; }
}

/// <summary>
/// Token identifier referencing a specific token on a specific page.
/// </summary>
public class TokenId
{
    /// <summary>
    /// Zero-based page index.
    /// </summary>
    public int PageIndex { get; set; }

    /// <summary>
    /// Zero-based token index within the page.
    /// </summary>
    public int TokenIndex { get; set; }
}

/// <summary>
/// Text span with character offsets for annotation positioning.
/// </summary>
public class TextSpan
{
    /// <summary>
    /// Optional span identifier.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Start character offset (0-indexed).
    /// </summary>
    public int Start { get; set; }

    /// <summary>
    /// End character offset (exclusive).
    /// </summary>
    public int End { get; set; }

    /// <summary>
    /// The text content of this span.
    /// </summary>
    public string Text { get; set; } = "";
}

/// <summary>
/// Relationship between annotations.
/// </summary>
public class OpenContractsRelationship
{
    /// <summary>
    /// Unique relationship identifier.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Label describing the relationship type.
    /// </summary>
    public string RelationshipLabel { get; set; } = "";

    /// <summary>
    /// IDs of source annotations.
    /// </summary>
    public List<string> SourceAnnotationIds { get; set; } = new();

    /// <summary>
    /// IDs of target annotations.
    /// </summary>
    public List<string> TargetAnnotationIds { get; set; } = new();

    /// <summary>
    /// Whether this is a structural relationship.
    /// </summary>
    public bool Structural { get; set; }
}

#endregion

#region Internal Structure Classes

/// <summary>
/// Internal representation of an extracted text block with position info.
/// </summary>
internal class ExtractedTextBlock
{
    public string Text { get; set; } = "";
    public int StartOffset { get; set; }
    public int EndOffset { get; set; }
    public string ElementType { get; set; } = "";
    public string? ElementId { get; set; }
    public int? SectionIndex { get; set; }
    public int? PageIndex { get; set; }
    public string? ParentId { get; set; }
}

/// <summary>
/// Internal representation of document section info.
/// </summary>
internal class ExtractedSection
{
    public int Index { get; set; }
    public double PageWidthPt { get; set; }
    public double PageHeightPt { get; set; }
    public double MarginTopPt { get; set; }
    public double MarginBottomPt { get; set; }
    public double MarginLeftPt { get; set; }
    public double MarginRightPt { get; set; }
    public int StartOffset { get; set; }
    public int EndOffset { get; set; }
    public List<ExtractedTextBlock> Blocks { get; set; } = new();
}

#endregion

/// <summary>
/// Exports DOCX documents to the OpenContracts format for interoperability
/// with the OpenContracts ecosystem.
/// </summary>
public static class OpenContractExporter
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    // Default page dimensions (US Letter in points)
    private const double DefaultPageWidthPt = 612.0;  // 8.5 inches
    private const double DefaultPageHeightPt = 792.0; // 11 inches
    private const double DefaultMarginPt = 72.0;      // 1 inch

    // EMU to points conversion (914400 EMUs per inch, 72 points per inch)
    private const double EmuPerPoint = 914400.0 / 72.0;

    /// <summary>
    /// Export a WmlDocument to OpenContracts format.
    /// </summary>
    /// <param name="doc">The document to export.</param>
    /// <returns>The exported document in OpenContracts format.</returns>
    public static OpenContractDocExport Export(WmlDocument doc)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));

        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);

        return Export(wordDoc);
    }

    /// <summary>
    /// Export a WordprocessingDocument to OpenContracts format.
    /// </summary>
    /// <param name="wordDoc">The document to export.</param>
    /// <returns>The exported document in OpenContracts format.</returns>
    public static OpenContractDocExport Export(WordprocessingDocument wordDoc)
    {
        if (wordDoc == null) throw new ArgumentNullException(nameof(wordDoc));

        var result = new OpenContractDocExport();

        // Extract title from document properties
        result.Title = ExtractTitle(wordDoc);
        result.Description = ExtractDescription(wordDoc);

        // Extract all text content
        var contentBuilder = new StringBuilder();
        var textBlocks = new List<ExtractedTextBlock>();
        var sections = new List<ExtractedSection>();

        ExtractContent(wordDoc, contentBuilder, textBlocks, sections);

        result.Content = contentBuilder.ToString();

        // Calculate page count and generate PAWLS content
        var pageInfo = CalculatePages(sections, result.Content.Length);
        result.PageCount = pageInfo.pageCount;
        result.PawlsFileContent = GeneratePawlsContent(sections, textBlocks, pageInfo.pageAssignments);

        // Convert existing annotations if present
        if (wordDoc.MainDocumentPart != null)
        {
            var wmlDoc = new WmlDocument("temp.docx", ReadToByteArray(wordDoc));
            if (AnnotationManager.HasAnnotations(wmlDoc))
            {
                var annotations = AnnotationManager.GetAnnotations(wmlDoc);
                result.LabelledText = ConvertAnnotations(annotations, result.Content, textBlocks);
            }
        }

        // Generate structural annotations for sections, paragraphs, tables
        var structuralAnnotations = GenerateStructuralAnnotations(textBlocks, sections);
        result.LabelledText.AddRange(structuralAnnotations);

        // Generate relationships between structural elements
        result.Relationships = GenerateRelationships(structuralAnnotations);

        return result;
    }

    #region Title and Description Extraction

    private static string ExtractTitle(WordprocessingDocument wordDoc)
    {
        try
        {
            var coreProps = wordDoc.PackageProperties;
            if (!string.IsNullOrEmpty(coreProps?.Title))
            {
                return coreProps.Title;
            }
        }
        catch
        {
            // Ignore errors reading properties
        }

        return "Untitled Document";
    }

    private static string? ExtractDescription(WordprocessingDocument wordDoc)
    {
        try
        {
            var coreProps = wordDoc.PackageProperties;
            if (!string.IsNullOrEmpty(coreProps?.Description))
            {
                return coreProps.Description;
            }
            if (!string.IsNullOrEmpty(coreProps?.Subject))
            {
                return coreProps.Subject;
            }
        }
        catch
        {
            // Ignore errors reading properties
        }

        return null;
    }

    #endregion

    #region Content Extraction

    private static void ExtractContent(
        WordprocessingDocument wordDoc,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        List<ExtractedSection> sections)
    {
        var mainPart = wordDoc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        var body = XElement.Parse(mainPart.Document.Body.OuterXml);

        // Get section properties for page dimensions
        var sectionProps = ExtractSectionProperties(body);

        int currentSection = 0;
        int blockIndex = 0;
        var currentSectionData = new ExtractedSection
        {
            Index = 0,
            PageWidthPt = sectionProps.Count > 0 ? sectionProps[0].PageWidthPt : DefaultPageWidthPt,
            PageHeightPt = sectionProps.Count > 0 ? sectionProps[0].PageHeightPt : DefaultPageHeightPt,
            MarginTopPt = sectionProps.Count > 0 ? sectionProps[0].MarginTopPt : DefaultMarginPt,
            MarginBottomPt = sectionProps.Count > 0 ? sectionProps[0].MarginBottomPt : DefaultMarginPt,
            MarginLeftPt = sectionProps.Count > 0 ? sectionProps[0].MarginLeftPt : DefaultMarginPt,
            MarginRightPt = sectionProps.Count > 0 ? sectionProps[0].MarginRightPt : DefaultMarginPt,
            StartOffset = 0
        };

        // Extract headers and footers first (at document level)
        ExtractHeadersAndFooters(wordDoc, contentBuilder, textBlocks, ref blockIndex);

        // Extract main document body
        foreach (var element in body.Elements())
        {
            // Check for section break
            var sectPr = element.Element(W + "pPr")?.Element(W + "sectPr");
            if (sectPr != null)
            {
                // End current section
                currentSectionData.EndOffset = contentBuilder.Length;
                sections.Add(currentSectionData);

                // Start new section
                currentSection++;
                var props = currentSection < sectionProps.Count ? sectionProps[currentSection] : sectionProps.LastOrDefault();
                currentSectionData = new ExtractedSection
                {
                    Index = currentSection,
                    PageWidthPt = props?.PageWidthPt ?? DefaultPageWidthPt,
                    PageHeightPt = props?.PageHeightPt ?? DefaultPageHeightPt,
                    MarginTopPt = props?.MarginTopPt ?? DefaultMarginPt,
                    MarginBottomPt = props?.MarginBottomPt ?? DefaultMarginPt,
                    MarginLeftPt = props?.MarginLeftPt ?? DefaultMarginPt,
                    MarginRightPt = props?.MarginRightPt ?? DefaultMarginPt,
                    StartOffset = contentBuilder.Length
                };
            }

            if (element.Name == W + "p")
            {
                ExtractParagraph(element, contentBuilder, textBlocks, ref blockIndex, currentSection, $"doc/p-{blockIndex}");
            }
            else if (element.Name == W + "tbl")
            {
                ExtractTable(element, contentBuilder, textBlocks, ref blockIndex, currentSection, $"doc/tbl-{blockIndex}");
            }
            else if (element.Name == W + "sdt")
            {
                // Structured document tag (content control) - extract content
                ExtractSdtContent(element, contentBuilder, textBlocks, ref blockIndex, currentSection);
            }
        }

        // Finalize last section
        currentSectionData.EndOffset = contentBuilder.Length;
        sections.Add(currentSectionData);

        // Extract footnotes and endnotes
        ExtractFootnotes(wordDoc, contentBuilder, textBlocks, ref blockIndex);
        ExtractEndnotes(wordDoc, contentBuilder, textBlocks, ref blockIndex);
    }

    private static List<ExtractedSection> ExtractSectionProperties(XElement body)
    {
        var sections = new List<ExtractedSection>();
        int index = 0;

        // Find section properties within paragraphs (for section breaks)
        foreach (var sectPr in body.Descendants(W + "sectPr"))
        {
            var section = ParseSectionProperties(sectPr, index++);
            sections.Add(section);
        }

        // If no sections found, use defaults
        if (sections.Count == 0)
        {
            sections.Add(new ExtractedSection
            {
                Index = 0,
                PageWidthPt = DefaultPageWidthPt,
                PageHeightPt = DefaultPageHeightPt,
                MarginTopPt = DefaultMarginPt,
                MarginBottomPt = DefaultMarginPt,
                MarginLeftPt = DefaultMarginPt,
                MarginRightPt = DefaultMarginPt
            });
        }

        return sections;
    }

    private static ExtractedSection ParseSectionProperties(XElement sectPr, int index)
    {
        var section = new ExtractedSection { Index = index };

        // Page size
        var pgSz = sectPr.Element(W + "pgSz");
        if (pgSz != null)
        {
            var w = pgSz.Attribute(W + "w")?.Value;
            var h = pgSz.Attribute(W + "h")?.Value;

            if (double.TryParse(w, out var width))
                section.PageWidthPt = width / 20.0; // Twips to points
            else
                section.PageWidthPt = DefaultPageWidthPt;

            if (double.TryParse(h, out var height))
                section.PageHeightPt = height / 20.0;
            else
                section.PageHeightPt = DefaultPageHeightPt;
        }
        else
        {
            section.PageWidthPt = DefaultPageWidthPt;
            section.PageHeightPt = DefaultPageHeightPt;
        }

        // Margins
        var pgMar = sectPr.Element(W + "pgMar");
        if (pgMar != null)
        {
            var top = pgMar.Attribute(W + "top")?.Value;
            var bottom = pgMar.Attribute(W + "bottom")?.Value;
            var left = pgMar.Attribute(W + "left")?.Value;
            var right = pgMar.Attribute(W + "right")?.Value;

            section.MarginTopPt = double.TryParse(top, out var t) ? t / 20.0 : DefaultMarginPt;
            section.MarginBottomPt = double.TryParse(bottom, out var b) ? b / 20.0 : DefaultMarginPt;
            section.MarginLeftPt = double.TryParse(left, out var l) ? l / 20.0 : DefaultMarginPt;
            section.MarginRightPt = double.TryParse(right, out var r) ? r / 20.0 : DefaultMarginPt;
        }
        else
        {
            section.MarginTopPt = DefaultMarginPt;
            section.MarginBottomPt = DefaultMarginPt;
            section.MarginLeftPt = DefaultMarginPt;
            section.MarginRightPt = DefaultMarginPt;
        }

        return section;
    }

    private static void ExtractParagraph(
        XElement para,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        ref int blockIndex,
        int sectionIndex,
        string elementId)
    {
        var startOffset = contentBuilder.Length;
        var paraText = new StringBuilder();

        foreach (var element in para.Elements())
        {
            if (element.Name == W + "r")
            {
                ExtractRun(element, paraText);
            }
            else if (element.Name == W + "hyperlink")
            {
                foreach (var run in element.Elements(W + "r"))
                {
                    ExtractRun(run, paraText);
                }
            }
            else if (element.Name == W + "bookmarkStart" || element.Name == W + "bookmarkEnd")
            {
                // Skip bookmarks
            }
            else if (element.Name == W + "sdt")
            {
                // Content control within paragraph
                var sdtContent = element.Element(W + "sdtContent");
                if (sdtContent != null)
                {
                    foreach (var run in sdtContent.Descendants(W + "r"))
                    {
                        ExtractRun(run, paraText);
                    }
                }
            }
        }

        var text = paraText.ToString();
        if (text.Length > 0 || true) // Include empty paragraphs as line breaks
        {
            contentBuilder.Append(text);
            contentBuilder.Append('\n'); // Paragraph separator

            textBlocks.Add(new ExtractedTextBlock
            {
                Text = text,
                StartOffset = startOffset,
                EndOffset = contentBuilder.Length - 1, // Exclude the newline
                ElementType = "Paragraph",
                ElementId = elementId,
                SectionIndex = sectionIndex
            });
        }

        blockIndex++;
    }

    private static void ExtractRun(XElement run, StringBuilder textBuilder)
    {
        foreach (var element in run.Elements())
        {
            if (element.Name == W + "t")
            {
                textBuilder.Append(element.Value);
            }
            else if (element.Name == W + "tab")
            {
                textBuilder.Append('\t');
            }
            else if (element.Name == W + "br")
            {
                var type = element.Attribute(W + "type")?.Value;
                if (type == "page")
                {
                    textBuilder.Append('\f'); // Form feed for page break
                }
                else
                {
                    textBuilder.Append('\n'); // Line break
                }
            }
            else if (element.Name == W + "sym")
            {
                // Symbol - try to get the character
                var charCode = element.Attribute(W + "char")?.Value;
                if (!string.IsNullOrEmpty(charCode) && int.TryParse(charCode, System.Globalization.NumberStyles.HexNumber, null, out var code))
                {
                    textBuilder.Append((char)code);
                }
            }
            else if (element.Name == W + "drawing")
            {
                // Image placeholder
                textBuilder.Append("[IMAGE]");
            }
            else if (element.Name == W + "object")
            {
                // OLE object placeholder
                textBuilder.Append("[OBJECT]");
            }
            else if (element.Name == W + "pict")
            {
                // Legacy picture
                textBuilder.Append("[PICTURE]");
            }
        }
    }

    private static void ExtractTable(
        XElement table,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        ref int blockIndex,
        int sectionIndex,
        string elementId)
    {
        var startOffset = contentBuilder.Length;
        int rowIndex = 0;

        foreach (var row in table.Elements(W + "tr"))
        {
            int cellIndex = 0;
            foreach (var cell in row.Elements(W + "tc"))
            {
                var cellId = $"{elementId}/tr-{rowIndex}/tc-{cellIndex}";

                // Extract cell content (can contain paragraphs and nested tables)
                foreach (var element in cell.Elements())
                {
                    if (element.Name == W + "p")
                    {
                        ExtractParagraph(element, contentBuilder, textBlocks, ref blockIndex, sectionIndex, $"{cellId}/p-{blockIndex}");
                    }
                    else if (element.Name == W + "tbl")
                    {
                        // Nested table
                        ExtractTable(element, contentBuilder, textBlocks, ref blockIndex, sectionIndex, $"{cellId}/tbl-{blockIndex}");
                    }
                }

                cellIndex++;
            }
            rowIndex++;
        }

        // Add table block
        textBlocks.Add(new ExtractedTextBlock
        {
            Text = contentBuilder.ToString(startOffset, contentBuilder.Length - startOffset),
            StartOffset = startOffset,
            EndOffset = contentBuilder.Length,
            ElementType = "Table",
            ElementId = elementId,
            SectionIndex = sectionIndex
        });
    }

    private static void ExtractSdtContent(
        XElement sdt,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        ref int blockIndex,
        int sectionIndex)
    {
        var sdtContent = sdt.Element(W + "sdtContent");
        if (sdtContent == null) return;

        foreach (var element in sdtContent.Elements())
        {
            if (element.Name == W + "p")
            {
                ExtractParagraph(element, contentBuilder, textBlocks, ref blockIndex, sectionIndex, $"doc/sdt/p-{blockIndex}");
            }
            else if (element.Name == W + "tbl")
            {
                ExtractTable(element, contentBuilder, textBlocks, ref blockIndex, sectionIndex, $"doc/sdt/tbl-{blockIndex}");
            }
        }
    }

    private static void ExtractHeadersAndFooters(
        WordprocessingDocument wordDoc,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        ref int blockIndex)
    {
        var mainPart = wordDoc.MainDocumentPart;
        if (mainPart == null) return;

        // Extract from all header parts
        foreach (var headerPart in mainPart.HeaderParts)
        {
            try
            {
                var headerXml = headerPart.Header?.OuterXml;
                if (string.IsNullOrEmpty(headerXml)) continue;

                var header = XElement.Parse(headerXml);
                var startOffset = contentBuilder.Length;

                foreach (var para in header.Descendants(W + "p"))
                {
                    ExtractParagraph(para, contentBuilder, textBlocks, ref blockIndex, 0, $"header/p-{blockIndex}");
                }

                if (contentBuilder.Length > startOffset)
                {
                    textBlocks.Add(new ExtractedTextBlock
                    {
                        Text = contentBuilder.ToString(startOffset, contentBuilder.Length - startOffset),
                        StartOffset = startOffset,
                        EndOffset = contentBuilder.Length,
                        ElementType = "Header"
                    });
                }
            }
            catch
            {
                // Skip invalid header parts
            }
        }

        // Extract from all footer parts
        foreach (var footerPart in mainPart.FooterParts)
        {
            try
            {
                var footerXml = footerPart.Footer?.OuterXml;
                if (string.IsNullOrEmpty(footerXml)) continue;

                var footer = XElement.Parse(footerXml);
                var startOffset = contentBuilder.Length;

                foreach (var para in footer.Descendants(W + "p"))
                {
                    ExtractParagraph(para, contentBuilder, textBlocks, ref blockIndex, 0, $"footer/p-{blockIndex}");
                }

                if (contentBuilder.Length > startOffset)
                {
                    textBlocks.Add(new ExtractedTextBlock
                    {
                        Text = contentBuilder.ToString(startOffset, contentBuilder.Length - startOffset),
                        StartOffset = startOffset,
                        EndOffset = contentBuilder.Length,
                        ElementType = "Footer"
                    });
                }
            }
            catch
            {
                // Skip invalid footer parts
            }
        }
    }

    private static void ExtractFootnotes(
        WordprocessingDocument wordDoc,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        ref int blockIndex)
    {
        var mainPart = wordDoc.MainDocumentPart;
        var footnotesPart = mainPart?.FootnotesPart;
        if (footnotesPart?.Footnotes == null) return;

        try
        {
            var footnotesXml = footnotesPart.Footnotes.OuterXml;
            var footnotes = XElement.Parse(footnotesXml);

            foreach (var footnote in footnotes.Elements(W + "footnote"))
            {
                var id = footnote.Attribute(W + "id")?.Value;
                // Skip special footnotes (separator, continuation separator)
                if (id == "0" || id == "-1") continue;

                var startOffset = contentBuilder.Length;

                foreach (var para in footnote.Elements(W + "p"))
                {
                    ExtractParagraph(para, contentBuilder, textBlocks, ref blockIndex, 0, $"footnote-{id}/p-{blockIndex}");
                }

                if (contentBuilder.Length > startOffset)
                {
                    textBlocks.Add(new ExtractedTextBlock
                    {
                        Text = contentBuilder.ToString(startOffset, contentBuilder.Length - startOffset),
                        StartOffset = startOffset,
                        EndOffset = contentBuilder.Length,
                        ElementType = "Footnote",
                        ElementId = $"footnote-{id}"
                    });
                }
            }
        }
        catch
        {
            // Skip if footnotes can't be parsed
        }
    }

    private static void ExtractEndnotes(
        WordprocessingDocument wordDoc,
        StringBuilder contentBuilder,
        List<ExtractedTextBlock> textBlocks,
        ref int blockIndex)
    {
        var mainPart = wordDoc.MainDocumentPart;
        var endnotesPart = mainPart?.EndnotesPart;
        if (endnotesPart?.Endnotes == null) return;

        try
        {
            var endnotesXml = endnotesPart.Endnotes.OuterXml;
            var endnotes = XElement.Parse(endnotesXml);

            foreach (var endnote in endnotes.Elements(W + "endnote"))
            {
                var id = endnote.Attribute(W + "id")?.Value;
                // Skip special endnotes
                if (id == "0" || id == "-1") continue;

                var startOffset = contentBuilder.Length;

                foreach (var para in endnote.Elements(W + "p"))
                {
                    ExtractParagraph(para, contentBuilder, textBlocks, ref blockIndex, 0, $"endnote-{id}/p-{blockIndex}");
                }

                if (contentBuilder.Length > startOffset)
                {
                    textBlocks.Add(new ExtractedTextBlock
                    {
                        Text = contentBuilder.ToString(startOffset, contentBuilder.Length - startOffset),
                        StartOffset = startOffset,
                        EndOffset = contentBuilder.Length,
                        ElementType = "Endnote",
                        ElementId = $"endnote-{id}"
                    });
                }
            }
        }
        catch
        {
            // Skip if endnotes can't be parsed
        }
    }

    #endregion

    #region Page Calculation and PAWLS Generation

    private static (int pageCount, Dictionary<int, int> pageAssignments) CalculatePages(
        List<ExtractedSection> sections,
        int totalContentLength)
    {
        if (totalContentLength == 0)
        {
            return (1, new Dictionary<int, int>());
        }

        // Estimate pages based on content length and section properties
        // Average characters per page estimation (varies by font size, margins)
        const int avgCharsPerPage = 2500;

        var pageAssignments = new Dictionary<int, int>();
        int totalPages = 0;
        int currentOffset = 0;

        foreach (var section in sections)
        {
            var sectionLength = section.EndOffset - section.StartOffset;
            if (sectionLength <= 0) continue;

            // Calculate content area
            var contentWidth = section.PageWidthPt - section.MarginLeftPt - section.MarginRightPt;
            var contentHeight = section.PageHeightPt - section.MarginTopPt - section.MarginBottomPt;

            // Adjust chars per page based on content area ratio to default
            var areaRatio = (contentWidth * contentHeight) / ((DefaultPageWidthPt - 2 * DefaultMarginPt) * (DefaultPageHeightPt - 2 * DefaultMarginPt));
            var adjustedCharsPerPage = (int)(avgCharsPerPage * areaRatio);
            adjustedCharsPerPage = Math.Max(adjustedCharsPerPage, 500); // Minimum

            var sectionPages = Math.Max(1, (int)Math.Ceiling((double)sectionLength / adjustedCharsPerPage));

            for (int i = 0; i < sectionLength; i++)
            {
                var pageWithinSection = Math.Min(i / adjustedCharsPerPage, sectionPages - 1);
                pageAssignments[currentOffset + i] = totalPages + pageWithinSection;
            }

            totalPages += sectionPages;
            currentOffset += sectionLength;
        }

        return (Math.Max(1, totalPages), pageAssignments);
    }

    private static List<PawlsPage> GeneratePawlsContent(
        List<ExtractedSection> sections,
        List<ExtractedTextBlock> textBlocks,
        Dictionary<int, int> pageAssignments)
    {
        var pages = new Dictionary<int, PawlsPage>();

        // Group text blocks by estimated page
        foreach (var block in textBlocks)
        {
            if (string.IsNullOrEmpty(block.Text)) continue;

            var pageIndex = pageAssignments.TryGetValue(block.StartOffset, out var p) ? p : 0;
            var section = sections.FirstOrDefault(s => block.SectionIndex == s.Index) ?? sections.FirstOrDefault();

            if (!pages.ContainsKey(pageIndex))
            {
                pages[pageIndex] = new PawlsPage
                {
                    Page = new PawlsPageBoundary
                    {
                        Index = pageIndex,
                        Width = section?.PageWidthPt ?? DefaultPageWidthPt,
                        Height = section?.PageHeightPt ?? DefaultPageHeightPt
                    }
                };
            }

            // Generate tokens for each word in the block
            var page = pages[pageIndex];
            var words = block.Text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            double x = section?.MarginLeftPt ?? DefaultMarginPt;
            double y = section?.MarginTopPt ?? DefaultMarginPt;
            double lineHeight = 12.0; // Estimated line height in points
            double charWidth = 6.0;   // Estimated character width in points

            foreach (var word in words)
            {
                var tokenWidth = word.Length * charWidth;

                // Check if we need to wrap to next line
                var maxX = (section?.PageWidthPt ?? DefaultPageWidthPt) - (section?.MarginRightPt ?? DefaultMarginPt);
                if (x + tokenWidth > maxX)
                {
                    x = section?.MarginLeftPt ?? DefaultMarginPt;
                    y += lineHeight;
                }

                page.Tokens.Add(new PawlsToken
                {
                    X = x,
                    Y = y,
                    Width = tokenWidth,
                    Height = lineHeight,
                    Text = word
                });

                x += tokenWidth + charWidth; // Add space after word
            }
        }

        // Ensure we have at least one page
        if (pages.Count == 0)
        {
            pages[0] = new PawlsPage
            {
                Page = new PawlsPageBoundary
                {
                    Index = 0,
                    Width = DefaultPageWidthPt,
                    Height = DefaultPageHeightPt
                }
            };
        }

        return pages.OrderBy(kvp => kvp.Key).Select(kvp => kvp.Value).ToList();
    }

    #endregion

    #region Annotation Conversion

    private static List<OpenContractsAnnotation> ConvertAnnotations(
        List<DocumentAnnotation> annotations,
        string fullContent,
        List<ExtractedTextBlock> textBlocks)
    {
        var result = new List<OpenContractsAnnotation>();

        foreach (var ann in annotations)
        {
            var converted = new OpenContractsAnnotation
            {
                Id = ann.Id,
                AnnotationLabel = ann.LabelId ?? ann.Label ?? "",
                RawText = ann.AnnotatedText ?? "",
                Page = ann.StartPage ?? 0,
                ParentId = null,
                AnnotationType = "text",
                Structural = false
            };

            // Find the text span in the content
            if (!string.IsNullOrEmpty(ann.AnnotatedText))
            {
                var startIndex = fullContent.IndexOf(ann.AnnotatedText, StringComparison.Ordinal);
                if (startIndex >= 0)
                {
                    converted.AnnotationJson = new TextSpan
                    {
                        Start = startIndex,
                        End = startIndex + ann.AnnotatedText.Length,
                        Text = ann.AnnotatedText
                    };
                }
            }

            result.Add(converted);
        }

        return result;
    }

    private static List<OpenContractsAnnotation> GenerateStructuralAnnotations(
        List<ExtractedTextBlock> textBlocks,
        List<ExtractedSection> sections)
    {
        var result = new List<OpenContractsAnnotation>();
        int annotationId = 1;

        // Add section annotations
        foreach (var section in sections)
        {
            result.Add(new OpenContractsAnnotation
            {
                Id = $"section-{section.Index}",
                AnnotationLabel = "SECTION",
                RawText = "",
                Page = 0, // Will be calculated
                AnnotationJson = new TextSpan
                {
                    Start = section.StartOffset,
                    End = section.EndOffset,
                    Text = ""
                },
                AnnotationType = "structural",
                Structural = true
            });
        }

        // Add paragraph annotations
        foreach (var block in textBlocks.Where(b => b.ElementType == "Paragraph"))
        {
            result.Add(new OpenContractsAnnotation
            {
                Id = $"para-{annotationId++}",
                AnnotationLabel = "PARAGRAPH",
                RawText = block.Text,
                Page = 0,
                AnnotationJson = new TextSpan
                {
                    Start = block.StartOffset,
                    End = block.EndOffset,
                    Text = block.Text
                },
                ParentId = block.SectionIndex.HasValue ? $"section-{block.SectionIndex}" : null,
                AnnotationType = "structural",
                Structural = true
            });
        }

        // Add table annotations
        foreach (var block in textBlocks.Where(b => b.ElementType == "Table"))
        {
            result.Add(new OpenContractsAnnotation
            {
                Id = $"table-{annotationId++}",
                AnnotationLabel = "TABLE",
                RawText = block.Text,
                Page = 0,
                AnnotationJson = new TextSpan
                {
                    Start = block.StartOffset,
                    End = block.EndOffset,
                    Text = block.Text
                },
                ParentId = block.SectionIndex.HasValue ? $"section-{block.SectionIndex}" : null,
                AnnotationType = "structural",
                Structural = true
            });
        }

        return result;
    }

    private static List<OpenContractsRelationship> GenerateRelationships(
        List<OpenContractsAnnotation> annotations)
    {
        var result = new List<OpenContractsRelationship>();
        int relationshipId = 1;

        // Create parent-child relationships
        foreach (var ann in annotations.Where(a => !string.IsNullOrEmpty(a.ParentId)))
        {
            result.Add(new OpenContractsRelationship
            {
                Id = $"rel-{relationshipId++}",
                RelationshipLabel = "CONTAINS",
                SourceAnnotationIds = new List<string> { ann.ParentId! },
                TargetAnnotationIds = new List<string> { ann.Id! },
                Structural = true
            });
        }

        return result;
    }

    #endregion

    #region Utility Methods

    private static byte[] ReadToByteArray(WordprocessingDocument wordDoc)
    {
        // Save to memory stream and return bytes
        using var ms = new MemoryStream();
        wordDoc.Clone(ms);
        return ms.ToArray();
    }

    #endregion
}
