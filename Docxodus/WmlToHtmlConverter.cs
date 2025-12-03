// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using SkiaSharp;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

// 200e lrm - LTR
// 200f rlm - RTL

// todo need to set the HTTP "Content-Language" header, for instance:
// Content-Language: en-US
// Content-Language: fr-FR

namespace Docxodus
{
    /// <summary>
    /// Specifies how comments are rendered in the HTML output.
    /// </summary>
    public enum CommentRenderMode
    {
        /// <summary>
        /// Comments are rendered as a section at the end of the document
        /// with bidirectional links to/from commented text (default).
        /// </summary>
        EndnoteStyle,

        /// <summary>
        /// Comments are rendered inline as data attributes on the highlighted text.
        /// Uses title attribute for tooltip display.
        /// </summary>
        Inline,

        /// <summary>
        /// Comments are rendered in a margin area using CSS positioning.
        /// Best for print-style layouts.
        /// </summary>
        Margin
    }

    /// <summary>
    /// Specifies how pagination is rendered in the HTML output.
    /// </summary>
    public enum PaginationMode
    {
        /// <summary>
        /// No pagination - content flows continuously (default, current behavior).
        /// </summary>
        None,

        /// <summary>
        /// Paginated view - outputs page containers with document dimensions
        /// and content with data attributes for client-side pagination.
        /// Creates a PDF.js-style page preview experience.
        /// </summary>
        Paginated
    }

    public partial class WmlDocument
    {
        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertToHtml(WmlToHtmlConverterSettings htmlConverterSettings)
        {
            return WmlToHtmlConverter.ConvertToHtml(this, htmlConverterSettings);
        }

        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertToHtml(HtmlConverterSettings htmlConverterSettings)
        {
            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(this, settings);
        }
    }

    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Global")]
    public class WmlToHtmlConverterSettings
    {
        public string PageTitle;
        public string CssClassPrefix;
        public bool FabricateCssClasses;
        public string GeneralCss;
        public string AdditionalCss;
        public bool RestrictToSupportedLanguages;
        public bool RestrictToSupportedNumberingFormats;
        public Dictionary<string, Func<string, int, string, string>> ListItemImplementations;
        public Func<ImageInfo, XElement> ImageHandler;

        /// <summary>
        /// If true, render tracked changes visually in HTML output.
        /// If false (default), accept all revisions before conversion.
        /// </summary>
        public bool RenderTrackedChanges;

        /// <summary>
        /// CSS class prefix for revision elements (default: "rev-")
        /// </summary>
        public string RevisionCssClassPrefix;

        /// <summary>
        /// If true, include revision metadata (author, date) as data attributes
        /// </summary>
        public bool IncludeRevisionMetadata;

        /// <summary>
        /// If true, show deleted content with strikethrough (default: true)
        /// If false, hide deleted content entirely
        /// </summary>
        public bool ShowDeletedContent;

        /// <summary>
        /// Custom colors for different authors (author name -> CSS color)
        /// </summary>
        public Dictionary<string, string> AuthorColors;

        /// <summary>
        /// If true, render move operations as separate from/to (default: true)
        /// If false, render moves as regular delete + insert
        /// </summary>
        public bool RenderMoveOperations;

        /// <summary>
        /// If true, render footnotes and endnotes at the end of the HTML document.
        /// If false (default), footnotes and endnotes are stripped from the output.
        /// </summary>
        public bool RenderFootnotesAndEndnotes;

        /// <summary>
        /// If true, render headers and footers in the HTML document.
        /// If false (default), headers and footers are not rendered.
        /// </summary>
        public bool RenderHeadersAndFooters;

        /// <summary>
        /// If true, render comments in HTML output.
        /// If false (default), comments are stripped from the output.
        /// </summary>
        public bool RenderComments;

        /// <summary>
        /// How to render comments in the HTML output (default: EndnoteStyle).
        /// </summary>
        public CommentRenderMode CommentRenderMode;

        /// <summary>
        /// CSS class prefix for comment elements (default: "comment-")
        /// </summary>
        public string CommentCssClassPrefix;

        /// <summary>
        /// If true, include comment metadata (author, date) as data attributes
        /// </summary>
        public bool IncludeCommentMetadata;

        /// <summary>
        /// If not None, render document with page containers in PDF.js style.
        /// Default: None (continuous scrolling layout).
        /// </summary>
        public PaginationMode RenderPagination;

        /// <summary>
        /// Scale factor for page rendering in paginated mode (1.0 = 100%).
        /// Default: 1.0
        /// </summary>
        public double PaginationScale;

        /// <summary>
        /// CSS class prefix for pagination elements.
        /// Default: "page-"
        /// </summary>
        public string PaginationCssClassPrefix;

        /// <summary>
        /// If true, render custom annotations as highlights in HTML output.
        /// Default: false
        /// </summary>
        public bool RenderAnnotations;

        /// <summary>
        /// CSS class prefix for annotation elements.
        /// Default: "annot-"
        /// </summary>
        public string AnnotationCssClassPrefix;

        /// <summary>
        /// How to display annotation labels.
        /// Default: Above
        /// </summary>
        public AnnotationLabelMode AnnotationLabelMode;

        /// <summary>
        /// If true, include annotation metadata as data attributes.
        /// Default: true
        /// </summary>
        public bool IncludeAnnotationMetadata;

        public WmlToHtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
            RestrictToSupportedLanguages = false;
            RestrictToSupportedNumberingFormats = false;
            ListItemImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations;
            RenderTrackedChanges = false;
            RevisionCssClassPrefix = "rev-";
            IncludeRevisionMetadata = true;
            ShowDeletedContent = true;
            RenderMoveOperations = true;
            RenderFootnotesAndEndnotes = false;
            RenderHeadersAndFooters = false;
            RenderComments = false;
            CommentRenderMode = CommentRenderMode.EndnoteStyle;
            CommentCssClassPrefix = "comment-";
            IncludeCommentMetadata = true;
            RenderPagination = PaginationMode.None;
            PaginationScale = 1.0;
            PaginationCssClassPrefix = "page-";
            RenderAnnotations = false;
            AnnotationCssClassPrefix = "annot-";
            AnnotationLabelMode = AnnotationLabelMode.Above;
            IncludeAnnotationMetadata = true;
        }

        public WmlToHtmlConverterSettings(HtmlConverterSettings htmlConverterSettings)
        {
            PageTitle = htmlConverterSettings.PageTitle;
            CssClassPrefix = htmlConverterSettings.CssClassPrefix;
            FabricateCssClasses = htmlConverterSettings.FabricateCssClasses;
            GeneralCss = htmlConverterSettings.GeneralCss;
            AdditionalCss = htmlConverterSettings.AdditionalCss;
            RestrictToSupportedLanguages = htmlConverterSettings.RestrictToSupportedLanguages;
            RestrictToSupportedNumberingFormats = htmlConverterSettings.RestrictToSupportedNumberingFormats;
            ListItemImplementations = htmlConverterSettings.ListItemImplementations;
            ImageHandler = htmlConverterSettings.ImageHandler;
            RenderTrackedChanges = htmlConverterSettings.RenderTrackedChanges;
            RevisionCssClassPrefix = htmlConverterSettings.RevisionCssClassPrefix;
            IncludeRevisionMetadata = htmlConverterSettings.IncludeRevisionMetadata;
            ShowDeletedContent = htmlConverterSettings.ShowDeletedContent;
            AuthorColors = htmlConverterSettings.AuthorColors;
            RenderMoveOperations = htmlConverterSettings.RenderMoveOperations;
            RenderFootnotesAndEndnotes = htmlConverterSettings.RenderFootnotesAndEndnotes;
            RenderHeadersAndFooters = htmlConverterSettings.RenderHeadersAndFooters;
            RenderComments = htmlConverterSettings.RenderComments;
            CommentRenderMode = htmlConverterSettings.CommentRenderMode;
            CommentCssClassPrefix = htmlConverterSettings.CommentCssClassPrefix;
            IncludeCommentMetadata = htmlConverterSettings.IncludeCommentMetadata;
            RenderPagination = htmlConverterSettings.RenderPagination;
            PaginationScale = htmlConverterSettings.PaginationScale;
            PaginationCssClassPrefix = htmlConverterSettings.PaginationCssClassPrefix;
            RenderAnnotations = htmlConverterSettings.RenderAnnotations;
            AnnotationCssClassPrefix = htmlConverterSettings.AnnotationCssClassPrefix;
            AnnotationLabelMode = htmlConverterSettings.AnnotationLabelMode;
            IncludeAnnotationMetadata = htmlConverterSettings.IncludeAnnotationMetadata;
        }
    }

    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Global")]
    public class HtmlConverterSettings
    {
        public string PageTitle;
        public string CssClassPrefix;
        public bool FabricateCssClasses;
        public string GeneralCss;
        public string AdditionalCss;
        public bool RestrictToSupportedLanguages;
        public bool RestrictToSupportedNumberingFormats;
        public Dictionary<string, Func<string, int, string, string>> ListItemImplementations;
        public Func<ImageInfo, XElement> ImageHandler;

        /// <summary>
        /// If true, render tracked changes visually in HTML output.
        /// If false (default), accept all revisions before conversion.
        /// </summary>
        public bool RenderTrackedChanges;

        /// <summary>
        /// CSS class prefix for revision elements (default: "rev-")
        /// </summary>
        public string RevisionCssClassPrefix;

        /// <summary>
        /// If true, include revision metadata (author, date) as data attributes
        /// </summary>
        public bool IncludeRevisionMetadata;

        /// <summary>
        /// If true, show deleted content with strikethrough (default: true)
        /// If false, hide deleted content entirely
        /// </summary>
        public bool ShowDeletedContent;

        /// <summary>
        /// Custom colors for different authors (author name -> CSS color)
        /// </summary>
        public Dictionary<string, string> AuthorColors;

        /// <summary>
        /// If true, render move operations as separate from/to (default: true)
        /// If false, render moves as regular delete + insert
        /// </summary>
        public bool RenderMoveOperations;

        /// <summary>
        /// If true, render footnotes and endnotes at the end of the HTML document.
        /// If false (default), footnotes and endnotes are stripped from the output.
        /// </summary>
        public bool RenderFootnotesAndEndnotes;

        /// <summary>
        /// If true, render headers and footers in the HTML document.
        /// If false (default), headers and footers are not rendered.
        /// </summary>
        public bool RenderHeadersAndFooters;

        /// <summary>
        /// If true, render comments in HTML output.
        /// If false (default), comments are stripped from the output.
        /// </summary>
        public bool RenderComments;

        /// <summary>
        /// How to render comments in the HTML output (default: EndnoteStyle).
        /// </summary>
        public CommentRenderMode CommentRenderMode;

        /// <summary>
        /// CSS class prefix for comment elements (default: "comment-")
        /// </summary>
        public string CommentCssClassPrefix;

        /// <summary>
        /// If true, include comment metadata (author, date) as data attributes
        /// </summary>
        public bool IncludeCommentMetadata;

        /// <summary>
        /// If not None, render document with page containers in PDF.js style.
        /// Default: None (continuous scrolling layout).
        /// </summary>
        public PaginationMode RenderPagination;

        /// <summary>
        /// Scale factor for page rendering in paginated mode (1.0 = 100%).
        /// Default: 1.0
        /// </summary>
        public double PaginationScale;

        /// <summary>
        /// CSS class prefix for pagination elements.
        /// Default: "page-"
        /// </summary>
        public string PaginationCssClassPrefix;

        /// <summary>
        /// If true, render custom annotations as highlights in HTML output.
        /// Default: false
        /// </summary>
        public bool RenderAnnotations;

        /// <summary>
        /// CSS class prefix for annotation elements.
        /// Default: "annot-"
        /// </summary>
        public string AnnotationCssClassPrefix;

        /// <summary>
        /// How to display annotation labels.
        /// Default: Above
        /// </summary>
        public AnnotationLabelMode AnnotationLabelMode;

        /// <summary>
        /// If true, include annotation metadata as data attributes.
        /// Default: true
        /// </summary>
        public bool IncludeAnnotationMetadata;

        public HtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
            RestrictToSupportedLanguages = false;
            RestrictToSupportedNumberingFormats = false;
            ListItemImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations;
            RenderTrackedChanges = false;
            RevisionCssClassPrefix = "rev-";
            IncludeRevisionMetadata = true;
            ShowDeletedContent = true;
            RenderMoveOperations = true;
            RenderFootnotesAndEndnotes = false;
            RenderHeadersAndFooters = false;
            RenderComments = false;
            CommentRenderMode = CommentRenderMode.EndnoteStyle;
            CommentCssClassPrefix = "comment-";
            IncludeCommentMetadata = true;
            RenderPagination = PaginationMode.None;
            PaginationScale = 1.0;
            PaginationCssClassPrefix = "page-";
            RenderAnnotations = false;
            AnnotationCssClassPrefix = "annot-";
            AnnotationLabelMode = AnnotationLabelMode.Above;
            IncludeAnnotationMetadata = true;
        }
    }

    public static class HtmlConverter
    {
        public static XElement ConvertToHtml(WmlDocument wmlDoc, HtmlConverterSettings htmlConverterSettings)
        {
            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);
        }

        public static XElement ConvertToHtml(WordprocessingDocument wDoc, HtmlConverterSettings htmlConverterSettings)
        {
            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        }
    }

    [SuppressMessage("ReSharper", "NotAccessedField.Global")]
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public class ImageInfo
    {
        public SKBitmap? Bitmap;
        public byte[]? ImageBytes;
        public XAttribute? ImgStyleAttribute;
        public string? ContentType;
        public XElement? DrawingElement;
        public string? AltText;

        public const int EmusPerInch = 914400;
        public const int EmusPerCm = 360000;

        /// <summary>
        /// Saves the image to a file using the specified format.
        /// </summary>
        public void SaveImage(string filePath, SKEncodedImageFormat format, int quality = 100)
        {
            if (ImageBytes != null)
            {
                // If we have raw bytes, save them directly for formats that match
                System.IO.File.WriteAllBytes(filePath, ImageBytes);
                return;
            }
            if (Bitmap != null)
            {
                using var image = SKImage.FromBitmap(Bitmap);
                using var data = image.Encode(format, quality);
                using var stream = System.IO.File.OpenWrite(filePath);
                data.SaveTo(stream);
            }
        }
    }

    /// <summary>
    /// Information about a single comment in the document.
    /// </summary>
    public class CommentInfo
    {
        public int Id { get; set; }
        public string Author { get; set; }
        public string Date { get; set; }
        public string Initials { get; set; }
        public List<XElement> ContentParagraphs { get; set; } = new List<XElement>();
    }

    /// <summary>
    /// Tracks comment state during HTML transformation.
    /// </summary>
    internal class CommentTracker
    {
        /// <summary>
        /// All comments loaded from the document, keyed by ID.
        /// </summary>
        public Dictionary<int, CommentInfo> Comments { get; } = new Dictionary<int, CommentInfo>();

        /// <summary>
        /// IDs of comment ranges that are currently open (between start and end markers).
        /// </summary>
        public HashSet<int> OpenRanges { get; } = new HashSet<int>();

        /// <summary>
        /// IDs of comments that have been referenced in the document (for rendering order).
        /// </summary>
        public List<int> ReferencedCommentIds { get; } = new List<int>();
    }

    /// <summary>
    /// Tracks annotation state during HTML transformation.
    /// </summary>
    internal class AnnotationTracker
    {
        /// <summary>
        /// All annotations loaded from the document, keyed by ID.
        /// </summary>
        public Dictionary<string, DocumentAnnotation> Annotations { get; } = new Dictionary<string, DocumentAnnotation>();

        /// <summary>
        /// Maps bookmark names to annotation IDs.
        /// </summary>
        public Dictionary<string, string> BookmarkToAnnotationId { get; } = new Dictionary<string, string>();

        /// <summary>
        /// IDs of annotation ranges that are currently open (between bookmarkStart and bookmarkEnd).
        /// </summary>
        public HashSet<string> OpenRanges { get; } = new HashSet<string>();

        /// <summary>
        /// Tracks whether this is the first segment (paragraph) of an annotation for label positioning.
        /// Key: annotation ID, Value: true if first segment has been rendered.
        /// </summary>
        public HashSet<string> FirstSegmentRendered { get; } = new HashSet<string>();
    }

    /// <summary>
    /// Metadata for a single section in the document (for lazy loading).
    /// All dimension values are in points (1/72 inch).
    /// </summary>
    public class SectionMetadata
    {
        /// <summary>Section index (0-based)</summary>
        public int SectionIndex { get; set; }

        /// <summary>Page width in points</summary>
        public double PageWidthPt { get; set; }

        /// <summary>Page height in points</summary>
        public double PageHeightPt { get; set; }

        /// <summary>Top margin in points</summary>
        public double MarginTopPt { get; set; }

        /// <summary>Right margin in points</summary>
        public double MarginRightPt { get; set; }

        /// <summary>Bottom margin in points</summary>
        public double MarginBottomPt { get; set; }

        /// <summary>Left margin in points</summary>
        public double MarginLeftPt { get; set; }

        /// <summary>Content width (page minus margins) in points</summary>
        public double ContentWidthPt { get; set; }

        /// <summary>Content height (page minus margins) in points</summary>
        public double ContentHeightPt { get; set; }

        /// <summary>Header distance from top in points</summary>
        public double HeaderPt { get; set; }

        /// <summary>Footer distance from bottom in points</summary>
        public double FooterPt { get; set; }

        /// <summary>Number of paragraphs in this section</summary>
        public int ParagraphCount { get; set; }

        /// <summary>Number of tables in this section</summary>
        public int TableCount { get; set; }

        /// <summary>Whether this section has a default header</summary>
        public bool HasHeader { get; set; }

        /// <summary>Whether this section has a default footer</summary>
        public bool HasFooter { get; set; }

        /// <summary>Whether this section has a first page header (titlePg enabled)</summary>
        public bool HasFirstPageHeader { get; set; }

        /// <summary>Whether this section has a first page footer (titlePg enabled)</summary>
        public bool HasFirstPageFooter { get; set; }

        /// <summary>Whether this section has an even page header</summary>
        public bool HasEvenPageHeader { get; set; }

        /// <summary>Whether this section has an even page footer</summary>
        public bool HasEvenPageFooter { get; set; }

        /// <summary>Start paragraph index (0-based, global across document)</summary>
        public int StartParagraphIndex { get; set; }

        /// <summary>End paragraph index (exclusive, global across document)</summary>
        public int EndParagraphIndex { get; set; }

        /// <summary>Start table index (0-based, global across document)</summary>
        public int StartTableIndex { get; set; }

        /// <summary>End table index (exclusive, global across document)</summary>
        public int EndTableIndex { get; set; }
    }

    /// <summary>
    /// Document metadata for lazy loading pagination.
    /// Provides fast access to document structure without full HTML rendering.
    /// </summary>
    public class DocumentMetadata
    {
        /// <summary>List of sections with their metadata</summary>
        public List<SectionMetadata> Sections { get; set; } = new List<SectionMetadata>();

        /// <summary>Total number of paragraphs in the document</summary>
        public int TotalParagraphs { get; set; }

        /// <summary>Total number of tables in the document</summary>
        public int TotalTables { get; set; }

        /// <summary>Whether the document has any footnotes</summary>
        public bool HasFootnotes { get; set; }

        /// <summary>Whether the document has any endnotes</summary>
        public bool HasEndnotes { get; set; }

        /// <summary>Whether the document has tracked changes</summary>
        public bool HasTrackedChanges { get; set; }

        /// <summary>Whether the document has comments</summary>
        public bool HasComments { get; set; }

        /// <summary>Estimated total page count (rough estimate based on content)</summary>
        public int EstimatedPageCount { get; set; }
    }

    public static class WmlToHtmlConverter
    {
        public static XElement ConvertToHtml(WmlDocument doc, WmlToHtmlConverterSettings htmlConverterSettings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    return ConvertToHtml(document, htmlConverterSettings);
                }
            }
        }

        public static XElement ConvertToHtml(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings htmlConverterSettings)
        {
            // Only accept revisions if NOT rendering tracked changes AND document has tracked changes
            // This optimization saves ~9% of conversion time for documents without revisions
            if (!htmlConverterSettings.RenderTrackedChanges && HasTrackedChanges(wordDoc))
            {
                RevisionAccepter.AcceptRevisions(wordDoc);
            }

            SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
            {
                RemoveComments = !htmlConverterSettings.RenderComments,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = !htmlConverterSettings.RenderFootnotesAndEndnotes,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                RemoveGoBackBookmark = true,
                ReplaceTabsWithSpaces = false,
            };
            MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

            FormattingAssemblerSettings formattingAssemblerSettings = new FormattingAssemblerSettings
            {
                RemoveStyleNamesFromParagraphAndRunProperties = false,
                ClearStyles = false,
                RestrictToSupportedLanguages = htmlConverterSettings.RestrictToSupportedLanguages,
                RestrictToSupportedNumberingFormats = htmlConverterSettings.RestrictToSupportedNumberingFormats,
                CreateHtmlConverterAnnotationAttributes = true,
                OrderElementsPerStandard = false,
                ListItemRetrieverSettings =
                    htmlConverterSettings.ListItemImplementations == null ?
                    new ListItemRetrieverSettings()
                    {
                        ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                    } :
                    new ListItemRetrieverSettings()
                    {
                        ListItemTextImplementations = htmlConverterSettings.ListItemImplementations,
                    },
            };

            FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);

            InsertAppropriateNonbreakingSpaces(wordDoc);
            CalculateSpanWidthForTabs(wordDoc);
            ReverseTableBordersForRtlTables(wordDoc);
            AdjustTableBorders(wordDoc);
            XElement rootElement = wordDoc.MainDocumentPart.GetXDocument().Root;
            FieldRetriever.AnnotateWithFieldInfo(wordDoc.MainDocumentPart);
            AnnotateForSections(wordDoc);

            // Load comments if rendering is enabled and store as annotation
            var commentTracker = new CommentTracker();
            if (htmlConverterSettings.RenderComments)
            {
                LoadComments(wordDoc, commentTracker);
            }
            // Store tracker as annotation on the root element for access during transform
            rootElement.AddAnnotation(commentTracker);

            // Load custom annotations if rendering is enabled
            var annotationTracker = new AnnotationTracker();
            if (htmlConverterSettings.RenderAnnotations)
            {
                LoadAnnotations(wordDoc, annotationTracker);
            }
            rootElement.AddAnnotation(annotationTracker);

            XElement xhtml = (XElement)ConvertToHtmlTransform(wordDoc, htmlConverterSettings,
                rootElement, false, 0m);

            ReifyStylesAndClasses(htmlConverterSettings, xhtml);

            // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
            // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities will not be serialized properly.

            return xhtml;
        }

        /// <summary>
        /// Quickly checks if a document contains any tracked changes (revisions).
        /// This is used to skip the expensive AcceptRevisions call when not needed.
        /// </summary>
        /// <param name="wordDoc">The Word document to check</param>
        /// <returns>True if the document has any tracked changes</returns>
        private static bool HasTrackedChanges(WordprocessingDocument wordDoc)
        {
            var mainPart = wordDoc.MainDocumentPart;
            if (mainPart == null) return false;

            var xDoc = mainPart.GetXDocument();
            var body = xDoc.Root?.Element(W.body);
            if (body == null) return false;

            // Check for any revision elements - this is a fast descendant scan
            return body.Descendants().Any(e =>
                e.Name == W.ins ||
                e.Name == W.del ||
                e.Name == W.moveFrom ||
                e.Name == W.moveTo ||
                e.Name == W.rPrChange ||
                e.Name == W.pPrChange ||
                e.Name == W.tblPrChange ||
                e.Name == W.tcPrChange ||
                e.Name == W.sectPrChange);
        }

        /// <summary>
        /// Extracts document metadata for lazy loading pagination.
        /// This is a fast operation that doesn't perform full HTML conversion.
        /// </summary>
        /// <param name="doc">The Word document to analyze</param>
        /// <returns>Document metadata including section dimensions and content counts</returns>
        public static DocumentMetadata GetDocumentMetadata(WmlDocument doc)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
                {
                    return GetDocumentMetadata(document);
                }
            }
        }

        /// <summary>
        /// Extracts document metadata for lazy loading pagination.
        /// This is a fast operation that doesn't perform full HTML conversion.
        /// </summary>
        /// <param name="wordDoc">The Word document to analyze</param>
        /// <returns>Document metadata including section dimensions and content counts</returns>
        public static DocumentMetadata GetDocumentMetadata(WordprocessingDocument wordDoc)
        {
            var metadata = new DocumentMetadata();

            var mainDocPart = wordDoc.MainDocumentPart;
            if (mainDocPart == null)
                return metadata;

            var xDoc = mainDocPart.GetXDocument();
            var body = xDoc.Root?.Element(W.body);
            if (body == null)
                return metadata;

            // Check for footnotes
            var footnotesPart = mainDocPart.FootnotesPart;
            metadata.HasFootnotes = footnotesPart != null &&
                footnotesPart.GetXDocument().Root?.Elements(W.footnote)
                    .Any(fn => (string)fn.Attribute(W.type) != "separator" &&
                              (string)fn.Attribute(W.type) != "continuationSeparator") == true;

            // Check for endnotes
            var endnotesPart = mainDocPart.EndnotesPart;
            metadata.HasEndnotes = endnotesPart != null &&
                endnotesPart.GetXDocument().Root?.Elements(W.endnote)
                    .Any(en => (string)en.Attribute(W.type) != "separator" &&
                              (string)en.Attribute(W.type) != "continuationSeparator") == true;

            // Check for comments
            var commentsPart = mainDocPart.WordprocessingCommentsPart;
            metadata.HasComments = commentsPart != null &&
                commentsPart.GetXDocument().Root?.Elements(W.comment).Any() == true;

            // Check for tracked changes (insertions, deletions, moves)
            metadata.HasTrackedChanges = body.Descendants()
                .Any(e => e.Name == W.ins || e.Name == W.del ||
                         e.Name == W.moveFrom || e.Name == W.moveTo);

            // Collect all section properties and their associated content
            var sectionData = CollectSectionData(body);

            int globalParagraphIndex = 0;
            int globalTableIndex = 0;
            int sectionIndex = 0;

            foreach (var (sectPr, paragraphs, tables) in sectionData)
            {
                var sectionMeta = new SectionMetadata
                {
                    SectionIndex = sectionIndex,
                    ParagraphCount = paragraphs.Count,
                    TableCount = tables.Count,
                    StartParagraphIndex = globalParagraphIndex,
                    EndParagraphIndex = globalParagraphIndex + paragraphs.Count,
                    StartTableIndex = globalTableIndex,
                    EndTableIndex = globalTableIndex + tables.Count
                };

                // Extract page dimensions
                var dims = ExtractPageDimensions(sectPr);
                sectionMeta.PageWidthPt = dims.PageWidthPt;
                sectionMeta.PageHeightPt = dims.PageHeightPt;
                sectionMeta.MarginTopPt = dims.MarginTopPt;
                sectionMeta.MarginRightPt = dims.MarginRightPt;
                sectionMeta.MarginBottomPt = dims.MarginBottomPt;
                sectionMeta.MarginLeftPt = dims.MarginLeftPt;
                sectionMeta.ContentWidthPt = dims.ContentWidthPt;
                sectionMeta.ContentHeightPt = dims.ContentHeightPt;
                sectionMeta.HeaderPt = dims.HeaderPt;
                sectionMeta.FooterPt = dims.FooterPt;

                // Detect headers and footers
                DetectHeadersFooters(wordDoc, sectPr, sectionMeta);

                metadata.Sections.Add(sectionMeta);

                globalParagraphIndex += paragraphs.Count;
                globalTableIndex += tables.Count;
                sectionIndex++;
            }

            metadata.TotalParagraphs = globalParagraphIndex;
            metadata.TotalTables = globalTableIndex;

            // Estimate page count based on content
            metadata.EstimatedPageCount = EstimatePageCount(metadata);

            return metadata;
        }

        /// <summary>
        /// Collects section data including section properties and their associated paragraphs/tables.
        /// </summary>
        /// <remarks>
        /// <para>
        /// This method handles the following sectPr placements:
        /// <list type="bullet">
        ///   <item><description>sectPr inside paragraph properties (w:p/w:pPr/w:sectPr) - mid-document section breaks</description></item>
        ///   <item><description>Document-level sectPr at end of body (w:body/w:sectPr) - final section</description></item>
        /// </list>
        /// </para>
        /// <para>
        /// LIMITATION: sectPr elements inside tables or text boxes are NOT detected.
        /// This is an edge case - most documents don't have section breaks inside tables.
        /// See GitHub issue #51 for tracking this enhancement.
        /// </para>
        /// </remarks>
        private static List<(XElement sectPr, List<XElement> paragraphs, List<XElement> tables)> CollectSectionData(XElement body)
        {
            var result = new List<(XElement sectPr, List<XElement> paragraphs, List<XElement> tables)>();
            var currentParagraphs = new List<XElement>();
            var currentTables = new List<XElement>();

            // Get all top-level block elements
            var blockElements = body.Elements()
                .Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sectPr)
                .ToList();

            foreach (var element in blockElements)
            {
                if (element.Name == W.p)
                {
                    currentParagraphs.Add(element);

                    // Check for section break in paragraph properties
                    var pPr = element.Element(W.pPr);
                    var sectPr = pPr?.Element(W.sectPr);
                    if (sectPr != null)
                    {
                        // End of section
                        result.Add((sectPr, new List<XElement>(currentParagraphs), new List<XElement>(currentTables)));
                        currentParagraphs.Clear();
                        currentTables.Clear();
                    }
                }
                else if (element.Name == W.tbl)
                {
                    currentTables.Add(element);
                    // Also count paragraphs inside tables
                    var tableParagraphs = element.Descendants(W.p).ToList();
                    currentParagraphs.AddRange(tableParagraphs);
                }
                else if (element.Name == W.sectPr)
                {
                    // Document-level sectPr (at end of body)
                    result.Add((element, new List<XElement>(currentParagraphs), new List<XElement>(currentTables)));
                    currentParagraphs.Clear();
                    currentTables.Clear();
                }
            }

            // If there's remaining content without a sectPr, use default dimensions
            if (currentParagraphs.Count > 0 || currentTables.Count > 0)
            {
                // Try to find a trailing sectPr
                var trailingSectPr = body.Element(W.sectPr);
                result.Add((trailingSectPr, new List<XElement>(currentParagraphs), new List<XElement>(currentTables)));
            }

            // Ensure at least one section exists
            if (result.Count == 0)
            {
                var defaultSectPr = body.Element(W.sectPr);
                result.Add((defaultSectPr, new List<XElement>(), new List<XElement>()));
            }

            return result;
        }

        /// <summary>
        /// Extracts page dimensions from a sectPr element.
        /// Returns default US Letter dimensions if sectPr is null.
        /// </summary>
        private static (double PageWidthPt, double PageHeightPt, double MarginTopPt, double MarginRightPt,
            double MarginBottomPt, double MarginLeftPt, double ContentWidthPt, double ContentHeightPt,
            double HeaderPt, double FooterPt) ExtractPageDimensions(XElement sectPr)
        {
            // Default to US Letter: 8.5" x 11" (612pt x 792pt) with 1" margins (72pt)
            double pageWidthPt = 612;
            double pageHeightPt = 792;
            double marginTopPt = 72;
            double marginRightPt = 72;
            double marginBottomPt = 72;
            double marginLeftPt = 72;
            double headerPt = 36; // 0.5 inch
            double footerPt = 36;

            if (sectPr != null)
            {
                // Parse page size (w:pgSz) - values are in twips (1/20 of a point)
                var pgSz = sectPr.Element(W.pgSz);
                if (pgSz != null)
                {
                    if (int.TryParse((string)pgSz.Attribute(W._w), out int w))
                        pageWidthPt = w / 20.0;
                    if (int.TryParse((string)pgSz.Attribute(W.h), out int h))
                        pageHeightPt = h / 20.0;
                }

                // Parse page margins (w:pgMar) - values are in twips
                var pgMar = sectPr.Element(W.pgMar);
                if (pgMar != null)
                {
                    if (int.TryParse((string)pgMar.Attribute(W.top), out int top))
                        marginTopPt = top / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.right), out int right))
                        marginRightPt = right / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.bottom), out int bottom))
                        marginBottomPt = bottom / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.left), out int left))
                        marginLeftPt = left / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.header), out int header))
                        headerPt = header / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.footer), out int footer))
                        footerPt = footer / 20.0;
                }
            }

            double contentWidthPt = pageWidthPt - marginLeftPt - marginRightPt;
            double contentHeightPt = pageHeightPt - marginTopPt - marginBottomPt;

            return (pageWidthPt, pageHeightPt, marginTopPt, marginRightPt, marginBottomPt, marginLeftPt,
                contentWidthPt, contentHeightPt, headerPt, footerPt);
        }

        /// <summary>
        /// Detects presence of headers and footers for a section.
        /// </summary>
        private static void DetectHeadersFooters(WordprocessingDocument wordDoc, XElement sectPr, SectionMetadata sectionMeta)
        {
            if (sectPr == null)
                return;

            var mainDocPart = wordDoc.MainDocumentPart;
            bool hasTitlePage = sectPr.Element(W.titlePg) != null;

            // Check header references
            var headerRefs = sectPr.Elements(W.headerReference).ToList();
            foreach (var headerRef in headerRefs)
            {
                var type = (string)headerRef.Attribute(W.type);
                var headerId = (string)headerRef.Attribute(R.id);

                if (string.IsNullOrEmpty(headerId))
                    continue;

                try
                {
                    var headerPart = mainDocPart.GetPartById(headerId) as HeaderPart;
                    if (headerPart != null && headerPart.GetXDocument().Root?.HasElements == true)
                    {
                        switch (type)
                        {
                            case "default":
                                sectionMeta.HasHeader = true;
                                break;
                            case "first":
                                sectionMeta.HasFirstPageHeader = hasTitlePage;
                                break;
                            case "even":
                                sectionMeta.HasEvenPageHeader = true;
                                break;
                        }
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    // Invalid relationship ID - header reference points to non-existent part.
                    // This can happen with corrupted or manually edited documents.
                    // For metadata extraction, we silently skip invalid references.
                }
            }

            // Check footer references
            var footerRefs = sectPr.Elements(W.footerReference).ToList();
            foreach (var footerRef in footerRefs)
            {
                var type = (string)footerRef.Attribute(W.type);
                var footerId = (string)footerRef.Attribute(R.id);

                if (string.IsNullOrEmpty(footerId))
                    continue;

                try
                {
                    var footerPart = mainDocPart.GetPartById(footerId) as FooterPart;
                    if (footerPart != null && footerPart.GetXDocument().Root?.HasElements == true)
                    {
                        switch (type)
                        {
                            case "default":
                                sectionMeta.HasFooter = true;
                                break;
                            case "first":
                                sectionMeta.HasFirstPageFooter = hasTitlePage;
                                break;
                            case "even":
                                sectionMeta.HasEvenPageFooter = true;
                                break;
                        }
                    }
                }
                catch (ArgumentOutOfRangeException)
                {
                    // Invalid relationship ID - footer reference points to non-existent part.
                    // This can happen with corrupted or manually edited documents.
                    // For metadata extraction, we silently skip invalid references.
                }
            }
        }

        /// <summary>
        /// Estimates the total page count based on content volume.
        /// This is a rough estimate using average content density.
        /// </summary>
        private static int EstimatePageCount(DocumentMetadata metadata)
        {
            if (metadata.Sections.Count == 0)
                return 1;

            int totalPages = 0;

            foreach (var section in metadata.Sections)
            {
                // Rough estimate: ~25 paragraphs per page for average document
                // Tables count as ~3 paragraphs each due to their size
                double contentUnits = section.ParagraphCount + (section.TableCount * 3);

                // Adjust for page size (smaller pages = more pages needed)
                double pageAreaRatio = (section.ContentWidthPt * section.ContentHeightPt) / (468.0 * 648.0);
                if (pageAreaRatio > 0)
                {
                    contentUnits /= pageAreaRatio;
                }

                int sectionPages = Math.Max(1, (int)Math.Ceiling(contentUnits / 25.0));
                totalPages += sectionPages;
            }

            return Math.Max(1, totalPages);
        }


        private static void ReverseTableBordersForRtlTables(WordprocessingDocument wordDoc)
        {
            XDocument xd = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var tbl in xd.Descendants(W.tbl))
            {
                var bidiVisual = tbl.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
                if (bidiVisual == null)
                    continue;

                var tblBorders = tbl.Elements(W.tblPr).Elements(W.tblBorders).FirstOrDefault();
                if (tblBorders != null)
                {
                    var left = tblBorders.Element(W.left);
                    if (left != null)
                        left = new XElement(W.right, left.Attributes());

                    var right = tblBorders.Element(W.right);
                    if (right != null)
                        right = new XElement(W.left, right.Attributes());

                    var newTblBorders = new XElement(W.tblBorders,
                        tblBorders.Element(W.top),
                        left,
                        tblBorders.Element(W.bottom),
                        right);
                    tblBorders.ReplaceWith(newTblBorders);
                }

                foreach (var tc in tbl.Elements(W.tr).Elements(W.tc))
                {
                    var tcBorders = tc.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault();
                    if (tcBorders != null)
                    {
                        var left = tcBorders.Element(W.left);
                        if (left != null)
                            left = new XElement(W.right, left.Attributes());

                        var right = tcBorders.Element(W.right);
                        if (right != null)
                            right = new XElement(W.left, right.Attributes());

                        var newTcBorders = new XElement(W.tcBorders,
                            tcBorders.Element(W.top),
                            left,
                            tcBorders.Element(W.bottom),
                            right);
                        tcBorders.ReplaceWith(newTcBorders);
                    }
                }
            }
        }

        private static void ReifyStylesAndClasses(WmlToHtmlConverterSettings htmlConverterSettings, XElement xhtml)
        {
            if (htmlConverterSettings.FabricateCssClasses)
            {
                var usedCssClassNames = new HashSet<string>();
                var elementsThatNeedClasses = xhtml
                    .DescendantsAndSelf()
                    .Select(d => new
                    {
                        Element = d,
                        Styles = d.Annotation<Dictionary<string, string>>(),
                    })
                    .Where(z => z.Styles != null);
                var augmented = elementsThatNeedClasses
                    .Select(p => new
                    {
                        p.Element,
                        p.Styles,
                        StylesString = p.Element.Name.LocalName + "|" + p.Styles.OrderBy(k => k.Key).Select(s => string.Format("{0}: {1};", s.Key, s.Value)).StringConcatenate(),
                    })
                    .GroupBy(p => p.StylesString)
                    .ToList();
                int classCounter = 1000000;
                var sb = new StringBuilder();
                sb.Append(Environment.NewLine);
                foreach (var grp in augmented)
                {
                    string classNameToUse;
                    var firstOne = grp.First();
                    var styles = firstOne.Styles;
                    if (styles.ContainsKey("PtStyleName"))
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix + styles["PtStyleName"];
                        if (usedCssClassNames.Contains(classNameToUse))
                        {
                            classNameToUse = htmlConverterSettings.CssClassPrefix +
                                styles["PtStyleName"] + "-" +
                                classCounter.ToString().Substring(1);
                            classCounter++;
                        }
                    }
                    else
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix +
                            classCounter.ToString().Substring(1);
                        classCounter++;
                    }
                    usedCssClassNames.Add(classNameToUse);
                    sb.Append(firstOne.Element.Name.LocalName + "." + classNameToUse + " {" + Environment.NewLine);
                    foreach (var st in firstOne.Styles.Where(s => s.Key != "PtStyleName"))
                    {
                        var s = "    " + st.Key + ": " + st.Value + ";" + Environment.NewLine;
                        sb.Append(s);
                    }
                    sb.Append("}" + Environment.NewLine);
                    foreach (var gc in grp)
                    {
                        var existingClass = gc.Element.Attribute("class");
                        if (existingClass != null)
                            existingClass.Value = existingClass.Value + " " + classNameToUse;
                        else
                            gc.Element.Add(new XAttribute("class", classNameToUse));
                    }
                }
                var revisionCss = GenerateRevisionCss(htmlConverterSettings);
                var footnoteCss = GenerateFootnoteCss(htmlConverterSettings);
                var headerFooterCss = GenerateHeaderFooterCss(htmlConverterSettings);
                var commentCss = GenerateCommentCss(htmlConverterSettings);
                var paginationCss = GeneratePaginationCss(htmlConverterSettings);
                var annotationCss = GenerateAnnotationCss(htmlConverterSettings);
                var styleValue = htmlConverterSettings.GeneralCss + sb + revisionCss + footnoteCss + headerFooterCss + commentCss + paginationCss + annotationCss + htmlConverterSettings.AdditionalCss;

                SetStyleElementValue(xhtml, styleValue);
            }
            else
            {
                // Previously, the h:style element was not added at this point. However,
                // at least the General CSS will contain important settings.
                var revisionCss = GenerateRevisionCss(htmlConverterSettings);
                var footnoteCss = GenerateFootnoteCss(htmlConverterSettings);
                var headerFooterCss = GenerateHeaderFooterCss(htmlConverterSettings);
                var commentCss = GenerateCommentCss(htmlConverterSettings);
                var paginationCss = GeneratePaginationCss(htmlConverterSettings);
                var annotationCss = GenerateAnnotationCss(htmlConverterSettings);
                SetStyleElementValue(xhtml, htmlConverterSettings.GeneralCss + revisionCss + footnoteCss + headerFooterCss + commentCss + paginationCss + annotationCss + htmlConverterSettings.AdditionalCss);

                foreach (var d in xhtml.DescendantsAndSelf())
                {
                    var style = d.Annotation<Dictionary<string, string>>();
                    if (style == null)
                        continue;
                    var styleValue =
                        style
                        .Where(p => p.Key != "PtStyleName")
                        .OrderBy(p => p.Key)
                        .Select(e => string.Format("{0}: {1};", e.Key, e.Value))
                        .StringConcatenate();
                    XAttribute st = new XAttribute("style", styleValue);
                    if (d.Attribute("style") != null)
                        d.Attribute("style").Value += styleValue;
                    else
                        d.Add(st);
                }
            }
        }

        private static void SetStyleElementValue(XElement xhtml, string styleValue)
        {
            var styleElement = xhtml
                .Descendants(Xhtml.style)
                .FirstOrDefault();
            if (styleElement != null)
                styleElement.Value = styleValue;
            else
            {
                styleElement = new XElement(Xhtml.style, styleValue);
                var head = xhtml.Element(Xhtml.head);
                if (head != null)
                    head.Add(styleElement);
            }
        }

        private static string GenerateRevisionCss(WmlToHtmlConverterSettings settings)
        {
            if (!settings.RenderTrackedChanges)
                return string.Empty;

            var prefix = settings.RevisionCssClassPrefix ?? "rev-";
            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("/* Tracked Changes CSS */");

            // Insertions - underlined green text
            sb.AppendLine($"ins.{prefix}ins {{");
            sb.AppendLine("    text-decoration: underline;");
            sb.AppendLine("    color: #006400;");
            sb.AppendLine("    background-color: #e6ffe6;");
            sb.AppendLine("}");

            // Deletions - strikethrough red text
            sb.AppendLine($"del.{prefix}del {{");
            sb.AppendLine("    text-decoration: line-through;");
            sb.AppendLine("    color: #8b0000;");
            sb.AppendLine("    background-color: #ffe6e6;");
            sb.AppendLine("}");

            // Deletion marker (when content is hidden)
            sb.AppendLine($"span.{prefix}del-marker {{");
            sb.AppendLine("    display: inline-block;");
            sb.AppendLine("    width: 4px;");
            sb.AppendLine("    height: 1em;");
            sb.AppendLine("    background-color: #8b0000;");
            sb.AppendLine("    vertical-align: middle;");
            sb.AppendLine("}");

            // Move source - strikethrough purple text (content moved elsewhere)
            sb.AppendLine($"del.{prefix}move-from {{");
            sb.AppendLine("    text-decoration: line-through;");
            sb.AppendLine("    color: #4b0082;");
            sb.AppendLine("    background-color: #f0e6ff;");
            sb.AppendLine("}");

            // Move destination - underlined purple text (content moved here)
            sb.AppendLine($"ins.{prefix}move-to {{");
            sb.AppendLine("    text-decoration: underline;");
            sb.AppendLine("    color: #4b0082;");
            sb.AppendLine("    background-color: #e6f0ff;");
            sb.AppendLine("}");

            // Move from marker (when content is hidden)
            sb.AppendLine($"span.{prefix}move-from-marker {{");
            sb.AppendLine("    display: inline-block;");
            sb.AppendLine("    width: 4px;");
            sb.AppendLine("    height: 1em;");
            sb.AppendLine("    background-color: #4b0082;");
            sb.AppendLine("    vertical-align: middle;");
            sb.AppendLine("}");

            // Table row inserted
            sb.AppendLine($"tr.{prefix}row-ins {{");
            sb.AppendLine("    background-color: #e6ffe6;");
            sb.AppendLine("}");

            // Table row deleted
            sb.AppendLine($"tr.{prefix}row-del {{");
            sb.AppendLine("    background-color: #ffe6e6;");
            sb.AppendLine("    text-decoration: line-through;");
            sb.AppendLine("}");

            // Paragraph with inserted paragraph mark (paragraph was split)
            sb.AppendLine($".{prefix}para-ins {{");
            sb.AppendLine("    border-left: 3px solid #006400;");
            sb.AppendLine("    padding-left: 4px;");
            sb.AppendLine("}");

            // Paragraph with deleted paragraph mark (paragraphs were merged)
            sb.AppendLine($".{prefix}para-del {{");
            sb.AppendLine("    border-left: 3px solid #8b0000;");
            sb.AppendLine("    padding-left: 4px;");
            sb.AppendLine("}");

            // Deleted paragraph mark indicator (pilcrow)
            sb.AppendLine($"span.{prefix}para-mark-del {{");
            sb.AppendLine("    color: #8b0000;");
            sb.AppendLine("    font-size: 0.8em;");
            sb.AppendLine("    vertical-align: super;");
            sb.AppendLine("}");

            // Format changes (run property revisions)
            sb.AppendLine($"span.{prefix}format-change {{");
            sb.AppendLine("    border-bottom: 2px dotted #ffa500;");
            sb.AppendLine("    cursor: help;");
            sb.AppendLine("}");

            // Table cell inserted
            sb.AppendLine($"td.{prefix}cell-ins {{");
            sb.AppendLine("    background-color: #e6ffe6;");
            sb.AppendLine("}");

            // Table cell deleted
            sb.AppendLine($"td.{prefix}cell-del {{");
            sb.AppendLine("    background-color: #ffe6e6;");
            sb.AppendLine("    text-decoration: line-through;");
            sb.AppendLine("}");

            // Table cell merged
            sb.AppendLine($"td.{prefix}cell-merge {{");
            sb.AppendLine("    background-color: #fff0e6;");
            sb.AppendLine("    border: 2px dashed #ff8c00;");
            sb.AppendLine("}");

            // Author-specific colors if provided
            if (settings.AuthorColors != null)
            {
                foreach (var kvp in settings.AuthorColors)
                {
                    // Escape author name for use in CSS attribute selector
                    var safeAuthor = kvp.Key.Replace("\"", "\\\"");
                    sb.AppendLine($"[data-author=\"{safeAuthor}\"] {{");
                    sb.AppendLine($"    border-left: 3px solid {kvp.Value};");
                    sb.AppendLine($"    padding-left: 2px;");
                    sb.AppendLine("}");
                }
            }

            return sb.ToString();
        }

        private static string GenerateFootnoteCss(WmlToHtmlConverterSettings settings)
        {
            if (!settings.RenderFootnotesAndEndnotes)
                return string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("/* Footnotes and Endnotes CSS */");

            // Footnote/Endnote reference links
            sb.AppendLine("a.footnote-ref, a.endnote-ref {");
            sb.AppendLine("    text-decoration: none;");
            sb.AppendLine("    color: #0066cc;");
            sb.AppendLine("}");

            // Footnotes section
            sb.AppendLine("section.footnotes, section.endnotes {");
            sb.AppendLine("    margin-top: 2em;");
            sb.AppendLine("    font-size: 0.9em;");
            sb.AppendLine("}");

            // Footnote/Endnote list
            sb.AppendLine("section.footnotes ol, section.endnotes ol {");
            sb.AppendLine("    padding-left: 1.5em;");
            sb.AppendLine("}");

            // Back reference links
            sb.AppendLine("a.footnote-backref, a.endnote-backref {");
            sb.AppendLine("    text-decoration: none;");
            sb.AppendLine("    color: #0066cc;");
            sb.AppendLine("    margin-left: 0.5em;");
            sb.AppendLine("}");

            // Per-page footnotes (for paginated mode)
            var prefix = settings.PaginationCssClassPrefix ?? "page-";
            sb.AppendLine();
            sb.AppendLine("/* Per-page Footnotes (Paginated Mode) */");

            // Page footnotes container
            sb.AppendLine($".{prefix}footnotes {{");
            sb.AppendLine("    font-size: 0.85em;");
            sb.AppendLine("    line-height: 1.4;");
            sb.AppendLine("}");

            // Separator line above footnotes - using background instead of border
            // to avoid subpixel rendering issues that can cause the line to disappear during scrolling
            sb.AppendLine($".{prefix}footnotes hr {{");
            sb.AppendLine("    border: none;");
            sb.AppendLine("    height: 1px;");
            sb.AppendLine("    background-color: #999;");
            sb.AppendLine("    width: 33%;");
            sb.AppendLine("    margin: 0 0 6pt 0;");
            sb.AppendLine("    opacity: 0.6;");
            sb.AppendLine("}");

            // Individual footnote item (in registry and on page)
            sb.AppendLine(".footnote-item {");
            sb.AppendLine("    margin-bottom: 4pt;");
            sb.AppendLine("}");

            // Footnote number - inline with superscript styling
            sb.AppendLine(".footnote-number {");
            sb.AppendLine("    font-weight: normal;");
            sb.AppendLine("    display: inline;");
            sb.AppendLine("    vertical-align: super;");
            sb.AppendLine("    font-size: 0.85em;");
            sb.AppendLine("    margin-right: 2pt;");
            sb.AppendLine("}");

            // Footnote content (inline with number)
            sb.AppendLine(".footnote-content {");
            sb.AppendLine("    display: inline;");
            sb.AppendLine("}");

            // Make first paragraph in footnote content inline to flow with number
            // Use :first-of-type instead of :first-child because XML serialization adds
            // whitespace text nodes that would prevent :first-child from matching
            sb.AppendLine(".footnote-content > p:first-of-type {");
            sb.AppendLine("    display: inline;");
            sb.AppendLine("}");

            // Subsequent paragraphs in footnote get normal block display with indent
            sb.AppendLine(".footnote-content > p:not(:first-of-type) {");
            sb.AppendLine("    display: block;");
            sb.AppendLine("    margin-top: 2pt;");
            sb.AppendLine("    margin-left: 12pt;");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static string GenerateHeaderFooterCss(WmlToHtmlConverterSettings settings)
        {
            if (!settings.RenderHeadersAndFooters)
                return string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("/* Document Headers and Footers CSS */");

            // Document header
            sb.AppendLine("header.document-header {");
            sb.AppendLine("    margin-bottom: 1.5em;");
            sb.AppendLine("    padding-bottom: 0.5em;");
            sb.AppendLine("    border-bottom: 1px solid #ccc;");
            sb.AppendLine("}");

            // Document footer
            sb.AppendLine("footer.document-footer {");
            sb.AppendLine("    margin-top: 1.5em;");
            sb.AppendLine("    padding-top: 0.5em;");
            sb.AppendLine("    border-top: 1px solid #ccc;");
            sb.AppendLine("    font-size: 0.9em;");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static string GenerateCommentCss(WmlToHtmlConverterSettings settings)
        {
            if (!settings.RenderComments)
                return string.Empty;

            var prefix = settings.CommentCssClassPrefix ?? "comment-";
            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("/* Comments CSS */");

            // Comment highlights (the highlighted text in the document)
            sb.AppendLine($"span.{prefix}highlight {{");
            sb.AppendLine("    background-color: #fff9c4;");
            sb.AppendLine("    border-bottom: 2px solid #fbc02d;");
            sb.AppendLine("}");

            // Comment marker (the [1] reference link)
            sb.AppendLine($"a.{prefix}marker {{");
            sb.AppendLine("    color: #1976d2;");
            sb.AppendLine("    text-decoration: none;");
            sb.AppendLine("    margin-left: 2px;");
            sb.AppendLine("}");

            sb.AppendLine($"a.{prefix}marker:hover {{");
            sb.AppendLine("    text-decoration: underline;");
            sb.AppendLine("}");

            // Comments section (EndnoteStyle mode)
            sb.AppendLine($"aside.{prefix.TrimEnd('-')}s-section {{");
            sb.AppendLine("    margin-top: 2em;");
            sb.AppendLine("    padding-top: 1em;");
            sb.AppendLine("    border-top: 2px solid #ccc;");
            sb.AppendLine("}");

            sb.AppendLine($"aside.{prefix.TrimEnd('-')}s-section h2 {{");
            sb.AppendLine("    font-size: 1.2em;");
            sb.AppendLine("    margin-bottom: 0.5em;");
            sb.AppendLine("}");

            sb.AppendLine($"ol.{prefix.TrimEnd('-')}s-list {{");
            sb.AppendLine("    list-style: none;");
            sb.AppendLine("    padding: 0;");
            sb.AppendLine("}");

            // Individual comment item
            sb.AppendLine($"li.{prefix.TrimEnd('-')} {{");
            sb.AppendLine("    margin-bottom: 1em;");
            sb.AppendLine("    padding: 0.75em;");
            sb.AppendLine("    background-color: #f5f5f5;");
            sb.AppendLine("    border-left: 3px solid #1976d2;");
            sb.AppendLine("    border-radius: 0 4px 4px 0;");
            sb.AppendLine("}");

            // Comment header
            sb.AppendLine($"div.{prefix}header {{");
            sb.AppendLine("    display: flex;");
            sb.AppendLine("    align-items: center;");
            sb.AppendLine("    gap: 0.5em;");
            sb.AppendLine("    margin-bottom: 0.5em;");
            sb.AppendLine("    font-size: 0.85em;");
            sb.AppendLine("}");

            sb.AppendLine($"span.{prefix}author {{");
            sb.AppendLine("    font-weight: bold;");
            sb.AppendLine("    color: #1976d2;");
            sb.AppendLine("}");

            sb.AppendLine($"span.{prefix}date {{");
            sb.AppendLine("    color: #666;");
            sb.AppendLine("}");

            sb.AppendLine($"a.{prefix}backref {{");
            sb.AppendLine("    margin-left: auto;");
            sb.AppendLine("    text-decoration: none;");
            sb.AppendLine("    color: #1976d2;");
            sb.AppendLine("}");

            // Comment body
            sb.AppendLine($"div.{prefix}body p {{");
            sb.AppendLine("    margin: 0;");
            sb.AppendLine("}");

            // Inline mode - tooltip on hover
            sb.AppendLine($"span.{prefix}highlight[title] {{");
            sb.AppendLine("    cursor: help;");
            sb.AppendLine("}");

            // Margin mode styles
            if (settings.CommentRenderMode == CommentRenderMode.Margin)
            {
                sb.AppendLine();
                sb.AppendLine("/* Margin Mode Comments */");

                // Container that holds both content and margin
                sb.AppendLine($"div.{prefix}margin-container {{");
                sb.AppendLine("    display: flex;");
                sb.AppendLine("    flex-direction: row;");
                sb.AppendLine("    gap: 1em;");
                sb.AppendLine("}");

                // Main content area
                sb.AppendLine($"div.{prefix}margin-content {{");
                sb.AppendLine("    flex: 1;");
                sb.AppendLine("    min-width: 0;");
                sb.AppendLine("}");

                // Margin column for comments
                sb.AppendLine($"aside.{prefix}margin-column {{");
                sb.AppendLine("    width: 250px;");
                sb.AppendLine("    flex-shrink: 0;");
                sb.AppendLine("    position: relative;");
                sb.AppendLine("}");

                // Individual margin comment
                sb.AppendLine($"div.{prefix}margin-note {{");
                sb.AppendLine("    position: relative;");
                sb.AppendLine("    margin-bottom: 0.5em;");
                sb.AppendLine("    padding: 0.5em;");
                sb.AppendLine("    background-color: #fff9c4;");
                sb.AppendLine("    border-left: 3px solid #fbc02d;");
                sb.AppendLine("    border-radius: 0 4px 4px 0;");
                sb.AppendLine("    font-size: 0.85em;");
                sb.AppendLine("    box-shadow: 0 1px 3px rgba(0,0,0,0.1);");
                sb.AppendLine("}");

                // Margin comment header
                sb.AppendLine($"div.{prefix}margin-note-header {{");
                sb.AppendLine("    display: flex;");
                sb.AppendLine("    justify-content: space-between;");
                sb.AppendLine("    align-items: center;");
                sb.AppendLine("    margin-bottom: 0.25em;");
                sb.AppendLine("    font-size: 0.9em;");
                sb.AppendLine("}");

                sb.AppendLine($"span.{prefix}margin-author {{");
                sb.AppendLine("    font-weight: bold;");
                sb.AppendLine("    color: #f57f17;");
                sb.AppendLine("}");

                sb.AppendLine($"span.{prefix}margin-date {{");
                sb.AppendLine("    color: #666;");
                sb.AppendLine("    font-size: 0.85em;");
                sb.AppendLine("}");

                // Margin comment body
                sb.AppendLine($"div.{prefix}margin-note-body {{");
                sb.AppendLine("    color: #333;");
                sb.AppendLine("}");

                sb.AppendLine($"div.{prefix}margin-note-body p {{");
                sb.AppendLine("    margin: 0;");
                sb.AppendLine("}");

                // Back reference link
                sb.AppendLine($"a.{prefix}margin-backref {{");
                sb.AppendLine("    color: #1976d2;");
                sb.AppendLine("    text-decoration: none;");
                sb.AppendLine("    font-size: 0.85em;");
                sb.AppendLine("}");

                sb.AppendLine($"a.{prefix}margin-backref:hover {{");
                sb.AppendLine("    text-decoration: underline;");
                sb.AppendLine("}");

                // Highlighted text in margin mode - add anchor styling
                sb.AppendLine($"span.{prefix}highlight[data-comment-id] {{");
                sb.AppendLine("    cursor: pointer;");
                sb.AppendLine("}");

                // Print styles for margin mode
                sb.AppendLine("@media print {");
                sb.AppendLine($"    div.{prefix}margin-container {{");
                sb.AppendLine("        display: block;");
                sb.AppendLine("    }");
                sb.AppendLine($"    aside.{prefix}margin-column {{");
                sb.AppendLine("        width: auto;");
                sb.AppendLine("        page-break-inside: avoid;");
                sb.AppendLine("    }");
                sb.AppendLine("}");
            }

            return sb.ToString();
        }

        private static string GeneratePaginationCss(WmlToHtmlConverterSettings settings)
        {
            if (settings.RenderPagination != PaginationMode.Paginated)
                return string.Empty;

            var prefix = settings.PaginationCssClassPrefix ?? "page-";
            var scale = settings.PaginationScale > 0 ? settings.PaginationScale : 1.0;
            var sb = new StringBuilder();

            sb.AppendLine();
            sb.AppendLine("/* Pagination CSS */");

            // CSS variable for scale
            sb.AppendLine(":root {");
            sb.AppendLine(string.Format(NumberFormatInfo.InvariantInfo, "    --{0}scale: {1:F2};", prefix, scale));
            sb.AppendLine("}");

            // Staging area - hidden for measurement by client-side JavaScript
            sb.AppendLine($".{prefix}staging {{");
            sb.AppendLine("    position: absolute;");
            sb.AppendLine("    left: -9999px;");
            sb.AppendLine("    visibility: hidden;");
            sb.AppendLine("}");

            // Main container with dark background (PDF.js style)
            sb.AppendLine($".{prefix}container {{");
            sb.AppendLine("    display: flex;");
            sb.AppendLine("    flex-direction: column;");
            sb.AppendLine("    align-items: center;");
            sb.AppendLine("    gap: 20px;");
            sb.AppendLine("    padding: 20px;");
            sb.AppendLine("    background: #525659;");
            sb.AppendLine("    min-height: 100vh;");
            sb.AppendLine("}");

            // Page box with shadow
            sb.AppendLine($".{prefix}box {{");
            sb.AppendLine("    background: white;");
            sb.AppendLine("    box-shadow: 0 2px 8px rgba(0,0,0,0.3);");
            sb.AppendLine("    position: relative;");
            sb.AppendLine("    overflow: hidden;");
            sb.AppendLine("    box-sizing: border-box;");
            sb.AppendLine("}");

            // Content area within page
            sb.AppendLine($".{prefix}content {{");
            sb.AppendLine("    position: absolute;");
            sb.AppendLine("    overflow: hidden;");
            sb.AppendLine("    transform-origin: top left;");
            sb.AppendLine("}");

            // Header area within page (positioned at top, constrained to top margin height)
            sb.AppendLine($".{prefix}header {{");
            sb.AppendLine("    position: absolute;");
            sb.AppendLine("    top: 0;");
            sb.AppendLine("    overflow: hidden;");
            sb.AppendLine("    box-sizing: border-box;");
            sb.AppendLine("    display: flex;");
            sb.AppendLine("    flex-direction: column;");
            sb.AppendLine("    justify-content: flex-end;"); // Align content to bottom of header area
            sb.AppendLine("}");

            // Footer area within page (positioned at bottom, constrained to bottom margin height)
            sb.AppendLine($".{prefix}footer {{");
            sb.AppendLine("    position: absolute;");
            sb.AppendLine("    bottom: 0;");
            sb.AppendLine("    overflow: hidden;");
            sb.AppendLine("    box-sizing: border-box;");
            sb.AppendLine("    display: flex;");
            sb.AppendLine("    flex-direction: column;");
            sb.AppendLine("    justify-content: flex-start;"); // Align content to top of footer area
            sb.AppendLine("}");

            // Page number indicator
            sb.AppendLine($".{prefix}number {{");
            sb.AppendLine("    position: absolute;");
            sb.AppendLine("    bottom: 8px;");
            sb.AppendLine("    width: 100%;");
            sb.AppendLine("    text-align: center;");
            sb.AppendLine("    font-size: 11px;");
            sb.AppendLine("    color: #666;");
            sb.AppendLine("    pointer-events: none;");
            sb.AppendLine("}");

            // Hide system page number when page has a document footer
            sb.AppendLine($".{prefix}box:has(.{prefix}footer) .{prefix}number {{");
            sb.AppendLine("    display: none;");
            sb.AppendLine("}");

            // Page break marker (rendered by client-side as separator)
            sb.AppendLine($".{prefix}break {{");
            sb.AppendLine("    display: none;"); // Hidden in staging; processed by pagination engine
            sb.AppendLine("}");

            // Column break marker
            sb.AppendLine($".{prefix}column-break {{");
            sb.AppendLine("    display: none;");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static string GenerateAnnotationCss(WmlToHtmlConverterSettings settings)
        {
            if (!settings.RenderAnnotations)
                return string.Empty;

            var prefix = settings.AnnotationCssClassPrefix ?? "annot-";
            var sb = new StringBuilder();

            sb.AppendLine();
            sb.AppendLine("/* Custom Annotations CSS */");

            // Annotation highlight base
            sb.AppendLine($".{prefix}highlight {{");
            sb.AppendLine("    position: relative;");
            sb.AppendLine("    display: inline;");
            sb.AppendLine("    background-color: color-mix(in srgb, var(--annot-color, #FFFF00) 35%, transparent);");
            sb.AppendLine("    border-bottom: 2px solid var(--annot-color, #FFFF00);");
            sb.AppendLine("    padding: 1px 2px;");
            sb.AppendLine("    border-radius: 2px;");
            sb.AppendLine("    transition: background-color 0.15s ease;");
            sb.AppendLine("}");

            // Hover state
            sb.AppendLine($".{prefix}highlight:hover {{");
            sb.AppendLine("    background-color: color-mix(in srgb, var(--annot-color, #FFFF00) 50%, transparent);");
            sb.AppendLine("}");

            // Floating label above highlight
            sb.AppendLine($".{prefix}label {{");
            sb.AppendLine("    position: absolute;");
            sb.AppendLine("    top: -1.7em;");
            sb.AppendLine("    left: 0;");
            sb.AppendLine("    font-size: 0.7em;");
            sb.AppendLine("    font-weight: 600;");
            sb.AppendLine("    background: var(--annot-color, #FFFF00);");
            sb.AppendLine("    color: #000;");
            sb.AppendLine("    padding: 2px 6px;");
            sb.AppendLine("    border-radius: 3px;");
            sb.AppendLine("    white-space: nowrap;");
            sb.AppendLine("    box-shadow: 0 1px 3px rgba(0,0,0,0.2);");
            sb.AppendLine("    z-index: 100;");
            sb.AppendLine("    pointer-events: none;");
            sb.AppendLine("    line-height: 1.2;");
            sb.AppendLine("}");

            // Only show label on first segment of multi-paragraph annotations
            sb.AppendLine($".{prefix}continuation .{prefix}label,");
            sb.AppendLine($".{prefix}end .{prefix}label {{");
            sb.AppendLine("    display: none;");
            sb.AppendLine("}");

            // Inline label mode
            sb.AppendLine($".{prefix}highlight[data-label-mode=\"inline\"] .{prefix}label {{");
            sb.AppendLine("    position: static;");
            sb.AppendLine("    display: inline;");
            sb.AppendLine("    margin-right: 4px;");
            sb.AppendLine("    font-size: 0.8em;");
            sb.AppendLine("    vertical-align: middle;");
            sb.AppendLine("}");

            // Tooltip mode - show on hover
            sb.AppendLine($".{prefix}highlight[data-label-mode=\"tooltip\"] .{prefix}label {{");
            sb.AppendLine("    display: none;");
            sb.AppendLine("    top: auto;");
            sb.AppendLine("    bottom: 100%;");
            sb.AppendLine("    margin-bottom: 4px;");
            sb.AppendLine("}");

            sb.AppendLine($".{prefix}highlight[data-label-mode=\"tooltip\"]:hover .{prefix}label {{");
            sb.AppendLine("    display: block;");
            sb.AppendLine("}");

            // No label mode
            sb.AppendLine($".{prefix}highlight[data-label-mode=\"none\"] .{prefix}label {{");
            sb.AppendLine("    display: none;");
            sb.AppendLine("}");

            // Handle nested/overlapping annotations
            sb.AppendLine($".{prefix}highlight .{prefix}highlight {{");
            sb.AppendLine("    background: none;");
            sb.AppendLine("    border-bottom-style: dashed;");
            sb.AppendLine("    padding: 0;");
            sb.AppendLine("}");

            // Ensure labels don't overlap badly for nested annotations
            sb.AppendLine($".{prefix}highlight .{prefix}highlight .{prefix}label {{");
            sb.AppendLine("    top: -3.2em;");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static object ConvertToHtmlTransform(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XNode node,
            bool suppressTrailingWhiteSpace,
            decimal currentMarginLeft)
        {
            var element = node as XElement;
            if (element == null) return null;

            // Transform the w:document element to the XHTML h:html element.
            // The h:head element is laid out based on the W3C's recommended layout, i.e.,
            // the charset (using the HTML5-compliant form), the title (which is always
            // there but possibly empty), and other meta tags.
            if (element.Name == W.document)
            {
                return new XElement(Xhtml.html,
                    new XElement(Xhtml.head,
                        new XElement(Xhtml.meta, new XAttribute("charset", "UTF-8")),
                        settings.PageTitle != null
                            ? new XElement(Xhtml.title, new XText(settings.PageTitle))
                            : new XElement(Xhtml.title, new XText(string.Empty)),
                        new XElement(Xhtml.meta,
                            new XAttribute("name", "Generator"),
                            new XAttribute("content", "PowerTools for Open XML"))),
                    element.Elements()
                        .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));
            }

            // Transform the w:body element to the XHTML h:body element.
            if (element.Name == W.body)
            {
                var bodyContent = new List<object>();

                // In pagination mode, headers/footers are rendered into the registry
                // and cloned per-page by the client-side pagination engine.
                // Only render document-level headers when NOT using pagination.
                bool usePaginatedHeadersFooters =
                    settings.RenderHeadersAndFooters &&
                    settings.RenderPagination == PaginationMode.Paginated;

                // Add headers at the top if enabled (non-paginated mode only)
                if (settings.RenderHeadersAndFooters && !usePaginatedHeadersFooters)
                {
                    var headersSection = RenderHeadersSection(wordDoc, settings);
                    if (headersSection != null)
                        bodyContent.Add(headersSection);
                }

                // Get main document content
                var mainContent = CreateSectionDivs(wordDoc, settings, element);

                // For margin mode, wrap content in a flex container with margin column
                if (settings.RenderComments && settings.CommentRenderMode == CommentRenderMode.Margin)
                {
                    var prefix = settings.CommentCssClassPrefix ?? "comment-";
                    var tracker = GetCommentTracker(element);

                    var marginContainer = new XElement(Xhtml.div,
                        new XAttribute("class", prefix + "margin-container"),
                        new XElement(Xhtml.div,
                            new XAttribute("class", prefix + "margin-content"),
                            mainContent),
                        RenderMarginCommentsColumn(wordDoc, settings, tracker, prefix));

                    bodyContent.Add(marginContainer);
                }
                else
                {
                    bodyContent.Add(mainContent);
                }

                // Add footnotes and endnotes sections if enabled
                if (settings.RenderFootnotesAndEndnotes)
                {
                    // In paginated mode, use footnote registry for client-side per-page distribution
                    bool usePaginatedFootnotes = settings.RenderPagination == PaginationMode.Paginated;

                    if (usePaginatedFootnotes)
                    {
                        // Footnote registry goes into the staging area for the pagination engine
                        // The registry is already added to the staging area in CreateSectionDivs
                        // (we'll handle this there to keep it with the staging container)
                    }
                    else
                    {
                        // Non-paginated mode: render footnotes section at the end
                        var footnotesSection = RenderFootnotesSection(wordDoc, settings);
                        if (footnotesSection != null)
                            bodyContent.Add(footnotesSection);
                    }

                    // Endnotes always render at document end (not per-page)
                    var endnotesSection = RenderEndnotesSection(wordDoc, settings);
                    if (endnotesSection != null)
                        bodyContent.Add(endnotesSection);
                }

                // Add footers at the bottom if enabled (non-paginated mode only)
                if (settings.RenderHeadersAndFooters && !usePaginatedHeadersFooters)
                {
                    var footersSection = RenderFootersSection(wordDoc, settings);
                    if (footersSection != null)
                        bodyContent.Add(footersSection);
                }

                // Add comments section if enabled (EndnoteStyle mode)
                if (settings.RenderComments && settings.CommentRenderMode == CommentRenderMode.EndnoteStyle)
                {
                    var tracker = GetCommentTracker(element);
                    if (tracker != null)
                    {
                        var commentsSection = RenderCommentsSection(wordDoc, settings, tracker);
                        if (commentsSection != null)
                            bodyContent.Add(commentsSection);
                    }
                }

                return new XElement(Xhtml.body, bodyContent);
            }

            // Transform the w:p element to the XHTML h:h1-h6 or h:p element (if the previous paragraph does not
            // have a style separator).
            if (element.Name == W.p)
            {
                return ProcessParagraph(wordDoc, settings, element, suppressTrailingWhiteSpace, currentMarginLeft);
            }

            // Transform hyperlinks to the XHTML h:a element.
            if (element.Name == W.hyperlink && element.Attribute(R.id) != null)
            {
                try
                {
                    var a = new XElement(Xhtml.a,
                        new XAttribute("href",
                            wordDoc.MainDocumentPart
                                .HyperlinkRelationships
                                .First(x => x.Id == (string)element.Attribute(R.id))
                                .Uri
                            ),
                        element.Elements(W.r).Select(run => ConvertRun(wordDoc, settings, run))
                        );
                    if (!a.Nodes().Any())
                        a.Add(new XText(""));
                    return a;
                }
                catch (UriFormatException)
                {
                    return element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
                }
            }

            // Transform hyperlinks to bookmarks to the XHTML h:a element.
            if (element.Name == W.hyperlink && element.Attribute(W.anchor) != null)
            {
                return ProcessHyperlinkToBookmark(wordDoc, settings, element);
            }

            // Transform contents of runs.
            if (element.Name == W.r)
            {
                return ConvertRun(wordDoc, settings, element);
            }

            // Transform w:bookmarkStart into anchor (and track annotation ranges)
            if (element.Name == W.bookmarkStart)
            {
                return ProcessBookmarkStart(settings, element);
            }

            // Handle bookmarkEnd for closing annotation ranges
            if (element.Name == W.bookmarkEnd)
            {
                return ProcessBookmarkEnd(settings, element);
            }

            // Transform every w:t element to a text node.
            if (element.Name == W.t)
            {
                // We don't need to convert characters to entities in a UTF-8 document.
                // Further, we don't need &nbsp; entities for significant whitespace
                // because we are wrapping the text nodes in <span> elements within
                // which all whitespace is significant.
                return new XText(element.Value);
            }

            // Transform symbols to spans
            if (element.Name == W.sym)
            {
                var cs = (string)element.Attribute(W._char);
                var c = Convert.ToInt32(cs, 16);
                return new XElement(Xhtml.span, new XEntity(string.Format("#{0}", c)));
            }

            // Transform tabs that have the pt:TabWidth attribute set
            if (element.Name == W.tab)
            {
                return ProcessTab(element);
            }

            // Transform w:br to h:br (or page break div in pagination mode).
            if (element.Name == W.br || element.Name == W.cr)
            {
                return ProcessBreak(element, settings);
            }

            // Transform w:noBreakHyphen to '-'
            if (element.Name == W.noBreakHyphen)
            {
                return new XText("-");
            }

            // Transform w:tbl to h:tbl.
            if (element.Name == W.tbl)
            {
                return ProcessTable(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform w:tr to h:tr.
            if (element.Name == W.tr)
            {
                return ProcessTableRow(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform w:tc to h:td.
            if (element.Name == W.tc)
            {
                return ProcessTableCell(wordDoc, settings, element);
            }

            // Transform images
            if (element.Name == W.drawing || element.Name == W.pict || element.Name == W._object)
            {
                return ProcessImage(wordDoc, element, settings.ImageHandler);
            }

            // Transform content controls.
            if (element.Name == W.sdt)
            {
                return ProcessContentControl(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform smart tags and simple fields.
            if (element.Name == W.smartTag || element.Name == W.fldSimple)
            {
                return CreateBorderDivs(wordDoc, settings, element.Elements());
            }

            // Transform tracked changes - insertions
            if (element.Name == W.ins)
            {
                return ProcessInsertion(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform tracked changes - deletions
            if (element.Name == W.del)
            {
                return ProcessDeletion(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform deleted text (w:delText) to text node when rendering tracked changes
            if (element.Name == W.delText)
            {
                if (settings.RenderTrackedChanges && settings.ShowDeletedContent)
                {
                    return new XText(element.Value);
                }
                return null;
            }

            // Transform tracked changes - move from (source of moved content)
            if (element.Name == W.moveFrom)
            {
                return ProcessMoveFrom(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform tracked changes - move to (destination of moved content)
            if (element.Name == W.moveTo)
            {
                return ProcessMoveTo(wordDoc, settings, element, currentMarginLeft);
            }

            // Handle footnote references
            if (element.Name == W.footnoteReference)
            {
                return ProcessFootnoteReference(wordDoc, settings, element);
            }

            // Handle endnote references
            if (element.Name == W.endnoteReference)
            {
                return ProcessEndnoteReference(wordDoc, settings, element);
            }

            // Handle comment range start
            if (element.Name == W.commentRangeStart)
            {
                return ProcessCommentRangeStart(settings, element);
            }

            // Handle comment range end
            if (element.Name == W.commentRangeEnd)
            {
                return ProcessCommentRangeEnd(settings, element);
            }

            // Handle comment reference (the marker in the text)
            if (element.Name == W.commentReference)
            {
                return ProcessCommentReference(settings, element);
            }

            // Ignore element.
            return null;
        }

        private static object ProcessHyperlinkToBookmark(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            var style = new Dictionary<string, string>();
            var a = new XElement(Xhtml.a,
                new XAttribute("href", "#" + (string) element.Attribute(W.anchor)),
                element.Elements(W.r).Select(run => ConvertRun(wordDoc, settings, run)));
            if (!a.Nodes().Any())
                a.Add(new XText(""));
            style.Add("text-decoration", "none");
            a.AddAnnotation(style);
            return a;
        }

        private static object ProcessBookmarkStart(WmlToHtmlConverterSettings settings, XElement element)
        {
            var name = (string) element.Attribute(W.name);
            if (name == null) return null;

            // Check if this is an annotation bookmark and track it
            if (settings.RenderAnnotations && name.StartsWith(AnnotationManager.BookmarkPrefix))
            {
                var tracker = GetAnnotationTracker(element);
                if (tracker != null && tracker.BookmarkToAnnotationId.TryGetValue(name, out var annotationId))
                {
                    tracker.OpenRanges.Add(annotationId);
                }
            }

            var style = new Dictionary<string, string>();
            var a = new XElement(Xhtml.a,
                new XAttribute("id", name),
                new XText(""));
            if (!a.Nodes().Any())
                a.Add(new XText(""));
            style.Add("text-decoration", "none");
            a.AddAnnotation(style);
            return a;
        }

        private static object ProcessBookmarkEnd(WmlToHtmlConverterSettings settings, XElement element)
        {
            // bookmarkEnd uses w:id to reference the bookmark, not w:name
            // We need to find the corresponding bookmarkStart to get the name
            var id = (string)element.Attribute(W.id);
            if (id == null)
                return null;

            if (!settings.RenderAnnotations)
                return null;

            var tracker = GetAnnotationTracker(element);
            if (tracker == null)
                return null;

            // Find the corresponding bookmarkStart in the document to get the name
            var root = element.AncestorsAndSelf().LastOrDefault();
            if (root != null)
            {
                var bookmarkStart = root.Descendants(W.bookmarkStart)
                    .FirstOrDefault(bs => (string)bs.Attribute(W.id) == id);

                if (bookmarkStart != null)
                {
                    var name = (string)bookmarkStart.Attribute(W.name);
                    if (name != null && name.StartsWith(AnnotationManager.BookmarkPrefix) &&
                        tracker.BookmarkToAnnotationId.TryGetValue(name, out var annotationId))
                    {
                        tracker.OpenRanges.Remove(annotationId);
                    }
                }
            }

            return null;
        }

        private static object ProcessInsertion(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
        {
            if (!settings.RenderTrackedChanges)
            {
                // When not rendering tracked changes, just process children normally
                return element.Elements()
                    .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
            }

            var ins = new XElement(Xhtml.ins);

            // Add CSS class
            var className = (settings.RevisionCssClassPrefix ?? "rev-") + "ins";
            ins.Add(new XAttribute("class", className));

            // Add metadata if requested
            if (settings.IncludeRevisionMetadata)
            {
                var author = (string)element.Attribute(W.author);
                var date = (string)element.Attribute(W.date);

                if (author != null)
                    ins.Add(new XAttribute("data-author", author));
                if (date != null)
                    ins.Add(new XAttribute("data-date", date));
            }

            // Process children
            var content = element.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                .ToList();

            ins.Add(content);

            // Ensure the element isn't empty (browsers may ignore empty ins/del)
            if (!ins.Nodes().Any())
                ins.Add(new XText(""));

            return ins;
        }

        private static object ProcessDeletion(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
        {
            if (!settings.RenderTrackedChanges)
            {
                // When not rendering tracked changes, deletions are removed entirely
                return null;
            }

            if (!settings.ShowDeletedContent)
            {
                // Show marker but not content
                var marker = new XElement(Xhtml.span,
                    new XAttribute("class", (settings.RevisionCssClassPrefix ?? "rev-") + "del-marker"),
                    new XAttribute("title", "Deleted content"));
                return marker;
            }

            var del = new XElement(Xhtml.del);

            // Add CSS class
            var className = (settings.RevisionCssClassPrefix ?? "rev-") + "del";
            del.Add(new XAttribute("class", className));

            // Add metadata if requested
            if (settings.IncludeRevisionMetadata)
            {
                var author = (string)element.Attribute(W.author);
                var date = (string)element.Attribute(W.date);

                if (author != null)
                    del.Add(new XAttribute("data-author", author));
                if (date != null)
                    del.Add(new XAttribute("data-date", date));
            }

            // Process children - note: w:del contains w:delText instead of w:t
            var content = element.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                .ToList();

            del.Add(content);

            // Ensure the element isn't empty
            if (!del.Nodes().Any())
                del.Add(new XText(""));

            return del;
        }

        private static object ProcessMoveFrom(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
        {
            if (!settings.RenderTrackedChanges)
            {
                // When not rendering tracked changes, move sources are removed (content moved elsewhere)
                return null;
            }

            // If not rendering move operations separately, treat as deletion
            if (!settings.RenderMoveOperations)
            {
                return ProcessDeletion(wordDoc, settings, element, currentMarginLeft);
            }

            if (!settings.ShowDeletedContent)
            {
                // Show marker but not content
                var marker = new XElement(Xhtml.span,
                    new XAttribute("class", (settings.RevisionCssClassPrefix ?? "rev-") + "move-from-marker"),
                    new XAttribute("title", "Moved content (source)"));
                return marker;
            }

            var del = new XElement(Xhtml.del);

            // Add CSS class for move source
            var className = (settings.RevisionCssClassPrefix ?? "rev-") + "move-from";
            del.Add(new XAttribute("class", className));

            // Add metadata if requested
            if (settings.IncludeRevisionMetadata)
            {
                var author = (string)element.Attribute(W.author);
                var date = (string)element.Attribute(W.date);
                var moveId = (string)element.Attribute(W.id);

                if (author != null)
                    del.Add(new XAttribute("data-author", author));
                if (date != null)
                    del.Add(new XAttribute("data-date", date));
                if (moveId != null)
                    del.Add(new XAttribute("data-move-id", moveId));
            }

            // Process children - move source contains delText like w:del
            var content = element.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                .ToList();

            del.Add(content);

            // Ensure the element isn't empty
            if (!del.Nodes().Any())
                del.Add(new XText(""));

            return del;
        }

        private static object ProcessMoveTo(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
        {
            if (!settings.RenderTrackedChanges)
            {
                // When not rendering tracked changes, just process children normally
                return element.Elements()
                    .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
            }

            // If not rendering move operations separately, treat as insertion
            if (!settings.RenderMoveOperations)
            {
                return ProcessInsertion(wordDoc, settings, element, currentMarginLeft);
            }

            var ins = new XElement(Xhtml.ins);

            // Add CSS class for move destination
            var className = (settings.RevisionCssClassPrefix ?? "rev-") + "move-to";
            ins.Add(new XAttribute("class", className));

            // Add metadata if requested
            if (settings.IncludeRevisionMetadata)
            {
                var author = (string)element.Attribute(W.author);
                var date = (string)element.Attribute(W.date);
                var moveId = (string)element.Attribute(W.id);

                if (author != null)
                    ins.Add(new XAttribute("data-author", author));
                if (date != null)
                    ins.Add(new XAttribute("data-date", date));
                if (moveId != null)
                    ins.Add(new XAttribute("data-move-id", moveId));
            }

            // Process children
            var content = element.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                .ToList();

            ins.Add(content);

            // Ensure the element isn't empty
            if (!ins.Nodes().Any())
                ins.Add(new XText(""));

            return ins;
        }

        private static object ProcessFootnoteReference(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement element)
        {
            if (!settings.RenderFootnotesAndEndnotes)
            {
                return null;
            }

            var footnoteId = (string)element.Attribute(W.id);
            if (footnoteId == null)
                return null;

            var style = new Dictionary<string, string>
            {
                { "vertical-align", "super" },
                { "font-size", "smaller" }
            };

            var anchor = new XElement(Xhtml.a,
                new XAttribute("href", $"#footnote-{footnoteId}"),
                new XAttribute("id", $"footnote-ref-{footnoteId}"),
                new XAttribute("class", "footnote-ref"),
                new XAttribute("data-footnote-id", footnoteId), // For pagination engine to track footnotes per page
                new XText($"[{footnoteId}]"));

            anchor.AddAnnotation(style);
            return anchor;
        }

        private static object ProcessEndnoteReference(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement element)
        {
            if (!settings.RenderFootnotesAndEndnotes)
            {
                return null;
            }

            var endnoteId = (string)element.Attribute(W.id);
            if (endnoteId == null)
                return null;

            var style = new Dictionary<string, string>
            {
                { "vertical-align", "super" },
                { "font-size", "smaller" }
            };

            var anchor = new XElement(Xhtml.a,
                new XAttribute("href", $"#endnote-{endnoteId}"),
                new XAttribute("id", $"endnote-ref-{endnoteId}"),
                new XAttribute("class", "endnote-ref"),
                new XText($"[{endnoteId}]"));

            anchor.AddAnnotation(style);
            return anchor;
        }

        private static XElement RenderFootnotesSection(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings)
        {
            var footnotesPart = wordDoc.MainDocumentPart.FootnotesPart;
            if (footnotesPart == null)
                return null;

            var footnotesXDoc = footnotesPart.GetXDocument();
            var footnotes = footnotesXDoc.Root?.Elements(W.footnote)
                .Where(fn =>
                {
                    var typeAttr = (string)fn.Attribute(W.type);
                    // Skip separator and continuationSeparator footnotes
                    return typeAttr != "separator" && typeAttr != "continuationSeparator";
                })
                .ToList();

            if (footnotes == null || !footnotes.Any())
                return null;

            var footnotesSection = new XElement(Xhtml.section,
                new XAttribute("class", "footnotes"),
                new XElement(Xhtml.hr),
                new XElement(Xhtml.ol,
                    footnotes.Select(fn => RenderFootnoteItem(wordDoc, settings, fn, "footnote"))));

            return footnotesSection;
        }

        private static XElement RenderEndnotesSection(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings)
        {
            var endnotesPart = wordDoc.MainDocumentPart.EndnotesPart;
            if (endnotesPart == null)
                return null;

            var endnotesXDoc = endnotesPart.GetXDocument();
            var endnotes = endnotesXDoc.Root?.Elements(W.endnote)
                .Where(en =>
                {
                    var typeAttr = (string)en.Attribute(W.type);
                    // Skip separator and continuationSeparator endnotes
                    return typeAttr != "separator" && typeAttr != "continuationSeparator";
                })
                .ToList();

            if (endnotes == null || !endnotes.Any())
                return null;

            var endnotesSection = new XElement(Xhtml.section,
                new XAttribute("class", "endnotes"),
                new XElement(Xhtml.hr),
                new XElement(Xhtml.ol,
                    endnotes.Select(en => RenderFootnoteItem(wordDoc, settings, en, "endnote"))));

            return endnotesSection;
        }

        private static XElement RenderFootnoteItem(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, XElement noteElement, string noteType)
        {
            var noteId = (string)noteElement.Attribute(W.id);
            if (noteId == null)
                return null;

            // Convert the content of the footnote/endnote
            var content = noteElement.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .ToList();

            var li = new XElement(Xhtml.li,
                new XAttribute("id", $"{noteType}-{noteId}"),
                new XAttribute("value", noteId),
                content,
                new XText(" "),
                new XElement(Xhtml.a,
                    new XAttribute("href", $"#{noteType}-ref-{noteId}"),
                    new XAttribute("class", $"{noteType}-backref"),
                    new XText("↩")));

            return li;
        }

        /// <summary>
        /// Renders a footnote registry for paginated mode.
        /// In paginated mode, footnotes are stored in a hidden registry and distributed
        /// to each page by the client-side pagination engine based on which footnote
        /// references appear on each page.
        /// </summary>
        private static XElement RenderPaginatedFootnoteRegistry(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings)
        {
            var footnotesPart = wordDoc.MainDocumentPart.FootnotesPart;
            if (footnotesPart == null)
                return null;

            var footnotesXDoc = footnotesPart.GetXDocument();
            var footnotes = footnotesXDoc.Root?.Elements(W.footnote)
                .Where(fn =>
                {
                    var typeAttr = (string)fn.Attribute(W.type);
                    // Skip separator and continuationSeparator footnotes
                    return typeAttr != "separator" && typeAttr != "continuationSeparator";
                })
                .ToList();

            if (footnotes == null || !footnotes.Any())
                return null;

            var registry = new XElement(Xhtml.div,
                new XAttribute("id", "pagination-footnote-registry"),
                new XAttribute("style", "display:none"));

            foreach (var fn in footnotes)
            {
                var footnoteId = (string)fn.Attribute(W.id);
                if (footnoteId == null)
                    continue;

                // Convert the content of the footnote
                var content = fn.Elements()
                    .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                    .ToList();

                // Create a footnote item div with data attribute for lookup
                var footnoteItem = new XElement(Xhtml.div,
                    new XAttribute("data-footnote-id", footnoteId),
                    new XAttribute("class", "footnote-item"),
                    new XElement(Xhtml.span,
                        new XAttribute("class", "footnote-number"),
                        new XText($"{footnoteId}. ")),
                    new XElement(Xhtml.span,
                        new XAttribute("class", "footnote-content"),
                        content),
                    new XText(" "),
                    new XElement(Xhtml.a,
                        new XAttribute("href", $"#footnote-ref-{footnoteId}"),
                        new XAttribute("class", "footnote-backref"),
                        new XText("↩")));

                registry.Add(footnoteItem);
            }

            return registry.HasElements ? registry : null;
        }

        private static XElement RenderHeadersSection(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings)
        {
            var headerParts = wordDoc.MainDocumentPart.HeaderParts.ToList();
            if (!headerParts.Any())
                return null;

            // Get the section properties from the document body or last sectPr
            var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
            var body = mainDoc.Root?.Element(W.body);
            var sectPr = body?.Element(W.sectPr) ?? body?.Elements(W.p).LastOrDefault()?.Element(W.pPr)?.Element(W.sectPr);

            // Find the default header reference
            var defaultHeaderRef = sectPr?.Elements(W.headerReference)
                .FirstOrDefault(hr => (string)hr.Attribute(W.type) == "default");

            if (defaultHeaderRef == null)
            {
                // If no default header, try to get the first header
                defaultHeaderRef = sectPr?.Elements(W.headerReference).FirstOrDefault();
            }

            if (defaultHeaderRef == null && headerParts.Any())
            {
                // Just use the first header part if no reference found
                var firstHeaderPart = headerParts.First();
                return RenderHeaderPart(wordDoc, settings, firstHeaderPart);
            }

            if (defaultHeaderRef != null)
            {
                var headerId = (string)defaultHeaderRef.Attribute(R.id);
                if (headerId != null)
                {
                    var headerPart = wordDoc.MainDocumentPart.GetPartById(headerId) as HeaderPart;
                    if (headerPart != null)
                    {
                        return RenderHeaderPart(wordDoc, settings, headerPart);
                    }
                }
            }

            return null;
        }

        private static XElement RenderHeaderPart(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, HeaderPart headerPart)
        {
            var headerXDoc = headerPart.GetXDocument();
            var headerRoot = headerXDoc.Root;

            if (headerRoot == null || !headerRoot.Elements().Any())
                return null;

            // Convert the content of the header
            var content = headerRoot.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .Where(c => c != null)
                .ToList();

            if (!content.Any())
                return null;

            return new XElement(Xhtml.header,
                new XAttribute("class", "document-header"),
                content);
        }

        private static XElement RenderFootersSection(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings)
        {
            var footerParts = wordDoc.MainDocumentPart.FooterParts.ToList();
            if (!footerParts.Any())
                return null;

            // Get the section properties from the document body or last sectPr
            var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
            var body = mainDoc.Root?.Element(W.body);
            var sectPr = body?.Element(W.sectPr) ?? body?.Elements(W.p).LastOrDefault()?.Element(W.pPr)?.Element(W.sectPr);

            // Find the default footer reference
            var defaultFooterRef = sectPr?.Elements(W.footerReference)
                .FirstOrDefault(fr => (string)fr.Attribute(W.type) == "default");

            if (defaultFooterRef == null)
            {
                // If no default footer, try to get the first footer
                defaultFooterRef = sectPr?.Elements(W.footerReference).FirstOrDefault();
            }

            if (defaultFooterRef == null && footerParts.Any())
            {
                // Just use the first footer part if no reference found
                var firstFooterPart = footerParts.First();
                return RenderFooterPart(wordDoc, settings, firstFooterPart);
            }

            if (defaultFooterRef != null)
            {
                var footerId = (string)defaultFooterRef.Attribute(R.id);
                if (footerId != null)
                {
                    var footerPart = wordDoc.MainDocumentPart.GetPartById(footerId) as FooterPart;
                    if (footerPart != null)
                    {
                        return RenderFooterPart(wordDoc, settings, footerPart);
                    }
                }
            }

            return null;
        }

        private static XElement RenderFooterPart(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, FooterPart footerPart)
        {
            var footerXDoc = footerPart.GetXDocument();
            var footerRoot = footerXDoc.Root;

            if (footerRoot == null || !footerRoot.Elements().Any())
                return null;

            // Convert the content of the footer
            var content = footerRoot.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .Where(c => c != null)
                .ToList();

            if (!content.Any())
                return null;

            return new XElement(Xhtml.footer,
                new XAttribute("class", "document-footer"),
                content);
        }

        /// <summary>
        /// Renders headers and footers into a hidden registry for pagination mode.
        /// Each header/footer is stored with section index and type metadata for client-side cloning.
        /// </summary>
        private static XElement RenderPaginatedHeaderFooterRegistry(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement bodyElement)
        {
            var registry = new XElement(Xhtml.div,
                new XAttribute("id", "pagination-hf-registry"),
                new XAttribute("style", "display:none"));

            // Collect all section properties from the document
            var sectionProperties = new List<XElement>();

            // Get section properties from paragraph annotations
            foreach (var para in bodyElement.Descendants(W.p))
            {
                var sectAnnotation = para.Annotation<SectionAnnotation>();
                if (sectAnnotation?.SectionElement != null)
                {
                    var sectPr = sectAnnotation.SectionElement;
                    if (!sectionProperties.Contains(sectPr))
                        sectionProperties.Add(sectPr);
                }
            }

            // If no sections found from annotations, get from document body
            if (sectionProperties.Count == 0)
            {
                var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
                var body = mainDoc.Root?.Element(W.body);
                var sectPr = body?.Element(W.sectPr) ?? body?.Elements(W.p).LastOrDefault()?.Element(W.pPr)?.Element(W.sectPr);
                if (sectPr != null)
                    sectionProperties.Add(sectPr);
            }

            // Render headers/footers for each section
            for (int sectionIndex = 0; sectionIndex < sectionProperties.Count; sectionIndex++)
            {
                var sectPr = sectionProperties[sectionIndex];

                // Check if section has different first page headers/footers
                bool hasTitlePage = sectPr.Element(W.titlePg) != null;

                // Render default header
                var defaultHeaderContent = RenderHeaderForSection(wordDoc, settings, sectPr, "default");
                if (defaultHeaderContent != null)
                {
                    registry.Add(new XElement(Xhtml.div,
                        new XAttribute("data-section", sectionIndex),
                        new XAttribute("data-hf-type", "header-default"),
                        defaultHeaderContent));
                }

                // Render first page header if different first page is enabled
                if (hasTitlePage)
                {
                    var firstHeaderContent = RenderHeaderForSection(wordDoc, settings, sectPr, "first");
                    if (firstHeaderContent != null)
                    {
                        registry.Add(new XElement(Xhtml.div,
                            new XAttribute("data-section", sectionIndex),
                            new XAttribute("data-hf-type", "header-first"),
                            firstHeaderContent));
                    }
                }

                // Render even header if exists
                var evenHeaderContent = RenderHeaderForSection(wordDoc, settings, sectPr, "even");
                if (evenHeaderContent != null)
                {
                    registry.Add(new XElement(Xhtml.div,
                        new XAttribute("data-section", sectionIndex),
                        new XAttribute("data-hf-type", "header-even"),
                        evenHeaderContent));
                }

                // Render default footer
                var defaultFooterContent = RenderFooterForSection(wordDoc, settings, sectPr, "default");
                if (defaultFooterContent != null)
                {
                    registry.Add(new XElement(Xhtml.div,
                        new XAttribute("data-section", sectionIndex),
                        new XAttribute("data-hf-type", "footer-default"),
                        defaultFooterContent));
                }

                // Render first page footer if different first page is enabled
                if (hasTitlePage)
                {
                    var firstFooterContent = RenderFooterForSection(wordDoc, settings, sectPr, "first");
                    if (firstFooterContent != null)
                    {
                        registry.Add(new XElement(Xhtml.div,
                            new XAttribute("data-section", sectionIndex),
                            new XAttribute("data-hf-type", "footer-first"),
                            firstFooterContent));
                    }
                }

                // Render even footer if exists
                var evenFooterContent = RenderFooterForSection(wordDoc, settings, sectPr, "even");
                if (evenFooterContent != null)
                {
                    registry.Add(new XElement(Xhtml.div,
                        new XAttribute("data-section", sectionIndex),
                        new XAttribute("data-hf-type", "footer-even"),
                        evenFooterContent));
                }
            }

            // Only return registry if it has content
            return registry.HasElements ? registry : null;
        }

        /// <summary>
        /// Renders a header of a specific type for a section.
        /// </summary>
        private static object RenderHeaderForSection(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement sectPr,
            string headerType)
        {
            var headerRef = sectPr?.Elements(W.headerReference)
                .FirstOrDefault(hr => (string)hr.Attribute(W.type) == headerType);

            if (headerRef == null)
                return null;

            var headerId = (string)headerRef.Attribute(R.id);
            if (headerId == null)
                return null;

            var headerPart = wordDoc.MainDocumentPart.GetPartById(headerId) as HeaderPart;
            if (headerPart == null)
                return null;

            var headerXDoc = headerPart.GetXDocument();
            var headerRoot = headerXDoc.Root;

            if (headerRoot == null || !headerRoot.Elements().Any())
                return null;

            // Convert the content of the header
            return headerRoot.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .Where(c => c != null)
                .ToList();
        }

        /// <summary>
        /// Renders a footer of a specific type for a section.
        /// </summary>
        private static object RenderFooterForSection(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement sectPr,
            string footerType)
        {
            var footerRef = sectPr?.Elements(W.footerReference)
                .FirstOrDefault(fr => (string)fr.Attribute(W.type) == footerType);

            if (footerRef == null)
                return null;

            var footerId = (string)footerRef.Attribute(R.id);
            if (footerId == null)
                return null;

            var footerPart = wordDoc.MainDocumentPart.GetPartById(footerId) as FooterPart;
            if (footerPart == null)
                return null;

            var footerXDoc = footerPart.GetXDocument();
            var footerRoot = footerXDoc.Root;

            if (footerRoot == null || !footerRoot.Elements().Any())
                return null;

            // Convert the content of the footer
            return footerRoot.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .Where(c => c != null)
                .ToList();
        }

        private static void LoadComments(WordprocessingDocument wordDoc, CommentTracker tracker)
        {
            var commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;
            if (commentsPart == null)
                return;

            var commentsXDoc = commentsPart.GetXDocument();
            if (commentsXDoc.Root == null)
                return;

            foreach (var comment in commentsXDoc.Root.Elements(W.comment))
            {
                var id = (int?)comment.Attribute(W.id);
                if (id == null)
                    continue;

                tracker.Comments[id.Value] = new CommentInfo
                {
                    Id = id.Value,
                    Author = (string)comment.Attribute(W.author),
                    Date = (string)comment.Attribute(W.date),
                    Initials = (string)comment.Attribute(W.initials),
                    ContentParagraphs = comment.Elements(W.p).ToList()
                };
            }
        }

        private static CommentTracker GetCommentTracker(XElement element)
        {
            // Walk up to the document root to find the annotation
            var root = element.AncestorsAndSelf().LastOrDefault();
            return root?.Annotation<CommentTracker>();
        }

        private static void LoadAnnotations(WordprocessingDocument wordDoc, AnnotationTracker tracker)
        {
            // Look for the custom XML part with our namespace
            var customXmlParts = wordDoc.MainDocumentPart.CustomXmlParts;
            if (customXmlParts == null)
                return;

            foreach (var customXmlPart in customXmlParts)
            {
                try
                {
                    var xDoc = XDocument.Load(customXmlPart.GetStream());
                    if (xDoc.Root?.Name.LocalName == "annotations" &&
                        xDoc.Root.Name.NamespaceName == AnnotationManager.AnnotationsNamespace)
                    {
                        // Found our annotations - parse them
                        var ns = XNamespace.Get(AnnotationManager.AnnotationsNamespace);
                        foreach (var annotElement in xDoc.Root.Elements(ns + "annotation"))
                        {
                            // Annotation data is stored as attributes, not elements
                            var annotation = new DocumentAnnotation
                            {
                                Id = (string)annotElement.Attribute("id"),
                                LabelId = (string)annotElement.Attribute("labelId"),
                                Label = (string)annotElement.Attribute("label"),
                                Color = (string)annotElement.Attribute("color"),
                                Author = (string)annotElement.Attribute("author")
                            };

                            // BookmarkName is in a child range element
                            var rangeElement = annotElement.Element(ns + "range");
                            if (rangeElement != null)
                            {
                                annotation.BookmarkName = (string)rangeElement.Attribute("bookmarkName");
                            }

                            var createdStr = (string)annotElement.Attribute("created");
                            if (!string.IsNullOrEmpty(createdStr) && DateTime.TryParse(createdStr, out var created))
                                annotation.Created = created;

                            // Page span is in a child element
                            var pageSpanElement = annotElement.Element(ns + "pageSpan");
                            if (pageSpanElement != null)
                            {
                                var startPageStr = (string)pageSpanElement.Attribute("startPage");
                                if (!string.IsNullOrEmpty(startPageStr) && int.TryParse(startPageStr, out var startPage))
                                    annotation.StartPage = startPage;

                                var endPageStr = (string)pageSpanElement.Attribute("endPage");
                                if (!string.IsNullOrEmpty(endPageStr) && int.TryParse(endPageStr, out var endPage))
                                    annotation.EndPage = endPage;
                            }

                            // Load metadata
                            var metadataElement = annotElement.Element(ns + "metadata");
                            if (metadataElement != null)
                            {
                                foreach (var item in metadataElement.Elements(ns + "item"))
                                {
                                    var key = (string)item.Attribute("key");
                                    var value = item.Value;
                                    if (!string.IsNullOrEmpty(key))
                                        annotation.Metadata[key] = value;
                                }
                            }

                            if (!string.IsNullOrEmpty(annotation.Id))
                            {
                                tracker.Annotations[annotation.Id] = annotation;

                                // Map bookmark name to annotation ID
                                if (!string.IsNullOrEmpty(annotation.BookmarkName))
                                {
                                    tracker.BookmarkToAnnotationId[annotation.BookmarkName] = annotation.Id;
                                }
                            }
                        }
                        break; // Found our part, no need to continue
                    }
                }
                catch
                {
                    // Skip invalid XML parts
                }
            }
        }

        private static AnnotationTracker GetAnnotationTracker(XElement element)
        {
            // Walk up to the document root to find the tracker
            var root = element.AncestorsAndSelf().LastOrDefault();
            return root?.Annotation<AnnotationTracker>();
        }

        private static object ProcessCommentRangeStart(WmlToHtmlConverterSettings settings, XElement element)
        {
            if (!settings.RenderComments)
                return null;

            var tracker = GetCommentTracker(element);
            if (tracker == null)
                return null;

            var id = (int?)element.Attribute(W.id);
            if (id != null)
            {
                tracker.OpenRanges.Add(id.Value);
            }

            // Don't emit anything - we'll wrap content in ConvertRun
            return null;
        }

        private static object ProcessCommentRangeEnd(WmlToHtmlConverterSettings settings, XElement element)
        {
            if (!settings.RenderComments)
                return null;

            var tracker = GetCommentTracker(element);
            if (tracker == null)
                return null;

            var id = (int?)element.Attribute(W.id);
            if (id != null)
            {
                tracker.OpenRanges.Remove(id.Value);
            }

            return null;
        }

        private static object ProcessCommentReference(WmlToHtmlConverterSettings settings, XElement element)
        {
            if (!settings.RenderComments)
                return null;

            var tracker = GetCommentTracker(element);
            if (tracker == null)
                return null;

            var id = (int?)element.Attribute(W.id);
            if (id == null)
                return null;

            // Track that this comment was referenced (for rendering order)
            if (!tracker.ReferencedCommentIds.Contains(id.Value))
            {
                tracker.ReferencedCommentIds.Add(id.Value);
            }

            var prefix = settings.CommentCssClassPrefix ?? "comment-";
            tracker.Comments.TryGetValue(id.Value, out var comment);

            var marker = new XElement(Xhtml.a,
                new XAttribute("href", $"#comment-{id}"),
                new XAttribute("id", $"comment-ref-{id}"),
                new XAttribute("class", prefix + "marker"));

            if (comment != null && settings.IncludeCommentMetadata && comment.Author != null)
            {
                marker.Add(new XAttribute("title", $"Comment by {comment.Author}"));
            }

            marker.Add(new XText($"[{id}]"));

            var style = new Dictionary<string, string>
            {
                { "vertical-align", "super" },
                { "font-size", "smaller" }
            };
            marker.AddAnnotation(style);

            return marker;
        }

        private static XElement RenderCommentsSection(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, CommentTracker tracker)
        {
            if (!settings.RenderComments || !tracker.Comments.Any())
                return null;

            var prefix = settings.CommentCssClassPrefix ?? "comment-";

            // Use referenced order if available, otherwise use comment ID order
            var orderedComments = tracker.ReferencedCommentIds.Any()
                ? tracker.ReferencedCommentIds
                    .Where(id => tracker.Comments.ContainsKey(id))
                    .Select(id => tracker.Comments[id])
                : tracker.Comments.Values.OrderBy(c => c.Id);

            var commentItems = orderedComments
                .Select(c => RenderCommentItem(wordDoc, settings, c, prefix))
                .Where(item => item != null)
                .ToList();

            if (!commentItems.Any())
                return null;

            return new XElement(Xhtml.aside,
                new XAttribute("class", prefix.TrimEnd('-') + "s-section"),
                new XElement(Xhtml.h2, "Comments"),
                new XElement(Xhtml.ol,
                    new XAttribute("class", prefix.TrimEnd('-') + "s-list"),
                    commentItems));
        }

        private static XElement RenderCommentItem(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, CommentInfo comment, string prefix)
        {
            var li = new XElement(Xhtml.li,
                new XAttribute("id", $"comment-{comment.Id}"),
                new XAttribute("class", prefix.TrimEnd('-')));

            if (settings.IncludeCommentMetadata)
            {
                if (comment.Author != null)
                    li.Add(new XAttribute("data-author", comment.Author));
                if (comment.Date != null)
                    li.Add(new XAttribute("data-date", comment.Date));
            }

            // Header with author, date, and back link
            var header = new XElement(Xhtml.div,
                new XAttribute("class", prefix + "header"));

            if (comment.Author != null)
            {
                header.Add(new XElement(Xhtml.span,
                    new XAttribute("class", prefix + "author"),
                    comment.Author));
            }

            if (comment.Date != null)
            {
                // Format date nicely
                if (DateTime.TryParse(comment.Date, out var dt))
                {
                    header.Add(new XElement(Xhtml.span,
                        new XAttribute("class", prefix + "date"),
                        dt.ToString("MMM d, yyyy")));
                }
            }

            header.Add(new XElement(Xhtml.a,
                new XAttribute("href", $"#comment-ref-{comment.Id}"),
                new XAttribute("class", prefix + "backref"),
                "↩"));

            li.Add(header);

            // Comment body - extract text content from paragraphs
            var body = new XElement(Xhtml.div,
                new XAttribute("class", prefix + "body"));

            foreach (var para in comment.ContentParagraphs)
            {
                // Skip the annotation reference run and get text content
                var textContent = para.Descendants(W.t)
                    .Where(t => !t.Ancestors(W.r).Any(r => r.Elements(W.annotationRef).Any()))
                    .Select(t => t.Value)
                    .StringConcatenate();

                if (!string.IsNullOrWhiteSpace(textContent))
                {
                    body.Add(new XElement(Xhtml.p, textContent));
                }
            }

            // Only add if body has content
            if (body.HasElements)
            {
                li.Add(body);
            }
            else
            {
                // Add empty paragraph to avoid empty li
                li.Add(new XElement(Xhtml.div,
                    new XAttribute("class", prefix + "body"),
                    new XElement(Xhtml.p, "(empty comment)")));
            }

            return li;
        }

        private static XElement RenderMarginCommentsColumn(WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings, CommentTracker tracker, string prefix)
        {
            var marginColumn = new XElement(Xhtml.aside,
                new XAttribute("class", prefix + "margin-column"));

            if (tracker == null || !tracker.Comments.Any())
                return marginColumn;

            // Use referenced order if available, otherwise use comment ID order
            var orderedComments = tracker.ReferencedCommentIds.Any()
                ? tracker.ReferencedCommentIds
                    .Where(id => tracker.Comments.ContainsKey(id))
                    .Select(id => tracker.Comments[id])
                : tracker.Comments.Values.OrderBy(c => c.Id);

            foreach (var comment in orderedComments)
            {
                var marginNote = RenderMarginCommentNote(settings, comment, prefix);
                if (marginNote != null)
                    marginColumn.Add(marginNote);
            }

            return marginColumn;
        }

        private static XElement RenderMarginCommentNote(WmlToHtmlConverterSettings settings,
            CommentInfo comment, string prefix)
        {
            var note = new XElement(Xhtml.div,
                new XAttribute("id", $"comment-{comment.Id}"),
                new XAttribute("class", prefix + "margin-note"),
                new XAttribute("data-comment-id", comment.Id.ToString()));

            if (settings.IncludeCommentMetadata)
            {
                if (comment.Author != null)
                    note.Add(new XAttribute("data-author", comment.Author));
                if (comment.Date != null)
                    note.Add(new XAttribute("data-date", comment.Date));
            }

            // Header with author, date, and back link
            var header = new XElement(Xhtml.div,
                new XAttribute("class", prefix + "margin-note-header"));

            if (comment.Author != null)
            {
                header.Add(new XElement(Xhtml.span,
                    new XAttribute("class", prefix + "margin-author"),
                    comment.Author));
            }

            if (comment.Date != null)
            {
                // Format date nicely
                if (DateTime.TryParse(comment.Date, out var dt))
                {
                    header.Add(new XElement(Xhtml.span,
                        new XAttribute("class", prefix + "margin-date"),
                        dt.ToString("MMM d")));
                }
            }

            header.Add(new XElement(Xhtml.a,
                new XAttribute("href", $"#comment-ref-{comment.Id}"),
                new XAttribute("class", prefix + "margin-backref"),
                "↩"));

            note.Add(header);

            // Comment body - extract text content from paragraphs
            var body = new XElement(Xhtml.div,
                new XAttribute("class", prefix + "margin-note-body"));

            foreach (var para in comment.ContentParagraphs)
            {
                // Skip the annotation reference run and get text content
                var textContent = para.Descendants(W.t)
                    .Where(t => !t.Ancestors(W.r).Any(r => r.Elements(W.annotationRef).Any()))
                    .Select(t => t.Value)
                    .StringConcatenate();

                if (!string.IsNullOrWhiteSpace(textContent))
                {
                    body.Add(new XElement(Xhtml.p, textContent));
                }
            }

            // Only add if body has content
            if (body.HasElements)
            {
                note.Add(body);
            }
            else
            {
                // Add empty paragraph to avoid empty note
                note.Add(new XElement(Xhtml.div,
                    new XAttribute("class", prefix + "margin-note-body"),
                    new XElement(Xhtml.p, "(empty comment)")));
            }

            return note;
        }

        private static object ProcessTab(XElement element)
        {
            var tabWidthAtt = element.Attribute(PtOpenXml.TabWidth);
            if (tabWidthAtt == null) return null;

            var leader = (string) element.Attribute(PtOpenXml.Leader);
            var tabWidth = (decimal) tabWidthAtt;
            var style = new Dictionary<string, string>();
            XElement span;
            if (leader != null)
            {
                var leaderChar = ".";
                if (leader == "hyphen")
                    leaderChar = "-";
                else if (leader == "dot")
                    leaderChar = ".";
                else if (leader == "underscore")
                    leaderChar = "_";

                var runContainingTabToReplace = element.Ancestors(W.r).First();
                var fontNameAtt = runContainingTabToReplace.Attribute(PtOpenXml.pt + "FontName") ??
                                  runContainingTabToReplace.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");

                var dummyRun = new XElement(W.r,
                    fontNameAtt,
                    runContainingTabToReplace.Elements(W.rPr),
                    new XElement(W.t, leaderChar));

                var widthOfLeaderChar = CalcWidthOfRunInTwips(dummyRun);

                bool forceArial = false;
                if (widthOfLeaderChar == 0)
                {
                    dummyRun = new XElement(W.r,
                        new XAttribute(PtOpenXml.FontName, "Arial"),
                        runContainingTabToReplace.Elements(W.rPr),
                        new XElement(W.t, leaderChar));
                    widthOfLeaderChar = CalcWidthOfRunInTwips(dummyRun);
                    forceArial = true;
                }

                if (widthOfLeaderChar != 0)
                {
                    var numberOfLeaderChars = (int) (Math.Floor((tabWidth*1440)/widthOfLeaderChar));
                    if (numberOfLeaderChars < 0)
                        numberOfLeaderChars = 0;
                    span = new XElement(Xhtml.span,
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        " " + "".PadRight(numberOfLeaderChars, leaderChar[0]) + " ");
                    style.Add("margin", "0 0 0 0");
                    style.Add("padding", "0 0 0 0");
                    style.Add("width", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", tabWidth));
                    style.Add("text-align", "center");
                    if (forceArial)
                        style.Add("font-family", "Arial");
                }
                else
                {
                    span = new XElement(Xhtml.span, new XAttribute(XNamespace.Xml + "space", "preserve"), " ");
                    style.Add("margin", "0 0 0 0");
                    style.Add("padding", "0 0 0 0");
                    style.Add("width", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", tabWidth));
                    style.Add("text-align", "center");
                    if (leader == "underscore")
                    {
                        style.Add("text-decoration", "underline");
                    }
                }
            }
            else
            {
#if false
                            var bidi = element
                                .Ancestors(W.p)
                                .Take(1)
                                .Elements(W.pPr)
                                .Elements(W.bidi)
                                .Where(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true)
                                .FirstOrDefault();
                            var isBidi = bidi != null;
                            if (isBidi)
                                span = new XElement(Xhtml.span, new XEntity("#x200f")); // RLM
                            else
                                span = new XElement(Xhtml.span, new XEntity("#x200e")); // LRM
#else
                span = new XElement(Xhtml.span, new XEntity("#x00a0"));
#endif
                style.Add("margin", string.Format(NumberFormatInfo.InvariantInfo, "0 0 0 {0:0.00}in", tabWidth));
                style.Add("padding", "0 0 0 0");
            }
            span.AddAnnotation(style);
            return span;
        }

        private static object ProcessBreak(XElement element, WmlToHtmlConverterSettings settings)
        {
            // Check for page break (w:br with w:type="page")
            var breakType = (string)element.Attribute(W.type);

            // In pagination mode, emit a page break marker div for page breaks
            if (settings.RenderPagination == PaginationMode.Paginated && breakType == "page")
            {
                var prefix = settings.PaginationCssClassPrefix ?? "page-";
                return new XElement(Xhtml.div,
                    new XAttribute("class", prefix + "break"),
                    new XAttribute("data-page-break", "true"));
            }

            // Column breaks - also mark for pagination but render as line break normally
            if (settings.RenderPagination == PaginationMode.Paginated && breakType == "column")
            {
                var prefix = settings.PaginationCssClassPrefix ?? "page-";
                return new XElement(Xhtml.div,
                    new XAttribute("class", prefix + "column-break"),
                    new XAttribute("data-column-break", "true"));
            }

            XElement span = null;
            var tabWidth = (decimal?) element.Attribute(PtOpenXml.TabWidth);
            if (tabWidth != null)
            {
                span = new XElement(Xhtml.span);
                span.AddAnnotation(new Dictionary<string, string>
                {
                    { "margin", string.Format(NumberFormatInfo.InvariantInfo, "0 0 0 {0:0.00}in", tabWidth) },
                    { "padding", "0 0 0 0" }
                });
            }

            var paragraph = element.Ancestors(W.p).FirstOrDefault();
            var isBidi = paragraph != null &&
                         paragraph.Elements(W.pPr).Elements(W.bidi).Any(b => b.Attribute(W.val) == null ||
                                                                             b.Attribute(W.val).ToBoolean() == true);
            var zeroWidthChar = isBidi ? new XEntity("#x200f") : new XEntity("#x200e");

            return new object[]
            {
                new XElement(Xhtml.br),
                zeroWidthChar,
                span,
            };
        }

        private static object ProcessContentControl(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            XElement element, decimal currentMarginLeft)
        {
            var relevantAncestors = element.Ancestors().TakeWhile(a => a.Name != W.txbxContent);
            var isRunLevelContentControl = relevantAncestors.Any(a => a.Name == W.p);
            if (isRunLevelContentControl)
            {
                return element.Elements(W.sdtContent).Elements()
                    .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                    .ToList();
            }
            return CreateBorderDivs(wordDoc, settings, element.Elements(W.sdtContent).Elements());
        }

        // Transform the w:p element, including the following sibling w:p element(s)
        // in case the w:p element has a style separator. The sibling(s) will be
        // transformed to h:span elements rather than h:p elements and added to
        // the element (e.g., h:h2) created from the w:p element having the (first)
        // style separator (i.e., a w:specVanish element).
        private static object ProcessParagraph(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            XElement element, bool suppressTrailingWhiteSpace, decimal currentMarginLeft)
        {
            // Ignore this paragraph if the previous paragraph has a style separator.
            // We have already transformed this one together with the previous one.
            var previousParagraph = element.ElementsBeforeSelf(W.p).LastOrDefault();
            if (HasStyleSeparator(previousParagraph)) return null;

            var elementName = GetParagraphElementName(element, wordDoc);
            var isBidi = IsBidi(element);
            var paragraph = (XElement) ConvertParagraph(wordDoc, settings, element, elementName,
                suppressTrailingWhiteSpace, currentMarginLeft, isBidi);

            // The paragraph conversion might have created empty spans.
            // These can and should be removed because empty spans are
            // invalid in HTML5.
            paragraph.Elements(Xhtml.span).Where(e => e.IsEmpty).Remove();

            foreach (var span in paragraph.Elements(Xhtml.span).ToList())
            {
                var v = span.Value;
                if (v.Length > 0 && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1])) && span.Attribute(XNamespace.Xml + "space") == null)
                    span.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
            }

            while (HasStyleSeparator(element))
            {
                element = element.ElementsAfterSelf(W.p).FirstOrDefault();
                if (element == null) break;

                elementName = Xhtml.span;
                isBidi = IsBidi(element);
                var span = (XElement)ConvertParagraph(wordDoc, settings, element, elementName,
                    suppressTrailingWhiteSpace, currentMarginLeft, isBidi);
                var v = span.Value;
                if (v.Length > 0 && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1])) && span.Attribute(XNamespace.Xml + "space") == null)
                    span.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                paragraph.Add(span);
            }

            // Handle paragraph mark revisions (when paragraph was inserted/deleted)
            if (settings.RenderTrackedChanges)
            {
                var pPr = element.Element(W.pPr);
                var rPr = pPr?.Element(W.rPr);
                var paraIns = rPr?.Element(W.ins);
                var paraDel = rPr?.Element(W.del);

                if (paraIns != null)
                {
                    // Paragraph mark was inserted (paragraph was split from another)
                    var existingClass = (string)paragraph.Attribute("class");
                    var newClass = (settings.RevisionCssClassPrefix ?? "rev-") + "para-ins";
                    paragraph.SetAttributeValue("class", existingClass != null ? existingClass + " " + newClass : newClass);

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)paraIns.Attribute(W.author);
                        var date = (string)paraIns.Attribute(W.date);
                        if (author != null && paragraph.Attribute("data-author") == null)
                            paragraph.Add(new XAttribute("data-author", author));
                        if (date != null && paragraph.Attribute("data-date") == null)
                            paragraph.Add(new XAttribute("data-date", date));
                    }
                }
                else if (paraDel != null)
                {
                    // Paragraph mark was deleted (paragraph was merged with next)
                    var existingClass = (string)paragraph.Attribute("class");
                    var newClass = (settings.RevisionCssClassPrefix ?? "rev-") + "para-del";
                    paragraph.SetAttributeValue("class", existingClass != null ? existingClass + " " + newClass : newClass);

                    // Add a pilcrow marker at the end to show the deleted paragraph mark
                    var prefix = settings.RevisionCssClassPrefix ?? "rev-";
                    paragraph.Add(new XElement(Xhtml.span,
                        new XAttribute("class", prefix + "para-mark-del"),
                        new XAttribute("title", "Paragraph mark deleted"),
                        new XText("¶")));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)paraDel.Attribute(W.author);
                        var date = (string)paraDel.Attribute(W.date);
                        if (author != null && paragraph.Attribute("data-author") == null)
                            paragraph.Add(new XAttribute("data-author", author));
                        if (date != null && paragraph.Attribute("data-date") == null)
                            paragraph.Add(new XAttribute("data-date", date));
                    }
                }
            }

            // Add pagination-related data attributes when pagination is enabled
            if (settings.RenderPagination == PaginationMode.Paginated)
            {
                AddPaginationDataAttributes(element, paragraph);
            }

            return paragraph;
        }

        /// <summary>
        /// Adds pagination-related data attributes to the HTML paragraph element.
        /// These attributes help the client-side pagination engine make better decisions.
        /// </summary>
        private static void AddPaginationDataAttributes(XElement wordParagraph, XElement htmlParagraph)
        {
            var pPr = wordParagraph.Element(W.pPr);
            if (pPr == null) return;

            // w:keepNext - keep this paragraph with the next one on the same page
            var keepNext = pPr.Element(W.keepNext);
            if (keepNext != null && ((string)keepNext.Attribute(W.val) == null || keepNext.Attribute(W.val).ToBoolean() == true))
            {
                htmlParagraph.Add(new XAttribute("data-keep-with-next", "true"));
            }

            // w:keepLines - keep all lines of this paragraph together on one page
            var keepLines = pPr.Element(W.keepLines);
            if (keepLines != null && ((string)keepLines.Attribute(W.val) == null || keepLines.Attribute(W.val).ToBoolean() == true))
            {
                htmlParagraph.Add(new XAttribute("data-keep-lines", "true"));
            }

            // w:pageBreakBefore - force a page break before this paragraph
            var pageBreakBefore = pPr.Element(W.pageBreakBefore);
            if (pageBreakBefore != null && ((string)pageBreakBefore.Attribute(W.val) == null || pageBreakBefore.Attribute(W.val).ToBoolean() == true))
            {
                htmlParagraph.Add(new XAttribute("data-page-break-before", "true"));
            }

            // w:widowControl - control widow/orphan lines
            var widowControl = pPr.Element(W.widowControl);
            if (widowControl != null)
            {
                var val = widowControl.Attribute(W.val);
                // If val is null or true, widow control is enabled
                bool isEnabled = val == null || val.ToBoolean() == true;
                htmlParagraph.Add(new XAttribute("data-widow-control", isEnabled ? "true" : "false"));
            }
        }

        private static object ProcessTable(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
        {
            var style = new Dictionary<string, string>();
            style.AddIfMissing("border-collapse", "collapse");
            style.AddIfMissing("border", "none");
            var bidiVisual = element.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
            var tblW = element.Elements(W.tblPr).Elements(W.tblW).FirstOrDefault();
            if (tblW != null)
            {
                var type = (string)tblW.Attribute(W.type);
                if (type != null && type == "pct")
                {
                    var w = (int)tblW.Attribute(W._w);
                    style.AddIfMissing("width", (w / 50) + "%");
                }
            }
            var tblInd = element.Elements(W.tblPr).Elements(W.tblInd).FirstOrDefault();
            if (tblInd != null)
            {
                var tblIndType = (string)tblInd.Attribute(W.type);
                if (tblIndType != null)
                {
                    if (tblIndType == "dxa")
                    {
                        var width = (decimal?)tblInd.Attribute(W._w);
                        if (width != null)
                        {
                            style.AddIfMissing("margin-left",
                                width > 0m
                                    ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", width / 20m)
                                    : "0");
                        }
                    }
                }
            }
            var tableDirection = bidiVisual != null ? new XAttribute("dir", "rtl") : new XAttribute("dir", "ltr");
            style.AddIfMissing("margin-bottom", ".001pt");
            var table = new XElement(Xhtml.table,
                // TODO: Revisit and make sure the omission is covered by appropriate CSS.
                // new XAttribute("border", "1"),
                // new XAttribute("cellspacing", 0),
                // new XAttribute("cellpadding", 0),
                tableDirection,
                element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));
            table.AddAnnotation(style);
            var jc = (string)element.Elements(W.tblPr).Elements(W.jc).Attributes(W.val).FirstOrDefault() ?? "left";
            XAttribute dir = null;
            XAttribute jcToUse = null;
            if (bidiVisual != null)
            {
                dir = new XAttribute("dir", "rtl");
                if (jc == "left")
                    jcToUse = new XAttribute("align", "right");
                else if (jc == "right")
                    jcToUse = new XAttribute("align", "left");
                else if (jc == "center")
                    jcToUse = new XAttribute("align", "center");
            }
            else
            {
                jcToUse = new XAttribute("align", jc);
            }
            var tableDiv = new XElement(Xhtml.div,
                dir,
                jcToUse,
                table);
            return tableDiv;
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        private static object ProcessTableCell(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            var style = new Dictionary<string, string>();
            XAttribute colSpan = null;
            XAttribute rowSpan = null;

            var tcPr = element.Element(W.tcPr);
            if (tcPr != null)
            {
                if ((string) tcPr.Elements(W.vMerge).Attributes(W.val).FirstOrDefault() == "restart")
                {
                    var currentRow = element.Parent.ElementsBeforeSelf(W.tr).Count();
                    var currentCell = element.ElementsBeforeSelf(W.tc).Count();
                    var tbl = element.Parent.Parent;
                    int rowSpanCount = 1;
                    currentRow += 1;
                    while (true)
                    {
                        var row = tbl.Elements(W.tr).Skip(currentRow).FirstOrDefault();
                        if (row == null)
                            break;
                        var cell2 = row.Elements(W.tc).Skip(currentCell).FirstOrDefault();
                        if (cell2 == null)
                            break;
                        if (cell2.Elements(W.tcPr).Elements(W.vMerge).FirstOrDefault() == null)
                            break;
                        if ((string) cell2.Elements(W.tcPr).Elements(W.vMerge).Attributes(W.val).FirstOrDefault() == "restart")
                            break;
                        currentRow += 1;
                        rowSpanCount += 1;
                    }
                    rowSpan = new XAttribute("rowspan", rowSpanCount);
                }

                if (tcPr.Element(W.vMerge) != null &&
                    (string) tcPr.Elements(W.vMerge).Attributes(W.val).FirstOrDefault() != "restart")
                    return null;

                if (tcPr.Element(W.vAlign) != null)
                {
                    var vAlignVal = (string) tcPr.Elements(W.vAlign).Attributes(W.val).FirstOrDefault();
                    if (vAlignVal == "top")
                        style.AddIfMissing("vertical-align", "top");
                    else if (vAlignVal == "center")
                        style.AddIfMissing("vertical-align", "middle");
                    else if (vAlignVal == "bottom")
                        style.AddIfMissing("vertical-align", "bottom");
                    else
                        style.AddIfMissing("vertical-align", "middle");
                }
                style.AddIfMissing("vertical-align", "top");

                if ((string) tcPr.Elements(W.tcW).Attributes(W.type).FirstOrDefault() == "dxa")
                {
                    decimal width = (int) tcPr.Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    style.AddIfMissing("width", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", width/20m));
                }
                if ((string) tcPr.Elements(W.tcW).Attributes(W.type).FirstOrDefault() == "pct")
                {
                    decimal width = (int) tcPr.Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    style.AddIfMissing("width", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}%", width/50m));
                }

                var tcBorders = tcPr.Element(W.tcBorders);
                GenerateBorderStyle(tcBorders, W.top, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.right, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.bottom, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.left, style, BorderType.Cell);

                CreateStyleFromShd(style, tcPr.Element(W.shd));

                var gridSpan = tcPr.Elements(W.gridSpan).Attributes(W.val).Select(a => (int?) a).FirstOrDefault();
                if (gridSpan != null)
                    colSpan = new XAttribute("colspan", (int) gridSpan);
            }
            style.AddIfMissing("padding-top", "0");
            style.AddIfMissing("padding-bottom", "0");

            var cell = new XElement(Xhtml.td,
                rowSpan,
                colSpan,
                CreateBorderDivs(wordDoc, settings, element.Elements()));
            cell.AddAnnotation(style);

            // Handle table cell revisions
            if (settings.RenderTrackedChanges && tcPr != null)
            {
                var cellIns = tcPr.Element(W.cellIns);
                var cellDel = tcPr.Element(W.cellDel);
                var cellMerge = tcPr.Element(W.cellMerge);

                if (cellIns != null)
                {
                    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "cell-ins";
                    cell.Add(new XAttribute("class", className));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)cellIns.Attribute(W.author);
                        var date = (string)cellIns.Attribute(W.date);
                        if (author != null)
                            cell.Add(new XAttribute("data-author", author));
                        if (date != null)
                            cell.Add(new XAttribute("data-date", date));
                    }
                }
                else if (cellDel != null)
                {
                    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "cell-del";
                    cell.Add(new XAttribute("class", className));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)cellDel.Attribute(W.author);
                        var date = (string)cellDel.Attribute(W.date);
                        if (author != null)
                            cell.Add(new XAttribute("data-author", author));
                        if (date != null)
                            cell.Add(new XAttribute("data-date", date));
                    }
                }
                else if (cellMerge != null)
                {
                    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "cell-merge";
                    cell.Add(new XAttribute("class", className));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)cellMerge.Attribute(W.author);
                        var date = (string)cellMerge.Attribute(W.date);
                        if (author != null)
                            cell.Add(new XAttribute("data-author", author));
                        if (date != null)
                            cell.Add(new XAttribute("data-date", date));
                    }
                }
            }

            return cell;
        }

        private static object ProcessTableRow(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element,
            decimal currentMarginLeft)
        {
            var style = new Dictionary<string, string>();
            int? trHeight = (int?) element.Elements(W.trPr).Elements(W.trHeight).Attributes(W.val).FirstOrDefault();
            if (trHeight != null)
                style.AddIfMissing("height",
                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", (decimal) trHeight/1440m));
            var htmlRow = new XElement(Xhtml.tr,
                element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));
            if (style.Any())
                htmlRow.AddAnnotation(style);

            // Handle table row revision tracking
            if (settings.RenderTrackedChanges)
            {
                var trPr = element.Element(W.trPr);
                var rowIns = trPr?.Element(W.ins);
                var rowDel = trPr?.Element(W.del);

                if (rowIns != null)
                {
                    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "row-ins";
                    htmlRow.Add(new XAttribute("class", className));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)rowIns.Attribute(W.author);
                        var date = (string)rowIns.Attribute(W.date);
                        if (author != null)
                            htmlRow.Add(new XAttribute("data-author", author));
                        if (date != null)
                            htmlRow.Add(new XAttribute("data-date", date));
                    }
                }
                else if (rowDel != null)
                {
                    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "row-del";
                    htmlRow.Add(new XAttribute("class", className));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)rowDel.Attribute(W.author);
                        var date = (string)rowDel.Attribute(W.date);
                        if (author != null)
                            htmlRow.Add(new XAttribute("data-author", author));
                        if (date != null)
                            htmlRow.Add(new XAttribute("data-date", date));
                    }
                }
            }

            return htmlRow;
        }

        private static bool HasStyleSeparator(XElement element)
        {
            return element != null && element.Elements(W.pPr).Elements(W.rPr).Any(e => GetBoolProp(e, W.specVanish));
        }

        private static bool IsBidi(XElement element)
        {
            return element
                .Elements(W.pPr)
                .Elements(W.bidi)
                .Any(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);
        }

        private static XName GetParagraphElementName(XElement element, WordprocessingDocument wordDoc)
        {
            var elementName = Xhtml.p;

            var styleId = (string) element.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
            if (styleId == null) return elementName;

            var style = GetStyle(styleId, wordDoc);
            if (style == null) return elementName;

            var outlineLevel =
                (int?) style.Elements(W.pPr).Elements(W.outlineLvl).Attributes(W.val).FirstOrDefault();
            if (outlineLevel != null && outlineLevel <= 5)
            {
                elementName = Xhtml.xhtml + string.Format("h{0}", outlineLevel + 1);
            }

            return elementName;
        }

        private static XElement GetStyle(string styleId, WordprocessingDocument wordDoc)
        {
            var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart == null) return null;

            var styles = stylesPart.GetXDocument().Root;
            return styles != null
                ? styles.Elements(W.style).FirstOrDefault(s => (string) s.Attribute(W.styleId) == styleId)
                : null;
        }

        private static object CreateSectionDivs(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement element)
        {
            // Group elements by section. In pagination mode, we preserve section boundaries.
            // Without pagination, adjacent sections with identical formatting are conflated.
            var groupedIntoDivs = element
                .Elements()
                .GroupAdjacent(e => {
                    var sectAnnotation = e.Annotation<SectionAnnotation>();
                    return sectAnnotation != null ? sectAnnotation.SectionElement.ToString() : "";
                });

            int sectionIndex = 0;
            var divList = groupedIntoDivs
                .Select(g =>
                {
                    var sectAnnotation = g.First().Annotation<SectionAnnotation>();
                    XElement bidi = null;
                    PageDimensions dims = null;

                    if (sectAnnotation != null)
                    {
                        bidi = sectAnnotation
                            .SectionElement
                            .Elements(W.bidi)
                            .FirstOrDefault(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);

                        // Parse page dimensions for pagination mode
                        if (settings.RenderPagination == PaginationMode.Paginated)
                        {
                            dims = PageDimensions.FromSectionProperties(sectAnnotation.SectionElement);
                        }
                    }

                    var div = new XElement(Xhtml.div,
                        bidi != null ? new XAttribute("dir", "rtl") : null,
                        CreateBorderDivs(wordDoc, settings, g));

                    // Add pagination data attributes when enabled
                    if (settings.RenderPagination == PaginationMode.Paginated)
                    {
                        div.Add(new XAttribute("data-section-index", sectionIndex));

                        if (dims != null)
                        {
                            div.Add(new XAttribute("data-page-width", dims.PageWidthPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-page-height", dims.PageHeightPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-content-width", dims.ContentWidthPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-content-height", dims.ContentHeightPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-margin-top", dims.MarginTopPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-margin-right", dims.MarginRightPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-margin-bottom", dims.MarginBottomPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-margin-left", dims.MarginLeftPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-header-height", dims.HeaderPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                            div.Add(new XAttribute("data-footer-height", dims.FooterPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
                        }
                    }

                    sectionIndex++;
                    return div;
                })
                .ToList();

            // In pagination mode, wrap content in staging structure for client-side processing
            if (settings.RenderPagination == PaginationMode.Paginated)
            {
                var prefix = settings.PaginationCssClassPrefix ?? "page-";
                var stagingContent = new List<object>();

                // Add header/footer registry if headers/footers are enabled
                if (settings.RenderHeadersAndFooters)
                {
                    var hfRegistry = RenderPaginatedHeaderFooterRegistry(wordDoc, settings, element);
                    if (hfRegistry != null)
                        stagingContent.Add(hfRegistry);
                }

                // Add footnote registry if footnotes are enabled
                if (settings.RenderFootnotesAndEndnotes)
                {
                    var footnoteRegistry = RenderPaginatedFootnoteRegistry(wordDoc, settings);
                    if (footnoteRegistry != null)
                        stagingContent.Add(footnoteRegistry);
                }

                // Add section content
                stagingContent.AddRange(divList);

                return new object[]
                {
                    // Staging area containing the content (hidden by CSS for client-side measurement)
                    new XElement(Xhtml.div,
                        new XAttribute("id", "pagination-staging"),
                        new XAttribute("class", prefix + "staging"),
                        stagingContent),
                    // Container where paginated content will be rendered by client-side JavaScript
                    new XElement(Xhtml.div,
                        new XAttribute("id", "pagination-container"),
                        new XAttribute("class", prefix + "container"))
                };
            }

            return divList;
        }

        private enum BorderType
        {
            Paragraph,
            Cell,
        };

        /*
         * Notes on line spacing
         *
         * the w:line and w:lineRule attributes control spacing between lines - including between lines within a paragraph
         *
         * If w:spacing w:lineRule="auto" then
         *   w:spacing w:line is a percentage where 240 == 100%
         *
         *   (line value / 240) * 100 = percentage of line
         *
         * If w:spacing w:lineRule="exact" or w:lineRule="atLeast" then
         *   w:spacing w:line is in twips
         *   1440 = exactly one inch from line to line
         *
         * Handle
         * - ind
         * - jc
         * - numPr
         * - pBdr
         * - shd
         * - spacing
         * - textAlignment
         *
         * Don't Handle (yet)
         * - adjustRightInd?
         * - autoSpaceDE
         * - autoSpaceDN
         * - bidi
         * - contextualSpacing
         * - divId
         * - framePr
         * - keepLines
         * - keepNext
         * - kinsoku
         * - mirrorIndents
         * - overflowPunct
         * - pageBreakBefore
         * - snapToGrid
         * - suppressAutoHyphens
         * - suppressLineNumbers
         * - suppressOverlap
         * - tabs
         * - textBoxTightWrap
         * - textDirection
         * - topLinePunct
         * - widowControl
         * - wordWrap
         *
         */

        private static object ConvertParagraph(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            XElement paragraph, XName elementName, bool suppressTrailingWhiteSpace, decimal currentMarginLeft, bool isBidi)
        {
            var style = DefineParagraphStyle(paragraph, elementName, suppressTrailingWhiteSpace, currentMarginLeft, isBidi);
            var rtl = isBidi ? new XAttribute("dir", "rtl") : new XAttribute("dir", "ltr");
            var firstMark = isBidi ? new XEntity("#x200f") : null;

            // Analyze initial runs to see whether we have a tab, in which case we will render
            // a span with a defined width and ignore the tab rather than rendering the text
            // preceding the tab and the tab as a span with a computed width.
            var firstTabRun = paragraph
                .Elements(W.r)
                .FirstOrDefault(run => run.Elements(W.tab).Any());
            var elementsPrecedingTab = firstTabRun != null
                ? paragraph.Elements(W.r).TakeWhile(e => e != firstTabRun)
                    .Where(e => e.Elements().Any(c => c.Attributes(PtOpenXml.TabWidth).Any())).ToList()
                : Enumerable.Empty<XElement>().ToList();

            // TODO: Revisit
            // For the time being, if a hyperlink field precedes the tab, we'll render it as before.
            var hyperlinkPrecedesTab = elementsPrecedingTab
                .Elements(W.r)
                .Elements(W.instrText)
                .Select(e => e.Value)
                .Any(value => value != null && value.TrimStart().ToUpper().StartsWith("HYPERLINK"));
            if (hyperlinkPrecedesTab)
            {
                var paraElement1 = new XElement(elementName,
                    rtl,
                    firstMark,
                    ConvertContentThatCanContainFields(wordDoc, settings, paragraph.Elements()));
                paraElement1.AddAnnotation(style);
                return paraElement1;
            }

            var txElementsPrecedingTab = TransformElementsPrecedingTab(wordDoc, settings, elementsPrecedingTab, firstTabRun);
            var elementsSucceedingTab = firstTabRun != null
                ? paragraph.Elements().SkipWhile(e => e != firstTabRun).Skip(1)
                : paragraph.Elements();
            var paraElement = new XElement(elementName,
                rtl,
                firstMark,
                txElementsPrecedingTab,
                ConvertContentThatCanContainFields(wordDoc, settings, elementsSucceedingTab));
            paraElement.AddAnnotation(style);

            return paraElement;
        }

        private static List<object> TransformElementsPrecedingTab(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            List<XElement> elementsPrecedingTab, XElement firstTabRun)
        {
            var tabWidth = firstTabRun != null
                ? (decimal?) firstTabRun.Elements(W.tab).Attributes(PtOpenXml.TabWidth).FirstOrDefault() ?? 0m
                : 0m;
            var precedingElementsWidth = elementsPrecedingTab
                .Elements()
                .Where(c => c.Attributes(PtOpenXml.TabWidth).Any())
                .Select(e => (decimal) e.Attribute(PtOpenXml.TabWidth))
                .Sum();
            var totalWidth = precedingElementsWidth + tabWidth;

            var txElementsPrecedingTab = elementsPrecedingTab
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .ToList();
            if (txElementsPrecedingTab.Count > 1)
            {
                var span = new XElement(Xhtml.span, txElementsPrecedingTab);
                // Use min-width instead of width so the container expands to fit content
                // when text is wider than the calculated tab position. This fixes issues
                // where list numbers (e.g., "2.3") overlap with heading text because the
                // TabWidth for text elements is 0 (due to font measurement limitations).
                var spanStyle = new Dictionary<string, string>
                {
                    { "display", "inline-block" },
                    { "text-indent", "0" },
                    { "min-width", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}in", totalWidth) }
                };
                span.AddAnnotation(spanStyle);
            }
            else if (txElementsPrecedingTab.Count == 1)
            {
                var element = txElementsPrecedingTab.First() as XElement;
                if (element != null)
                {
                    var spanStyle = element.Annotation<Dictionary<string, string>>();
                    spanStyle.AddIfMissing("display", "inline-block");
                    spanStyle.AddIfMissing("text-indent", "0");
                    // Use min-width instead of width to allow content to expand naturally
                    spanStyle.AddIfMissing("min-width", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}in", totalWidth));
                }
            }
            return txElementsPrecedingTab;
        }

        private static Dictionary<string, string> DefineParagraphStyle(XElement paragraph, XName elementName,
            bool suppressTrailingWhiteSpace, decimal currentMarginLeft, bool isBidi)
        {
            var style = new Dictionary<string, string>();

            var styleName = (string) paragraph.Attribute(PtOpenXml.StyleName);
            if (styleName != null)
                style.Add("PtStyleName", styleName);

            var pPr = paragraph.Element(W.pPr);
            if (pPr == null) return style;

            CreateStyleFromSpacing(style, pPr.Element(W.spacing), elementName, suppressTrailingWhiteSpace);
            CreateStyleFromInd(style, pPr.Element(W.ind), elementName, currentMarginLeft, isBidi);

            // todo need to handle
            // - both
            // - mediumKashida
            // - distribute
            // - numTab
            // - highKashida
            // - lowKashida
            // - thaiDistribute

            CreateStyleFromJc(style, pPr.Element(W.jc), isBidi);
            CreateStyleFromShd(style, pPr.Element(W.shd));

            // Pt.FontName
            var font = (string) paragraph.Attributes(PtOpenXml.FontName).FirstOrDefault();
            if (font != null)
                CreateFontCssProperty(font, style);

            DefineFontSize(style, paragraph);
            DefineLineHeight(style, paragraph);

            // vertical text alignment as of December 2013 does not work in any major browsers.
            CreateStyleFromTextAlignment(style, pPr.Element(W.textAlignment));

            style.AddIfMissing("margin-top", "0");
            style.AddIfMissing("margin-left", "0");
            style.AddIfMissing("margin-right", "0");
            style.AddIfMissing("margin-bottom", ".001pt");

            return style;
        }

        private static void CreateStyleFromInd(Dictionary<string, string> style, XElement ind, XName elementName,
            decimal currentMarginLeft, bool isBidi)
        {
            if (ind == null) return;

            var left = (decimal?) ind.Attribute(W.left);
            if (left != null && elementName != Xhtml.span)
            {
                var leftInInches = (decimal) left/1440 - currentMarginLeft;
                style.AddIfMissing(isBidi ? "margin-right" : "margin-left",
                    leftInInches > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", leftInInches)
                        : "0");
            }

            var right = (decimal?) ind.Attribute(W.right);
            if (right != null)
            {
                var rightInInches = (decimal) right/1440;
                style.AddIfMissing(isBidi ? "margin-left" : "margin-right",
                    rightInInches > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", rightInInches)
                        : "0");
            }

            var firstLine = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.firstLine));
            if (firstLine != null && elementName != Xhtml.span)
            {
                var firstLineInInches = (decimal) firstLine/1440m;
                style.AddIfMissing("text-indent",
                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", firstLineInInches));
            }

            var hanging = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.hanging));
            if (hanging != null && elementName != Xhtml.span)
            {
                var hangingInInches = (decimal) -hanging/1440m;
                style.AddIfMissing("text-indent",
                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", hangingInInches));
            }
        }

        private static void CreateStyleFromJc(Dictionary<string, string> style, XElement jc, bool isBidi)
        {
            if (jc != null)
            {
                var jcVal = (string)jc.Attributes(W.val).FirstOrDefault() ?? "left";
                if (jcVal == "left")
                    style.AddIfMissing("text-align", isBidi ? "right" : "left");
                else if (jcVal == "right")
                    style.AddIfMissing("text-align", isBidi ? "left" : "right");
                else if (jcVal == "center")
                    style.AddIfMissing("text-align", "center");
                else if (jcVal == "both")
                    style.AddIfMissing("text-align", "justify");
            }
        }

        private static void CreateStyleFromSpacing(Dictionary<string, string> style, XElement spacing, XName elementName,
            bool suppressTrailingWhiteSpace)
        {
            if (spacing == null) return;

            var spacingBefore = (decimal?) spacing.Attribute(W.before);
            if (spacingBefore != null && elementName != Xhtml.span)
                style.AddIfMissing("margin-top",
                    spacingBefore > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", spacingBefore/20.0m)
                        : "0");

            var lineRule = (string) spacing.Attribute(W.lineRule);
            if (lineRule == "auto")
            {
                var line = (decimal) spacing.Attribute(W.line);
                if (line != 240m)
                {
                    var pct = (line/240m)*100m;
                    style.Add("line-height", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}%", pct));
                }
            }
            if (lineRule == "exact")
            {
                var line = (decimal) spacing.Attribute(W.line);
                var points = line/20m;
                style.Add("line-height", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", points));
            }
            if (lineRule == "atLeast")
            {
                var line = (decimal) spacing.Attribute(W.line);
                var points = line/20m;
                if (points >= 14m)
                    style.Add("line-height", string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", points));
            }

            var spacingAfter = suppressTrailingWhiteSpace ? 0 : WordprocessingMLUtil.AttributeToTwips(spacing.Attribute(W.after));
            if (spacingAfter != null)
                style.AddIfMissing("margin-bottom",
                    spacingAfter > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", spacingAfter/20.0m)
                        : "0");
        }

        private static void CreateStyleFromTextAlignment(Dictionary<string, string> style, XElement textAlignment)
        {
            if (textAlignment == null) return;

            var verticalTextAlignment = (string)textAlignment.Attributes(W.val).FirstOrDefault();
            if (verticalTextAlignment == null || verticalTextAlignment == "auto") return;

            if (verticalTextAlignment == "top")
                style.AddIfMissing("vertical-align", "top");
            else if (verticalTextAlignment == "center")
                style.AddIfMissing("vertical-align", "middle");
            else if (verticalTextAlignment == "baseline")
                style.AddIfMissing("vertical-align", "baseline");
            else if (verticalTextAlignment == "bottom")
                style.AddIfMissing("vertical-align", "bottom");
        }

        private static void DefineFontSize(Dictionary<string, string> style, XElement paragraph)
        {
            var sz = paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.r)
                .Select(r => GetFontSize(r))
                .Max();
            if (sz != null)
                style.AddIfMissing("font-size", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", sz / 2.0m));
        }

        private static void DefineLineHeight(Dictionary<string, string> style, XElement paragraph)
        {
            var allRunsAreUniDirectional = paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.r)
                .Select(run => (string)run.Attribute(PtOpenXml.LanguageType))
                .All(lt => lt != "bidi");
            if (allRunsAreUniDirectional)
                style.AddIfMissing("line-height", "108%");
        }

        /*
         * Handle:
         * - b
         * - bdr
         * - caps
         * - color
         * - dstrike
         * - highlight
         * - i
         * - position
         * - rFonts
         * - shd
         * - smallCaps
         * - spacing
         * - strike
         * - sz
         * - u
         * - vanish
         * - vertAlign
         *
         * Don't handle:
         * - em
         * - emboss
         * - fitText
         * - imprint
         * - kern
         * - outline
         * - shadow
         * - w
         *
         */

        private static object ConvertRun(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, XElement run)
        {
            var rPr = run.Element(W.rPr);
            if (rPr == null)
                return run.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m));

            // hide all content that contains the w:rPr/w:webHidden element
            if (rPr.Element(W.webHidden) != null)
                return null;

            var style = DefineRunStyle(run);
            object content = run.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m));

            // Wrap content in h:sup or h:sub elements as necessary.
            if (rPr.Element(W.vertAlign) != null)
            {
                XElement newContent = null;
                var vertAlignVal = (string)rPr.Elements(W.vertAlign).Attributes(W.val).FirstOrDefault();
                switch (vertAlignVal)
                {
                    case "superscript":
                        newContent = new XElement(Xhtml.sup, content);
                        break;
                    case "subscript":
                        newContent = new XElement(Xhtml.sub, content);
                        break;
                }
                if (newContent != null && newContent.Nodes().Any())
                    content = newContent;
            }

            var langAttribute = GetLangAttribute(run);

            XEntity runStartMark;
            XEntity runEndMark;
            DetermineRunMarks(run, rPr, style, out runStartMark, out runEndMark);

            if (style.Any() || langAttribute != null || runStartMark != null)
            {
                style.AddIfMissing("margin", "0");
                style.AddIfMissing("padding", "0");
                var xe = new XElement(Xhtml.span,
                    langAttribute,
                    runStartMark,
                    content,
                    runEndMark);

                xe.AddAnnotation(style);
                content = xe;
            }

            // Handle run property changes (formatting revisions)
            if (settings.RenderTrackedChanges && rPr != null)
            {
                var rPrChange = rPr.Element(W.rPrChange);
                if (rPrChange != null)
                {
                    var formatChangeSpan = new XElement(Xhtml.span);
                    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "format-change";
                    formatChangeSpan.Add(new XAttribute("class", className));

                    // Generate a title describing the change
                    var changeDescription = DescribeFormatChange(rPr, rPrChange);
                    if (!string.IsNullOrEmpty(changeDescription))
                        formatChangeSpan.Add(new XAttribute("title", changeDescription));

                    if (settings.IncludeRevisionMetadata)
                    {
                        var author = (string)rPrChange.Attribute(W.author);
                        var date = (string)rPrChange.Attribute(W.date);
                        if (author != null)
                            formatChangeSpan.Add(new XAttribute("data-author", author));
                        if (date != null)
                            formatChangeSpan.Add(new XAttribute("data-date", date));
                    }

                    formatChangeSpan.Add(content);
                    content = formatChangeSpan;
                }
            }

            // Wrap content in comment highlight spans if inside an open comment range
            if (settings.RenderComments)
            {
                var tracker = GetCommentTracker(run);
                if (tracker != null && tracker.OpenRanges.Any())
                {
                    var prefix = settings.CommentCssClassPrefix ?? "comment-";

                    // For each open comment range, wrap the content
                    foreach (var commentId in tracker.OpenRanges.OrderBy(id => id))
                    {
                        var highlightSpan = new XElement(Xhtml.span,
                            new XAttribute("class", prefix + "highlight"),
                            new XAttribute("data-comment-id", commentId.ToString()));

                        // Add tooltip in inline mode
                        if (settings.CommentRenderMode == CommentRenderMode.Inline)
                        {
                            if (tracker.Comments.TryGetValue(commentId, out var comment))
                            {
                                // Build inline tooltip
                                var tooltipText = comment.ContentParagraphs
                                    .SelectMany(p => p.Descendants(W.t)
                                        .Where(t => !t.Ancestors(W.r).Any(r => r.Elements(W.annotationRef).Any())))
                                    .Select(t => t.Value)
                                    .StringConcatenate();

                                if (!string.IsNullOrEmpty(tooltipText))
                                {
                                    var tooltip = comment.Author != null
                                        ? $"{comment.Author}: {tooltipText}"
                                        : tooltipText;
                                    highlightSpan.Add(new XAttribute("title", tooltip));
                                }

                                if (settings.IncludeCommentMetadata)
                                {
                                    highlightSpan.Add(new XAttribute("data-comment", tooltipText ?? ""));
                                    if (comment.Author != null)
                                        highlightSpan.Add(new XAttribute("data-author", comment.Author));
                                    if (comment.Date != null)
                                        highlightSpan.Add(new XAttribute("data-date", comment.Date));
                                }
                            }
                        }

                        highlightSpan.Add(content);
                        content = highlightSpan;
                    }
                }
            }

            // Wrap content in annotation highlight spans if inside an open annotation range
            if (settings.RenderAnnotations)
            {
                var annotTracker = GetAnnotationTracker(run);
                if (annotTracker != null && annotTracker.OpenRanges.Any())
                {
                    var prefix = settings.AnnotationCssClassPrefix ?? "annot-";

                    // For each open annotation range, wrap the content
                    foreach (var annotationId in annotTracker.OpenRanges.OrderBy(id => id))
                    {
                        if (annotTracker.Annotations.TryGetValue(annotationId, out var annotation))
                        {
                            var highlightSpan = new XElement(Xhtml.span,
                                new XAttribute("class", prefix + "highlight"),
                                new XAttribute("data-annotation-id", annotationId));

                            // Set the color as a CSS custom property
                            var inlineStyle = new Dictionary<string, string>();
                            if (!string.IsNullOrEmpty(annotation.Color))
                            {
                                inlineStyle["--annot-color"] = annotation.Color;
                            }

                            // Add label mode
                            var labelMode = settings.AnnotationLabelMode.ToString().ToLowerInvariant();
                            highlightSpan.Add(new XAttribute("data-label-mode", labelMode));

                            // Add metadata if requested
                            if (settings.IncludeAnnotationMetadata)
                            {
                                highlightSpan.Add(new XAttribute("data-label-id", annotation.LabelId ?? ""));
                                if (annotation.Author != null)
                                    highlightSpan.Add(new XAttribute("data-author", annotation.Author));
                                if (annotation.Created.HasValue)
                                    highlightSpan.Add(new XAttribute("data-created", annotation.Created.Value.ToString("o")));
                                if (annotation.StartPage.HasValue)
                                    highlightSpan.Add(new XAttribute("data-start-page", annotation.StartPage.Value.ToString()));
                                if (annotation.EndPage.HasValue)
                                    highlightSpan.Add(new XAttribute("data-end-page", annotation.EndPage.Value.ToString()));
                            }

                            // Add floating label if this is the first segment of the annotation
                            // and label mode is not None
                            if (settings.AnnotationLabelMode != AnnotationLabelMode.None &&
                                !annotTracker.FirstSegmentRendered.Contains(annotationId))
                            {
                                annotTracker.FirstSegmentRendered.Add(annotationId);
                                var labelSpan = new XElement(Xhtml.span,
                                    new XAttribute("class", prefix + "label"),
                                    new XText(annotation.Label ?? annotation.LabelId ?? ""));
                                highlightSpan.AddFirst(labelSpan);
                            }
                            else if (annotTracker.FirstSegmentRendered.Contains(annotationId))
                            {
                                // Mark as continuation for CSS styling
                                var existingClass = (string)highlightSpan.Attribute("class");
                                highlightSpan.SetAttributeValue("class", existingClass + " " + prefix + "continuation");
                            }

                            // Apply inline style
                            if (inlineStyle.Any())
                            {
                                highlightSpan.AddAnnotation(inlineStyle);
                            }

                            highlightSpan.Add(content);
                            content = highlightSpan;
                        }
                    }
                }
            }

            return content;
        }

        private static string DescribeFormatChange(XElement currentRPr, XElement rPrChange)
        {
            var changes = new List<string>();
            var previousRPr = rPrChange.Element(W.rPr);

            // Check for bold change
            var currentBold = currentRPr.Element(W.b) != null;
            var previousBold = previousRPr?.Element(W.b) != null;
            if (currentBold != previousBold)
                changes.Add(currentBold ? "Bold added" : "Bold removed");

            // Check for italic change
            var currentItalic = currentRPr.Element(W.i) != null;
            var previousItalic = previousRPr?.Element(W.i) != null;
            if (currentItalic != previousItalic)
                changes.Add(currentItalic ? "Italic added" : "Italic removed");

            // Check for underline change
            var currentUnderline = currentRPr.Element(W.u) != null;
            var previousUnderline = previousRPr?.Element(W.u) != null;
            if (currentUnderline != previousUnderline)
                changes.Add(currentUnderline ? "Underline added" : "Underline removed");

            // Check for strikethrough change
            var currentStrike = currentRPr.Element(W.strike) != null;
            var previousStrike = previousRPr?.Element(W.strike) != null;
            if (currentStrike != previousStrike)
                changes.Add(currentStrike ? "Strikethrough added" : "Strikethrough removed");

            // Check for font size change
            var currentSz = (string)currentRPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
            var previousSz = (string)previousRPr?.Elements(W.sz).Attributes(W.val).FirstOrDefault();
            if (currentSz != previousSz)
                changes.Add("Font size changed");

            // Check for font change
            var currentFont = (string)currentRPr.Elements(W.rFonts).Attributes(W.ascii).FirstOrDefault();
            var previousFont = (string)previousRPr?.Elements(W.rFonts).Attributes(W.ascii).FirstOrDefault();
            if (currentFont != previousFont)
                changes.Add("Font changed");

            // Check for color change
            var currentColor = (string)currentRPr.Elements(W.color).Attributes(W.val).FirstOrDefault();
            var previousColor = (string)previousRPr?.Elements(W.color).Attributes(W.val).FirstOrDefault();
            if (currentColor != previousColor)
                changes.Add("Color changed");

            if (changes.Count == 0)
                return "Format changed";

            return string.Join(", ", changes);
        }

        [SuppressMessage("ReSharper", "FunctionComplexityOverflow")]
        private static Dictionary<string, string> DefineRunStyle(XElement run)
        {
            var style = new Dictionary<string, string>();

            var rPr = run.Elements(W.rPr).First();

            var styleName = (string) run.Attribute(PtOpenXml.StyleName);
            if (styleName != null)
                style.Add("PtStyleName", styleName);

            // W.bdr
            if (rPr.Element(W.bdr) != null && (string) rPr.Elements(W.bdr).Attributes(W.val).FirstOrDefault() != "none")
            {
                style.AddIfMissing("border", "solid windowtext 1.0pt");
                style.AddIfMissing("padding", "0");
            }

            // W.color
            var color = (string) rPr.Elements(W.color).Attributes(W.val).FirstOrDefault();
            if (color != null)
                CreateColorProperty("color", color, style);

            // W.highlight
            var highlight = (string) rPr.Elements(W.highlight).Attributes(W.val).FirstOrDefault();
            if (highlight != null)
                CreateColorProperty("background", highlight, style);

            // W.shd
            var shade = (string) rPr.Elements(W.shd).Attributes(W.fill).FirstOrDefault();
            if (shade != null)
                CreateColorProperty("background", shade, style);

            // Pt.FontName
            var sym = run.Element(W.sym);
            var font = sym != null
                ? (string) sym.Attributes(W.font).FirstOrDefault()
                : (string) run.Attributes(PtOpenXml.FontName).FirstOrDefault();
            if (font != null)
                CreateFontCssProperty(font, style);

            // W.sz
            var languageType = (string)run.Attribute(PtOpenXml.LanguageType);
            var sz = GetFontSize(languageType, rPr);
            if (sz != null)
                style.AddIfMissing("font-size", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", sz/2.0m));

            // W.caps
            if (GetBoolProp(rPr, W.caps))
                style.AddIfMissing("text-transform", "uppercase");

            // W.smallCaps
            if (GetBoolProp(rPr, W.smallCaps))
                style.AddIfMissing("font-variant", "small-caps");

            // W.spacing
            var spacingInTwips = (decimal?) rPr.Elements(W.spacing).Attributes(W.val).FirstOrDefault();
            if (spacingInTwips != null)
                style.AddIfMissing("letter-spacing",
                    spacingInTwips > 0m
                        ? string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", spacingInTwips/20)
                        : "0");

            // W.position
            var position = (decimal?) rPr.Elements(W.position).Attributes(W.val).FirstOrDefault();
            if (position != null)
            {
                style.AddIfMissing("position", "relative");
                style.AddIfMissing("top", string.Format(NumberFormatInfo.InvariantInfo, "{0}pt", -(position/2)));
            }

            // W.vanish
            if (GetBoolProp(rPr, W.vanish) && !GetBoolProp(rPr, W.specVanish))
                style.AddIfMissing("display", "none");

            // W.u
            if (rPr.Element(W.u) != null && (string) rPr.Elements(W.u).Attributes(W.val).FirstOrDefault() != "none")
                style.AddIfMissing("text-decoration", "underline");

            // W.i
            style.AddIfMissing("font-style", GetBoolProp(rPr, W.i) ? "italic" : "normal");

            // W.b
            style.AddIfMissing("font-weight", GetBoolProp(rPr, W.b) ? "bold" : "normal");

            // W.strike
            if (GetBoolProp(rPr, W.strike) || GetBoolProp(rPr, W.dstrike))
                style.AddIfMissing("text-decoration", "line-through");

            return style;
        }

        private static decimal? GetFontSize(XElement e)
        {
            var languageType = (string)e.Attribute(PtOpenXml.LanguageType);
            if (e.Name == W.p)
            {
                return GetFontSize(languageType, e.Elements(W.pPr).Elements(W.rPr).FirstOrDefault());
            }
            if (e.Name == W.r)
            {
                return GetFontSize(languageType, e.Element(W.rPr));
            }
            return null;
        }

        private static decimal? GetFontSize(string languageType, XElement rPr)
        {
            if (rPr == null) return null;
            return languageType == "bidi"
                ? (decimal?) rPr.Elements(W.szCs).Attributes(W.val).FirstOrDefault()
                : (decimal?) rPr.Elements(W.sz).Attributes(W.val).FirstOrDefault();
        }

        private static void DetermineRunMarks(XElement run, XElement rPr, Dictionary<string, string> style, out XEntity runStartMark, out XEntity runEndMark)
        {
            runStartMark = null;
            runEndMark = null;

            // Only do the following for text runs.
            if (run.Element(W.t) == null) return;

            // Can't add directional marks if the font-family is symbol - they are visible, and display as a ?
            var addDirectionalMarks = true;
            if (style.ContainsKey("font-family"))
            {
                if (style["font-family"].ToLower() == "symbol")
                    addDirectionalMarks = false;
            }
            if (!addDirectionalMarks) return;

            var isRtl = rPr.Element(W.rtl) != null;
            if (isRtl)
            {
                runStartMark = new XEntity("#x200f"); // RLM
                runEndMark = new XEntity("#x200f"); // RLM
            }
            else
            {
                var paragraph = run.Ancestors(W.p).First();
                var paraIsBidi = paragraph
                    .Elements(W.pPr)
                    .Elements(W.bidi)
                    .Any(b => b.Attribute(W.val) == null || b.Attribute(W.val).ToBoolean() == true);

                if (paraIsBidi)
                {
                    runStartMark = new XEntity("#x200e"); // LRM
                    runEndMark = new XEntity("#x200e"); // LRM
                }
            }
        }

        private static XAttribute GetLangAttribute(XElement run)
        {
            const string defaultLanguage = "en-US"; // todo need to get defaultLanguage

            var rPr = run.Elements(W.rPr).First();
            var languageType = (string)run.Attribute(PtOpenXml.LanguageType);

            string lang = null;
            if (languageType == "western")
                lang = (string) rPr.Elements(W.lang).Attributes(W.val).FirstOrDefault();
            else if (languageType == "bidi")
                lang = (string) rPr.Elements(W.lang).Attributes(W.bidi).FirstOrDefault();
            else if (languageType == "eastAsia")
                lang = (string) rPr.Elements(W.lang).Attributes(W.eastAsia).FirstOrDefault();
            if (lang == null)
                lang = defaultLanguage;

            return lang != defaultLanguage ? new XAttribute("lang", lang) : null;
        }

        private static void AdjustTableBorders(WordprocessingDocument wordDoc)
        {
            // Note: when implementing a paging version of the HTML transform, this needs to be done
            // for all content parts, not just the main document part.

            var xd = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var tbl in xd.Descendants(W.tbl))
                AdjustTableBorders(tbl);
            wordDoc.MainDocumentPart.PutXDocument();
        }

        private static void AdjustTableBorders(XElement tbl)
        {
            var ta = tbl
                .Elements(W.tr)
                .Select(r => r
                    .Elements(W.tc)
                    .SelectMany(c =>
                        Enumerable.Repeat(c,
                            (int?) c.Elements(W.tcPr).Elements(W.gridSpan).Attributes(W.val).FirstOrDefault() ?? 1))
                    .ToArray())
                .ToArray();

            for (var y = 0; y < ta.Length; y++)
            {
                for (var x = 0; x < ta[y].Length; x++)
                {
                    var thisCell = ta[y][x];
                    FixTopBorder(ta, thisCell, x, y);
                    FixLeftBorder(ta, thisCell, x, y);
                    FixBottomBorder(ta, thisCell, x, y);
                    FixRightBorder(ta, thisCell, x, y);
                }
            }
        }

        private static void FixTopBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (y > 0)
            {
                var rowAbove = ta[y - 1];
                if (x < rowAbove.Length - 1)
                {
                    XElement cellAbove = ta[y - 1][x];
                    if (cellAbove != null &&
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                        cellAbove.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                    {
                        ResolveCellBorder(
                            thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.top).FirstOrDefault(),
                            cellAbove.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.bottom).FirstOrDefault());
                    }
                }
            }
        }

        private static void FixLeftBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (x > 0)
            {
                XElement cellLeft = ta[y][x - 1];
                if (cellLeft != null &&
                    thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                    cellLeft.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                {
                    ResolveCellBorder(
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.left).FirstOrDefault(),
                        cellLeft.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.right).FirstOrDefault());
                }
            }
        }

        private static void FixBottomBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (y < ta.Length - 1)
            {
                var rowBelow = ta[y + 1];
                if (x < rowBelow.Length - 1)
                {
                    XElement cellBelow = ta[y + 1][x];
                    if (cellBelow != null &&
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                        cellBelow.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                    {
                        ResolveCellBorder(
                            thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.bottom).FirstOrDefault(),
                            cellBelow.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.top).FirstOrDefault());
                    }
                }
            }
        }

        private static void FixRightBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (x < ta[y].Length - 1)
            {
                XElement cellRight = ta[y][x + 1];
                if (cellRight != null &&
                    thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null &&
                    cellRight.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() != null)
                {
                    ResolveCellBorder(
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.right).FirstOrDefault(),
                        cellRight.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.left).FirstOrDefault());
                }
            }
        }

        private static readonly Dictionary<string, int> BorderTypePriority = new Dictionary<string, int>()
        {
            { "single", 1 },
            { "thick", 2 },
            { "double", 3 },
            { "dotted", 4 },
        };

        private static readonly Dictionary<string, int> BorderNumber = new Dictionary<string, int>()
        {
            {"single", 1 },
            {"thick", 2 },
            {"double", 3 },
            {"dotted", 4 },
            {"dashed", 5 },
            {"dotDash", 5 },
            {"dotDotDash", 5 },
            {"triple", 6 },
            {"thinThickSmallGap", 6 },
            {"thickThinSmallGap", 6 },
            {"thinThickThinSmallGap", 6 },
            {"thinThickMediumGap", 6 },
            {"thickThinMediumGap", 6 },
            {"thinThickThinMediumGap", 6 },
            {"thinThickLargeGap", 6 },
            {"thickThinLargeGap", 6 },
            {"thinThickThinLargeGap", 6 },
            {"wave", 7 },
            {"doubleWave", 7 },
            {"dashSmallGap", 5 },
            {"dashDotStroked", 5 },
            {"threeDEmboss", 7 },
            {"threeDEngrave", 7 },
            {"outset", 7 },
            {"inset", 7 },
        };

        private static void ResolveCellBorder(XElement border1, XElement border2)
        {
            if (border1 == null || border2 == null)
                return;
            if ((string)border1.Attribute(W.val) == "nil" || (string)border2.Attribute(W.val) == "nil")
                return;
            if ((string)border1.Attribute(W.sz) == "nil" || (string)border2.Attribute(W.sz) == "nil")
                return;

            var border1Val = (string)border1.Attribute(W.val);
            var border1Weight = 1;
            if (BorderNumber.ContainsKey(border1Val))
                border1Weight = BorderNumber[border1Val];

            var border2Val = (string)border2.Attribute(W.val);
            var border2Weight = 1;
            if (BorderNumber.ContainsKey(border2Val))
                border2Weight = BorderNumber[border2Val];

            if (border1Weight != border2Weight)
            {
                if (border1Weight < border2Weight)
                    BorderOverride(border2, border1);
                else
                    BorderOverride(border1, border2);
            }

            if ((decimal)border1.Attribute(W.sz) > (decimal)border2.Attribute(W.sz))
            {
                BorderOverride(border1, border2);
                return;
            }

            if ((decimal)border1.Attribute(W.sz) < (decimal)border2.Attribute(W.sz))
            {
                BorderOverride(border2, border1);
                return;
            }

            var border1Type = (string)border1.Attribute(W.val);
            var border2Type = (string)border2.Attribute(W.val);
            if (BorderTypePriority.ContainsKey(border1Type) &&
                BorderTypePriority.ContainsKey(border2Type))
            {
                var border1Pri = BorderTypePriority[border1Type];
                var border2Pri = BorderTypePriority[border2Type];
                if (border1Pri < border2Pri)
                {
                    BorderOverride(border2, border1);
                    return;
                }
                if (border2Pri < border1Pri)
                {
                    BorderOverride(border1, border2);
                    return;
                }
            }

            var color1Str = (string)border1.Attribute(W.color);
            if (color1Str == "auto")
                color1Str = "000000";
            var color2Str = (string)border2.Attribute(W.color);
            if (color2Str == "auto")
                color2Str = "000000";
            if (color1Str != null && color2Str != null && color1Str != color2Str)
            {
                try
                {
                    var color1 = Convert.ToInt32(color1Str, 16);
                    var color2 = Convert.ToInt32(color2Str, 16);
                    if (color1 < color2)
                    {
                        BorderOverride(border1, border2);
                        return;
                    }
                    if (color2 < color1)
                    {
                        BorderOverride(border2, border1);
                    }
                }
                // if the above throws ArgumentException, FormatException, or OverflowException, then abort
                catch (Exception)
                {
                    // Ignore
                }
            }
        }

        private static void BorderOverride(XElement fromBorder, XElement toBorder)
        {
            toBorder.Attribute(W.val).Value = fromBorder.Attribute(W.val).Value;
            if (fromBorder.Attribute(W.color) != null)
                toBorder.SetAttributeValue(W.color, fromBorder.Attribute(W.color).Value);
            if (fromBorder.Attribute(W.sz) != null)
                toBorder.SetAttributeValue(W.sz, fromBorder.Attribute(W.sz).Value);
            if (fromBorder.Attribute(W.themeColor) != null)
                toBorder.SetAttributeValue(W.themeColor, fromBorder.Attribute(W.themeColor).Value);
            if (fromBorder.Attribute(W.themeTint) != null)
                toBorder.SetAttributeValue(W.themeTint, fromBorder.Attribute(W.themeTint).Value);
        }

        private static void CalculateSpanWidthForTabs(WordprocessingDocument wordDoc)
        {
            // Note: when implementing a paging version of the HTML transform, this needs to be done
            // for all content parts, not just the main document part.

            // w:defaultTabStop in settings
            var sxd = wordDoc.MainDocumentPart.DocumentSettingsPart.GetXDocument();
            var defaultTabStopValue = (string)sxd.Descendants(W.defaultTabStop).Attributes(W.val).FirstOrDefault();
            var defaultTabStop = defaultTabStopValue != null ? WordprocessingMLUtil.StringToTwips(defaultTabStopValue) : 720;

            var pxd = wordDoc.MainDocumentPart.GetXDocument();
            var root = pxd.Root;
            if (root == null) return;

            var newRoot = (XElement)CalculateSpanWidthTransform(root, defaultTabStop);
            root.ReplaceWith(newRoot);
            wordDoc.MainDocumentPart.PutXDocument();
        }

        // TODO: Refactor. This method is way too long.
        [SuppressMessage("ReSharper", "FunctionComplexityOverflow")]
        private static object CalculateSpanWidthTransform(XNode node, int defaultTabStop)
        {
            var element = node as XElement;
            if (element == null) return node;

            // if it is not a paragraph or if there are no tabs in the paragraph,
            // then no need to continue processing.
            if (element.Name != W.p ||
                !element.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.r).Elements(W.tab).Any())
            {
                // TODO: Revisit. Can we just return the node if it is a paragraph that does not have any tab?
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CalculateSpanWidthTransform(n, defaultTabStop)));
            }

            var clonedPara = new XElement(element);

            var leftInTwips = 0;
            var firstInTwips = 0;

            var ind = clonedPara.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
            if (ind != null)
            {
                // todo need to handle start and end attributes

                var left = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.left));
                if (left != null)
                    leftInTwips = (int)left;

                var firstLine = 0;
                var firstLineAtt = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.firstLine));
                if (firstLineAtt != null)
                    firstLine = (int)firstLineAtt;

                var hangingAtt = WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.hanging));
                if (hangingAtt != null)
                    firstLine = -(int)hangingAtt;

                firstInTwips = leftInTwips + firstLine;
            }

            // calculate the tab stops, in twips
            var tabs = clonedPara
                .Elements(W.pPr)
                .Elements(W.tabs)
                .FirstOrDefault();

            if (tabs == null)
            {
                if (leftInTwips == 0)
                {
                    tabs = new XElement(W.tabs,
                        Enumerable.Range(1, 100)
                            .Select(r => new XElement(W.tab,
                                new XAttribute(W.val, "left"),
                                new XAttribute(W.pos, r * defaultTabStop))));
                }
                else
                {
                    tabs = new XElement(W.tabs,
                        new XElement(W.tab,
                            new XAttribute(W.val, "left"),
                            new XAttribute(W.pos, leftInTwips)));
                    tabs = AddDefaultTabsAfterLastTab(tabs, defaultTabStop);
                }
            }
            else
            {
                if (leftInTwips != 0)
                {
                    tabs.Add(
                        new XElement(W.tab,
                            new XAttribute(W.val, "left"),
                            new XAttribute(W.pos, leftInTwips)));
                }
                tabs = AddDefaultTabsAfterLastTab(tabs, defaultTabStop);
            }

            var twipCounter = firstInTwips;
            var contentToMeasure = element.DescendantsTrimmed(z => z.Name == W.txbxContent || z.Name == W.pPr || z.Name == W.rPr).ToArray();
            var currentElementIdx = 0;
            while (true)
            {
                var currentElement = contentToMeasure[currentElementIdx];

                if (currentElement.Name == W.br)
                {
                    twipCounter = leftInTwips;

                    currentElement.Add(new XAttribute(PtOpenXml.TabWidth,
                        string.Format(NumberFormatInfo.InvariantInfo,
                            "{0:0.000}", (decimal)firstInTwips / 1440m)));

                    currentElementIdx++;
                    if (currentElementIdx >= contentToMeasure.Length)
                        break; // we're done
                }

                if (currentElement.Name == W.tab)
                {
                    var runContainingTabToReplace = currentElement.Parent;
                    var fontNameAtt = runContainingTabToReplace.Attribute(PtOpenXml.pt + "FontName") ??
                                      runContainingTabToReplace.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");

                    var testAmount = twipCounter;

                    var tabAfterText = tabs
                        .Elements(W.tab)
                        .FirstOrDefault(t => WordprocessingMLUtil.StringToTwips((string)t.Attribute(W.pos)) > testAmount);

                    if (tabAfterText == null)
                    {
                        // something has gone wrong, so put 1/2 inch in
                        if (currentElement.Attribute(PtOpenXml.TabWidth) == null)
                            currentElement.Add(
                                new XAttribute(PtOpenXml.TabWidth, 720m));
                        break;
                    }

                    var tabVal = (string)tabAfterText.Attribute(W.val);
                    if (tabVal == "right" || tabVal == "end")
                    {
                        var textAfterElements = contentToMeasure
                            .Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z =>
                                z.Name != W.tab &&
                                z.Name != W.br &&
                                z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        var dummyRun2 = new XElement(W.r,
                            fontNameAtt,
                            runContainingTabToReplace.Elements(W.rPr),
                            new XElement(W.t, textAfterTab));

                        var widthOfTextAfterTab = CalcWidthOfRunInTwips(dummyRun2);
                        var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfTextAfterTab - twipCounter;
                        if (delta2 < 0)
                            delta2 = 0;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth,
                                string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", (decimal)delta2 / 1440m)),
                            GetLeader(tabAfterText));
                        twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)), twipCounter + widthOfTextAfterTab);

                        var lastElement = textElementsToMeasure.LastOrDefault();
                        if (lastElement == null)
                            break; // we're done

                        currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                        if (currentElementIdx >= contentToMeasure.Length)
                            break; // we're done

                        continue;
                    }
                    if (tabVal == "decimal")
                    {
                        var textAfterElements = contentToMeasure
                            .Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z =>
                                z.Name != W.tab &&
                                z.Name != W.br &&
                                z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        if (textAfterTab.Contains("."))
                        {
                            var mantissa = textAfterTab.Split('.')[0];

                            var dummyRun4 = new XElement(W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, mantissa));

                            var widthOfMantissa = CalcWidthOfRunInTwips(dummyRun4);
                            var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfMantissa - twipCounter;
                            if (delta2 < 0)
                                delta2 = 0;
                            currentElement.Add(
                                new XAttribute(PtOpenXml.TabWidth,
                                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", (decimal)delta2 / 1440m)),
                                GetLeader(tabAfterText));

                            var decims = textAfterTab.Substring(textAfterTab.IndexOf('.'));
                            dummyRun4 = new XElement(W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, decims));

                            var widthOfDecims = CalcWidthOfRunInTwips(dummyRun4);
                            twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) + widthOfDecims, twipCounter + widthOfMantissa + widthOfDecims);

                            var lastElement = textElementsToMeasure.LastOrDefault();
                            if (lastElement == null)
                                break; // we're done

                            currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                            if (currentElementIdx >= contentToMeasure.Length)
                                break; // we're done

                            continue;
                        }
                        else
                        {
                            var dummyRun2 = new XElement(W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, textAfterTab));

                            var widthOfTextAfterTab = CalcWidthOfRunInTwips(dummyRun2);
                            var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - widthOfTextAfterTab - twipCounter;
                            if (delta2 < 0)
                                delta2 = 0;
                            currentElement.Add(
                                new XAttribute(PtOpenXml.TabWidth,
                                    string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", (decimal)delta2 / 1440m)),
                                GetLeader(tabAfterText));
                            twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)), twipCounter + widthOfTextAfterTab);

                            var lastElement = textElementsToMeasure.LastOrDefault();
                            if (lastElement == null)
                                break; // we're done

                            currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                            if (currentElementIdx >= contentToMeasure.Length)
                                break; // we're done

                            continue;
                        }
                    }
                    if ((string)tabAfterText.Attribute(W.val) == "center")
                    {
                        var textAfterElements = contentToMeasure
                            .Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z =>
                                z.Name != W.tab &&
                                z.Name != W.br &&
                                z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        var dummyRun4 = new XElement(W.r,
                            fontNameAtt,
                            runContainingTabToReplace.Elements(W.rPr),
                            new XElement(W.t, textAfterTab));

                        var widthOfText = CalcWidthOfRunInTwips(dummyRun4);
                        var delta2 = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - (widthOfText / 2) - twipCounter;
                        if (delta2 < 0)
                            delta2 = 0;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth,
                                string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", (decimal)delta2 / 1440m)),
                            GetLeader(tabAfterText));
                        twipCounter = Math.Max(WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) + widthOfText / 2, twipCounter + widthOfText);

                        var lastElement = textElementsToMeasure.LastOrDefault();
                        if (lastElement == null)
                            break; // we're done

                        currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                        if (currentElementIdx >= contentToMeasure.Length)
                            break; // we're done

                        continue;
                    }
                    if (tabVal == "left" || tabVal == "start" || tabVal == "num")
                    {
                        var delta = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos)) - twipCounter;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth,
                                string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", (decimal)delta / 1440m)),
                            GetLeader(tabAfterText));
                        twipCounter = WordprocessingMLUtil.StringToTwips((string)tabAfterText.Attribute(W.pos));

                        currentElementIdx++;
                        if (currentElementIdx >= contentToMeasure.Length)
                            break; // we're done

                        continue;
                    }
                }

                if (currentElement.Name == W.t)
                {
                    // TODO: Revisit. This is a quick fix because it doesn't work on Azure.
                    // Given the changes we've made elsewhere, though, this is not required
                    // for the first tab at least. We could also enhance that other change
                    // to deal with all tabs.
                    //var runContainingTabToReplace = currentElement.Parent;
                    //var paragraphForRun = runContainingTabToReplace.Ancestors(W.p).First();
                    //var fontNameAtt = runContainingTabToReplace.Attribute(PtOpenXml.FontName) ??
                    //                  paragraphForRun.Attribute(PtOpenXml.FontName);
                    //var languageTypeAtt = runContainingTabToReplace.Attribute(PtOpenXml.LanguageType) ??
                    //                      paragraphForRun.Attribute(PtOpenXml.LanguageType);

                    //var dummyRun3 = new XElement(W.r, fontNameAtt, languageTypeAtt,
                    //    runContainingTabToReplace.Elements(W.rPr),
                    //    currentElement);
                    //var widthOfText = CalcWidthOfRunInTwips(dummyRun3);
                    const int widthOfText = 0;
                    currentElement.Add(new XAttribute(PtOpenXml.TabWidth,
                        string.Format(NumberFormatInfo.InvariantInfo, "{0:0.000}", (decimal) widthOfText/1440m)));
                    twipCounter += widthOfText;

                    currentElementIdx++;
                    if (currentElementIdx >= contentToMeasure.Length)
                        break; // we're done

                    continue;
                }

                currentElementIdx++;
                if (currentElementIdx >= contentToMeasure.Length)
                    break; // we're done
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => CalculateSpanWidthTransform(n, defaultTabStop)));
        }

        private static XAttribute GetLeader(XElement tabAfterText)
        {
            var leader = (string)tabAfterText.Attribute(W.leader);
            if (leader == null)
                return null;
            return new XAttribute(PtOpenXml.Leader, leader);
        }

        private static XElement AddDefaultTabsAfterLastTab(XElement tabs, int defaultTabStop)
        {
            var lastTabElement = tabs
                .Elements(W.tab)
                .Where(t => (string)t.Attribute(W.val) != "clear" && (string)t.Attribute(W.val) != "bar")
                .OrderBy(t => WordprocessingMLUtil.StringToTwips((string)t.Attribute(W.pos)))
                .LastOrDefault();
            if (lastTabElement != null)
            {
                if (defaultTabStop == 0)
                    defaultTabStop = 720;
                var rangeStart = WordprocessingMLUtil.StringToTwips((string)lastTabElement.Attribute(W.pos)) / defaultTabStop + 1;
                var tempTabs = new XElement(W.tabs,
                    tabs.Elements().Where(t => (string)t.Attribute(W.val) != "clear" && (string)t.Attribute(W.val) != "bar"),
                    Enumerable.Range(rangeStart, 100)
                    .Select(r => new XElement(W.tab,
                        new XAttribute(W.val, "left"),
                        new XAttribute(W.pos, r * defaultTabStop))));
                tempTabs = new XElement(W.tabs,
                    tempTabs.Elements().OrderBy(t => WordprocessingMLUtil.StringToTwips((string)t.Attribute(W.pos))));
                return tempTabs;
            }
            else
            {
                tabs = new XElement(W.tabs,
                    Enumerable.Range(1, 100)
                    .Select(r => new XElement(W.tab,
                        new XAttribute(W.val, "left"),
                        new XAttribute(W.pos, r * defaultTabStop))));
            }
            return tabs;
        }

        private static readonly HashSet<string> UnknownFonts = new HashSet<string>();
        private static HashSet<string> _knownFamilies;

        private static HashSet<string> KnownFamilies
        {
            get
            {
                if (_knownFamilies == null)
                {
                    _knownFamilies = new HashSet<string>();
                    var families = SKFontManager.Default.FontFamilies;
                    foreach (var fam in families)
                        _knownFamilies.Add(fam);
                }
                return _knownFamilies;
            }
        }

        private static int CalcWidthOfRunInTwips(XElement r)
        {
            var fontName = (string)r.Attribute(PtOpenXml.pt + "FontName") ??
                           (string)r.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");
            if (fontName == null)
                throw new DocxodusException("Internal Error, should have FontName attribute");
            if (UnknownFonts.Contains(fontName))
                return 0;

            var rPr = r.Element(W.rPr);
            if (rPr == null)
                throw new DocxodusException("Internal Error, should have run properties");

            var sz = GetFontSize(r) ?? 22m;

            // unknown font families will throw ArgumentException, in which case just return 0
            if (!KnownFamilies.Contains(fontName))
                return 0;

            var bold = GetBoolProp(rPr, W.b) || GetBoolProp(rPr, W.bCs);
            var italic = GetBoolProp(rPr, W.i) || GetBoolProp(rPr, W.iCs);

            // Appended blank as a quick fix to accommodate &nbsp; that will get
            // appended to some layout-critical runs such as list item numbers.
            // In some cases, this might not be required or even wrong, so this
            // must be revisited.
            // TODO: Revisit.
            var runText = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.t)
                .Select(t => (string) t)
                .StringConcatenate() + " ";

            var tabLength = r.DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.tab)
                .Select(t => (decimal)t.Attribute(PtOpenXml.TabWidth))
                .Sum();

            if (runText.Length == 0 && tabLength == 0)
                return 0;

            int multiplier = 1;
            if (runText.Length <= 2)
                multiplier = 100;
            else if (runText.Length <= 4)
                multiplier = 50;
            else if (runText.Length <= 8)
                multiplier = 25;
            else if (runText.Length <= 16)
                multiplier = 12;
            else if (runText.Length <= 32)
                multiplier = 6;
            if (multiplier != 1)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < multiplier; i++)
                    sb.Append(runText);
                runText = sb.ToString();
            }

            var w = MetricsGetter.GetTextWidth(fontName, bold, italic, sz, runText);

            return (int)(w / 96m * 1440m / multiplier + tabLength * 1440m);
        }

        private static void InsertAppropriateNonbreakingSpaces(WordprocessingDocument wordDoc)
        {
            foreach (var part in wordDoc.ContentParts())
            {
                var pxd = part.GetXDocument();
                var root = pxd.Root;
                if (root == null) return;

                var newRoot = (XElement)InsertAppropriateNonbreakingSpacesTransform(root);
                root.ReplaceWith(newRoot);
                part.PutXDocument();
            }
        }

        // Non-breaking spaces are not required if we use appropriate CSS, i.e., "white-space: pre-wrap;".
        // We only need to make sure that empty w:p elements are translated into non-empty h:p elements,
        // because empty h:p elements would be ignored by browsers.
        // Further, in addition to not being required, non-breaking spaces would change the layout behavior
        // of spans having consecutive spaces. Therefore, avoiding non-breaking spaces has the additional
        // benefit of leading to a more faithful representation of the Word document in HTML.
        private static object InsertAppropriateNonbreakingSpacesTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                // child content of run to look for
                // W.br
                // W.cr
                // W.dayLong
                // W.dayShort
                // W.drawing
                // W.monthLong
                // W.monthShort
                // W.noBreakHyphen
                // W.object
                // W.pgNum
                // W.pTab
                // W.separator
                // W.softHyphen
                // W.sym
                // W.t
                // W.tab
                // W.yearLong
                // W.yearShort
                if (element.Name == W.p)
                {
                    // Translate empty paragraphs to paragraphs having one run with
                    // a normal space. A non-breaking space, i.e., \x00A0, is not
                    // required if we use appropriate CSS.
                    bool hasContent = element
                        .Elements()
                        .Where(e => e.Name != W.pPr)
                        .DescendantsAndSelf()
                        .Any(e =>
                            e.Name == W.dayLong ||
                            e.Name == W.dayShort ||
                            e.Name == W.drawing ||
                            e.Name == W.monthLong ||
                            e.Name == W.monthShort ||
                            e.Name == W.noBreakHyphen ||
                            e.Name == W._object ||
                            e.Name == W.pgNum ||
                            e.Name == W.ptab ||
                            e.Name == W.separator ||
                            e.Name == W.softHyphen ||
                            e.Name == W.sym ||
                            e.Name == W.t ||
                            e.Name == W.tab ||
                            e.Name == W.yearLong ||
                            e.Name == W.yearShort
                        );
                    if (hasContent == false)
                        return new XElement(element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => InsertAppropriateNonbreakingSpacesTransform(n)),
                            new XElement(W.r,
                                element.Elements(W.pPr).Elements(W.rPr),
                                new XElement(W.t, " ")));
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => InsertAppropriateNonbreakingSpacesTransform(n)));
            }
            return node;
        }

        private class SectionAnnotation
        {
            public XElement SectionElement;
        }

        /// <summary>
        /// Represents page dimensions extracted from w:sectPr.
        /// All values are in points (1/72 inch) for CSS compatibility.
        /// </summary>
        private class PageDimensions
        {
            /// <summary>Page width in points (from w:pgSz w:w)</summary>
            public double PageWidthPt { get; set; }

            /// <summary>Page height in points (from w:pgSz w:h)</summary>
            public double PageHeightPt { get; set; }

            /// <summary>Top margin in points (from w:pgMar w:top)</summary>
            public double MarginTopPt { get; set; }

            /// <summary>Right margin in points (from w:pgMar w:right)</summary>
            public double MarginRightPt { get; set; }

            /// <summary>Bottom margin in points (from w:pgMar w:bottom)</summary>
            public double MarginBottomPt { get; set; }

            /// <summary>Left margin in points (from w:pgMar w:left)</summary>
            public double MarginLeftPt { get; set; }

            /// <summary>Header distance in points (from w:pgMar w:header)</summary>
            public double HeaderPt { get; set; }

            /// <summary>Footer distance in points (from w:pgMar w:footer)</summary>
            public double FooterPt { get; set; }

            /// <summary>Content width (page width minus left and right margins)</summary>
            public double ContentWidthPt => PageWidthPt - MarginLeftPt - MarginRightPt;

            /// <summary>Content height (page height minus top and bottom margins)</summary>
            public double ContentHeightPt => PageHeightPt - MarginTopPt - MarginBottomPt;

            /// <summary>
            /// Creates PageDimensions from a w:sectPr element.
            /// Returns US Letter defaults (8.5"x11" with 1" margins) if sectPr is null.
            /// </summary>
            public static PageDimensions FromSectionProperties(XElement sectPr)
            {
                // Default to US Letter: 8.5" x 11" (612pt x 792pt) with 1" margins (72pt)
                var dims = new PageDimensions
                {
                    PageWidthPt = 612,   // 8.5 inches
                    PageHeightPt = 792,  // 11 inches
                    MarginTopPt = 72,    // 1 inch
                    MarginRightPt = 72,
                    MarginBottomPt = 72,
                    MarginLeftPt = 72,
                    HeaderPt = 36,       // 0.5 inch
                    FooterPt = 36
                };

                if (sectPr == null) return dims;

                // Parse page size (w:pgSz)
                // w:w and w:h are in twips (1/20 of a point, or 1/1440 of an inch)
                var pgSz = sectPr.Element(W.pgSz);
                if (pgSz != null)
                {
                    if (int.TryParse((string)pgSz.Attribute(W._w), out int w))
                        dims.PageWidthPt = w / 20.0;
                    if (int.TryParse((string)pgSz.Attribute(W.h), out int h))
                        dims.PageHeightPt = h / 20.0;
                }

                // Parse page margins (w:pgMar)
                var pgMar = sectPr.Element(W.pgMar);
                if (pgMar != null)
                {
                    if (int.TryParse((string)pgMar.Attribute(W.top), out int top))
                        dims.MarginTopPt = top / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.right), out int right))
                        dims.MarginRightPt = right / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.bottom), out int bottom))
                        dims.MarginBottomPt = bottom / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.left), out int left))
                        dims.MarginLeftPt = left / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.header), out int header))
                        dims.HeaderPt = header / 20.0;
                    if (int.TryParse((string)pgMar.Attribute(W.footer), out int footer))
                        dims.FooterPt = footer / 20.0;
                }

                return dims;
            }
        }

        private static void AnnotateForSections(WordprocessingDocument wordDoc)
        {
            var xd = wordDoc.MainDocumentPart.GetXDocument();

            var document = xd.Root;
            if (document == null) return;

            var body = document.Element(W.body);
            if (body == null) return;

            // move last sectPr into last paragraph
            var lastSectPr = body.Elements(W.sectPr).LastOrDefault();
            if (lastSectPr != null)
            {
                // if the last thing in the document is a table, Word will always insert a paragraph following that.
                var lastPara = body
                    .DescendantsTrimmed(W.txbxContent)
                    .LastOrDefault(p => p.Name == W.p);

                if (lastPara != null)
                {
                    var lastParaProps = lastPara.Element(W.pPr);
                    if (lastParaProps != null)
                        lastParaProps.Add(lastSectPr);
                    else
                        lastPara.Add(new XElement(W.pPr, lastSectPr));

                    lastSectPr.Remove();
                }
            }

            var reverseDescendants = xd.Descendants().Reverse().ToList();
            var currentSection = InitializeSectionAnnotation(reverseDescendants);

            foreach (var d in reverseDescendants)
            {
                if (d.Name == W.sectPr)
                {
                    if (d.Attribute(XNamespace.Xmlns + "w") == null)
                        d.Add(new XAttribute(XNamespace.Xmlns + "w", W.w));

                    currentSection = new SectionAnnotation()
                    {
                        SectionElement = d
                    };
                }
                else
                    d.AddAnnotation(currentSection);
            }
        }

        private static SectionAnnotation InitializeSectionAnnotation(IEnumerable<XElement> reverseDescendants)
        {
            var currentSection = new SectionAnnotation()
            {
                SectionElement = reverseDescendants.FirstOrDefault(e => e.Name == W.sectPr)
            };
            if (currentSection.SectionElement != null &&
                currentSection.SectionElement.Attribute(XNamespace.Xmlns + "w") == null)
                currentSection.SectionElement.Add(new XAttribute(XNamespace.Xmlns + "w", W.w));

            // todo what should the default section props be?
            if (currentSection.SectionElement == null)
                currentSection = new SectionAnnotation()
                {
                    SectionElement = new XElement(W.sectPr,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        new XElement(W.pgSz,
                            new XAttribute(W._w, 12240),
                            new XAttribute(W.h, 15840)),
                        new XElement(W.pgMar,
                            new XAttribute(W.top, 1440),
                            new XAttribute(W.right, 1440),
                            new XAttribute(W.bottom, 1440),
                            new XAttribute(W.left, 1440),
                            new XAttribute(W.header, 720),
                            new XAttribute(W.footer, 720),
                            new XAttribute(W.gutter, 0)),
                        new XElement(W.cols,
                            new XAttribute(W.space, 720)),
                        new XElement(W.docGrid,
                            new XAttribute(W.linePitch, 360)))
                };

            return currentSection;
        }

        private static object CreateBorderDivs(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings, IEnumerable<XElement> elements)
        {
            return elements.GroupAdjacent(e =>
                {
                    var pBdr = e.Elements(W.pPr).Elements(W.pBdr).FirstOrDefault();
                    if (pBdr != null)
                    {
                        var indStr = string.Empty;
                        var ind = e.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
                        if (ind != null)
                            indStr = ind.ToString(SaveOptions.DisableFormatting);
                        return pBdr.ToString(SaveOptions.DisableFormatting) + indStr;
                    }
                    return e.Name == W.tbl ? "table" : string.Empty;
                })
                .Select(g =>
                {
                    if (g.Key == string.Empty)
                    {
                        return (object) GroupAndVerticallySpaceNumberedParagraphs(wordDoc, settings, g, 0m);
                    }
                    if (g.Key == "table")
                    {
                        return g.Select(gc => ConvertToHtmlTransform(wordDoc, settings, gc, false, 0));
                    }
                    var pPr = g.First().Elements(W.pPr).First();
                    var pBdr = pPr.Element(W.pBdr);
                    var style = new Dictionary<string, string>();
                    GenerateBorderStyle(pBdr, W.top, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.right, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.bottom, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.left, style, BorderType.Paragraph);

                    var currentMarginLeft = 0m;
                    var ind = pPr.Element(W.ind);
                    if (ind != null)
                    {
                        var leftInInches = (decimal?) ind.Attribute(W.left)/1440m ?? 0;
                        var hangingInInches = -(decimal?) ind.Attribute(W.hanging)/1440m ?? 0;
                        currentMarginLeft = leftInInches + hangingInInches;

                        style.AddIfMissing("margin-left",
                            currentMarginLeft > 0m
                                ? string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", currentMarginLeft)
                                : "0");
                    }

                    var div = new XElement(Xhtml.div,
                        GroupAndVerticallySpaceNumberedParagraphs(wordDoc, settings, g, currentMarginLeft));
                    div.AddAnnotation(style);
                    return div;
                })
            .ToList();
        }

        private static IEnumerable<object> GroupAndVerticallySpaceNumberedParagraphs(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            IEnumerable<XElement> elements, decimal currentMarginLeft)
        {
            var grouped = elements
                .GroupAdjacent(e =>
                {
                    var abstractNumId = (string)e.Attribute(PtOpenXml.pt + "AbstractNumId");
                    if (abstractNumId != null)
                        return "num:" + abstractNumId;
                    var contextualSpacing = e.Elements(W.pPr).Elements(W.contextualSpacing).FirstOrDefault();
                    if (contextualSpacing != null)
                    {
                        var styleName = (string)e.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                        if (styleName == null)
                            return "";
                        return "sty:" + styleName;
                    }
                    return "";
                })
                .ToList();
            var newContent = grouped
                .Select(g =>
                {
                    if (g.Key == "")
                        return g.Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
                    var last = g.Count() - 1;
                    return g.Select((e, i) => ConvertToHtmlTransform(wordDoc, settings, e, i != last, currentMarginLeft));
                });
            return (IEnumerable<object>)newContent;
        }

        private class BorderMappingInfo
        {
            public string CssName;
            public decimal CssSize;
        }

        private static readonly Dictionary<string, BorderMappingInfo> BorderStyleMap = new Dictionary<string, BorderMappingInfo>()
        {
            { "single", new BorderMappingInfo() { CssName = "solid", CssSize = 1.0m }},
            { "dotted", new BorderMappingInfo() { CssName = "dotted", CssSize = 1.0m }},
            { "dashSmallGap", new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }},
            { "dashed", new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }},
            { "dotDash", new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }},
            { "dotDotDash", new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }},
            { "double", new BorderMappingInfo() { CssName = "double", CssSize = 2.5m }},
            { "triple", new BorderMappingInfo() { CssName = "double", CssSize = 2.5m }},
            { "thinThickSmallGap", new BorderMappingInfo() { CssName = "double", CssSize = 4.5m }},
            { "thickThinSmallGap", new BorderMappingInfo() { CssName = "double", CssSize = 4.5m }},
            { "thinThickThinSmallGap", new BorderMappingInfo() { CssName = "double", CssSize = 6.0m }},
            { "thickThinMediumGap", new BorderMappingInfo() { CssName = "double", CssSize = 6.0m }},
            { "thinThickMediumGap", new BorderMappingInfo() { CssName = "double", CssSize = 6.0m }},
            { "thinThickThinMediumGap", new BorderMappingInfo() { CssName = "double", CssSize = 9.0m }},
            { "thinThickLargeGap", new BorderMappingInfo() { CssName = "double", CssSize = 5.25m }},
            { "thickThinLargeGap", new BorderMappingInfo() { CssName = "double", CssSize = 5.25m }},
            { "thinThickThinLargeGap", new BorderMappingInfo() { CssName = "double", CssSize = 9.0m }},
            { "wave", new BorderMappingInfo() { CssName = "solid", CssSize = 3.0m }},
            { "doubleWave", new BorderMappingInfo() { CssName = "double", CssSize = 5.25m }},
            { "dashDotStroked", new BorderMappingInfo() { CssName = "solid", CssSize = 3.0m }},
            { "threeDEmboss", new BorderMappingInfo() { CssName = "ridge", CssSize = 6.0m }},
            { "threeDEngrave", new BorderMappingInfo() { CssName = "groove", CssSize = 6.0m }},
            { "outset", new BorderMappingInfo() { CssName = "outset", CssSize = 4.5m }},
            { "inset", new BorderMappingInfo() { CssName = "inset", CssSize = 4.5m }},
        };

        private static void GenerateBorderStyle(XElement pBdr, XName sideXName, Dictionary<string, string> style, BorderType borderType)
        {
            string whichSide;
            if (sideXName == W.top)
                whichSide = "top";
            else if (sideXName == W.right)
                whichSide = "right";
            else if (sideXName == W.bottom)
                whichSide = "bottom";
            else
                whichSide = "left";
            if (pBdr == null)
            {
                style.Add("border-" + whichSide, "none");
                if (borderType == BorderType.Cell &&
                    (whichSide == "left" || whichSide == "right"))
                    style.Add("padding-" + whichSide, "5.4pt");
                return;
            }

            var side = pBdr.Element(sideXName);
            if (side == null)
            {
                style.Add("border-" + whichSide, "none");
                if (borderType == BorderType.Cell &&
                    (whichSide == "left" || whichSide == "right"))
                    style.Add("padding-" + whichSide, "5.4pt");
                return;
            }
            var type = (string)side.Attribute(W.val);
            if (type == "nil" || type == "none")
            {
                style.Add("border-" + whichSide + "-style", "none");

                var space = (decimal?)side.Attribute(W.space) ?? 0;
                if (borderType == BorderType.Cell &&
                    (whichSide == "left" || whichSide == "right"))
                    if (space < 5.4m)
                        space = 5.4m;
                style.Add("padding-" + whichSide,
                    space == 0 ? "0" : string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", space));

            }
            else
            {
                var sz = (int)side.Attribute(W.sz);
                var space = (decimal?)side.Attribute(W.space) ?? 0;
                var color = (string)side.Attribute(W.color);
                if (color == null || color == "auto")
                    color = "windowtext";
                else
                    color = ConvertColor(color);

                decimal borderWidthInPoints = Math.Max(1m, Math.Min(96m, Math.Max(2m, sz)) / 8m);

                var borderStyle = "solid";
                if (BorderStyleMap.ContainsKey(type))
                {
                    var borderInfo = BorderStyleMap[type];
                    borderStyle = borderInfo.CssName;
                    if (type == "double")
                    {
                        if (sz <= 8)
                            borderWidthInPoints = 2.5m;
                        else if (sz <= 18)
                            borderWidthInPoints = 6.75m;
                        else
                            borderWidthInPoints = sz / 3m;
                    }
                    else if (type == "triple")
                    {
                        if (sz <= 8)
                            borderWidthInPoints = 8m;
                        else if (sz <= 18)
                            borderWidthInPoints = 11.25m;
                        else
                            borderWidthInPoints = 11.25m;
                    }
                    else if (type.ToLower().Contains("dash"))
                    {
                        if (sz <= 4)
                            borderWidthInPoints = 1m;
                        else if (sz <= 12)
                            borderWidthInPoints = 1.5m;
                        else
                            borderWidthInPoints = 2m;
                    }
                    else if (type != "single")
                        borderWidthInPoints = borderInfo.CssSize;
                }
                if (type == "outset" || type == "inset")
                    color = "";
                var borderWidth = string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", borderWidthInPoints);

                style.Add("border-" + whichSide, borderStyle + " " + color + " " + borderWidth);
                if (borderType == BorderType.Cell &&
                    (whichSide == "left" || whichSide == "right"))
                    if (space < 5.4m)
                        space = 5.4m;

                style.Add("padding-" + whichSide,
                    space == 0 ? "0" : string.Format(NumberFormatInfo.InvariantInfo, "{0:0.0}pt", space));
            }
        }

        private static readonly Dictionary<string, Func<string, string, string>> ShadeMapper = new Dictionary<string,Func<string, string, string>>()
        {
            { "auto", (c, f) => c },
            { "clear", (c, f) => f },
            { "nil", (c, f) => f },
            { "solid", (c, f) => c },
            { "diagCross", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "diagStripe", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "horzCross", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "horzStripe", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "pct10", (c, f) => ConvertColorFillPct(c, f, .1) },
            { "pct12", (c, f) => ConvertColorFillPct(c, f, .125) },
            { "pct15", (c, f) => ConvertColorFillPct(c, f, .15) },
            { "pct20", (c, f) => ConvertColorFillPct(c, f, .2) },
            { "pct25", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "pct30", (c, f) => ConvertColorFillPct(c, f, .3) },
            { "pct35", (c, f) => ConvertColorFillPct(c, f, .35) },
            { "pct37", (c, f) => ConvertColorFillPct(c, f, .375) },
            { "pct40", (c, f) => ConvertColorFillPct(c, f, .4) },
            { "pct45", (c, f) => ConvertColorFillPct(c, f, .45) },
            { "pct50", (c, f) => ConvertColorFillPct(c, f, .50) },
            { "pct55", (c, f) => ConvertColorFillPct(c, f, .55) },
            { "pct60", (c, f) => ConvertColorFillPct(c, f, .60) },
            { "pct62", (c, f) => ConvertColorFillPct(c, f, .625) },
            { "pct65", (c, f) => ConvertColorFillPct(c, f, .65) },
            { "pct70", (c, f) => ConvertColorFillPct(c, f, .7) },
            { "pct75", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "pct80", (c, f) => ConvertColorFillPct(c, f, .8) },
            { "pct85", (c, f) => ConvertColorFillPct(c, f, .85) },
            { "pct87", (c, f) => ConvertColorFillPct(c, f, .875) },
            { "pct90", (c, f) => ConvertColorFillPct(c, f, .9) },
            { "pct95", (c, f) => ConvertColorFillPct(c, f, .95) },
            { "reverseDiagStripe", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "thinDiagCross", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "thinDiagStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinHorzCross", (c, f) => ConvertColorFillPct(c, f, .3) },
            { "thinHorzStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinReverseDiagStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinVertStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
        };

        private static readonly Dictionary<string, string> ShadeCache = new Dictionary<string, string>();

        // fill is the background, color is the foreground
        private static string ConvertColorFillPct(string color, string fill, double pct)
        {
            if (color == "auto")
                color = "000000";
            if (fill == "auto")
                fill = "ffffff";
            var key = color + fill + pct.ToString(CultureInfo.InvariantCulture);
            if (ShadeCache.ContainsKey(key))
                return ShadeCache[key];
            var fillRed = Convert.ToInt32(fill.Substring(0, 2), 16);
            var fillGreen = Convert.ToInt32(fill.Substring(2, 2), 16);
            var fillBlue = Convert.ToInt32(fill.Substring(4, 2), 16);
            var colorRed = Convert.ToInt32(color.Substring(0, 2), 16);
            var colorGreen = Convert.ToInt32(color.Substring(2, 2), 16);
            var colorBlue = Convert.ToInt32(color.Substring(4, 2), 16);
            var finalRed = (int)(fillRed - (fillRed - colorRed) * pct);
            var finalGreen = (int)(fillGreen - (fillGreen - colorGreen) * pct);
            var finalBlue = (int)(fillBlue - (fillBlue - colorBlue) * pct);
            var returnValue = string.Format("{0:x2}{1:x2}{2:x2}", finalRed, finalGreen, finalBlue);
            ShadeCache.Add(key, returnValue);
            return returnValue;
        }

        private static void CreateStyleFromShd(Dictionary<string, string> style, XElement shd)
        {
            if (shd == null)
                return;
            var shadeType = (string)shd.Attribute(W.val);
            var color = (string)shd.Attribute(W.color);
            var fill = (string)shd.Attribute(W.fill);
            if (ShadeMapper.ContainsKey(shadeType))
            {
                color = ShadeMapper[shadeType](color, fill);
            }
            if (color != null)
            {
                var cvtColor = ConvertColor(color);
                if (!string.IsNullOrEmpty(cvtColor))
                    style.AddIfMissing("background", cvtColor);
            }
        }

        private static readonly Dictionary<string, string> NamedColors = new Dictionary<string, string>()
        {
            {"black", "black"},
            {"blue", "blue" },
            {"cyan", "aqua" },
            {"green", "green" },
            {"magenta", "fuchsia" },
            {"red", "red" },
            {"yellow", "yellow" },
            {"white", "white" },
            {"darkBlue", "#00008B" },
            {"darkCyan", "#008B8B" },
            {"darkGreen", "#006400" },
            {"darkMagenta", "#800080" },
            {"darkRed", "#8B0000" },
            {"darkYellow", "#808000" },
            {"darkGray", "#A9A9A9" },
            {"lightGray", "#D3D3D3" },
            {"none", "" },
        };

        private static void CreateColorProperty(string propertyName, string color, Dictionary<string, string> style)
        {
            if (color == null)
                return;

            // "auto" color is black for "color" and white for "background" property.
            if (color == "auto")
                color = propertyName == "color" ? "black" : "white";

            if (NamedColors.ContainsKey(color))
            {
                var lc = NamedColors[color];
                if (lc == "")
                    return;
                style.AddIfMissing(propertyName, lc);
                return;
            }
            style.AddIfMissing(propertyName, "#" + color);
        }

        private static string ConvertColor(string color)
        {
            // "auto" color is black for "color" and white for "background" property.
            // As this method is only called for "background" colors, "auto" is translated
            // to "white" and never "black".
            if (color == "auto")
                color = "white";

            if (NamedColors.ContainsKey(color))
            {
                var lc = NamedColors[color];
                if (lc == "")
                    return "black";
                return lc;
            }
            return "#" + color;
        }

        private static readonly Dictionary<string, string> FontFallback = new Dictionary<string, string>()
        {
            { "Arial", @"'{0}', 'sans-serif'" },
            { "Arial Narrow", @"'{0}', 'sans-serif'" },
            { "Arial Rounded MT Bold", @"'{0}', 'sans-serif'" },
            { "Arial Unicode MS", @"'{0}', 'sans-serif'" },
            { "Baskerville Old Face", @"'{0}', 'serif'" },
            { "Berlin Sans FB", @"'{0}', 'sans-serif'" },
            { "Berlin Sans FB Demi", @"'{0}', 'sans-serif'" },
            { "Calibri Light", @"'{0}', 'sans-serif'" },
            { "Gill Sans MT", @"'{0}', 'sans-serif'" },
            { "Gill Sans MT Condensed", @"'{0}', 'sans-serif'" },
            { "Lucida Sans", @"'{0}', 'sans-serif'" },
            { "Lucida Sans Unicode", @"'{0}', 'sans-serif'" },
            { "Segoe UI", @"'{0}', 'sans-serif'" },
            { "Segoe UI Light", @"'{0}', 'sans-serif'" },
            { "Segoe UI Semibold", @"'{0}', 'sans-serif'" },
            { "Tahoma", @"'{0}', 'sans-serif'" },
            { "Trebuchet MS", @"'{0}', 'sans-serif'" },
            { "Verdana", @"'{0}', 'sans-serif'" },
            { "Book Antiqua", @"'{0}', 'serif'" },
            { "Bookman Old Style", @"'{0}', 'serif'" },
            { "Californian FB", @"'{0}', 'serif'" },
            { "Cambria", @"'{0}', 'serif'" },
            { "Constantia", @"'{0}', 'serif'" },
            { "Garamond", @"'{0}', 'serif'" },
            { "Lucida Bright", @"'{0}', 'serif'" },
            { "Lucida Fax", @"'{0}', 'serif'" },
            { "Palatino Linotype", @"'{0}', 'serif'" },
            { "Times New Roman", @"'{0}', 'serif'" },
            { "Wide Latin", @"'{0}', 'serif'" },
            { "Courier New", @"'{0}'" },
            { "Lucida Console", @"'{0}'" },
        };

        private static void CreateFontCssProperty(string font, Dictionary<string, string> style)
        {
            if (FontFallback.ContainsKey(font))
            {
                style.AddIfMissing("font-family", string.Format(FontFallback[font], font));
                return;
            }
            style.AddIfMissing("font-family", font);
        }

        private static bool GetBoolProp(XElement runProps, XName xName)
        {
            var p = runProps.Element(xName);
            if (p == null)
                return false;
            var v = p.Attribute(W.val);
            if (v == null)
                return true;
            var s = v.Value.ToLower();
            if (s == "0" || s == "false")
                return false;
            if (s == "1" || s == "true")
                return true;
            return false;
        }

        private static object ConvertContentThatCanContainFields(WordprocessingDocument wordDoc, WmlToHtmlConverterSettings settings,
            IEnumerable<XElement> elements)
        {
            var grouped = elements
                .GroupAdjacent(e =>
                {
                    var stack = e.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                    return stack == null || !stack.Any() ? (int?)null : stack.Select(st => st.Id).Min();
                })
                .ToList();

            var txformed = grouped
                .Select(g =>
                {
                    var key = g.Key;
                    if (key == null)
                        return (object)g.Select(n => ConvertToHtmlTransform(wordDoc, settings, n, false, 0m));

                    var instrText = FieldRetriever.InstrText(g.First().Ancestors().Last(), (int)key)
                        .TrimStart('{').TrimEnd('}');

                    var parsed = FieldRetriever.ParseField(instrText);
                    if (parsed.FieldType != "HYPERLINK")
                        return g.Select(n => ConvertToHtmlTransform(wordDoc, settings, n, false, 0m));

                    var content = g.DescendantsAndSelf(W.r).Select(run => ConvertRun(wordDoc, settings, run));
                    var a = parsed.Arguments.Length > 0
                        ? new XElement(Xhtml.a, new XAttribute("href", parsed.Arguments[0]), content)
                        : new XElement(Xhtml.a, content);
                    var a2 = a as XElement;
                    if (!a2.Nodes().Any())
                    {
                        a2.Add(new XText(""));
                        return a2;
                    }
                    return a;
                })
                .ToList();

            return txformed;
        }

        #region Image Processing

        // Don't process wmf files (with contentType == "image/x-wmf") because GDI consumes huge amounts
        // of memory when dealing with wmf perhaps because it loads a DLL to do the rendering?
        // It actually works, but is not recommended.
        private static readonly List<string> ImageContentTypes = new List<string>
        {
            "image/png", "image/gif", "image/tiff", "image/jpeg"
        };


        public static XElement ProcessImage(WordprocessingDocument wordDoc,
            XElement element, Func<ImageInfo, XElement> imageHandler)
        {
            if (imageHandler == null)
            {
                return null;
            }
            if (element.Name == W.drawing)
            {
                return ProcessDrawing(wordDoc, element, imageHandler);
            }
            if (element.Name == W.pict || element.Name == W._object)
            {
                return ProcessPictureOrObject(wordDoc, element, imageHandler);
            }
            return null;
        }

        private static XElement ProcessDrawing(WordprocessingDocument wordDoc,
            XElement element, Func<ImageInfo, XElement> imageHandler)
        {
            var containerElement = element.Elements()
                .FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
            if (containerElement == null) return null;

            string hyperlinkUri = null;
            var hyperlinkElement = element
                .Elements(WP.inline)
                .Elements(WP.docPr)
                .Elements(A.hlinkClick)
                .FirstOrDefault();
            if (hyperlinkElement != null)
            {
                var rId = (string)hyperlinkElement.Attribute(R.id);
                if (rId != null)
                {
                    var hyperlinkRel = wordDoc.MainDocumentPart.HyperlinkRelationships.FirstOrDefault(hlr => hlr.Id == rId);
                    if (hyperlinkRel != null)
                    {
                        hyperlinkUri = hyperlinkRel.Uri.ToString();
                    }
                }
            }

            var extentCx = (int?)containerElement.Elements(WP.extent)
                .Attributes(NoNamespace.cx).FirstOrDefault();
            var extentCy = (int?)containerElement.Elements(WP.extent)
                .Attributes(NoNamespace.cy).FirstOrDefault();
            var altText = (string)containerElement.Elements(WP.docPr).Attributes(NoNamespace.descr).FirstOrDefault() ??
                          ((string)containerElement.Elements(WP.docPr).Attributes(NoNamespace.name).FirstOrDefault() ?? "");

            var blipFill = containerElement.Elements(A.graphic)
                .Elements(A.graphicData)
                .Elements(Pic._pic).Elements(Pic.blipFill).FirstOrDefault();
            if (blipFill == null) return null;

            var imageRid = (string)blipFill.Elements(A.blip).Attributes(R.embed).FirstOrDefault();
            if (imageRid == null) return null;

            var pp3 = wordDoc.MainDocumentPart.Parts.FirstOrDefault(pp => pp.RelationshipId == imageRid);
            if (pp3 == null) return null;

            var imagePart = (ImagePart)pp3.OpenXmlPart;
            if (imagePart == null) return null;

            // If the image markup points to a NULL image, then following will throw an ArgumentOutOfRangeException
            try
            {
                imagePart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(imageRid);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }

            var contentType = imagePart.ContentType;
            if (!ImageContentTypes.Contains(contentType))
                return null;

            using (var partStream = imagePart.GetStream())
            using (var memoryStream = new System.IO.MemoryStream())
            {
                partStream.CopyTo(memoryStream);
                var imageBytes = memoryStream.ToArray();
                using var bitmap = SKBitmap.Decode(imageBytes);

                if (extentCx != null && extentCy != null)
                {
                    var imageInfo = new ImageInfo()
                    {
                        Bitmap = bitmap,
                        ImageBytes = imageBytes,
                        ImgStyleAttribute = new XAttribute("style",
                            string.Format(NumberFormatInfo.InvariantInfo,
                                "width: {0}in; height: {1}in",
                                (float)extentCx / (float)ImageInfo.EmusPerInch,
                                (float)extentCy / (float)ImageInfo.EmusPerInch)),
                        ContentType = contentType,
                        DrawingElement = element,
                        AltText = altText,
                    };
                    var imgElement2 = imageHandler(imageInfo);
                    if (hyperlinkUri != null)
                    {
                        return new XElement(XhtmlNoNamespace.a,
                            new XAttribute(XhtmlNoNamespace.href, hyperlinkUri),
                            imgElement2);
                    }
                    return imgElement2;
                }

                var imageInfo2 = new ImageInfo()
                {
                    Bitmap = bitmap,
                    ImageBytes = imageBytes,
                    ContentType = contentType,
                    DrawingElement = element,
                    AltText = altText,
                };
                var imgElement = imageHandler(imageInfo2);
                if (hyperlinkUri != null)
                {
                    return new XElement(XhtmlNoNamespace.a,
                        new XAttribute(XhtmlNoNamespace.href, hyperlinkUri),
                        imgElement);
                }
                return imgElement;
            }
        }

        private static XElement ProcessPictureOrObject(WordprocessingDocument wordDoc,
            XElement element, Func<ImageInfo, XElement> imageHandler)
        {
            var imageRid = (string)element.Elements(VML.shape).Elements(VML.imagedata).Attributes(R.id).FirstOrDefault();
            if (imageRid == null) return null;

            try
            {
                var pp = wordDoc.MainDocumentPart.Parts.FirstOrDefault(pp2 => pp2.RelationshipId == imageRid);
                if (pp == null) return null;

                var imagePart = (ImagePart)pp.OpenXmlPart;
                if (imagePart == null) return null;

                var contentType = imagePart.ContentType;
                if (!ImageContentTypes.Contains(contentType))
                    return null;

                using (var partStream = imagePart.GetStream())
                using (var memoryStream = new System.IO.MemoryStream())
                {
                    try
                    {
                        partStream.CopyTo(memoryStream);
                        var imageBytes = memoryStream.ToArray();
                        using var bitmap = SKBitmap.Decode(imageBytes);

                        var imageInfo = new ImageInfo()
                        {
                            Bitmap = bitmap,
                            ImageBytes = imageBytes,
                            ContentType = contentType,
                            DrawingElement = element
                        };

                        var style = (string?)element.Elements(VML.shape).Attributes("style").FirstOrDefault();
                        if (style == null) return imageHandler(imageInfo);

                        var tokens = style.Split(';');
                        var widthInPoints = WidthInPoints(tokens);
                        var heightInPoints = HeightInPoints(tokens);
                        if (widthInPoints != null && heightInPoints != null)
                        {
                            imageInfo.ImgStyleAttribute = new XAttribute("style",
                                string.Format(NumberFormatInfo.InvariantInfo,
                                    "width: {0}pt; height: {1}pt", widthInPoints, heightInPoints));
                        }
                        return imageHandler(imageInfo);
                    }
                    catch (OutOfMemoryException)
                    {
                        // SKBitmap.Decode can throw OutOfMemoryException for corrupted images
                        return null;
                    }
                    catch (ArgumentException)
                    {
                        return null;
                    }
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private static float? HeightInPoints(IEnumerable<string> tokens)
        {
            return SizeInPoints(tokens, "height");
        }

        private static float? WidthInPoints(IEnumerable<string> tokens)
        {
            return SizeInPoints(tokens, "width");
        }

        private static float? SizeInPoints(IEnumerable<string> tokens, string name)
        {
            var sizeString = tokens
                .Select(t => new
                {
                    Name = t.Split(':').First(),
                    Value = t.Split(':').Skip(1).Take(1).FirstOrDefault()
                })
                .Where(p => p.Name == name)
                .Select(p => p.Value)
                .FirstOrDefault();

            if (sizeString != null &&
                sizeString.Length > 2 &&
                sizeString.Substring(sizeString.Length - 2) == "pt")
            {
                float size;
                if (float.TryParse(sizeString.Substring(0, sizeString.Length - 2), out size))
                    return size;
            }
            return null;
        }

        #endregion
    }

    public static class HtmlConverterExtensions
    {
        public static void AddIfMissing(this Dictionary<string, string> style, string propName, string value)
        {
            if (style.ContainsKey(propName))
                return;
            style.Add(propName, value);
        }
    }
}
