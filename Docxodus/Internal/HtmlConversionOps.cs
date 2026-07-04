#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Options for <see cref="HtmlConversionOps"/>. Mirrors the parameter set of the
/// WASM <c>DocumentConverter.ConvertDocxToHtmlComplete</c> shell so every surface
/// renders identically. Integer-coded modes match the existing WASM wire contract:
/// CommentRenderMode -1=disabled,0=Endnote,1=Inline,2=Margin;
/// PaginationMode 0=None,1=Paginated; AnnotationLabelMode 0=Above,1=Inline,2=Tooltip,3=None.
/// </summary>
internal sealed class HtmlConversionOptions
{
    public string PageTitle { get; init; } = "Document";
    public string CssClassPrefix { get; init; } = "docx-";
    public bool FabricateCssClasses { get; init; } = true;
    public string AdditionalCss { get; init; } = "";
    public int CommentRenderMode { get; init; } = -1;
    public string CommentCssClassPrefix { get; init; } = "comment-";
    public int PaginationMode { get; init; }
    public double PaginationScale { get; init; } = 1.0;
    public string PaginationCssClassPrefix { get; init; } = "page-";
    public bool RenderAnnotations { get; init; }
    public int AnnotationLabelMode { get; init; }
    public string AnnotationCssClassPrefix { get; init; } = "annot-";
    public bool RenderFootnotesAndEndnotes { get; init; }
    public bool RenderHeadersAndFooters { get; init; }
    public bool RenderTrackedChanges { get; init; }
    public bool ShowDeletedContent { get; init; } = true;
    public bool RenderMoveOperations { get; init; } = true;
    public bool RenderUnsupportedContentPlaceholders { get; init; }
    public string? DocumentLanguage { get; init; }

    /// <summary>
    /// When true, assign deterministic content-addressable Unids and stamp
    /// block-level HTML elements with <c>data-anchor</c> so the editor can address
    /// blocks in the DOM. Anchors match the markdown projector / DocxSession.
    /// </summary>
    public bool StampAnchors { get; init; }
}

/// <summary>
/// Single owner of the DOCX-bytes + <see cref="HtmlConversionOptions"/> →
/// HTML-string mapping. Both the WASM <c>DocumentConverter</c> bridge and the
/// stdio Python host route through here, so render behavior lives in one place.
/// Throws on invalid input; callers serialize errors at their boundary.
/// </summary>
internal static class HtmlConversionOps
{
    /// <summary>Render raw DOCX bytes to a self-contained HTML string.</summary>
    public static string ConvertToHtml(byte[] docxBytes, HtmlConversionOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(docxBytes));
        ArgumentNullException.ThrowIfNull(options);

        // Writable stream required: WmlToHtmlConverter runs RevisionAccepter internally.
        using var memoryStream = new MemoryStream();
        memoryStream.Write(docxBytes, 0, docxBytes.Length);
        memoryStream.Position = 0;
        using var wordDoc = WordprocessingDocument.Open(memoryStream, true);

        if (options.StampAnchors)
        {
            // Deterministic, content-addressable Unids — identical to the markdown
            // projector / DocxSession, so editor anchors line up across surfaces.
            UnidHelper.AssignToAllElementsDeterministic(wordDoc.MainDocumentPart!.GetXDocument().Root!);
        }

        var renderComments = options.CommentRenderMode >= 0;

        var settings = new WmlToHtmlConverterSettings
        {
            PageTitle = options.PageTitle,
            CssClassPrefix = options.CssClassPrefix,
            FabricateCssClasses = options.FabricateCssClasses,
            AdditionalCss = options.AdditionalCss,
            GeneralCss = "body { font-family: Arial, sans-serif; margin: 20px; } " +
                         "span { white-space: pre-wrap; }",
            RenderComments = renderComments,
            CommentRenderMode = renderComments
                ? (CommentRenderMode)options.CommentRenderMode
                : CommentRenderMode.EndnoteStyle,
            CommentCssClassPrefix = options.CommentCssClassPrefix,
            IncludeCommentMetadata = true,
            RenderPagination = (PaginationMode)options.PaginationMode,
            PaginationScale = options.PaginationScale > 0 ? options.PaginationScale : 1.0,
            PaginationCssClassPrefix = options.PaginationCssClassPrefix,
            RenderAnnotations = options.RenderAnnotations,
            AnnotationLabelMode = (AnnotationLabelMode)options.AnnotationLabelMode,
            AnnotationCssClassPrefix = options.AnnotationCssClassPrefix,
            IncludeAnnotationMetadata = true,
            RenderFootnotesAndEndnotes = options.RenderFootnotesAndEndnotes,
            RenderHeadersAndFooters = options.RenderHeadersAndFooters,
            RenderTrackedChanges = options.RenderTrackedChanges,
            ShowDeletedContent = options.ShowDeletedContent,
            RenderMoveOperations = options.RenderMoveOperations,
            IncludeRevisionMetadata = true,
            RenderUnsupportedContentPlaceholders = options.RenderUnsupportedContentPlaceholders,
            UnsupportedContentCssClassPrefix = "unsupported-",
            IncludeUnsupportedContentMetadata = true,
            DocumentLanguage = options.DocumentLanguage,
            StampAnchors = options.StampAnchors,
            // Embed images as base64 data URIs — no SkiaSharp needed (WASM-safe).
            ImageHandler = CreateBase64ImageHandler(),
        };

        var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
        return htmlElement.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Render a live session's current (possibly edited) state to HTML.</summary>
    public static string ConvertToHtml(DocxSession session, HtmlConversionOptions options)
    {
        if (session == null) throw new ArgumentNullException(nameof(session));
        return ConvertToHtml(session.Save(), options);
    }

    /// <summary>Render the session registered under <paramref name="handle"/> to HTML.</summary>
    public static string ConvertToHtml(int handle, HtmlConversionOptions options) =>
        ConvertToHtml(SessionRegistry.Get(handle), options);

    /// <summary>
    /// Render a single block (addressed by a <c>kind:scope:unid</c> anchor) to faithful
    /// HTML. Builds a throwaway document that copies the source's styles/numbering/theme
    /// parts and contains just the one block, then runs the standard converter. The full
    /// document render is the faithfulness oracle — this must match the corresponding
    /// <c>data-anchor</c> element from a full render. Known limits: a list item loses
    /// numbering continuation, and an inline image loses its (uncopied) image part.
    /// </summary>
    public static string RenderBlockHtml(byte[] docxBytes, string anchorId, HtmlConversionOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(docxBytes));
        if (string.IsNullOrWhiteSpace(anchorId))
            throw new ArgumentException("No anchor id provided", nameof(anchorId));
        ArgumentNullException.ThrowIfNull(options);

        using var sourceStream = new MemoryStream();
        sourceStream.Write(docxBytes, 0, docxBytes.Length);
        sourceStream.Position = 0;
        using var sourceDoc = WordprocessingDocument.Open(sourceStream, true);

        // Stateless path: no live session, so assign deterministic Unids here (the same
        // call the full render uses) so the anchor resolves by construction.
        UnidHelper.AssignToAllElementsDeterministic(sourceDoc.MainDocumentPart!.GetXDocument().Root!);

        var unid = AnchorUnid(anchorId);
        var blockElement = FindByUnid(sourceDoc, unid)
            ?? throw new ArgumentException($"anchor not found: {anchorId}", nameof(anchorId));
        return RenderResolvedBlock(sourceDoc, blockElement, options);
    }

    /// <summary>
    /// Session-attached single-block render. Resolves the block from the live session
    /// document WITHOUT re-opening bytes or re-assigning Unids over the whole document —
    /// the optimized path for an editor's incremental per-block re-render after an edit.
    /// Read-only with respect to the session (the block is cloned, parts are read).
    /// </summary>
    public static string RenderBlockHtml(DocxSession session, string anchorId, HtmlConversionOptions options)
    {
        if (session is null) throw new ArgumentNullException(nameof(session));
        if (string.IsNullOrWhiteSpace(anchorId))
            throw new ArgumentException("No anchor id provided", nameof(anchorId));
        ArgumentNullException.ThrowIfNull(options);

        var unid = AnchorUnid(anchorId);
        var liveDoc = session.LiveDocument;

        var blockElement = FindByUnid(liveDoc, unid);
        if (blockElement is null)
        {
            // Anchor not on the live tree yet — ensure Unids are assigned/persisted
            // (one projection) and retry once.
            session.Project();
            blockElement = FindByUnid(liveDoc, unid);
        }
        if (blockElement is null)
            throw new ArgumentException($"anchor not found: {anchorId}", nameof(anchorId));

        // Reuse a per-session formatting "shell" (the formatting parts + an empty body, serialized
        // once) so a keystroke commit doesn't re-clone the source's whole style gallery every render
        // — the dominant cost on a large gallery. The shell is rebuilt only when the formatting parts
        // actually change (signature), which only happens on a format op (add style / numbering /
        // level), never on a text edit, so it survives normal typing.
        long sig = ComputeFormattingSignature(liveDoc);
        if (session.RenderShellBytes is null || session.RenderShellSignature != sig)
        {
            session.RenderShellBytes = BuildShellDocBytes(liveDoc);
            session.RenderShellSignature = sig;
        }
        return RenderBlockFromShell(session.RenderShellBytes, blockElement, options);
    }

    /// <summary>Session-attached render for a registered session handle.</summary>
    public static string RenderBlockHtml(int handle, string anchorId, HtmlConversionOptions options) =>
        RenderBlockHtml(SessionRegistry.Get(handle), anchorId, options);

    private static string AnchorUnid(string anchorId) =>
        anchorId.Substring(anchorId.LastIndexOf(':') + 1);

    /// <summary>Find the element bearing PtOpenXml:Unid == unid across body/header/footer parts.</summary>
    private static XElement? FindByUnid(WordprocessingDocument doc, string unid)
    {
        var main = doc.MainDocumentPart;
        if (main is null) return null;
        bool Match(XElement e) => (string?)e.Attribute(PtOpenXml.Unid) == unid;

        var hit = main.GetXDocument().Root?.DescendantsAndSelf().FirstOrDefault(Match);
        if (hit != null) return hit;
        foreach (var part in main.HeaderParts.Cast<OpenXmlPart>().Concat(main.FooterParts))
        {
            hit = part.GetXDocument().Root?.DescendantsAndSelf().FirstOrDefault(Match);
            if (hit != null) return hit;
        }
        return null;
    }

    /// <summary>
    /// Render one resolved block element to HTML via a throwaway document that copies the
    /// source's formatting parts. Read-only w.r.t. <paramref name="sourceDoc"/> (the block is
    /// cloned, parts are read), so it is safe to call on a live session document. This is the
    /// STATELESS path (no per-session shell cache); the session-attached overload reuses a cached
    /// shell via <see cref="RenderBlockFromShell"/>.
    /// </summary>
    private static string RenderResolvedBlock(WordprocessingDocument sourceDoc, XElement blockElement,
        HtmlConversionOptions options)
    {
        var unid = (string?)blockElement.Attribute(PtOpenXml.Unid);

        // Build a throwaway doc: copied formatting parts + just this block.
        using var blockStream = new MemoryStream();
        using (var blockDoc = WordprocessingDocument.Create(
                   blockStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = blockDoc.AddMainDocumentPart();
            AddFormattingParts(blockDoc, sourceDoc);
            main.PutXDocument(BuildBodyDocument(new XElement(blockElement)));
        }
        blockStream.Position = 0;
        using var renderDoc = WordprocessingDocument.Open(blockStream, true);
        var htmlElement = WmlToHtmlConverter.ConvertToHtml(renderDoc, BuildBlockConverterSettings(options));
        return ExtractBlockHtml(htmlElement, unid);
    }

    /// <summary>
    /// Build the reusable per-session "shell": a serialized throwaway .docx holding the copied
    /// formatting parts and an EMPTY body. Built once (per formatting signature) and cached on the
    /// session; <see cref="RenderBlockFromShell"/> drops the block into its body per render. This
    /// front-loads the (expensive on a large style gallery) part clone+serialize so it is paid once
    /// rather than every keystroke commit.
    /// </summary>
    private static byte[] BuildShellDocBytes(WordprocessingDocument sourceDoc)
    {
        using var shellStream = new MemoryStream();
        using (var shellDoc = WordprocessingDocument.Create(
                   shellStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = shellDoc.AddMainDocumentPart();
            AddFormattingParts(shellDoc, sourceDoc);
            main.PutXDocument(BuildBodyDocument(/* empty body */));
        }
        return shellStream.ToArray();
    }

    /// <summary>
    /// Render one block from a cached shell: open a fresh copy of <paramref name="shellBytes"/>
    /// (so the converter's in-place mutation never touches the cache), drop the cloned block into
    /// the empty body, convert, and extract the block's HTML. Output is identical to
    /// <see cref="RenderResolvedBlock"/> for the same block + parts.
    /// </summary>
    private static string RenderBlockFromShell(byte[] shellBytes, XElement blockElement,
        HtmlConversionOptions options)
    {
        var unid = (string?)blockElement.Attribute(PtOpenXml.Unid);
        using var ms = new MemoryStream();
        ms.Write(shellBytes, 0, shellBytes.Length);
        ms.Position = 0;
        using var renderDoc = WordprocessingDocument.Open(ms, true);
        var bodyEl = renderDoc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        bodyEl.RemoveNodes();
        bodyEl.Add(new XElement(blockElement));
        var htmlElement = WmlToHtmlConverter.ConvertToHtml(renderDoc, BuildBlockConverterSettings(options));
        return ExtractBlockHtml(htmlElement, unid);
    }

    /// <summary>
    /// Cheap content signature of the formatting parts that affect a block render. It changes when a
    /// format op adds a style / numbering / level (the only mid-session formatting-part mutations —
    /// see DocxSession's StyleFactory / NumberingFactory call sites); text edits never touch these
    /// parts. Computed from the already-parsed (cached) XDocuments, so it is ~microseconds and
    /// reflects in-memory mutations regardless of stream flush. NOTE: the edit API never mutates the
    /// theme / fontTable / settings parts, so they are not part of the signature; if that ever
    /// changes, add them here (or the cached shell could go stale).
    /// </summary>
    private static long ComputeFormattingSignature(WordprocessingDocument doc)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var main = doc.MainDocumentPart;
        if (main is null) return 0;
        long sig = 17;
        void Mix(long v) => sig = unchecked(sig * 1000003 + v);

        var styles = main.StyleDefinitionsPart?.GetXDocument().Root;
        Mix(styles?.Elements(w + "style").Count() ?? -1);
        var swe = main.StylesWithEffectsPart?.GetXDocument().Root;
        Mix(swe?.Elements(w + "style").Count() ?? -1);
        var num = main.NumberingDefinitionsPart?.GetXDocument().Root;
        Mix(num?.Elements(w + "num").Count() ?? -1);
        Mix(num?.Elements(w + "abstractNum").Count() ?? -1);
        Mix(num?.Descendants(w + "lvl").Count() ?? -1);
        return sig;
    }

    /// <summary>Copy the formatting parts (styles/numbering/theme/font/settings) from src into the
    /// throwaway doc, ensuring a DocumentSettingsPart exists (the converter reads w:defaultTabStop
    /// with no null check).</summary>
    private static void AddFormattingParts(WordprocessingDocument blockDoc, WordprocessingDocument sourceDoc)
    {
        CopyPartXml(sourceDoc, blockDoc, p => p.StyleDefinitionsPart);
        CopyPartXml(sourceDoc, blockDoc, p => p.StylesWithEffectsPart);
        CopyPartXml(sourceDoc, blockDoc, p => p.NumberingDefinitionsPart);
        CopyPartXml(sourceDoc, blockDoc, p => p.ThemePart);
        CopyPartXml(sourceDoc, blockDoc, p => p.FontTablePart);
        CopyPartXml(sourceDoc, blockDoc, p => p.DocumentSettingsPart);
        if (blockDoc.MainDocumentPart!.DocumentSettingsPart is null)
        {
            blockDoc.MainDocumentPart.AddNewPart<DocumentSettingsPart>()
                .PutXDocument(new XDocument(
                    new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
        }
    }

    /// <summary>A minimal <c>w:document</c> wrapping <paramref name="bodyContent"/> (or an empty body).</summary>
    private static XDocument BuildBodyDocument(params object[] bodyContent) =>
        new XDocument(
            new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "r", R.r),
                new XElement(W.body, bodyContent)));

    private static WmlToHtmlConverterSettings BuildBlockConverterSettings(HtmlConversionOptions options) =>
        new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = options.FabricateCssClasses,
            CssClassPrefix = options.CssClassPrefix,
            StampAnchors = true,
            // The throwaway doc copies the source's (possibly huge) style gallery verbatim;
            // re-simplifying it every render is the dominant single-block cost (~70ms on a 160-style
            // python-docx doc) and only strips rsids, which never reach the HTML. Skip it — the
            // resolved formatting, and thus the rendered block, are identical to the full render.
            SkipFormattingPartsSimplification = true,
        };

    /// <summary>Extract the rendered block (located by its stamped data-anchor) from the full
    /// converter output, not the <c>&lt;html&gt;</c> wrapper.</summary>
    private static string ExtractBlockHtml(XElement htmlElement, string? unid)
    {
        XElement? inner = null;
        if (unid != null)
            inner = htmlElement.Descendants().FirstOrDefault(e => (string?)e.Attribute("data-anchor") == unid);
        if (inner is null)
        {
            var body = htmlElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "body");
            inner = body?.Elements().FirstOrDefault() ?? htmlElement;
        }
        return inner.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Clone a whole formatting part (styles/numbering/theme/font) from src to dst.</summary>
    private static void CopyPartXml<TPart>(WordprocessingDocument src, WordprocessingDocument dst,
        Func<MainDocumentPart, TPart?> get) where TPart : OpenXmlPart, IFixedContentTypePart
    {
        var srcPart = get(src.MainDocumentPart!);
        if (srcPart is null) return;
        var srcRoot = srcPart.GetXDocument().Root;
        if (srcRoot is null) return;
        var dstPart = dst.MainDocumentPart!.AddNewPart<TPart>();
        dstPart.PutXDocument(new XDocument(new XElement(srcRoot)));
    }

    private static Func<ImageInfo, XElement> CreateBase64ImageHandler()
    {
        return imageInfo =>
        {
            if (imageInfo.ImageBytes == null || imageInfo.ImageBytes.Length == 0)
                return null!;

            var mimeType = imageInfo.ContentType ?? "image/png";
            var base64 = Convert.ToBase64String(imageInfo.ImageBytes);
            var dataUri = $"data:{mimeType};base64,{base64}";

            var imgElement = new XElement(XhtmlNoNamespace.img,
                new XAttribute("src", dataUri));

            if (imageInfo.ImgStyleAttribute != null)
                imgElement.Add(imageInfo.ImgStyleAttribute);

            if (!string.IsNullOrEmpty(imageInfo.AltText))
                imgElement.Add(new XAttribute("alt", imageInfo.AltText));

            return imgElement;
        };
    }
}
