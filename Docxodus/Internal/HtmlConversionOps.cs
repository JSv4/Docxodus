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
    /// <summary>
    /// Assign the deterministic, content-addressable Unids that editor anchors are derived from —
    /// identical to the markdown projector / <see cref="DocxSession"/>, so anchors line up across
    /// surfaces. Covers the main document part <em>and</em> the footnotes/endnotes parts: note
    /// bodies render as ordinary blocks in the footnotes section, so their paragraphs need
    /// <c>data-anchor</c> stamps exactly like body paragraphs — without them a rendered footnote is
    /// visible but not addressable, so an editor can show it and not edit it.
    /// </summary>
    private static void AssignAnchorUnids(WordprocessingDocument doc)
    {
        _ = WmlToMarkdownConverter.BuildAnchorIndexOnly(doc,
            new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.All });
    }

    /// <summary>Render raw DOCX bytes to a self-contained HTML string.</summary>
    public static string ConvertToHtml(byte[] docxBytes, HtmlConversionOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(docxBytes));
        ArgumentNullException.ThrowIfNull(options);

        // WmlToHtmlConverter's XDocument pipeline is transitional-namespace based. Match
        // DocxDiff's open behavior so Word Strict packages render rather than presenting an
        // unrecognized purl.oclc.org body to the converter.
        docxBytes = StrictOoxmlNormalizer.NormalizeToTransitional(docxBytes);

        // Writable stream required: WmlToHtmlConverter runs RevisionAccepter internally.
        using var memoryStream = new MemoryStream();
        memoryStream.Write(docxBytes, 0, docxBytes.Length);
        memoryStream.Position = 0;
        using var wordDoc = WordprocessingDocument.Open(memoryStream, true);

        var renderComments = options.CommentRenderMode >= 0;
        bool renderPagination = options.PaginationMode == (int)PaginationMode.Paginated;
        // PageMap source identity is required for every paginated render, including the
        // stateless viewer's default stampAnchors=false path. Bare editable data-anchor stamps
        // remain opt-in; canonical kind:scope:unid identity does not.
        if (options.StampAnchors || renderPagination)
            AssignAnchorUnids(wordDoc);
        // The paginated React viewer injects this document HTML into its capture host. A body margin
        // therefore applies to the HOST body as well as the document staging tree and makes every fixed-size
        // page box overflow onto a second printed page. Keep the comfortable standalone-document margin, but
        // leave the host flush when pagination owns page geometry.
        string bodyMargin = renderPagination ? "0" : "20px";

        var settings = new WmlToHtmlConverterSettings
        {
            PageTitle = options.PageTitle,
            CssClassPrefix = options.CssClassPrefix,
            FabricateCssClasses = options.FabricateCssClasses,
            AdditionalCss = options.AdditionalCss,
            GeneralCss = $"body {{ font-family: Arial, sans-serif; margin: {bodyMargin}; }} " +
                         "span { white-space: pre-wrap; }",
            RenderComments = renderComments,
            CommentRenderMode = renderComments
                ? (CommentRenderMode)options.CommentRenderMode
                : CommentRenderMode.EndnoteStyle,
            CommentCssClassPrefix = options.CommentCssClassPrefix,
            IncludeCommentMetadata = true,
            RenderPagination = (PaginationMode)options.PaginationMode,
            // Paginated output is printed from fixed-size page boxes. Emit physical page CSS for
            // every paginated document: uniform sections use one global rule; mixed sections use
            // named rules selected by each rendered page box.
            GeneratePageCss = renderPagination,
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
            StampCanonicalSourceAnchors = options.StampAnchors || renderPagination,
            // Embed images as base64 data URIs — no SkiaSharp needed (WASM-safe).
            ImageHandler = CreateBase64ImageHandler(),
        };

        var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
        return htmlElement.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Render a live session's current (possibly edited) state to HTML.</summary>
    /// <remarks>
    /// Serializes with <c>persistAnchorIds: true</c> REGARDLESS of the session's setting. These
    /// bytes are an internal hop that is rendered and discarded, and the anchors stamped into the
    /// HTML have to address the LIVE session: a Unid is content-hashed, so stripping it here would
    /// make the converter re-derive a fresh one for any block edited since the session assigned
    /// its id, and the rendered <c>data-anchor</c> would no longer resolve — the editor sees that
    /// as a block that has silently stopped being editable.
    ///
    /// That coupling is why the browser editor used to open its session with
    /// <c>PersistAnchorIds</c>, which then leaked the bookkeeping into the file its users
    /// downloaded. Pinning it to the RENDER, where the requirement actually lives, lets a save
    /// mean what a save should mean.
    /// </remarks>
    public static string ConvertToHtml(DocxSession session, HtmlConversionOptions options)
    {
        if (session == null) throw new ArgumentNullException(nameof(session));
        return ConvertToHtml(session.Save(persistAnchorIds: true), options);
    }

    /// <summary>Render the session registered under <paramref name="handle"/> to HTML.</summary>
    public static string ConvertToHtml(int handle, HtmlConversionOptions options) =>
        ConvertToHtml(SessionRegistry.Get(handle), options);

    /// <summary>
    /// The single definition of the option profile a mutation-batch preview renders with
    /// (<see cref="MutationPreviewHtmlMode.Full"/>). A preview answers "what would the document
    /// become", so it shows everything the applied document would carry — tracked changes,
    /// comments, annotations, notes, headers/footers — rather than the editor's authoring view.
    /// Every surface MUST consume this rather than restating the flags: the typed core, the
    /// handle façade, both bridges and both clients must agree about what a preview shows, or
    /// two callers previewing the same batch see materially different documents.
    /// </summary>
    public static HtmlConversionOptions PreviewDocumentOptions() => new()
    {
        CommentRenderMode = 0,
        RenderAnnotations = true,
        RenderFootnotesAndEndnotes = true,
        RenderHeadersAndFooters = true,
        RenderTrackedChanges = true,
        StampAnchors = true,
    };

    /// <summary>
    /// The single definition of the option profile a scoped
    /// (<see cref="MutationPreviewHtmlMode.Scoped"/>) preview renders one block with. Tracked
    /// changes stay on for the same reason as <see cref="PreviewDocumentOptions"/>: a scoped
    /// redline preview that hides its own redlines shows the caller nothing.
    /// </summary>
    public static HtmlConversionOptions PreviewBlockOptions() => new()
    {
        RenderTrackedChanges = true,
        RenderFootnotesAndEndnotes = true,
        StampAnchors = true,
    };

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

        docxBytes = StrictOoxmlNormalizer.NormalizeToTransitional(docxBytes);

        using var sourceStream = new MemoryStream();
        sourceStream.Write(docxBytes, 0, docxBytes.Length);
        sourceStream.Position = 0;
        using var sourceDoc = WordprocessingDocument.Open(sourceStream, true);

        // Stateless path: no live session, so assign deterministic Unids here (the same
        // call the full render uses) so the anchor resolves by construction — note parts
        // included, else a footnote-paragraph anchor could never resolve on this path.
        AssignAnchorUnids(sourceDoc);

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

        // Single-block render IS a one-element batch — one owner for anchor resolution,
        // neighbor context (contextualSpacing) and list-annotation transplant (true
        // marker numbers in isolation), so every incremental swap path shares them.
        var map = RenderBlocksCore(session, new[] { anchorId }, options);
        return map[anchorId] ?? throw new ArgumentException($"anchor not found: {anchorId}", nameof(anchorId));
    }

    /// <summary>
    /// Batch session-attached block render: N anchors, ONE throwaway document, one
    /// converter run — the per-call shell setup that dominates single-block renders is
    /// paid once. Returns a JSON object mapping each anchor id to its HTML element
    /// (<c>null</c> for an anchor that fails to resolve — callers fall back per block).
    /// Each rendered block matches the corresponding <c>data-anchor</c> element of a
    /// full render with the same options: real siblings are cloned around each target
    /// (so <c>w:contextualSpacing</c> margins resolve) and the live document's
    /// list-numbering annotations ride along on the clones (so a list item deep in a
    /// list renders its true number, not "1."). A <c>fn:</c>/<c>en:</c> anchor renders
    /// as the concatenation of the note's paragraphs (no list-item wrapper — that is
    /// notes-section chrome the client owns).
    /// </summary>
    public static string RenderBlocksHtml(DocxSession session, IReadOnlyList<string> anchorIds, HtmlConversionOptions options)
    {
        if (session is null) throw new ArgumentNullException(nameof(session));
        ArgumentNullException.ThrowIfNull(anchorIds);
        ArgumentNullException.ThrowIfNull(options);

        var map = RenderBlocksCore(session, anchorIds, options);
        var sb = new System.Text.StringBuilder(256);
        sb.Append('{');
        bool first = true;
        foreach (var id in anchorIds)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(DocxSessionJson.JsonString(id)).Append(':');
            sb.Append(map.TryGetValue(id, out var html) && html is not null
                ? DocxSessionJson.JsonString(html)
                : "null");
        }
        sb.Append('}');
        return sb.ToString();
    }

    /// <summary>Batch overload for a registered session handle (anchor ids as a JSON string array).</summary>
    public static string RenderBlocksHtml(int handle, string anchorIdsJson, HtmlConversionOptions options)
    {
        var ids = new List<string>();
        using (var doc = System.Text.Json.JsonDocument.Parse(anchorIdsJson))
        {
            foreach (var e in doc.RootElement.EnumerateArray())
            {
                if (e.GetString() is { } s) ids.Add(s);
            }
        }
        return RenderBlocksHtml(SessionRegistry.Get(handle), ids, options);
    }

    private static Dictionary<string, string?> RenderBlocksCore(
        DocxSession session, IReadOnlyList<string> anchorIds, HtmlConversionOptions options)
    {
        var liveDoc = session.LiveDocument;
        EnsureListAnnotations(liveDoc);

        var results = new Dictionary<string, string?>(StringComparer.Ordinal);
        // Block-level render targets: the element to place in the throwaway body, keyed
        // by the unid its HTML is extracted under.
        var targets = new List<XElement>();
        // fn/en anchors expand to their child paragraphs; remember which unids to
        // concatenate back per note anchor.
        var noteChildUnids = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        foreach (var anchorId in anchorIds.Distinct(StringComparer.Ordinal))
        {
            var el = ResolveSessionAnchor(session, anchorId);
            if (el is null) { results[anchorId] = null; continue; }
            if (el.Name == W.footnote || el.Name == W.endnote)
            {
                var childUnids = new List<string>();
                foreach (var p in el.Elements(W.p))
                {
                    var u = (string?)p.Attribute(PtOpenXml.Unid);
                    if (u is null) continue;
                    childUnids.Add(u);
                    targets.Add(p);
                }
                noteChildUnids[anchorId] = childUnids;
            }
            else
            {
                targets.Add(el);
            }
        }

        if (targets.Count > 0)
        {
            var htmlByUnid = RenderTargetsFromShell(session, liveDoc, targets, options);
            foreach (var anchorId in anchorIds)
            {
                if (results.ContainsKey(anchorId)) continue; // resolution failure already recorded
                if (noteChildUnids.TryGetValue(anchorId, out var childUnids))
                {
                    var parts = new List<string>(childUnids.Count);
                    foreach (var u in childUnids)
                    {
                        if (htmlByUnid.TryGetValue(u, out var h) && h is not null) parts.Add(h);
                    }
                    results[anchorId] = parts.Count > 0 ? string.Concat(parts) : null;
                }
                else
                {
                    var unid = AnchorUnid(anchorId);
                    results[anchorId] = htmlByUnid.TryGetValue(unid, out var h) ? h : null;
                }
            }
        }
        return results;
    }

    /// <summary>
    /// Resolve an anchor against the live session: index-first (it knows which PART the
    /// anchor lives in — content-addressed unids collide across parts, e.g. empty
    /// default/first/even header stories), unid-scan fallback, one projection retry for
    /// anchors not yet on the live tree.
    /// </summary>
    /// <remarks>
    /// When the session's anchor-index cache is COLD (the preceding mutation invalidated
    /// it — i.e. on every single-block re-render an editor performs), a body-scope anchor
    /// resolves by a direct unid walk of the main part instead: the walk costs a fraction
    /// of the whole-document index rebuild <see cref="DocxSession.FindAnchor"/> would
    /// trigger, and for main-part elements it is exactly as authoritative (a body/unid
    /// collision resolves to the main part in both). Non-body scopes keep the index path —
    /// only the index knows which header/footer part owns a colliding unid.
    /// </remarks>
    private static XElement? ResolveSessionAnchor(DocxSession session, string anchorId)
    {
        if (string.IsNullOrWhiteSpace(anchorId)) return null;
        var liveDoc = session.LiveDocument;
        var unid = AnchorUnid(anchorId);

        XElement? el = null;
        if (!session.HasCachedAnchorIndex && AnchorScope(anchorId) == "body")
        {
            bool Match(XElement e) => (string?)e.Attribute(PtOpenXml.Unid) == unid;
            el = liveDoc.MainDocumentPart?.GetXDocument().Root?.DescendantsAndSelf().FirstOrDefault(Match);
        }

        el ??= session.FindAnchor(anchorId)?.Resolve(liveDoc);
        el ??= FindByUnid(liveDoc, unid);
        if (el is null)
        {
            session.Project();
            el = session.FindAnchor(anchorId)?.Resolve(liveDoc) ?? FindByUnid(liveDoc, unid);
        }
        return el;
    }

    /// <summary>The scope segment of a <c>kind:scope:unid</c> anchor id ("" when malformed).</summary>
    private static string AnchorScope(string anchorId)
    {
        int first = anchorId.IndexOf(':');
        int last = anchorId.LastIndexOf(':');
        return first >= 0 && last > first ? anchorId.Substring(first + 1, last - first - 1) : "";
    }

    /// <summary>
    /// Render the target block elements through ONE fresh copy of the session's cached
    /// shell and return each target's HTML keyed by its unid. Targets are grouped per
    /// parent into contiguous sibling runs padded with one real neighbor on each side,
    /// so <c>w:contextualSpacing</c> resolves exactly as in the full render; only
    /// requested unids are extracted (context clones are scaffolding).
    /// </summary>
    private static Dictionary<string, string?> RenderTargetsFromShell(
        DocxSession session, WordprocessingDocument liveDoc, List<XElement> targets, HtmlConversionOptions options)
    {
        long sig = ComputeFormattingSignature(liveDoc);
        if (session.RenderShellDoc is null || session.RenderShellSignature != sig)
        {
            session.DisposeRenderShell();
            var shellBytes = BuildShellDocBytes(liveDoc);
            var shellStream = new MemoryStream();
            shellStream.Write(shellBytes, 0, shellBytes.Length);
            shellStream.Position = 0;
            session.RenderShellStream = shellStream;
            session.RenderShellDoc = WordprocessingDocument.Open(shellStream, true);
            session.RenderShellSignature = sig;
        }

        var wantedUnids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var t in targets)
        {
            if ((string?)t.Attribute(PtOpenXml.Unid) is { } u) wantedUnids.Add(u);
        }

        // The session's own anchor index IS the identity the caller addressed these blocks by,
        // and it is cached, so repeated stamped renders do not rebuild it.
        var identity = BlockSourceIdentity.For(options.StampAnchors, session.AnchorIndex(), liveDoc);

        // Per parent: order block-level children, merge each target's ±1 window into runs.
        var bodyContent = new List<XElement>();
        foreach (var parentGroup in targets.Where(t => t.Parent is not null).GroupBy(t => t.Parent!))
        {
            var siblings = parentGroup.Key.Elements()
                .Where(e => e.Name == W.p || e.Name == W.tbl)
                .ToList();
            var posOf = new Dictionary<XElement, int>();
            for (int i = 0; i < siblings.Count; i++) posOf[siblings[i]] = i;

            var positions = parentGroup.Select(t => posOf.TryGetValue(t, out var p) ? p : -1)
                .Where(p => p >= 0).Distinct().OrderBy(p => p).ToList();
            int idx = 0;
            while (idx < positions.Count)
            {
                int start = Math.Max(0, positions[idx] - 1);
                int end = Math.Min(siblings.Count - 1, positions[idx] + 1);
                while (idx + 1 < positions.Count && positions[idx + 1] - 1 <= end + 1)
                {
                    idx++;
                    end = Math.Min(siblings.Count - 1, positions[idx] + 1);
                }
                idx++;
                for (int i = start; i <= end; i++)
                {
                    var clone = CloneWithListAnnotations(siblings[i]);
                    identity?.Record(siblings[i], clone);
                    bodyContent.Add(clone);
                }
            }
        }

        var htmlByUnid = new Dictionary<string, string?>(StringComparer.Ordinal);
        if (bodyContent.Count == 0) return htmlByUnid;

        // Render through the OPEN shell. Each render replaces the main part's XDocument with a
        // brand-new body document (PutXDocument swaps the cached XDocument annotation), so the
        // converter's per-render root annotations (comment/footnote trackers, section info, field
        // info) can never leak from one render into the next — while the formatting-part
        // XDocuments, and the style/numbering resolution caches the converter annotates onto
        // them, persist across renders. That persistence is the point: re-opening the shell per
        // render paid the package open + styles/numbering parse + cache rebuild on every
        // keystroke commit, and it dominated single-block render time.
        var renderDoc = session.RenderShellDoc;
        renderDoc.MainDocumentPart!.PutXDocument(
            BuildBodyDocument(bodyContent.Cast<object>().ToArray()));

        var blockSettings = BuildBlockConverterSettings(options);
        if (identity is not null) blockSettings.SourceAnchorIdentityProvider = identity.Resolve;

        var htmlElement = WmlToHtmlConverter.ConvertToHtml(renderDoc, blockSettings);
        foreach (var e in htmlElement.Descendants())
        {
            var u = (string?)e.Attribute("data-anchor");
            if (u is null || !wantedUnids.Contains(u) || htmlByUnid.ContainsKey(u)) continue;
            // A table always renders inside a generated single-child alignment <div>
            // (see the converter's tableDiv). Return that wrapper so an incremental
            // renderer inserts the same node shape a full render produces — a bare
            // <table> would lose alignment and leave wrapper husks on replace. Only
            // tables: paragraph wrappers (border <div>s) can GROUP several blocks, so
            // they are the client's remount-fallback territory, not extraction chrome.
            var outer = e;
            if (e.Name.LocalName == "table"
                && e.Parent is { } p
                && p.Name.LocalName == "div"
                && p.Attribute("data-anchor") is null
                && p.Elements().Count() == 1)
            {
                outer = p;
            }
            htmlByUnid[u] = outer.ToString(SaveOptions.DisableFormatting);
        }
        return htmlByUnid;
    }

    /// <summary>
    /// Make sure the live document carries <see cref="ListItemRetriever"/> annotations
    /// (per-paragraph <c>ListItemInfo</c> + per-item <c>LevelNumbers</c> counter
    /// vectors). One <see cref="ListItemRetriever.RetrieveListItem(WordprocessingDocument, XElement)"/>
    /// call on any unannotated paragraph initializes every content part.
    /// </summary>
    private static void EnsureListAnnotations(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main?.NumberingDefinitionsPart is null || main.StyleDefinitionsPart is null) return;
        var missing = main.GetXDocument().Root?
            .Descendants(W.p)
            .FirstOrDefault(p => p.Annotation<ListItemRetriever.ListItemInfo>() is null);
        if (missing is not null) ListItemRetriever.RetrieveListItem(doc, missing);
    }

    /// <summary>
    /// Clone a block element and transplant the LIVE document's list-numbering
    /// annotations onto the clone's paragraphs (XElement cloning drops annotations).
    /// The throwaway converter then reads the live counters instead of recomputing
    /// them from the throwaway's tiny body — where every list item would count from 1.
    /// Annotation reads are first-added-wins, so even if the converter re-initializes
    /// the throwaway document, the transplanted values hold.
    /// </summary>
    private static XElement CloneWithListAnnotations(XElement src)
    {
        var clone = new XElement(src);
        using var s = src.DescendantsAndSelf().GetEnumerator();
        using var c = clone.DescendantsAndSelf().GetEnumerator();
        while (s.MoveNext() && c.MoveNext())
        {
            if (s.Current.Name != W.p) continue;
            if (s.Current.Annotation<ListItemRetriever.ListItemInfo>() is { } lii) c.Current.AddAnnotation(lii);
            if (s.Current.Annotation<ListItemRetriever.LevelNumbers>() is { } ln) c.Current.AddAnnotation(ln);
            if (s.Current.Annotation<ListItemRetriever.ContinuationInfo>() is { } ci) c.Current.AddAnnotation(ci);
        }
        return clone;
    }

    /// <summary>
    /// Carries canonical <c>data-source-anchor-id</c> provenance from the ORIGINAL package onto
    /// the throwaway shell's clones.
    /// </summary>
    /// <remarks>
    /// The full render builds that identity by indexing the document it is converting
    /// (<see cref="WmlToHtmlConverter"/>'s <c>StampCanonicalSourceAnchors</c> block). A block
    /// shell must NOT do the same: its body is not the source's body — <c>fn:</c>/<c>en:</c>
    /// note paragraphs are hoisted into it, so a shell-derived index would resolve them to the
    /// <c>body</c> scope and stamp a confidently WRONG id onto the very attribute
    /// <c>npm/src/pagination.ts</c> resolves citations by. The scope is not inferred here, it is
    /// carried: the caller already knows which live element produced each clone.
    ///
    /// Two tiers, mirroring the full render's own provider: clone object identity first, then a
    /// Unid fallback for the case where converter preprocessing rebuilds an element (attributes,
    /// including <c>PtOpenXml:Unid</c>, ride along — that is what <c>data-anchor</c> depends on).
    /// Content-addressed Unids can collide across parts, so an ambiguous Unid stamps NOTHING
    /// rather than the wrong story's id.
    /// </remarks>
    private sealed class BlockSourceIdentity
    {
        private readonly Dictionary<XElement, string> _liveIdByElement;
        private readonly Dictionary<XElement, string> _idByClone = new();
        private readonly Dictionary<string, string?> _idByUnid = new(StringComparer.Ordinal);

        private BlockSourceIdentity(Dictionary<XElement, string> liveIdByElement) =>
            _liveIdByElement = liveIdByElement;

        /// <summary>
        /// Build a carrier over <paramref name="index"/>, or null when the caller did not ask
        /// for anchor stamping (the editor's incremental swap path) — the index walk is not
        /// paid for a render that will not use it.
        /// </summary>
        public static BlockSourceIdentity? For(
            bool enabled, IReadOnlyDictionary<string, AnchorTarget> index, WordprocessingDocument doc)
        {
            if (!enabled) return null;
            var liveIdByElement = new Dictionary<XElement, string>();
            foreach (var target in index.Values)
            {
                var source = target.Resolve(doc);
                if (source is not null) liveIdByElement[source] = target.Anchor.Id;
            }
            return new BlockSourceIdentity(liveIdByElement);
        }

        /// <summary>Record the source-to-clone correspondence for one cloned block subtree.</summary>
        public void Record(XElement source, XElement clone)
        {
            using var s = source.DescendantsAndSelf().GetEnumerator();
            using var c = clone.DescendantsAndSelf().GetEnumerator();
            while (s.MoveNext() && c.MoveNext())
            {
                if (!_liveIdByElement.TryGetValue(s.Current, out var id)) continue;
                _idByClone[c.Current] = id;
                if ((string?)c.Current.Attribute(PtOpenXml.Unid) is not { Length: > 0 } unid) continue;
                if (_idByUnid.TryGetValue(unid, out var existing))
                {
                    if (!string.Equals(existing, id, StringComparison.Ordinal)) _idByUnid[unid] = null;
                }
                else
                {
                    _idByUnid[unid] = id;
                }
            }
        }

        /// <summary>The <see cref="WmlToHtmlConverterSettings.SourceAnchorIdentityProvider"/>.</summary>
        public string? Resolve(XElement element)
        {
            if (_idByClone.TryGetValue(element, out var id)) return id;
            return (string?)element.Attribute(PtOpenXml.Unid) is { Length: > 0 } unid
                && _idByUnid.TryGetValue(unid, out var byUnid)
                    ? byUnid
                    : null;
        }
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
        // Peer stories that can own an addressable block: header/footer parts, and the
        // footnotes/endnotes parts (note bodies render as editable blocks in the notes section).
        foreach (var part in NoteAndStoryParts(main))
        {
            hit = part.GetXDocument().Root?.DescendantsAndSelf().FirstOrDefault(Match);
            if (hit != null) return hit;
        }
        return null;
    }

    /// <summary>Non-main parts that can own a block an anchor addresses.</summary>
    private static IEnumerable<OpenXmlPart> NoteAndStoryParts(MainDocumentPart main)
    {
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }

    /// <summary>
    /// Render one resolved block element to HTML via a throwaway document that copies the
    /// source's formatting parts. Read-only w.r.t. <paramref name="sourceDoc"/> (the block is
    /// cloned, parts are read), so it is safe to call on a live session document. This is the
    /// STATELESS path (no per-session shell cache); the session-attached paths batch through
    /// <see cref="RenderTargetsFromShell"/> instead.
    /// </summary>
    private static string RenderResolvedBlock(WordprocessingDocument sourceDoc, XElement blockElement,
        HtmlConversionOptions options)
    {
        var unid = (string?)blockElement.Attribute(PtOpenXml.Unid);

        // Same contract as the batch path: canonical source provenance is carried from the
        // source package, never re-derived from the throwaway body.
        var identity = BlockSourceIdentity.For(
            options.StampAnchors,
            WmlToMarkdownConverter.BuildAnchorIndexOnly(
                sourceDoc, new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.All }),
            sourceDoc);
        var blockClone = new XElement(blockElement);
        identity?.Record(blockElement, blockClone);

        // Build a throwaway doc: copied formatting parts + just this block.
        using var blockStream = new MemoryStream();
        using (var blockDoc = WordprocessingDocument.Create(
                   blockStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = blockDoc.AddMainDocumentPart();
            AddFormattingParts(blockDoc, sourceDoc);
            main.PutXDocument(BuildBodyDocument(blockClone));
        }
        blockStream.Position = 0;
        using var renderDoc = WordprocessingDocument.Open(blockStream, true);
        var blockSettings = BuildBlockConverterSettings(options);
        if (identity is not null) blockSettings.SourceAnchorIdentityProvider = identity.Resolve;
        var htmlElement = WmlToHtmlConverter.ConvertToHtml(renderDoc, blockSettings);
        return ExtractBlockHtml(htmlElement, unid);
    }

    /// <summary>
    /// Build the reusable per-session "shell": a serialized throwaway .docx holding the copied
    /// formatting parts and an EMPTY body. Built once (per formatting signature); the caller
    /// opens it and keeps it OPEN on the session (<see cref="DocxSession.RenderShellDoc"/>), and
    /// <see cref="RenderTargetsFromShell"/> replaces its main-part body document per render.
    /// This front-loads the part clone+serialize AND the package open + formatting-part XML
    /// parse (both expensive on a large style gallery) so they are paid once rather than on
    /// every keystroke commit.
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
    /// throwaway doc. Settings may be absent; ConvertToHtml defaults tab stop to 720 twips.</summary>
    private static void AddFormattingParts(WordprocessingDocument blockDoc, WordprocessingDocument sourceDoc)
    {
        CopyPartXml(sourceDoc, blockDoc, p => p.StyleDefinitionsPart);
        CopyPartXml(sourceDoc, blockDoc, p => p.StylesWithEffectsPart);
        CopyPartXml(sourceDoc, blockDoc, p => p.NumberingDefinitionsPart);
        CopyPartXml(sourceDoc, blockDoc, p => p.ThemePart);
        CopyPartXml(sourceDoc, blockDoc, p => p.FontTablePart);
        CopyPartXml(sourceDoc, blockDoc, p => p.DocumentSettingsPart);
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
            // Must FOLLOW the caller's profile: with it off, ProcessFootnoteReference
            // returns null and a re-rendered citing paragraph silently loses its
            // citation marker from the DOM (the XML keeps the reference).
            RenderFootnotesAndEndnotes = options.RenderFootnotesAndEndnotes,
            // These settings affect within-block output too. An incremental swap must not
            // accept revisions that the matching full render preserves.
            RenderTrackedChanges = options.RenderTrackedChanges,
            ShowDeletedContent = options.ShowDeletedContent,
            RenderMoveOperations = options.RenderMoveOperations,
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
