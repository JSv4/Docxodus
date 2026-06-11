#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>
/// Which parts of the OOXML package to include in the markdown projection.
/// </summary>
[Flags]
public enum ProjectionScopes
{
    Body = 1,
    Headers = 2,
    Footers = 4,
    Footnotes = 8,
    Endnotes = 16,
    Comments = 32,
    All = Body | Headers | Footers | Footnotes | Endnotes | Comments
}

/// <summary>
/// How anchor markers are rendered in the markdown output.
/// </summary>
public enum AnchorRenderMode
{
    /// <summary>Anchor appears on its own line before each block element (default).</summary>
    Block,
    /// <summary>Block anchors plus inline {#...} markers for spans (comments, hyperlinks, fields).</summary>
    BlockAndInline,
    /// <summary>No anchor markers in the output (projection only, no addressing).</summary>
    None
}

/// <summary>
/// Strategy for converting <c>w:tbl</c> elements that don't fit GFM pipe-table constraints.
/// </summary>
public enum TableRenderMode
{
    /// <summary>Emit GFM pipe tables when possible, opaque anchor blocks otherwise (default).</summary>
    GfmWithOpaqueFallback,
    /// <summary>Always emit GFM pipe tables, flattening complex structure with possible loss.</summary>
    AlwaysGfm,
    /// <summary>Always emit opaque anchor blocks. Useful when callers will fetch tables via the SDK.</summary>
    AlwaysOpaque
}

/// <summary>
/// How tracked changes are handled when projecting to markdown.
/// </summary>
public enum TrackedChangeMode
{
    /// <summary>Accept all revisions before conversion (default).</summary>
    Accept,
    /// <summary>Render insertions and deletions inline as <c>{+ins+}</c> / <c>{-del-}</c>.</summary>
    RenderInline,
    /// <summary>Accept insertions, drop deletions.</summary>
    StripDeletions
}

/// <summary>
/// Settings controlling markdown projection of a Word document. See
/// <c>docs/architecture/markdown_projection.md</c> for the full specification.
/// </summary>
public class WmlToMarkdownConverterSettings
{
    /// <summary>Which package parts to include in the projection.</summary>
    public ProjectionScopes Scopes { get; set; } = ProjectionScopes.All;

    /// <summary>
    /// Offset added to Word heading levels when emitting markdown headings.
    /// Example: an offset of 1 turns Word Heading1 into a markdown <c>##</c>.
    /// </summary>
    public int HeadingLevelOffset { get; set; } = 0;

    /// <summary>How anchor markers are emitted in the output.</summary>
    public AnchorRenderMode AnchorMode { get; set; } = AnchorRenderMode.Block;

    /// <summary>Strategy for tables that don't fit GFM pipe-table syntax.</summary>
    public TableRenderMode TableMode { get; set; } = TableRenderMode.GfmWithOpaqueFallback;

    /// <summary>
    /// Maximum characters per cell before a simple table is downgraded to an opaque
    /// anchor block. Ignored when <see cref="TableMode"/> is <see cref="TableRenderMode.AlwaysGfm"/>
    /// or <see cref="TableRenderMode.AlwaysOpaque"/>.
    /// </summary>
    public int TableInlineCellMax { get; set; } = 80;

    /// <summary>How tracked changes (<c>w:ins</c>, <c>w:del</c>) are projected.</summary>
    public TrackedChangeMode TrackedChanges { get; set; } = TrackedChangeMode.Accept;

    /// <summary>
    /// Resolve list numbering (<c>w:numPr</c>) to literal markers (<c>1.</c>, <c>a.</c>, etc.).
    /// When false, list items are emitted as plain markdown unordered items.
    /// </summary>
    public bool ResolveNumbering { get; set; } = true;

    /// <summary>
    /// Builds the URI used for image references in the output. Defaults to
    /// <c>docxodus://img/{unid}</c>; callers that want data URIs or HTTP URLs can override.
    /// </summary>
    public Func<ImageInfo, string>? ImageUriBuilder { get; set; }

    /// <summary>
    /// How empty paragraphs (paragraphs with no visible runs — Word's "spacer" paragraphs)
    /// render in the projection. Default <see cref="EmptyParagraphMode.AnchorOnly"/>
    /// preserves the addressable anchor on its own line so callers can target the spacer
    /// for edits; <see cref="EmptyParagraphMode.MarkedEmpty"/> tags it visibly for agents
    /// pattern-matching; <see cref="EmptyParagraphMode.Suppress"/> drops empty paragraphs
    /// entirely (and from the anchor index) for callers that want a denser projection.
    /// </summary>
    public EmptyParagraphMode EmptyParagraphs { get; set; } = EmptyParagraphMode.AnchorOnly;

    /// <summary>How anchor ids are rendered in <c>{#…}</c> tokens and keyed in
    /// the resulting <see cref="MarkdownProjection.AnchorIndex"/>. Default
    /// <see cref="AnchorIdRendering.FullUnid"/> matches legacy behavior.</summary>
    public AnchorIdRendering AnchorIdRendering { get; set; } = AnchorIdRendering.FullUnid;
}

/// <summary>How empty paragraphs are rendered by <see cref="WmlToMarkdownConverter"/>.</summary>
public enum EmptyParagraphMode
{
    /// <summary>Default: emit <c>{#p:body:UNID}\n</c> (anchor + newline, no body text).</summary>
    AnchorOnly,

    /// <summary>Emit <c>{#p:body:UNID} ∅\n</c> so agents can pattern-match empty paragraphs.</summary>
    MarkedEmpty,

    /// <summary>Skip empty paragraphs entirely — they don't appear in the markdown or the anchor index.</summary>
    Suppress,
}

/// <summary>How anchor ids are rendered inside <c>{#…}</c> tokens and keyed in
/// <see cref="MarkdownProjection.AnchorIndex"/>.</summary>
public enum AnchorIdRendering
{
    /// <summary>Full 32-char hex Unid (e.g. <c>{#h:body:a1b2c3d4e5f6789012345678901234ab}</c>).
    /// Default; matches the projection's existing behavior.</summary>
    FullUnid = 0,

    /// <summary>Shortest unique prefix per (kind, scope) bucket, with a 4-char
    /// floor (e.g. <c>{#h:body:a1b2}</c>). Saves 5-10% of projection-token
    /// budget for LLM consumption. The <see cref="MarkdownProjection.AnchorIndex"/>
    /// is dual-keyed — lookups by either full Unid or abbreviated form work.</summary>
    Abbreviated = 1,

    /// <summary>Sequential numeric ids per (kind, scope) bucket, in document
    /// order (e.g. <c>{#h:body:1} {#h:body:2}</c>). Maximally token-efficient
    /// for one-shot LLM contexts where anchor stability across sessions is
    /// unnecessary. These ids are NOT stable across <c>Project()</c> calls and
    /// must NOT be persisted; the dual-keyed <c>AnchorIndex</c> contains the
    /// numeric form too, but the canonical <see cref="Anchor.Id"/> on each
    /// <c>AnchorTarget</c> still carries the full Unid.</summary>
    Sequential = 2,
}

/// <summary>
/// Identifies a single addressable element in the markdown projection.
/// Anchor ids have the form <c>kind:scope:unid</c> (e.g.
/// <c>p:body:a1b2c3d4e5f6789012345678901234ab</c> — the example is the canonical
/// 32-char hex Unid form; abbreviated/sequential renderings may shorten it).
/// </summary>
public readonly record struct Anchor(string Id, string Kind, string Scope, string Unid)
{
    /// <summary>
    /// The canonical block-level token built from <see cref="Id"/> (e.g.
    /// <c>{#p:body:a1b2c3d4e5f6789012345678901234ab}</c>). Always uses the full
    /// Unid form regardless of the projection's
    /// <see cref="WmlToMarkdownConverterSettings.AnchorIdRendering"/> setting —
    /// for the actually-rendered token in <see cref="MarkdownProjection.Markdown"/>,
    /// the rendered id is in the markdown text and the dual-keyed
    /// <see cref="MarkdownProjection.AnchorIndex"/> resolves either form.
    /// </summary>
    public string Token => $"{{#{Id}}}";
}

/// <summary>
/// Resolved location of an anchor in the underlying OOXML package — sufficient to walk
/// back to the source <see cref="XElement"/> and apply edits via the Open XML SDK.
/// </summary>
public sealed class AnchorTarget
{
    /// <summary>The anchor this target resolves.</summary>
    required public Anchor Anchor { get; init; }

    /// <summary>URI of the package part containing the element (e.g. main document, header part).</summary>
    required public string PartUri { get; init; }

    /// <summary>Stable Unid of the element. Use with Docxodus' Unid lookup to fetch the live XElement.</summary>
    required public string Unid { get; init; }

    /// <summary>
    /// First ~80 characters of the element's flat text, suitable for showing in
    /// agent context windows or UI lists. Computed during projection so agents
    /// don't need to <see cref="Resolve"/> + re-walk the element for previews.
    /// Empty for elements with no text (e.g. empty paragraphs, section breaks).
    /// </summary>
    public string TextPreview { get; init; } = string.Empty;

    /// <summary>
    /// Resolved auto-numbering prefix Word would render for this element — e.g.
    /// <c>"First"</c>, <c>"1."</c>, <c>"1.1"</c>. <c>null</c> when the element has
    /// no <c>w:numPr</c>, when its style doesn't contribute numbering, or when
    /// the kind doesn't carry numbering (everything except <c>p</c>/<c>h</c>/<c>li</c>).
    /// <para>
    /// This bridges a foot-gun: the markdown projection emits the resolved prefix
    /// inline (so <c>"# First The total number…"</c> in the rendered markdown),
    /// but the underlying run text contains only <c>"The total number…"</c>. A
    /// caller searching the doc via <see cref="DocxSession.Grep"/> for what they
    /// see in the projection won't find <c>"First"</c> on its own — it lives here.
    /// </para>
    /// </summary>
    public string? AutoNumberPrefix { get; init; }

    /// <summary>
    /// The element's text as a reader would see it: <see cref="AutoNumberPrefix"/>
    /// joined with <see cref="TextPreview"/> by a single space when a prefix is
    /// present, otherwise just <see cref="TextPreview"/>. Convenience for UI / log
    /// surfaces that want "what does this block say?" without re-resolving numbering.
    /// </summary>
    public string FullText =>
        string.IsNullOrEmpty(AutoNumberPrefix)
            ? TextPreview
            : string.IsNullOrEmpty(TextPreview)
                ? AutoNumberPrefix!
                : AutoNumberPrefix + " " + TextPreview;

    /// <summary>
    /// Resolves this anchor to its current <see cref="XElement"/> inside the given document.
    /// Returns null if the element has been removed since projection.
    /// </summary>
    public XElement? Resolve(WordprocessingDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var main = document.MainDocumentPart;
        if (main == null) return null;

        OpenXmlPart? part = null;
        var targetUri = PartUri;
        if (main.Uri.ToString() == targetUri) part = main;
        else
        {
            foreach (var h in main.HeaderParts) if (h.Uri.ToString() == targetUri) { part = h; break; }
            if (part == null)
                foreach (var f in main.FooterParts) if (f.Uri.ToString() == targetUri) { part = f; break; }
            if (part == null && main.FootnotesPart?.Uri.ToString() == targetUri) part = main.FootnotesPart;
            if (part == null && main.EndnotesPart?.Uri.ToString() == targetUri) part = main.EndnotesPart;
            if (part == null && main.WordprocessingCommentsPart?.Uri.ToString() == targetUri) part = main.WordprocessingCommentsPart;
        }

        var root = part?.GetXDocument().Root;
        return root?.DescendantsAndSelf()
            .FirstOrDefault(e => (string?)e.Attribute(PtOpenXml.Unid) == Unid);
    }
}

/// <summary>
/// The output of a markdown projection: the rendered text plus the anchor index that maps
/// each <c>{#...}</c> token back to a location in the OOXML package.
/// </summary>
public sealed class MarkdownProjection
{
    required public string Markdown { get; init; }
    required public IReadOnlyDictionary<string, AnchorTarget> AnchorIndex { get; init; }
}

public partial class WmlDocument
{
    /// <summary>
    /// Project this document to anchor-addressed markdown. See
    /// <c>docs/architecture/markdown_projection.md</c> for the projection spec.
    /// </summary>
    public MarkdownProjection ConvertToMarkdown(WmlToMarkdownConverterSettings settings)
    {
        return WmlToMarkdownConverter.Convert(this, settings);
    }
}

/// <summary>
/// Projects a Word document to anchor-addressed markdown — a stable, deterministic text view
/// suitable for tools that want to read, search, and edit Word documents the way they would
/// source files (LLM editing pipelines, structured search indexers, diff/review UIs).
///
/// This is a scaffold. The public surface is fixed; the implementation is staged in phases
/// described in <c>docs/architecture/markdown_projection.md</c>.
/// </summary>
public static class WmlToMarkdownConverter
{
    /// <summary>
    /// Convert a <see cref="WmlDocument"/> to its markdown projection. Persists any newly
    /// assigned <c>PtOpenXml.Unid</c> attributes back into <paramref name="document"/>'s
    /// underlying byte array so subsequent <see cref="AnchorTarget.Resolve"/> calls against
    /// the same bytes succeed.
    /// </summary>
    public static MarkdownProjection Convert(WmlDocument document, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(settings);

        using var stream = new OpenXmlMemoryStreamDocument(document);
        MarkdownProjection projection;
        using (var wdoc = stream.GetWordprocessingDocument())
        {
            projection = Convert(wdoc, settings);
        }
        var modified = stream.GetModifiedWmlDocument();
        document.DocumentByteArray = modified.DocumentByteArray;
        return projection;
    }

    /// <summary>
    /// Convert an open <see cref="WordprocessingDocument"/> to its markdown projection
    /// without round-tripping through <see cref="WmlDocument"/>. Mutates the in-memory
    /// XDocument of each in-scope part to add missing <c>PtOpenXml.Unid</c> attributes and
    /// persists those mutations via <c>PutXDocument</c>; the caller is responsible for
    /// saving the package itself.
    /// </summary>
    public static MarkdownProjection Convert(WordprocessingDocument document, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(settings);

        var (index, scopes, renderMap) = BuildAnchorIndex(document, settings);
        var markdown = EmitMarkdown(document, settings, scopes, renderMap);
        return new MarkdownProjection { Markdown = markdown, AnchorIndex = index };
    }

    // ------------------------------------------------------------------
    // Phase 1: anchor index
    // ------------------------------------------------------------------

    private sealed class ScopeInfo
    {
        required public string Name { get; init; }
        required public OpenXmlPart Part { get; init; }
        required public XElement Root { get; init; }
    }

    /// <summary>
    /// Per-projection map from full Unid → rendered id. Used by every <c>{#…}</c>
    /// emission site so the projection's anchor tokens consistently use the
    /// caller-requested rendering. The <see cref="Render"/> method returns the full
    /// Unid unchanged when the rendering is <see cref="AnchorIdRendering.FullUnid"/>
    /// or when an unknown Unid is passed (defensive fallback).
    /// </summary>
    internal sealed class AnchorIdMap
    {
        private readonly Dictionary<string, string> _map = new(StringComparer.Ordinal);

        public string Render(string fullUnid) => _map.TryGetValue(fullUnid, out var r) ? r : fullUnid;

        internal void Set(string fullUnid, string renderedUnid) => _map[fullUnid] = renderedUnid;

        internal IReadOnlyDictionary<string, string> Map => _map;
    }

    private const int TextPreviewMaxLength = 80;

    private static string ComputeTextPreview(XElement element)
    {
        var text = string.Concat(element.Descendants(W.t).Select(t => (string)t));
        return text.Length > TextPreviewMaxLength
            ? text.Substring(0, TextPreviewMaxLength) + "…"
            : text;
    }

    private static (IReadOnlyDictionary<string, AnchorTarget> Index, List<ScopeInfo> Scopes, AnchorIdMap RenderMap)
        BuildAnchorIndex(WordprocessingDocument doc, WmlToMarkdownConverterSettings settings)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");

        var scopes = new List<ScopeInfo>();
        if (settings.Scopes.HasFlag(ProjectionScopes.Body))
            scopes.Add(new ScopeInfo { Name = "body", Part = main, Root = main.GetXDocument().Root! });
        if (settings.Scopes.HasFlag(ProjectionScopes.Headers))
        {
            var i = 1;
            foreach (var hp in main.HeaderParts)
                scopes.Add(new ScopeInfo { Name = $"hdr{i++}", Part = hp, Root = hp.GetXDocument().Root! });
        }
        if (settings.Scopes.HasFlag(ProjectionScopes.Footers))
        {
            var i = 1;
            foreach (var fp in main.FooterParts)
                scopes.Add(new ScopeInfo { Name = $"ftr{i++}", Part = fp, Root = fp.GetXDocument().Root! });
        }
        if (settings.Scopes.HasFlag(ProjectionScopes.Footnotes) && main.FootnotesPart != null)
            scopes.Add(new ScopeInfo { Name = "fn", Part = main.FootnotesPart, Root = main.FootnotesPart.GetXDocument().Root! });
        if (settings.Scopes.HasFlag(ProjectionScopes.Endnotes) && main.EndnotesPart != null)
            scopes.Add(new ScopeInfo { Name = "en", Part = main.EndnotesPart, Root = main.EndnotesPart.GetXDocument().Root! });
        if (settings.Scopes.HasFlag(ProjectionScopes.Comments) && main.WordprocessingCommentsPart != null)
            scopes.Add(new ScopeInfo { Name = "cmt", Part = main.WordprocessingCommentsPart, Root = main.WordprocessingCommentsPart.GetXDocument().Root! });

        var index = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
        foreach (var scope in scopes)
        {
            // Deterministic Unid path so the same docx bytes produce identical
            // anchor ids across sessions. WmlComparer continues to use the random
            // path via AssignToAllElements — see UnidHelper class doc for why
            // the two consumers stay split.
            UnidHelper.AssignToAllElementsDeterministic(scope.Root);
            // Stash the owning part on the root so downstream emitters (hyperlinks, etc.)
            // can resolve relationship-bound URIs without threading the part through every call.
            if (scope.Root.Annotation<OpenXmlPart>() == null)
                scope.Root.AddAnnotation(scope.Part);

            // Word-reserved footnote/endnote separators (type="separator" / type="continuationSeparator")
            // are structural plumbing that cannot be deleted and should not appear in the agent-facing
            // AnchorIndex. Pre-collect them and their descendants so the walker skips both the notes
            // themselves and any paragraphs/runs they contain.
            var skip = new HashSet<XElement>();
            if (scope.Name is "fn" or "en")
            {
                var noteName = scope.Name == "fn" ? W.footnote : W.endnote;
                foreach (var n in scope.Root.Elements(noteName))
                {
                    if (IsBoilerplateNote(n))
                    {
                        skip.Add(n);
                        foreach (var d in n.Descendants()) skip.Add(d);
                    }
                }
            }

            foreach (var el in scope.Root.DescendantsAndSelf())
            {
                if (skip.Contains(el)) continue;
                var kind = KindFor(el);
                if (kind == null) continue;
                var unid = (string?)el.Attribute(PtOpenXml.Unid);
                if (unid == null) continue;
                // Suppress-mode: drop empty paragraphs from the AnchorIndex too,
                // so callers iterating the index don't see anchors that have no
                // corresponding line in the projection. Mirrors what EmitParagraph does.
                if (settings.EmptyParagraphs == EmptyParagraphMode.Suppress
                    && el.Name == W.p
                    && kind is "p" or "h" or "li"
                    && !el.Descendants(W.t).Any(t => !string.IsNullOrEmpty((string)t)))
                    continue;
                var id = $"{kind}:{scope.Name}:{unid}";
                if (index.ContainsKey(id)) continue;
                var anchor = new Anchor(id, kind, scope.Name, unid);
                index[id] = new AnchorTarget
                {
                    Anchor = anchor,
                    PartUri = scope.Part.Uri.ToString(),
                    Unid = unid,
                    TextPreview = ComputeTextPreview(el),
                    AutoNumberPrefix = kind is "p" or "h" or "li" && scope.Name == "body"
                        ? Internal.ListNumberResolver.Resolve(el, doc)
                        : null,
                };
            }
            scope.Part.PutXDocument();
        }

        // Build the AnchorIdMap based on the requested rendering mode.
        var renderMap = new AnchorIdMap();
        if (settings.AnchorIdRendering == AnchorIdRendering.Abbreviated)
        {
            // Group anchors by (kind, scope). Within each group, find the shortest
            // prefix length n >= 4 such that every member's Unid[..n] is unique.
            foreach (var bucket in index.Values.GroupBy(t => (t.Anchor.Kind, t.Anchor.Scope)))
            {
                var members = bucket.ToList();
                if (members.Count == 0) continue;
                int n = 4;
                while (true)
                {
                    var prefixes = new HashSet<string>(StringComparer.Ordinal);
                    bool unique = true;
                    foreach (var t in members)
                    {
                        var prefix = t.Unid.Length >= n ? t.Unid.Substring(0, n) : t.Unid;
                        if (!prefixes.Add(prefix)) { unique = false; break; }
                    }
                    if (unique) break;
                    n++;
                    if (n >= 32) break;   // fall back to full Unid
                }
                foreach (var t in members)
                {
                    var prefix = t.Unid.Length >= n ? t.Unid.Substring(0, n) : t.Unid;
                    renderMap.Set(t.Unid, prefix);
                }
            }
        }
        else if (settings.AnchorIdRendering == AnchorIdRendering.Sequential)
        {
            // Assign 1-based numeric ids per (kind, scope) bucket in document order
            // (which is the natural order of index.Values since Dictionary preserves
            // insertion order and BuildAnchorIndex inserts in descendant-walk order).
            var counters = new Dictionary<(string Kind, string Scope), int>();
            foreach (var t in index.Values)
            {
                var bucket = (t.Anchor.Kind, t.Anchor.Scope);
                if (!counters.TryGetValue(bucket, out var n)) n = 0;
                n++;
                counters[bucket] = n;
                renderMap.Set(t.Unid, n.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        // Dual-key the index: add alias entries with the rendered id substituted.
        if (settings.AnchorIdRendering != AnchorIdRendering.FullUnid)
        {
            var aliases = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
            foreach (var (key, target) in index)
            {
                var rendered = renderMap.Render(target.Unid);
                if (rendered == target.Unid) continue;
                var aliasKey = $"{target.Anchor.Kind}:{target.Anchor.Scope}:{rendered}";
                aliases[aliasKey] = target;
            }
            foreach (var (key, target) in aliases)
                index[key] = target;
        }

        return (index, scopes, renderMap);
    }

    /// <summary>
    /// Classify an element to its anchor <c>kind</c>. Returns <c>null</c> for elements that
    /// are not addressable by the projection (runs, inline children, formatting properties).
    /// </summary>
    internal static string? KindFor(XElement el)
    {
        var n = el.Name;
        if (n == W.p) return IsHeading(el) ? "h" : IsListItem(el) ? "li" : "p";
        if (n == W.tbl) return "tbl";
        if (n == W.tr) return "tr";
        if (n == W.tc) return "tc";
        if (n == W.sectPr) return "sec";
        if (n == W.footnote) return "fn";
        if (n == W.endnote) return "en";
        if (n == W.comment) return "cmt";
        return null;
    }

    internal static bool IsHeading(XElement p)
    {
        var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (string.IsNullOrEmpty(styleId)) return false;
        return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            || styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
            || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsListItem(XElement p)
    {
        // Inline w:numPr wins.
        if (p.Element(W.pPr)?.Element(W.numPr) != null) return true;
        // Otherwise the paragraph is a list item if its pStyle chain contributes numPr.
        // Resolved via the styles part reachable from the in-memory OpenXmlPart annotation
        // that BuildAnchorIndex stashes on each scope root.
        var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (string.IsNullOrEmpty(styleId)) return false;
        var part = p.Document?.Root?.Annotation<OpenXmlPart>();
        if (part?.OpenXmlPackage is not WordprocessingDocument wDoc) return false;
        var stylesRoot = wDoc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
        if (stylesRoot is null) return false;

        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        for (int i = 0; i < 16 && current is not null; i++)
        {
            if (!visited.Add(current)) return false;  // cycle
            var style = stylesRoot.Elements(W.style)
                .FirstOrDefault(s => (string?)s.Attribute(W.styleId) == current);
            if (style is null) return false;
            if (style.Element(W.pPr)?.Element(W.numPr) != null) return true;
            current = (string?)style.Element(W.basedOn)?.Attribute(W.val);
        }
        return false;
    }

    // ------------------------------------------------------------------
    // Markdown emission. Phase 2+ implementation. Per-element handlers are
    // dispatched by element name from EmitBlocks; the EmitContext threads the
    // current scope name and a StringBuilder so handlers stay pure of state.
    // ------------------------------------------------------------------

    private sealed class EmitContext
    {
        public StringBuilder Sb { get; } = new();
        required public WmlToMarkdownConverterSettings Settings { get; init; }
        required public WordprocessingDocument Document { get; init; }
        required public AnchorIdMap AnchorIdMap { get; init; }
        public string Scope { get; set; } = "body";
        public ListItemRetrieverSettings ListItemRetrieverSettings { get; } = new();
        public bool InsideListBlock { get; set; }
    }

    private static string EmitMarkdown(
        WordprocessingDocument document,
        WmlToMarkdownConverterSettings settings,
        List<ScopeInfo> scopes,
        AnchorIdMap renderMap)
    {
        var ctx = new EmitContext { Settings = settings, Document = document, AnchorIdMap = renderMap };

        // Each scope's emitter appends to ctx.Sb. After each scope, we record whether anything
        // was actually written; a divider is emitted between two non-empty scopes so
        // downstream parsers can split the markdown into per-scope chunks without inspecting
        // heading text.
        var anyScopeEmitted = false;

        var bodyScope = scopes.FirstOrDefault(s => s.Name == "body");
        if (bodyScope != null)
        {
            ctx.Sb.AppendLine("# Document");
            ctx.Sb.AppendLine();
            ctx.Scope = "body";
            var body = bodyScope.Root.Element(W.body);
            if (body != null) EmitBlocks(body.Elements(), ctx);
            anyScopeEmitted = true;
        }

        // Many DOCXs declare 6+ header/footer parts for first-page/even-page/default variants
        // and leave the unused ones blank. Suppress scopes with no text content so the
        // projection isn't padded with empty "## hdrN" titles.
        var headerScopes = scopes
            .Where(s => s.Name.StartsWith("hdr", StringComparison.Ordinal) && ScopeHasContent(s))
            .ToList();
        if (headerScopes.Count > 0)
        {
            if (anyScopeEmitted) AppendScopeDivider(ctx.Sb);
            ctx.Sb.AppendLine("# Headers");
            ctx.Sb.AppendLine();
            foreach (var s in headerScopes)
            {
                ctx.Sb.Append("## ").AppendLine(s.Name);
                ctx.Sb.AppendLine();
                ctx.Scope = s.Name;
                EmitBlocks(s.Root.Elements(), ctx);
            }
            anyScopeEmitted = true;
        }

        var footerScopes = scopes
            .Where(s => s.Name.StartsWith("ftr", StringComparison.Ordinal) && ScopeHasContent(s))
            .ToList();
        if (footerScopes.Count > 0)
        {
            if (anyScopeEmitted) AppendScopeDivider(ctx.Sb);
            ctx.Sb.AppendLine("# Footers");
            ctx.Sb.AppendLine();
            foreach (var s in footerScopes)
            {
                ctx.Sb.Append("## ").AppendLine(s.Name);
                ctx.Sb.AppendLine();
                ctx.Scope = s.Name;
                EmitBlocks(s.Root.Elements(), ctx);
            }
            anyScopeEmitted = true;
        }

        var fnScope = scopes.FirstOrDefault(s => s.Name == "fn");
        if (fnScope != null)
        {
            var before = ctx.Sb.Length;
            EmitNoteDefinitions(fnScope, ctx, "Footnotes", "fn", W.footnote, anyScopeEmitted);
            if (ctx.Sb.Length > before) anyScopeEmitted = true;
        }

        var enScope = scopes.FirstOrDefault(s => s.Name == "en");
        if (enScope != null)
        {
            var before = ctx.Sb.Length;
            EmitNoteDefinitions(enScope, ctx, "Endnotes", "en", W.endnote, anyScopeEmitted);
            if (ctx.Sb.Length > before) anyScopeEmitted = true;
        }

        var cmtScope = scopes.FirstOrDefault(s => s.Name == "cmt");
        if (cmtScope != null)
        {
            var before = ctx.Sb.Length;
            EmitComments(cmtScope, ctx, anyScopeEmitted);
            if (ctx.Sb.Length > before) anyScopeEmitted = true;
        }

        return ctx.Sb.ToString();
    }

    private static void AppendScopeDivider(StringBuilder sb)
    {
        sb.AppendLine("---");
        sb.AppendLine();
    }

    /// <summary>
    /// True when a header/footer scope has any non-whitespace text. Used to suppress
    /// scope sections that exist in the package but carry no user-visible content.
    /// </summary>
    private static bool ScopeHasContent(ScopeInfo scope)
    {
        foreach (var t in scope.Root.Descendants(W.t))
        {
            if (!string.IsNullOrWhiteSpace(t.Value)) return true;
        }
        return false;
    }

    private static void EmitNoteDefinitions(ScopeInfo scope, EmitContext ctx, string header, string kindPrefix, XName elementName, bool precedingContent)
    {
        // The footnotes/endnotes parts always carry the separator/continuationSeparator
        // boilerplate notes with type="separator"/type="continuationSeparator". Filter those
        // out so only user-authored notes reach the projection.
        var notes = scope.Root
            .Elements(elementName)
            .Where(n => !IsBoilerplateNote(n))
            .ToList();
        if (notes.Count == 0) return;

        if (precedingContent) AppendScopeDivider(ctx.Sb);
        ctx.Sb.Append("# ").AppendLine(header);
        ctx.Sb.AppendLine();
        ctx.Scope = scope.Name;
        foreach (var note in notes)
        {
            var unid = (string?)note.Attribute(PtOpenXml.Unid) ?? "0";
            var label = $"{kindPrefix}-{NoteLabelSuffix(unid, ctx)}";
            ctx.Sb.Append("[^").Append(label).Append("]: ");
            // Notes contain paragraphs; flatten their text inline for the definition.
            var first = true;
            foreach (var p in note.Elements(W.p))
            {
                if (!first) ctx.Sb.Append(' ');
                first = false;
                EmitInlineRuns(p, ctx);
            }
            ctx.Sb.AppendLine();
            ctx.Sb.AppendLine();
        }
    }

    internal static bool IsBoilerplateNote(XElement note)
    {
        var type = (string?)note.Attribute(W.type);
        return type is "separator" or "continuationSeparator";
    }

    private static void EmitComments(ScopeInfo scope, EmitContext ctx, bool precedingContent)
    {
        var comments = scope.Root.Elements(W.comment).ToList();
        if (comments.Count == 0) return;

        if (precedingContent) AppendScopeDivider(ctx.Sb);
        ctx.Sb.AppendLine("# Comments");
        ctx.Sb.AppendLine();
        ctx.Scope = "cmt";
        foreach (var c in comments)
        {
            var unid = (string?)c.Attribute(PtOpenXml.Unid) ?? "0";
            var author = (string?)c.Attribute(W.author) ?? "unknown";
            var date = (string?)c.Attribute(W.date);
            var renderedUnid = ctx.AnchorIdMap.Render(unid);
            ctx.Sb.Append($"- {{#cmt:cmt:{renderedUnid}}} **{author}**");
            if (!string.IsNullOrEmpty(date)) ctx.Sb.Append(" (").Append(date).Append(')');
            ctx.Sb.Append(": ");
            foreach (var p in c.Elements(W.p))
            {
                EmitInlineRuns(p, ctx);
                ctx.Sb.Append(' ');
            }
            ctx.Sb.AppendLine();
        }
        ctx.Sb.AppendLine();
    }

    private static string ShortUnid(string unid) =>
        unid.Length >= 8 ? unid.Substring(0, 8) : unid;

    /// <summary>
    /// Compute the suffix used in <c>[^fn-…]</c> / <c>[^en-…]</c> footnote/endnote labels.
    /// In FullUnid mode we keep the legacy 8-char truncation (the full Unid is too long to
    /// embed in a label); in Abbreviated/Sequential modes we route through the AnchorIdMap
    /// so the rendered form is consistent with the corresponding <c>{#fn:…}</c> anchor token.
    /// </summary>
    private static string NoteLabelSuffix(string unid, EmitContext ctx) =>
        ctx.Settings.AnchorIdRendering == AnchorIdRendering.FullUnid
            ? ShortUnid(unid)
            : ctx.AnchorIdMap.Render(unid);

    private static void EmitBlocks(IEnumerable<XElement> blocks, EmitContext ctx)
    {
        var blocksList = blocks.ToList();
        for (var i = 0; i < blocksList.Count; i++)
        {
            var b = blocksList[i];
            if (b.Name == W.p)
            {
                // Track whether the next block continues a list so we don't append blank lines
                // between adjacent list items (the spec says lists are not separated internally).
                var nextIsListItem = i + 1 < blocksList.Count
                    && blocksList[i + 1].Name == W.p
                    && IsListItem(blocksList[i + 1]);
                ctx.InsideListBlock = IsListItem(b);
                EmitParagraph(b, ctx);
                if (IsListItem(b) && !nextIsListItem)
                {
                    // End of a list block — emit the trailing blank line that the list item
                    // emitter intentionally omits.
                    ctx.Sb.AppendLine();
                }
            }
            else if (b.Name == W.tbl) EmitTable(b, ctx);
            else if (b.Name == W.sectPr) { /* phase 6: section breaks */ }
        }
    }

    private static void EmitParagraph(XElement p, EmitContext ctx)
    {
        var anchor = AnchorPrefix(p, ctx);

        if (IsHeading(p))
        {
            // Clamp upper bound is 9 (Word's outline depth), not 6: silently collapsing
            // Heading7-9 to ###### loses outline depth that callers (LLMs, ToC builders,
            // diff renderers) rely on. ATX headings beyond 6 are not standard CommonMark,
            // but downstream parsers and LLM consumers handle 7-9 hashes naturally; strict
            // renderers degrade to literal text, which is still better than silent collapse.
            var level = Math.Clamp(HeadingLevel(p) + ctx.Settings.HeadingLevelOffset, 1, 9);
            ctx.Sb.Append(anchor);
            ctx.Sb.Append(new string('#', level));
            ctx.Sb.Append(' ');
            // Legal-style headings often style each clause Heading{N} AND attach w:numPr so
            // Word renders "FIRST: …" / "1.1 …". Resolve that prefix here so it survives
            // projection — without it, callers see "## : The name…" with the auto-number gone.
            if (ctx.Settings.ResolveNumbering)
            {
                var numberPrefix = ResolveHeadingNumberPrefix(p, ctx);
                if (numberPrefix != null) ctx.Sb.Append(numberPrefix).Append(' ');
            }
            EmitInlineRuns(p, ctx);
            ctx.Sb.AppendLine();
            ctx.Sb.AppendLine();
            EmitInlineSectionBreak(p, ctx);
            return;
        }

        if (IsListItem(p))
        {
            EmitListItem(p, ctx);
            EmitInlineSectionBreak(p, ctx);
            return;
        }

        // Empty-paragraph handling honors EmptyParagraphMode. Detected by "no visible
        // inline content emitted between the anchor token and the line terminator."
        var modeForEmpty = ctx.Settings.EmptyParagraphs;
        if (modeForEmpty == EmptyParagraphMode.Suppress && !HasVisibleInlineContent(p, ctx))
        {
            // Skip the paragraph entirely. Section breaks still need to surface
            // (they're metadata, not content) so emit them even when the spacer is dropped.
            EmitInlineSectionBreak(p, ctx);
            return;
        }

        var beforeContent = ctx.Sb.Length;
        ctx.Sb.Append(anchor);
        var afterAnchor = ctx.Sb.Length;
        EmitInlineRuns(p, ctx);
        // If the paragraph had no visible runs, either tag it with the marked-empty
        // sentinel (∅) or strip the anchor's dangling trailing separator space.
        if (ctx.Sb.Length == afterAnchor && afterAnchor > beforeContent)
        {
            if (modeForEmpty == EmptyParagraphMode.MarkedEmpty)
            {
                ctx.Sb.Append('∅');
            }
            else if (ctx.Sb[ctx.Sb.Length - 1] == ' ')
            {
                ctx.Sb.Length--;
            }
        }
        ctx.Sb.AppendLine();
        ctx.Sb.AppendLine();
        EmitInlineSectionBreak(p, ctx);
    }

    /// <summary>True when the paragraph would emit any visible inline content under
    /// the current tracked-change mode. Used by <see cref="EmptyParagraphMode.Suppress"/>
    /// to decide whether to skip a paragraph entirely.</summary>
    private static bool HasVisibleInlineContent(XElement p, EmitContext ctx)
    {
        foreach (var (_, runs) in GroupInlineRuns(p))
        {
            foreach (var r in runs)
            {
                if (r.Descendants(W.t).Any(t => !string.IsNullOrEmpty((string)t))) return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Emit a thematic break (with the section's anchor) when this paragraph carries a
    /// <c>w:sectPr</c> in its <c>w:pPr</c> — that marks an in-document section transition.
    /// The trailing top-level <c>w:sectPr</c> at the end of the body is metadata for the last
    /// section, not a transition, so it is handled by <see cref="EmitBlocks"/> instead.
    /// </summary>
    private static void EmitInlineSectionBreak(XElement p, EmitContext ctx)
    {
        var sectPr = p.Element(W.pPr)?.Element(W.sectPr);
        if (sectPr == null) return;
        var unid = (string?)sectPr.Attribute(PtOpenXml.Unid);
        if (unid == null) return;
        if (ctx.Settings.AnchorMode != AnchorRenderMode.None)
        {
            var rendered = ctx.AnchorIdMap.Render(unid);
            ctx.Sb.Append("{#sec:").Append(ctx.Scope).Append(':').Append(rendered).AppendLine("}");
        }
        ctx.Sb.AppendLine("---");
        ctx.Sb.AppendLine();
    }

    /// <summary>Build the anchor prefix (with trailing space) for a block element, or empty string when AnchorMode==None.</summary>
    private static string AnchorPrefix(XElement el, EmitContext ctx)
    {
        if (ctx.Settings.AnchorMode == AnchorRenderMode.None) return string.Empty;
        var kind = KindFor(el) ?? "unk";
        var unid = (string?)el.Attribute(PtOpenXml.Unid) ?? "0";
        var rendered = ctx.AnchorIdMap.Render(unid);
        return $"{{#{kind}:{ctx.Scope}:{rendered}}} ";
    }

    internal static int HeadingLevel(XElement p)
    {
        var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val) ?? string.Empty;
        if (styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)) return 1;
        if (styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)) return 2;
        var digits = new string(styleId.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out var n) && n >= 1 && n <= 9 ? n : 1;
    }

    // ------------------------------------------------------------------
    // Phase 3: inline runs. Adjacent runs sharing the same formatting are
    // merged into one delimiter pair so "**a****b**" becomes "**ab**". Each
    // run's text is escaped before delimiters are wrapped around it.
    // ------------------------------------------------------------------

    private enum Revision { None, Inserted, Deleted }

    private readonly record struct RunFormatting(
        bool Bold,
        bool Italic,
        bool Code,
        bool Strike,
        string? HyperlinkUrl,
        Revision Revision);

    private static void EmitInlineRuns(XElement p, EmitContext ctx)
    {
        foreach (var (fmt, runs) in GroupInlineRuns(p))
        {
            // Tracked-change mode shapes what we keep before any formatting is applied.
            switch (ctx.Settings.TrackedChanges)
            {
                case TrackedChangeMode.Accept:
                    if (fmt.Revision == Revision.Deleted) continue;
                    break;
                case TrackedChangeMode.StripDeletions:
                    if (fmt.Revision == Revision.Deleted) continue;
                    break;
                case TrackedChangeMode.RenderInline:
                    // Both ins and del are wrapped in brace markers below.
                    break;
            }

            if (fmt.HyperlinkUrl != null)
            {
                ctx.Sb.Append('[');
                foreach (var r in runs) AppendRunText(r, ctx);
                ctx.Sb.Append("](").Append(fmt.HyperlinkUrl).Append(')');
                continue;
            }

            var (open, close) = MarkdownDelimiters(fmt);
            if (ctx.Settings.TrackedChanges == TrackedChangeMode.RenderInline)
            {
                if (fmt.Revision == Revision.Inserted) { open = "{+" + open; close += "+}"; }
                else if (fmt.Revision == Revision.Deleted) { open = "{-" + open; close += "-}"; }
            }

            ctx.Sb.Append(open);
            foreach (var r in runs) AppendRunText(r, ctx);
            ctx.Sb.Append(close);
        }
    }

    /// <summary>
    /// Walk the inline children of a paragraph (runs and hyperlinks containing runs) and
    /// return groups of adjacent runs that share the same formatting. Hyperlinks always form
    /// their own group regardless of neighbours so delimiter merging never crosses a link.
    /// </summary>
    private static List<(RunFormatting Fmt, List<XElement> Runs)> GroupInlineRuns(XElement p)
    {
        var groups = new List<(RunFormatting, List<XElement>)>();
        var buf = new List<XElement>();
        RunFormatting bufFmt = default;
        var primed = false;

        void Flush()
        {
            if (primed && buf.Count > 0)
                groups.Add((bufFmt, new List<XElement>(buf)));
            buf.Clear();
            primed = false;
        }

        void Add(XElement run, RunFormatting fmt)
        {
            if (!primed)
            {
                bufFmt = fmt;
                buf.Add(run);
                primed = true;
                return;
            }
            // Hyperlinks never merge across adjacent runs even if formatting matches —
            // each hyperlink is its own [text](url) span.
            if (fmt.HyperlinkUrl == null && bufFmt.HyperlinkUrl == null && fmt.Equals(bufFmt))
            {
                buf.Add(run);
                return;
            }
            Flush();
            bufFmt = fmt;
            buf.Add(run);
            primed = true;
        }

        foreach (var child in p.Elements())
        {
            if (child.Name == W.r)
            {
                Add(child, ReadRunFormatting(child, hyperlinkUrl: null, revision: Revision.None));
            }
            else if (child.Name == W.hyperlink)
            {
                var url = ResolveHyperlinkUrl(child);
                Flush();
                foreach (var r in child.Elements(W.r))
                    Add(r, ReadRunFormatting(r, hyperlinkUrl: url, revision: Revision.None));
                Flush();
            }
            else if (child.Name == W.ins)
            {
                Flush();
                foreach (var r in child.Descendants(W.r))
                    Add(r, ReadRunFormatting(r, hyperlinkUrl: null, revision: Revision.Inserted));
                Flush();
            }
            else if (child.Name == W.del)
            {
                Flush();
                foreach (var r in child.Descendants(W.r))
                    Add(r, ReadRunFormatting(r, hyperlinkUrl: null, revision: Revision.Deleted));
                Flush();
            }
        }
        Flush();
        return groups;
    }

    private static RunFormatting ReadRunFormatting(XElement run, string? hyperlinkUrl, Revision revision)
    {
        var rPr = run.Element(W.rPr);
        return new RunFormatting(
            Bold: HasToggle(rPr, W.b),
            Italic: HasToggle(rPr, W.i),
            Code: IsCodeRun(rPr),
            Strike: HasToggle(rPr, W.strike),
            HyperlinkUrl: hyperlinkUrl,
            Revision: revision);
    }

    private static bool HasToggle(XElement? rPr, XName name)
    {
        if (rPr == null) return false;
        var el = rPr.Element(name);
        if (el == null) return false;
        var val = (string?)el.Attribute(W.val);
        // OOXML toggle properties: absent => default (false), present-with-no-val => true,
        // present with val="0"/"false" => false, otherwise true.
        return val == null || (val != "0" && !val.Equals("false", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsCodeRun(XElement? rPr)
    {
        if (rPr == null) return false;
        var styleId = (string?)rPr.Element(W.rStyle)?.Attribute(W.val);
        if (styleId != null &&
            (styleId.Equals("Code", StringComparison.OrdinalIgnoreCase)
             || styleId.Equals("HTMLCode", StringComparison.OrdinalIgnoreCase)
             || styleId.Equals("VerbatimChar", StringComparison.OrdinalIgnoreCase)))
            return true;
        // Heuristic: monospace ascii font.
        var ascii = (string?)rPr.Element(W.rFonts)?.Attribute(W.ascii);
        if (ascii != null && (ascii.Contains("Mono", StringComparison.OrdinalIgnoreCase)
            || ascii.Contains("Courier", StringComparison.OrdinalIgnoreCase)
            || ascii.Contains("Consolas", StringComparison.OrdinalIgnoreCase)))
            return true;
        return false;
    }

    private static (string Open, string Close) MarkdownDelimiters(RunFormatting fmt)
    {
        if (fmt.Code) return ("`", "`"); // GFM: code is exclusive with bold/italic markup.
        var open = new StringBuilder();
        var close = new StringBuilder();
        if (fmt.Strike) { open.Append("~~"); close.Insert(0, "~~"); }
        if (fmt.Bold) { open.Append("**"); close.Insert(0, "**"); }
        if (fmt.Italic) { open.Append('*'); close.Insert(0, '*'); }
        return (open.ToString(), close.ToString());
    }

    private static string? ResolveHyperlinkUrl(XElement hyperlink)
    {
        // External: w:hyperlink with r:id pointing at a HyperlinkRelationship.
        var relId = (string?)hyperlink.Attribute(R.id);
        if (relId != null)
        {
            // Walk up to the containing part by finding the document root and chasing the
            // matching relationship from any ancestor; the relationship is on the part that
            // owns the XDocument, so we have to resolve at emit time.
            // We stash the part reference on the XDocument's annotations earlier? Simpler:
            // find the doc-level _parentPart by scanning ancestors of `hyperlink` until the
            // root element, then look up the relationship from the document via Annotations.
            var url = LookupRelationshipUrl(hyperlink, relId);
            if (url != null) return url;
        }
        // Internal: w:anchor for bookmark navigation.
        var anchor = (string?)hyperlink.Attribute(W.anchor);
        if (anchor != null) return $"#{anchor}";
        return null;
    }

    private static string? LookupRelationshipUrl(XElement el, string relId)
    {
        // BuildAnchorIndex stashes the owning OpenXmlPart on the root via XElement.AddAnnotation.
        var root = el.AncestorsAndSelf().Last();
        var part = root.Annotation<OpenXmlPart>();
        if (part == null) return null;
        foreach (var rel in part.HyperlinkRelationships)
        {
            if (rel.Id == relId) return rel.Uri.ToString();
        }
        return null;
    }

    private static void AppendRunText(XElement r, EmitContext ctx)
    {
        foreach (var node in r.Elements())
        {
            if (node.Name == W.t)
                ctx.Sb.Append(EscapeMarkdown((string)node));
            else if (node.Name == W.delText)
                ctx.Sb.Append(EscapeMarkdown((string)node));
            else if (node.Name == W.br)
                ctx.Sb.Append("  \n"); // hard line break in markdown
            else if (node.Name == W.tab)
                ctx.Sb.Append("    ");
            else if (node.Name == W.footnoteReference)
                AppendNoteRefMarker(node, ctx, "fn", W.footnote, ctx.Document.MainDocumentPart?.FootnotesPart);
            else if (node.Name == W.endnoteReference)
                AppendNoteRefMarker(node, ctx, "en", W.endnote, ctx.Document.MainDocumentPart?.EndnotesPart);
        }
    }

    private static void AppendNoteRefMarker(XElement reference, EmitContext ctx, string prefix, XName noteName, OpenXmlPart? notePart)
    {
        if (notePart == null) return;
        var id = (string?)reference.Attribute(W.id);
        if (id == null) return;
        var note = notePart.GetXDocument().Root?.Elements(noteName)
            .FirstOrDefault(n => (string?)n.Attribute(W.id) == id);
        var unid = (string?)note?.Attribute(PtOpenXml.Unid);
        if (unid == null) return;
        ctx.Sb.Append("[^").Append(prefix).Append('-').Append(NoteLabelSuffix(unid, ctx)).Append(']');
    }

    private static readonly System.Text.RegularExpressions.Regex MarkdownMetaPattern =
        new(@"([\\`*_{}\[\]()#+\-!|>~])", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string EscapeMarkdown(string s) => MarkdownMetaPattern.Replace(s, @"\$1");

    // ------------------------------------------------------------------
    // Phase 4: list items. ListItemRetriever does the numbering math; we
    // translate its result ("1.", "1.2.", "a.", "·" …) into a markdown
    // marker and indent by 2 spaces per level.
    // ------------------------------------------------------------------

    private static void EmitListItem(XElement p, EmitContext ctx)
    {
        var ilvl = (int?)p.Element(W.pPr)?.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
        var indent = new string(' ', Math.Max(0, ilvl) * 2);
        var marker = ResolveListMarker(p, ctx);
        var anchor = AnchorPrefix(p, ctx);

        ctx.Sb.Append(indent).Append(anchor).Append(marker).Append(' ');
        EmitInlineRuns(p, ctx);
        ctx.Sb.AppendLine();
        // The trailing blank line that separates list blocks from following content is
        // emitted by EmitBlocks once the run of list items ends.
    }

    /// <summary>
    /// Resolve the numbering prefix for a Heading paragraph that carries <c>w:numPr</c>.
    /// Delegates to <see cref="Internal.ListNumberResolver"/> so <see cref="DocxSession.ReplaceText"/>
    /// can use the same resolver to strip the prefix from agent payloads (otherwise an agent
    /// that echoes back what it sees in the projection ends up doubling the prefix).
    /// </summary>
    private static string? ResolveHeadingNumberPrefix(XElement p, EmitContext ctx) =>
        Internal.ListNumberResolver.Resolve(p, ctx.Document, ctx.ListItemRetrieverSettings);

    private static string ResolveListMarker(XElement p, EmitContext ctx)
    {
        if (!ctx.Settings.ResolveNumbering) return "-";
        try
        {
            var resolved = ListItemRetriever.RetrieveListItem(ctx.Document, p, ctx.ListItemRetrieverSettings);
            if (string.IsNullOrEmpty(resolved)) return "-";
            // Bullet-format levels produce a literal bullet glyph (e.g. "·" / ""); render as "-".
            if (resolved.Length == 1 && !char.IsLetterOrDigit(resolved, 0)) return "-";
            return resolved.TrimEnd();
        }
        catch
        {
            // ListItemRetriever throws on malformed numbering setups (e.g. missing
            // NumberingDefinitionsPart). Falling back to "-" matches the spec's promise
            // that "lossy by design, honestly" — we degrade visibly, never silently.
            return "-";
        }
    }

    // ------------------------------------------------------------------
    // Phase 5: tables. GFM pipe tables when the shape is simple; otherwise
    // an opaque ```table``` block referenced by the table's anchor, with cell
    // contents addressable individually via {#tc:body:…} anchors in the index.
    // ------------------------------------------------------------------

    private static void EmitTable(XElement tbl, EmitContext ctx)
    {
        var anchor = AnchorPrefix(tbl, ctx).TrimEnd();
        if (ctx.Settings.TableMode == TableRenderMode.AlwaysOpaque || !CanRenderAsGfm(tbl, ctx))
        {
            EmitOpaqueTable(tbl, anchor, ctx);
            return;
        }
        EmitGfmTable(tbl, anchor, ctx);
    }

    private static bool CanRenderAsGfm(XElement tbl, EmitContext ctx)
    {
        if (ctx.Settings.TableMode == TableRenderMode.AlwaysGfm) return true;

        // Merged cells disqualify (Word's gridSpan val>1 or any vMerge).
        if (tbl.Descendants(W.gridSpan).Any(g => ((int?)g.Attribute(W.val) ?? 1) > 1)) return false;
        if (tbl.Descendants(W.vMerge).Any()) return false;
        // Nested tables disqualify.
        if (tbl.Elements(W.tr).Elements(W.tc).Elements(W.tbl).Any()) return false;
        // Per-cell text length cap.
        var max = ctx.Settings.TableInlineCellMax;
        foreach (var tc in tbl.Elements(W.tr).Elements(W.tc))
        {
            if (CellTextRaw(tc).Length > max) return false;
        }
        return true;
    }

    private static void EmitGfmTable(XElement tbl, string anchor, EmitContext ctx)
    {
        if (anchor.Length > 0) { ctx.Sb.Append(anchor); ctx.Sb.AppendLine(); }
        var rows = tbl.Elements(W.tr).ToList();
        if (rows.Count == 0) return;

        var headerCells = rows[0].Elements(W.tc).Select(CellTextForGfm).ToList();
        ctx.Sb.Append("| ").Append(string.Join(" | ", headerCells)).AppendLine(" |");
        ctx.Sb.Append('|').Append(string.Concat(Enumerable.Repeat(" --- |", headerCells.Count))).AppendLine();
        foreach (var r in rows.Skip(1))
        {
            var cells = r.Elements(W.tc).Select(CellTextForGfm);
            ctx.Sb.Append("| ").Append(string.Join(" | ", cells)).AppendLine(" |");
        }
        ctx.Sb.AppendLine();
    }

    private static void EmitOpaqueTable(XElement tbl, string anchor, EmitContext ctx)
    {
        var rows = tbl.Elements(W.tr).Count();
        var cols = tbl.Elements(W.tr).FirstOrDefault()?.Elements(W.tc).Count() ?? 0;
        if (anchor.Length > 0) { ctx.Sb.Append(anchor); ctx.Sb.AppendLine(); }
        ctx.Sb.AppendLine("```table");
        ctx.Sb.Append("rows: ").Append(rows).AppendLine();
        ctx.Sb.Append("cols: ").Append(cols).AppendLine();
        ctx.Sb.AppendLine("```");
        ctx.Sb.AppendLine();
    }

    private static string CellTextRaw(XElement tc)
        => string.Concat(tc.Descendants(W.t).Select(t => (string)t));

    private static string CellTextForGfm(XElement tc)
    {
        // GFM pipe tables: every "|" must be escaped, every newline collapsed. Strip the
        // anchor markers — they live in the AnchorIndex but would corrupt the pipe layout.
        var raw = CellTextRaw(tc).Replace('\n', ' ').Replace('\r', ' ').Replace("|", @"\|").Trim();
        return raw.Length == 0 ? " " : raw;
    }
}
