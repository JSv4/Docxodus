#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Docxodus.Ir;

/// <summary>
/// IR-consuming reimplementation of the markdown projection (M1.4). Consumes an
/// <see cref="IrDocument"/> and produces a <see cref="MarkdownProjection"/>-shaped result that is
/// intended to be byte-equivalent to <see cref="WmlToMarkdownConverter.Convert(WmlDocument, WmlToMarkdownConverterSettings)"/>
/// — the shipped converter is the ORACLE and stays byte-untouched; this is the equivalence target.
/// </summary>
/// <remarks>
/// <para><b>Task 1 scope.</b> BODY paragraphs only with DEFAULT settings (FullUnid anchor rendering,
/// <see cref="EmptyParagraphMode.AnchorOnly"/>): headings (<c>#</c>-prefix from the pStyle heading
/// level), plain paragraphs, list items (bullet/number marker + 2-space-per-ilvl indent), block
/// anchors <c>{#kind:scope:unid}</c>, inline formatting (bold/italic/code/strike), hyperlinks, note
/// references, tabs, and breaks. Tables, images, opaque blocks, multipart scopes, section breaks,
/// auto-number HEADING prefixes, and the non-default settings modes are deliberately emitted as a
/// clearly-wrong placeholder (or skipped) here and land in Tasks 2/3. Those fixtures are simply not
/// on the must-pass list yet.</para>
///
/// <para><b>Auto-number markers (TODO(M1.4-T3)).</b> The oracle resolves list markers via
/// <c>ListItemRetriever</c>'s full counter walk against the live package. The IR carries only the
/// numbering FORMAT string (<c>bullet</c>/<c>decimal</c>/…) on <see cref="IrListInfo"/>, not the
/// resolved counter. For Task 1 we render <c>bullet</c>-format levels as <c>-</c> (which matches the
/// oracle exactly for bulleted lists) and emit a clearly-wrong <c>?.</c> placeholder for
/// counter-bearing formats — so numbered-list fixtures are off the must-pass list until the counter
/// walk is ported. Heading auto-number prefixes (legal clause numbering) are likewise stubbed.</para>
/// </remarks>
internal static class IrMarkdownEmitter
{
    /// <summary>The IR-emitter result: the markdown text plus the anchor index, mirroring
    /// <see cref="MarkdownProjection"/>. Reuses the public anchor-index types so equivalence
    /// comparisons against the oracle compare like-for-like.</summary>
    internal sealed class IrMarkdownResult
    {
        public required string Markdown { get; init; }
        public required IReadOnlyDictionary<string, AnchorTarget> AnchorIndex { get; init; }

        public MarkdownProjection ToProjection() =>
            new() { Markdown = Markdown, AnchorIndex = AnchorIndex };
    }

    private const int TextPreviewMaxLength = 80;

    public static IrMarkdownResult Emit(IrDocument ir, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(ir);
        ArgumentNullException.ThrowIfNull(settings);

        var (index, renderMap) = BuildAnchorIndex(ir, settings);
        var markdown = EmitMarkdown(ir, settings, renderMap);
        return new IrMarkdownResult { Markdown = markdown, AnchorIndex = index };
    }

    // ------------------------------------------------------------------
    // Anchor index (body scope; mirrors the oracle's BuildAnchorIndex order + AnchorIdMap)
    // ------------------------------------------------------------------

    private static (IReadOnlyDictionary<string, AnchorTarget> Index, AnchorIdMap RenderMap)
        BuildAnchorIndex(IrDocument ir, WmlToMarkdownConverterSettings settings)
    {
        var index = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);

        // Index ALL scopes in the oracle's BuildAnchorIndex order — body, hdr1.., ftr1.., fn, en, cmt —
        // so the AnchorIdMap (built from index.Values in insertion order) matches the oracle for every
        // AnchorIdRendering mode. The harness compares only BODY entries, but the markdown's note labels
        // and {#cmt:…}/{#sec:…} tokens in non-default modes route through this map.
        var autoNumber = BuildAutoNumberResolver(ir);

        // --- body ---
        var bodyPartUri = ResolveBodyPartUri(ir);
        foreach (var (anchor, preview) in WalkAnchorsForIndex(ir.Body.Blocks, settings))
            AddIndexEntry(index, anchor, bodyPartUri, preview, autoNumber);

        // --- headers / footers (each in part-enumeration order) ---
        foreach (var hf in ir.Headers.Concat(ir.Footers))
        {
            var partUri = ResolveScopePartUri(ir, hf.Scope.Blocks);
            foreach (var (anchor, preview) in WalkAnchorsForIndex(hf.Scope.Blocks, settings))
                AddIndexEntry(index, anchor, partUri, preview, autoNumber);
        }

        // --- footnotes / endnotes (the note ELEMENT is the addressable anchor) ---
        IndexNoteScope(index, ir, ir.Footnotes, IrAnchorKind.Fn);
        IndexNoteScope(index, ir, ir.Endnotes, IrAnchorKind.En);

        // --- comments ---
        if (ir.Comments.Comments.Count > 0)
        {
            var partUri = ResolveCommentsPartUri(ir);
            foreach (var c in ir.Comments.Comments)
                AddIndexEntry(index, c.Anchor, partUri, ComputeScopeTextPreview(c.Blocks), autoNumber);
        }

        // Build the AnchorIdMap. Mirror the oracle exactly: the map is constructed by iterating
        // index.Values in INSERTION order (which the walk above keeps identical to the oracle's
        // DescendantsAndSelf order), so Abbreviated prefixes and Sequential counters match byte-for-byte.
        var renderMap = BuildAnchorIdMap(index, settings);

        // Dual-key the index with the rendered id substituted (oracle parity).
        if (settings.AnchorIdRendering != AnchorIdRendering.FullUnid)
        {
            var aliases = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
            foreach (var (_, target) in index)
            {
                var rendered = renderMap.Render(target.Unid);
                if (rendered == target.Unid) continue;
                var aliasKey = $"{target.Anchor.Kind}:{target.Anchor.Scope}:{rendered}";
                aliases[aliasKey] = target;
            }
            foreach (var (key, target) in aliases)
                index[key] = target;
        }

        return (index, renderMap);
    }

    private static void AddIndexEntry(
        Dictionary<string, AnchorTarget> index, IrAnchor anchor, string partUri, string preview,
        AutoNumberResolver autoNumber)
    {
        var id = anchor.ToString();
        if (index.ContainsKey(id)) return;
        index[id] = new AnchorTarget
        {
            Anchor = ToPublicAnchor(anchor),
            PartUri = partUri,
            Unid = anchor.Unid,
            TextPreview = preview,
            // AutoNumberPrefix is resolved (heading/list numbering) only for p/h/li in the body scope,
            // matching the oracle (kind is "p"/"h"/"li" && scope.Name == "body").
            AutoNumberPrefix = anchor.Scope == "body"
                && anchor.Kind is IrAnchorKind.P or IrAnchorKind.H or IrAnchorKind.Li
                ? autoNumber.Prefix(anchor)
                : null,
        };
    }

    /// <summary>Index a note scope's note ELEMENTS (kind fn/en, scope fn/en). The note's anchor Unid
    /// comes from <see cref="IrNoteStore.NoteUnids"/>; its TextPreview is the flat text of its blocks,
    /// in the note store's insertion (document) order — mirroring the oracle's
    /// <c>scope.Root.Elements(noteName)</c> walk over real (non-boilerplate) notes.</summary>
    private static void IndexNoteScope(
        Dictionary<string, AnchorTarget> index, IrDocument ir, IrNoteStore store, IrAnchorKind kind)
    {
        if (store.Notes.Count == 0) return;
        var scope = IrAnchor.KindToken(kind);
        var partUri = ResolveScopePartUriBySuffix(ir, scope == "fn" ? "/footnotes.xml" : "/endnotes.xml");
        foreach (var (noteId, noteScope) in store.Notes)
        {
            if (!store.NoteUnids.TryGetValue(noteId, out var unid)) continue;
            var anchor = new IrAnchor(kind, scope, unid);
            AddIndexEntry(index, anchor, partUri, ComputeScopeTextPreview(noteScope.Blocks),
                AutoNumberResolver.None);
        }
    }

    private static string ComputeScopeTextPreview(IrNodeList<IrBlock> blocks)
    {
        var sb = new StringBuilder();
        foreach (var b in blocks) AppendFlatText(b, sb);
        var text = sb.ToString();
        return text.Length > TextPreviewMaxLength
            ? text.Substring(0, TextPreviewMaxLength) + "…"
            : text;
    }

    private static string ResolveScopePartUri(IrDocument ir, IrNodeList<IrBlock> blocks)
    {
        foreach (var b in blocks)
            if (b.Source.PartUri is { } uri) return uri.ToString();
        return ir.Sources.Keys.FirstOrDefault()?.ToString() ?? "/word/document.xml";
    }

    private static string ResolveScopePartUriBySuffix(IrDocument ir, string suffix)
    {
        var match = ir.Sources.Keys.FirstOrDefault(u => u.ToString().EndsWith(suffix, StringComparison.Ordinal));
        return match?.ToString() ?? bodyPartUriFallback(ir);

        static string bodyPartUriFallback(IrDocument ir) =>
            ir.Sources.Keys.FirstOrDefault(u => u.ToString().EndsWith("/document.xml", StringComparison.Ordinal))?.ToString()
            ?? "/word/document.xml";
    }

    private static string ResolveCommentsPartUri(IrDocument ir)
    {
        // Comment blocks carry their part provenance; fall back to a Sources scan by suffix.
        foreach (var c in ir.Comments.Comments)
            foreach (var b in c.Blocks)
                if (b.Source.PartUri is { } uri) return uri.ToString();
        return ResolveScopePartUriBySuffix(ir, "/comments.xml");
    }

    /// <summary>
    /// Per-projection map from full Unid → rendered id, ported from the oracle's
    /// <c>WmlToMarkdownConverter.AnchorIdMap</c>. <see cref="Render"/> returns the full Unid unchanged
    /// for <see cref="AnchorIdRendering.FullUnid"/> or an unknown Unid (defensive fallback).
    /// </summary>
    internal sealed class AnchorIdMap
    {
        private readonly Dictionary<string, string> _map = new(StringComparer.Ordinal);
        public string Render(string fullUnid) => _map.TryGetValue(fullUnid, out var r) ? r : fullUnid;
        internal void Set(string fullUnid, string renderedUnid) => _map[fullUnid] = renderedUnid;
    }

    /// <summary>Port of the oracle's AnchorIdMap construction: Abbreviated = shortest unique
    /// per-(kind,scope) prefix with a 4-char floor; Sequential = 1-based per-(kind,scope) counter in
    /// insertion (document) order; FullUnid = empty map (Render is identity).</summary>
    private static AnchorIdMap BuildAnchorIdMap(
        Dictionary<string, AnchorTarget> index, WmlToMarkdownConverterSettings settings)
    {
        var renderMap = new AnchorIdMap();
        if (settings.AnchorIdRendering == AnchorIdRendering.Abbreviated)
        {
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
                    if (n >= 32) break;
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
            var counters = new Dictionary<(string Kind, string Scope), int>();
            foreach (var t in index.Values)
            {
                var bucket = (t.Anchor.Kind, t.Anchor.Scope);
                if (!counters.TryGetValue(bucket, out var num)) num = 0;
                num++;
                counters[bucket] = num;
                renderMap.Set(t.Unid, num.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }
        return renderMap;
    }

    /// <summary>
    /// Walk the body blocks yielding each addressable anchor with its TextPreview, in the oracle's
    /// <c>DescendantsAndSelf</c> order: a paragraph (then its in-pPr <c>sec</c> if any), a table then
    /// its rows then cells then cell blocks, a standalone section break, an opaque block. Empty
    /// preview for sectPr/opaque (no <c>w:t</c> descendants), mirroring the oracle. Suppress-mode
    /// drops empty paragraphs from the index too.
    /// </summary>
    private static IEnumerable<(IrAnchor Anchor, string Preview)> WalkAnchorsForIndex(
        IrNodeList<IrBlock> blocks, WmlToMarkdownConverterSettings settings)
    {
        foreach (var b in blocks)
        {
            switch (b)
            {
                case IrParagraph p:
                    // Suppress-mode: drop empty paragraphs from the index (oracle parity). Note a
                    // paragraph whose only "text" lives in a textbox is NOT empty under the oracle's
                    // index walk: KindFor checks Descendants(w:t), which sees textbox text — so a
                    // textbox-bearing paragraph keeps its own index entry even in Suppress mode (and
                    // ParagraphHasVisibleTextOrTextbox reflects that). Its textbox inner blocks are
                    // still indexed regardless (the oracle reaches them via DescendantsAndSelf).
                    if (settings.EmptyParagraphs == EmptyParagraphMode.Suppress
                        && !ParagraphHasVisibleTextOrTextbox(p))
                    {
                        // The in-pPr sectPr is metadata, not content — it still appears in the index.
                        if (p.InlineSectionBreakAnchor is { } supSec)
                            yield return (supSec, string.Empty);
                        foreach (var inner in WalkTextboxAnchors(p, settings))
                            yield return inner;
                        break;
                    }
                    yield return (p.Anchor, ComputeTextPreview(p));
                    if (p.InlineSectionBreakAnchor is { } sec)
                        yield return (sec, string.Empty);
                    // Textbox inner blocks are addressable (the oracle's DescendantsAndSelf walk reaches
                    // them); index them right after the containing paragraph, in inline document order.
                    foreach (var inner in WalkTextboxAnchors(p, settings))
                        yield return inner;
                    break;
                case IrTable t:
                    yield return (t.Anchor, ComputeTextPreview(t));
                    foreach (var row in t.Rows)
                    {
                        yield return (row.Anchor, ComputeTextPreview(row));
                        foreach (var cell in row.Cells)
                        {
                            yield return (cell.Anchor, ComputeTextPreview(cell));
                            foreach (var inner in WalkAnchorsForIndex(cell.Blocks, settings))
                                yield return inner;
                        }
                    }
                    break;
                case IrSectionBreak s:
                    // Trailing/standalone body sectPr: indexed (empty preview), not rendered.
                    yield return (s.Anchor, string.Empty);
                    break;
                case IrOpaqueBlock o:
                    // KindFor returns null for unmodeled block elements, so the oracle does NOT index
                    // them. Match that: emit no index entry for opaque blocks.
                    break;
            }
        }
    }

    /// <summary>Yield the index anchors for every textbox inner block of <paramref name="p"/>, in
    /// inline document order, recursing through the normal block walk (so nested textboxes, tables, and
    /// their inner paragraphs are all reached) — mirroring the oracle's <c>DescendantsAndSelf</c> index
    /// walk, which descends into <c>w:txbxContent</c> inner paragraphs.</summary>
    private static IEnumerable<(IrAnchor Anchor, string Preview)> WalkTextboxAnchors(
        IrParagraph p, WmlToMarkdownConverterSettings settings)
    {
        foreach (var inline in p.Inlines)
            if (inline is IrTextbox tb)
                foreach (var inner in WalkAnchorsForIndex(tb.Blocks, settings))
                    yield return inner;
    }

    private static Anchor ToPublicAnchor(IrAnchor a) =>
        new(a.ToString(), IrAnchor.KindToken(a.Kind), a.Scope, a.Unid);

    // ------------------------------------------------------------------
    // Auto-number prefixes / list markers (port of the oracle's two display rules over the IR fact
    // IrParagraph.ResolvedListMarker, which the reader resolved against the live package). The raw
    // marker (e.g. "1.", "a.", "·", "First Article") is shared; the heading-PREFIX rule drops a single
    // non-alphanumeric glyph to null (bullets aren't meaningful heading prefixes), while the list-ITEM
    // MARKER rule keeps such a glyph as "-".
    // ------------------------------------------------------------------

    /// <summary>Resolves a block anchor to its index <c>AutoNumberPrefix</c> — the heading/list-number
    /// prefix Word renders, or null. Built once per emit from the IR's resolved markers.</summary>
    private sealed class AutoNumberResolver
    {
        public static readonly AutoNumberResolver None = new(new Dictionary<string, string?>(StringComparer.Ordinal));

        private readonly Dictionary<string, string?> _byAnchor;
        public AutoNumberResolver(Dictionary<string, string?> byAnchor) => _byAnchor = byAnchor;

        /// <summary>The AutoNumberPrefix for an index entry: the heading-prefix form of the resolved
        /// marker (single non-alnum glyph → null), or null when the paragraph carries no marker.</summary>
        public string? Prefix(IrAnchor anchor) =>
            _byAnchor.TryGetValue(anchor.ToString(), out var v) ? v : null;
    }

    /// <summary>Walk every paragraph in every scope, mapping its anchor to the heading-prefix form of
    /// its <see cref="IrParagraph.ResolvedListMarker"/> (the projection's <c>ResolveHeadingNumberPrefix</c>
    /// rule). Used for the index AutoNumberPrefix field.</summary>
    private static AutoNumberResolver BuildAutoNumberResolver(IrDocument ir)
    {
        var map = new Dictionary<string, string?>(StringComparer.Ordinal);
        void Visit(IrNodeList<IrBlock> blocks)
        {
            foreach (var b in blocks)
            {
                if (b is IrParagraph p)
                {
                    var prefix = HeadingNumberPrefix(p);
                    if (prefix != null) map[p.Anchor.ToString()] = prefix;
                    // Textbox inner paragraphs are body-scope p/h/li too, so the oracle resolves their
                    // AutoNumberPrefix as well; descend so the index field matches.
                    foreach (var inline in p.Inlines)
                        if (inline is IrTextbox tb)
                            Visit(tb.Blocks);
                }
                else if (b is IrTable t)
                {
                    foreach (var row in t.Rows)
                        foreach (var cell in row.Cells)
                            Visit(cell.Blocks);
                }
            }
        }
        Visit(ir.Body.Blocks);
        foreach (var hf in ir.Headers.Concat(ir.Footers)) Visit(hf.Scope.Blocks);
        foreach (var s in ir.Footnotes.Notes.Values) Visit(s.Blocks);
        foreach (var s in ir.Endnotes.Notes.Values) Visit(s.Blocks);
        foreach (var c in ir.Comments.Comments) Visit(c.Blocks);
        return new AutoNumberResolver(map);
    }

    /// <summary>Port of the oracle's <c>ResolveHeadingNumberPrefix</c> →
    /// <c>ListNumberResolver.Resolve</c> display rule over the IR's resolved marker: trim the trailing
    /// whitespace; a single non-alphanumeric glyph (a bullet) → null; empty → null.</summary>
    private static string? HeadingNumberPrefix(IrParagraph p)
    {
        var raw = p.ResolvedListMarker;
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var trimmed = raw.TrimEnd();
        if (trimmed.Length == 1 && !char.IsLetterOrDigit(trimmed, 0)) return null;
        return trimmed.Length == 0 ? null : trimmed;
    }

    private static string ResolveBodyPartUri(IrDocument ir)
    {
        // Prefer the provenance pin carried on the first body block.
        foreach (var b in ir.Body.Blocks)
        {
            var uri = b.Source.PartUri;
            if (uri is not null) return uri.ToString();
        }
        // Fallback: the single source whose part is the main document part. Sources is keyed by URI;
        // the main document part is conventionally "/word/document.xml".
        var main = ir.Sources.Keys.FirstOrDefault(u => u.ToString().EndsWith("/document.xml", StringComparison.Ordinal));
        return main?.ToString() ?? ir.Sources.Keys.FirstOrDefault()?.ToString() ?? "/word/document.xml";
    }

    // ------------------------------------------------------------------
    // TextPreview (mirrors the oracle's ComputeTextPreview: flat w:t concat, 80-char cap + ellipsis)
    // ------------------------------------------------------------------

    private static string ComputeTextPreview(object node)
    {
        var sb = new StringBuilder();
        AppendFlatText(node, sb);
        var text = sb.ToString();
        return text.Length > TextPreviewMaxLength
            ? text.Substring(0, TextPreviewMaxLength) + "…"
            : text;
    }

    /// <summary>Concatenate the flat text of a node exactly as the oracle's
    /// <c>string.Concat(element.Descendants(W.t))</c> would — i.e. only <c>w:t</c> text, which in the
    /// IR is the text carried by <see cref="IrTextRun"/> (including field cached-result runs and
    /// hyperlink interiors). Tabs/breaks/notes/images contribute nothing, matching <c>w:t</c>-only.</summary>
    private static void AppendFlatText(object node, StringBuilder sb)
    {
        switch (node)
        {
            case IrParagraph p:
                AppendInlineText(p.Inlines, sb);
                break;
            case IrTable t:
                foreach (var row in t.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var b in cell.Blocks)
                            AppendFlatText(b, sb);
                break;
            case IrRow r:
                foreach (var cell in r.Cells)
                    foreach (var b in cell.Blocks)
                        AppendFlatText(b, sb);
                break;
            case IrCell c:
                foreach (var b in c.Blocks)
                    AppendFlatText(b, sb);
                break;
            // IrSectionBreak / IrOpaqueBlock contribute no w:t text.
        }
    }

    private static void AppendInlineText(IrNodeList<IrInline> inlines, StringBuilder sb)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun tr:
                    sb.Append(tr.Text);
                    break;
                case IrHyperlink h:
                    AppendInlineText(h.Inlines, sb);
                    break;
                case IrFieldRun f:
                    AppendInlineText(f.CachedResult, sb);
                    break;
                case IrTextbox tb:
                    // Textbox w:t text IS visible to the oracle's Descendants(w:t) — it flows into the
                    // containing paragraph's TextPreview, into ScopeHasContent (so header/footer
                    // detection sees textbox-only content), and into table cell text. Append the inner
                    // blocks' flat text at the textbox's document position. Nested textboxes recurse
                    // through this same path (an inner paragraph's IrTextbox hits this case again).
                    foreach (var b in tb.Blocks)
                        AppendFlatText(b, sb);
                    break;
                // tab/break/note-ref/image/opaque: no w:t text.
            }
        }
    }

    /// <summary>
    /// Mirror the oracle's <c>HasVisibleInlineContent</c>: true when the paragraph would emit any
    /// visible text under the markdown grouping (which DROPS inline-SDT-spliced runs and w:fldSimple
    /// text — see GroupInlineRuns). Used by Suppress mode and the empty-paragraph trim, so it must match
    /// what is actually rendered, NOT the richer TextPreview text.
    /// </summary>
    private static bool ParagraphHasVisibleText(IrParagraph p)
    {
        foreach (var (_, runs) in GroupInlineRuns(p.Inlines))
            foreach (var r in runs)
                if (RunHasText(r))
                    return true;
        return false;
    }

    private static bool RunHasText(IrInline inline) => inline switch
    {
        IrTextRun tr => tr.Text.Length > 0,
        IrHyperlink h => h.Inlines.Any(RunHasText),
        IrFieldRun f => !f.IsSimpleField && f.CachedResult.Any(RunHasText),
        _ => false,
    };

    /// <summary>
    /// Whether the paragraph keeps its own AnchorIndex entry under Suppress mode. The oracle's index
    /// walk drops a paragraph only when <c>Descendants(w:t)</c> is empty — and that walk SEES textbox
    /// w:t text (and inline-SDT/fldSimple text, which the rendered-text predicate drops). So a
    /// paragraph whose only text lives in a textbox is NOT dropped from the index, even though its
    /// rendered markdown line is empty. This is the index-walk parity predicate, distinct from
    /// <see cref="ParagraphHasVisibleText"/> (the rendered-text predicate used for the markdown spacer).
    /// </summary>
    private static bool ParagraphHasVisibleTextOrTextbox(IrParagraph p)
    {
        var sb = new StringBuilder();
        AppendInlineText(p.Inlines, sb); // mirrors Descendants(w:t): includes textbox + SDT + field text
        return sb.Length > 0;
    }

    // ------------------------------------------------------------------
    // Markdown emission (body scope; multipart scopes land in T3)
    // ------------------------------------------------------------------

    private static string EmitMarkdown(
        IrDocument ir, WmlToMarkdownConverterSettings settings, AnchorIdMap renderMap)
    {
        var sb = new StringBuilder();
        var ctx = new EmitCtx
        {
            Settings = settings,
            Scope = "body",
            AnchorIdMap = renderMap,
            // Note ref labels resolve w:id → the note element's pt:Unid through these stores.
            Footnotes = ir.Footnotes,
            Endnotes = ir.Endnotes,
        };

        // A divider (---) separates two non-empty scopes so downstream parsers can split per-scope
        // chunks without inspecting heading text — mirroring the oracle's anyScopeEmitted bookkeeping.
        var anyScopeEmitted = false;

        // --- body: opens with the fixed, non-addressable "# Document" marker, then a blank line. ---
        ctx.Scope = "body";
        sb.AppendLine("# Document");
        sb.AppendLine();
        EmitBlocks(ir.Body.Blocks, sb, ctx);
        anyScopeEmitted = true;

        // --- headers (suppress scopes with no non-whitespace text, like the oracle's ScopeHasContent) ---
        var headerScopes = ir.Headers.Where(ScopeHasContent).ToList();
        if (headerScopes.Count > 0)
        {
            if (anyScopeEmitted) AppendScopeDivider(sb);
            sb.AppendLine("# Headers");
            sb.AppendLine();
            foreach (var hf in headerScopes)
            {
                sb.Append("## ").AppendLine(hf.ScopeName);
                sb.AppendLine();
                ctx.Scope = hf.ScopeName;
                EmitBlocks(hf.Scope.Blocks, sb, ctx);
            }
            anyScopeEmitted = true;
        }

        var footerScopes = ir.Footers.Where(ScopeHasContent).ToList();
        if (footerScopes.Count > 0)
        {
            if (anyScopeEmitted) AppendScopeDivider(sb);
            sb.AppendLine("# Footers");
            sb.AppendLine();
            foreach (var hf in footerScopes)
            {
                sb.Append("## ").AppendLine(hf.ScopeName);
                sb.AppendLine();
                ctx.Scope = hf.ScopeName;
                EmitBlocks(hf.Scope.Blocks, sb, ctx);
            }
            anyScopeEmitted = true;
        }

        // --- footnotes / endnotes ---
        if (EmitNoteDefinitions(ir.Footnotes, sb, ctx, "Footnotes", "fn", anyScopeEmitted))
            anyScopeEmitted = true;
        if (EmitNoteDefinitions(ir.Endnotes, sb, ctx, "Endnotes", "en", anyScopeEmitted))
            anyScopeEmitted = true;

        // --- comments ---
        if (EmitComments(ir.Comments, sb, ctx, anyScopeEmitted))
            anyScopeEmitted = true;

        return sb.ToString();
    }

    private static void AppendScopeDivider(StringBuilder sb)
    {
        sb.AppendLine("---");
        sb.AppendLine();
    }

    /// <summary>True when a header/footer scope has any non-whitespace text (the oracle's
    /// <c>ScopeHasContent</c>: <c>Descendants(w:t)</c> with a non-blank value).</summary>
    private static bool ScopeHasContent(IrHeaderFooter hf)
    {
        var sb = new StringBuilder();
        foreach (var b in hf.Scope.Blocks) AppendFlatText(b, sb);
        for (int i = 0; i < sb.Length; i++)
            if (!char.IsWhiteSpace(sb[i])) return true;
        return false;
    }

    /// <summary>
    /// Port of the oracle's <c>EmitNoteDefinitions</c>: a <c># Footnotes</c>/<c># Endnotes</c> header
    /// (preceded by a divider when content already emitted), then one <c>[^fn-…]: …</c>/<c>[^en-…]: …</c>
    /// line per real note, the note's paragraphs flattened inline and joined by a single space. The
    /// label suffix derives from the note element's <c>pt:Unid</c> (<see cref="IrNoteStore.NoteUnids"/>).
    /// Notes are walked in the store's insertion (document) order. Returns true if anything was emitted.
    /// </summary>
    private static bool EmitNoteDefinitions(
        IrNoteStore store, StringBuilder sb, EmitCtx ctx, string header, string kindPrefix,
        bool precedingContent)
    {
        if (store.Notes.Count == 0) return false;

        if (precedingContent) AppendScopeDivider(sb);
        sb.Append("# ").AppendLine(header);
        sb.AppendLine();
        ctx.Scope = kindPrefix;
        foreach (var (noteId, noteScope) in store.Notes)
        {
            var unid = store.NoteUnids.TryGetValue(noteId, out var u) ? u : "0";
            var label = $"{kindPrefix}-{NoteLabelSuffix(unid, ctx)}";
            sb.Append("[^").Append(label).Append("]: ");
            var first = true;
            foreach (var p in noteScope.Blocks.OfType<IrParagraph>())
            {
                if (!first) sb.Append(' ');
                first = false;
                EmitInlineRuns(p, sb, ctx);
            }
            sb.AppendLine();
            sb.AppendLine();
        }
        return true;
    }

    /// <summary>
    /// Port of the oracle's <c>EmitComments</c>: a <c># Comments</c> header (preceded by a divider when
    /// content already emitted), then one <c>- {#cmt:cmt:…} **author** (date): …</c> line per comment,
    /// each comment paragraph flattened inline with a trailing space, closing with a blank line.
    /// Returns true if anything was emitted.
    /// </summary>
    private static bool EmitComments(
        IrCommentStore store, StringBuilder sb, EmitCtx ctx, bool precedingContent)
    {
        if (store.Comments.Count == 0) return false;

        if (precedingContent) AppendScopeDivider(sb);
        sb.AppendLine("# Comments");
        sb.AppendLine();
        ctx.Scope = "cmt";
        foreach (var c in store.Comments)
        {
            var author = string.IsNullOrEmpty(c.Author) ? "unknown" : c.Author;
            var renderedUnid = ctx.AnchorIdMap.Render(c.Anchor.Unid);
            sb.Append($"- {{#cmt:cmt:{renderedUnid}}} **{author}**");
            if (!string.IsNullOrEmpty(c.Date)) sb.Append(" (").Append(c.Date).Append(')');
            sb.Append(": ");
            foreach (var p in c.Blocks.OfType<IrParagraph>())
            {
                EmitInlineRuns(p, sb, ctx);
                sb.Append(' ');
            }
            sb.AppendLine();
        }
        sb.AppendLine();
        return true;
    }

    private static string ShortUnid(string unid) =>
        unid.Length >= 8 ? unid.Substring(0, 8) : unid;

    /// <summary>Port of the oracle's <c>NoteLabelSuffix</c>: FullUnid mode keeps the legacy 8-char
    /// truncation; Abbreviated/Sequential route through the AnchorIdMap for consistency with the
    /// corresponding <c>{#fn:…}</c> token.</summary>
    private static string NoteLabelSuffix(string unid, EmitCtx ctx) =>
        ctx.Settings.AnchorIdRendering == AnchorIdRendering.FullUnid
            ? ShortUnid(unid)
            : ctx.AnchorIdMap.Render(unid);

    private sealed class EmitCtx
    {
        public required WmlToMarkdownConverterSettings Settings { get; init; }
        public required string Scope { get; set; }
        public required AnchorIdMap AnchorIdMap { get; init; }
        public IrNoteStore Footnotes { get; init; } = IrNoteStore.Empty;
        public IrNoteStore Endnotes { get; init; } = IrNoteStore.Empty;
    }

    private static void EmitBlocks(IrNodeList<IrBlock> blocks, StringBuilder sb, EmitCtx ctx)
    {
        // Skip blocks a block-level w:sdt delivered: the oracle's EmitBlocks dispatches only direct
        // w:p/w:tbl/w:sectPr children and silently skips the w:sdt wrapper, so it never renders these.
        // They stay in the IR + anchor index (the oracle indexes them via Descendants).
        var list = blocks.Where(b => !b.Source.FromBlockSdt).ToList();
        for (var i = 0; i < list.Count; i++)
        {
            var b = list[i];
            if (b is IrParagraph p)
            {
                // The trailing-blank-line rule keys on the oracle's IsListItem (numPr present, inline
                // OR via the pStyle chain) — NOT the anchor kind alone. Two cases the IR must union to
                // match it: (a) a plain list item is anchor-kind "li"; (b) a Heading{N} paragraph that
                // ALSO carries numPr is anchor-kind "h" but IS a list item. p.List carries resolved
                // numbering facts but is null for some style-inherited empty list items, so OR it with
                // the "li" anchor kind — together they reproduce the oracle's IsListItem exactly.
                var nextIsListItem = i + 1 < list.Count
                    && list[i + 1] is IrParagraph np && IsListItemForBlankRule(np);
                EmitParagraph(p, sb, ctx);
                if (IsListItemForBlankRule(p) && !nextIsListItem)
                    sb.AppendLine();
            }
            else if (b is IrTable t)
            {
                EmitTable(t, sb, ctx);
            }
            // IrSectionBreak (standalone/trailing body sectPr): the oracle's EmitBlocks treats a
            // top-level w:sectPr as a no-op (it is last-section metadata, not a transition). Match it.
            // IrOpaqueBlock: the oracle's EmitBlocks only dispatches w:p/w:tbl/w:sectPr — any other
            // top-level element is silently skipped. Match that too (no markdown, no index entry).
        }
    }

    /// <summary>The IR proxy for the oracle's <c>IsListItem</c>, used by the EmitBlocks trailing-blank
    /// rule: a paragraph is a list item iff its anchor kind is <c>li</c> OR it carries resolved
    /// numbering facts (<see cref="IrParagraph.List"/> — true for a Heading-with-numPr).</summary>
    private static bool IsListItemForBlankRule(IrParagraph p) =>
        p.Anchor.Kind == IrAnchorKind.Li || p.List is not null;

    private static void EmitParagraph(IrParagraph p, StringBuilder sb, EmitCtx ctx)
    {
        var anchor = AnchorPrefix(p.Anchor, ctx);

        if (p.Anchor.Kind == IrAnchorKind.H)
        {
            var level = Math.Clamp(HeadingLevel(p) + ctx.Settings.HeadingLevelOffset, 1, 9);
            sb.Append(anchor);
            sb.Append('#', level);
            sb.Append(' ');
            // Legal-style headings often style each clause Heading{N} AND attach numbering so Word
            // renders "First …" / "1.1 …". Emit that prefix (the heading-prefix form of the reader-
            // resolved marker) so it survives projection, mirroring ResolveHeadingNumberPrefix.
            if (ctx.Settings.ResolveNumbering)
            {
                var numberPrefix = HeadingNumberPrefix(p);
                if (numberPrefix != null) sb.Append(numberPrefix).Append(' ');
            }
            EmitInlineRuns(p, sb, ctx);
            sb.AppendLine();
            sb.AppendLine();
            EmitInlineSectionBreak(p, sb, ctx);
            return;
        }

        if (p.Anchor.Kind == IrAnchorKind.Li)
        {
            EmitListItem(p, sb, ctx);
            EmitInlineSectionBreak(p, sb, ctx);
            return;
        }

        // Plain paragraph. Default settings => EmptyParagraphMode.AnchorOnly.
        var mode = ctx.Settings.EmptyParagraphs;
        if (mode == EmptyParagraphMode.Suppress && !ParagraphHasVisibleText(p))
        {
            // The spacer is dropped, but a section transition is metadata, not content — still emit it.
            EmitInlineSectionBreak(p, sb, ctx);
            return;
        }

        var beforeAnchor = sb.Length;
        sb.Append(anchor);
        var afterAnchor = sb.Length;
        EmitInlineRuns(p, sb, ctx);
        if (sb.Length == afterAnchor && afterAnchor > beforeAnchor)
        {
            // No visible runs emitted: honor empty-paragraph mode (default trims the dangling space).
            if (mode == EmptyParagraphMode.MarkedEmpty)
                sb.Append('∅');
            else if (sb[sb.Length - 1] == ' ')
                sb.Length--;
        }
        sb.AppendLine();
        sb.AppendLine();
        EmitInlineSectionBreak(p, sb, ctx);
    }

    /// <summary>
    /// Mirror the oracle's <c>EmitInlineSectionBreak</c>: when the paragraph carries an in-pPr
    /// <c>w:sectPr</c> (captured by the reader as <see cref="IrParagraph.InlineSectionBreakAnchor"/>),
    /// emit the section anchor token (unless AnchorMode==None) followed by a <c>---</c> thematic break.
    /// </summary>
    private static void EmitInlineSectionBreak(IrParagraph p, StringBuilder sb, EmitCtx ctx)
    {
        if (p.InlineSectionBreakAnchor is not { } sec) return;
        if (ctx.Settings.AnchorMode != AnchorRenderMode.None)
        {
            var rendered = ctx.AnchorIdMap.Render(sec.Unid);
            sb.Append("{#sec:").Append(ctx.Scope).Append(':').Append(rendered).AppendLine("}");
        }
        sb.AppendLine("---");
        sb.AppendLine();
    }

    private static void EmitListItem(IrParagraph p, StringBuilder sb, EmitCtx ctx)
    {
        var ilvl = p.List?.Ilvl ?? 0;
        var indent = new string(' ', Math.Max(0, ilvl) * 2);
        var marker = ResolveListMarker(p, ctx);
        var anchor = AnchorPrefix(p.Anchor, ctx);

        sb.Append(indent).Append(anchor).Append(marker).Append(' ');
        EmitInlineRuns(p, sb, ctx);
        sb.AppendLine();
        // Trailing blank line between a list block and following content is emitted by EmitBlocks.
    }

    /// <summary>
    /// Resolve the list marker, porting the oracle's <c>ResolveListMarker</c> display rule over the
    /// IR's <see cref="IrParagraph.ResolvedListMarker"/> (resolved by the reader against the live
    /// package). When <see cref="WmlToMarkdownConverterSettings.ResolveNumbering"/> is false, or no
    /// marker resolved, render <c>-</c>; a single non-alphanumeric glyph (a bullet) renders <c>-</c>;
    /// otherwise the trimmed marker (e.g. <c>1.</c>, <c>a.</c>, <c>1.1</c>) renders verbatim.
    /// </summary>
    private static string ResolveListMarker(IrParagraph p, EmitCtx ctx)
    {
        if (!ctx.Settings.ResolveNumbering) return "-";
        var raw = p.ResolvedListMarker;
        if (string.IsNullOrEmpty(raw)) return "-";
        var trimmed = raw.TrimEnd();
        if (trimmed.Length == 0) return "-";
        if (trimmed.Length == 1 && !char.IsLetterOrDigit(trimmed, 0)) return "-";
        return trimmed;
    }

    // ------------------------------------------------------------------
    // Inline runs — mirrors the oracle's GroupInlineRuns + EmitInlineRuns,
    // consuming the already-walked IR inline list (revisions accepted, fields
    // flattened, SDTs spliced, runs coalesced by the reader).
    // ------------------------------------------------------------------

    private readonly record struct RunFmt(bool Bold, bool Italic, bool Code, bool Strike, string? HyperlinkUrl);

    private static void EmitInlineRuns(IrParagraph p, StringBuilder sb, EmitCtx ctx)
    {
        foreach (var (fmt, runs) in GroupInlineRuns(p.Inlines))
        {
            if (fmt.HyperlinkUrl != null)
            {
                sb.Append('[');
                foreach (var r in runs) AppendRunText(r, sb, ctx);
                sb.Append("](").Append(fmt.HyperlinkUrl).Append(')');
                continue;
            }

            var (open, close) = MarkdownDelimiters(fmt);
            sb.Append(open);
            foreach (var r in runs) AppendRunText(r, sb, ctx);
            sb.Append(close);
        }
    }

    /// <summary>
    /// Group the paragraph's inline children into runs of shared formatting, mirroring the oracle's
    /// <c>GroupInlineRuns</c>: hyperlinks each form their own group; adjacent same-format text runs
    /// merge. The IR has already coalesced same-format <see cref="IrTextRun"/>s, but we regroup here
    /// because the oracle's RunFmt comparison key (bold/italic/code/strike/url) is COARSER than the
    /// IR's full-format coalescing key — two runs the IR kept separate (e.g. differing color) still
    /// merge under the markdown key.
    /// </summary>
    private static List<(RunFmt Fmt, List<IrInline> Runs)> GroupInlineRuns(IrNodeList<IrInline> inlines)
    {
        var groups = new List<(RunFmt, List<IrInline>)>();
        var buf = new List<IrInline>();
        RunFmt bufFmt = default;
        var primed = false;

        void Flush()
        {
            if (primed && buf.Count > 0)
                groups.Add((bufFmt, new List<IrInline>(buf)));
            buf.Clear();
            primed = false;
        }

        void Add(IrInline inline, RunFmt fmt)
        {
            if (!primed)
            {
                bufFmt = fmt;
                buf.Add(inline);
                primed = true;
                return;
            }
            if (fmt.HyperlinkUrl == null && bufFmt.HyperlinkUrl == null && fmt.Equals(bufFmt))
            {
                buf.Add(inline);
                return;
            }
            Flush();
            bufFmt = fmt;
            buf.Add(inline);
            primed = true;
        }

        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrHyperlink h:
                    Flush();
                    var url = h.Target;
                    // The oracle groups each hyperlink's interior runs under the link url and flushes
                    // around it. We treat the whole hyperlink as one group carrying its inner runs.
                    Add(h, new RunFmt(false, false, false, false, url));
                    Flush();
                    break;
                case IrTextRun tr:
                    // Inline-SDT/smartTag-spliced runs are DROPPED from the markdown (the oracle's
                    // GroupInlineRuns never visits a w:sdt), though they still count toward TextPreview
                    // (AppendInlineText keeps them, mirroring the oracle's Descendants(w:t)).
                    if (tr.FromInlineSdt) break;
                    Add(tr, ReadRunFmt(tr.Format));
                    break;
                case IrFieldRun f:
                    // A w:fldSimple is DROPPED from the markdown (the oracle's GroupInlineRuns walks
                    // only w:r/w:hyperlink/w:ins/w:del, never w:fldSimple) — though its text still
                    // counts toward TextPreview/cell text (Descendants(w:t)), handled in AppendInlineText.
                    if (f.IsSimpleField) break;
                    // Run-based field result runs ARE direct w:r children → the oracle emits them.
                    Flush();
                    foreach (var rr in f.CachedResult)
                        if (rr is IrTextRun ftr) Add(ftr, ReadRunFmt(ftr.Format));
                    Flush();
                    break;
                case IrTab tab:
                    // A w:tab is a child of a w:r, so the oracle groups it under THAT run's formatting
                    // (e.g. an italic run's tab lands inside the *…* span). The IR carries the run's
                    // format on the tab, so group by it — matching the oracle's per-run grouping.
                    Add(tab, ReadRunFmt(tab.Format));
                    break;
                case IrBreak:
                case IrNoteRef:
                    // Breaks/note refs carry no run-format on the IR node; group them with a default
                    // key. (A w:br inside a formatted run is a rare delimiter-placement edge case left
                    // to triage; note refs carry no formatting toggle.)
                    Add(inline, default);
                    break;
                case IrTextbox:
                    // DROPPED from the rendered markdown: the oracle's GroupInlineRuns walks only
                    // w:r/w:hyperlink/w:ins/w:del and never descends into a w:drawing/w:pict, so textbox
                    // content produces no markdown. Its text still counts toward TextPreview/cell text
                    // (AppendInlineText) and its inner blocks are still indexed (WalkTextboxAnchors).
                    break;
                // IrInlineImage / IrOpaqueInline: TODO(M1.4-T2). Skipped here.
            }
        }
        Flush();
        return groups;
    }

    private static RunFmt ReadRunFmt(IrRunFormat f) =>
        new(
            Bold: f.Bold == true,
            Italic: f.Italic == true,
            Code: IsCodeRun(f),
            Strike: f.Strike == true,
            HyperlinkUrl: null);

    /// <summary>Mirror the oracle's <c>IsCodeRun</c>: a Code/HTMLCode/VerbatimChar character style, or
    /// a monospace ascii font (Mono/Courier/Consolas).</summary>
    private static bool IsCodeRun(IrRunFormat f)
    {
        var styleId = f.StyleId;
        if (styleId != null &&
            (styleId.Equals("Code", StringComparison.OrdinalIgnoreCase)
             || styleId.Equals("HTMLCode", StringComparison.OrdinalIgnoreCase)
             || styleId.Equals("VerbatimChar", StringComparison.OrdinalIgnoreCase)))
            return true;
        var ascii = f.FontAscii;
        if (ascii != null && (ascii.Contains("Mono", StringComparison.OrdinalIgnoreCase)
            || ascii.Contains("Courier", StringComparison.OrdinalIgnoreCase)
            || ascii.Contains("Consolas", StringComparison.OrdinalIgnoreCase)))
            return true;
        return false;
    }

    private static (string Open, string Close) MarkdownDelimiters(RunFmt fmt)
    {
        if (fmt.Code) return ("`", "`");
        var open = new StringBuilder();
        var close = new StringBuilder();
        if (fmt.Strike) { open.Append("~~"); close.Insert(0, "~~"); }
        if (fmt.Bold) { open.Append("**"); close.Insert(0, "**"); }
        if (fmt.Italic) { open.Append('*'); close.Insert(0, '*'); }
        return (open.ToString(), close.ToString());
    }

    /// <summary>Append a single inline's text, escaped, mirroring the oracle's <c>AppendRunText</c>:
    /// text/delText escaped, <c>w:br</c> → hard break, <c>w:tab</c> → 4 spaces, note refs →
    /// <c>[^fn-…]</c>/<c>[^en-…]</c>. Hyperlink interiors recurse to their text runs.</summary>
    private static void AppendRunText(IrInline inline, StringBuilder sb, EmitCtx ctx)
    {
        switch (inline)
        {
            case IrTextRun tr:
                sb.Append(EscapeMarkdown(tr.Text));
                break;
            case IrHyperlink h:
                foreach (var inner in h.Inlines)
                    AppendRunText(inner, sb, ctx);
                break;
            case IrFieldRun f:
                // A w:fldSimple is not rendered (see GroupInlineRuns); run-based field result recurses.
                if (f.IsSimpleField) break;
                foreach (var inner in f.CachedResult)
                    AppendRunText(inner, sb, ctx);
                break;
            case IrBreak br when br.Kind == IrBreakKind.Line:
                sb.Append("  \n");
                break;
            case IrBreak:
                // Page/column breaks: the oracle only special-cases w:br as a hard line break.
                sb.Append("  \n");
                break;
            case IrTab:
                sb.Append("    ");
                break;
            case IrNoteRef nr:
                // Port of the oracle's AppendNoteRefMarker: resolve w:id → the note element's pt:Unid
                // (through the note store) → label suffix → [^fn-<suffix>] / [^en-<suffix>]. A ref to a
                // missing/boilerplate note (not in the store) emits nothing, matching the oracle's
                // unid==null early return.
                AppendNoteRefMarker(nr, sb, ctx);
                break;
            // IrInlineImage / IrOpaqueInline: TODO(M1.4-T2).
        }
    }

    /// <summary>Mirror the oracle's <c>AppendNoteRefMarker</c>: resolve the note ref's <c>w:id</c>
    /// through the matching note store to the note element's Unid, then emit
    /// <c>[^fn-&lt;suffix&gt;]</c>/<c>[^en-&lt;suffix&gt;]</c>. Emits nothing when the id resolves to no
    /// real note (the oracle returns early when the Unid is null — e.g. a ref to a boilerplate or
    /// missing note).</summary>
    private static void AppendNoteRefMarker(IrNoteRef nr, StringBuilder sb, EmitCtx ctx)
    {
        var (store, prefix) = nr.Kind == IrNoteKind.Footnote
            ? (ctx.Footnotes, "fn")
            : (ctx.Endnotes, "en");
        if (!store.NoteUnids.TryGetValue(nr.NoteId, out var unid)) return;
        sb.Append("[^").Append(prefix).Append('-').Append(NoteLabelSuffix(unid, ctx)).Append(']');
    }

    private static readonly System.Text.RegularExpressions.Regex MarkdownMetaPattern =
        new(@"([\\`*_{}\[\]()#+\-!|>~])", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string EscapeMarkdown(string s) => MarkdownMetaPattern.Replace(s, @"\$1");

    // ------------------------------------------------------------------
    // Heading level + anchor prefix (mirror the oracle exactly)
    // ------------------------------------------------------------------

    /// <summary>Mirror the oracle's <c>HeadingLevel</c>: Title→1, Subtitle→2, else the digits in the
    /// style id clamped to 1..9 (default 1).</summary>
    private static int HeadingLevel(IrParagraph p)
    {
        var styleId = p.Format.StyleId ?? string.Empty;
        if (styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)) return 1;
        if (styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)) return 2;
        var digits = new string(styleId.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out var n) && n >= 1 && n <= 9 ? n : 1;
    }

    /// <summary>Build the block anchor prefix (with trailing space), substituting the rendered Unid
    /// per the AnchorIdMap, or empty string when AnchorMode==None. Mirrors the oracle's
    /// <c>AnchorPrefix</c> (which renders <c>{#kind:scope:rendered}</c>).</summary>
    private static string AnchorPrefix(IrAnchor anchor, EmitCtx ctx)
    {
        if (ctx.Settings.AnchorMode == AnchorRenderMode.None) return string.Empty;
        var rendered = ctx.AnchorIdMap.Render(anchor.Unid);
        return $"{{#{IrAnchor.KindToken(anchor.Kind)}:{anchor.Scope}:{rendered}}} ";
    }

    // ------------------------------------------------------------------
    // Tables — ported from the oracle's EmitTable / CanRenderAsGfm / EmitGfmTable /
    // EmitOpaqueTable / CellTextForGfm / CellTextRaw. Simple tables → GFM pipe tables;
    // merges/nesting/over-long cells → an opaque ```table rows/cols summary.
    // ------------------------------------------------------------------

    private static void EmitTable(IrTable tbl, StringBuilder sb, EmitCtx ctx)
    {
        // The oracle takes AnchorPrefix(tbl).TrimEnd() — the {#tbl:…} token without the trailing space.
        var anchor = AnchorPrefix(tbl.Anchor, ctx).TrimEnd();
        if (ctx.Settings.TableMode == TableRenderMode.AlwaysOpaque || !CanRenderAsGfm(tbl, ctx))
        {
            EmitOpaqueTable(tbl, anchor, sb);
            return;
        }
        EmitGfmTable(tbl, anchor, sb, ctx);
    }

    /// <summary>
    /// Port of the oracle's <c>CanRenderAsGfm</c> simplicity predicate: any gridSpan&gt;1 or any
    /// vMerge disqualifies; a nested table in any cell disqualifies; any first-level cell whose raw
    /// text exceeds <see cref="WmlToMarkdownConverterSettings.TableInlineCellMax"/> disqualifies.
    /// <see cref="TableRenderMode.AlwaysGfm"/> bypasses all checks.
    /// </summary>
    private static bool CanRenderAsGfm(IrTable tbl, EmitCtx ctx)
    {
        if (ctx.Settings.TableMode == TableRenderMode.AlwaysGfm) return true;

        var max = ctx.Settings.TableInlineCellMax;
        // Merged-cell checks mirror the oracle's tbl.Descendants(w:gridSpan)/Descendants(w:vMerge):
        // they reach into SDT-delivered rows/cells, so probe ALL rows and ALL cells here.
        foreach (var row in tbl.Rows)
            foreach (var cell in row.Cells)
            {
                if (cell.GridSpan > 1) return false;
                if (cell.VMerge != IrVMerge.None) return false;
            }
        // Nested-table and per-cell length checks mirror tbl.Elements(w:tr).Elements(w:tc)[.Elements(w:tbl)]
        // — direct rows AND direct cells only — so a nested table or over-long cell that an SDT delivers
        // does NOT disqualify GFM. Restrict to the oracle-visible rows/cells.
        foreach (var row in OracleVisibleRows(tbl))
            foreach (var cell in OracleVisibleCells(row))
            {
                if (cell.Blocks.Any(b => b is IrTable)) return false;
                if (CellTextRaw(cell).Length > max) return false;
            }
        return true;
    }

    /// <summary>The rows the oracle's table walk SEES — direct <c>w:tr</c> children only
    /// (<c>tbl.Elements(w:tr)</c>). Excludes rows a table-level <c>w:sdt</c> delivered
    /// (<see cref="IrRow.FromTableSdt"/>), which the IR keeps + indexes for fidelity but the oracle's
    /// table markdown never renders.</summary>
    private static IEnumerable<IrRow> OracleVisibleRows(IrTable tbl) =>
        tbl.Rows.Where(r => !r.FromTableSdt);

    /// <summary>The cells of a row the oracle's table walk SEES — direct <c>w:tc</c> children only
    /// (<c>Elements(w:tr).Elements(w:tc)</c>). Excludes cells a row-level <c>w:sdt</c> delivered
    /// (<see cref="IrCell.FromRowSdt"/>), which the IR keeps for content fidelity but the oracle's
    /// table markdown never renders.</summary>
    private static IEnumerable<IrCell> OracleVisibleCells(IrRow row) =>
        row.Cells.Where(c => !c.FromRowSdt);

    private static void EmitGfmTable(IrTable tbl, string anchor, StringBuilder sb, EmitCtx ctx)
    {
        if (anchor.Length > 0) { sb.Append(anchor); sb.AppendLine(); }
        var rows = OracleVisibleRows(tbl).ToList();
        if (rows.Count == 0) return;

        var headerCells = OracleVisibleCells(rows[0]).Select(CellTextForGfm).ToList();
        sb.Append("| ").Append(string.Join(" | ", headerCells)).AppendLine(" |");
        sb.Append('|').Append(string.Concat(Enumerable.Repeat(" --- |", headerCells.Count))).AppendLine();
        foreach (var r in rows.Skip(1))
        {
            var cells = OracleVisibleCells(r).Select(CellTextForGfm);
            sb.Append("| ").Append(string.Join(" | ", cells)).AppendLine(" |");
        }
        sb.AppendLine();
    }

    private static void EmitOpaqueTable(IrTable tbl, string anchor, StringBuilder sb)
    {
        // rows/cols mirror the oracle's tbl.Elements(w:tr).Count() and the first direct row's direct
        // cell count — SDT-delivered rows/cells are excluded from both.
        var visibleRows = OracleVisibleRows(tbl).ToList();
        var rows = visibleRows.Count;
        var cols = visibleRows.FirstOrDefault() is { } first ? OracleVisibleCells(first).Count() : 0;
        if (anchor.Length > 0) { sb.Append(anchor); sb.AppendLine(); }
        sb.AppendLine("```table");
        sb.Append("rows: ").Append(rows).AppendLine();
        sb.Append("cols: ").Append(cols).AppendLine();
        sb.AppendLine("```");
        sb.AppendLine();
    }

    /// <summary>Raw flat cell text — the <c>w:t</c> concat, mirroring the oracle's <c>CellTextRaw</c>
    /// (which is <c>Descendants(W.t)</c>). In the IR that is the cell's flat inline text.</summary>
    private static string CellTextRaw(IrCell cell)
    {
        var sb = new StringBuilder();
        foreach (var b in cell.Blocks) AppendFlatText(b, sb);
        return sb.ToString();
    }

    /// <summary>Port of the oracle's <c>CellTextForGfm</c>: collapse newlines to spaces, escape pipes,
    /// trim, and substitute a single space for an empty cell so the pipe layout stays well-formed.</summary>
    private static string CellTextForGfm(IrCell cell)
    {
        var raw = CellTextRaw(cell).Replace('\n', ' ').Replace('\r', ' ').Replace("|", @"\|").Trim();
        return raw.Length == 0 ? " " : raw;
    }
}
