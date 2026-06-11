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

        var index = BuildAnchorIndex(ir, settings);
        var markdown = EmitMarkdown(ir, settings);
        return new IrMarkdownResult { Markdown = markdown, AnchorIndex = index };
    }

    // ------------------------------------------------------------------
    // Anchor index (Task 1: body scope only)
    // ------------------------------------------------------------------

    private static IReadOnlyDictionary<string, AnchorTarget> BuildAnchorIndex(
        IrDocument ir, WmlToMarkdownConverterSettings settings)
    {
        var index = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);

        // Body scope only for Task 1. The body part URI is the provenance pin on the first block's
        // Source; fall back to scanning Sources for the main document part.
        var partUri = ResolveBodyPartUri(ir);

        foreach (var block in WalkBlocksForIndex(ir.Body.Blocks))
        {
            var anchor = AnchorOf(block);
            if (anchor is null) continue;

            // Suppress-mode parity: drop empty paragraphs (no visible text) from the index too.
            // Task 1 is default settings (AnchorOnly) so this branch is normally inert, but mirror
            // the oracle's BuildAnchorIndex so the field is honored when Task 2 turns it on.
            if (settings.EmptyParagraphs == EmptyParagraphMode.Suppress
                && block is IrParagraph p
                && !ParagraphHasVisibleText(p))
                continue;

            var id = anchor.Value.ToString();
            if (index.ContainsKey(id)) continue;

            var publicAnchor = ToPublicAnchor(anchor.Value);
            index[id] = new AnchorTarget
            {
                Anchor = publicAnchor,
                PartUri = partUri,
                Unid = anchor.Value.Unid,
                TextPreview = ComputeTextPreview(block),
                // AutoNumberPrefix: the oracle resolves this via ListNumberResolver's counter walk.
                // TODO(M1.4-T3): port the counter walk onto IR list facts. For Task 1 we leave it
                // null and EXCLUDE it from must-pass index comparison (documented in the harness).
                AutoNumberPrefix = null,
            };
        }

        return index;
    }

    /// <summary>Block-tree walk matching the oracle's DescendantsAndSelf order for index purposes:
    /// each block, and (for tables) the table then its rows then cells then cell blocks. Task 1 only
    /// asserts parity on paragraph anchors, but we walk tables so the index is structurally complete
    /// once Task 2 lands.</summary>
    private static IEnumerable<object> WalkBlocksForIndex(IrNodeList<IrBlock> blocks)
    {
        foreach (var b in blocks)
        {
            yield return b;
            if (b is IrTable t)
            {
                foreach (var row in t.Rows)
                {
                    yield return row;
                    foreach (var cell in row.Cells)
                    {
                        yield return cell;
                        foreach (var inner in WalkBlocksForIndex(cell.Blocks))
                            yield return inner;
                    }
                }
            }
        }
    }

    private static IrAnchor? AnchorOf(object node) => node switch
    {
        IrBlock b => b.Anchor,
        IrRow r => r.Anchor,
        IrCell c => c.Anchor,
        _ => null,
    };

    private static Anchor ToPublicAnchor(IrAnchor a) =>
        new(a.ToString(), IrAnchor.KindToken(a.Kind), a.Scope, a.Unid);

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
                // tab/break/note-ref/image/opaque: no w:t text.
            }
        }
    }

    private static bool ParagraphHasVisibleText(IrParagraph p)
    {
        var sb = new StringBuilder();
        AppendInlineText(p.Inlines, sb);
        return sb.Length > 0;
    }

    // ------------------------------------------------------------------
    // Markdown emission (Task 1: body scope only)
    // ------------------------------------------------------------------

    private static string EmitMarkdown(IrDocument ir, WmlToMarkdownConverterSettings settings)
    {
        var sb = new StringBuilder();
        var ctx = new EmitCtx { Settings = settings, Scope = "body" };

        // The body scope opens with the fixed, non-addressable "# Document" marker, then a blank line.
        sb.AppendLine("# Document");
        sb.AppendLine();

        EmitBlocks(ir.Body.Blocks, sb, ctx);

        // NOTE(Task 2/3): multipart scopes (headers/footers/notes/comments) and the scope dividers
        // are not emitted here yet. On the must-pass body-simple fixtures the oracle's other scopes
        // are empty/suppressed so this still reaches byte-equality; richer fixtures are off-list.
        return sb.ToString();
    }

    private sealed class EmitCtx
    {
        public required WmlToMarkdownConverterSettings Settings { get; init; }
        public required string Scope { get; init; }
    }

    private static void EmitBlocks(IrNodeList<IrBlock> blocks, StringBuilder sb, EmitCtx ctx)
    {
        var list = blocks.ToList();
        for (var i = 0; i < list.Count; i++)
        {
            var b = list[i];
            if (b is IrParagraph p)
            {
                var nextIsListItem = i + 1 < list.Count
                    && list[i + 1] is IrParagraph np && np.Anchor.Kind == IrAnchorKind.Li;
                EmitParagraph(p, sb, ctx);
                if (p.Anchor.Kind == IrAnchorKind.Li && !nextIsListItem)
                    sb.AppendLine();
            }
            else if (b is IrTable)
            {
                // TODO(M1.4-T2): GFM/opaque table rendering. Clearly-wrong placeholder for now.
                sb.AppendLine("<!-- IR-EMIT-UNPORTED: table -->");
                sb.AppendLine();
            }
            // IrSectionBreak (top-level): TODO(M1.4-T2) section breaks. IrOpaqueBlock: TODO(M1.4-T2).
        }
    }

    private static void EmitParagraph(IrParagraph p, StringBuilder sb, EmitCtx ctx)
    {
        var anchor = AnchorPrefix(p.Anchor, ctx);

        if (p.Anchor.Kind == IrAnchorKind.H)
        {
            var level = Math.Clamp(HeadingLevel(p) + ctx.Settings.HeadingLevelOffset, 1, 9);
            sb.Append(anchor);
            sb.Append('#', level);
            sb.Append(' ');
            // TODO(M1.4-T3): resolve heading auto-number prefix (legal clause numbering) via the
            // counter walk. Headings carrying w:numPr are therefore off the must-pass list.
            EmitInlineRuns(p, sb, ctx);
            sb.AppendLine();
            sb.AppendLine();
            // TODO(M1.4-T2): EmitInlineSectionBreak.
            return;
        }

        if (p.Anchor.Kind == IrAnchorKind.Li)
        {
            EmitListItem(p, sb, ctx);
            // TODO(M1.4-T2): EmitInlineSectionBreak.
            return;
        }

        // Plain paragraph. Default settings => EmptyParagraphMode.AnchorOnly.
        var mode = ctx.Settings.EmptyParagraphs;
        if (mode == EmptyParagraphMode.Suppress && !ParagraphHasVisibleText(p))
        {
            // TODO(M1.4-T2): EmitInlineSectionBreak even when the spacer is dropped.
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
        // TODO(M1.4-T2): EmitInlineSectionBreak.
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
    /// Resolve the list marker from IR list facts. <c>bullet</c>-format levels (and the
    /// <see cref="WmlToMarkdownConverterSettings.ResolveNumbering"/>=false case) render as <c>-</c>,
    /// matching the oracle exactly. Counter-bearing formats (decimal, lowerLetter, …) need the full
    /// counter walk the oracle performs; TODO(M1.4-T3) ports it. Until then we emit a clearly-wrong
    /// <c>?.</c> placeholder so numbered-list fixtures stay off the must-pass list and visibly differ.
    /// </summary>
    private static string ResolveListMarker(IrParagraph p, EmitCtx ctx)
    {
        if (!ctx.Settings.ResolveNumbering) return "-";
        var fmt = p.List?.NumberFormat;
        if (string.IsNullOrEmpty(fmt)) return "-";
        // The oracle renders a single-char bullet glyph (·, , etc.) as "-". The IR carries the
        // numFmt string, so a "bullet" level maps to "-" directly.
        if (fmt == "bullet" || fmt == "none") return "-";
        // TODO(M1.4-T3): counter walk for decimal/lowerLetter/upperRoman/etc.
        return "?.";
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
                    Add(tr, ReadRunFmt(tr.Format));
                    break;
                case IrFieldRun f:
                    // Field cached-result runs are visible text; emit them as plain runs (no per-run
                    // format key beyond default — TODO(M1.4-T2) thread result-run formats).
                    Flush();
                    foreach (var rr in f.CachedResult)
                        if (rr is IrTextRun ftr) Add(ftr, ReadRunFmt(ftr.Format));
                    Flush();
                    break;
                case IrTab:
                case IrBreak:
                case IrNoteRef:
                    // Whitespace/structural inlines carry no formatting toggle; group them with a
                    // default key so they emit between delimiters like the oracle's non-text runs.
                    Add(inline, default);
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
            case IrNoteRef:
                // TODO(M1.4-T3): the oracle resolves the note's Unid and emits [^fn-<suffix>]. The IR
                // IrNoteRef carries only the w:id, not the note's Unid, so a faithful label needs the
                // note store. Emit a clearly-wrong placeholder; note-ref fixtures stay off-list.
                sb.Append("[^ir-noteref]");
                break;
            // IrInlineImage / IrOpaqueInline: TODO(M1.4-T2).
        }
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

    private static string AnchorPrefix(IrAnchor anchor, EmitCtx ctx)
    {
        if (ctx.Settings.AnchorMode == AnchorRenderMode.None) return string.Empty;
        // Default settings => FullUnid rendering: the rendered token uses the full Unid verbatim.
        // TODO(M1.4-T2): Abbreviated/Sequential rendering via an AnchorIdMap equivalent.
        return $"{{#{anchor.ToString()}}} ";
    }
}
