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
}

/// <summary>
/// Identifies a single addressable element in the markdown projection.
/// Anchor ids have the form <c>kind:scope:unid</c> (e.g. <c>p:body:a1b2c3d4</c>).
/// </summary>
public readonly record struct Anchor(string Id, string Kind, string Scope, string Unid)
{
    /// <summary>The block-level token rendered in the projection (e.g. <c>{#p:body:a1b2c3d4}</c>).</summary>
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

        var (index, scopes) = BuildAnchorIndex(document, settings);
        var markdown = EmitMarkdown(document, settings, scopes);
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

    private static (IReadOnlyDictionary<string, AnchorTarget> Index, List<ScopeInfo> Scopes)
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
            UnidHelper.AssignToAllElements(scope.Root);
            // Stash the owning part on the root so downstream emitters (hyperlinks, etc.)
            // can resolve relationship-bound URIs without threading the part through every call.
            if (scope.Root.Annotation<OpenXmlPart>() == null)
                scope.Root.AddAnnotation(scope.Part);
            foreach (var el in scope.Root.DescendantsAndSelf())
            {
                var kind = KindFor(el);
                if (kind == null) continue;
                var unid = (string?)el.Attribute(PtOpenXml.Unid);
                if (unid == null) continue;
                var id = $"{kind}:{scope.Name}:{unid}";
                if (index.ContainsKey(id)) continue;
                var anchor = new Anchor(id, kind, scope.Name, unid);
                index[id] = new AnchorTarget
                {
                    Anchor = anchor,
                    PartUri = scope.Part.Uri.ToString(),
                    Unid = unid,
                };
            }
            scope.Part.PutXDocument();
        }

        return (index, scopes);
    }

    /// <summary>
    /// Classify an element to its anchor <c>kind</c>. Returns <c>null</c> for elements that
    /// are not addressable by the projection (runs, inline children, formatting properties).
    /// </summary>
    private static string? KindFor(XElement el)
    {
        var n = el.Name;
        if (n == W.p) return IsHeading(el) ? "h" : IsListItem(el) ? "li" : "p";
        if (n == W.tbl) return "tbl";
        if (n == W.tr) return "tr";
        if (n == W.tc) return "tc";
        if (n == W.footnote) return "fn";
        if (n == W.endnote) return "en";
        if (n == W.comment) return "cmt";
        return null;
    }

    private static bool IsHeading(XElement p)
    {
        var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (string.IsNullOrEmpty(styleId)) return false;
        return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            || styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
            || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsListItem(XElement p)
        => p.Element(W.pPr)?.Element(W.numPr) != null;

    // ------------------------------------------------------------------
    // Markdown emission. Phase 2+ implementation. Per-element handlers are
    // dispatched by element name from EmitBlocks; the EmitContext threads the
    // current scope name and a StringBuilder so handlers stay pure of state.
    // ------------------------------------------------------------------

    private sealed class EmitContext
    {
        public StringBuilder Sb { get; } = new();
        required public WmlToMarkdownConverterSettings Settings { get; init; }
        public string Scope { get; set; } = "body";
    }

    private static string EmitMarkdown(
        WordprocessingDocument document,
        WmlToMarkdownConverterSettings settings,
        List<ScopeInfo> scopes)
    {
        var ctx = new EmitContext { Settings = settings };

        var bodyScope = scopes.FirstOrDefault(s => s.Name == "body");
        if (bodyScope != null)
        {
            ctx.Sb.AppendLine("# Document");
            ctx.Sb.AppendLine();
            ctx.Scope = "body";
            var body = bodyScope.Root.Element(W.body);
            if (body != null) EmitBlocks(body.Elements(), ctx);
        }

        // Phases 6+ append headers/footers/footnotes/endnotes/comments here.

        return ctx.Sb.ToString();
    }

    private static void EmitBlocks(IEnumerable<XElement> blocks, EmitContext ctx)
    {
        foreach (var b in blocks)
        {
            if (b.Name == W.p) EmitParagraph(b, ctx);
            else if (b.Name == W.tbl) EmitTable(b, ctx);
            else if (b.Name == W.sectPr) { /* phase 6: section breaks */ }
        }
    }

    private static void EmitParagraph(XElement p, EmitContext ctx)
    {
        var anchor = AnchorPrefix(p, ctx);

        if (IsHeading(p))
        {
            var level = Math.Clamp(HeadingLevel(p) + ctx.Settings.HeadingLevelOffset, 1, 6);
            ctx.Sb.Append(anchor);
            ctx.Sb.Append(new string('#', level));
            ctx.Sb.Append(' ');
            EmitInlineRuns(p, ctx);
            ctx.Sb.AppendLine();
            ctx.Sb.AppendLine();
            return;
        }

        if (IsListItem(p))
        {
            EmitListItem(p, ctx);
            return;
        }

        ctx.Sb.Append(anchor);
        EmitInlineRuns(p, ctx);
        ctx.Sb.AppendLine();
        ctx.Sb.AppendLine();
    }

    /// <summary>Build the anchor prefix (with trailing space) for a block element, or empty string when AnchorMode==None.</summary>
    private static string AnchorPrefix(XElement el, EmitContext ctx)
    {
        if (ctx.Settings.AnchorMode == AnchorRenderMode.None) return string.Empty;
        var kind = KindFor(el) ?? "unk";
        var unid = (string?)el.Attribute(PtOpenXml.Unid) ?? "0";
        return $"{{#{kind}:{ctx.Scope}:{unid}}} ";
    }

    private static int HeadingLevel(XElement p)
    {
        var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val) ?? string.Empty;
        if (styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)) return 1;
        if (styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)) return 2;
        var digits = new string(styleId.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out var n) && n >= 1 && n <= 6 ? n : 1;
    }

    // ------------------------------------------------------------------
    // Phase 3: inline runs. Adjacent runs sharing the same formatting are
    // merged into one delimiter pair so "**a****b**" becomes "**ab**". Each
    // run's text is escaped before delimiters are wrapped around it.
    // ------------------------------------------------------------------

    private readonly record struct RunFormatting(
        bool Bold,
        bool Italic,
        bool Code,
        bool Strike,
        string? HyperlinkUrl);

    private static void EmitInlineRuns(XElement p, EmitContext ctx)
    {
        foreach (var (fmt, runs) in GroupInlineRuns(p))
        {
            if (fmt.HyperlinkUrl != null)
            {
                ctx.Sb.Append('[');
                foreach (var r in runs) AppendRunText(r, ctx);
                ctx.Sb.Append("](").Append(fmt.HyperlinkUrl).Append(')');
                continue;
            }

            var (open, close) = MarkdownDelimiters(fmt);
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
                Add(child, ReadRunFormatting(child, hyperlinkUrl: null));
            }
            else if (child.Name == W.hyperlink)
            {
                var url = ResolveHyperlinkUrl(child);
                // Hyperlink always forms its own group.
                Flush();
                foreach (var r in child.Elements(W.r))
                    Add(r, ReadRunFormatting(r, hyperlinkUrl: url));
                Flush();
            }
        }
        Flush();
        return groups;
    }

    private static RunFormatting ReadRunFormatting(XElement run, string? hyperlinkUrl)
    {
        var rPr = run.Element(W.rPr);
        return new RunFormatting(
            Bold: HasToggle(rPr, W.b),
            Italic: HasToggle(rPr, W.i),
            Code: IsCodeRun(rPr),
            Strike: HasToggle(rPr, W.strike),
            HyperlinkUrl: hyperlinkUrl);
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
            else if (node.Name == W.br)
                ctx.Sb.Append("  \n"); // hard line break in markdown
            else if (node.Name == W.tab)
                ctx.Sb.Append("    ");
        }
    }

    private static readonly System.Text.RegularExpressions.Regex MarkdownMetaPattern =
        new(@"([\\`*_{}\[\]()#+\-!|>~])", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string EscapeMarkdown(string s) => MarkdownMetaPattern.Replace(s, @"\$1");

    // Phase 4 placeholders — fleshed out later.
    private static void EmitListItem(XElement p, EmitContext ctx)
    {
        // Until Phase 4 lands, render list items the same as paragraphs but with a "-" prefix.
        var anchor = AnchorPrefix(p, ctx);
        ctx.Sb.Append(anchor).Append("- ");
        EmitInlineRuns(p, ctx);
        ctx.Sb.AppendLine();
    }

    private static void EmitTable(XElement tbl, EmitContext ctx)
    {
        // Phase 5 fills this in.
    }
}
