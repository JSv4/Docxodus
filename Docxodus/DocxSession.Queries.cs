// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

public sealed partial class DocxSession
{
    internal AnchorTarget? FindAnchor(string? anchorId)
    {
        if (anchorId is null) return null;
        var index = AnchorIndex();
        if (index.TryGetValue(anchorId, out var direct)) return direct;
        int lastColon = anchorId.LastIndexOf(':');
        if (lastColon <= 0 || lastColon == anchorId.Length - 1) return null;
        var unid = anchorId.Substring(lastColon + 1);
        foreach (var v in index.Values)
        {
            if (v.Unid == unid) return v;
        }
        return null;
    }

    /// <summary>
    /// Reverse-resolve an <see cref="Anchor"/> from the current projection by Unid, preferring
    /// the entry that lives in <paramref name="preferPartUri"/>.
    /// </summary>
    /// <remarks>
    /// Unids are CONTENT-ADDRESSED, so identical content in DIFFERENT package parts yields the
    /// SAME unid — a document with empty default/first/even header stories has one unid shared
    /// across several header parts (Word writes exactly that). A bare-unid reverse lookup then
    /// returns whichever part the projection happened to index first, so an <see cref="EditResult"/>
    /// would report an anchor pointing at the WRONG story, and a caller that addressed the
    /// returned anchor next (as an editor does) would silently write into the wrong part.
    /// Scoping by the part the edit actually touched keeps the round trip unambiguous; the
    /// unid-only fallback preserves behavior when the part isn't known.
    /// </remarks>
    private Anchor? AnchorForUnid(string? unid, string? preferPartUri)
    {
        if (unid is null) return null;
        AnchorTarget? fallback = null;
        foreach (var t in AnchorIndex().Values)
        {
            if (t.Unid != unid) continue;
            if (preferPartUri is not null && t.PartUri == preferPartUri) return t.Anchor;
            fallback ??= t;
        }
        return fallback?.Anchor;
    }

    /// <summary>URI of the package part that owns <paramref name="element"/>, or <c>null</c>
    /// when it belongs to no projected part (e.g. a detached element).</summary>
    private string? PartUriOf(XElement element)
    {
        var root = element.AncestorsAndSelf().Last();
        foreach (var part in EnumerateProjectedParts())
        {
            if (ReferenceEquals(part.GetXDocument().Root, root)) return part.Uri.ToString();
        }
        return null;
    }

    /// <summary>Reverse-resolve the anchor of a live element, scoped to its owning part.</summary>
    private Anchor? AnchorForElement(XElement element) =>
        AnchorForUnid((string?)element.Attribute(PtOpenXml.Unid), PartUriOf(element));

    public bool Exists(string anchorId)
    {
        ThrowIfDisposed();
        return FindAnchor(anchorId) is not null;
    }

    public AnchorInfo? GetAnchorInfo(string anchorId)
    {
        ThrowIfDisposed();
        _ = Project(); // AnchorInfo's product IS the enrichment — never serve the index-only (empty-preview) entries.
        var target = FindAnchor(anchorId);
        if (target is null) return null;
        var element = target.Resolve(_doc!);
        return new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, target.TextPreview)
        {
            AutoNumberPrefix = target.AutoNumberPrefix,
            ContentHash = element is null ? string.Empty : UnidHelper.ContentHash(element),
            VisibleText = element is null ? string.Empty : ExactVisibleText(target, element),
        };
    }

    /// <summary>
    /// Bulk variant of <see cref="GetAnchorInfo"/>. Resolves every requested anchor
    /// from the projection's cached <c>AnchorIndex</c> in a single pass. Unknown
    /// anchor ids map to <c>null</c> in the returned dictionary so callers can
    /// distinguish "anchor doesn't exist" from "anchor exists with empty preview."
    /// </summary>
    public IReadOnlyDictionary<string, AnchorInfo?> GetAnchorInfos(IEnumerable<string> anchorIds)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorIds);
        _ = Project(); // See GetAnchorInfo — enrichment required, index-only entries won't do.

        var result = new Dictionary<string, AnchorInfo?>(StringComparer.Ordinal);
        foreach (var id in anchorIds)
        {
            if (id is null) continue;
            if (result.ContainsKey(id)) continue;
            var target = FindAnchor(id);
            var element = target?.Resolve(_doc!);
            result[id] = target is null || element is null
                ? null
                : new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, target.TextPreview)
                {
                    AutoNumberPrefix = target.AutoNumberPrefix,
                    ContentHash = UnidHelper.ContentHash(element),
                    VisibleText = ExactVisibleText(target, element),
                };
        }
        return result;
    }

    private static string FlatElementText(XElement element) =>
        string.Concat(element.Descendants(W.t).Select(t => (string)t));

    private string ExactVisibleText(AnchorTarget target, XElement element)
    {
        var text = FlatElementText(element);
        var prefix = target.Anchor.Kind is "p" or "h" or "li" && target.Anchor.Scope == "body"
            ? target.AutoNumberPrefix ?? Internal.ListNumberResolver.Resolve(element, _doc!)
            : null;
        return string.IsNullOrEmpty(prefix)
            ? text
            : string.IsNullOrEmpty(text) ? prefix : prefix + " " + text;
    }

    private PreconditionTarget CurrentPreconditionTarget(string? anchorId)
    {
        if (string.IsNullOrEmpty(anchorId)) return new PreconditionTarget { Exists = false };
        var target = FindAnchor(anchorId);
        var element = target?.Resolve(_doc!);
        if (target is null || element is null)
            return new PreconditionTarget { Exists = false, AnchorId = anchorId };
        return new PreconditionTarget
        {
            Exists = true,
            AnchorId = target.Anchor.Id,
            Kind = target.Anchor.Kind,
            Scope = target.Anchor.Scope,
            ContentHash = UnidHelper.ContentHash(element),
            VisibleText = ExactVisibleText(target, element),
        };
    }

    private EditError PreconditionError(
        string condition, object? expected, object? actual, string? anchorId,
        PreconditionTarget? currentTarget = null) =>
        new(EditErrorCode.PreconditionFailed,
            $"precondition failed: {condition} expected {expected ?? "null"}, actual {actual ?? "null"}",
            anchorId)
        {
            Precondition = new PreconditionFailure(
                condition, expected, actual, _version,
                currentTarget ?? (anchorId is null ? null : CurrentPreconditionTarget(anchorId))),
        };

    /// <summary>
    /// Resolves block-level metadata (style id + name, outline level, list
    /// membership, formatting probe) for <paramref name="anchorId"/>. Returns
    /// <c>null</c> when the anchor doesn't exist. See <see cref="BlockMetadata"/>
    /// for the field reference.
    /// </summary>
    public BlockMetadata? GetBlockMetadata(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetBlockMetadata(_doc!, target);
    }

    /// <summary>Resolve a canonical <c>tbl</c> anchor to the live table's complete structural
    /// metadata. This never climbs from a nested table to an enclosing table.</summary>
    public TableMetadataResult GetTableMetadata(string tableAnchorId)
    {
        if (_disposed) return new TableMetadataResult
        {
            Error = new EditError(EditErrorCode.SessionDisposed, "session disposed", tableAnchorId),
        };
        var target = FindAnchor(tableAnchorId);
        if (target is null) return new TableMetadataResult
        {
            Error = new EditError(EditErrorCode.AnchorNotFound, "table anchor not found", tableAnchorId),
        };
        var table = target.Resolve(_doc!);
        if (target.Anchor.Kind != "tbl" || table?.Name != W.tbl) return new TableMetadataResult
        {
            Error = new EditError(EditErrorCode.AnchorWrongKind,
                "GetTableMetadata requires the table's canonical tbl anchor", tableAnchorId),
        };
        return new TableMetadataResult
        {
            Success = true,
            Metadata = Internal.TableGridModel.BuildMetadata(table, AnchorForElement),
        };
    }

    /// <summary>Resolve a canonical <c>tc</c> anchor to its table-grid coordinate and spans.</summary>
    public TableCellResolutionResult ResolveTableCellAnchor(string cellAnchorId)
    {
        if (_disposed) return CellResolutionFail(EditErrorCode.SessionDisposed, "session disposed", cellAnchorId);
        var target = FindAnchor(cellAnchorId);
        if (target is null)
            return CellResolutionFail(EditErrorCode.AnchorNotFound, "cell anchor not found", cellAnchorId);
        var cell = target.Resolve(_doc!);
        if (target.Anchor.Kind != "tc" || cell?.Name != W.tc)
            return CellResolutionFail(EditErrorCode.TableAnchorMigrationRequired,
                "ResolveTableCellAnchor requires a canonical tc anchor; obtain one from table metadata or ResolveTableCellCoordinate",
                cellAnchorId);
        var row = cell.Ancestors(W.tr).FirstOrDefault();
        var table = row?.Ancestors(W.tbl).FirstOrDefault();
        if (row is null || table is null)
            return CellResolutionFail(EditErrorCode.InternalError, "malformed table cell", cellAnchorId);
        var metadata = Internal.TableGridModel.BuildMetadata(table, AnchorForElement);
        var resolved = metadata.Rows.SelectMany(item => item.Cells)
            .FirstOrDefault(item => item.Anchor.Id == target.Anchor.Id);
        return resolved is null
            ? CellResolutionFail(EditErrorCode.InternalError, "cell is absent from its table grid", cellAnchorId)
            : new TableCellResolutionResult { Success = true, Cell = resolved };
    }

    /// <summary>Resolve a zero-based table-grid coordinate to the physical cell covering it.
    /// Coordinates omitted by <c>w:gridBefore</c>/<c>w:gridAfter</c> resolve as not found;
    /// every coordinate covered by a <c>w:gridSpan</c> resolves to the same cell anchor.</summary>
    public TableCellResolutionResult ResolveTableCellCoordinate(
        string tableAnchorId, int rowIndex, int columnIndex)
    {
        var tableResult = GetTableMetadata(tableAnchorId);
        if (!tableResult.Success)
            return new TableCellResolutionResult { Error = tableResult.Error };
        var cell = Internal.TableGridModel.CellAt(tableResult.Metadata!, rowIndex, columnIndex);
        return cell is null
            ? CellResolutionFail(EditErrorCode.AnchorNotFound,
                $"no table cell covers coordinate ({rowIndex}, {columnIndex})", tableAnchorId)
            : new TableCellResolutionResult { Success = true, Cell = cell };
    }

    private static TableCellResolutionResult CellResolutionFail(
        EditErrorCode code, string message, string? anchorId) =>
        new() { Error = new EditError(code, message, anchorId) };

    /// <summary>
    /// Bulk variant of <see cref="GetBlockMetadata"/>. Unknown anchor ids map
    /// to <c>null</c>; duplicate ids are deduped; iteration order matches
    /// input order for keys that appear first.
    /// </summary>
    public IReadOnlyDictionary<string, BlockMetadata?> GetBlockMetadatas(IEnumerable<string> anchorIds)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorIds);

        var result = new Dictionary<string, BlockMetadata?>(StringComparer.Ordinal);
        foreach (var id in anchorIds)
        {
            if (id is null) continue;
            if (result.ContainsKey(id)) continue;
            var target = FindAnchor(id);
            result[id] = target is null ? null : Internal.BlockMetadataOps.GetBlockMetadata(_doc!, target);
        }
        return result;
    }

    /// <summary>
    /// Resolves the <see cref="ListMembership"/> for a list-item paragraph;
    /// returns <c>null</c> when the anchor has no <c>w:numPr</c> (inline or
    /// inherited from style) or doesn't exist.
    /// </summary>
    public ListMembership? GetListMembership(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetListMembership(_doc!, target);
    }

    /// <summary>
    /// Resolves the <see cref="SectionInfo"/> for the <c>w:sectPr</c> that
    /// governs <paramref name="anchorId"/>. Returns <c>null</c> when the
    /// anchor lives outside the body part (footnotes, endnotes, headers,
    /// footers, comments) or doesn't exist.
    /// </summary>
    public SectionInfo? GetSectionInfo(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetSectionInfo(_doc!, target);
    }

    /// <summary>
    /// Enumerates the document's explicit paragraph, character, table, and numbering style
    /// definitions in declaration order. Each entry includes inheritance/gallery metadata and
    /// high-signal effective properties; returned style ids are the ids accepted by the matching
    /// paragraph/run style mutation fields.
    /// </summary>
    public IReadOnlyList<StyleInfo> ListStyles()
    {
        ThrowIfDisposed();
        return Internal.FormattingIntrospectionOps.ListStyles(_doc!);
    }

    /// <summary>
    /// Inspect direct and effective paragraph formatting plus every text-bearing run's direct and
    /// effective formatting for a paragraph/heading/list-item anchor. Returns <c>null</c> for an
    /// unknown or non-paragraph anchor.
    /// </summary>
    public FormattingInspection? GetFormatting(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null
            ? null
            : Internal.FormattingIntrospectionOps.GetFormatting(_doc!, target);
    }

    /// <summary>
    /// Enumerate text-bearing inline runs for a paragraph/heading/list-item anchor. Every result
    /// carries a block-relative <see cref="CharSpan"/> that can be passed directly to
    /// <see cref="ApplyFormat(string, CharSpan?, FormatOp)"/>. Unknown/non-paragraph anchors return
    /// an empty list.
    /// </summary>
    public IReadOnlyList<InlineSpan> ListInlineSpans(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null
            ? Array.Empty<InlineSpan>()
            : Internal.FormattingIntrospectionOps.ListInlineSpans(_doc!, target);
    }

    /// <summary>
    /// Searches the flat text of every paragraph/heading/list-item in <paramref name="scope"/>
    /// for matches of <paramref name="pattern"/> and returns them in document order, each
    /// with the run fragments it spans. The fragment list lets callers rewrite a match in
    /// place while preserving each fragment's formatting — see #143 for design context.
    /// </summary>
    /// <param name="pattern">Regular-expression pattern (use <c>Regex.Escape</c> for literal text).</param>
    /// <param name="options">Standard <see cref="System.Text.RegularExpressions.RegexOptions"/> flags.</param>
    /// <param name="scope">Which package parts to search. Defaults to <see cref="ProjectionScopes.Body"/>.</param>
    /// <param name="contextChars">Number of characters of surrounding text to include in
    /// <see cref="TextMatch.ContextBefore"/> and <see cref="TextMatch.ContextAfter"/>.</param>
    public IReadOnlyList<TextMatch> Grep(
        string pattern,
        System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 80,
        WhitespaceMode whitespace = WhitespaceMode.Preserve,
        ContextBoundary boundary = ContextBoundary.Char,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(pattern)) return Array.Empty<TextMatch>();

        var regex = new System.Text.RegularExpressions.Regex(pattern, options);
        var results = new List<TextMatch>();

        // Walk the projection's AnchorIndex so document order is the same order
        // an agent sees in the projection. Only block-level kinds that hold runs
        // qualify (paragraphs/headings/list-items/table cells); other kinds either
        // don't contain text directly (tbl, tr, sec) or live in non-body scopes
        // we filter explicitly below.
        var index = Project().AnchorIndex;
        foreach (var target in index.Values)
        {
            if (!ScopeMatches(target.Anchor.Scope, scope)) continue;
            if (target.Anchor.Kind is not ("p" or "h" or "li" or "tc")) continue;

            var element = target.Resolve(_doc!);
            if (element is null) continue;

            // Table cells contain paragraphs; recurse so a Grep over the body
            // also hits cell text. Other kinds operate on the element directly.
            if (target.Anchor.Kind == "tc")
            {
                // Cell paragraphs are reachable via their own AnchorIndex entries,
                // so skip the cell wrapper to avoid double-counting matches.
                continue;
            }

            var map = Internal.RunTextMap.Build(element);
            if (map.FlatText.Length == 0) continue;

            // Look up the owner part once per anchor so the hyperlink resolver
            // doesn't have to walk back up to the root annotation per run.
            var ownerPart = ResolvePart(target.PartUri);

            // For Normalize mode: match against a whitespace-normalized COPY of the
            // flat text while keeping the segment offset map pointing at the original
            // positions. Match indices apply unchanged because the substitutions are
            // 1:1 (NBSP → space, narrow-NBSP → space, thin-space → space) — same
            // character count, just different code points.
            var matchText = whitespace == WhitespaceMode.Normalize
                ? NormalizeWhitespace(map.FlatText)
                : map.FlatText;

            foreach (System.Text.RegularExpressions.Match m in regex.Matches(matchText))
            {
                if (!m.Success || m.Length == 0) continue;

                var pieces = Internal.RunTextMap.ResolveRange(map, m.Index, m.Length);
                if (pieces.Count == 0) continue;

                var fragments = new List<RunFragment>(pieces.Count);
                foreach (var (seg, offsetInRun, len) in pieces)
                {
                    var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
                    var runText = RunText(seg.Run);
                    fragments.Add(new RunFragment
                    {
                        Unid = runUnid,
                        Text = runText.Substring(offsetInRun, len),
                        SpanInElement = new CharSpan(offsetInRun, len),
                        Formatting = ExtractFormatting(seg.Run, ownerPart),
                    });
                }

                var (ctxBefore, ctxAfter) = WalkContext(map.FlatText, m.Index, m.Length, contextChars, boundary);

                var groups = new string[m.Groups.Count];
                for (int i = 0; i < m.Groups.Count; i++) groups[i] = m.Groups[i].Value;

                results.Add(new TextMatch
                {
                    Text = m.Value,
                    EnclosingAnchor = target,
                    Span = new CharSpan(m.Index, m.Length),
                    Fragments = fragments,
                    ContextBefore = ctxBefore,
                    ContextAfter = ctxAfter,
                    Groups = groups,
                    Citation = citationRequest is null
                        ? null
                        : GetPageCitation(target.Anchor.Id, citationRequest),
                });
            }
        }

        return results;
    }

    /// <summary>
    /// Searches the flat text of every block-level element in <paramref name="scope"/>, like
    /// <see cref="Grep"/>, but lets a single match span <em>adjacent</em> block-level siblings
    /// (paragraphs/headings/list items) sharing the same direct parent. Returns matches in
    /// document order, each with a per-block <see cref="BlockSlice"/> breakdown. See issue #146.
    ///
    /// Block boundaries are represented in the concatenated text by a single <c>\n</c>, so
    /// <c>^</c>/<c>$</c> with <see cref="System.Text.RegularExpressions.RegexOptions.Multiline"/>
    /// anchor at boundaries; <c>.</c> won't cross unless
    /// <see cref="System.Text.RegularExpressions.RegexOptions.Singleline"/> is set.
    ///
    /// Matches never cross:
    /// <list type="bullet">
    ///   <item><description>OOXML package parts (e.g. body → footnote, header → body).</description></item>
    ///   <item><description>Container boundaries (e.g. body paragraph → table-cell paragraph).</description></item>
    ///   <item><description>Non-paragraph siblings (a <c>w:tbl</c> or section property between two paragraphs breaks the run).</description></item>
    /// </list>
    ///
    /// Superset of <see cref="Grep"/>: single-block matches are still returned (with one
    /// <see cref="BlockSlice"/>). Callers that want only cross-block hits can filter
    /// <c>Slices.Count &gt; 1</c>.
    /// </summary>
    public IReadOnlyList<CrossBlockMatch> GrepCrossBlock(
        string pattern,
        System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 80,
        WhitespaceMode whitespace = WhitespaceMode.Preserve,
        ContextBoundary boundary = ContextBoundary.Char,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(pattern)) return Array.Empty<CrossBlockMatch>();

        var regex = new System.Text.RegularExpressions.Regex(pattern, options);
        var results = new List<CrossBlockMatch>();

        // Build groups of consecutive block-level siblings under the same parent.
        // Document order comes from AnchorIndex iteration; the parent check ensures
        // we don't bridge a body paragraph to a table-cell paragraph or a header to a
        // body paragraph. Any non-eligible anchor (kind != p/h/li, or out of scope,
        // or unresolved) breaks the run.
        var index = Project().AnchorIndex;
        var groups = new List<List<(AnchorTarget Target, XElement Element)>>();
        List<(AnchorTarget, XElement)>? current = null;
        XElement? currentParent = null;

        foreach (var target in index.Values)
        {
            if (!ScopeMatches(target.Anchor.Scope, scope)) { current = null; continue; }
            if (target.Anchor.Kind is not ("p" or "h" or "li")) { current = null; continue; }

            var element = target.Resolve(_doc!);
            if (element is null) { current = null; continue; }

            if (current is not null && ReferenceEquals(element.Parent, currentParent))
            {
                current.Add((target, element));
            }
            else
            {
                current = new List<(AnchorTarget, XElement)> { (target, element) };
                currentParent = element.Parent;
                groups.Add(current);
            }
        }

        foreach (var group in groups)
        {
            // Build per-block maps + a parallel boundary array (start offset of each
            // block in the concatenated text, length of the block's flat text). A
            // single '\n' between blocks acts as the sentinel.
            var maps = new List<Internal.RunTextMap.Map>(group.Count);
            var starts = new int[group.Count];
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < group.Count; i++)
            {
                if (i > 0) sb.Append('\n');
                starts[i] = sb.Length;
                var map = Internal.RunTextMap.Build(group[i].Element);
                maps.Add(map);
                sb.Append(map.FlatText);
            }
            var concat = sb.ToString();
            if (concat.Length == 0) continue;

            var matchText = whitespace == WhitespaceMode.Normalize
                ? NormalizeWhitespace(concat)
                : concat;

            // Cache owner-part lookup per group; every block in a group lives in the
            // same package part (siblings share a parent), so one lookup suffices.
            var ownerPart = ResolvePart(group[0].Target.PartUri);

            foreach (System.Text.RegularExpressions.Match m in regex.Matches(matchText))
            {
                if (!m.Success || m.Length == 0) continue;

                var slices = new List<BlockSlice>();
                var anchors = new List<AnchorTarget>();
                for (int i = 0; i < group.Count; i++)
                {
                    var blockStart = starts[i];
                    var blockEnd = blockStart + maps[i].FlatText.Length;
                    if (blockEnd <= m.Index) continue;
                    if (blockStart >= m.Index + m.Length) break;

                    var overlapStart = Math.Max(m.Index, blockStart) - blockStart;
                    var overlapLen = Math.Min(m.Index + m.Length, blockEnd) - blockStart - overlapStart;

                    var pieces = overlapLen > 0
                        ? Internal.RunTextMap.ResolveRange(maps[i], overlapStart, overlapLen)
                        : new List<(Internal.RunTextMap.RunSegment, int, int)>();

                    var fragments = new List<RunFragment>(pieces.Count);
                    foreach (var (seg, offsetInRun, len) in pieces)
                    {
                        var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
                        var runText = RunText(seg.Run);
                        fragments.Add(new RunFragment
                        {
                            Unid = runUnid,
                            Text = runText.Substring(offsetInRun, len),
                            SpanInElement = new CharSpan(offsetInRun, len),
                            Formatting = ExtractFormatting(seg.Run, ownerPart),
                        });
                    }

                    slices.Add(new BlockSlice
                    {
                        Anchor = group[i].Target,
                        SpanInBlock = new CharSpan(overlapStart, overlapLen),
                        Fragments = fragments,
                    });
                    anchors.Add(group[i].Target);
                }

                if (slices.Count == 0) continue;

                var (ctxBefore, ctxAfter) = WalkContext(concat, m.Index, m.Length, contextChars, boundary);

                var groups2 = new string[m.Groups.Count];
                for (int i = 0; i < m.Groups.Count; i++) groups2[i] = m.Groups[i].Value;

                results.Add(new CrossBlockMatch
                {
                    Text = m.Value,
                    EnclosingAnchors = anchors,
                    Slices = slices,
                    ContextBefore = ctxBefore,
                    ContextAfter = ctxAfter,
                    Groups = groups2,
                    Citations = citationRequest is null
                        ? null
                        : anchors.Select(a => GetPageCitation(a.Anchor.Id, citationRequest)).ToArray(),
                });
            }
        }

        return results;
    }

    /// <summary>
    /// Finds the first anchor whose flat text contains <paramref name="needle"/>, or null.
    /// Thin wrapper over <see cref="Grep"/> — every consumer was reimplementing the same
    /// scan with its own quirks (case sensitivity, NBSP, scope filter). See issue #137.
    /// </summary>
    public AnchorTarget? FindByText(string needle, FindOptions? options = null) =>
        FindAllByText(needle, options).FirstOrDefault();

    /// <summary>
    /// All anchors whose flat text contains <paramref name="needle"/>, in document order.
    /// Duplicates removed (one entry per enclosing anchor regardless of how many times
    /// the needle appears inside it).
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindAllByText(string needle, FindOptions? options = null)
    {
        if (string.IsNullOrEmpty(needle)) return Array.Empty<AnchorTarget>();
        var opts = options ?? new FindOptions();
        var regexOpts = opts.IgnoreCase
            ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
            : System.Text.RegularExpressions.RegexOptions.None;
        return FindMatchesFiltered(System.Text.RegularExpressions.Regex.Escape(needle), regexOpts, opts);
    }

    /// <summary>
    /// All anchors with at least one match for <paramref name="pattern"/>, in document order.
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindByRegex(
        string pattern,
        System.Text.RegularExpressions.RegexOptions regexOptions = System.Text.RegularExpressions.RegexOptions.None,
        FindOptions? options = null) =>
        FindMatchesFiltered(pattern, regexOptions, options ?? new FindOptions());

    /// <summary>
    /// All anchors of a given kind (and optionally scope), in document order. Direct read
    /// over the projection's <c>AnchorIndex</c>; no text scan, so no <see cref="FindOptions"/>.
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindByKind(
        string kind,
        string? scope = null,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        var result = new List<AnchorTarget>();
        foreach (var target in Project().AnchorIndex.Values)
        {
            if (target.Anchor.Kind != kind) continue;
            if (scope is not null && target.Anchor.Scope != scope) continue;
            result.Add(AttachCitation(target, citationRequest));
        }
        return result;
    }

    private IReadOnlyList<AnchorTarget> FindMatchesFiltered(
        string pattern,
        System.Text.RegularExpressions.RegexOptions regexOptions,
        FindOptions options)
    {
        ThrowIfDisposed();
        // Prefer Scopes (typed, composable) for the underlying Grep walker. The
        // string ScopeFilter still applies as a finer post-filter below for
        // callers targeting a single named part like "hdr1".
        var matches = Grep(
            pattern,
            regexOptions,
            options.Scopes,
            contextChars: 0,
            whitespace: options.IgnoreWhitespace ? WhitespaceMode.Normalize : WhitespaceMode.Preserve,
            citationRequest: options.CitationRequest);

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<AnchorTarget>();
        foreach (var m in matches)
        {
            var anchor = m.EnclosingAnchor;
            if (options.KindFilter is not null && anchor.Anchor.Kind != options.KindFilter) continue;
            if (options.ScopeFilter is not null && anchor.Anchor.Scope != options.ScopeFilter) continue;
            if (!seen.Add(anchor.Anchor.Id)) continue;
            result.Add(AttachCitation(anchor, options.CitationRequest));
        }
        return result;
    }

    /// <summary>
    /// Enumerate every anchor whose scope belongs to <paramref name="scopes"/>, in
    /// projection order. Convenience over walking <c>Project().AnchorIndex</c> and
    /// filtering by scope name — common for callers that want to operate on every
    /// header paragraph, every footnote, etc.
    /// </summary>
    /// <example>
    /// <code>
    /// // Every paragraph in any header or footer:
    /// foreach (var t in session.AnchorsByScope(ProjectionScopes.Headers | ProjectionScopes.Footers))
    ///     Console.WriteLine($"{t.Anchor.Scope}: {t.TextPreview}");
    /// </code>
    /// </example>
    public IReadOnlyList<AnchorTarget> AnchorsByScope(ProjectionScopes scopes)
    {
        ThrowIfDisposed();
        var result = new List<AnchorTarget>();
        foreach (var t in Project().AnchorIndex.Values)
            if (scopes.IncludesScope(t.Anchor.Scope))
                result.Add(t);
        return result;
    }

    // ─── Annotation-based anchor discovery (#132) ────────────────────────

    /// <summary>
    /// Resolves an annotation's range to the block-level markdown anchors covering it,
    /// in document order. The bridge between the read-side annotation API
    /// (<see cref="AnnotationManager"/>) and the write-side session: an agent that wants
    /// to edit "the indemnification clause" looks the annotation up by id and gets the
    /// anchors it can hand to <see cref="ReplaceText"/> / <see cref="DeleteBlock"/> /
    /// <see cref="Raw"/>. Returns an empty list when the id is unknown or the annotation's
    /// bookmark is missing/malformed.
    /// </summary>
    /// <remarks>
    /// v1 returns the enclosing block anchors — every paragraph/heading/list-item/cell/
    /// row/table whose subtree overlaps the bookmark range. Bookmarks that sit inside a
    /// single paragraph yield that paragraph's anchor; bookmarks spanning multiple blocks
    /// yield each in document order. A finer-grained <see cref="CharSpan"/>-aware return
    /// is left to a follow-up (see the issue's "Out of scope for v1").
    /// </remarks>
    public IReadOnlyList<AnchorTarget> FindByAnnotation(
        string annotationId,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(annotationId)) return Array.Empty<AnchorTarget>();
        var ann = AnnotationManager.GetAnnotations(_doc!)
            .FirstOrDefault(a => string.Equals(a.Id, annotationId, StringComparison.Ordinal));
        if (ann is null || string.IsNullOrEmpty(ann.BookmarkName))
            return Array.Empty<AnchorTarget>();
        return ResolveBookmarkAnchors(ann.BookmarkName)
            .Select(target => AttachCitation(target, citationRequest)).ToArray();
    }

    /// <summary>
    /// Finds every annotation whose <see cref="DocumentAnnotation.LabelId"/> equals
    /// <paramref name="labelId"/> and resolves each of their ranges. The result is keyed
    /// by annotation id so callers can disambiguate when the same label was applied to
    /// multiple regions (e.g. three separate "WARRANTY" annotations). Annotations whose
    /// bookmark is missing or resolves to no anchors are omitted from the result.
    /// </summary>
    public IReadOnlyDictionary<string, IReadOnlyList<AnchorTarget>> FindByLabel(
        string labelId,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        var map = new Dictionary<string, IReadOnlyList<AnchorTarget>>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(labelId)) return map;
        foreach (var ann in AnnotationManager.GetAnnotations(_doc!))
        {
            if (!string.Equals(ann.LabelId, labelId, StringComparison.Ordinal)) continue;
            if (string.IsNullOrEmpty(ann.BookmarkName)) continue;
            if (ann.Id is null) continue;
            var anchors = ResolveBookmarkAnchors(ann.BookmarkName)
                .Select(target => AttachCitation(target, citationRequest)).ToArray();
            if (anchors.Length > 0) map[ann.Id] = anchors;
        }
        return map;
    }

    /// <summary>
    /// Resolves any bookmark in body/header/footer/footnote/endnote parts (Docxodus-managed or user-authored)
    /// to the block-level anchors covering its range, in document order. Empty when the
    /// bookmark name is unknown or its end marker is missing. Use this for raw bookmark
    /// names that didn't come from <see cref="AnnotationManager"/>.
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindByBookmark(
        string bookmarkName,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(bookmarkName)) return Array.Empty<AnchorTarget>();
        return ResolveBookmarkAnchors(bookmarkName)
            .Select(target => AttachCitation(target, citationRequest)).ToArray();
    }

    private AnchorTarget AttachCitation(AnchorTarget target, PageCitationRequest? request) =>
        request is null
            ? target
            : new AnchorTarget
            {
                Anchor = target.Anchor,
                PartUri = target.PartUri,
                Unid = target.Unid,
                TextPreview = target.TextPreview,
                AutoNumberPrefix = target.AutoNumberPrefix,
                Citation = GetPageCitation(target.Anchor.Id, request),
            };

    /// <summary>
    /// Enumerates every annotation persisted in the document — id, label id/text, color,
    /// author, and (when the bookmark resolves) the annotated text it covers. Lets an
    /// agent prime itself with "here are the labeled regions you can target" before
    /// committing to a specific id.
    /// </summary>
    public IReadOnlyList<DocumentAnnotation> ListAnnotations()
    {
        ThrowIfDisposed();
        return AnnotationManager.GetAnnotations(_doc!);
    }

    /// <summary>
    /// Walks the bookmark's owning story part once: locates the bookmark by name, then collects
    /// every block-level anchor whose subtree overlaps the bookmark range, deduplicated
    /// and sorted in document order. Pre-order positions are recomputed per call rather
    /// than cached — callers in agentic loops should resolve once and reuse the result.
    /// </summary>
    private IReadOnlyList<AnchorTarget> ResolveBookmarkAnchors(string bookmarkName)
    {
        var matches = Internal.OwnedPartRelationships.StoryParts(_doc!)
            .SelectMany(o => o.Part.GetXDocument().Descendants(W.bookmarkStart).Select(start => (Owner: o, Start: start)))
            .Where(x => (string?)x.Start.Attribute(W.name) == bookmarkName).ToList();
        if (matches.Count != 1) return Array.Empty<AnchorTarget>();
        var owner = matches[0].Owner;
        var start = matches[0].Start;
        var root = owner.Part.GetXDocument().Root!;
        var bookmarkId = (string?)start.Attribute(W.id);
        if (bookmarkId is null) return Array.Empty<AnchorTarget>();
        var end = root.Descendants(W.bookmarkEnd)
            .FirstOrDefault(b => (string?)b.Attribute(W.id) == bookmarkId);
        if (end is null) return Array.Empty<AnchorTarget>();

        // Force Project() so Unids are assigned on every block and the AnchorIndex is
        // populated. Building a Unid → AnchorTarget reverse map lets us look up each
        // candidate block without re-running the converter's KindFor classifier here.
        var index = Project().AnchorIndex;
        var byUnid = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
        foreach (var t in index.Values)
            if (t.PartUri == owner.PartUri) byUnid[t.Unid] = t;

        // Pre-order positions support two operations: (a) deciding whether a block's
        // subtree overlaps the bookmark range, (b) sorting the collected hits back into
        // document order. O(N) per call — fine for in-session use where Project() is
        // already O(N).
        var pos = new Dictionary<XElement, int>(ReferenceEqualityComparer.Instance);
        int counter = 0;
        foreach (var el in root.DescendantsAndSelf()) pos[el] = counter++;

        if (!pos.TryGetValue(start, out var startPos) || !pos.TryGetValue(end, out var endPos))
            return Array.Empty<AnchorTarget>();
        if (endPos <= startPos) return Array.Empty<AnchorTarget>();

        var hits = new List<(int Pos, AnchorTarget Target)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var el in root.Descendants())
        {
            var unid = (string?)el.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            if (!byUnid.TryGetValue(unid, out var target)) continue;
            if (!string.Equals(target.PartUri, owner.PartUri, StringComparison.Ordinal)) continue;

            var elStart = pos[el];
            var lastDesc = el.DescendantsAndSelf().Last();
            var elEnd = pos[lastDesc];
            // Strict overlap on the marker positions themselves: a bookmark sitting
            // exactly between two paragraphs shouldn't pick up either of them.
            if (elEnd <= startPos) continue;
            if (elStart >= endPos) continue;
            if (!seen.Add(target.Anchor.Id)) continue;
            hits.Add((elStart, target));
        }

        hits.Sort((a, b) => a.Pos.CompareTo(b.Pos));
        var result = new AnchorTarget[hits.Count];
        for (int i = 0; i < hits.Count; i++) result[i] = hits[i].Target;
        return result;
    }
}
