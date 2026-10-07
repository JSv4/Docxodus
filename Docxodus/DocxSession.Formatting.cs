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
    // ─── Tier C: formatting ──────────────────────────────────────────────

    /// <summary>
    /// Convenience: find <paramref name="substring"/> in the anchor's flat text and apply
    /// <paramref name="op"/> to the first occurrence. Eliminates the offset-arithmetic
    /// trap where an auto-number prefix shifts the visible text vs the run-text indices
    /// the underlying <see cref="ApplyFormat(string, CharSpan?, FormatOp)"/> overload
    /// expects — see issue #138. Named distinctly (rather than overloading) so existing
    /// <c>ApplyFormat(anchor, null, op)</c> calls (whole-paragraph format) stay
    /// unambiguous to the C# overload resolver.
    /// </summary>
    public EditResult ApplyFormatToSubstring(string anchorId, string substring, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (string.IsNullOrEmpty(substring))
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "substring must be non-empty", anchorId);

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ApplyFormat requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var map = Internal.RunTextMap.Build(element);
        var idx = map.FlatText.IndexOf(substring, StringComparison.Ordinal);
        if (idx < 0) return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
            $"substring not found in anchor's text", anchorId);

        return ApplyFormat(anchorId, new CharSpan(idx, substring.Length), op);
    }

    /// <summary>
    /// Convenience: apply <paramref name="op"/> to the exact span covered by a
    /// <see cref="TextMatch"/> (typically from <see cref="Grep"/>). The match's
    /// <see cref="TextMatch.EnclosingAnchor"/> + <see cref="TextMatch.Span"/> address
    /// one specific occurrence even when several identical needles share the same block.
    /// </summary>
    public EditResult ApplyFormat(TextMatch match, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");
        return ApplyFormat(
            match.EnclosingAnchor.Anchor.Id,
            new CharSpan(match.Span.Start, match.Span.Length),
            op);
    }

    public EditResult ApplyFormat(string anchorId, CharSpan? span, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (op is null) return EditResult.Fail(EditErrorCode.MalformedMarkdown, "null format op", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ApplyFormat requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var totalText = ParagraphText(element);
        var actualSpan = span ?? new CharSpan(0, totalText.Length);
        if (actualSpan.Start < 0 || actualSpan.Length < 0 ||
            actualSpan.Start + actualSpan.Length > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span [{actualSpan.Start},{actualSpan.Start + actualSpan.Length}) out of [0,{totalText.Length})", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Inline code references a "Code" character style by id; ensure it actually
            // exists so the run renders monospace instead of pointing at a phantom style.
            if (op.Code is true) Internal.StyleFactory.EnsureCodeCharacterStyle(_doc!);

            SplitRunsAtOffset(element, actualSpan.Start);
            SplitRunsAtOffset(element, actualSpan.Start + actualSpan.Length);

            var trackFormatChanges = _trackedChanges == TrackedChangeMode.RenderInline;
            var revisionAuthor = _revisionAuthor ?? "docxodus";
            // Every run touched by one ApplyFormat call belongs to the same user action.
            // Give its per-run rPrChange markers one timestamp for coherent attribution
            // and grouping. The value is monotonic at tick precision so two adjacent,
            // same-author ApplyFormat calls remain independently selectable even when
            // the wall clock would otherwise stamp them in the same second.
            var revisionDate = trackFormatChanges
                ? NextTrackedFormatRevisionDate()
                : null;

            int consumed = 0;
            foreach (var run in InlineRuns(element).ToList())
            {
                var runText = RunText(run);
                int runStart = consumed;
                int runEnd = consumed + runText.Length;
                consumed = runEnd;
                if (runEnd <= actualSpan.Start || runStart >= actualSpan.Start + actualSpan.Length) continue;
                if (trackFormatChanges)
                    ApplyFormatToRunTracked(run, op, revisionAuthor, revisionDate!);
                else
                    ApplyFormatToRun(run, op);
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    private string NextTrackedFormatRevisionDate()
    {
        while (true)
        {
            var observed = System.Threading.Interlocked.Read(ref _lastFormatRevisionTicks);
            var now = DateTime.UtcNow.Ticks;
            var next = now > observed ? now : observed + 1;
            if (System.Threading.Interlocked.CompareExchange(
                    ref _lastFormatRevisionTicks, next, observed) == observed)
            {
                return new DateTime(next, DateTimeKind.Utc).ToString(
                    "yyyy-MM-ddTHH:mm:ss.fffffffZ",
                    System.Globalization.CultureInfo.InvariantCulture);
            }
        }
    }

    public EditResult SetParagraphStyle(string anchorId, string styleId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetParagraphStyle requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        if (_trackedChanges == TrackedChangeMode.RenderInline
            && !Internal.StyleFactory.HasParagraphStyle(_doc!, styleId))
            return TrackedStructureUnsupported(
                $"SetParagraphStyle requiring synthesis of style '{styleId}'", anchorId);

        // Style synthesis mutates the styles part. Capture before it so rejecting the generated
        // pPrChange (or undoing this direct-mode op) restores package state as well as paragraph XML.
        // Tracked mode preflights above because native pPrChange has no styles-part before-image.
        var preOp = TakeSnapshot();

        // Find-or-create well-known built-in styles (Title, Subtitle, Heading1-9) the document
        // hasn't defined yet, so applying one works instead of silently failing. Mirrors the inline
        // "Code" character style. A truly unknown custom id is left untouched and still rejected.
        if (!Internal.StyleFactory.EnsureParagraphStyle(_doc!, styleId))
            return EditResult.Fail(EditErrorCode.UnknownStyle, $"style id not found: {styleId}", anchorId);

        var oldPPr = new XElement(element.Element(W.pPr) ?? new XElement(W.pPr));
        _history.RecordPreOp(preOp);
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            pPr.Element(W.pStyle)?.Remove();
            pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));
            if (_trackedChanges == TrackedChangeMode.RenderInline)
                TrackPropertyMutation(pPr, oldPPr, W.pPrChange,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate(), W.rPr, W.sectPr);

            InvalidateProjectionCache();
            // Anchor kind may have flipped (e.g., p → h); look it up in the fresh index.
            var freshIndex = AnchorIndex();
            var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>Return the existing w:pPr child of <paramref name="name"/> (attributes intact),
    /// or slot a new empty one in at its correct CT_PPr position.</summary>
    private static XElement GetOrCreatePPrChild(XElement pPr, XName name)
    {
        var child = pPr.Element(name);
        if (child is null)
        {
            child = new XElement(name);
            SetPPrChildInOrder(pPr, child);
        }
        return child;
    }

    /// <summary>Insert (replacing any existing) a w:pPr child at its correct CT_PPr position.</summary>
    private static void SetPPrChildInOrder(XElement pPr, XElement child)
    {
        pPr.Elements(child.Name).Remove();
        int idx = Array.IndexOf(PPrChildOrder, child.Name.LocalName);
        XElement? after = null;
        foreach (var e in pPr.Elements())
        {
            int ei = Array.IndexOf(PPrChildOrder, e.Name.LocalName);
            if (ei >= 0 && ei < idx) after = e;
            else if (ei >= idx) break;
        }
        if (after is null) pPr.AddFirst(child);
        else after.AddAfterSelf(child);
    }

    private static XElement BorderEdgeElement(XName edgeName, ParagraphBorderEdge edge) =>
        new XElement(edgeName,
            new XAttribute(W.val, string.IsNullOrEmpty(edge.Style) ? "single" : edge.Style),
            new XAttribute(W.sz, edge.Size ?? 6),
            new XAttribute(W.space, edge.Space ?? 1),
            new XAttribute(W.color, string.IsNullOrEmpty(edge.Color) ? "auto" : edge.Color));

    /// <summary>Insert/replace a single <c>w:pBdr</c> edge, keeping CT_PBdr child order.</summary>
    private static void SetBorderEdgeInOrder(XElement pBdr, XName edgeName, XElement edge)
    {
        pBdr.Elements(edgeName).Remove();
        int idx = Array.IndexOf(PBdrEdgeOrder, edgeName.LocalName);
        XElement? after = null;
        foreach (var e in pBdr.Elements())
        {
            int ei = Array.IndexOf(PBdrEdgeOrder, e.Name.LocalName);
            if (ei >= 0 && ei < idx) after = e;
            else if (ei >= idx) break;
        }
        if (after is null) pBdr.AddFirst(edge);
        else after.AddAfterSelf(edge);
    }

    /// <summary>Apply top/bottom border edges (and an optional clear) to a paragraph's pPr, in place.</summary>
    private static void ApplyParagraphBorders(XElement pPr, ParagraphBorderEdge? top, ParagraphBorderEdge? bottom, bool clear)
    {
        if (clear) pPr.Element(W.pBdr)?.Remove();
        if (top is null && bottom is null) return;
        var pBdr = pPr.Element(W.pBdr);
        bool isNew = pBdr is null;
        pBdr ??= new XElement(W.pBdr);
        if (top is not null) SetBorderEdgeInOrder(pBdr, W.top, BorderEdgeElement(W.top, top));
        if (bottom is not null) SetBorderEdgeInOrder(pBdr, W.bottom, BorderEdgeElement(W.bottom, bottom));
        if (isNew) SetPPrChildInOrder(pPr, pBdr);
    }

    /// <summary>
    /// Set paragraph-level formatting (alignment, indent delta, first-line/hanging indent,
    /// before/after/line spacing, page-break-before, borders) on the paragraph the anchor
    /// names. Only the non-null fields of <paramref name="op"/> change.
    /// </summary>
    public EditResult SetParagraphFormat(string anchorId, ParagraphFormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetParagraphFormat requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        if (op.FirstLineIndent is not null && op.HangingIndent is not null)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "firstLineIndent and hangingIndent are mutually exclusive (w:ind holds one or the other)", anchorId);
        if (op.FirstLineIndent is < 0 || op.HangingIndent is < 0 ||
            op.SpacingBefore is < 0 || op.SpacingAfter is < 0 || op.LineSpacing is < 0)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "indent/spacing values are unsigned twips and must be >= 0", anchorId);
        if (op.LineSpacingRule is not null && op.LineSpacing is null)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "lineSpacingRule requires lineSpacing (w:lineRule qualifies w:line)", anchorId);

        var oldPPr = new XElement(element.Element(W.pPr) ?? new XElement(W.pPr));
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }

            if (op.Alignment is { } align)
            {
                var val = align switch
                {
                    ParagraphAlignment.Left => "left",
                    ParagraphAlignment.Center => "center",
                    ParagraphAlignment.Right => "right",
                    ParagraphAlignment.Justify => "both",
                    _ => "left",
                };
                SetPPrChildInOrder(pPr, new XElement(W.jc, new XAttribute(W.val, val)));
            }

            if (op.PageBreakBefore is { } pbb)
            {
                pPr.Element(W.pageBreakBefore)?.Remove();
                if (pbb) SetPPrChildInOrder(pPr, new XElement(W.pageBreakBefore));
            }

            if (op.IndentDelta is { } delta && delta != 0)
            {
                var ind = pPr.Element(W.ind);
                // Parse the current left indent tolerantly: documents exported by Google Docs (and
                // others) emit non-integer twips like w:left="12.996749877929688", which a bare
                // (int?) cast rejects with a FormatException. AttributeToTwips is the same helper the
                // HTML converter uses (decimal → truncate), so we read what the doc renders and write
                // back a clean integer.
                // Adjust the edge in whichever spelling the paragraph already uses (w:start or
                // w:left), so the value written is the one every reader sees.
                var leading = WordprocessingMLUtil.IndLeadingAttribute(ind);
                int cur = WordprocessingMLUtil.AttributeToTwips(leading) ?? 0;
                int next = Math.Max(0, cur + delta);
                if (ind is null)
                {
                    ind = new XElement(W.ind);
                    SetPPrChildInOrder(pPr, ind);
                }
                ind.SetAttributeValue(leading?.Name ?? W.left, next);
            }

            // firstLine/hanging share one w:ind slot in Word: writing either evicts the other
            // (validation above already rejected an op carrying both).
            if (op.FirstLineIndent is { } firstLine)
            {
                var ind = GetOrCreatePPrChild(pPr, W.ind);
                ind.SetAttributeValue(W.firstLine, firstLine);
                ind.SetAttributeValue(W.hanging, null);
            }
            if (op.HangingIndent is { } hanging)
            {
                var ind = GetOrCreatePPrChild(pPr, W.ind);
                ind.SetAttributeValue(W.hanging, hanging);
                ind.SetAttributeValue(W.firstLine, null);
            }

            if (op.SpacingBefore is not null || op.SpacingAfter is not null || op.LineSpacing is not null)
            {
                var spacing = GetOrCreatePPrChild(pPr, W.spacing);
                // A direct beforeAutospacing/afterAutospacing flag makes Word ignore the explicit
                // value, so writing one clears the matching flag — Word's own Paragraph dialog
                // does the same when a typed value replaces "Auto".
                if (op.SpacingBefore is { } before)
                {
                    spacing.SetAttributeValue(W.before, before);
                    spacing.SetAttributeValue(W.beforeAutospacing, null);
                }
                if (op.SpacingAfter is { } after)
                {
                    spacing.SetAttributeValue(W.after, after);
                    spacing.SetAttributeValue(W.afterAutospacing, null);
                }
                if (op.LineSpacing is { } line)
                {
                    spacing.SetAttributeValue(W.line, line);
                    spacing.SetAttributeValue(W.lineRule, (op.LineSpacingRule ?? LineSpacingRule.Auto) switch
                    {
                        LineSpacingRule.Exact => "exact",
                        LineSpacingRule.AtLeast => "atLeast",
                        _ => "auto",
                    });
                }
            }

            if (op.ClearBorders is true || op.TopBorder is not null || op.BottomBorder is not null)
                ApplyParagraphBorders(pPr, op.TopBorder, op.BottomBorder, op.ClearBorders is true);

            if (_trackedChanges == TrackedChangeMode.RenderInline)
                TrackPropertyMutation(pPr, oldPPr, W.pPrChange,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate(), W.rPr, W.sectPr);

            InvalidateProjectionCache();
            // pPr-only writes (jc/ind/spacing/pBdr/pageBreakBefore) can't change an anchor's
            // kind — KindFor derives it from pStyle/numPr, which this op never touches — nor
            // its scope or unid, so the cached anchor stays valid. Skipping the eager
            // whole-document index rebuild here roughly halves this op's latency on a real
            // document (the invalidated cache refills lazily on the next lookup).
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// Insert an empty paragraph carrying a bottom border — an S-1-style horizontal rule —
    /// before/after the block named by <paramref name="anchorId"/>. <paramref name="rule"/>
    /// styles the line (default: a single 12-eighths ≈1.5pt black rule).
    /// </summary>
    /// <summary>
    /// Mint a complete, blank single-paragraph DOCX (Normal style, doc defaults, settings, and a
    /// US-Letter portrait section) as bytes — a "New document" seed for editors that draft from
    /// scratch. The result opens cleanly in Word and as a <see cref="DocxSession"/>.
    /// </summary>
    public static byte[] CreateBlankDocxBytes() => Internal.BlankDocumentFactory.CreateBytes();

    /// <summary>
    /// Insert an empty paragraph carrying a bottom border — an S-1-style horizontal rule —
    /// before/after the block named by <paramref name="anchorId"/>. <paramref name="rule"/>
    /// styles the line (default: a single 12-eighths ≈1.5pt black rule).
    /// </summary>
    public EditResult InsertHorizontalRule(string anchorId, Position pos, ParagraphBorderEdge? rule = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var edge = rule ?? new ParagraphBorderEdge { Style = "single", Size = 12, Color = "auto" };
            var pPr = new XElement(W.pPr);
            ApplyParagraphBorders(pPr, top: null, bottom: edge, clear: false);
            var p = new XElement(W.p, pPr);
            UnidHelper.AssignToSelfAndDescendants(p);

            if (pos == Position.Before) element.AddBeforeSelf(p);
            else element.AddAfterSelf(p);

            // A rule is a paragraph, so it records as one — same marking InsertParagraph applies,
            // so rejecting the revision takes the rule back out instead of leaving it behind.
            if (_trackedChanges == TrackedChangeMode.RenderInline)
                MarkParagraphContentAndMark(p, W.ins, _revisionAuthor ?? "docxodus",
                    NextTrackedFormatRevisionDate());

            var unid = (string)p.Attribute(PtOpenXml.Unid)!;
            InvalidateProjectionCache();
            var created = AnchorForUnid(unid, target.PartUri)
                ?? new Anchor($"p:{target.Anchor.Scope}:{unid}", "p", target.Anchor.Scope, unid);

            return new EditResult
            {
                Success = true,
                Created = new[] { created },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    private EditResult? RefuseNestedTrackedParagraphPropertyChange(XElement paragraph, string anchorId)
    {
        if (_trackedChanges != TrackedChangeMode.RenderInline) return null;
        return paragraph.Element(W.pPr)?.Element(W.pPrChange) is null
            ? null
            : EditResult.Fail(EditErrorCode.UnresolvedStructuralRevision,
                "paragraph has an unresolved property revision; resolve it before another tracked property mutation",
                anchorId);
    }
}
