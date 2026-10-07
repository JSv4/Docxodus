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
    private void TrackListPropertyMutation(XElement paragraph, XElement oldPPr, bool insertedNumPr)
    {
        if (_trackedChanges != TrackedChangeMode.RenderInline) return;
        var author = _revisionAuthor ?? "docxodus";
        var date = NextTrackedFormatRevisionDate();
        var pPr = paragraph.Element(W.pPr);
        var oldBase = PropertySnapshot(oldPPr, W.pPr, W.pPrChange, W.rPr, W.sectPr);
        var newBase = PropertySnapshot(pPr, W.pPr, W.pPrChange, W.rPr, W.sectPr);
        if (XNode.DeepEquals(oldBase, newBase)) return;
        if (pPr is null)
        {
            pPr = new XElement(W.pPr);
            paragraph.AddFirst(pPr);
        }
        if (insertedNumPr && pPr.Element(W.numPr) is { } numPr)
        {
            numPr.Add(CreateRevisionEnvelope(W.ins, author, date));
            return;
        }
        pPr.Add(CreateRevisionEnvelope(W.pPrChange, author, date, oldBase));
    }

    public EditResult SetListLevel(string anchorId, int levelDelta)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetListLevel requires a list-item anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        var pPr = element.Element(W.pPr);
        var numPr = pPr?.Element(W.numPr);
        var oldPPr = new XElement(pPr ?? new XElement(W.pPr));
        bool insertedNumPr = numPr is null;

        // Resolve the effective (numId, current ilvl). A direct w:numPr wins; otherwise the
        // paragraph is a list item only via its pStyle chain (e.g. python-docx "List Bullet",
        // which carries numPr on the STYLE, not the paragraph). In that case read the effective
        // values from the style and materialize a direct w:numPr below — exactly what Word does
        // when you Tab a styled list item, and the only way to control ilvl per paragraph.
        int current;
        int? effectiveNumId;
        if (numPr is not null)
        {
            current = (int?)numPr.Element(W.ilvl)?.Attribute(W.val) ?? 0;
            effectiveNumId = (int?)numPr.Element(W.numId)?.Attribute(W.val);
        }
        else
        {
            (effectiveNumId, current) = ResolveStyleNumbering(element);
            if (effectiveNumId is null)
                return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                    "no numPr on this paragraph or its style", anchorId);
        }

        int next = current + levelDelta;
        if (next < 0 || next > 8)
            return EditResult.Fail(EditErrorCode.InvalidListLevel,
                $"resulting list level {next} out of [0,8]", anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline && effectiveNumId.HasValue
            && Internal.NumberingFactory.WouldEnsureLevelDefinedMutate(
                _doc!, effectiveNumId.Value, next))
            return TrackedStructureUnsupported(
                $"SetListLevel requiring numbering level {next} synthesis", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        // Nesting only renders if the abstractNum actually DEFINES the target level — many docs
        // define just level 0, so synthesize any missing levels before bumping ilvl.
        if (effectiveNumId.HasValue)
            Internal.NumberingFactory.EnsureLevelDefined(_doc!, effectiveNumId.Value, next);

        if (numPr is not null)
        {
            numPr.Element(W.ilvl)?.Remove();
            numPr.AddFirst(new XElement(W.ilvl, new XAttribute(W.val, next))); // ilvl precedes numId
        }
        else
        {
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            SetPPrChildInOrder(pPr, new XElement(W.numPr,
                new XElement(W.ilvl, new XAttribute(W.val, next)),
                new XElement(W.numId, new XAttribute(W.val, effectiveNumId!.Value))));
        }
        TrackListPropertyMutation(element, oldPPr, insertedNumPr);
        // Flush the body mutation to the part stream immediately — same as NumberingFactory does for
        // the numbering part. Without this the materialized w:numPr lives only in the in-memory
        // XDocument; under WASM the typed-DOM/XDocument divergence means a later Save() serializes
        // the un-flushed state and the nest silently vanishes on save and re-render. (Body lists are
        // body-scoped; flushing the main part covers them.)
        _doc!.MainDocumentPart!.PutXDocument();
        InvalidateProjectionCache();
        return new EditResult
        {
            Success = true,
            Modified = new[] { target.Anchor },
            Patch = PatchFor(target),
        };
    }

    /// <summary>
    /// Resolve the effective <c>(numId, ilvl)</c> a paragraph inherits from its pStyle chain, for
    /// a list item whose numbering comes from a style rather than a direct <c>w:numPr</c>. Walks
    /// <c>basedOn</c> (cycle-guarded). Returns <c>(null, 0)</c> when no style contributes a numId.
    /// </summary>
    private (int? numId, int ilvl) ResolveStyleNumbering(XElement paragraph)
    {
        var styleId = (string?)paragraph.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (string.IsNullOrEmpty(styleId)) return (null, 0);
        var stylesRoot = _doc!.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
        if (stylesRoot is null) return (null, 0);

        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        for (int i = 0; i < 16 && current is not null; i++)
        {
            if (!visited.Add(current)) break; // cycle
            var style = stylesRoot.Elements(W.style)
                .FirstOrDefault(s => (string?)s.Attribute(W.styleId) == current);
            if (style is null) break;
            var styleNumPr = style.Element(W.pPr)?.Element(W.numPr);
            var numId = (int?)styleNumPr?.Element(W.numId)?.Attribute(W.val);
            if (numId is not null)
                return (numId, (int?)styleNumPr!.Element(W.ilvl)?.Attribute(W.val) ?? 0);
            current = (string?)style.Element(W.basedOn)?.Attribute(W.val);
        }
        return (null, 0);
    }

    public EditResult RemoveListMembership(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "RemoveListMembership requires a paragraph, heading, or list-item anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        var pPr = element.Element(W.pPr);
        var oldPPr = new XElement(pPr ?? new XElement(W.pPr));
        var directNumPr = pPr?.Element(W.numPr);
        // Removing a direct numPr can expose numbering inherited from the paragraph style.
        // Materialize Word's numId=0 sentinel in that case so the explicit removal wins over
        // the style chain. This also lets callers create an unnumbered heading in documents
        // whose Heading styles carry legal-outline numbering.
        bool needsStyleOverride = ResolveStyleNumbering(element).numId is not null;

        _history.RecordPreOp(TakeSnapshot());
        directNumPr?.Remove();
        if (needsStyleOverride)
        {
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            SetPPrChildInOrder(pPr, new XElement(W.numPr,
                new XElement(W.ilvl, new XAttribute(W.val, 0)),
                new XElement(W.numId, new XAttribute(W.val, 0))));
        }
        TrackListPropertyMutation(element, oldPPr, insertedNumPr: false);
        InvalidateProjectionCache();
        var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;
        return new EditResult
        {
            Success = true,
            Modified = new[] { updated },
            Patch = PatchFor(target),
        };
    }

    /// <summary>
    /// Make the paragraph a bullet or numbered list item, or remove list membership.
    /// Unlike <see cref="SetListLevel"/>/<see cref="RemoveListMembership"/> (which require an
    /// existing list item), this PROMOTES a plain paragraph: it ensures a reusable numbering
    /// definition exists (synthesizing one in the numbering part if needed) and sets the
    /// paragraph's <c>w:numPr</c>. <see cref="ListFormat.None"/> strips inline list membership.
    /// </summary>
    public EditResult ApplyListFormat(string anchorId, ListFormat kind)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ApplyListFormat requires a paragraph anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        var oldPPr = new XElement(element.Element(W.pPr) ?? new XElement(W.pPr));
        bool insertedNumPr = oldPPr.Element(W.numPr) is null && kind != ListFormat.None;
        if (_trackedChanges == TrackedChangeMode.RenderInline
            && kind != ListFormat.None && !insertedNumPr
            && Internal.NumberingFactory.WouldEnsureNumberingMutate(_doc!, kind))
            return TrackedStructureUnsupported(
                $"ApplyListFormat requiring synthesis of {kind} numbering", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (kind == ListFormat.None)
            {
                pPr?.Element(W.numPr)?.Remove();
            }
            else
            {
                if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
                int numId = Internal.NumberingFactory.EnsureNumbering(_doc!, kind);
                int ilvl = (int?)pPr.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
                pPr.Element(W.numPr)?.Remove();
                SetPPrChildInOrder(pPr, new XElement(W.numPr,
                    new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
                    new XElement(W.numId, new XAttribute(W.val, numId))));
            }

            TrackListPropertyMutation(element, oldPPr, insertedNumPr);

            InvalidateProjectionCache();
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

    /// <summary>
    /// <see cref="ApplyListFormat"/> across a contiguous sibling run of paragraphs, from
    /// <paramref name="firstAnchorId"/> to <paramref name="lastAnchorId"/> INCLUSIVE (they may
    /// be the same anchor, and may be passed in either document order). Every member gets the
    /// same shared <c>w:num</c> instance, so the numbering sequence stays intact — the per-item
    /// op cannot guarantee that. One snapshot is recorded, so the whole range is a single
    /// <see cref="Undo"/> step. Non-paragraph siblings inside the range (a table, an sdt) are
    /// skipped — they cannot carry <c>w:numPr</c>. Each paragraph keeps its own <c>w:ilvl</c>,
    /// so a nested run converts in place. <see cref="ListFormat.None"/> strips inline list
    /// membership from every member.
    /// </summary>
    public EditResult ApplyListFormatRange(string firstAnchorId, string lastAnchorId, ListFormat kind)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("ApplyListFormatRange", firstAnchorId);
        var firstTarget = FindAnchor(firstAnchorId);
        if (firstTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"first anchor not found: {firstAnchorId}", firstAnchorId);
        var lastTarget = FindAnchor(lastAnchorId);
        if (lastTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"last anchor not found: {lastAnchorId}", lastAnchorId);
        if (firstTarget.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ApplyListFormatRange requires paragraph anchors; first kind={firstTarget.Anchor.Kind}", firstAnchorId);
        if (lastTarget.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ApplyListFormatRange requires paragraph anchors; last kind={lastTarget.Anchor.Kind}", lastAnchorId);
        if (firstTarget.Anchor.Scope != lastTarget.Anchor.Scope)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                $"ApplyListFormatRange anchors must live in the same package part; first={firstTarget.Anchor.Scope} last={lastTarget.Anchor.Scope}",
                firstAnchorId);

        var firstElement = firstTarget.Resolve(_doc!);
        var lastElement = lastTarget.Resolve(_doc!);
        if (firstElement is null || lastElement is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", firstAnchorId);
        if (firstElement.Parent != lastElement.Parent)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "ApplyListFormatRange anchors must share a direct parent (no spanning into nested containers)",
                firstAnchorId);

        // Normalize order: same parent is established, so if `last` is not a following sibling
        // of `first`, the caller passed them reversed — swap rather than erroring.
        if (firstElement != lastElement && !firstElement.ElementsAfterSelf().Contains(lastElement))
            (firstElement, lastElement) = (lastElement, firstElement);

        // The w:p members of the run, first..last inclusive, with their unids captured pre-op
        // so the post-op anchors (kind may flip p↔li) can be reported in Modified.
        var members = new List<XElement>();
        for (var cursor = firstElement; cursor is not null; cursor = cursor.ElementsAfterSelf().FirstOrDefault())
        {
            if (cursor.Name == W.p) members.Add(cursor);
            if (cursor == lastElement) break;
        }
        var memberUnids = members.Select(m => (string?)m.Attribute(PtOpenXml.Unid)).ToList();
        var partUri = firstTarget.PartUri;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (kind == ListFormat.None)
            {
                foreach (var member in members)
                    member.Element(W.pPr)?.Element(W.numPr)?.Remove();
            }
            else
            {
                // One find-or-create up front — every member points at the SAME numId.
                int numId = Internal.NumberingFactory.EnsureNumbering(_doc!, kind);
                foreach (var member in members)
                {
                    var pPr = member.Element(W.pPr);
                    if (pPr is null) { pPr = new XElement(W.pPr); member.AddFirst(pPr); }
                    int ilvl = (int?)pPr.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
                    pPr.Element(W.numPr)?.Remove();
                    SetPPrChildInOrder(pPr, new XElement(W.numPr,
                        new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
                        new XElement(W.numId, new XAttribute(W.val, numId))));
                }
            }

            InvalidateProjectionCache();
            _ = AnchorIndex();
            var modified = new List<Anchor>();
            foreach (var unid in memberUnids)
            {
                if (unid is not null && AnchorForUnid(unid, partUri) is { } updated)
                    modified.Add(updated);
            }
            return new EditResult
            {
                Success = true,
                Modified = modified,
                Patch = PatchFor(firstTarget),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, firstAnchorId);
        }
    }

    /// <summary>
    /// Restart (or seed) the anchored list item's numbering at <paramref name="value"/> — Word's
    /// <em>Set Numbering Value… → Set value to</em> (issue #314). Writes a
    /// <c>w:lvlOverride[@w:ilvl]/w:startOverride[@w:val]</c> on a DEDICATED <c>w:num</c> instance:
    /// the item's current num is cloned (never mutated because it may be shared), and the anchored
    /// paragraph plus every FOLLOWING paragraph of
    /// the same numbering instance in the part is repointed at the clone. An anchored item
    /// mid-sequence therefore splits the sequence exactly like Word: earlier items keep their
    /// numbers, the anchored item shows <paramref name="value"/>, and the tail continues from it.
    /// Style-derived members get a direct <c>w:numPr</c> materialized (ilvl preserved), the same
    /// way <see cref="SetListLevel"/> does. Undo restores every repointed paragraph.
    /// </summary>
    public EditResult SetListStartOverride(string anchorId, int value)
    {
        if (value < 0)
            return EditResult.Fail(EditErrorCode.InvalidListStartValue,
                $"list start value cannot be negative (got {value})", anchorId);
        return ApplyListStartOverride(anchorId, value);
    }

    /// <summary>
    /// Remove the numbering restart from the anchored item's list sequence — the inverse of
    /// <see cref="SetListStartOverride"/>. EVERY paragraph of the same numbering instance in the
    /// part (before and after the anchor — they move together, so relative continuation is
    /// preserved) is repointed at a clone of the instance WITHOUT the
    /// <c>w:startOverride</c> at the item's level; the sequence reverts to the abstract
    /// definition's own <c>w:start</c>. A sequence with no override at the item's level is a
    /// successful no-op that consumes no undo history.
    /// </summary>
    public EditResult ClearListStartOverride(string anchorId) =>
        ApplyListStartOverride(anchorId, null);

    /// <summary>Shared engine for <see cref="SetListStartOverride"/> (split the sequence at the
    /// anchor onto a clone carrying the override) and <see cref="ClearListStartOverride"/>
    /// (move the whole sequence onto a clone without it).</summary>
    private EditResult ApplyListStartOverride(string anchorId, int? value)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported(
                value is null ? "ClearListStartOverride" : "SetListStartOverride", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "SetListStartOverride requires a list-item anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var (numId, ilvl) = EffectiveNumberingOf(element);
        // numId 0 is OOXML for "numbering removed" — not an instance a start override can target.
        if (numId is null or 0)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "no numPr on this paragraph or its style", anchorId);

        // Clearing a sequence that has no override at this level is a no-op; return BEFORE
        // TakeSnapshot so it cannot evict real edits from the bounded undo ring.
        if (value is null && Internal.NumberingFactory.GetStartOverride(_doc!, numId.Value, ilvl) is null)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var newNumId = Internal.NumberingFactory.CloneNumWithStartOverride(_doc!, numId.Value, ilvl, value);
            if (newNumId is null)
            {
                _ = _history.PopForUndo();
                return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                    $"numbering instance {numId} is not defined in the numbering part", anchorId);
            }

            // Repoint the sequence: for a set, the anchored paragraph and everything after it
            // (the split Word performs); for a clear, every member (the whole sequence moves).
            var partRoot = element.AncestorsAndSelf().Last();
            var repointedUnids = new List<string?>();
            bool reached = false;
            foreach (var p in partRoot.Descendants(W.p))
            {
                if (p == element) reached = true;
                if (value is not null && !reached) continue;
                var (pNumId, pIlvl) = EffectiveNumberingOf(p);
                if (pNumId != numId) continue;
                RepointListInstance(p, pIlvl, newNumId.Value);
                repointedUnids.Add((string?)p.Attribute(PtOpenXml.Unid));
            }

            // Flush the body mutation to the part stream immediately — same WASM typed-DOM /
            // XDocument divergence rationale as SetListLevel.
            (ResolvePart(target.PartUri) ?? _doc!.MainDocumentPart!).PutXDocument();
            InvalidateProjectionCache();
            _ = AnchorIndex();
            var modified = new List<Anchor>();
            foreach (var unid in repointedUnids)
            {
                if (unid is not null && AnchorForUnid(unid, target.PartUri) is { } updated)
                    modified.Add(updated);
            }
            return new EditResult
            {
                Success = true,
                Modified = modified,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// The effective <c>(numId, ilvl)</c> of a paragraph: a direct <c>w:numPr</c> wins per
    /// attribute (its <c>w:numId</c>/<c>w:ilvl</c> each fall back to the pStyle chain via
    /// <see cref="ResolveStyleNumbering"/> when the child is absent). <c>(null, 0)</c> when
    /// neither contributes a numId.
    /// </summary>
    private (int? numId, int ilvl) EffectiveNumberingOf(XElement paragraph)
    {
        var numPr = paragraph.Element(W.pPr)?.Element(W.numPr);
        var directNumId = (int?)numPr?.Element(W.numId)?.Attribute(W.val);
        var directIlvl = (int?)numPr?.Element(W.ilvl)?.Attribute(W.val);
        if (directNumId is not null) return (directNumId, directIlvl ?? 0);
        var (styleNumId, styleIlvl) = ResolveStyleNumbering(paragraph);
        return (styleNumId, directIlvl ?? styleIlvl);
    }

    /// <summary>Point <paramref name="paragraph"/>'s numbering at <paramref name="newNumId"/>,
    /// keeping its effective <paramref name="ilvl"/> — editing the direct <c>w:numPr</c> in place
    /// when one carries a <c>w:numId</c>, else materializing one (the style-derived case).</summary>
    private void RepointListInstance(XElement paragraph, int ilvl, int newNumId)
    {
        var pPr = paragraph.Element(W.pPr);
        var numPr = pPr?.Element(W.numPr);
        if (numPr?.Element(W.numId) is { } numIdEl)
        {
            numIdEl.SetAttributeValue(W.val, newNumId);
            return;
        }
        if (pPr is null) { pPr = new XElement(W.pPr); paragraph.AddFirst(pPr); }
        pPr.Element(W.numPr)?.Remove();
        SetPPrChildInOrder(pPr, new XElement(W.numPr,
            new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
            new XElement(W.numId, new XAttribute(W.val, newNumId))));
    }
}
