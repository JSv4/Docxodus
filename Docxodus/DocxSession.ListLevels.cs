// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// A list's levels, as Word's <i>Define New Multilevel List</i> dialog sets them (number format,
/// number text such as <c>1.1</c>, start, positions), and <i>Continue Numbering</i>.
/// </summary>
public sealed partial class DocxSession
{
    /// <summary>
    /// The levels of the list the anchored paragraph is an item of (directly or through its style):
    /// for each level its format, number text, start, positions and legal flag. Null when the
    /// anchor names no list item.
    /// </summary>
    public IReadOnlyList<ListLevelInfo>? GetListLevels(string anchorId)
    {
        if (_disposed) return null;
        var target = FindAnchor(anchorId);
        var element = target?.Resolve(_doc!);
        if (element is null || element.Name != W.p) return null;
        var (numId, _) = EffectiveNumberingOf(element);
        if (numId is null or 0) return null;
        return Internal.NumberingFactory.ReadLevels(_doc!, numId.Value);
    }

    /// <summary>
    /// Change levels of the list the anchored paragraph is an item of: each
    /// <see cref="ListLevelSpec"/> names a level, and its non-null fields are written there. Only
    /// that list changes: when its levels come from a definition other lists share (Docxodus's own
    /// definitions, which every new list of a format reuses), the list gets a copy of its own first.
    /// A document's own definition is changed in place, as Word does: every part of that list, a
    /// restarted part too, takes the change. Missing levels up to 8 are added first. One undo step.
    /// </summary>
    /// <remarks>
    /// Refused (<see cref="EditErrorCode.InvalidListLevel"/>) for a level outside 0–8, an unknown
    /// format, a number text naming a deeper level than its own (<c>%3</c> on level 2), a negative
    /// start, and a list whose levels come from a list style (<c>w:numStyleLink</c>). Numbering
    /// definitions cannot carry tracked changes, so under
    /// <see cref="TrackedChangeMode.RenderInline"/> it is refused as other numbering edits are.
    /// </remarks>
    public EditResult SetListLevels(string anchorId, IReadOnlyList<ListLevelSpec> levels)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("SetListLevels", anchorId);
        if (ValidateLevelSpecs(levels) is { } invalid)
            return EditResult.Fail(EditErrorCode.InvalidListLevel, invalid, anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null || element.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetListLevels requires a list item", anchorId);
        var (numId, _) = EffectiveNumberingOf(element);
        if (numId is null or 0)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "no numbering on this paragraph or its style", anchorId);
        if (levels.Count == 0) return new EditResult { Success = true, Modified = new[] { target.Anchor } };
        if (Internal.NumberingFactory.DefiningAbstractNum(_doc!, numId.Value) is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, $"numbering instance {numId} is not defined", anchorId);
        if (Internal.NumberingFactory.IsStyleLinked(_doc!, numId.Value))
            return EditResult.Fail(EditErrorCode.InvalidListLevel,
                "this list takes its levels from a list style; change the style instead", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.NumberingFactory.EnsureLevelDefined(_doc!, numId.Value, 8);
            var own = Internal.NumberingFactory.OwnAbstractNum(_doc!, numId.Value);
            var num = Internal.NumberingFactory.FindNum(_doc!, numId.Value);
            if (own is null || num is null)
            {
                _ = _history.PopForUndo();
                return EditResult.Fail(EditErrorCode.InvalidListLevel, "this list's levels cannot be changed here", anchorId);
            }
            foreach (var spec in levels) Internal.NumberingFactory.ApplyLevelSpec(own, num, spec);
            _doc!.MainDocumentPart!.NumberingDefinitionsPart!.PutXDocument();
            InvalidateProjectionCache();
            _ = AnchorIndex();
            return new EditResult
            {
                Success = true,
                Modified = ListMembersOf(element, target.PartUri, (string)own.Attribute(W.abstractNumId)!),
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// Word's <i>Continue Numbering</i>: the anchored item, and the items of its list after it,
    /// join the list before it (the nearest earlier list item of another list in the same story),
    /// keeping their levels; their numbers go on from that list's. Items of its list before the
    /// anchor stay as they are. Refused when no list comes before it.
    /// </summary>
    public EditResult ContinueNumbering(string anchorId)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("ContinueNumbering", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null || element.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ContinueNumbering requires a list item", anchorId);
        var (numId, _) = EffectiveNumberingOf(element);
        if (numId is null or 0)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "no numbering on this paragraph or its style", anchorId);

        var paragraphs = element.AncestorsAndSelf().Last().Descendants(W.p).ToList();
        int at = paragraphs.IndexOf(element);
        int? previous = null;
        for (int i = at - 1; i >= 0 && previous is null; i--)
        {
            var (other, _) = EffectiveNumberingOf(paragraphs[i]);
            if (other is not null and not 0 && other != numId) previous = other;
        }
        if (previous is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "there is no list before this one to continue", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var repointed = new List<string?>();
            for (int i = at; i < paragraphs.Count; i++)
            {
                var (pNumId, pIlvl) = EffectiveNumberingOf(paragraphs[i]);
                if (pNumId != numId) continue;
                RepointListInstance(paragraphs[i], pIlvl, previous.Value);
                repointed.Add((string?)paragraphs[i].Attribute(PtOpenXml.Unid));
            }
            (ResolvePart(target.PartUri) ?? _doc!.MainDocumentPart!).PutXDocument();
            InvalidateProjectionCache();
            _ = AnchorIndex();
            var modified = new List<Anchor>();
            foreach (var unid in repointed)
            {
                if (unid is not null && AnchorForUnid(unid, target.PartUri) is { } updated) modified.Add(updated);
            }
            return new EditResult { Success = true, Modified = modified, Patch = PatchFor(target) };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>Why <paramref name="levels"/> cannot be written, or null when they can.</summary>
    private static string? ValidateLevelSpecs(IReadOnlyList<ListLevelSpec> levels)
    {
        foreach (var spec in levels)
        {
            if (spec.Level is < 0 or > 8) return $"list level {spec.Level} is outside 0-8";
            if (spec.Format is not null && !Internal.NumberingFactory.EditableLevelFormats.Contains(spec.Format))
                return $"unknown number format \"{spec.Format}\"";
            if (spec.Start is < 0) return $"level {spec.Level}: the start cannot be negative";
            if (spec.Text is { } text)
            {
                for (int i = text.IndexOf('%'); i >= 0 && i + 1 < text.Length; i = text.IndexOf('%', i + 1))
                {
                    if (char.IsDigit(text[i + 1]) && text[i + 1] - '0' is var n && (n < 1 || n > spec.Level + 1))
                        return $"level {spec.Level}: \"%{n}\" names a level below this one";
                }
            }
        }
        return null;
    }

    /// <summary>The anchors of the paragraphs in <paramref name="paragraph"/>'s part whose list
    /// takes its levels from abstract definition <paramref name="abstractId"/>.</summary>
    private List<Anchor> ListMembersOf(XElement paragraph, string partUri, string abstractId)
    {
        var modified = new List<Anchor>();
        foreach (var p in paragraph.AncestorsAndSelf().Last().Descendants(W.p))
        {
            var (pNumId, _) = EffectiveNumberingOf(p);
            if (pNumId is null or 0) continue;
            var aid = (string?)Internal.NumberingFactory.FindNum(_doc!, pNumId.Value)?.Element(W.abstractNumId)?.Attribute(W.val);
            if (aid != abstractId) continue;
            if ((string?)p.Attribute(PtOpenXml.Unid) is { } unid && AnchorForUnid(unid, partUri) is { } anchor) modified.Add(anchor);
        }
        return modified;
    }
}
