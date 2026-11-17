// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

public sealed partial class DocxSession
{
    /// <summary>
    /// Make the anchored paragraph an item of an existing list: numbering instance
    /// <paramref name="numId"/> (a <see cref="ListMembership.NumId"/> read from another item) at
    /// <paramref name="level"/> (0-8). The paragraph then numbers on in that list, as Word's Format
    /// Painter does when it copies a numbered paragraph's formatting. A plain paragraph becomes a
    /// list item; a list item moves to that list. Refused when the document has no such
    /// <c>w:num</c> or the level is outside 0-8. One undo step.
    /// </summary>
    public EditResult SetListMembership(string anchorId, int numId, int level)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetListMembership requires a paragraph anchor", anchorId);
        if (level is < 0 or > 8)
            return EditResult.Fail(EditErrorCode.InvalidListLevel, $"list level {level} is outside 0-8", anchorId);
        var numbering = _doc!.MainDocumentPart?.NumberingDefinitionsPart?.GetXDocument().Root;
        bool known = numId > 0 && numbering is not null
            && numbering.Elements(W.num).Any(n => (int?)n.Attribute(W.numId) == numId);
        if (!known)
            return EditResult.Fail(EditErrorCode.InvalidListLevel, $"the document has no list {numId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        var oldPPr = new XElement(element.Element(W.pPr) ?? new XElement(W.pPr));
        var oldNumPr = oldPPr.Element(W.numPr);
        if ((int?)oldNumPr?.Element(W.numId)?.Attribute(W.val) == numId
            && ((int?)oldNumPr?.Element(W.ilvl)?.Attribute(W.val) ?? 0) == level)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            pPr.Element(W.numPr)?.Remove();
            SetPPrChildInOrder(pPr, new XElement(W.numPr,
                new XElement(W.ilvl, new XAttribute(W.val, level)),
                new XElement(W.numId, new XAttribute(W.val, numId))));
            TrackListPropertyMutation(element, oldPPr, insertedNumPr: oldNumPr is null);
            InvalidateProjectionCache();
            _ = AnchorIndex();
            var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;
            return new EditResult { Success = true, Modified = new[] { updated }, Patch = PatchFor(target) };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }
}
