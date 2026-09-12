// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Explicit, previewable repair of native revision markup the registry refuses to resolve
/// (issues #754–#758). Ordinary listing, accept and reject never repair anything; a caller
/// reads <see cref="Propose"/> for an entry, sees exactly which carriers are defective and
/// whether a unique repair exists, and asks for it by kind. Every repair is bounded to the
/// registry entry's own carriers, is refused when the package holds no evidence of the
/// intended state, and never invents review metadata: wrapping orphan deleted text takes
/// the author and date from the caller.
/// </summary>
internal static class RevisionRepairOps
{
    /// <summary>The repairs the registry can offer for one entry — none when its diagnostic
    /// has no defined repair (a supported entry, an unsupported family, a merge state the
    /// package cannot restore).</summary>
    internal static IReadOnlyList<RevisionRepairProposal> Propose(RevisionOps.RevisionGroup group)
    {
        if (group.ResolutionStatus == RevisionResolutionStatus.Supported || group.Diagnostic is null)
            return Array.Empty<RevisionRepairProposal>();

        // Repairs key on what the defective markup IS, not only on the code the listing shows:
        // a numbering or cell marker outside its property container is inventoried as an
        // unsupported family when the walker never reached it as a unit, and as an orphan
        // when it did. Both are the same misplacement with the same unique-owner repair.
        var units = group.Units;
        if (group.Diagnostic.Code is "missing_revision_id" or "invalid_revision_id" or "duplicate_revision_id")
        {
            var carriers = Carriers(group);
            var defective = carriers.Count(c => !RevisionOps.HasCanonicalRevisionId(c));
            return new[]
            {
                Proposal(group, RevisionRepairKind.AssignIdentity, repairable: true,
                    group.Diagnostic.Code == "duplicate_revision_id"
                        ? $"Renumber every one of this entry's {carriers.Count} carriers with fresh document-unique ids; markers sharing one old id keep sharing the new one, so range pairs stay paired."
                        : $"Assign fresh document-unique ids to {defective} carrier(s) lacking a canonical w:id and renumber the rest of the entry consistently."),
            };
        }

        if (units.Count > 0 && units.All(u => u.Element.Name == W.numberingChange
                && u.Element.Parent?.Name != W.numPr && u.Element.Parent?.Name != W.fldChar))
        {
            var owner = NumberingOwner(group, out var reason);
            return new[]
            {
                Proposal(group, RevisionRepairKind.ReattachNumberingChange, owner is not null, reason),
            };
        }

        if (units.Count > 0 && units.All(u =>
                (u.Element.Name == W.cellIns || u.Element.Name == W.cellDel || u.Element.Name == W.cellMerge)
                && (u.Element.Parent?.Name != W.tcPr || u.MarkedCell is null)))
        {
            var ok = CellOwners(group, out var reason);
            return new[]
            {
                Proposal(group, RevisionRepairKind.ReattachCellMarker, ok, reason),
            };
        }

        if (IsOrphanDeletedPayload(group))
        {
            var run = units[0].Element.Parent;
            bool wrappable = run is not null && run.Name == W.r && run.Elements().All(child =>
                child.Name == W.rPr || child.Name == W.delText || child.Name == W.delInstrText);
            return new[]
            {
                Proposal(group, RevisionRepairKind.RestoreOrphanText, repairable: true,
                    "Treat the payload as live text: w:delText becomes w:t and w:delInstrText becomes w:instrText; no revision remains."),
                Proposal(group, RevisionRepairKind.WrapOrphanTextAsDeletion, wrappable,
                    wrappable
                        ? "Wrap the owning run in a w:del carrying the author and date the caller supplies, making it an ordinary deletion."
                        : "The owning run also holds live text, so wrapping it would delete text the document shows."),
            };
        }

        return Array.Empty<RevisionRepairProposal>();
    }

    /// <summary>Apply one requested repair to its entry. The caller has already checked
    /// <see cref="Propose"/> admits the kind and holds the mutation snapshot.</summary>
    internal static RevisionRepairOutcome Apply(
        RevisionOps.RevisionGroup group,
        RevisionRepairRequest request,
        Func<int> nextRevisionId)
    {
        var identities = new List<RevisionCarrierIdentity>();
        switch (request.Kind)
        {
            case RevisionRepairKind.AssignIdentity:
            {
                // Markers sharing one old id (a range pair, or a legal role reuse inside the
                // entry) keep sharing the new one; carriers with no usable id each get their own.
                var fresh = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var carrier in Carriers(group))
                {
                    var old = (string?)carrier.Attribute(W.id);
                    int assigned = old is not null && fresh.TryGetValue(old, out var mapped)
                        ? mapped
                        : nextRevisionId();
                    if (old is not null) fresh[old] = assigned;
                    carrier.SetAttributeValue(W.id, assigned);
                    identities.Add(new RevisionCarrierIdentity(CarrierKey(carrier), old,
                        assigned.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }
                break;
            }
            case RevisionRepairKind.ReattachNumberingChange:
            {
                var owner = NumberingOwner(group, out _)
                    ?? throw new InvalidOperationException("no unique numbering owner");
                foreach (var unit in group.Units)
                {
                    var marker = unit.Element;
                    var key = CarrierKey(marker);
                    marker.Remove();
                    // CT_NumPr order: ilvl, numId, numberingChange, ins.
                    if (marker.Name == W.numberingChange && owner.Element(W.ins) is { } ins)
                        ins.AddBeforeSelf(marker);
                    else
                        owner.Add(marker);
                    identities.Add(new RevisionCarrierIdentity(key, (string?)marker.Attribute(W.id),
                        (string?)marker.Attribute(W.id) ?? ""));
                }
                break;
            }
            case RevisionRepairKind.ReattachCellMarker:
            {
                if (!CellOwners(group, out _))
                    throw new InvalidOperationException("no owning table cell");
                foreach (var unit in group.Units)
                {
                    var marker = unit.Element;
                    var key = CarrierKey(marker);
                    var cell = marker.Ancestors(W.tc).First();
                    marker.Remove();
                    var tcPr = cell.Element(W.tcPr);
                    if (tcPr is null)
                    {
                        tcPr = new XElement(W.tcPr);
                        cell.AddFirst(tcPr);
                    }
                    // CT_TcPr order: … cellIns/cellDel/cellMerge, tcPrChange.
                    if (tcPr.Element(W.tcPrChange) is { } change) change.AddBeforeSelf(marker);
                    else tcPr.Add(marker);
                    identities.Add(new RevisionCarrierIdentity(key, (string?)marker.Attribute(W.id),
                        (string?)marker.Attribute(W.id) ?? ""));
                }
                break;
            }
            case RevisionRepairKind.RestoreOrphanText:
            {
                var payload = group.Units[0].Element;
                var key = CarrierKey(payload);
                var live = payload.Name == W.delText ? W.t : W.instrText;
                payload.ReplaceWith(new XElement(live, payload.Attributes(), payload.Nodes()));
                identities.Add(new RevisionCarrierIdentity(key, null, ""));
                break;
            }
            case RevisionRepairKind.WrapOrphanTextAsDeletion:
            {
                var payload = group.Units[0].Element;
                var run = payload.Parent ?? throw new InvalidOperationException("orphan payload has no run");
                var key = CarrierKey(payload);
                int id = nextRevisionId();
                var wrapper = new XElement(W.del,
                    new XAttribute(W.id, id),
                    new XAttribute(W.author, request.Author!),
                    new XAttribute(W.date, request.Date!));
                run.ReplaceWith(wrapper);
                wrapper.Add(run);
                identities.Add(new RevisionCarrierIdentity(key, null,
                    id.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                break;
            }
            default:
                throw new InvalidOperationException($"unknown repair kind {request.Kind}");
        }

        return new RevisionRepairOutcome
        {
            RevisionId = group.Id,
            Kind = request.Kind,
            PartUri = group.PartUri,
            Identities = identities,
        };
    }

    // ─── Owners ─────────────────────────────────────────────────────────

    private static XElement? NumberingOwner(RevisionOps.RevisionGroup group, out string reason)
    {
        var owners = new HashSet<XElement>();
        foreach (var unit in group.Units)
        {
            var paragraph = unit.Element.Ancestors(W.p).FirstOrDefault();
            var numPr = paragraph?.Element(W.pPr)?.Element(W.numPr);
            if (numPr is null || unit.Element.Ancestors().Contains(numPr))
            {
                reason = "The marker's paragraph carries no w:numPr the marker could be reattached to; the package records no numbering owner.";
                return null;
            }
            owners.Add(numPr);
        }
        if (owners.Count != 1)
        {
            reason = "The entry's markers belong to different paragraphs, so no single numbering owner exists.";
            return null;
        }
        reason = "Move the marker into its paragraph's own w:numPr, the unique candidate owner.";
        return owners.Single();
    }

    private static bool CellOwners(RevisionOps.RevisionGroup group, out string reason)
    {
        foreach (var unit in group.Units)
        {
            if (unit.Element.Ancestors(W.tc).FirstOrDefault() is null)
            {
                reason = "The marker is outside any table cell; the package records no cell that owns it.";
                return false;
            }
        }
        reason = "Move each marker into the w:tcPr of the nearest enclosing cell, creating the property set when absent.";
        return true;
    }

    private static bool IsOrphanDeletedPayload(RevisionOps.RevisionGroup group) =>
        group.Units.Count == 1
        && group.Units[0].Kind == RevisionOps.UnitKind.Unsupported
        && (group.Units[0].Element.Name == W.delText || group.Units[0].Element.Name == W.delInstrText);

    private static List<XElement> Carriers(RevisionOps.RevisionGroup group) =>
        group.Units.Select(u => u.Element).Concat(group.RangeMarkers).Distinct().ToList();

    internal static string CarrierKey(XElement element) =>
        RevisionOps.PrefixedName(element.Name) + "@" + RevisionOps.ElementPath(element);

    private static RevisionRepairProposal Proposal(
        RevisionOps.RevisionGroup group, RevisionRepairKind kind, bool repairable, string reason) =>
        new()
        {
            RevisionId = group.Id,
            Kind = kind,
            PartUri = group.PartUri,
            Diagnostic = group.Diagnostic!,
            Carriers = Carriers(group).Select(CarrierKey).ToList(),
            Repairable = repairable,
            Reason = reason,
            RequiresAuthorship = kind == RevisionRepairKind.WrapOrphanTextAsDeletion,
        };
}
