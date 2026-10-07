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
    // ─── Tracked revisions: markup-native listing + selective resolution (issue #318) ───

    /// <summary>
    /// Enumerate the document's tracked revisions from a live, part-aware registry.
    /// Cell structure, content-control envelopes, and numbering families are atomic
    /// entries alongside the existing content/move/property families. Bad native
    /// markup remains visible with a diagnostic and cannot be resolved accidentally.
    /// </summary>
    public IReadOnlyList<RevisionListEntry> ListRevisions()
        => GetRevisionInventory().Entries;

    /// <summary>
    /// Build the public revision view and the exact native carrier count from one registry pass.
    /// The latter counts units and range markers rather than UI groups, so coalescing cannot hide
    /// adversarial work from verification callers.
    /// </summary>
    internal (IReadOnlyList<RevisionListEntry> Entries, long NativeElementCount,
        bool Complete, bool EvidenceComplete) GetRevisionInventory(
            int maximumNativeElements = int.MaxValue,
            int maximumEvidenceItems = int.MaxValue,
            long maximumEvidenceTextCharacters = long.MaxValue)
    {
        ThrowIfDisposed();

        // Count native carriers before grouping/sorting/hashing them. The manifest's compact
        // facts intentionally do not model every range-marker family, so relying on it alone
        // would let a single coalesced move name hide adversarial registry work.
        var parts = RevisionStoryParts();
        long carrierCount = 0;
        foreach (var (_, root, _) in parts)
        {
            foreach (var element in root.DescendantsAndSelf())
            {
                if (!Internal.RevisionOps.IsNativeRevisionCarrierForInventory(element)) continue;
                carrierCount++;
                if (carrierCount > maximumNativeElements)
                    return (Array.Empty<RevisionListEntry>(), carrierCount, false, true);
            }
        }

        _ = AnchorIndex(); // guarantees Unids so entries can carry block anchors
        var registry = BuildRevisionRegistry(parts);
        var result = new List<RevisionListEntry>(registry.Entries.Count);
        int evidenceItems = 0;
        long evidenceTextCharacters = 0;
        long remainingAnchorTraversal = maximumEvidenceItems == int.MaxValue
            ? long.MaxValue
            : Math.Max(1024L, checked((long)maximumEvidenceItems * 8));
        foreach (var g in registry.Entries)
        {
            var constituentIds = Internal.RevisionOps.ConstituentIds(g);
            var constituentKeys = Internal.RevisionOps.ConstituentKeys(g);
            if (evidenceItems > maximumEvidenceItems - constituentIds.Count
                || evidenceItems + constituentIds.Count
                    > maximumEvidenceItems - constituentKeys.Count)
                return (Array.Empty<RevisionListEntry>(), carrierCount, true, false);
            evidenceItems += constituentIds.Count + constituentKeys.Count;

            bool AddText(params string?[] values)
            {
                foreach (var value in values)
                {
                    if (value is null) continue;
                    if (evidenceTextCharacters
                        > maximumEvidenceTextCharacters - value.Length)
                        return false;
                    evidenceTextCharacters += value.Length;
                }
                return true;
            }

            if (!AddText(
                    g.Id, g.PartUri, g.Scope, g.Type, g.Author, g.Date, g.DateUtc,
                    g.Diagnostic?.Code, g.Diagnostic?.Message)
                || constituentIds.Any(value => !AddText(value))
                || constituentKeys.Any(value => !AddText(value)))
                return (Array.Empty<RevisionListEntry>(), carrierCount, true, false);

            int remainingAnchors = maximumEvidenceItems - evidenceItems;
            if (!TryRevisionGroupAnchors(
                    g, g.PartUri, remainingAnchors,
                    ref remainingAnchorTraversal, out var affected))
                return (Array.Empty<RevisionListEntry>(), carrierCount, true, false);
            evidenceItems += affected.Count;
            if (affected.Any(anchor => !AddText(anchor.Id)))
                return (Array.Empty<RevisionListEntry>(), carrierCount, true, false);

            long remainingText = maximumEvidenceTextCharacters - evidenceTextCharacters;
            var groupText = Internal.RevisionOps.GroupText(
                g, remainingText, out var groupTextComplete);
            if (!groupTextComplete || !AddText(groupText))
                return (Array.Empty<RevisionListEntry>(), carrierCount, true, false);
            result.Add(new RevisionListEntry
            {
                Id = g.Id,
                Type = g.Type,
                Family = g.Family,
                ConstituentIds = constituentIds,
                ConstituentKeys = constituentKeys,
                Author = g.Author,
                Date = g.Date,
                DateUtc = g.DateUtc,
                Text = groupText,
                PartUri = g.PartUri,
                Scope = g.Scope,
                AnchorId = affected.Count == 0 ? null : affected[0].Id,
                AffectedAnchors = affected,
                ResolutionStatus = g.ResolutionStatus,
                Diagnostic = g.Diagnostic,
            });
        }
        return (result, carrierCount, true, true);
    }

    /// <summary>Accept ONE revision by the id <see cref="ListRevisions"/> reported —
    /// insertions keep their content (markup unwrapped), deletions are carried out,
    /// a move materializes at its destination, a format change keeps the new
    /// properties. An undoable session mutation; every other revision's markup (and
    /// id) is left untouched.</summary>
    public EditResult AcceptRevision(string revisionId) => ResolveRevision(revisionId, accept: true);

    /// <summary>Reject ONE revision by id — the inverse of <see cref="AcceptRevision"/>:
    /// insertions are removed, deleted content is restored (<c>w:delText</c> back to
    /// <c>w:t</c>), a move stays at its source, a format change restores the stored
    /// old properties.</summary>
    public EditResult RejectRevision(string revisionId) => ResolveRevision(revisionId, accept: false);

    /// <summary>Accept every live revision through the same fail-closed resolver used by
    /// <see cref="AcceptRevision"/>. The complete operation is one undo step.</summary>
    public EditResult AcceptAllRevisions() => ResolveAllRevisions(accept: true);

    /// <summary>Reject every live revision through the same fail-closed resolver used by
    /// <see cref="RejectRevision"/>. The complete operation is one undo step.</summary>
    public EditResult RejectAllRevisions() => ResolveAllRevisions(accept: false);

    private EditResult ResolveRevision(string revisionId, bool accept)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (string.IsNullOrEmpty(revisionId))
            return EditResult.Fail(EditErrorCode.RevisionNotFound, "revision id is empty");

        _ = AnchorIndex();
        var registry = BuildRevisionRegistry();
        var group = registry.Find(revisionId);
        if (group is null)
            return EditResult.Fail(EditErrorCode.RevisionNotFound, $"revision not found: {revisionId}");

        if (RevisionResolutionError(group) is { } resolutionError)
            return resolutionError;

        var partUri = group.PartUri;
        var owningPart = ResolvePart(partUri);
        var relationshipCandidates = UnprotectedRevisionRelationshipIds(group, owningPart);
        var rejectedNumberingIds = accept
            ? Array.Empty<int>()
            : NumberingIdsIntroducedBy(group)
                .Except(_settings.ProtectedRevisionNumberingIds).ToArray();

        // Capture the block anchors the resolution touches BEFORE applying — elements
        // detach during Apply and can no longer be resolved to a part afterwards.
        var modified = RevisionGroupAnchors(group, partUri);
        var referencedNotesBefore = ReferencedNoteIds();

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var removedElements = registry.Resolve(
                group,
                accept,
                _settings.ProofSafeRevisionResolution,
                _settings.ProtectedRevisionEmptyContainerKeys);
            // Numbering is intentionally outside ordinary projected saves: merely opening and
            // saving a document must not reserialize numbering.xml. A revision in that part
            // mutates its cached XDocument directly, so flush only that selected owner here.
            if (owningPart is NumberingDefinitionsPart)
                owningPart.PutXDocument();
            if (owningPart is not null)
            {
                SweepOrphanedStoryRelationships(owningPart, relationshipCandidates);
            }
            if (!accept && rejectedNumberingIds.Length > 0)
                PruneUnreferencedDocxodusNumbering(rejectedNumberingIds);
            var prunedNotes = PruneOrphanedNotes(referencedNotesBefore);

            var removed = new List<Anchor>();
            var seenRemoved = new HashSet<string>(StringComparer.Ordinal);
            foreach (var el in removedElements)
            {
                foreach (var d in el.DescendantsAndSelf())
                {
                    var unid = (string?)d.Attribute(PtOpenXml.Unid);
                    if (unid is null) continue;
                    if (AnchorForUnid(unid, partUri) is { } anch && seenRemoved.Add(anch.Id))
                        removed.Add(anch);
                }
            }

            AppendPrunedNoteAnchors(prunedNotes, removed, seenRemoved);

            // The relationship sweep above is limited to ids carried by this group, so an
            // unrelated orphan survives. Media is still swept at the mutation boundary as every
            // edit does — a pruned note's image lives in a part no group root reaches — except
            // under a whole-package reversibility proof, where pre-existing orphans must stay.
            InvalidateProjectionCache(sweepOrphanedImages: !_settings.ProofSafeRevisionResolution);
            return new EditResult
            {
                Success = true,
                Modified = modified.Where(m => !seenRemoved.Contains(m.Id)).ToList(),
                Removed = removed,
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    private EditResult ResolveAllRevisions(bool accept)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        _ = AnchorIndex();
        var registry = BuildRevisionRegistry();
        if (registry.Entries.Count == 0)
            return new EditResult { Success = true };

        var blocked = registry.Entries.FirstOrDefault(g =>
            g.ResolutionStatus != RevisionResolutionStatus.Supported);
        if (blocked is not null)
            return RevisionResolutionError(blocked)!;

        var modified = registry.Entries.SelectMany(g => RevisionGroupAnchors(g, g.PartUri))
            .GroupBy(a => a.Id, StringComparer.Ordinal).Select(g => g.First()).ToList();
        var relationshipCandidates = registry.Entries
            .GroupBy(group => group.PartUri, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group =>
                {
                    var owner = ResolvePart(group.Key);
                    return (IReadOnlyCollection<string>)group
                        .SelectMany(revision => UnprotectedRevisionRelationshipIds(revision, owner))
                        .Distinct(StringComparer.Ordinal).ToArray();
                },
                StringComparer.Ordinal);
        var rejectedNumberingIds = accept
            ? Array.Empty<int>()
            : registry.Entries.SelectMany(NumberingIdsIntroducedBy)
                .Except(_settings.ProtectedRevisionNumberingIds).Distinct().ToArray();
        var referencedNotesBefore = ReferencedNoteIds();

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var removedElements = registry.ResolveAll(
                accept,
                _settings.ProofSafeRevisionResolution,
                _settings.ProtectedRevisionEmptyContainerKeys);
            foreach (var story in RevisionStoryParts())
                if (story.Part is NumberingDefinitionsPart)
                    story.Part.PutXDocument();
            // Sweeping one story can delete another: accepting a deleted section-break
            // paragraph mark orphans the header/footer parts its w:sectPr referenced. Resolve
            // each owner afresh so a part an earlier sweep removed is skipped, not touched.
            foreach (var (partUri, candidates) in relationshipCandidates)
                if (ResolvePart(partUri) is { } owner)
                    SweepOrphanedStoryRelationships(owner, candidates);
            if (!accept && rejectedNumberingIds.Length > 0)
                PruneUnreferencedDocxodusNumbering(rejectedNumberingIds);
            var prunedNotes = PruneOrphanedNotes(referencedNotesBefore);
            var removed = new List<Anchor>();
            var seenRemoved = new HashSet<string>(StringComparer.Ordinal);
            foreach (var element in removedElements)
            {
                var partUri = PartUriOf(element) ?? registry.Entries
                    .FirstOrDefault(g => g.Units.Any(u => ReferenceEquals(u.Element, element)
                        || ReferenceEquals(u.MarkedCell, element)
                        || ReferenceEquals(u.MarkedRow, element)
                        || ReferenceEquals(u.StructuredWrapper, element)))?.PartUri;
                foreach (var descendant in element.DescendantsAndSelf())
                {
                    var unid = (string?)descendant.Attribute(PtOpenXml.Unid);
                    if (unid is null) continue;
                    if (AnchorForUnid(unid, partUri) is { } anchor && seenRemoved.Add(anchor.Id))
                        removed.Add(anchor);
                }
            }

            AppendPrunedNoteAnchors(prunedNotes, removed, seenRemoved);

            // ResolveAll has already swept only relationships referenced by the groups it
            // resolved; media orphaned in another part (a pruned note's image) is swept at the
            // mutation boundary unless a reversibility proof needs unrelated orphans kept.
            InvalidateProjectionCache(sweepOrphanedImages: !_settings.ProofSafeRevisionResolution);
            return new EditResult
            {
                Success = true,
                Modified = modified.Where(a => !seenRemoved.Contains(a.Id)).ToList(),
                Removed = removed,
            };
        }
        catch (Internal.RevisionResolutionException ex)
        {
            RollbackFailedOp();
            return RevisionResolutionError(ex.Group)!;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>
    /// The explicit repairs the registry offers for entries it refuses to resolve (issues
    /// #754–#758): which carriers are defective, whether a unique repair exists, and what it
    /// would do. Read-only; ordinary listing, accept and reject never repair.
    /// </summary>
    public IReadOnlyList<RevisionRepairProposal> ListRevisionRepairs()
    {
        ThrowIfDisposed();
        _ = AnchorIndex();
        return BuildRevisionRegistry().Entries
            .SelectMany(Internal.RevisionRepairOps.Propose)
            .ToList();
    }

    /// <summary>
    /// Perform requested repairs atomically as one undo step. Each request names a listed entry
    /// and a kind <see cref="ListRevisionRepairs"/> offered as repairable; anything else, or a
    /// wrap-as-deletion without author and date, refuses the whole call with
    /// <see cref="EditErrorCode.RevisionRepairRejected"/> and no mutation. Afterwards the repaired
    /// carriers must no longer carry their diagnostic, or the call rolls back.
    /// </summary>
    public RevisionRepairResult RepairRevisions(IReadOnlyList<RevisionRepairRequest> repairs)
    {
        if (_disposed) return RevisionRepairResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (repairs is null || repairs.Count == 0)
            return RevisionRepairResult.Fail(EditErrorCode.RevisionRepairRejected, "no repairs requested");

        _ = AnchorIndex();
        var registry = BuildRevisionRegistry();
        var planned = new List<(Internal.RevisionOps.RevisionGroup Group, RevisionRepairRequest Request)>();
        foreach (var request in repairs)
        {
            var group = registry.Find(request.RevisionId);
            if (group is null)
                return RevisionRepairResult.Fail(EditErrorCode.RevisionNotFound,
                    $"revision not found: {request.RevisionId}");
            if (planned.Any(p => ReferenceEquals(p.Group, group)))
                return RevisionRepairResult.Fail(EditErrorCode.RevisionRepairRejected,
                    $"revision {request.RevisionId} is named more than once");
            var proposal = Internal.RevisionRepairOps.Propose(group)
                .FirstOrDefault(candidate => candidate.Kind == request.Kind);
            if (proposal is null)
                return RevisionRepairResult.Fail(EditErrorCode.RevisionRepairRejected,
                    $"revision {request.RevisionId} ({group.Diagnostic?.Code ?? "supported"}) offers no {request.Kind} repair");
            if (!proposal.Repairable)
                return RevisionRepairResult.Fail(EditErrorCode.RevisionRepairRejected,
                    $"revision {request.RevisionId} cannot be repaired by {request.Kind}: {proposal.Reason}");
            if (proposal.RequiresAuthorship)
            {
                if (string.IsNullOrWhiteSpace(request.Author))
                    return RevisionRepairResult.Fail(EditErrorCode.RevisionRepairRejected,
                        $"{request.Kind} requires an author; the registry never fabricates review metadata");
                if (request.Date is null || !Internal.RevisionOps.IsValidRevisionDate(request.Date))
                    return RevisionRepairResult.Fail(EditErrorCode.RevisionRepairRejected,
                        $"{request.Kind} requires a canonical XML Schema date-time; got '{request.Date}'");
            }
            planned.Add((group, request));
        }

        var modified = planned.SelectMany(p => RevisionGroupAnchors(p.Group, p.Group.PartUri))
            .GroupBy(a => a.Id, StringComparer.Ordinal).Select(g => g.First()).ToList();
        var touchedParts = planned.Select(p => ResolvePart(p.Group.PartUri))
            .Where(part => part is not null).Distinct().ToList();

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var outcomes = new List<RevisionRepairOutcome>();
            foreach (var (group, request) in planned)
                outcomes.Add(Internal.RevisionRepairOps.Apply(group, request, NextRevisionId));
            foreach (var part in touchedParts)
                if (part is NumberingDefinitionsPart) part!.PutXDocument();

            // The repair must have removed the defect it addressed; a carrier still listed under
            // the same diagnostic means the package was not what the proposal assumed.
            var after = BuildRevisionRegistry();
            foreach (var (group, request) in planned)
            {
                var carriers = group.Units.Select(u => u.Element).Concat(group.RangeMarkers).ToHashSet();
                var lingering = after.Entries.FirstOrDefault(entry =>
                    entry.Diagnostic?.Code == group.Diagnostic?.Code
                    && entry.Units.Select(u => u.Element).Concat(entry.RangeMarkers).Any(carriers.Contains));
                if (lingering is not null)
                    throw new InvalidOperationException(
                        $"repair {request.Kind} of {request.RevisionId} left {lingering.Diagnostic!.Code} in place");
            }

            InvalidateProjectionCache();
            return new RevisionRepairResult { Success = true, Repairs = outcomes, Modified = modified };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return RevisionRepairResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    private static EditResult? RevisionResolutionError(Internal.RevisionOps.RevisionGroup group)
    {
        var code = group.ResolutionStatus switch
        {
            RevisionResolutionStatus.Unsupported => EditErrorCode.RevisionUnsupported,
            RevisionResolutionStatus.Malformed => EditErrorCode.RevisionMalformed,
            RevisionResolutionStatus.Ambiguous => EditErrorCode.RevisionAmbiguous,
            _ => (EditErrorCode?)null,
        };
        return code is null ? null : EditResult.Fail(code.Value,
            group.Diagnostic?.Message ?? "revision cannot be resolved safely");
    }

    /// <summary>Numbering instances a revision brought in: a <c>w:numPr/w:ins</c> mark, or the
    /// numbering of a wholly inserted paragraph, which leaves with the paragraph on reject.</summary>
    private static IReadOnlyList<int> NumberingIdsIntroducedBy(
        Internal.RevisionOps.RevisionGroup group) =>
        group.Units.Select(unit =>
                unit.Kind == Internal.RevisionOps.UnitKind.NumberingPropertiesInsert
                    ? unit.Element.Parent?.Element(W.numId)
                    : unit.Kind == Internal.RevisionOps.UnitKind.ParaMark
                        && unit.Element.Name == W.ins
                        ? unit.Paragraph?.Element(W.pPr)?.Element(W.numPr)?.Element(W.numId)
                        : null)
            .Select(numId => (string?)numId?.Attribute(W.val))
            .Where(value => int.TryParse(value, out _))
            .Select(value => int.Parse(
                value!, System.Globalization.CultureInfo.InvariantCulture))
            .Distinct()
            .ToArray();

    /// <summary>
    /// Remove Docxodus-authored numbering instances made unreachable by rejecting their native
    /// numPr insertion. Proof sessions exclude definitions present in the selected baseline.
    /// </summary>
    private void PruneUnreferencedDocxodusNumbering(IReadOnlyCollection<int> candidateIds)
    {
        var main = _doc!.MainDocumentPart;
        var part = main?.NumberingDefinitionsPart;
        var root = part?.GetXDocument().Root;
        if (main is null || part is null || root is null || candidateIds.Count == 0) return;

        var referenced = new HashSet<int>();
        foreach (var referencePart in EnumerateProjectedParts()
                     .Where(projected => !ReferenceEquals(projected, part)))
        {
            foreach (var numId in referencePart.GetXDocument().Descendants(W.numId))
                if (int.TryParse((string?)numId.Attribute(W.val), out var value) && value > 0)
                    referenced.Add(value);
        }

        var candidates = candidateIds.ToHashSet();
        bool changed = false;
        foreach (var num in root.Elements(W.num).ToList())
        {
            if (!int.TryParse((string?)num.Attribute(W.numId), out var numId)
                || !candidates.Contains(numId) || referenced.Contains(numId))
                continue;
            var abstractId = (string?)num.Element(W.abstractNumId)?.Attribute(W.val);
            var abstractNum = root.Elements(W.abstractNum).FirstOrDefault(candidate =>
                (string?)candidate.Attribute(W.abstractNumId) == abstractId);
            if (abstractNum is null
                || !_settings.ProofSafeRevisionResolution
                    && !Internal.NumberingFactory.IsDocxodusDefinition(abstractNum))
                continue;

            num.Remove();
            changed = true;
            bool protectedAbstract = int.TryParse(abstractId, out var parsedAbstractId)
                && _settings.ProtectedRevisionAbstractNumberingIds.Contains(parsedAbstractId);
            if (!protectedAbstract && !root.Elements(W.num).Any(other =>
                    (string?)other.Element(W.abstractNumId)?.Attribute(W.val) == abstractId))
                abstractNum.Remove();
        }

        if (!changed) return;
        if (!root.Elements().Any()) main.DeletePart(part);
        else part.PutXDocument();
    }

    /// <summary>Footnote/endnote ids cited anywhere in the package, captured before an op that
    /// can carry a citation away. See <see cref="Internal.NoteReferenceOps"/> for the rule.</summary>
    private (HashSet<int> Footnotes, HashSet<int> Endnotes) ReferencedNoteIds() =>
        Internal.NoteReferenceOps.ReferencedNoteIds(_doc?.MainDocumentPart);

    /// <summary>
    /// Word-faithful note cleanup after an op that removes body content (issues #516, #591): a note
    /// whose LAST citation this op carried away is deleted from its part. The rule and the reasons
    /// live in <see cref="Internal.NoteReferenceOps"/>, which the stateless
    /// <see cref="RevisionProcessor"/> path shares so a redline is reversible on every transport.
    /// </summary>
    /// <returns>The removed note elements with their part uri, for removed-anchor reporting.</returns>
    private List<(XElement Note, string PartUri)> PruneOrphanedNotes(
        (HashSet<int> Footnotes, HashSet<int> Endnotes) before) =>
        Internal.NoteReferenceOps.PruneOrphanedNotes(_doc?.MainDocumentPart, before);

    private static bool IsSeparatorNote(XElement note) =>
        Internal.NoteReferenceOps.IsSeparatorNote(note);

    /// <summary>Appends the anchors of pruned note definitions (with their descendants) to
    /// <paramref name="removed"/>, deduplicated through <paramref name="seenRemoved"/> — the
    /// shared removed-anchor reporting for every op that calls
    /// <see cref="PruneOrphanedNotes"/>.</summary>
    private void AppendPrunedNoteAnchors(
        List<(XElement Note, string PartUri)> prunedNotes,
        List<Anchor> removed,
        HashSet<string> seenRemoved)
    {
        foreach (var (note, notePartUri) in prunedNotes)
        {
            foreach (var descendant in note.DescendantsAndSelf())
            {
                var unid = (string?)descendant.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                if (AnchorForUnid(unid, notePartUri) is { } anchor && seenRemoved.Add(anchor.Id))
                    removed.Add(anchor);
            }
        }
    }

    internal IReadOnlyList<int> DefinedNumberingIds()
    {
        ThrowIfDisposed();
        return _doc?.MainDocumentPart?.NumberingDefinitionsPart?.GetXDocument().Root?
            .Elements(W.num)
            .Select(element => (string?)element.Attribute(W.numId))
            .Where(value => int.TryParse(value, out _))
            .Select(value => int.Parse(
                value!, System.Globalization.CultureInfo.InvariantCulture))
            .Distinct().OrderBy(value => value).ToArray()
            ?? Array.Empty<int>();
    }

    internal IReadOnlyList<int> DefinedAbstractNumberingIds()
    {
        ThrowIfDisposed();
        return _doc?.MainDocumentPart?.NumberingDefinitionsPart?.GetXDocument().Root?
            .Elements(W.abstractNum)
            .Select(element => (string?)element.Attribute(W.abstractNumId))
            .Where(value => int.TryParse(value, out _))
            .Select(value => int.Parse(
                value!, System.Globalization.CultureInfo.InvariantCulture))
            .Distinct().OrderBy(value => value).ToArray()
            ?? Array.Empty<int>();
    }

    internal IReadOnlyCollection<string> RevisionRelationshipKeys()
    {
        ThrowIfDisposed();
        return EnumeratePackageParts(_doc!)
            .SelectMany(owner => owner.Parts
                .Select(relationship => relationship.RelationshipId)
                .Concat(owner.HyperlinkRelationships.Select(relationship => relationship.Id))
                .Concat(owner.ExternalRelationships.Select(relationship => relationship.Id))
                .Concat(owner.DataPartReferenceRelationships.Select(relationship => relationship.Id))
                .Select(relationshipId => RevisionRelationshipKey(
                    owner.Uri.ToString(), relationshipId)))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    internal IReadOnlyCollection<string> EmptyRevisionPropertyContainerKeys()
    {
        ThrowIfDisposed();
        _ = AnchorIndex();
        return RevisionStoryParts()
            .SelectMany(story => story.Root.DescendantsAndSelf()
                .Where(element => Internal.RevisionOps.IsEmptyRemovablePropertyContainer(element))
                .SelectMany(element => Internal.RevisionOps.EmptyPropertyContainerKeys(
                    story.Part.Uri.ToString(), element)))
            .Distinct()
            .ToArray();
    }

    private static IReadOnlyCollection<string> RevisionRelationshipIds(
        Internal.RevisionOps.RevisionGroup group,
        OpenXmlPart? owner)
    {
        if (owner is null) return Array.Empty<string>();
        var relationshipIds = owner.Parts.Select(relationship => relationship.RelationshipId)
            .Concat(owner.HyperlinkRelationships.Select(relationship => relationship.Id))
            .Concat(owner.ExternalRelationships.Select(relationship => relationship.Id))
            .Concat(owner.DataPartReferenceRelationships.Select(relationship => relationship.Id))
            .ToHashSet(StringComparer.Ordinal);
        if (relationshipIds.Count == 0) return Array.Empty<string>();

        var roots = group.Units.SelectMany(unit => new[]
            {
                unit.Element,
                unit.MarkedCell,
                unit.MarkedRow,
                unit.StructuredWrapper,
                unit.Kind == Internal.RevisionOps.UnitKind.PropsChange
                    ? unit.Element.Parent : null,
                unit.Kind == Internal.RevisionOps.UnitKind.ParaMark
                    ? unit.Paragraph : null,
            })
            .Where(element => element is not null)
            .Select(element => element!)
            .Concat(group.Units.SelectMany(unit => unit.Element.Ancestors()
                .Where(ancestor => ancestor.Attributes().Any(attribute =>
                    !attribute.IsNamespaceDeclaration
                    && relationshipIds.Contains(attribute.Value)))))
            .Distinct();
        return roots.SelectMany(root => root.DescendantsAndSelf())
            .SelectMany(element => element.Attributes())
            .Where(attribute => !attribute.IsNamespaceDeclaration
                && (attribute.Name.Namespace == R.r
                    || attribute.Name.NamespaceName
                        == "http://purl.oclc.org/ooxml/officeDocument/relationships"
                    || (attribute.Name.NamespaceName
                            == "urn:schemas-microsoft-com:office:office"
                        && attribute.Name.LocalName == "relid")))
            .Select(attribute => attribute.Value)
            .Where(relationshipIds.Contains)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private IReadOnlyCollection<string> UnprotectedRevisionRelationshipIds(
        Internal.RevisionOps.RevisionGroup group,
        OpenXmlPart? owner)
    {
        if (owner is null) return Array.Empty<string>();
        var ownerUri = owner.Uri.ToString();
        return RevisionRelationshipIds(group, owner)
            .Where(relationshipId => !_settings.ProtectedRevisionRelationshipKeys.Contains(
                RevisionRelationshipKey(ownerUri, relationshipId), StringComparer.Ordinal))
            .ToArray();
    }

    private static string RevisionRelationshipKey(string ownerPartUri, string relationshipId) =>
        ownerPartUri + "\n" + relationshipId;

    private List<Anchor> RevisionGroupAnchors(
        Internal.RevisionOps.RevisionGroup group, string partUri)
    {
        long unlimitedTraversal = long.MaxValue;
        _ = TryRevisionGroupAnchors(
            group, partUri, int.MaxValue, ref unlimitedTraversal, out var anchors);
        return anchors;
    }

    private bool TryRevisionGroupAnchors(
        Internal.RevisionOps.RevisionGroup group,
        string partUri,
        int maximumAnchors,
        ref long remainingTraversal,
        out List<Anchor> anchors)
    {
        var collected = new List<Anchor>();
        anchors = collected;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var preferredByUnid = new Dictionary<string, Anchor>(StringComparer.Ordinal);
        var fallbackByUnid = new Dictionary<string, Anchor>(StringComparer.Ordinal);
        foreach (var target in AnchorIndex().Values)
        {
            fallbackByUnid.TryAdd(target.Unid, target.Anchor);
            if (target.PartUri == partUri)
                preferredByUnid.TryAdd(target.Unid, target.Anchor);
        }
        // Anchor is a struct: GetValueOrDefault would hand back an all-empty anchor for a
        // Unid no addressable element owns (a w:customXml wrapper, say), and the null test
        // below would then admit it.
        Anchor? FindAnchor(string unid) => preferredByUnid.TryGetValue(unid, out var preferred)
            ? preferred
            : fallbackByUnid.TryGetValue(unid, out var fallback) ? fallback : null;
        bool TryAdd(Anchor anchor)
        {
            if (!seen.Add(anchor.Id)) return true;
            if (collected.Count >= maximumAnchors) return false;
            collected.Add(anchor);
            return true;
        }

        var structuralTables = group.Units.Where(u => u.Kind == Internal.RevisionOps.UnitKind.CellMark)
            .Select(u => u.Table).Where(t => t is not null).Select(t => t!).Distinct().ToList();
        foreach (var table in structuralTables)
        {
            foreach (var element in table.DescendantsAndSelf())
            {
                if (!TryConsumeRevisionEvidenceTraversal(ref remainingTraversal)) return false;
                var unid = (string?)element.Attribute(PtOpenXml.Unid);
                if (unid is not null && FindAnchor(unid) is { } anchor
                    && !TryAdd(anchor))
                    return false;
            }
        }

        foreach (var unit in group.Units)
        {
            int anchorCountBeforeUnit = collected.Count;
            var start = unit.MarkedCell ?? unit.Paragraph ?? unit.MarkedRow
                ?? unit.StructuredWrapper ?? unit.Element;
            if (unit.StructuredWrapper is { } wrapper)
            {
                foreach (var element in wrapper.DescendantsAndSelf())
                {
                    if (!TryConsumeRevisionEvidenceTraversal(ref remainingTraversal)) return false;
                    var descendantUnid = (string?)element.Attribute(PtOpenXml.Unid);
                    if (descendantUnid is not null
                        && FindAnchor(descendantUnid) is { } descendantAnchor
                        && !TryAdd(descendantAnchor))
                        return false;
                }
            }
            for (var element = start;
                element is not null; element = element.Parent)
            {
                if (!TryConsumeRevisionEvidenceTraversal(ref remainingTraversal)) return false;
                var unid = (string?)element.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                if (FindAnchor(unid) is { } anchor)
                {
                    if (!TryAdd(anchor)) return false;
                    break;
                }
            }

            if (collected.Count == anchorCountBeforeUnit
                && unit.Element.Name == W.sectPrChange
                && unit.Element.Ancestors(W.sectPr).FirstOrDefault() is { } sectionProperties
                && sectionProperties.Parent?.Name == W.body)
            {
                var previousBlock = sectionProperties.ElementsBeforeSelf()
                    .LastOrDefault(element => element.Name == W.p || element.Name == W.tbl);
                if (previousBlock is not null)
                {
                    foreach (var element in previousBlock.DescendantsAndSelf())
                    {
                        if (!TryConsumeRevisionEvidenceTraversal(ref remainingTraversal)) return false;
                        var unid = (string?)element.Attribute(PtOpenXml.Unid);
                        if (unid is not null
                            && FindAnchor(unid) is { } anchor)
                        {
                            if (!TryAdd(anchor)) return false;
                            break;
                        }
                    }
                }
            }
        }
        return true;
    }

    private static bool TryConsumeRevisionEvidenceTraversal(ref long remaining)
    {
        if (remaining == long.MaxValue) return true;
        if (remaining == 0) return false;
        remaining--;
        return true;
    }

    private Internal.RevisionRegistry BuildRevisionRegistry()
    {
        var parts = RevisionStoryParts();
        return BuildRevisionRegistry(parts);
    }

    private static Internal.RevisionRegistry BuildRevisionRegistry(
        IReadOnlyList<(OpenXmlPart Part, XElement Root, string Scope)> parts)
    {
        var registry = Internal.RevisionRegistry.Build(parts.Select(p =>
            new Internal.RevisionRegistry.Part(
                p.Part.Uri.ToString(), p.Scope, p.Root)).ToList());
        foreach (var group in registry.Entries.Where(group => group.Scope is
                     "glossary" or "stylesWithEffects" or "glossaryStyles"
                     or "glossaryStylesWithEffects" or "glossaryNumbering"))
        {
            group.ResolutionStatus = RevisionResolutionStatus.Unsupported;
            group.Diagnostic = new RevisionDiagnostic(
                "unsupported_revision_part",
                "Revisions in this auxiliary style/glossary part are inventoried but selective resolution is unsupported.");
        }
        return registry;
    }

    /// <summary>The story parts revision markup lives in, in the fixed order the
    /// revision enumeration indexes them. Comments and styles are resolvable; glossary markup is
    /// inventoried explicitly and marked unsupported rather than disappearing from proof evidence.</summary>
    private List<(OpenXmlPart Part, XElement Root, string Scope)> RevisionStoryParts()
    {
        var list = new List<(OpenXmlPart, XElement, string)>();
        void Add(OpenXmlPart part, string scope)
        {
            var root = part.GetXDocument().Root;
            if (root is not null) list.Add((part, root, scope));
        }

        var main = _doc!.MainDocumentPart;
        if (main is null) return list;
        Add(main, "body");
        int index = 1;
        foreach (var header in main.HeaderParts) Add(header, $"hdr{index++}");
        index = 1;
        foreach (var footer in main.FooterParts) Add(footer, $"ftr{index++}");
        if (main.FootnotesPart is not null) Add(main.FootnotesPart, "fn");
        if (main.EndnotesPart is not null) Add(main.EndnotesPart, "en");
        if (main.WordprocessingCommentsPart is not null)
            Add(main.WordprocessingCommentsPart, "cmt");
        if (main.StyleDefinitionsPart is not null)
            Add(main.StyleDefinitionsPart, "styles");
        if (main.NumberingDefinitionsPart is not null)
            Add(main.NumberingDefinitionsPart, "numbering");
        if (main.StylesWithEffectsPart is not null)
            Add(main.StylesWithEffectsPart, "stylesWithEffects");
        if (main.GlossaryDocumentPart is not null)
        {
            Add(main.GlossaryDocumentPart, "glossary");
            if (main.GlossaryDocumentPart.StyleDefinitionsPart is not null)
                Add(main.GlossaryDocumentPart.StyleDefinitionsPart, "glossaryStyles");
            if (main.GlossaryDocumentPart.StylesWithEffectsPart is not null)
                Add(main.GlossaryDocumentPart.StylesWithEffectsPart,
                    "glossaryStylesWithEffects");
            if (main.GlossaryDocumentPart.NumberingDefinitionsPart is not null)
                Add(main.GlossaryDocumentPart.NumberingDefinitionsPart, "glossaryNumbering");
        }
        return list;
    }
}
