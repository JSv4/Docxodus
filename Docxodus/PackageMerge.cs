// Package-merge plumbing extracted verbatim from the removed WmlComparer (v11.0.0). These helpers
// copy styles, numbering definitions and related package parts BETWEEN two packages; none of them
// compare anything. DocxDiff's markup renderers have always called them, so they outlive the
// comparison engine they happened to live inside.

using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>
/// Cross-package merge helpers: import a document's missing styles and numbering definitions into
/// another package, and move the package parts a cloned element depends on along with it.
/// </summary>
internal static class PackageMerge
{
    /// <summary>
    /// Attribute names that carry a relationship id. A cloned element bearing any of these depends on a
    /// package part that must travel with it. Lifted from the removed comparison engine's
    /// <c>ComparisonUnitWord</c>, whose only role here was to own this list.
    /// </summary>
    private static readonly XName[] RelationshipAttributeNames =
    {
        R.embed, R.link, R.id, R.cs, R.dm, R.lo, R.qs, R.href, R.pict,
    };

        internal static void CopyMissingStylesFromOneDocToAnother(WordprocessingDocument wDocFrom, WordprocessingDocument wDocTo)
        {
            var revisionsStylesXDoc = wDocTo.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument();
            var afterStylesXDoc = wDocFrom.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument();
            foreach (var style in afterStylesXDoc.Root!.Elements(W.style))
            {
                var type = (string?)style.Attribute(W.type);
                var styleId = (string?)style.Attribute(W.styleId);
                var styleInRevDoc = revisionsStylesXDoc
                    .Root!
                    .Elements(W.style)
                    .FirstOrDefault(st => (string?)st.Attribute(W.type) == type &&
                                          (string?)st.Attribute(W.styleId) == styleId);
                if (styleInRevDoc != null)
                    continue;
                var cloned = new XElement(style);
                cloned.Attribute(W._default)?.Remove();
                revisionsStylesXDoc.Root!.Add(cloned);
            }
            wDocTo.MainDocumentPart!.StyleDefinitionsPart!.PutXDocument();
        }

        /// <summary>
        /// Copies numbering definitions from one document to another, handling ID conflicts.
        /// This ensures that when comparing documents with different numbering styles (e.g., legal numbering),
        /// the numbering definitions from the revised document are preserved in the comparison result.
        /// Fixes GitHub issue: https://github.com/dotnet/Open-XML-SDK/issues/1634
        /// </summary>
        // internal (was private): reused by Docxodus.Ir.Diff.IrMarkupRenderer for numbering continuity
        // (esp. legal-numbering preservation, GitHub #1634) when right-only content carries numbering the
        // LEFT package lacks.
        /// <summary>Copies numbering definitions missing from <paramref name="wDocTo"/> out of
        /// <paramref name="wDocFrom"/>. Returns the numId translation table for source definitions
        /// that had to be RENUMBERED around an id collision (source numId → destination numId);
        /// references to those ids in content cloned from the source document must be rebound by
        /// the caller — the ids they carry resolve to the destination's (different) definition.
        /// <para>When <paramref name="alignedNumIdPairs"/> is non-null the copy switches to the
        /// alignment-aware Word-parity rule (see <see cref="CopyNumberingPreservingListIdentity"/>):
        /// a USED imported source list instance gets its own FRESH cloned <c>w:abstractNum</c> unless
        /// the evidence proves it is the SAME list as a destination instance;
        /// <paramref name="usedFromNumIds"/> (null ⇒ treat every instance as used) says which source
        /// numIds surviving output content actually references — unreferenced definitions keep the
        /// legacy content-dedup treatment, so they are not needlessly duplicated. Null
        /// <paramref name="alignedNumIdPairs"/> keeps the legacy content-dedup behavior for the
        /// WmlComparer and Consolidate call sites.</para></summary>
        internal static Dictionary<int, int> CopyMissingNumberingFromOneDocToAnother(WordprocessingDocument wDocFrom, WordprocessingDocument wDocTo,
            IReadOnlyCollection<(int FromNumId, int ToNumId)>? alignedNumIdPairs = null,
            IReadOnlySet<int>? usedFromNumIds = null)
        {
            var numIdMap = new Dictionary<int, int>();
            var fromNumberingPart = wDocFrom.MainDocumentPart!.NumberingDefinitionsPart;
            if (fromNumberingPart == null)
                return numIdMap;

            var toNumberingPart = wDocTo.MainDocumentPart!.NumberingDefinitionsPart;
            XDocument toNumberingXDoc;

            if (toNumberingPart == null)
            {
                // Create a new NumberingDefinitionsPart if one doesn't exist
                toNumberingPart =
                    wDocTo.MainDocumentPart.AddDeterministicPart<NumberingDefinitionsPart>("rIdNumbering");
                toNumberingXDoc = new XDocument(
                    new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(W.numbering,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        new XAttribute(XNamespace.Xmlns + "r", R.r)));
                toNumberingPart.PutXDocument(toNumberingXDoc);
            }
            else
            {
                toNumberingXDoc = toNumberingPart.GetXDocument();
            }

            var fromNumberingXDoc = fromNumberingPart.GetXDocument();

            if (alignedNumIdPairs != null)
            {
                var identityNumIdMap = CopyNumberingPreservingListIdentity(
                    fromNumberingXDoc, toNumberingXDoc, alignedNumIdPairs, usedFromNumIds);
                toNumberingPart.PutXDocument(toNumberingXDoc);
                return identityNumIdMap;
            }

            // Find the maximum IDs in the destination document to avoid conflicts
            int maxAbstractNumId = toNumberingXDoc.Root!
                .Elements(W.abstractNum)
                .Select(e => (int?)e.Attribute(W.abstractNumId) ?? 0)
                .DefaultIfEmpty(0)
                .Max();

            int maxNumId = toNumberingXDoc.Root!
                .Elements(W.num)
                .Select(e => (int?)e.Attribute(W.numId) ?? 0)
                .DefaultIfEmpty(0)
                .Max();

            // Dictionary to track abstractNumId remapping (source ID -> destination ID)
            var abstractNumIdMap = new Dictionary<int, int>();

            // Copy abstractNum elements, reusing existing definitions with matching content
            foreach (var abstractNum in fromNumberingXDoc.Root!.Elements(W.abstractNum))
            {
                var fromAbstractNumId = GetIntAttribute(abstractNum, W.abstractNumId);
                if (fromAbstractNumId == null)
                    continue; // Skip malformed elements

                var normalizedFrom = NormalizeAbstractNumForComparison(abstractNum);

                // First, check if ANY existing abstractNum has matching content (regardless of ID)
                var matchingByContent = toNumberingXDoc.Root!
                    .Elements(W.abstractNum)
                    .FirstOrDefault(e => XNode.DeepEquals(NormalizeAbstractNumForComparison(e), normalizedFrom));

                if (matchingByContent != null)
                {
                    // Reuse existing abstractNum with matching content
                    var existingId = GetIntAttribute(matchingByContent, W.abstractNumId);
                    if (existingId != null)
                    {
                        abstractNumIdMap[fromAbstractNumId.Value] = existingId.Value;
                        continue;
                    }
                }

                // No matching content found - check if the ID is already taken
                var existingWithSameId = toNumberingXDoc.Root!
                    .Elements(W.abstractNum)
                    .FirstOrDefault(e => GetIntAttribute(e, W.abstractNumId) == fromAbstractNumId);

                int targetId;
                if (existingWithSameId != null)
                {
                    // ID conflict - assign a new ID
                    maxAbstractNumId++;
                    targetId = maxAbstractNumId;
                }
                else
                {
                    // ID is free, use it
                    targetId = fromAbstractNumId.Value;
                }

                var cloned = new XElement(abstractNum);
                cloned.SetAttributeValue(W.abstractNumId, targetId);
                abstractNumIdMap[fromAbstractNumId.Value] = targetId;

                WordprocessingMLUtil.InsertNumberingChildInOrder(toNumberingXDoc.Root!, cloned);
            }

            // Copy num elements that don't exist in destination
            foreach (var num in fromNumberingXDoc.Root!.Elements(W.num))
            {
                var fromNumId = GetIntAttribute(num, W.numId);
                var fromAbstractNumIdRef = GetIntAttribute(num.Element(W.abstractNumId), W.val);
                if (fromNumId == null || fromAbstractNumIdRef == null)
                    continue; // Skip malformed elements

                // Determine the mapped abstractNumId for this num
                int mappedAbstractNumId = abstractNumIdMap.TryGetValue(fromAbstractNumIdRef.Value, out var mapped)
                    ? mapped
                    : fromAbstractNumIdRef.Value;

                var existingNum = toNumberingXDoc.Root!
                    .Elements(W.num)
                    .FirstOrDefault(e => GetIntAttribute(e, W.numId) == fromNumId);

                if (existingNum != null)
                {
                    // Check if it references the same (mapped) abstractNum
                    var existingAbstractNumIdRef = GetIntAttribute(existingNum.Element(W.abstractNumId), W.val);
                    if (existingAbstractNumIdRef == mappedAbstractNumId)
                    {
                        // Same num with same abstractNum reference, skip
                        continue;
                    }

                    // Different abstractNum reference - need a new numId
                    maxNumId++;
                    var cloned = new XElement(num);
                    cloned.SetAttributeValue(W.numId, maxNumId);
                    var abstractNumIdElement = cloned.Element(W.abstractNumId);
                    if (abstractNumIdElement != null)
                        abstractNumIdElement.SetAttributeValue(W.val, mappedAbstractNumId);
                    WordprocessingMLUtil.InsertNumberingChildInOrder(toNumberingXDoc.Root!, cloned);
                    numIdMap[fromNumId.Value] = maxNumId;
                }
                else
                {
                    // No existing num with this ID, copy with remapped abstractNumId
                    var cloned = new XElement(num);
                    if (mappedAbstractNumId != fromAbstractNumIdRef.Value)
                    {
                        var abstractNumIdElement = cloned.Element(W.abstractNumId);
                        if (abstractNumIdElement != null)
                            abstractNumIdElement.SetAttributeValue(W.val, mappedAbstractNumId);
                    }
                    WordprocessingMLUtil.InsertNumberingChildInOrder(toNumberingXDoc.Root!, cloned);
                }
            }

            toNumberingPart.PutXDocument(toNumberingXDoc);
            return numIdMap;
        }

        /// <summary>
        /// Alignment-aware numbering import (the IR diff renderer's variant of
        /// <see cref="CopyMissingNumberingFromOneDocToAnother"/>). Word's compare output gives every
        /// imported foreign list instance its own FRESH cloned <c>w:abstractNum</c>; it never
        /// deduplicates an imported definition onto a content-equal destination abstractNum. That
        /// matters because LibreOffice keys list COUNTERS by abstractNumId — mapping two different
        /// list instances (one destination-native, one imported) onto one shared abstractNum makes it
        /// CONTINUE numbering across them where Word's output RESTARTS each list. The one exception
        /// is a source list that is genuinely the SAME list as a surviving destination list:
        /// <paramref name="alignedNumIdPairs"/> carries (source numId, destination numId) pairs
        /// harvested from paragraph pairs the diff aligned as present on BOTH sides; when such a
        /// pair exists and the two definitions agree, the destination definition is reused so the
        /// counter continues across an inserted item joining that list (recording a numId rebind
        /// when the two ids differ). Imported instances preserve the SOURCE document's own
        /// num→abstractNum topology (source nums sharing one abstract share its clone), so the
        /// accepted output renders like the source document.
        /// <para>The fresh-clone rule applies only to instances surviving output content actually
        /// REFERENCES (<paramref name="usedFromNumIds"/>; null ⇒ all). An unreferenced source
        /// definition cannot exhibit the counter defect, so it keeps the legacy content-dedup
        /// treatment — forking it would only duplicate definitions (and any schema noise they
        /// carry) to no rendering effect.</para>
        /// </summary>
        private static Dictionary<int, int> CopyNumberingPreservingListIdentity(
            XDocument fromNumberingXDoc, XDocument toNumberingXDoc,
            IReadOnlyCollection<(int FromNumId, int ToNumId)> alignedNumIdPairs,
            IReadOnlySet<int>? usedFromNumIds)
        {
            var numIdMap = new Dictionary<int, int>();

            // Seed the id high-water marks with BOTH sides' maxima so a freshly-allocated id can
            // never collide with a source id processed later in the loop.
            int maxAbstractNumId = toNumberingXDoc.Root!
                .Elements(W.abstractNum)
                .Concat(fromNumberingXDoc.Root!.Elements(W.abstractNum))
                .Select(e => (int?)e.Attribute(W.abstractNumId) ?? 0)
                .DefaultIfEmpty(0)
                .Max();

            int maxNumId = toNumberingXDoc.Root!
                .Elements(W.num)
                .Concat(fromNumberingXDoc.Root!.Elements(W.num))
                .Select(e => (int?)e.Attribute(W.numId) ?? 0)
                .DefaultIfEmpty(0)
                .Max();

            static XElement? FindAbstract(XDocument numberingXDoc, int id) => numberingXDoc.Root!
                .Elements(W.abstractNum)
                .FirstOrDefault(e => GetIntAttribute(e, W.abstractNumId) == id);

            static XElement? FindNum(XDocument numberingXDoc, int id) => numberingXDoc.Root!
                .Elements(W.num)
                .FirstOrDefault(e => GetIntAttribute(e, W.numId) == id);

            // One clone per SOURCE abstractNum, not per instance: source nums that shared an
            // abstract keep sharing its clone, preserving the source's own counter topology.
            // Forked (used) and deduped (unused) resolutions are cached separately — a used
            // instance must never land on a content-matched destination abstract.
            var forkedAbstractIds = new Dictionary<int, int>();
            var dedupedAbstractIds = new Dictionary<int, int>();

            int ResolveTargetAbstractId(XElement fromAbstract, int fromAbstractId, bool allowContentDedup)
            {
                var cache = allowContentDedup ? dedupedAbstractIds : forkedAbstractIds;
                if (cache.TryGetValue(fromAbstractId, out var cached))
                    return cached;

                if (allowContentDedup)
                {
                    var normalized = NormalizeAbstractNumForComparison(fromAbstract);
                    var matchingByContent = toNumberingXDoc.Root!
                        .Elements(W.abstractNum)
                        .FirstOrDefault(e => XNode.DeepEquals(NormalizeAbstractNumForComparison(e), normalized));
                    var matchingId = matchingByContent == null
                        ? null
                        : GetIntAttribute(matchingByContent, W.abstractNumId);
                    if (matchingId != null)
                    {
                        cache[fromAbstractId] = matchingId.Value;
                        return matchingId.Value;
                    }
                }

                var targetId = FindAbstract(toNumberingXDoc, fromAbstractId) == null
                    ? fromAbstractId
                    : ++maxAbstractNumId;
                var clonedAbstract = new XElement(fromAbstract);
                clonedAbstract.SetAttributeValue(W.abstractNumId, targetId);
                WordprocessingMLUtil.InsertNumberingChildInOrder(toNumberingXDoc.Root!, clonedAbstract);
                cache[fromAbstractId] = targetId;
                return targetId;
            }

            foreach (var num in fromNumberingXDoc.Root!.Elements(W.num).ToList())
            {
                var fromNumId = GetIntAttribute(num, W.numId);
                var fromAbstractRef = GetIntAttribute(num.Element(W.abstractNumId), W.val);
                if (fromNumId == null || fromAbstractRef == null)
                    continue; // Skip malformed elements

                var fromAbstract = FindAbstract(fromNumberingXDoc, fromAbstractRef.Value);
                if (fromAbstract == null)
                    continue; // Dangling source reference — nothing to import (the caller's repair pass owns refs)

                var normalizedFrom = NormalizeAbstractNumForComparison(fromAbstract);

                // Identity test: this source instance is "the same list" as a destination instance
                // when the diff aligned at least one surviving paragraph pair carrying both numIds
                // AND the definitions agree. Prefer the same-id pairing (no rebind needed), then
                // the smallest destination id, for determinism.
                int? sameListNumId = null;
                foreach (var candidateId in alignedNumIdPairs
                             .Where(p => p.FromNumId == fromNumId.Value)
                             .Select(p => p.ToNumId)
                             .Distinct()
                             .OrderBy(id => id == fromNumId.Value ? 0 : 1)
                             .ThenBy(id => id))
                {
                    var candidateNum = FindNum(toNumberingXDoc, candidateId);
                    if (candidateNum == null)
                        continue;
                    var candidateAbstractRef = GetIntAttribute(candidateNum.Element(W.abstractNumId), W.val);
                    if (candidateAbstractRef == null)
                        continue;
                    var candidateAbstract = FindAbstract(toNumberingXDoc, candidateAbstractRef.Value);
                    if (candidateAbstract != null &&
                        XNode.DeepEquals(NormalizeAbstractNumForComparison(candidateAbstract), normalizedFrom))
                    {
                        sameListNumId = candidateId;
                        break;
                    }
                }

                if (sameListNumId != null)
                {
                    // Same list: keep the destination definition so the counter continues.
                    if (sameListNumId.Value != fromNumId.Value)
                        numIdMap[fromNumId.Value] = sameListNumId.Value;
                    continue;
                }

                // A USED foreign instance imports under its own cloned abstractNum — NEVER onto a
                // content-equal destination abstract. An UNREFERENCED definition keeps the legacy
                // content-dedup import (it cannot render, so forking it is pure duplication).
                var isUsed = usedFromNumIds == null || usedFromNumIds.Contains(fromNumId.Value);
                var targetAbstractId = ResolveTargetAbstractId(
                    fromAbstract, fromAbstractRef.Value, allowContentDedup: !isUsed);

                int targetNumId;
                var existingNum = FindNum(toNumberingXDoc, fromNumId.Value);
                if (existingNum == null)
                {
                    targetNumId = fromNumId.Value;
                }
                else
                {
                    if (GetIntAttribute(existingNum.Element(W.abstractNumId), W.val) == targetAbstractId)
                        continue; // The destination already carries this exact instance.
                    targetNumId = ++maxNumId;
                    numIdMap[fromNumId.Value] = targetNumId;
                }

                var clonedNum = new XElement(num);
                clonedNum.SetAttributeValue(W.numId, targetNumId);
                var abstractNumIdElement = clonedNum.Element(W.abstractNumId);
                if (abstractNumIdElement != null)
                    abstractNumIdElement.SetAttributeValue(W.val, targetAbstractId);
                WordprocessingMLUtil.InsertNumberingChildInOrder(toNumberingXDoc.Root!, clonedNum);
            }

            return numIdMap;
        }

        /// <summary>
        /// Safely extracts an integer value from an XAttribute.
        /// </summary>
        private static int? GetIntAttribute(XElement? element, XName attributeName)
        {
            if (element == null)
                return null;
            var attr = element.Attribute(attributeName);
            if (attr == null)
                return null;
            if (int.TryParse(attr.Value, out var result))
                return result;
            return null;
        }

        /// <summary>
        /// Normalizes an abstractNum element for comparison by removing ID-based attributes
        /// that don't affect the functional behavior of the numbering definition.
        /// </summary>
        private static XElement NormalizeAbstractNumForComparison(XElement abstractNum)
        {
            var normalized = new XElement(abstractNum);
            // Remove attributes/elements that shouldn't affect functional comparison
            normalized.Attribute(W.abstractNumId)?.Remove();
            // nsid is a unique identifier that may differ between documents
            normalized.Element(W.nsid)?.Remove();
            // tmpl is auto-generated and may differ
            normalized.Element(W.tmpl)?.Remove();
            return normalized;
        }

        internal static XElement MoveRelatedPartsToDestination(PackagePart partOfDeletedContent, PackagePart partInNewDocument,
            XElement contentElement, bool skipDanglingRelationships = false, bool skipHeaderFooterReferences = false)
        {
            var state = new RelatedPartImportState(partOfDeletedContent, partInNewDocument);
            return MoveRelatedPartsToDestination(
                partOfDeletedContent, partInNewDocument, contentElement, state,
                skipDanglingRelationships, skipHeaderFooterReferences);
        }

        /// <summary>
        /// Per-root import state. Reusing one copied destination part for a repeated source target preserves
        /// relationship graphs (including a corrupt cyclic graph) without unbounded recursive cloning.
        /// </summary>
        private sealed class RelatedPartImportState
        {
            // Most package parts are globally identified by their source URI, so copying the same target once
            // correctly preserves sharing and cycles. DiagramDataPart is the exception: its
            // dsp:dataModelExt/@relId is resolved against the *owner that linked the data part*. A shared data
            // part can therefore carry distinct owner-local prebuilt drawing edges. Keep a separate cloned data
            // part for each such source owner instead of letting the first import rewrite the shared copy for all
            // subsequent owners.
            private readonly Dictionary<RelatedPartKey, PackagePart> _destinationsBySourceKey = new();

            private readonly record struct RelatedPartKey(
                Uri SourcePartUri,
                Uri? DiagramDataRelationshipOwnerUri);

            public RelatedPartImportState(PackagePart sourceRoot, PackagePart destinationRoot)
            {
                _destinationsBySourceKey.Add(new RelatedPartKey(sourceRoot.Uri, null), destinationRoot);
            }

            public bool TryGetDestination(
                PackagePart sourcePart, PackagePart sourceOwner, PackageRelationship sourceRelationship,
                [NotNullWhen(true)] out PackagePart? destinationPart) =>
                _destinationsBySourceKey.TryGetValue(
                    GetKey(sourcePart, sourceOwner, sourceRelationship), out destinationPart);

            public void Remember(
                PackagePart sourcePart, PackagePart sourceOwner, PackageRelationship sourceRelationship,
                PackagePart destinationPart) =>
                _destinationsBySourceKey.Add(
                    GetKey(sourcePart, sourceOwner, sourceRelationship), destinationPart);

            private static RelatedPartKey GetKey(
                PackagePart sourcePart, PackagePart sourceOwner, PackageRelationship sourceRelationship) =>
                new(
                    sourcePart.Uri,
                    sourceRelationship.RelationshipType.EndsWith("/diagramData", StringComparison.Ordinal)
                        ? sourceOwner.Uri
                        : null);
        }

        private static XElement MoveRelatedPartsToDestination(
            PackagePart partOfDeletedContent, PackagePart partInNewDocument, XElement contentElement,
            RelatedPartImportState state, bool skipDanglingRelationships, bool skipHeaderFooterReferences)
        {
            var elementsToUpdate = contentElement
                // Recursive graph import can receive an XML-part ROOT (not just a cloned paragraph/drawing).
                // Include that root so a relationship attribute located directly on it is copied/remapped too.
                .DescendantsAndSelf()
                .Where(d => d.Attributes().Any(a => RelationshipAttributeNames.Contains(a.Name)))
                // c:externalData is INCLUDED: skipping it (a leftover from the legacy DetachExternalData
                // flow, which the IR path never runs) left a copied chart part without its own rels —
                // the embedded workbook never came across and the chart's rId dangled, so LibreOffice
                // rendered the chart as a black box. A genuinely dangling externalData rel is skipped
                // (not thrown) below, preserving the legacy callers' tolerance.
                // The IR renderer clones whole blocks (paragraphs), which can carry a w:sectPr with
                // w:headerReference/w:footerReference. Header/footer scopes are NOT diffed, so the LEFT
                // package's parts (already present, same r:ids — both sides derive from one base) are
                // authoritative; importing the RIGHT's would duplicate them as P<guid> parts. The legacy
                // WmlComparer callers only ever pass a w:drawing, so they never opt in (default false).
                .Where(d => !skipHeaderFooterReferences
                            || (d.Name != W.headerReference && d.Name != W.footerReference))
                .ToList();
            foreach (var element in elementsToUpdate)
            {
                var attributesToUpdate = element
                    .Attributes()
                    .Where(a => RelationshipAttributeNames.Contains(a.Name))
                    .ToList();
                foreach (var att in attributesToUpdate)
                {
                    var rId = (string)att;

                    // A DANGLING rId — one that names no relationship at all in this source part — is the trigger.
                    // (Hyperlink/external rels are recreated by the caller BEFORE this runs, so a surviving
                    // unresolvable rId here is genuinely dangling, NOT an external-hyperlink rel.) The old engine
                    // treats that as a corrupt document and fails loudly; only the IR caller, which tolerates a
                    // dangling reference on unchanged-text content, opts into skipping it.
                    if (!partOfDeletedContent.RelationshipExists(rId))
                    {
                        if (skipDanglingRelationships || element.Name == C.externalData)
                            continue;
                        throw new FileFormatException(
                            $"Content references relationship id '{rId}' that does not exist in the source part.");
                    }

                    var relationshipForDeletedPart = partOfDeletedContent.GetRelationship(rId);
                    if (relationshipForDeletedPart == null)
                        throw new FileFormatException("Invalid document");

                    // External and hyperlink relationships have no package part. The old path attempted
                    // Package.GetPart on them, which made a nested chart c:externalData/@r:id detectable
                    // by the IR but impossible to render. Recreate them on the copied XML owner and remap.
                    if (relationshipForDeletedPart.TargetMode == TargetMode.External)
                    {
                        att.Value = ImportExternalRelationship(partInNewDocument, relationshipForDeletedPart);
                        continue;
                    }

                    if (!TryGetInternalTargetPart(partOfDeletedContent, relationshipForDeletedPart, out var relatedPackagePart))
                    {
                        if (skipDanglingRelationships || element.Name == C.externalData)
                            continue;
                        throw new FileFormatException(
                            $"Content relationship '{rId}' does not resolve to a package part.");
                    }

                    att.Value = ImportInternalRelationship(
                        partOfDeletedContent, partInNewDocument, relatedPackagePart, relationshipForDeletedPart,
                        state, skipDanglingRelationships, skipHeaderFooterReferences);
                }
            }
            return contentElement;
        }

        private static string ImportExternalRelationship(
            PackagePart destinationOwner, PackageRelationship sourceRelationship)
        {
            var newRid = NewRelationshipId(destinationOwner);
            destinationOwner.CreateRelationship(
                sourceRelationship.TargetUri, TargetMode.External, sourceRelationship.RelationshipType, newRid);
            return newRid;
        }

        private static string ImportInternalRelationship(
            PackagePart sourceOwner, PackagePart destinationOwner, PackagePart sourceTarget,
            PackageRelationship sourceRelationship, RelatedPartImportState state,
            bool skipDanglingRelationships, bool skipHeaderFooterReferences)
        {
            if (!state.TryGetDestination(sourceTarget, sourceOwner, sourceRelationship, out var destinationTarget))
            {
                destinationTarget = destinationOwner.Package.CreatePart(
                    NewRelatedPartUri(sourceTarget, destinationOwner.Package), sourceTarget.ContentType);
                state.Remember(sourceTarget, sourceOwner, sourceRelationship, destinationTarget);
                using (var oldPartStream = sourceTarget.GetStream())
                using (var newPartStream = destinationTarget.GetStream())
                    FileUtils.CopyStream(oldPartStream, newPartStream);

                FixupCopiedXmlPart(
                    sourceOwner, destinationOwner, sourceTarget, destinationTarget, sourceRelationship, state,
                    skipDanglingRelationships, skipHeaderFooterReferences);
            }

            var newRid = NewRelationshipId(destinationOwner);
            destinationOwner.CreateRelationship(
                destinationTarget.Uri, TargetMode.Internal, sourceRelationship.RelationshipType, newRid);
            return newRid;
        }

        private static void FixupCopiedXmlPart(
            PackagePart sourceOwner, PackagePart destinationOwner, PackagePart sourcePart, PackagePart copiedPart,
            PackageRelationship sourceRelationship, RelatedPartImportState state,
            bool skipDanglingRelationships, bool skipHeaderFooterReferences)
        {
            if (!copiedPart.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase))
                return;

            XDocument copiedXDoc;
            try
            {
                using var stream = copiedPart.GetStream();
                copiedXDoc = XDocument.Load(stream);
            }
            catch (Exception e) when (e is System.Xml.XmlException or ArgumentException)
            {
                // A readable malformed XML part receives a raw-byte identity in IrDrawingGraphHasher. It was
                // already copied intact; skipping recursive fixup keeps Compare total and lets Accept/Reject
                // retain its exact source bytes instead of turning a detected change into an import failure.
                return;
            }

            if (copiedXDoc.Root is null)
                return;

            MoveRelatedPartsToDestination(
                sourcePart, copiedPart, copiedXDoc.Root, state,
                skipDanglingRelationships, skipHeaderFooterReferences);
            // FileMode.Create TRUNCATES. The plain GetStream() overload opens OpenOrCreate, so a rewrite
            // shorter than the bytes just copied in leaves the tail of the original behind and the part stops
            // parsing. Remapping only ever GREW a part while relationship ids were "R" + a 32-character Guid;
            // with the deterministic ids (see NewRelationshipId) a fixed-up part is routinely shorter.
            using (var stream = copiedPart.GetStream(FileMode.Create, FileAccess.Write))
                copiedXDoc.Save(stream);

            // SmartArt's PREBUILT drawing rides an extension inside the DATA part (dsp:dataModelExt/@relId)
            // whose relationship lives on the owner that linked the data part, not the data part itself.
            if (sourceRelationship.RelationshipType.EndsWith("/diagramData", StringComparison.Ordinal))
            {
                ChaseDiagramDataModelExt(
                    sourceOwner, destinationOwner, copiedPart, state,
                    skipDanglingRelationships, skipHeaderFooterReferences);
            }
        }

        private static bool TryGetInternalTargetPart(
            PackagePart sourceOwner, PackageRelationship relationship, [NotNullWhen(true)] out PackagePart? targetPart)
        {
            try
            {
                var targetUri = PackUriHelper.ResolvePartUri(
                    new Uri(sourceOwner.Uri.ToString(), UriKind.RelativeOrAbsolute),
                    new Uri(relationship.TargetUri.ToString(), UriKind.RelativeOrAbsolute));
                if (sourceOwner.Package.PartExists(targetUri))
                {
                    targetPart = sourceOwner.Package.GetPart(targetUri);
                    return true;
                }
            }
            catch (ArgumentException)
            {
                // Treat a malformed internal target like the other dangling relationship forms above.
            }

            targetPart = null;
            return false;
        }

        /// <summary>
        /// Deterministic destination name for an imported part, keeping the historic "P" + 32 lowercase hex
        /// shape. A fresh <c>Guid</c> here used to churn the name of every imported media/diagram part on
        /// every run, and that churn propagated into <c>document.xml</c>, the <c>_rels</c> and the
        /// <c>[Content_Types].xml</c> overrides — so a redline that imported media was never byte-reproducible,
        /// contrary to what <see cref="Ir.Diff.IrDiffSettings.Deterministic"/> promises.
        ///
        /// The address is taken over the SOURCE part's bytes, not the copy's: an XML part is fixed up AFTER
        /// it is copied, and that fixup embeds the names of the parts below it, so the copy's own bytes are
        /// not yet known when the name has to be chosen. Identical source bytes therefore land on one name,
        /// and the lowest free "-N" suffix separates copies that must stay distinct anyway — a second import
        /// of the same source part (each cloned block imports under its own <see cref="RelatedPartImportState"/>)
        /// or the per-owner diagram-data rule that state documents. The suffix is stable because the import
        /// walk is document-ordered.
        /// </summary>
        private static Uri NewRelatedPartUri(PackagePart sourcePart, Package destinationPackage)
        {
            var uriSplit = sourcePart.Uri.ToString().Split('/');
            var last = uriSplit[uriSplit.Length - 1].Split('.');
            var stem = uriSplit.PtSkipLast(1).Select(p => p + "/").StringConcatenate() +
                "P" + ContentAddress(sourcePart);
            var extension = last.Length == 2 ? "." + last[1] : string.Empty;
            var kind = sourcePart.Uri.IsAbsoluteUri ? UriKind.Absolute : UriKind.Relative;

            var candidate = new Uri(stem + extension, kind);
            for (var n = 1; destinationPackage.PartExists(candidate); n++)
                candidate = new Uri(
                    stem + "-" + n.ToString(CultureInfo.InvariantCulture) + extension, kind);
            return candidate;
        }

        /// <summary>Lowercase hex SHA-256 of a part's bytes, truncated to the 32 characters the
        /// replaced <c>Guid</c> occupied. Collision only costs a "-N" suffix, never correctness.</summary>
        private static string ContentAddress(PackagePart part)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return Convert.ToHexStringLower(SHA256.HashData(stream)).Substring(0, 32);
        }

        /// <summary>
        /// Deterministic relationship id: the lowest free "R"+ordinal on the owner that will carry the
        /// relationship. Relationship ids are part-scoped, so clearing the ids already on this one owner is
        /// enough for validity, and the result is reproducible where the replaced <c>Guid</c> was not.
        /// Same probe shape as <c>IrMarkupRenderer.FreshRelationshipId</c>.
        /// </summary>
        private static string NewRelationshipId(PackagePart destinationOwner)
        {
            var n = 1;
            string candidate;
            do
            {
                candidate = "R" + n++.ToString(CultureInfo.InvariantCulture);
            }
            while (destinationOwner.RelationshipExists(candidate));
            return candidate;
        }

        /// <summary>See the call site: copy the MS-2007 prebuilt diagram drawing referenced by the
        /// copied data part's <c>dsp:dataModelExt/@relId</c> (resolved against the top-level source
        /// part) and rewire the attribute. A dangling or absent extension is skipped gracefully.</summary>
        private static void ChaseDiagramDataModelExt(
            PackagePart topLevelSourcePart, PackagePart partInNewDocument, PackagePart copiedDataPart,
            RelatedPartImportState state, bool skipDanglingRelationships, bool skipHeaderFooterReferences)
        {
            XNamespace dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";
            XDocument dataXDoc;
            try
            {
                using var s = copiedDataPart.GetStream();
                dataXDoc = XDocument.Load(s);
            }
            catch (Exception e) when (e is System.Xml.XmlException or ArgumentException)
            {
                return;
            }

            bool changed = false;
            foreach (var ext in dataXDoc.Descendants(dsp + "dataModelExt").ToList())
            {
                var relId = (string?)ext.Attribute("relId");
                if (string.IsNullOrEmpty(relId) || !topLevelSourcePart.RelationshipExists(relId))
                    continue;
                var rel = topLevelSourcePart.GetRelationship(relId);
                if (rel is null)
                    continue;

                if (rel.TargetMode == TargetMode.External)
                {
                    ext.SetAttributeValue("relId", ImportExternalRelationship(partInNewDocument, rel));
                    changed = true;
                    continue;
                }

                if (!TryGetInternalTargetPart(topLevelSourcePart, rel, out var srcPart))
                    continue;
                ext.SetAttributeValue(
                    "relId",
                    ImportInternalRelationship(
                        topLevelSourcePart, partInNewDocument, srcPart, rel, state,
                        skipDanglingRelationships, skipHeaderFooterReferences));
                changed = true;
            }

            if (changed)
                using (var s = copiedDataPart.GetStream(FileMode.Create, FileAccess.Write))
                    dataXDoc.Save(s);
        }

}
