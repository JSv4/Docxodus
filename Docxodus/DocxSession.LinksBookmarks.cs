// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

public sealed partial class DocxSession
{
    private static readonly XNamespace LinkR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly Regex BookmarkNamePattern = new("^[A-Za-z_][A-Za-z0-9_]{0,39}$", RegexOptions.CultureInvariant);

    /// <summary>
    /// Names Word allocates and rewrites for itself: <c>_GoBack</c> (last edit position), <c>_Toc*</c>
    /// (regenerated wholesale every time a TOC field refreshes), <c>_Ref*</c> (allocated when a
    /// cross-reference is inserted), and <c>_Hlt*</c>/<c>_Hlk*</c> (hyperlink bookkeeping).
    /// <para>Policy: this namespace is closed to <em>creation</em> — <see cref="AddBookmark"/> and the
    /// destination name of <see cref="RenameBookmark"/> refuse it, because a name Word owns will be
    /// silently reallocated or clobbered under the caller's feet. Bookmarks that Word already put there
    /// stay fully readable and mutable: renaming, moving, or removing one is a legitimate edit, and its
    /// inbound <c>REF</c>/<c>PAGEREF</c>/<c>NOTEREF</c>/<c>HYPERLINK \l</c> fields are retargeted (rename)
    /// or block the removal (remove) like any other. Note that renaming a <c>_Toc*</c> bookmark is not
    /// durable — the next TOC refresh in Word regenerates the whole family.</para>
    /// </summary>
    private static readonly Regex ReservedBookmarkNamePattern = new(
        @"^_(GoBack|Toc\d*|Ref\d*|Hlt\d*|Hlk\d*)$",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    /// <summary>Enumerate native hyperlinks in body, headers, footers, footnotes, and endnotes.</summary>
    public IReadOnlyList<HyperlinkInfo> ListHyperlinks(ProjectionScopes scopes = ProjectionScopes.All)
    {
        ThrowIfDisposed();
        _ = AnchorIndex();
        var result = new List<HyperlinkInfo>();
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
        {
            if (!scopes.IncludesScope(owner.Scope)) continue;
            var root = owner.Part.GetXDocument().Root;
            if (root is null) continue;
            foreach (var link in root.Descendants(W.hyperlink))
            {
                var paragraph = link.Ancestors(W.p).FirstOrDefault();
                var anchor = paragraph is null ? null : AnchorForElement(paragraph);
                if (paragraph is null || anchor is null) continue;
                var map = RunTextMap.Build(paragraph);
                var member = map.Segments.Where(s => s.Run.Ancestors(W.hyperlink).FirstOrDefault() == link).ToList();
                int start = member.Count == 0 ? MarkerOffset(paragraph, link) : member[0].StartOffsetInBlock;
                int end = member.Count == 0 ? start : member[^1].EndOffsetInBlock;
                var internalTarget = (string?)link.Attribute(W.anchor);
                var relationshipId = (string?)link.Attribute(LinkR + "id");
                HyperlinkKind kind;
                string? target;
                bool? external = null;
                bool broken;
                if (!string.IsNullOrEmpty(internalTarget))
                {
                    kind = HyperlinkKind.Internal;
                    target = internalTarget;
                    broken = ResolveBookmarkPair(internalTarget, hyperlinkTarget: true,
                        anchor.Value.Id, out _, out _) is not null;
                }
                else
                {
                    kind = HyperlinkKind.External;
                    var relationship = owner.Part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId);
                    target = relationship?.Uri.ToString();
                    external = relationship?.IsExternal;
                    broken = relationship is null;
                }
                result.Add(new HyperlinkInfo(
                    HyperlinkPublicId(owner, link), kind, owner.PartUri, owner.Scope,
                    anchor.Value.Id, new CharSpan(start, Math.Max(0, end - start)),
                    TextForRuns(member.Select(s => s.Run)), target, relationshipId, external, broken));
            }
        }
        return result;
    }

    /// <summary>Enumerate bookmark pairs, including pairs whose endpoints cross paragraphs in
    /// one owning part. Pairing is keyed by (part, w:id); unmatched starts and ambiguous
    /// ids/names are diagnostic. Orphan ends have no name/start coordinate and are not rows.</summary>
    public IReadOnlyList<BookmarkInfo> ListBookmarks(ProjectionScopes scopes = ProjectionScopes.All)
    {
        ThrowIfDisposed();
        _ = AnchorIndex();
        var owners = OwnedPartRelationships.StoryParts(_doc!).ToList();
        var nodes = owners.SelectMany((owner, ownerIndex) =>
            (owner.Part.GetXDocument().Root?.DescendantsAndSelf() ?? Enumerable.Empty<XElement>())
                .Select((element, nodeIndex) => (owner, ownerIndex, element, nodeIndex))).ToList();
        var ends = nodes.Where(n => n.element.Name == W.bookmarkEnd).ToList();
        var result = new List<BookmarkInfo>();

        foreach (var startNode in nodes.Where(n => n.element.Name == W.bookmarkStart))
        {
            if (!scopes.IncludesScope(startNode.owner.Scope)) continue;
            var name = (string?)startNode.element.Attribute(W.name) ?? string.Empty;
            var id = (string?)startNode.element.Attribute(W.id) ?? string.Empty;
            // w:id is story-part scoped in real files. Never pair a body start with a header/footer
            // end merely because Word reused the same decimal id in both parts.
            var candidateEnds = ends.Where(e => e.owner.PartUri == startNode.owner.PartUri
                && string.Equals((string?)e.element.Attribute(W.id), id, StringComparison.Ordinal)
                && XNode.DocumentOrderComparer.Compare(e.element, startNode.element) > 0).ToList();
            var endNode = candidateEnds.FirstOrDefault();
            int sameIdStarts = nodes.Count(n => n.element.Name == W.bookmarkStart
                && n.owner.PartUri == startNode.owner.PartUri
                && string.Equals((string?)n.element.Attribute(W.id), id, StringComparison.Ordinal));
            int sameIdEnds = ends.Count(n => n.owner.PartUri == startNode.owner.PartUri
                && string.Equals((string?)n.element.Attribute(W.id), id, StringComparison.Ordinal));
            int sameNameStarts = nodes.Count(n => n.element.Name == W.bookmarkStart
                && string.Equals((string?)n.element.Attribute(W.name), name, StringComparison.Ordinal));
            var startParagraph = startNode.element.Ancestors(W.p).FirstOrDefault();
            var endParagraph = endNode.element?.Ancestors(W.p).FirstOrDefault();
            var startAnchor = startParagraph is null ? null : AnchorForElement(startParagraph);
            var endAnchor = endParagraph is null ? null : AnchorForElement(endParagraph);
            DocumentRange? range = null;
            IReadOnlyList<BookmarkRangeSegment> segments = Array.Empty<BookmarkRangeSegment>();
            string text = string.Empty;
            if (startParagraph is not null && endParagraph is not null && startAnchor is not null && endAnchor is not null)
            {
                int startOffset = MarkerOffset(startParagraph, startNode.element);
                int endOffset = MarkerOffset(endParagraph, endNode.element!);
                range = new DocumentRange(startAnchor.Value.Id, startOffset, endAnchor.Value.Id, endOffset);
                segments = BuildBookmarkSegments(owners, startNode.owner, startParagraph, startOffset,
                    endNode.owner, endParagraph, endOffset);
                text = string.Join("\n", segments.Select(s => s.Text));
            }
            string? validationError = candidateEnds.Count == 0 ? "bookmarkEnd is missing"
                : candidateEnds.Count > 1 ? "multiple bookmarkEnd markers follow this start"
                : sameIdStarts > 1 ? "bookmark numeric id is duplicated in this story part"
                : sameIdEnds > 1 ? "bookmark numeric id has multiple ends in this story part"
                : sameNameStarts > 1 ? "bookmark name is duplicated"
                : null;
            result.Add(new BookmarkInfo(name, id, startNode.owner.PartUri, startNode.owner.Scope,
                endNode.element is null ? null : endNode.owner.PartUri,
                endNode.element is null ? null : endNode.owner.Scope,
                range, segments, text, endNode.element is not null,
                name.StartsWith(AnnotationManager.BookmarkPrefix, StringComparison.Ordinal),
                validationError is null, validationError));
        }
        return result;
    }

    public EditResult AddHyperlink(string anchorId, CharSpan span, HyperlinkTarget target)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
                "hyperlink mutations cannot be represented faithfully as tracked revisions", anchorId);
        var validatedTarget = ValidateHyperlinkTarget(target, anchorId);
        if (validatedTarget is not null) return validatedTarget;
        var anchor = FindAnchor(anchorId);
        if (anchor is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var paragraph = anchor.Resolve(_doc!);
        if (paragraph is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        if (paragraph.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "AddHyperlink requires a paragraph/heading/list-item anchor", anchorId);
        var map = RunTextMap.Build(paragraph);
        if (span.Length <= 0) return EditResult.Fail(EditErrorCode.EmptyHyperlinkSpan, "hyperlink span must contain text", anchorId);
        if (span.Start < 0 || span.Start + span.Length > map.FlatText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span [{span.Start},{span.Start + span.Length}) outside paragraph of length {map.FlatText.Length}", anchorId);
        var range = RunTextMap.ResolveRange(map, span.Start, span.Length);
        var unsupported = ValidateHyperlinkBoundary(paragraph, range.Select(x => x.Segment.Run));
        if (unsupported is not null) return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary, unsupported, anchorId);
        var owner = OwnedPartRelationships.FindOwner(_doc!, paragraph);
        if (owner is null) return EditResult.Fail(EditErrorCode.InternalError, "paragraph has no owning story part", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            SplitRunsAtOffset(paragraph, span.Start + span.Length);
            SplitRunsAtOffset(paragraph, span.Start);
            var splitMap = RunTextMap.Build(paragraph);
            var selected = splitMap.Segments
                .Where(s => s.StartOffsetInBlock >= span.Start && s.EndOffsetInBlock <= span.Start + span.Length)
                .Select(s => s.Run).ToList();
            if (selected.Count == 0 || selected.Any(r => r.Parent != paragraph))
                throw new InvalidOperationException("selection did not resolve to direct paragraph runs");
            var link = new XElement(W.hyperlink,
                new XAttribute(PtOpenXml.Unid, UnidHelper.GenerateUnid()));
            ApplyHyperlinkTarget(owner.Value.Part, link, target);
            // Relocate the whole CONTIGUOUS SIBLING RANGE, not just the selected w:r elements. A
            // zero-width marker sitting between two selected runs (w:bookmarkStart/End,
            // w:commentRangeStart/End, w:proofErr) is a legal w:hyperlink child (EG_PContent covers
            // EG_RunLevelElts), and moving only the runs would strand it AFTER the finished
            // w:hyperlink — which for a bookmark whose start lies inside the span puts the start
            // after its own end and permanently unresolvable. Everything outside the range keeps its
            // side of the link because the link is inserted where the first selected run stood, so
            // document order is preserved for every marker, inside the span or not.
            var siblings = paragraph.Elements().ToList();
            int firstIndex = siblings.FindIndex(e => ReferenceEquals(e, selected[0]));
            int lastIndex = siblings.FindIndex(e => ReferenceEquals(e, selected[^1]));
            if (firstIndex < 0 || lastIndex < firstIndex)
                throw new InvalidOperationException("selection did not resolve to a contiguous sibling range");
            var relocated = siblings.GetRange(firstIndex, lastIndex - firstIndex + 1);
            selected[0].AddBeforeSelf(link);
            foreach (var child in relocated) { child.Remove(); link.Add(child); }
            var id = HyperlinkPublicId(owner.Value, link);
            InvalidateProjectionCache();
            return new EditResult { Success = true, HyperlinkId = id, Modified = new[] { anchor.Anchor } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult UpdateHyperlink(string hyperlinkId, HyperlinkTarget target)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
                "hyperlink mutations cannot be represented faithfully as tracked revisions");
        var validatedTarget = ValidateHyperlinkTarget(target, null);
        if (validatedTarget is not null) return validatedTarget;
        var found = FindHyperlinkElement(hyperlinkId);
        if (found is null) return EditResult.Fail(EditErrorCode.HyperlinkNotFound, $"hyperlink not found: {hyperlinkId}");
        var (owner, link) = found.Value;
        var oldRelationshipId = (string?)link.Attribute(LinkR + "id");
        var paragraph = link.Ancestors(W.p).FirstOrDefault();
        var anchor = paragraph is null ? null : AnchorForElement(paragraph);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            ApplyHyperlinkTarget(owner.Part, link, target);
            OwnedPartRelationships.DeleteReferenceRelationshipIfOrphaned(owner.Part, oldRelationshipId, LinkR + "id");
            InvalidateProjectionCache();
            return new EditResult { Success = true, HyperlinkId = hyperlinkId,
                Modified = anchor is null ? Array.Empty<Anchor>() : new[] { anchor.Value } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    public EditResult RemoveHyperlink(string hyperlinkId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
                "hyperlink mutations cannot be represented faithfully as tracked revisions");
        var found = FindHyperlinkElement(hyperlinkId);
        if (found is null) return EditResult.Fail(EditErrorCode.HyperlinkNotFound, $"hyperlink not found: {hyperlinkId}");
        var (owner, link) = found.Value;
        var oldRelationshipId = (string?)link.Attribute(LinkR + "id");
        var paragraph = link.Ancestors(W.p).FirstOrDefault();
        var anchor = paragraph is null ? null : AnchorForElement(paragraph);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            link.ReplaceWith(link.Nodes());
            OwnedPartRelationships.DeleteReferenceRelationshipIfOrphaned(owner.Part, oldRelationshipId, LinkR + "id");
            InvalidateProjectionCache();
            return new EditResult { Success = true, HyperlinkId = hyperlinkId,
                Modified = anchor is null ? Array.Empty<Anchor>() : new[] { anchor.Value } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    public EditResult AddBookmark(string name, DocumentRange range)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var common = ValidateBookmarkMutation(name);
        if (common is not null) return common;
        if (BookmarkStarts(name).Count != 0)
            return EditResult.Fail(EditErrorCode.DuplicateBookmarkName, $"bookmark name already exists: {name}");
        var endpoints = ValidateDocumentRange(range);
        if (endpoints.Error is not null) return endpoints.Error;
        var bookmarkId = NextGlobalBookmarkId();
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            InsertRangeMarkers(endpoints, bookmarkId, name);
            InvalidateProjectionCache();
            return new EditResult { Success = true, BookmarkName = name,
                Modified = EndpointAnchors(endpoints) };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    public EditResult RenameBookmark(string name, string newName)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var oldValidation = ValidateBookmarkMutation(name, requireValidName: false);
        if (oldValidation is not null) return oldValidation;
        var newValidation = ValidateBookmarkMutation(newName);
        if (newValidation is not null) return newValidation;
        if (ResolveBookmarkPair(name, hyperlinkTarget: false, null,
            out var start, out _) is { } pairError) return pairError;
        if (!string.Equals(name, newName, StringComparison.Ordinal) && BookmarkStarts(newName).Count != 0)
            return EditResult.Fail(EditErrorCode.DuplicateBookmarkName, $"bookmark name already exists: {newName}");
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            start.SetAttributeValue(W.name, newName);
            // Retarget BOTH consumer families atomically: w:anchor links and REF/PAGEREF/NOTEREF/
            // HYPERLINK \l field instructions. Missing the fields is what turns every TOC entry over a
            // renamed _Toc bookmark into "Error! Bookmark not defined." the next time Word repaints.
            foreach (var reference in BookmarkReferences(name))
                RetargetBookmarkReference(reference, newName);
            InvalidateProjectionCache();
            return new EditResult { Success = true, BookmarkName = newName };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    public EditResult MoveBookmark(string name, DocumentRange range)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var common = ValidateBookmarkMutation(name, requireValidName: false);
        if (common is not null) return common;
        if (ResolveBookmarkPair(name, hyperlinkTarget: false, null,
            out var start, out var end) is { } pairError) return pairError;
        var id = (string?)start.Attribute(W.id);
        var endpoints = ValidateDocumentRange(range);
        if (endpoints.Error is not null) return endpoints.Error;
        // w:id is story-part scoped and Word reuses the same decimal across parts, so carrying the
        // source id into a DIFFERENT part can collide with a bookmark already living there — and since
        // ResolveBookmarkPair demands exactly one start and one end per id, the collision makes BOTH
        // bookmarks unresolvable. Only a same-part move keeps the id (LB004 pins that); a cross-part
        // move takes a fresh document-global one, allocated before the old markers are unlinked so it
        // can never be the id just freed.
        var sourceOwner = OwnedPartRelationships.FindOwner(_doc!, start);
        var destinationOwner = OwnedPartRelationships.FindOwner(_doc!, endpoints.StartParagraph);
        if (sourceOwner is null || destinationOwner is null || !string.Equals(
                sourceOwner.Value.PartUri, destinationOwner.Value.PartUri, StringComparison.Ordinal))
            id = NextGlobalBookmarkId();
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            start.Remove();
            end.Remove();
            InsertRangeMarkers(endpoints, id!, name);
            InvalidateProjectionCache();
            return new EditResult { Success = true, BookmarkName = name,
                Modified = EndpointAnchors(endpoints) };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    public EditResult RemoveBookmark(string name)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var common = ValidateBookmarkMutation(name, requireValidName: false);
        if (common is not null) return common;
        if (ResolveBookmarkPair(name, hyperlinkTarget: false, null,
            out var start, out var end) is { } pairError) return pairError;
        if (BookmarkReferences(name).Count != 0)
            return EditResult.Fail(EditErrorCode.BookmarkInUse,
                $"bookmark is targeted by one or more internal hyperlinks or cross-reference fields: {name}");
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            start.Remove();
            end.Remove();
            InvalidateProjectionCache();
            return new EditResult { Success = true, BookmarkName = name };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    private EditResult? ValidateHyperlinkTarget(HyperlinkTarget? target, string? anchorId)
    {
        if (target is null || string.IsNullOrWhiteSpace(target.Target))
            return EditResult.Fail(EditErrorCode.InvalidHyperlinkTarget, "hyperlink target is empty", anchorId);
        if (target.Kind is not (HyperlinkKind.Internal or HyperlinkKind.External))
            return EditResult.Fail(EditErrorCode.InvalidHyperlinkTarget,
                $"unknown hyperlink target kind: {target.Kind}", anchorId);
        if (target.Kind == HyperlinkKind.Internal)
        {
            if (ResolveBookmarkPair(target.Target, hyperlinkTarget: true, anchorId,
                out _, out _) is { } pairError) return pairError;
        }
        else if (!Uri.TryCreate(target.Target, UriKind.RelativeOrAbsolute, out _)
            || target.Target.StartsWith("#", StringComparison.Ordinal))
            return EditResult.Fail(EditErrorCode.InvalidHyperlinkTarget,
                $"invalid external hyperlink target: {target.Target}", anchorId);
        return null;
    }

    /// <summary>Validate Markdown parser's detached href markers before a mutation snapshots or
    /// changes XML. This gives <c>[text](#bookmark)</c> the same structured target rules as the
    /// first-class API.</summary>
    private EditResult? ValidatePendingHyperlinks(IEnumerable<XElement> elements, string? anchorId)
    {
        foreach (var link in elements.SelectMany(e => e.DescendantsAndSelf(W.hyperlink)))
        {
            var href = (string?)link.Attribute(MarkdownPayloadParser.HrefAttr);
            if (href is null) continue;
            var target = href.StartsWith("#", StringComparison.Ordinal)
                ? HyperlinkTarget.Internal(href.Substring(1))
                : HyperlinkTarget.External(href);
            if (ValidateHyperlinkTarget(target, anchorId) is { } error) return error;
        }
        return null;
    }

    private EditResult? ValidateBookmarkMutation(string name, bool requireValidName = true)
    {
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
                "bookmark mutations cannot be represented faithfully as tracked revisions");
        if (name.StartsWith(AnnotationManager.BookmarkPrefix, StringComparison.Ordinal))
            return EditResult.Fail(EditErrorCode.ManagedBookmark,
                $"bookmark is managed by the annotation subsystem: {name}");
        if (requireValidName && !BookmarkNamePattern.IsMatch(name))
            return EditResult.Fail(EditErrorCode.InvalidBookmarkName,
                "bookmark names must be 1-40 characters, start with a letter or underscore, and contain only letters, digits, or underscores");
        // requireValidName is exactly the creation side (AddBookmark's name, RenameBookmark's newName),
        // which is where the reserved-namespace policy on ReservedBookmarkNamePattern applies.
        if (requireValidName && ReservedBookmarkNamePattern.IsMatch(name))
            return EditResult.Fail(EditErrorCode.InvalidBookmarkName,
                "bookmark name is reserved for Word's own TOC/cross-reference/hyperlink bookkeeping "
                + $"and would be reallocated or clobbered by Word: {name}");
        return null;
    }

    private static string? ValidateHyperlinkBoundary(XElement paragraph, IEnumerable<XElement> selectedRuns)
    {
        if (paragraph.Descendants().Any(e => e.Name == W.ins || e.Name == W.del
            || e.Name == W.moveFrom || e.Name == W.moveTo))
            return "hyperlinks cannot be created across tracked-revision markup";
        if (paragraph.Descendants().Any(e => e.Name == W.fldChar || e.Name == W.instrText))
            return "hyperlinks cannot be created across a complex field boundary";
        foreach (var run in selectedRuns)
        {
            var containers = run.Ancestors().TakeWhile(a => a != paragraph).ToList();
            if (containers.Count != 0)
                return containers.Any(a => a.Name == W.hyperlink)
                    ? "hyperlink spans cannot overlap an existing hyperlink"
                    : "hyperlink span crosses an unsupported inline container";
        }
        return null;
    }

    private void ApplyHyperlinkTarget(OpenXmlPart owner, XElement link, HyperlinkTarget target)
    {
        link.Attribute(W.anchor)?.Remove();
        link.Attribute(LinkR + "id")?.Remove();
        if (target.Kind == HyperlinkKind.Internal)
            link.SetAttributeValue(W.anchor, target.Target);
        else
        {
            var relationship = OwnedPartRelationships.FindOrAddExternalHyperlink(
                owner, new Uri(target.Target, UriKind.RelativeOrAbsolute));
            link.SetAttributeValue(LinkR + "id", relationship.Id);
        }
    }

    private (OwnedPartRelationships.Owner Owner, XElement Link)? FindHyperlinkElement(string hyperlinkId)
    {
        _ = AnchorIndex();
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
            foreach (var link in owner.Part.GetXDocument().Descendants(W.hyperlink))
                if (string.Equals(HyperlinkPublicId(owner, link), hyperlinkId, StringComparison.Ordinal))
                    return (owner, link);
        return null;
    }

    private static string HyperlinkPublicId(OwnedPartRelationships.Owner owner, XElement link) =>
        $"hl:{owner.Scope}:{UnidHelper.ReadOrDeriveUnid(link)}";

    private List<XElement> BookmarkStarts(string name) => OwnedPartRelationships.StoryParts(_doc!)
        .SelectMany(o => o.Part.GetXDocument().Descendants(W.bookmarkStart))
        .Where(b => string.Equals((string?)b.Attribute(W.name), name, StringComparison.Ordinal)).ToList();

    // ─── Inbound bookmark references ────────────────────────────────────────────────────────────
    //
    // A bookmark has TWO kinds of consumer, and a rename or removal that sees only the first one
    // silently produces "Error! Bookmark not defined." in Word:
    //   1. w:hyperlink/@w:anchor            — the relationship-free internal link this API authors.
    //   2. field cross-references           — REF / PAGEREF / NOTEREF / HYPERLINK \l, carried either
    //      in w:fldSimple/@w:instr or in the w:instrText runs between a w:fldChar begin and its
    //      matching separate/end. Every Word TOC entry is a PAGEREF over a _Toc bookmark.
    // Both are enumerated by BookmarkReferences so rename retargets them and removal is blocked by
    // them, with one scan and one definition of "referenced".

    /// <summary>One field instruction plus the XML that stores it: either a <c>w:fldSimple</c>
    /// (instruction in <c>@w:instr</c>) or the ordered <c>w:instrText</c> elements of one
    /// <c>w:fldChar</c>-delimited field.</summary>
    private sealed record FieldInstruction(XElement? Simple, IReadOnlyList<XElement> InstrTexts, string Text);

    /// <summary>A whitespace- or quote-delimited field-instruction token and where it sits in the
    /// concatenated instruction, so a reference can be spliced without disturbing switches.</summary>
    private readonly record struct FieldInstrToken(string Value, int Start, int Length);

    /// <summary>One inbound reference to a bookmark name. <see cref="Element"/> is the XML a
    /// structural deletion would have to contain for the reference to disappear with it (the
    /// <c>w:hyperlink</c>, the <c>w:fldSimple</c>, or the field's first <c>w:instrText</c>).</summary>
    private sealed record BookmarkReference(XElement Element, FieldInstruction? Field, FieldInstrToken Token);

    /// <summary>Every field instruction in one story part, in document order. Nested fields are
    /// tracked with a stack so an inner <c>{ PAGE }</c> never swallows its host's instruction.</summary>
    private static List<FieldInstruction> FieldInstructionsIn(XElement root)
    {
        var result = new List<FieldInstruction>();
        foreach (var simple in root.DescendantsAndSelf(W.fldSimple))
            if ((string?)simple.Attribute(W.instr) is { } instr)
                result.Add(new FieldInstruction(simple, Array.Empty<XElement>(), instr));

        var open = new Stack<(List<XElement> Parts, bool InInstruction)>();
        foreach (var element in root.DescendantsAndSelf()
            .Where(e => e.Name == W.fldChar || e.Name == W.instrText))
        {
            if (element.Name == W.instrText)
            {
                if (open.Count > 0 && open.Peek().InInstruction) open.Peek().Parts.Add(element);
                continue;
            }
            switch ((string?)element.Attribute(W.fldCharType))
            {
                case "begin":
                    open.Push((new List<XElement>(), true));
                    break;
                case "separate" when open.Count > 0:
                    // The instruction is complete at the separator; the field stays open for its result.
                    var separated = open.Pop();
                    Emit(separated.Parts);
                    open.Push((separated.Parts, false));
                    break;
                case "end" when open.Count > 0:
                    var ended = open.Pop();
                    if (ended.InInstruction) Emit(ended.Parts);
                    break;
            }
        }
        return result;

        void Emit(List<XElement> parts)
        {
            if (parts.Count > 0)
                result.Add(new FieldInstruction(null, parts, string.Concat(parts.Select(p => p.Value))));
        }
    }

    private static List<FieldInstrToken> TokenizeInstruction(string text)
    {
        var tokens = new List<FieldInstrToken>();
        int i = 0;
        while (i < text.Length)
        {
            while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
            if (i >= text.Length) break;
            int start = i;
            if (text[i] == '"')
            {
                i++;
                var value = new StringBuilder();
                while (i < text.Length && text[i] != '"')
                {
                    if (text[i] == '\\' && i + 1 < text.Length) i++;
                    value.Append(text[i]);
                    i++;
                }
                if (i < text.Length) i++;
                tokens.Add(new FieldInstrToken(value.ToString(), start, i - start));
            }
            else
            {
                while (i < text.Length && !char.IsWhiteSpace(text[i])) i++;
                tokens.Add(new FieldInstrToken(text.Substring(start, i - start), start, i - start));
            }
        }
        return tokens;
    }

    /// <summary>The token naming a bookmark in a field instruction, or null when the field does not
    /// cross-reference one. <c>REF</c>/<c>PAGEREF</c>/<c>NOTEREF</c> name it as their first argument;
    /// <c>HYPERLINK</c> names it as the argument of the <c>\l</c> switch.</summary>
    private static FieldInstrToken? BookmarkReferenceToken(IReadOnlyList<FieldInstrToken> tokens)
    {
        if (tokens.Count < 2) return null;
        bool IsSwitch(FieldInstrToken t) => t.Value.StartsWith("\\", StringComparison.Ordinal);
        switch (tokens[0].Value.ToUpperInvariant())
        {
            case "REF":
            case "PAGEREF":
            case "NOTEREF":
                foreach (var token in tokens.Skip(1))
                    if (!IsSwitch(token)) return token;
                return null;
            case "HYPERLINK":
                for (int i = 1; i < tokens.Count - 1; i++)
                    if (string.Equals(tokens[i].Value, "\\l", StringComparison.OrdinalIgnoreCase))
                        return tokens[i + 1];
                return null;
            default:
                return null;
        }
    }

    /// <summary>Every inbound reference to <paramref name="name"/> across every story part.</summary>
    private List<BookmarkReference> BookmarkReferences(string name)
    {
        var result = new List<BookmarkReference>();
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
        {
            var root = owner.Part.GetXDocument().Root;
            if (root is null) continue;
            foreach (var link in root.Descendants(W.hyperlink))
                if (string.Equals((string?)link.Attribute(W.anchor), name, StringComparison.Ordinal))
                    result.Add(new BookmarkReference(link, null, default));
            foreach (var field in FieldInstructionsIn(root))
                if (BookmarkReferenceToken(TokenizeInstruction(field.Text)) is { } token
                    && string.Equals(token.Value, name, StringComparison.Ordinal))
                    result.Add(new BookmarkReference(field.Simple ?? field.InstrTexts[0], field, token));
        }
        return result;
    }

    private static void RetargetBookmarkReference(BookmarkReference reference, string newName)
    {
        if (reference.Field is not { } field)
        {
            reference.Element.SetAttributeValue(W.anchor, newName);
            return;
        }
        // Splice only the name token so switches (\h, \* MERGEFORMAT, …) survive verbatim, and keep
        // the quoting style the field already used.
        bool quote = field.Text[reference.Token.Start] == '"'
            || TokenizeInstruction(newName).Count != 1;
        string replacement = quote ? "\"" + newName.Replace("\"", "\\\"") + "\"" : newName;
        string updated = field.Text.Remove(reference.Token.Start, reference.Token.Length)
            .Insert(reference.Token.Start, replacement);
        if (field.Simple is not null)
        {
            field.Simple.SetAttributeValue(W.instr, updated);
            return;
        }
        // A split instruction is coalesced onto its first w:instrText: the run count, their order and
        // the surrounding fldChar plumbing are untouched, so the field stays a well-formed field, and
        // xml:space="preserve" keeps the leading/trailing spaces that separate the switches.
        for (int i = 0; i < field.InstrTexts.Count; i++)
        {
            field.InstrTexts[i].SetAttributeValue(XNamespace.Xml + "space", "preserve");
            field.InstrTexts[i].Value = i == 0 ? updated : string.Empty;
        }
    }

    private List<XElement> BookmarkEndsForStart(XElement start, string? id)
    {
        var owner = OwnedPartRelationships.FindOwner(_doc!, start);
        if (owner is null) return new List<XElement>();
        return owner.Value.Part.GetXDocument().Descendants(W.bookmarkEnd)
            .Where(b => string.Equals((string?)b.Attribute(W.id), id, StringComparison.Ordinal)
                && XNode.DocumentOrderComparer.Compare(b, start) > 0).ToList();
    }

    /// <summary>Resolve one globally named bookmark to one unambiguous, ordered start/end pair in
    /// a single story part. A name-only start is not a valid internal-link target or mutable
    /// bookmark: accepting it would preserve or create dangling Word markup.</summary>
    private EditResult? ResolveBookmarkPair(string name, bool hyperlinkTarget, string? anchorId,
        out XElement start, out XElement end)
    {
        start = null!;
        end = null!;
        var starts = BookmarkStarts(name);
        if (starts.Count == 0)
            return EditResult.Fail(hyperlinkTarget
                    ? EditErrorCode.MissingBookmarkTarget : EditErrorCode.BookmarkNotFound,
                $"bookmark {(hyperlinkTarget ? "target does not exist" : "not found")}: {name}", anchorId);
        if (starts.Count > 1)
            return EditResult.Fail(EditErrorCode.DuplicateBookmarkName,
                $"bookmark name is ambiguous: {name}", anchorId);

        start = starts[0];
        var owner = OwnedPartRelationships.FindOwner(_doc!, start);
        var id = (string?)start.Attribute(W.id);
        if (owner is not null && !string.IsNullOrEmpty(id))
        {
            var storyStarts = owner.Value.Part.GetXDocument().Descendants(W.bookmarkStart)
                .Where(marker => string.Equals((string?)marker.Attribute(W.id), id,
                    StringComparison.Ordinal)).ToList();
            var storyEnds = owner.Value.Part.GetXDocument().Descendants(W.bookmarkEnd)
                .Where(marker => string.Equals((string?)marker.Attribute(W.id), id,
                    StringComparison.Ordinal)).ToList();
            if (storyStarts.Count == 1 && storyEnds.Count == 1
                && XNode.DocumentOrderComparer.Compare(start, storyEnds[0]) < 0)
            {
                end = storyEnds[0];
                return null;
            }
        }

        return EditResult.Fail(hyperlinkTarget
                ? EditErrorCode.MissingBookmarkTarget : EditErrorCode.BookmarkNotFound,
            $"bookmark is not one coherent same-story start/end pair: {name}", anchorId);
    }

    /// <summary>
    /// Guard generic structural deletions from leaving half a bookmark pair or a dangling
    /// internal hyperlink. A complete unreferenced pair may be deleted with its containing
    /// content; ranges crossing the deletion boundary are rejected before the undo snapshot.
    /// </summary>
    private EditResult? ValidateBookmarkRemoval(IEnumerable<XElement> removalRoots, string anchorId)
    {
        var roots = removalRoots.Distinct().ToList();
        bool IsRemoved(XElement element) => roots.Any(root =>
            ReferenceEquals(root, element) || element.Ancestors().Any(a => ReferenceEquals(a, root)));

        var markers = roots.SelectMany(root => root.DescendantsAndSelf()
            .Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
            .Distinct().ToList();
        if (markers.Count == 0) return null;

        foreach (var start in markers.Where(e => e.Name == W.bookmarkStart))
        {
            var name = (string?)start.Attribute(W.name);
            var id = (string?)start.Attribute(W.id);
            var ends = BookmarkEndsForStart(start, id);
            if (name is null || ends.Count != 1 || !IsRemoved(ends[0]))
                return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                    "structural deletion would leave a bookmark range endpoint orphaned", anchorId);
            if (name.StartsWith(AnnotationManager.BookmarkPrefix, StringComparison.Ordinal))
                return EditResult.Fail(EditErrorCode.ManagedBookmark,
                    $"structural deletion includes an annotation-managed bookmark: {name}", anchorId);
            if (BookmarkReferences(name).Any(reference => !IsRemoved(reference.Element)))
                return EditResult.Fail(EditErrorCode.BookmarkInUse,
                    "structural deletion would remove a bookmark still targeted by an internal hyperlink "
                    + $"or cross-reference field: {name}", anchorId);
        }

        foreach (var end in markers.Where(e => e.Name == W.bookmarkEnd))
        {
            var owner = OwnedPartRelationships.FindOwner(_doc!, end);
            var id = (string?)end.Attribute(W.id);
            if (owner is null) return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                "structural deletion contains an ownerless bookmarkEnd", anchorId);
            var starts = owner.Value.Part.GetXDocument().Descendants(W.bookmarkStart)
                .Where(start =>
                {
                    var paired = BookmarkEndsForStart(start, id);
                    return string.Equals((string?)start.Attribute(W.id), id, StringComparison.Ordinal)
                        && paired.Count == 1 && ReferenceEquals(paired[0], end);
                })
                .ToList();
            if (starts.Count != 1 || !IsRemoved(starts[0]))
                return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                    "structural deletion would leave a bookmark range endpoint orphaned", anchorId);
        }
        return null;
    }

    private string NextGlobalBookmarkId()
    {
        int max = -1;
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
            foreach (var marker in owner.Part.GetXDocument().Descendants()
                .Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
                if (int.TryParse((string?)marker.Attribute(W.id), out var id)) max = Math.Max(max, id);
        return checked(max + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private sealed record ValidatedRange(
        XElement StartParagraph, Anchor StartAnchor, int StartOffset,
        XElement EndParagraph, Anchor EndAnchor, int EndOffset,
        EditResult? Error = null);

    private ValidatedRange ValidateDocumentRange(DocumentRange range)
    {
        var startTarget = FindAnchor(range.StartAnchorId);
        if (startTarget is null) return RangeError(EditErrorCode.AnchorNotFound, "start anchor not found", range.StartAnchorId);
        var endTarget = FindAnchor(range.EndAnchorId);
        if (endTarget is null) return RangeError(EditErrorCode.AnchorNotFound, "end anchor not found", range.EndAnchorId);
        var start = startTarget.Resolve(_doc!);
        var end = endTarget.Resolve(_doc!);
        if (start is null || end is null) return RangeError(EditErrorCode.AnchorNotFound, "range endpoint resolved null", range.StartAnchorId);
        if (start.Name != W.p || end.Name != W.p)
            return RangeError(EditErrorCode.AnchorWrongKind, "bookmark endpoints must be paragraphs/headings/list items", range.StartAnchorId);
        int startLength = RunTextMap.Build(start).FlatText.Length;
        int endLength = RunTextMap.Build(end).FlatText.Length;
        if (range.StartOffset < 0 || range.StartOffset > startLength || range.EndOffset < 0 || range.EndOffset > endLength)
            return RangeError(EditErrorCode.OffsetOutOfRange, "bookmark endpoint offset is outside its paragraph", range.StartAnchorId);
        if (ReferenceEquals(start, end) && range.EndOffset < range.StartOffset)
            return RangeError(EditErrorCode.OffsetOutOfRange, "bookmark end precedes its start", range.StartAnchorId);
        var startOwner = OwnedPartRelationships.FindOwner(_doc!, start);
        var endOwner = OwnedPartRelationships.FindOwner(_doc!, end);
        if (startOwner is null || endOwner is null
            || !string.Equals(startOwner.Value.PartUri, endOwner.Value.PartUri, StringComparison.Ordinal))
            return RangeError(EditErrorCode.UnsupportedInlineBoundary,
                "bookmark pairs cannot cross XML package parts; choose endpoints in the same body/header/footer/note story",
                range.StartAnchorId);
        if (HasUnsafeBookmarkBoundary(start, range.StartOffset) || HasUnsafeBookmarkBoundary(end, range.EndOffset))
            return RangeError(EditErrorCode.UnsupportedInlineBoundary,
                "bookmark endpoint falls inside a field, revision, or unsupported inline container", range.StartAnchorId);
        var ordered = OwnedPartRelationships.StoryParts(_doc!).SelectMany(o =>
            o.Part.GetXDocument().Descendants(W.p)).ToList();
        if (ordered.IndexOf(start) > ordered.IndexOf(end))
            return RangeError(EditErrorCode.OffsetOutOfRange, "bookmark end precedes its start in document order", range.StartAnchorId);
        return new ValidatedRange(start, startTarget.Anchor, range.StartOffset, end, endTarget.Anchor, range.EndOffset);
    }

    private static ValidatedRange RangeError(EditErrorCode code, string message, string anchorId) =>
        new(new XElement(W.p), new Anchor("", "", "", ""), 0,
            new XElement(W.p), new Anchor("", "", "", ""), 0, EditResult.Fail(code, message, anchorId));

    private static bool HasUnsafeBookmarkBoundary(XElement paragraph, int offset)
    {
        if (paragraph.Descendants().Any(e => e.Name == W.ins || e.Name == W.del
            || e.Name == W.moveFrom || e.Name == W.moveTo || e.Name == W.fldChar || e.Name == W.instrText))
            return true;
        int consumed = 0;
        foreach (var child in paragraph.Elements().Where(IsInlineChild))
        {
            int length = InlineChildTextLength(child);
            if (child.Name != W.r && child.Name != W.hyperlink
                && consumed < offset && offset < consumed + length)
                return true;
            consumed += length;
        }
        return false;
    }

    private static void InsertRangeMarkers(ValidatedRange range, string bookmarkId, string name)
    {
        var start = new XElement(W.bookmarkStart,
            new XAttribute(W.id, bookmarkId), new XAttribute(W.name, name));
        var end = new XElement(W.bookmarkEnd, new XAttribute(W.id, bookmarkId));
        if (ReferenceEquals(range.StartParagraph, range.EndParagraph)
            && range.StartOffset == range.EndOffset)
        {
            InsertCollapsedBookmarkAtOffset(range.StartParagraph, range.StartOffset, start, end);
            return;
        }
        // End first keeps its pre-split offset stable when both endpoints share a paragraph.
        InsertMarkerAtOffset(range.EndParagraph, range.EndOffset,
            end);
        InsertMarkerAtOffset(range.StartParagraph, range.StartOffset,
            start);
    }

    private static void InsertCollapsedBookmarkAtOffset(
        XElement paragraph, int offset, XElement start, XElement end)
    {
        InsertMarkersAtOffset(paragraph, offset, new[] { start, end });
    }

    private static void InsertMarkerAtOffset(XElement paragraph, int offset, XElement marker)
        => InsertMarkersAtOffset(paragraph, offset, new[] { marker });

    private static void InsertMarkersAtOffset(
        XElement paragraph, int offset, IReadOnlyList<XElement> markers)
    {
        SplitRunsAtOffset(paragraph, offset);
        SplitInlineContainersAtOffset(paragraph, offset);
        var map = RunTextMap.Build(paragraph);
        var right = map.Segments.FirstOrDefault(s => s.StartOffsetInBlock >= offset).Run;
        if (right is not null)
        {
            var boundary = right.AncestorsAndSelf().First(e => ReferenceEquals(e.Parent, paragraph));
            boundary.AddBeforeSelf(markers);
            return;
        }
        paragraph.Add(markers);
    }

    private static IReadOnlyList<Anchor> EndpointAnchors(ValidatedRange range) =>
        range.StartAnchor.Id == range.EndAnchor.Id
            ? new[] { range.StartAnchor }
            : new[] { range.StartAnchor, range.EndAnchor };

    private IReadOnlyList<BookmarkRangeSegment> BuildBookmarkSegments(
        IReadOnlyList<OwnedPartRelationships.Owner> owners,
        OwnedPartRelationships.Owner startOwner, XElement startParagraph, int startOffset,
        OwnedPartRelationships.Owner endOwner, XElement endParagraph, int endOffset)
    {
        var paragraphs = owners.SelectMany(o => o.Part.GetXDocument().Descendants(W.p).Select(p => (o, p))).ToList();
        int first = paragraphs.FindIndex(x => ReferenceEquals(x.p, startParagraph));
        int last = paragraphs.FindIndex(x => ReferenceEquals(x.p, endParagraph));
        if (first < 0 || last < first) return Array.Empty<BookmarkRangeSegment>();
        var result = new List<BookmarkRangeSegment>();
        for (int i = first; i <= last; i++)
        {
            var (owner, paragraph) = paragraphs[i];
            var anchor = AnchorForElement(paragraph);
            if (anchor is null) continue;
            var text = RunTextMap.Build(paragraph).FlatText;
            int from = i == first ? startOffset : 0;
            int to = i == last ? endOffset : text.Length;
            from = Math.Clamp(from, 0, text.Length);
            to = Math.Clamp(to, from, text.Length);
            result.Add(new BookmarkRangeSegment(owner.PartUri, owner.Scope, anchor.Value.Id,
                new CharSpan(from, to - from), text.Substring(from, to - from)));
        }
        return result;
    }

    private static int MarkerOffset(XElement paragraph, XElement marker)
    {
        int offset = 0;
        foreach (var run in InlineRuns(paragraph))
        {
            if (XNode.DocumentOrderComparer.Compare(run, marker) >= 0) break;
            offset += RunText(run).Length;
        }
        return offset;
    }

    private static string TextForRuns(IEnumerable<XElement> runs) =>
        string.Concat(runs.Select(RunText));
}
