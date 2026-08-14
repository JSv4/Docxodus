// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

public sealed partial class DocxSession
{
    private static readonly XNamespace LinkR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly Regex BookmarkNamePattern = new("^[A-Za-z_][A-Za-z0-9_]{0,39}$", RegexOptions.CultureInvariant);

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
            selected[0].AddBeforeSelf(link);
            foreach (var run in selected) { run.Remove(); link.Add(run); }
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
            foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
                foreach (var link in owner.Part.GetXDocument().Descendants(W.hyperlink)
                    .Where(h => string.Equals((string?)h.Attribute(W.anchor), name, StringComparison.Ordinal)))
                    link.SetAttributeValue(W.anchor, newName);
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
        if (OwnedPartRelationships.StoryParts(_doc!).Any(o => o.Part.GetXDocument().Descendants(W.hyperlink)
            .Any(h => string.Equals((string?)h.Attribute(W.anchor), name, StringComparison.Ordinal))))
            return EditResult.Fail(EditErrorCode.BookmarkInUse,
                $"bookmark is targeted by one or more internal hyperlinks: {name}");
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
            if (OwnedPartRelationships.StoryParts(_doc!).Any(owner =>
                owner.Part.GetXDocument().Descendants(W.hyperlink).Any(link =>
                    string.Equals((string?)link.Attribute(W.anchor), name, StringComparison.Ordinal)
                    && !IsRemoved(link))))
                return EditResult.Fail(EditErrorCode.BookmarkInUse,
                    $"structural deletion would remove a bookmark still targeted by an internal hyperlink: {name}", anchorId);
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
