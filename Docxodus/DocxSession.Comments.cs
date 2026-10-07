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
    /// <summary>
    /// Remove a native Word comment, addressed by its definition anchor (kind <c>cmt</c>):
    /// the <c>w:comment</c> definition, its body-side marker triple
    /// (<c>w:commentRangeStart</c>/<c>w:commentRangeEnd</c>/<c>w:commentReference</c>, wrapper
    /// run included) everywhere in the package, and any <c>commentsExtended</c>/
    /// <c>commentsIds</c> threading entries keyed by its paragraphs' <c>w14:paraId</c> — a
    /// surviving reply whose parent was removed becomes top-level. Delegates to the same
    /// teardown <see cref="DeleteBlock"/> performs for a <c>cmt</c> anchor; this wrapper adds
    /// only the comment-specific kind guard. The comments part itself is kept even when the
    /// last comment is removed (part deletion happens only via <see cref="Undo"/> of the
    /// create).
    /// </summary>
    public EditResult RemoveComment(string commentAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(commentAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {commentAnchorId}", commentAnchorId);
        if (target.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"RemoveComment requires a comment definition anchor (kind cmt); got kind={target.Anchor.Kind}",
                commentAnchorId);

        return DeleteBlock(commentAnchorId);
    }

    /// <summary>
    /// The document's native Word comments in comments-part order — see
    /// <see cref="CommentListEntry"/>. Read-only; returns an empty list when the document
    /// has no comments part.
    /// </summary>
    public IReadOnlyList<CommentListEntry> ListComments()
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids on the comments part

        var result = new List<CommentListEntry>();
        var main = _doc!.MainDocumentPart;
        var root = main?.WordprocessingCommentsPart?.GetXDocument().Root;
        if (root is null) return result;

        var comments = root.Elements(W.comment).ToList();
        var anchorByParaId = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var c in comments)
        {
            var unid = (string?)c.Attribute(PtOpenXml.Unid);
            var paraId = (string?)c.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
            if (unid is not null && paraId is not null)
                anchorByParaId[paraId] = $"cmt:cmt:{unid}";
        }

        var commentExByParaId = main?.WordprocessingCommentsExPart?.GetXDocument().Root?
            .Elements(Internal.CommentOps.W15 + "commentEx")
            .Where(e => (string?)e.Attribute(Internal.CommentOps.W15 + "paraId") is not null)
            .GroupBy(e => (string)e.Attribute(Internal.CommentOps.W15 + "paraId")!, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal)
            ?? new Dictionary<string, XElement>(StringComparer.Ordinal);

        foreach (var c in comments)
        {
            var unid = (string?)c.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;

            string? parentAnchorId = null;
            bool? resolved = null;
            var paraId = (string?)c.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
            if (paraId is not null && commentExByParaId.TryGetValue(paraId, out var commentEx))
            {
                resolved = Internal.CommentOps.ParseDone(
                    (string?)commentEx.Attribute(Internal.CommentOps.W15 + "done"));
                var parentParaId = (string?)commentEx.Attribute(Internal.CommentOps.W15 + "paraIdParent");
                if (parentParaId is not null)
                    anchorByParaId.TryGetValue(parentParaId, out parentAnchorId);
            }

            result.Add(new CommentListEntry(
                $"cmt:cmt:{unid}",
                (string?)c.Attribute(W.author) ?? "unknown",
                (string?)c.Attribute(W.initials),
                (string?)c.Attribute(W.date),
                Internal.CommentOps.FlattenBodyText(c))
            {
                Id = int.TryParse((string?)c.Attribute(W.id), System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var numericId) ? numericId : -1,
                ParentAnchorId = parentAnchorId,
                Resolved = resolved,
            });
        }
        return result;
    }

    // ─── Comments (issue #300) ───────────────────────────────────────────
    //
    // Native Word comment authoring, following the part-creation pattern the note ops above
    // established: find-or-create the WordprocessingCommentsPart + the CommentText/
    // CommentReference styles, bracket a character span with w:commentRangeStart/End, append
    // the run-level w:commentReference, and add the w:comment definition. Mechanics live in
    // Internal.CommentOps; part create/delete is undo/redo-reconciled by ReconcileCommentsPart.
    //
    // Editing a comment body needs no bespoke path beyond UpdateComment: comment paragraphs
    // project as kind p, scope cmt, so ReplaceText already accepts them; DeleteBlock already
    // removes a cmt definition together with its body-side marker triple.

    /// <summary>
    /// Add a <b>native Word comment</b> (a <c>w:comment</c> the Reviewing pane shows — not the
    /// <see cref="AddAnnotation"/> overlay) on the paragraph named by <paramref name="anchorId"/>.
    /// <paramref name="span"/> selects the commented character range; <c>null</c> comments the
    /// whole block. Creates the <c>WordprocessingCommentsPart</c> and the <c>CommentText</c>/
    /// <c>CommentReference</c> styles when absent. The comment body comes from
    /// <paramref name="markdownPayload"/> (same subset as <see cref="InsertFootnote"/>).
    /// <paramref name="date"/> is written only when provided, keeping output deterministic by
    /// default; an Unspecified-kind value is treated as UTC. Returns the created definition
    /// anchor (kind <c>cmt</c>) and its paragraph anchors (kind <c>p</c>, scope <c>cmt</c>) in
    /// <see cref="EditResult.Created"/> so a caller can immediately
    /// <see cref="UpdateComment"/>/<see cref="RemoveComment"/> it.
    /// </summary>
    /// <remarks>
    /// Body paragraphs only (kind <c>p</c>/<c>h</c>/<c>li</c>, scope <c>body</c>) — Word has no
    /// comments-on-comments, and v1 does not target header/footer/note stories. Spans are
    /// single-block; the numeric <c>w:id</c> is never surfaced (comments are addressed by anchor).
    /// </remarks>
    public EditResult AddComment(
        string anchorId, CharSpan? span, string author, string markdownPayload,
        string? initials = null, DateTime? date = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"AddComment requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"AddComment requires a body paragraph anchor; got scope '{target.Anchor.Scope}'", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        var totalText = ParagraphText(element);
        int spanStart, spanLength;
        if (span.HasValue)
        {
            spanStart = span.Value.Start;
            spanLength = span.Value.Length;
            if (spanLength <= 0)
                return EditResult.Fail(EditErrorCode.EmptyCommentSpan, "span length must be > 0", anchorId);
            if (spanStart < 0 || spanStart + spanLength > totalText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{spanStart},{spanStart + spanLength}) outside block of length {totalText.Length}", anchorId);
        }
        else
        {
            spanStart = 0;
            spanLength = totalText.Length;
            if (spanLength == 0)
                return EditResult.Fail(EditErrorCode.EmptyCommentSpan, "block has no text to comment", anchorId);
        }

        return AddCommentCore(author, markdownPayload, initials, date,
            placeMarkers: id =>
            {
                // Splits route through the same offset mechanism every other span op uses
                // (AnnotationOps.SplitRunsForSpan).
                var (startRun, endRun) = Internal.AnnotationOps.SplitRunsForSpan(
                    element, spanStart, spanLength);
                InsertCommentMarkers(id, startRun, endRun);
            },
            modified: new[] { target.Anchor },
            patchTarget: target,
            errorTargetId: anchorId);
    }

    /// <summary>
    /// Add a native Word comment anchored to the exact live markup extent of the tracked
    /// revision named by <paramref name="revisionId"/>. The id is one returned by
    /// <see cref="ListRevisions"/>; an unknown or already-resolved id fails with
    /// <see cref="EditErrorCode.RevisionNotFound"/>. Comment markers sit outside revision
    /// wrappers, so accepting/rejecting leaves the comment on surviving text or collapses its
    /// range to a point when that text vanishes.
    /// This is the named revision-target counterpart to the anchor/span
    /// <see cref="AddComment"/> operation.
    /// </summary>
    public EditResult AddCommentToRevision(
        string revisionId, string author, string markdownPayload,
        string? initials = null, DateTime? date = null)
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

        var commentTarget = Internal.RevisionOps.CommentTarget(group);
        if (commentTarget is null)
            return EditResult.Fail(EditErrorCode.RevisionNotFound,
                $"revision has no commentable extent: {revisionId}");

        var partUri = group.PartUri;
        var modified = RevisionGroupAnchors(group, partUri);
        return AddCommentCore(author, markdownPayload, initials, date,
            placeMarkers: id =>
            {
                if (commentTarget.First is not null && commentTarget.Last is not null)
                {
                    InsertCommentMarkers(id, commentTarget.First, commentTarget.Last);
                    return;
                }

                var point = commentTarget.PointParagraph
                    ?? throw new InvalidOperationException("revision comment target has no boundary");
                InsertPointCommentMarkers(id, point);
            },
            modified: modified,
            patchTarget: null,
            errorTargetId: null);
    }

    private EditResult AddCommentCore(
        string author, string markdownPayload, string? initials, DateTime? date,
        Action<int> placeMarkers, IReadOnlyList<Anchor> modified,
        AnchorTarget? patchTarget, string? errorTargetId)
    {
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", errorTargetId);

        // Parse the body BEFORE snapshotting so malformed payloads are clean no-ops.
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, errorTargetId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var part = Internal.CommentOps.EnsureCommentsPart(main);
            Internal.StyleFactory.EnsureCommentStyles(_doc!);
            var id = Internal.CommentOps.NextCommentId(main);
            var idStr = id.ToString(System.Globalization.CultureInfo.InvariantCulture);
            placeMarkers(id);

            Internal.CommentOps.ApplyCommentBodyStyle(paras);
            var comment = new XElement(W.comment,
                new XAttribute(W.id, idStr),
                new XAttribute(W.author, author));
            if (!string.IsNullOrEmpty(initials))
                comment.SetAttributeValue(W.initials, initials);
            if (date.HasValue)
                comment.SetAttributeValue(W.date, Internal.CommentOps.FormatDate(date.Value));
            foreach (var p in paras) comment.Add(p);
            part.GetXDocument().Root!.Add(comment);
            UnidHelper.AssignToSelfAndDescendants(comment);
            part.PutXDocument();

            CommentsVersion++;
            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var commentsPartUri = part.Uri.ToString();
            if (AnchorForUnid((string?)comment.Attribute(PtOpenXml.Unid), commentsPartUri) is { } defAnchor)
                created.Add(defAnchor);
            foreach (var p in comment.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), commentsPartUri) is { } pa)
                    created.Add(pa);

            return new EditResult
            {
                Success = true,
                Created = created,
                Modified = modified,
                Patch = patchTarget is null ? null : PatchFor(patchTarget),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, errorTargetId);
        }
    }

    private static void InsertCommentMarkers(int id, XElement first, XElement last)
    {
        var idStr = id.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var rangeStart = new XElement(W.commentRangeStart, new XAttribute(W.id, idStr));
        var rangeEnd = new XElement(W.commentRangeEnd, new XAttribute(W.id, idStr));
        first.AddBeforeSelf(rangeStart);
        last.AddAfterSelf(rangeEnd);
        var refRun = Internal.CommentOps.BuildReferenceRun(id);
        UnidHelper.AssignToSelfAndDescendants(refRun);
        rangeEnd.AddAfterSelf(refRun);
    }

    private static void InsertPointCommentMarkers(int id, XElement paragraph)
    {
        var idStr = id.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var rangeStart = new XElement(W.commentRangeStart, new XAttribute(W.id, idStr));
        var rangeEnd = new XElement(W.commentRangeEnd, new XAttribute(W.id, idStr));
        var refRun = Internal.CommentOps.BuildReferenceRun(id);
        UnidHelper.AssignToSelfAndDescendants(refRun);
        paragraph.Add(rangeStart, rangeEnd, refRun);
    }

    /// <summary>
    /// Add a native Word <b>reply</b> to the comment addressed by
    /// <paramref name="parentCommentAnchorId"/>. The reply receives its own
    /// <c>w:comment</c> definition and marker id, adds an adjacent reference at the parent's
    /// native thread anchor, and links to it through <c>w15:paraIdParent</c> in a find-or-created
    /// <c>commentsExtended.xml</c>. A matching <c>commentsIds.xml</c> entry is also created for
    /// both sides when absent. Metadata ids are allocated deterministically.
    /// </summary>
    /// <remarks>
    /// The parent may itself be a reply. An orphaned definition with no live
    /// <c>w:commentReference</c> cannot be replied to because it has no document position to
    /// share. Returns the new definition and comment-body paragraph anchors in
    /// <see cref="EditResult.Created"/>.
    /// </remarks>
    public EditResult AddCommentReply(
        string parentCommentAnchorId, string author, string markdownPayload,
        string? initials = null, DateTime? date = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var parentTarget = FindAnchor(parentCommentAnchorId);
        if (parentTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"anchor not found: {parentCommentAnchorId}", parentCommentAnchorId);
        if (parentTarget.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"AddCommentReply requires a comment definition anchor (kind cmt); got kind={parentTarget.Anchor.Kind}",
                parentCommentAnchorId);

        var parentComment = parentTarget.Resolve(_doc!);
        if (parentComment is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", parentCommentAnchorId);
        var main = _doc!.MainDocumentPart;
        if (main?.WordprocessingCommentsPart is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no comments part", parentCommentAnchorId);
        var parentId = (string?)parentComment.Attribute(W.id);
        if (string.IsNullOrEmpty(parentId))
            return EditResult.Fail(EditErrorCode.InternalError,
                "parent comment definition has no w:id", parentCommentAnchorId);
        if (!Internal.CommentOps.HasCommentReference(main, parentId))
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                "parent comment definition has no live document reference", parentCommentAnchorId);

        // Parse before snapshotting so malformed markdown is a clean no-op.
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, parentCommentAnchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.StyleFactory.EnsureCommentStyles(_doc!);
            var id = Internal.CommentOps.NextCommentId(main);
            var idStr = id.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // Word keeps range markers on the thread root; each reply adds only an adjacent
            // reference and inherits that range through commentsExtended parentage.
            var hostBlocks = Internal.CommentOps.InsertReplyReference(main, parentId, id);

            Internal.CommentOps.ApplyCommentBodyStyle(paras);
            var reply = new XElement(W.comment,
                new XAttribute(W.id, idStr),
                new XAttribute(W.author, author));
            if (!string.IsNullOrEmpty(initials))
                reply.SetAttributeValue(W.initials, initials);
            if (date.HasValue)
                reply.SetAttributeValue(W.date, Internal.CommentOps.FormatDate(date.Value));
            foreach (var p in paras) reply.Add(p);
            main.WordprocessingCommentsPart.GetXDocument().Root!.Add(reply);
            UnidHelper.AssignToSelfAndDescendants(reply);

            // Upgrade a flat parent only as far as needed: one extension/id entry for it and one
            // for the reply. Existing thread/resolve metadata is preserved.
            var parentParaId = Internal.CommentOps.EnsureThreadingMetadata(main, parentComment);
            Internal.CommentOps.EnsureThreadingMetadata(main, reply,
                parentParaId: parentParaId, resolved: false);

            CommentsVersion++;
            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var commentsPartUri = main.WordprocessingCommentsPart.Uri.ToString();
            if (AnchorForUnid((string?)reply.Attribute(PtOpenXml.Unid), commentsPartUri) is { } defAnchor)
                created.Add(defAnchor);
            foreach (var p in reply.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), commentsPartUri) is { } pa)
                    created.Add(pa);

            // The parent gains/participates in extension metadata, while every document-side
            // reference host gains a new run. Report both semantic mutations; patch the first
            // host because that is the rendered document block callers need to refresh.
            var modified = new List<Anchor> { parentTarget.Anchor };
            var seenModified = new HashSet<string>(StringComparer.Ordinal)
            {
                parentTarget.Anchor.Id,
            };
            string? patchHostAnchorId = null;
            foreach (var hostBlock in hostBlocks)
            {
                if (AnchorForElement(hostBlock) is not { } anchor) continue;
                patchHostAnchorId ??= anchor.Id;
                if (seenModified.Add(anchor.Id)) modified.Add(anchor);
            }
            var patchTarget = patchHostAnchorId is null ? null : FindAnchor(patchHostAnchorId);

            return new EditResult
            {
                Success = true,
                Created = created,
                Modified = modified,
                Patch = patchTarget is null ? null : PatchFor(patchTarget),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, parentCommentAnchorId);
        }
    }

    /// <summary>
    /// Mark a comment and its reply subtree resolved or reopened by setting <c>w15:done</c>. A flat
    /// comment is upgraded in place: its last paragraph receives a deterministic
    /// <c>w14:paraId</c>, and <c>commentsExtended.xml</c>/<c>commentsIds.xml</c> are
    /// find-or-created. Existing reply parentage is preserved. The mutation is fully undoable,
    /// including first-time part creation.
    /// </summary>
    public EditResult SetCommentResolved(string commentAnchorId, bool resolved)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(commentAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {commentAnchorId}", commentAnchorId);
        if (target.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"SetCommentResolved requires a comment definition anchor (kind cmt); got kind={target.Anchor.Kind}",
                commentAnchorId);

        var comment = target.Resolve(_doc!);
        if (comment is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", commentAnchorId);
        var main = _doc!.MainDocumentPart;
        if (main?.WordprocessingCommentsPart is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no comments part", commentAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var rootParaId = Internal.CommentOps.EnsureThreadingMetadata(main, comment, resolved: resolved);

            // Word treats resolution as a thread/subtree state. Walk paraIdParent edges so
            // resolving a root also marks every reply, while resolving a nested reply affects
            // only that reply and its descendants.
            var exPart = main.WordprocessingCommentsExPart!;
            var exEntries = exPart.GetXDocument().Root!
                .Elements(Internal.CommentOps.W15 + "commentEx").ToList();
            var affectedParaIds = new HashSet<string>(StringComparer.Ordinal) { rootParaId };
            bool expanded;
            do
            {
                expanded = false;
                foreach (var entry in exEntries)
                {
                    var paraId = (string?)entry.Attribute(Internal.CommentOps.W15 + "paraId");
                    var parentParaId = (string?)entry.Attribute(Internal.CommentOps.W15 + "paraIdParent");
                    if (paraId is not null && parentParaId is not null
                        && affectedParaIds.Contains(parentParaId) && affectedParaIds.Add(paraId))
                        expanded = true;
                }
            } while (expanded);

            foreach (var entry in exEntries.Where(e =>
                         affectedParaIds.Contains((string?)e.Attribute(Internal.CommentOps.W15 + "paraId") ?? "")))
                entry.SetAttributeValue(Internal.CommentOps.W15 + "done", resolved ? "1" : "0");
            exPart.PutXDocument();

            CommentsVersion++;
            InvalidateProjectionCache();
            var modified = new List<Anchor>();
            foreach (var affectedComment in main.WordprocessingCommentsPart.GetXDocument().Root!
                         .Elements(W.comment))
            {
                var paraId = (string?)affectedComment.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
                var unid = (string?)affectedComment.Attribute(PtOpenXml.Unid);
                if (paraId is not null && unid is not null && affectedParaIds.Contains(paraId))
                    modified.Add(new Anchor($"cmt:cmt:{unid}", "cmt", "cmt", unid));
            }
            return new EditResult
            {
                Success = true,
                Modified = modified.Count == 0 ? new[] { target.Anchor } : modified,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, commentAnchorId);
        }
    }

    /// <summary>
    /// Replace a comment's <b>body text</b> with <paramref name="markdownPayload"/>, addressed by
    /// its definition anchor (kind <c>cmt</c>, from <see cref="EditResult.Created"/> or the
    /// projection's <c># Comments</c> tokens). The comment's identity attributes
    /// (<c>w:id</c>/<c>w:author</c>/<c>w:initials</c>/<c>w:date</c>) are untouched. When the old
    /// last paragraph carried a <c>w14:paraId</c> (a Word-threaded comment), the id is re-stamped
    /// on the new last paragraph — <c>commentsExtended.xml</c> entries key on it, so a body edit
    /// must not orphan Word's reply/resolve metadata.
    /// </summary>
    public EditResult UpdateComment(string commentAnchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(commentAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {commentAnchorId}", commentAnchorId);
        if (target.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"UpdateComment requires a comment definition anchor (kind cmt); got kind={target.Anchor.Kind}",
                commentAnchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", commentAnchorId);
        var main = _doc!.MainDocumentPart;
        if (main?.WordprocessingCommentsPart is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no comments part", commentAnchorId);

        // Parse BEFORE snapshotting so a malformed payload is a clean no-op.
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, commentAnchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.StyleFactory.EnsureCommentStyles(_doc!);

            var oldParas = element.Elements(W.p).ToList();
            var preservedParaId = (string?)oldParas.LastOrDefault()?.Attribute(W14.paraId);

            // Collect the outgoing paragraph anchors before removal.
            var index = AnchorIndex();
            var removed = new List<Anchor>();
            foreach (var p in oldParas)
            {
                var unid = (string?)p.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                foreach (var kv in index)
                    if (kv.Value.Unid == unid)
                        removed.Add(kv.Value.Anchor);
            }

            foreach (var p in oldParas) p.Remove();
            Internal.CommentOps.ApplyCommentBodyStyle(paras);
            foreach (var p in paras) element.Add(p);
            if (preservedParaId is not null)
                paras[paras.Count - 1].SetAttributeValue(W14.paraId, preservedParaId);
            foreach (var p in paras) UnidHelper.AssignToSelfAndDescendants(p);
            main.WordprocessingCommentsPart.PutXDocument();
            SweepOrphanedStoryRelationships(main.WordprocessingCommentsPart);

            CommentsVersion++;
            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var commentsPartUri = main.WordprocessingCommentsPart.Uri.ToString();
            foreach (var p in element.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), commentsPartUri) is { } pa)
                    created.Add(pa);

            return new EditResult
            {
                Success = true,
                Created = created,
                Removed = removed,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, commentAnchorId);
        }
    }

    /// <summary>
    /// Insert <paramref name="newChild"/> into <paramref name="paragraph"/> at
    /// <paramref name="offset"/> characters into its text — before the first child that starts at
    /// or past the offset, else appended. Callers must have cleared the boundary first
    /// (<see cref="SplitRunsAtOffset"/> + <see cref="SplitInlineContainersAtOffset"/>); this is the
    /// insert-side counterpart of <see cref="MoveInlineChildrenAfter"/> and counts positions the
    /// same way, so zero-width markers sandwiched at the offset keep the ref inside their range.
    /// </summary>
    private static void InsertInlineAtOffset(XElement paragraph, int offset, XElement newChild)
    {
        int consumed = 0;
        foreach (var child in paragraph.Elements().ToList())
        {
            if (child.Name == W.pPr) continue;
            if (consumed >= offset) { child.AddBeforeSelf(newChild); return; }
            consumed += IsInlineChild(child) ? InlineChildTextLength(child) : 0;
        }
        paragraph.Add(newChild);
    }

    /// <summary>
    /// Insert a <paramref name="rows"/>×<paramref name="cols"/> table before/after the block named
    /// by <paramref name="anchorId"/>. <paramref name="options"/> controls borders, per-cell markdown
    /// (row-major), and cell alignment. Returns the created canonical <c>tc</c> anchors (row-major).
    /// </summary>
    public EditResult InsertTable(string anchorId, Position pos, int rows, int cols, TableInsertOptions? options = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (rows < 1 || cols < 1)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "table needs >= 1 row and >= 1 column", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        var opts = options ?? new TableInsertOptions();
        var contents = opts.CellContents;

        if (contents is not null)
        {
            foreach (var markdown in contents.Where(s => !string.IsNullOrEmpty(s)))
            {
                var parsedCell = Internal.MarkdownPayloadParser.Parse(markdown!);
                if (parsedCell.Success
                    && ValidatePendingHyperlinks(parsedCell.Blocks.SelectMany(b => b.RunElements), anchorId) is { } linkError)
                    return linkError;
            }
        }

        // Explicit per-column widths: one per column, all positive. A mismatched count is a
        // caller error — reject rather than silently equalize (no silent caps).
        var colWidths = opts.ColumnWidths;
        if (colWidths is not null && (colWidths.Count != cols || colWidths.Any(w => w <= 0)))
            return EditResult.Fail(EditErrorCode.MalformedMarkdown,
                $"ColumnWidths must have one positive width per column ({cols}); got {colWidths.Count}", anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("InsertTable", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            const int contentTwips = 9576;           // ~6.65", a US-Letter content width
            int colTwips = contentTwips / cols;
            int Width(int c) => colWidths is not null ? colWidths[c] : colTwips;

            // With explicit widths the table is sized to their sum (dxa); otherwise it fills
            // the content area (100% pct) and splits equally.
            var tblW = colWidths is not null
                ? new XElement(W.tblW, new XAttribute(W._w, colWidths.Sum()), new XAttribute(W.type, "dxa"))
                : new XElement(W.tblW, new XAttribute(W._w, 5000), new XAttribute(W.type, "pct"));

            var tblPr = new XElement(W.tblPr,
                tblW,
                BuildTableBorders(opts.Borderless),
                new XElement(W.tblLayout, new XAttribute(W.type, "fixed")));

            var tblGrid = new XElement(W.tblGrid);
            for (int c = 0; c < cols; c++)
                tblGrid.Add(new XElement(W.gridCol, new XAttribute(W._w, Width(c))));

            var tbl = new XElement(W.tbl, tblPr, tblGrid);
            var cellParagraphs = new List<XElement>();
            var cells = new List<XElement>();

            for (int r = 0; r < rows; r++)
            {
                var tr = new XElement(W.tr);
                for (int c = 0; c < cols; c++)
                {
                    var tc = new XElement(W.tc,
                        new XElement(W.tcPr, new XElement(W.tcW, new XAttribute(W._w, Width(c)), new XAttribute(W.type, "dxa"))));

                    int idx = r * cols + c;
                    string? md = contents is not null && idx < contents.Count ? contents[idx] : null;
                    var paras = BuildCellParagraphs(md, opts.CellAlignment);
                    foreach (var p in paras) tc.Add(p);
                    cellParagraphs.AddRange(paras);
                    cells.Add(tc);
                    tr.Add(tc);
                }
                tbl.Add(tr);
            }

            UnidHelper.AssignToSelfAndDescendants(tbl);

            if (pos == Position.Before) element.AddBeforeSelf(tbl);
            else element.AddAfterSelf(tbl);

            // A table must be followed by a paragraph: Word's convention is to keep a w:p after
            // every table, and an end-of-body table with no trailing paragraph leaves no editable
            // block below it (S-1 smoke-test finding 2). If nothing — or only a sectPr / another
            // table — follows, append an empty trailing paragraph.
            var afterTbl = tbl.ElementsAfterSelf().FirstOrDefault();
            if (afterTbl is null || afterTbl.Name == W.sectPr || afterTbl.Name == W.tbl)
            {
                var trailing = new XElement(W.p);
                UnidHelper.AssignToSelfAndDescendants(trailing);
                tbl.AddAfterSelf(trailing);
            }

            foreach (var p in cellParagraphs) PromoteHyperlinkRelationships(p);

            InvalidateProjectionCache();
            var created = ResolveAnchorsForElements(cells);
            var metadata = Internal.TableGridModel.BuildMetadata(tbl, AnchorForElement);

            return new EditResult
            {
                Success = true,
                Created = created,
                TableAnchors = Internal.TableGridModel.Map(null, metadata),
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>Build the cell's paragraph(s) from optional markdown + alignment. Always >= 1 paragraph.</summary>
    private List<XElement> BuildCellParagraphs(string? markdown, ParagraphAlignment? align)
    {
        var result = new List<XElement>();
        if (!string.IsNullOrEmpty(markdown))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdown);
            if (parsed.Success)
                foreach (var block in parsed.Blocks)
                    result.Add(BuildParagraphFromParsedBlock(block));
        }
        if (result.Count == 0) result.Add(new XElement(W.p));

        if (align is { } a)
        {
            var val = a switch
            {
                ParagraphAlignment.Center => "center",
                ParagraphAlignment.Right => "right",
                ParagraphAlignment.Justify => "both",
                _ => "left",
            };
            foreach (var p in result)
            {
                var pPr = p.Element(W.pPr);
                if (pPr is null) { pPr = new XElement(W.pPr); p.AddFirst(pPr); }
                SetPPrChildInOrder(pPr, new XElement(W.jc, new XAttribute(W.val, val)));
            }
        }
        return result;
    }

    private static XElement BuildTableBorders(bool borderless)
    {
        var edges = new[] { W.top, W.left, W.bottom, W.right, W.insideH, W.insideV };
        var bdr = new XElement(W.tblBorders);
        foreach (var e in edges)
            bdr.Add(borderless
                ? new XElement(e, new XAttribute(W.val, "none"), new XAttribute(W.sz, 0),
                    new XAttribute(W.space, 0), new XAttribute(W.color, "auto"))
                : new XElement(e, new XAttribute(W.val, "single"), new XAttribute(W.sz, 4),
                    new XAttribute(W.space, 0), new XAttribute(W.color, "auto")));
        return bdr;
    }
}
