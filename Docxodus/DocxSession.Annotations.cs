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
    // ─── Tier E: annotations ────────────────────────────────────────────

    /// <summary>
    /// Annotate the range <paramref name="span"/> inside the block addressed by
    /// <paramref name="anchorId"/>. When <paramref name="span"/> is null, the
    /// annotation wraps every inline run of the block. When
    /// <paramref name="annotation"/>.Id is null/empty, a 16-char hex id is
    /// generated. The bookmark name, AnnotatedText, Created, and PageInfoStale
    /// fields of the annotation are always set by this method.
    /// </summary>
    public EditResult AddAnnotation(string anchorId, CharSpan? span, DocumentAnnotation annotation)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (annotation is null)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "annotation is null", anchorId);

        var anchor = FindAnchor(anchorId);
        if (anchor is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Add(_doc!, anchor, span, annotation);
            if (result.Success) InvalidateProjectionCache();
            else _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>Removes an annotation (its bookmark and custom-XML entry) by id.</summary>
    public EditResult RemoveAnnotation(string annotationId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Remove(_doc!, annotationId, CanonicalizeAnchorByUnid);
            if (result.Success) InvalidateProjectionCache();
            else _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>Mutates label/color/author/metadata of an annotation without re-targeting.</summary>
    public EditResult UpdateAnnotation(string annotationId, AnnotationUpdate update)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (update is null)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "update is null");

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Update(_doc!, annotationId, update);
            if (!result.Success) _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>Re-targets an existing annotation to a new anchor + span.</summary>
    public EditResult MoveAnnotation(string annotationId, string newAnchorId, CharSpan? newSpan)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var anchor = FindAnchor(newAnchorId);
        if (anchor is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"anchor not found: {newAnchorId}", newAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Move(
                _doc!, annotationId, anchor, newSpan, CanonicalizeAnchorByUnid);
            if (result.Success) InvalidateProjectionCache();
            else _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, newAnchorId);
        }
    }

    /// <summary>
    /// Looks up the canonical <see cref="Anchor"/> for a Unid in the current
    /// projection. Used by annotation ops so that the <see cref="EditResult.Modified"/>
    /// anchor matches what <see cref="Project"/>'s AnchorIndex will return on the
    /// next tick — bypasses the local kind/scope classifier in <c>AnnotationOps</c>
    /// drifting from the projector.
    /// </summary>
    private Anchor? CanonicalizeAnchorByUnid(string unid)
    {
        var idx = AnchorIndex();
        return idx.Values.FirstOrDefault(t => t.Unid == unid)?.Anchor;
    }
}
