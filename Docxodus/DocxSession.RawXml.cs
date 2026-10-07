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
    // ─── Raw escape hatch ────────────────────────────────────────────────

    public RawDocxOps Raw => _raw ??= new RawDocxOps(this);

    internal string RawGetXmlInternal(string anchorId)
    {
        ThrowIfDisposed();
        var target = FindAnchor(anchorId);
        if (target is null)
            throw new ArgumentException($"anchor not found: {anchorId}");
        var element = target.Resolve(_doc!);
        return element?.ToString() ?? "";
    }

    /// <summary>
    /// The live, in-memory document backing this session. Exposed for read-only,
    /// in-assembly consumers (e.g. session-attached single-block HTML rendering) that
    /// must read the current tree/parts without the round-trip cost of <see cref="Save"/>.
    /// Do not mutate it outside the session's own edit methods.
    /// </summary>
    internal WordprocessingDocument LiveDocument
    {
        get
        {
            ThrowIfDisposed();
            return _doc!;
        }
    }

    /// <summary>Release the cached block-render shell (rebuild happens lazily on next render).</summary>
    internal void DisposeRenderShell()
    {
        DiscardPackage(RenderShellDoc, RenderShellStream);
        RenderShellDoc = null;
        RenderShellStream = null;
        DenseTextRenderTemplates.Clear();
    }

    /// <summary>
    /// Lets go of a package whose contents nobody will read again, without closing it.
    /// </summary>
    /// <remarks>
    /// Closing a package opened for editing rewrites its whole archive into the backing stream,
    /// deflating every part touched since it was opened. The session's own package, the render
    /// shell, and the package a snapshot restore replaces all reach this point with no reader
    /// for that rewrite: the stream is discarded with them. So the package is dropped for the
    /// garbage collector instead of disposed, and only the stream is closed. The package owns
    /// nothing else — the SDK reads a part through a stream it closes as it goes, and the
    /// uncompressed part data it keeps between reads is managed memory. Besides skipping a
    /// deflate pass into a dead buffer, this keeps the editor's close path out of that rewrite:
    /// in the browser runtime it hung inside it in one reproducible sequence (save, convert the
    /// saved bytes, format three paragraphs, close), and the .NET runtime never did. Closing
    /// the stream first and letting the package's dispose fail on it is not an alternative: the
    /// exception that raises from inside the SDK's dispose chain trips a Mono assertion in the
    /// WASM runtime and takes the whole page down.
    /// </remarks>
    private static void DiscardPackage(WordprocessingDocument? doc, MemoryStream? stream)
    {
        _ = doc;
        stream?.Dispose();
    }

    internal EditResult RawInsertXmlInternal(string anchorId, Position pos, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var (parsedXml, err) = ParseRawXml(xml);
        if (parsedXml is null)
            return new EditResult { Success = false, Error = err! with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        int baselineErrors = _settings.ValidateRawOps ? CountRealValidationErrors() : 0;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            UnidHelper.AssignToSelfAndDescendants(parsedXml);
            if (pos == Position.Before) element.AddBeforeSelf(parsedXml);
            else element.AddAfterSelf(parsedXml);

            if (_settings.ValidateRawOps && CountRealValidationErrors() > baselineErrors)
            {
                var preOp = _history.PopForUndo();
                if (preOp.ok) RestoreSnapshot(preOp.snapshot);
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            InvalidateProjectionCache();
            var freshIndex = AnchorIndex();
            var created = new List<Anchor>();
            foreach (var unid in CollectUnids(parsedXml))
            {
                var hit = AnchorForUnid(unid, target.PartUri);
                if (hit is { } h) created.Add(h);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
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

    internal EditResult RawReplaceXmlInternal(string anchorId, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var (parsedXml, err) = ParseRawXml(xml);
        if (parsedXml is null)
            return new EditResult { Success = false, Error = err! with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        var relationshipOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, element);

        int baselineErrors = _settings.ValidateRawOps ? CountRealValidationErrors() : 0;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            UnidHelper.AssignToSelfAndDescendants(parsedXml);
            element.ReplaceWith(parsedXml);

            if (_settings.ValidateRawOps && CountRealValidationErrors() > baselineErrors)
            {
                var preOp = _history.PopForUndo();
                if (preOp.ok) RestoreSnapshot(preOp.snapshot);
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            if (relationshipOwner is { } owner)
                SweepOrphanedStoryRelationships(owner.Part);

            InvalidateProjectionCache();
            var freshIndex = AnchorIndex();
            var newUnids = CollectUnids(parsedXml).ToHashSet();

            // Classify by Unid set membership: the documented Get→mutate→Replace
            // recipe preserves Unids, so the target anchor must surface as
            // Modified (not as a phantom Removed-then-Created pair). When the
            // replacement XML has fresh Unids — because the caller authored it
            // from scratch — the target is genuinely Removed and the new
            // element(s) are Created. See DS092 / DS092b.
            var modified = new List<Anchor>();
            var removed = new List<Anchor>();
            var created = new List<Anchor>();

            if (newUnids.Contains(target.Unid))
            {
                var hit = AnchorForUnid(target.Unid, target.PartUri);
                if (hit is { } h) modified.Add(h);
            }
            else
            {
                removed.Add(target.Anchor);
            }
            // Most new Unids belong to run/property/text elements, which are
            // not addressable blocks. Resolve against the small anchor index
            // once rather than scanning it for every element of a dense frame.
            var anchorsByUnid = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
            foreach (var candidate in freshIndex.Values)
            {
                if (!anchorsByUnid.TryGetValue(candidate.Unid, out var prior)
                    || (prior.PartUri != target.PartUri && candidate.PartUri == target.PartUri))
                    anchorsByUnid[candidate.Unid] = candidate;
            }
            foreach (var unid in newUnids)
            {
                if (unid == target.Unid) continue;
                if (anchorsByUnid.TryGetValue(unid, out var hit)) created.Add(hit.Anchor);
            }

            return new EditResult
            {
                Success = true,
                Removed = removed,
                Created = created,
                Modified = modified,
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

    private static (XElement? parsed, EditError? err) ParseRawXml(string xml)
    {
        try
        {
            var x = XElement.Parse(xml);
            foreach (var el in x.DescendantsAndSelf())
            {
                var ns = el.Name.NamespaceName;
                if (!string.IsNullOrEmpty(ns) && !AllowedXmlNamespaces.Contains(ns))
                    return (null, new EditError(EditErrorCode.DisallowedNamespace,
                        $"disallowed namespace: {ns}"));
            }
            return (x, null);
        }
        catch (System.Xml.XmlException ex)
        {
            return (null, new EditError(EditErrorCode.MalformedXml, ex.Message));
        }
    }

    private static IEnumerable<string> CollectUnids(XElement root)
    {
        foreach (var el in root.DescendantsAndSelf())
        {
            var unid = (string?)el.Attribute(PtOpenXml.Unid);
            if (unid is not null) yield return unid;
        }
    }

    // PtOpenXml:Unid is an internal-only attribute added by the projector for anchor
    // addressing; it is not in the OOXML schema, so the validator will emit
    // Sch_UndeclaredAttribute for every occurrence. Filter those out before counting.
    //
    // Mutations operate directly on the part's in-memory XDocument; the validator
    // reads the typed OOXML object model, which is hydrated from the part stream.
    // Flush the XDocument back to the stream first so the validator sees the
    // current state instead of the original document.
    private int CountRealValidationErrors()
    {
        _doc!.MainDocumentPart!.PutXDocument();
        var v = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        return v.Validate(_doc!)
            .Count(e => !(e.Description ?? string.Empty)
                .Contains("http://powertools.codeplex.com/2011", StringComparison.Ordinal));
    }
}
