// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

public sealed partial class DocxSession
{
    /// <summary>Delete one complete, directly contained complex field as one reversible change.</summary>
    /// <remarks>The nonempty span must equal the field's complete visible result. This operation
    /// refuses nested fields, mixed field-marker runs and field contents under revision or other
    /// inline containers; a text edit must not silently consume those structures.</remarks>
    public EditResult DeleteFieldAtSpan(string anchorId, int spanStart, int spanLength)
    {
        lock (_mutationGate)
        {
            if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
            var target = FindAnchor(anchorId);
            if (target is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
            if (target.Anchor.Kind is not ("p" or "h" or "li"))
                return EditResult.Fail(EditErrorCode.AnchorWrongKind, "Field deletion requires a paragraph anchor", anchorId);
            var paragraph = target.Resolve(_doc!);
            if (paragraph is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "paragraph not found", anchorId);
            var map = RunTextMap.Build(paragraph);
            if (spanStart < 0 || spanLength <= 0 || spanStart > map.FlatText.Length || spanLength > map.FlatText.Length - spanStart)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange, "Field result span is out of range or empty", anchorId);
            var fields = ParagraphFields(paragraph)
                .Where(field => field.BeginAt == spanStart && field.EndAt == spanStart + spanLength)
                .ToArray();
            if (fields.Length != 1) return Unsupported();
            var field = fields[0];
            var first = field.Begin.Parent;
            var last = field.End.Parent;
            if (first?.Name != W.r || last?.Name != W.r || first.Parent != paragraph
                || last.Parent != paragraph || IsInsideComplexField(first)) return Unsupported();
            var runs = first.ElementsAfterSelf().Prepend(first).TakeWhile(e => e != last).Append(last).ToList();
            if (runs.Any(run => run.Name != W.r || run.Descendants(W.rPrChange).Any())) return Unsupported();
            bool separated = false;
            bool instruction = false;
            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                var markers = run.Elements(W.fldChar).ToArray();
                if (markers.Length > 0)
                {
                    if (markers.Length != 1 || run.Elements().Any(e => e.Name != W.rPr && e.Name != W.fldChar)) return Unsupported();
                    var kind = (string?)markers[0].Attribute(W.fldCharType);
                    if (i == 0 && kind == "begin") continue;
                    if (i == runs.Count - 1 && kind == "end" && separated && instruction) continue;
                    if (kind == "separate" && !separated) { separated = true; continue; }
                    return Unsupported();
                }
                if (run.Elements().Any(e => e.Name != W.rPr && e.Name != W.instrText && e.Name != W.t)) return Unsupported();
                if (!separated && run.Elements(W.instrText).Any(e => !string.IsNullOrWhiteSpace(e.Value))) instruction = true;
            }
            var segments = map.Segments.Where(segment => runs.Contains(segment.Run)).ToArray();
            if (segments.Sum(segment => segment.Length) != spanLength) return Unsupported();
            _history.RecordPreOp(TakeSnapshot());
            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                var deletion = CreateRevisionEnvelope(W.del, NewDistinctRevisionStamp());
                runs[0].AddBeforeSelf(deletion);
                foreach (var run in runs)
                {
                    run.Remove();
                    ConvertTextToDeletedText(run);
                    deletion.Add(run);
                }
            }
            else foreach (var run in runs) run.Remove();
            InvalidateProjectionCache();
            return new EditResult { Success = true, Modified = new[] { target.Anchor }, Patch = PatchFor(target) };

            EditResult Unsupported() => EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                "The span does not identify one complete, direct, unnested complex field with a nonempty result", anchorId);
        }
    }
}
