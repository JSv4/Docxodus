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
    /// Surgical text replacement within a single paragraph/heading/list-item: finds every
    /// literal occurrence of <paramref name="find"/> in the anchor's flat text and replaces
    /// it with <paramref name="replace"/>, preserving the surrounding run formatting that
    /// the match didn't touch. Returns one <see cref="EditResult"/> per attempted match.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A <paramref name="find"/> that matches nothing fails with a single
    /// <see cref="EditErrorCode.TextNotFound"/> result naming the anchor and the needle —
    /// never an empty list a caller could mistake for success. The exceptions are the two
    /// explicit no-op spellings: <see cref="ReplaceOptions.ExpectedMatchCount"/> = 0
    /// (assert absence) and <see cref="ReplaceOptions.MaxReplacements"/> = 0 (found but
    /// deliberately unconsumed), which return an empty list.
    /// </para>
    /// <para>
    /// The replacement text is plain-text and inherits the formatting of the FIRST run the
    /// match spanned — middle/trailing runs keep their <c>w:rPr</c> but lose the slice of
    /// text the match consumed (so a bold run that contributed three chars to the match now
    /// has those three chars gone, but stays bold for everything else it held).
    /// </para>
    /// <para>
    /// Matches are applied in reverse document order so multiple matches in the same
    /// paragraph don't invalidate each other's offsets. The whole call records a single undo
    /// snapshot — <see cref="Undo"/> rolls back every replacement together.
    /// </para>
    /// <para>
    /// In <see cref="TrackedChangeMode.RenderInline"/>, the untouched prefix/suffix runs stay
    /// ordinary text, the selected slices are copied (with their original run properties) to
    /// <c>w:del/w:r/w:delText</c>, and one first-run-formatted replacement is emitted as
    /// <c>w:ins/w:r/w:t</c>. Inline containers and zero-width semantic markers remain in place.
    /// </para>
    /// </remarks>
    public IReadOnlyList<EditResult> ReplaceTextRange(
        string anchorId,
        string find,
        string replace,
        ReplaceOptions? options = null)
    {
        lock (_mutationGate)
            return ReplaceTextRangeCore(anchorId, find, replace, options);
    }

    private IReadOnlyList<EditResult> ReplaceTextRangeCore(
        string anchorId,
        string find,
        string replace,
        ReplaceOptions? options)
    {
        if (_disposed)
            return new[] { EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed") };
        if (string.IsNullOrEmpty(find))
            return new[] { EditResult.Fail(EditErrorCode.MalformedMarkdown, "find must be non-empty", anchorId) };

        var opts = options ?? new ReplaceOptions();
        var guards = opts.Preconditions;
        if (guards is not null && guards.AnchorId is null)
            guards = guards with { AnchorId = anchorId };
        if (opts.ExpectedMatchCount is { } expectedCount)
            guards = (guards ?? new MutationPreconditions { AnchorId = anchorId }) with
            {
                ExpectedMatchCount = expectedCount,
            };
        if (EvaluatePreconditions(guards) is { } initialPreconditionError)
            return new[] { new EditResult { Success = false, Error = initialPreconditionError } };

        var target = FindAnchor(anchorId);
        if (target is null)
            return new[] { EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId) };
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return new[] { EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceTextRange requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId) };

        var regexOpts = opts.IgnoreCase
            ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
            : System.Text.RegularExpressions.RegexOptions.None;
        var pattern = System.Text.RegularExpressions.Regex.Escape(find);
        replace = MaybeApplySmartQuotes(replace);

        var matches = Grep(pattern, regexOpts)
            .Where(m => m.EnclosingAnchor.Anchor.Id == target.Anchor.Id)
            .ToList();
        if (EvaluatePreconditions(guards, matches.Count) is { } countPreconditionError)
            return new[] { new EditResult { Success = false, Error = countPreconditionError } };
        if (matches.Count == 0)
        {
            // ExpectedMatchCount = 0 already asserted absence above, so an empty result is
            // the caller's asserted outcome; any other zero-match call is a targeting
            // failure the caller must be able to distinguish from success (issue #490).
            if (guards?.ExpectedMatchCount == 0) return Array.Empty<EditResult>();
            return new[] { EditResult.Fail(EditErrorCode.TextNotFound,
                $"text not found in anchor's visible text: \"{find}\"", anchorId) };
        }

        if (opts.MaxReplacements is int cap) matches = matches.Take(cap).ToList();
        if (matches.Count == 0) return Array.Empty<EditResult>();

        var element = target.Resolve(_doc!);
        if (element is null)
            return new[] { EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId) };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
            var revisionAuthor = _revisionAuthor ?? "docxodus";
            var revisionDate = tracked
                ? RevisionDateNow()
                : null;

            // Reverse offset order so earlier-offset matches' SpanInElement stays valid
            // after later-offset edits land — see DS112/DS115.
            foreach (var match in matches.OrderByDescending(m => m.Span.Start))
            {
                if (tracked)
                    ApplyFragmentReplacementTracked(
                        element, match, replace, revisionAuthor, revisionDate!);
                else
                    ApplyFragmentReplacement(element, match, replace);
            }

            InvalidateProjectionCache();
            var success = new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
            return Enumerable.Repeat(success, matches.Count).ToArray();
        }
        catch (Exception ex)
        {
            return new[] { FailInternal(ex, anchorId) };
        }
    }

    /// <summary>
    /// Convenience: replace a single <see cref="TextMatch"/> (typically from <see cref="Grep"/>)
    /// in place with <paramref name="replace"/>. Same fragment-formatting semantics as
    /// <see cref="ReplaceTextRange"/>.
    /// </summary>
    public EditResult ReplaceMatch(TextMatch match, string replace)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");
        return ReplaceTextAtSpan(match.EnclosingAnchor.Anchor.Id, match.Span.Start, match.Span.Length, replace);
    }

    /// <summary>
    /// Replace a paragraph-local text span and apply <paramref name="format"/> to exactly the
    /// replacement text as one atomic edit and undo/version unit. Uses the same text,
    /// insertion and tracked-format semantics as ReplaceTextAtSpan followed by ApplyFormat.
    /// An empty replacement deletes the span without formatting adjacent text.
    /// </summary>
    /// <remarks>
    /// Intended for interactive typing: returns an ordinary EditResult, with no package hash or
    /// batch receipt. Use ExecuteBatch when those receipts or arbitrary compositions are needed.
    /// </remarks>
    public EditResult ReplaceTextAtSpanWithFormat(
        string anchorId, int spanStart, int spanLength, string replace, FormatOp format)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (format is null) return EditResult.Fail(EditErrorCode.MalformedMarkdown, "null format op", anchorId);

        // These two operations touch only snapshot-scoped XML/styles and owned story images.
        // Reuse transaction rollback/history handling without serializing the entire OPC package.
        using var transaction = BeginTransaction(fullPackage: false);
        var result = ReplaceTextAtSpan(anchorId, spanStart, spanLength, replace);
        if (!result.Success) return result;
        if (replace.Length > 0)
        {
            result = ApplyFormat(anchorId, new CharSpan(spanStart, replace.Length), format);
            if (!result.Success) return result;
        }
        transaction.Commit();
        return result;
    }

    /// <summary>
    /// Replace the bracketed portion of a <see cref="TextMatch"/> with <paramref name="newInner"/>,
    /// preserving any prefix or suffix outside the brackets. Designed for
    /// <see cref="FindPlaceholders"/> matches like <c>$[___]</c> where the regex
    /// <c>\$?\[…\]</c> captures the leading <c>$</c>: <c>ReplaceInner(match, "0.20")</c>
    /// yields <c>$0.20</c> (not <c>0.20</c>). For matches without any prefix/suffix,
    /// this is equivalent to <see cref="ReplaceMatch"/> with the new inner value.
    /// Returns <see cref="EditErrorCode.MalformedMarkdown"/> if the match text does
    /// not contain balanced brackets.
    /// </summary>
    public EditResult ReplaceInner(TextMatch match, string newInner)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");

        int lb = match.Text.IndexOf('[');
        int rb = match.Text.LastIndexOf(']');
        if (lb < 0 || rb <= lb)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown,
                $"match text has no balanced brackets: '{match.Text}'");

        var prefix = match.Text[..lb];
        var suffix = match.Text[(rb + 1)..];
        return ReplaceMatch(match, prefix + newInner + suffix);
    }

    /// <summary>
    /// Surgical replacement of an exact character range within one block's flat text.
    /// The natural pair to <see cref="Grep"/>: pass the <see cref="TextMatch.EnclosingAnchor"/>'s
    /// id plus the <see cref="TextMatch.Span"/> coordinates to replace one specific match
    /// even when several identical needles share the same paragraph (the template-filling
    /// case where five <c>[___]</c> placeholders each get a different value).
    /// </summary>
    public EditResult ReplaceTextAtSpan(string anchorId, int spanStart, int spanLength, string replace)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceTextAtSpan requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        replace = MaybeApplySmartQuotes(replace);

        var map = Internal.RunTextMap.Build(element);
        if (spanStart < 0 || spanLength < 0 || spanStart + spanLength > map.FlatText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span {spanStart}+{spanLength} out of [0, {map.FlatText.Length}]", anchorId);

        if (spanLength == 0)
            return InsertTextAtBoundary(target, element, map, spanStart, replace);

        var pieces = Internal.RunTextMap.ResolveRange(map, spanStart, spanLength);
        if (pieces.Count == 0)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange, "span resolved to no runs", anchorId);

        // Synthesize fragments from the resolved pieces. The replacement helper only
        // reads Unid + SpanInElement, so the other fields are placeholders.
        var fragments = new List<RunFragment>(pieces.Count);
        foreach (var (seg, offsetInRun, len) in pieces)
        {
            var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
            fragments.Add(new RunFragment
            {
                Unid = runUnid,
                Text = string.Empty,
                SpanInElement = new CharSpan(offsetInRun, len),
                Formatting = new RunFormatting(),
            });
        }
        var synthetic = new TextMatch
        {
            Text = map.FlatText.Substring(spanStart, spanLength),
            EnclosingAnchor = target,
            Span = new CharSpan(spanStart, spanLength),
            Fragments = fragments,
            ContextBefore = string.Empty,
            ContextAfter = string.Empty,
        };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                ApplyFragmentReplacementTracked(
                    element,
                    synthetic,
                    replace,
                    _revisionAuthor ?? "docxodus",
                    RevisionDateNow());
            }
            else
            {
                ApplyFragmentReplacement(element, synthetic, replace);
            }
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// The zero-length form of <see cref="ReplaceTextAtSpan"/>: insert <paramref name="replace"/>
    /// as a new run, splitting an ordinary text run when the caret is inside it. The
    /// difference matters next to a complex field: a browser editor that types after
    /// <c>Page {PAGE} of {NUMPAGES}</c> must not put its keystrokes inside the NUMPAGES result run,
    /// where Word's next field update would discard them. The new run copies the formatting of
    /// the run it follows (or precedes, at offset 0), minus any revision marker, and lands
    /// OUTSIDE any field whose chrome surrounds the boundary — after the field's <c>end</c> run,
    /// or before its <c>begin</c>. Interior insertion accepts text runs with optional leading tabs
    /// directly in the paragraph, outside fields, and must not split a UTF-16 surrogate pair.
    /// </summary>
    private EditResult InsertTextAtBoundary(
        AnchorTarget target, XElement element, Internal.RunTextMap.Map map, int offset, string replace)
    {
        var anchorId = target.Anchor.Id;
        if (replace.Length == 0)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        // The run whose text ends exactly at `offset` (insert after it), else the run whose text
        // starts there (insert before it). Empty paragraphs and offset 0 with no run resolve to the
        // paragraph itself.
        XElement? after = null;
        XElement? before = null;
        bool splitRun = false;
        foreach (var seg in map.Segments)
        {
            if (seg.Length == 0) continue;
            if (seg.EndOffsetInBlock == offset) { after = seg.Run; break; }
            if (seg.StartOffsetInBlock == offset) { before = seg.Run; break; }
            if (seg.StartOffsetInBlock < offset && offset < seg.EndOffsetInBlock)
            {
                if (!CanSplitTextRun(seg.Run) || !ReferenceEquals(seg.Run.Parent, element)
                    || IsInsideComplexField(seg.Run)
                    || char.IsSurrogatePair(map.FlatText, offset - 1))
                    return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                        "interior insertion requires an ordinary character boundary in a text run with only optional leading tabs, outside fields and inline containers", anchorId);
                after = seg.Run;
                splitRun = true;
                break;
            }
        }
        if (after is null && before is null && offset != 0 && map.FlatText.Length != 0)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange, "span resolved to no runs", anchorId);

        var rPrSource = (after ?? before)?.Element(W.rPr);
        XElement? rPr = null;
        if (rPrSource is not null)
        {
            rPr = new XElement(rPrSource);
            rPr.Elements(W.rPrChange).Remove();
            if (!rPr.HasElements) rPr = null;
        }

        // Step out of an inline container whose EDGE the boundary sits on: `after` the last run
        // of another author's w:ins (or a w:moveTo, hyperlink or smart tag) means after the
        // container, not inside it — text typed at the end of B's tracked insertion with tracking
        // off must not become part of B's change, and Word does not grow a link from text typed
        // at its end. A boundary strictly inside the container stays inside it. Then step out of
        // a complex field: `after` inside a field means "after the field's end run"; `before`
        // inside one means "before its begin run". Containers first, so a field whose result
        // wraps its runs in a hyperlink (every TOC entry does) is still stepped out of.
        var afterBeforeFieldStep = after;
        var beforeBeforeFieldStep = before;
        if (after is not null) after = FieldEndRunOrSelf(StepOutOfContainersAtEnd(after, element));
        if (before is not null) before = FieldBeginRunOrSelf(StepOutOfContainersAtStart(before, element));

        // Coalesce into the adjacent run rather than fragmenting the paragraph. Appending "foo"
        // after a run should extend that run's text, not drop a sibling run beside it: a separate
        // run whose text has a leading or trailing space forces the converter to emit that space
        // as &#160; (a run boundary is where HTML would otherwise collapse it), so a plain typed
        // space came back as a non-breaking one. Merge only into a run we did NOT step out of a
        // field to reach (its identity is unchanged by the step above), that is a direct child of
        // the paragraph holding only text, and only with tracked changes off — a tracked insertion
        // needs its own w:ins run. The rPr already came from this same run, so formatting matches.
        if (!splitRun && _trackedChanges != TrackedChangeMode.RenderInline)
        {
            XElement? mergeInto = null;
            bool appendAtEnd = false;
            if (after is not null && ReferenceEquals(after, afterBeforeFieldStep)
                && ReferenceEquals(after.Parent, element) && IsPlainTextRun(after))
            {
                mergeInto = after;
                appendAtEnd = true;
            }
            else if (before is not null && ReferenceEquals(before, beforeBeforeFieldStep)
                && ReferenceEquals(before.Parent, element) && IsPlainTextRun(before))
            {
                mergeInto = before;
                appendAtEnd = false;
            }

            if (mergeInto is not null)
            {
                _history.RecordPreOp(TakeSnapshot());
                try
                {
                    var text = appendAtEnd
                        ? mergeInto.Elements(W.t).Last()
                        : mergeInto.Elements(W.t).First();
                    text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    text.Value = appendAtEnd ? text.Value + replace : replace + text.Value;
                    InvalidateProjectionCache();
                    return new EditResult
                    {
                        Success = true,
                        Modified = new[] { target.Anchor },
                        Patch = PatchFor(target),
                    };
                }
                catch (Exception ex)
                {
                    return FailInternal(ex, anchorId);
                }
            }
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (splitRun) SplitRunsAtOffset(element, offset);
            var run = new XElement(W.r, rPr,
                new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), replace));
            XElement node = run;
            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                node = CreateRevisionEnvelope(
                    W.ins, _revisionAuthor ?? "docxodus", RevisionDateNow());
                node.Add(run);
            }
            UnidHelper.AssignToSelfAndDescendants(node);
            if (after is not null) after.AddAfterSelf(node);
            else if (before is not null) before.AddBeforeSelf(node);
            else if (element.Elements().FirstOrDefault(e => e.Name != W.pPr) is { } firstContent)
            {
                // No run has any text (a paragraph holding only a field with an empty cached
                // result), so offset 0 matched nothing — but it is still the START of the
                // paragraph, and "Page " typed before a PAGE field belongs before it.
                firstContent.AddBeforeSelf(node);
            }
            else element.Add(node);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// True when <paramref name="run"/> is an ordinary text run — a <c>w:r</c> carrying only its
    /// optional <c>w:rPr</c> and one or more <c>w:t</c>, with no field char, break, tab, drawing,
    /// symbol or other run content. Only such a run can safely absorb inserted text by extending
    /// a <c>w:t</c>; anything else keeps the insertion in its own sibling run.
    /// </summary>
    private static bool IsPlainTextRun(XElement run) =>
        run.Name == W.r
        && run.Elements().All(e => e.Name == W.rPr || e.Name == W.t)
        && run.Elements(W.t).Any();

    // SplitRunsAtOffset keeps non-text children on the prefix, ahead of its text.
    // Leading tabs are safe there; a tab after text or other run content could move.
    private static bool CanSplitTextRun(XElement run) =>
        run.Name == W.r && run.Elements(W.t).Any()
        && run.Elements().Where(e => e.Name != W.rPr)
            .SkipWhile(e => e.Name == W.tab).All(e => e.Name == W.t);

    // Complex fields can cross paragraphs. Text boxes, notes and comments have separate stories.
    private static bool IsInsideComplexField(XElement run)
    {
        var root = run.Ancestors().FirstOrDefault(e => e.Name == W.txbxContent
            || e.Name == W.footnote || e.Name == W.endnote || e.Name == W.comment)
            ?? run.Ancestors().Last();
        int depth = 0;
        foreach (var element in root.DescendantsTrimmed(e =>
            e.Name == W.txbxContent || e.Name == W.del || e.Name == W.moveFrom))
        {
            if (ReferenceEquals(element, run)) break;
            if (element.Name != W.fldChar) continue;
            var kind = (string?)element.Attribute(W.fldCharType);
            if (kind == "begin") depth++;
            else if (kind == "end" && depth > 0) depth--;
        }
        return depth > 0;
    }

    /// <summary>If <paramref name="run"/> sits inside a complex field (between a <c>begin</c> and
    /// its <c>end</c> <c>w:fldChar</c> among its siblings), the field's <c>end</c> run; else the run.</summary>
    private static XElement FieldEndRunOrSelf(XElement run)
    {
        int depth = 0;
        foreach (var sibling in run.ElementsBeforeSelf(W.r))
        {
            var kind = (string?)sibling.Element(W.fldChar)?.Attribute(W.fldCharType);
            if (kind == "begin") depth++;
            else if (kind == "end" && depth > 0) depth--;
        }
        var self = (string?)run.Element(W.fldChar)?.Attribute(W.fldCharType);
        if (self == "begin") depth++;
        else if (self == "end" && depth > 0) depth--;
        if (depth == 0) return run;
        foreach (var sibling in run.ElementsAfterSelf(W.r))
        {
            var kind = (string?)sibling.Element(W.fldChar)?.Attribute(W.fldCharType);
            if (kind == "begin") depth++;
            else if (kind == "end" && --depth == 0) return sibling;
        }
        return run;
    }

    /// <summary>Range markers and proofing marks — siblings that do not make a run anything but
    /// the container's first or last content.</summary>
    private static bool IsInlineMarker(XElement e) =>
        e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd
        || e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd
        || e.Name == W.permStart || e.Name == W.permEnd
        || e.Name == W.proofErr;

    /// <summary>Climb from <paramref name="node"/> through every <see cref="BoundaryContainers"/>
    /// ancestor (short of <paramref name="paragraph"/>) in which it is the last content, so an
    /// insert "after" it lands after the container.</summary>
    private static XElement StepOutOfContainersAtEnd(XElement node, XElement paragraph)
    {
        while (node.Parent is { } parent && !ReferenceEquals(parent, paragraph)
            && BoundaryContainers.Contains(parent.Name)
            && !node.ElementsAfterSelf().Any(e => !IsInlineMarker(e)))
            node = parent;
        return node;
    }

    /// <summary>The mirror of <see cref="StepOutOfContainersAtEnd"/> for an insert "before"
    /// the container's first content.</summary>
    private static XElement StepOutOfContainersAtStart(XElement node, XElement paragraph)
    {
        while (node.Parent is { } parent && !ReferenceEquals(parent, paragraph)
            && BoundaryContainers.Contains(parent.Name)
            && !node.ElementsBeforeSelf().Any(e => !IsInlineMarker(e)))
            node = parent;
        return node;
    }

    /// <summary>The mirror of <see cref="FieldEndRunOrSelf"/>: the field's <c>begin</c> run when
    /// <paramref name="run"/> sits inside one, else the run.</summary>
    private static XElement FieldBeginRunOrSelf(XElement run)
    {
        int depth = 0;
        foreach (var sibling in run.ElementsAfterSelf(W.r))
        {
            var kind = (string?)sibling.Element(W.fldChar)?.Attribute(W.fldCharType);
            if (kind == "end") depth++;
            else if (kind == "begin" && depth > 0) depth--;
        }
        var self = (string?)run.Element(W.fldChar)?.Attribute(W.fldCharType);
        if (self == "end") depth++;
        else if (self == "begin" && depth > 0) depth--;
        if (depth == 0) return run;
        foreach (var sibling in run.ElementsBeforeSelf(W.r).Reverse())
        {
            var kind = (string?)sibling.Element(W.fldChar)?.Attribute(W.fldCharType);
            if (kind == "end") depth++;
            else if (kind == "begin" && --depth == 0) return sibling;
        }
        return run;
    }

    /// <summary>
    /// Enumerate the template placeholders in the document. A thin classifier over
    /// <see cref="Grep"/> that distinguishes <c>[___]</c> value blanks, <c>[bracketed
    /// alternative clauses]</c>, and <c>[insert X]</c> / <c>[*italic hint*]</c>
    /// instruction placeholders — the three families a template-filling agent treats
    /// differently. See <see cref="PlaceholderKind"/> for the taxonomy.
    /// </summary>
    /// <remarks>
    /// Nested brackets resolve to the INNERMOST bracket. A construct like
    /// <c>[under the name [Bluth Co.]]</c> produces a placeholder for the inner
    /// <c>[Bluth Co.]</c> only — usually what an agent cares about — but the outer
    /// optional-clause bracket isn't reported separately. Use <see cref="Grep"/> with
    /// a balanced-bracket regex if you need both.
    /// </remarks>
    public IReadOnlyList<TemplatePlaceholder> FindPlaceholders(
        PlaceholderKinds kinds = PlaceholderKinds.All,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 80,
        ContextBoundary boundary = ContextBoundary.Char,
        PageCitationRequest? citationRequest = null)
    {
        ThrowIfDisposed();
        if (kinds == 0) return Array.Empty<TemplatePlaceholder>();

        // Single bracket-or-dollar-bracket scan; classify by content after the match.
        // Non-greedy inner content + negated character class keeps the regex from
        // crossing into a sibling bracket pair on the same line.
        var matches = Grep(@"\$?\[[^\[\]]+\]",
            System.Text.RegularExpressions.RegexOptions.None, scope,
            contextChars, WhitespaceMode.Preserve, boundary, citationRequest);
        var results = new List<TemplatePlaceholder>(matches.Count);
        foreach (var m in matches)
        {
            var (classified, alternatives) = Classify(m.Text);
            if (classified is not PlaceholderKind kind) continue;
            if (!kinds.HasFlag(KindToFlag(kind))) continue;
            results.Add(new TemplatePlaceholder
            {
                Match = m,
                Kind = kind,
                Hint = kind == PlaceholderKind.Instruction ? ExtractHint(m.Text) : null,
                AlternativeKinds = alternatives,
            });
        }
        return results;

        static (PlaceholderKind? Primary, IReadOnlyList<PlaceholderKind> Alternatives) Classify(string text)
        {
            var inner = text.StartsWith('$') ? text[2..^1] : text[1..^1];

            // BlankFill: 2+ underscores anywhere inside (so "[__]" director-count slots,
            // "[___ times]" unit-suffix slots, and "[________ __, 20__]" date-shaped
            // slots all qualify). Tighter than "any underscore" to avoid false positives
            // on quoted identifiers like "[a_b]". Trade-off in writeup at the FindPlaceholders
            // section of docs/architecture/docx_mutation_api.md.
            bool isBlankFill = inner.Count(c => c == '_') >= 2;

            // Instruction: italicized (asterisk-wrapped) text, or starts with the
            // drafter verbs "insert" / "specify". Conservative leading-word check
            // so general prose in brackets doesn't mis-classify.
            bool isInstruction = false;
            if (inner.StartsWith('*') && inner.EndsWith('*') && inner.Length > 2) isInstruction = true;
            else
            {
                var firstWord = inner.TakeWhile(char.IsLetter).ToArray();
                var w = new string(firstWord).ToLowerInvariant();
                if (w is "insert" or "specify") isInstruction = true;
            }

            // Secondary classification: long-clause-with-blanks. When BlankFill fires but
            // the inner text reads like a multi-word clause (4+ spaces between words),
            // the placeholder is plausibly an AlternativeClause with an embedded blank.
            // Caller can detect via AlternativeKinds and strip the outer brackets, then
            // separately fill the inner _______ run.
            bool looksClause = inner.Count(c => c == ' ') >= 4;

            // Primary classification keeps the original priority order:
            //   BlankFill → Instruction → AlternativeClause
            if (isBlankFill)
            {
                var alts = looksClause ? new[] { PlaceholderKind.AlternativeClause } : Array.Empty<PlaceholderKind>();
                return (PlaceholderKind.BlankFill, alts);
            }
            if (isInstruction)
                return (PlaceholderKind.Instruction, Array.Empty<PlaceholderKind>());
            return (PlaceholderKind.AlternativeClause, Array.Empty<PlaceholderKind>());
        }

        static string ExtractHint(string text)
        {
            var inner = text.StartsWith('$') ? text[2..^1] : text[1..^1];
            // Strip a single pair of surrounding asterisks (italic markers from the projector).
            if (inner.StartsWith('*') && inner.EndsWith('*') && inner.Length > 2)
                inner = inner[1..^1];
            return inner.Trim();
        }

        static PlaceholderKinds KindToFlag(PlaceholderKind k) => k switch
        {
            PlaceholderKind.BlankFill => PlaceholderKinds.BlankFill,
            PlaceholderKind.AlternativeClause => PlaceholderKinds.AlternativeClause,
            PlaceholderKind.Instruction => PlaceholderKinds.Instruction,
            _ => 0,
        };
    }

    /// <summary>
    /// Thin discoverability alias for <see cref="FindPlaceholders"/>. Same return
    /// shape; the rename exists because "what's remaining?" reads more naturally
    /// at agent call sites than "find the placeholders."
    /// </summary>
    public IReadOnlyList<TemplatePlaceholder> RemainingPlaceholders(
        PlaceholderKinds kinds = PlaceholderKinds.All) =>
        FindPlaceholders(kinds);

    /// <summary>
    /// Picker-driven template fill. For every placeholder matching
    /// <see cref="FillOptions.Kinds"/>, calls <paramref name="picker"/>; if the picker
    /// returns a non-null string, the placeholder is replaced (with optional
    /// <c>$</c>-prefix preservation per <see cref="FillOptions.PreserveDollarPrefix"/>).
    /// Iterates until no more placeholders match (or until <see cref="FillOptions.MaxPasses"/>
    /// is reached, or a pass makes zero state changes) — important when
    /// <see cref="FillOptions.Kinds"/> includes <see cref="PlaceholderKinds.AlternativeClause"/>
    /// and the doc has nested brackets that surface only after the inner ones are stripped.
    /// Replacements within a paragraph are applied in reverse-offset order automatically.
    /// The picker may be invoked more than once for the same logical placeholder
    /// when <see cref="FillOptions.Kinds"/> includes <see cref="PlaceholderKinds.AlternativeClause"/>
    /// and inner brackets are stripped between passes; pickers must therefore be
    /// deterministic on <c>p.Match.Text</c> (return the same result for the same
    /// input text). Non-deterministic pickers can produce inconsistent fills.
    /// </summary>
    public BulkEditResult FillPlaceholders(
        Func<TemplatePlaceholder, string?> picker,
        FillOptions? options = null)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(picker);
        var opts = options ?? new FillOptions();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(opts.MaxPasses);

        int filled = 0;
        int workPasses = 0;
        var errors = new List<EditError>();
        var unfilled = new List<TemplatePlaceholder>();
        var seenSkipKeys = new HashSet<(string AnchorId, int Start, int Length)>();

        for (int pass = 1; pass <= opts.MaxPasses; pass++)
        {
            var placeholders = FindPlaceholders(opts.Kinds, opts.Scope, opts.ContextChars, opts.Boundary)
                .OrderByDescending(p => p.Match.EnclosingAnchor.Anchor.Id, StringComparer.Ordinal)
                .ThenByDescending(p => p.Match.Span.Start)
                .ToList();
            if (placeholders.Count == 0) break;

            int passChanges = 0;
            foreach (var p in placeholders)
            {
                var pick = picker(p);
                if (pick is null)
                {
                    // Count each skip exactly once per placeholder lifetime.
                    var key = (p.Match.EnclosingAnchor.Anchor.Id, p.Match.Span.Start, p.Match.Span.Length);
                    if (seenSkipKeys.Add(key))
                        unfilled.Add(p);
                    continue;
                }

                if (opts.PreserveDollarPrefix && p.Match.Text.StartsWith("$") && !pick.StartsWith("$"))
                    pick = "$" + pick;

                var r = opts.CoalesceWhitespaceAroundEmptyFill && pick.Length == 0
                    ? ReplaceMatchCoalescingNeighbors(p.Match)
                    : ReplaceMatch(p.Match, pick);
                if (r.Success)
                {
                    filled++;
                    passChanges++;
                }
                else if (r.Error is { } err)
                {
                    errors.Add(err);
                }
            }

            // Record this pass only if it did real work — observation alone
            // (placeholders found but all skipped or all errored) doesn't count.
            if (passChanges > 0)
                workPasses = pass;

            // If this pass made no changes, the picker is steady-state — stop iterating.
            if (passChanges == 0) break;
        }

        int stillPresent = FindPlaceholders(opts.Kinds, opts.Scope).Count;

        return new BulkEditResult
        {
            Filled = filled,
            Skipped = unfilled.Count,
            StillPresent = stillPresent,
            Passes = workPasses,
            Unfilled = unfilled,
            Errors = errors,
        };
    }

    /// <summary>
    /// Helper for <see cref="FillPlaceholders"/>'s
    /// <see cref="FillOptions.CoalesceWhitespaceAroundEmptyFill"/> path: deletes the
    /// match's span and, based on the chars immediately adjacent in the enclosing
    /// block's flat text, also absorbs surrounding whitespace / leading-space-before-punctuation
    /// / matched-brackets. See the option's docs for the exact rules. Falls back
    /// to a literal <see cref="ReplaceMatch"/> with empty string when no neighbor
    /// pattern matches.
    /// </summary>
    private EditResult ReplaceMatchCoalescingNeighbors(TextMatch match)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var anchorId = match.EnclosingAnchor.Anchor.Id;
        var target = FindAnchor(anchorId);
        if (target is null) return ReplaceMatch(match, string.Empty);
        var element = target.Resolve(_doc!);
        if (element is null) return ReplaceMatch(match, string.Empty);

        var flat = Internal.RunTextMap.Build(element).FlatText;
        int start = match.Span.Start;
        int end = start + match.Span.Length;
        if (start < 0 || end > flat.Length) return ReplaceMatch(match, string.Empty);

        char? leftChar = start > 0 ? flat[start - 1] : null;
        char? rightChar = end < flat.Length ? flat[end] : null;

        // Fold the Unicode whitespace variants Word documents commonly use
        // (NBSP, narrow NBSP, thin space) to ASCII space for the rules below so
        // an NBSP-on-either-side still gets coalesced like a regular space.
        static char? Fold(char? c) => c switch
        {
            ' ' or ' ' or ' ' => ' ',
            _ => c,
        };
        char? l = Fold(leftChar);
        char? r = Fold(rightChar);

        static bool IsAsciiSpace(char? c) => c is ' ' or '\t';
        static bool IsClauseTerminator(char? c) => c is '.' or ',' or ';' or ':' or '!' or '?';
        static bool IsOpenBracket(char? c) => c is '(' or '[' or '{';
        static bool IsCloseBracket(char? c) => c is ')' or ']' or '}';

        int extendLeft = 0;
        int extendRight = 0;

        if (IsAsciiSpace(l) && IsAsciiSpace(r))
        {
            // " [x] " → consume the trailing space, leaving one space.
            extendRight = 1;
        }
        else if (IsAsciiSpace(l) && IsClauseTerminator(r))
        {
            // " [x]." / " [x]," → drop the leading space.
            extendLeft = 1;
        }
        else if (IsOpenBracket(l) && IsCloseBracket(r))
        {
            // "([x])" / "[[x]]" → drop both surrounding brackets.
            extendLeft = 1;
            extendRight = 1;
        }

        if (extendLeft == 0 && extendRight == 0)
            return ReplaceMatch(match, string.Empty);

        return ReplaceTextAtSpan(
            anchorId,
            start - extendLeft,
            match.Span.Length + extendLeft + extendRight,
            string.Empty);
    }

    /// <summary>
    /// Apply <paramref name="match"/>'s fragment list to the live element, inserting
    /// <paramref name="replace"/> into the first fragment's run and removing each
    /// subsequent fragment's slice from its run (preserving each run's rPr).
    /// </summary>
    private static void ApplyFragmentReplacement(XElement blockElement, TextMatch match, string replace)
    {
        if (match.Fragments.Count == 0) return;

        // Build a unid → XElement run lookup once. The run XElements are the live
        // descendants of `blockElement` (walking hyperlink/sdt containers too).
        var runsByUnid = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var run in InlineRuns(blockElement))
        {
            var unid = (string?)run.Attribute(PtOpenXml.Unid);
            if (unid is not null) runsByUnid[unid] = run;
        }

        for (int i = 0; i < match.Fragments.Count; i++)
        {
            var fragment = match.Fragments[i];
            if (!runsByUnid.TryGetValue(fragment.Unid, out var run)) continue;

            var concat = RunText(run);
            var start = fragment.SpanInElement.Start;
            var len = fragment.SpanInElement.Length;
            if (start < 0 || start + len > concat.Length) continue;

            var before = concat.Substring(0, start);
            var after = concat.Substring(start + len);
            var newText = i == 0 ? before + replace + after : before + after;

            // Collapse all w:t descendants in this run into a single w:t with the new text.
            // Loses any inline <w:tab/>/<w:br/> inside the run's text section — they're rare
            // for placeholder slots and supporting them here would balloon the impl. Run's
            // rPr/proofErr siblings are untouched, which is the formatting-preservation contract.
            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                newText));
        }
    }

    /// <summary>
    /// Tracked-change counterpart to <see cref="ApplyFragmentReplacement"/>. Each selected
    /// run slice becomes a formatting-preserving <c>w:r/w:delText</c> inside one or more
    /// <c>w:del</c> envelopes, while the replacement is a single <c>w:r/w:t</c> inside
    /// <c>w:ins</c> and inherits the first affected run's <c>w:rPr</c>. Runs are split at
    /// the selection boundaries so untouched prefixes/suffixes remain ordinary text.
    /// Zero-width siblings (bookmarks, comment ranges, note-reference runs, proofing markers)
    /// are never moved into revision envelopes; when one separates selected runs it simply
    /// splits the deletion into adjacent Word-style envelopes around that marker.
    /// </summary>
    private void ApplyFragmentReplacementTracked(
        XElement blockElement,
        TextMatch match,
        string replace,
        string author,
        string date)
    {
        if (match.Fragments.Count == 0) return;

        var runsByUnid = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var run in InlineRuns(blockElement))
        {
            var unid = (string?)run.Attribute(PtOpenXml.Unid);
            if (unid is not null) runsByUnid[unid] = run;
        }

        var deletedRuns = new List<XElement>(match.Fragments.Count);
        XElement? insertedRun = null;

        for (int i = 0; i < match.Fragments.Count; i++)
        {
            var fragment = match.Fragments[i];
            if (!runsByUnid.TryGetValue(fragment.Unid, out var run))
                throw new InvalidOperationException(
                    $"tracked replacement run no longer exists: {fragment.Unid}");

            var concat = RunText(run);
            var start = fragment.SpanInElement.Start;
            var len = fragment.SpanInElement.Length;
            if (start < 0 || len <= 0 || start + len > concat.Length)
                throw new InvalidOperationException(
                    $"tracked replacement slice {start}+{len} is outside run text length {concat.Length}");

            var before = concat[..start];
            var removed = concat.Substring(start, len);
            var after = concat[(start + len)..];

            // Snapshot the original formatting before changing/reusing the live run.
            var deletedRun = CloneRunWithRevisionText(run, W.delText, removed);
            if (i == 0 && replace.Length > 0)
                insertedRun = CloneRunWithRevisionText(run, W.t, replace);

            var replacementNodes = new List<XElement>(4);
            if (before.Length > 0)
            {
                SetRunText(run, before);
                replacementNodes.Add(run);
            }
            else if (after.Length == 0 && HasNonTextRunContent(run))
            {
                // A text-bearing marker run is unusual but legal (notably note/comment
                // references). Keep its zero-width payload exactly once when all w:t text
                // was selected instead of dropping it with the old text.
                SetRunText(run, null);
                replacementNodes.Add(run);
            }

            replacementNodes.Add(deletedRun);

            if (after.Length > 0)
            {
                if (before.Length == 0)
                {
                    SetRunText(run, after);
                    replacementNodes.Add(run);
                }
                else
                {
                    replacementNodes.Add(CloneRunWithRevisionText(run, W.t, after));
                }
            }

            run.ReplaceWith(replacementNodes);
            deletedRuns.Add(deletedRun);
        }

        // Coalesce adjacent deleted slices under the same inline parent. This is the
        // shape Word emits for a replacement spanning differently formatted runs:
        // one w:del containing each original rPr-bearing run, followed by one w:ins.
        // A semantic marker or container boundary naturally starts another envelope.
        var deletionEnvelopes = new List<XElement>();
        XElement? currentEnvelope = null;
        foreach (var deletedRun in deletedRuns)
        {
            if (currentEnvelope is not null
                && ReferenceEquals(currentEnvelope.Parent, deletedRun.Parent)
                && ReferenceEquals(deletedRun.PreviousNode, currentEnvelope))
            {
                deletedRun.Remove();
                currentEnvelope.Add(deletedRun);
                continue;
            }

            currentEnvelope = CreateRevisionEnvelope(W.del, author, date);
            deletedRun.AddBeforeSelf(currentEnvelope);
            deletedRun.Remove();
            currentEnvelope.Add(deletedRun);
            deletionEnvelopes.Add(currentEnvelope);
        }

        if (insertedRun is null || deletionEnvelopes.Count == 0) return;

        // Keep a replacement wholly inside its original hyperlink/SDT/smartTag/fldSimple.
        // When every deletion envelope has that same parent, put the insertion after the
        // final deletion (the canonical Word replacement order). A cross-container match
        // inserts beside the first slice so it retains the first run's inline semantics.
        var firstEnvelope = deletionEnvelopes[0];
        var insertionAnchor = deletionEnvelopes.All(e => ReferenceEquals(e.Parent, firstEnvelope.Parent))
            ? deletionEnvelopes[^1]
            : firstEnvelope;
        var insertionEnvelope = CreateRevisionEnvelope(W.ins, author, date);
        insertionEnvelope.Add(insertedRun);
        insertionAnchor.AddAfterSelf(insertionEnvelope);
    }

    /// <summary>Create native revision markup and keep Word's document-level recording flag in
    /// sync. The flag does not make existing revisions render; it tells Word to track subsequent
    /// interactive edits after the generated document is opened.</summary>
    /// <summary>The author/date pair one tracked operation stamps on every mark it writes.
    /// Word stamps one action with one time; the registry relies on that when it folds a
    /// wrapper's payload marks into the wrapper's own envelope entry, so a stamp is taken
    /// once per operation and threaded through, never re-read from the clock per element.</summary>
    private readonly record struct RevisionStamp(string Author, string Date);

    private RevisionStamp NewRevisionStamp() => new(_revisionAuthor ?? "docxodus", RevisionDateNow());

    /// <summary>The <c>w:date</c> of a revision recorded now, formatted under the invariant
    /// culture. A culture-sensitive format swaps the time separator (fi-FI writes
    /// <c>14.27.09</c>) or the calendar (th-TH writes year 2569), neither of which is the
    /// <c>xsd:dateTime</c> the schema and every reader expect.</summary>
    private static string RevisionDateNow() => Internal.CommentOps.FormatDate(DateTime.UtcNow);

    private XElement CreateRevisionEnvelope(
        XName name, RevisionStamp stamp, params object[] content) =>
        CreateRevisionEnvelope(name, stamp.Author, stamp.Date, content);

    private XElement CreateRevisionEnvelope(
        XName name, string author, string date, params object[] content)
    {
        EnsureTrackRevisionsEnabled();
        var envelope = new XElement(name,
            new XAttribute(W.id, NextRevisionId()),
            new XAttribute(W.author, author),
            new XAttribute(W.date, date));
        envelope.Add(content);
        return envelope;
    }

    private void EnsureTrackRevisionsEnabled()
    {
        var main = _doc!.MainDocumentPart
            ?? throw new InvalidOperationException("document has no main document part");
        var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        var xDoc = settingsPart.GetXDocument();
        var root = xDoc.Root;
        if (root is null)
        {
            root = new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w));
            xDoc.Add(root);
        }

        if (root.Element(W.trackRevisions) is { } existing)
        {
            // Bare CT_OnOff is the canonical enabled form. In particular, do not leave an
            // inherited w:val="false" in place after this session has emitted a revision.
            if (existing.Attribute(W.val) is { } disabledOrExplicit)
            {
                disabledOrExplicit.Remove();
                settingsPart.PutXDocument();
            }
            return;
        }

        if (WordprocessingMLUtil.EnsureSettingsChildInOrder(
                root, new XElement(W.trackRevisions)))
            settingsPart.PutXDocument();
    }

    /// <summary>Create a text-only run that retains the source run's formatting/rsid
    /// attributes but gets its own internal Unid. <paramref name="textName"/> is
    /// <c>w:t</c> for ordinary/inserted text and <c>w:delText</c> for deletions.</summary>
    private static XElement CloneRunWithRevisionText(XElement source, XName textName, string text)
    {
        var clone = new XElement(W.r,
            source.Attributes()
                .Where(a => a.Name != PtOpenXml.Unid)
                .Select(a => new XAttribute(a)),
            source.Element(W.rPr) is { } rPr ? new XElement(rPr) : null,
            new XElement(textName,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                text));
        UnidHelper.AssignToSelfAndDescendants(clone);
        return clone;
    }

    /// <summary>Replace a run's direct w:t sequence with one text node while retaining
    /// rPr and every non-text child in place. Null removes text without adding an empty
    /// node (used to preserve a zero-width marker run).</summary>
    private static void SetRunText(XElement run, string? text)
    {
        var textNodes = run.Elements(W.t).ToList();
        if (textNodes.Count > 0)
        {
            if (text is null)
            {
                foreach (var t in textNodes) t.Remove();
                return;
            }

            var replacement = new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                text);
            textNodes[0].ReplaceWith(replacement);
            foreach (var t in textNodes.Skip(1)) t.Remove();
            return;
        }

        if (text is not null)
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                text));
    }

    private static bool HasNonTextRunContent(XElement run) =>
        run.Elements().Any(e => e.Name != W.rPr && e.Name != W.t);

    /// <summary>
    /// When <see cref="DocxSessionSettings.SmartQuotes"/> is on, replace ASCII <c>"</c>
    /// and <c>'</c> with typographic curly quotes. Heuristic: open quote at the start
    /// of the string, after whitespace, or after an open-bracket-like character;
    /// close quote everywhere else. 1:1 character substitution preserves offsets so
    /// downstream span math stays correct.
    /// </summary>
    private string MaybeApplySmartQuotes(string text)
    {
        if (!_settings.SmartQuotes || string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c != '"' && c != '\'') { sb.Append(c); continue; }

            // Look at the previous character (default to "start of string" = whitespace).
            char prev = i == 0 ? ' ' : text[i - 1];
            bool open = char.IsWhiteSpace(prev) || prev is '(' or '[' or '{' or '<';

            sb.Append(c switch
            {
                '"' => open ? '“' : '”',
                '\'' => open ? '‘' : '’',
                _ => c,
            });
        }
        return sb.ToString();
    }

    /// <summary>
    /// Maps the Unicode whitespace variants Word documents commonly use (NBSP, narrow
    /// NBSP, thin space) to ASCII space. Each substitution is one-character-for-one,
    /// so character offsets in the result map 1:1 to the input.
    /// </summary>
    private static string NormalizeWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length);
        foreach (var c in text)
        {
            sb.Append(c switch
            {
                ' ' => ' ', // non-breaking space
                ' ' => ' ', // narrow no-break space
                ' ' => ' ', // thin space
                _ => c,
            });
        }
        return sb.ToString();
    }

    /// <summary>
    /// Walks outward from a match span by character, stopping at either the
    /// <c>contextChars</c> cap or the nearest character that qualifies as a
    /// boundary under <paramref name="boundary"/>. Returns the <c>(before, after)</c>
    /// text slices. Used by both <see cref="Grep"/> and <see cref="GrepCrossBlock"/>.
    /// </summary>
    private static (string Before, string After) WalkContext(
        string text, int matchStart, int matchLength, int contextChars, ContextBoundary boundary)
    {
        int matchEnd = matchStart + matchLength;

        int leftCap = Math.Max(0, matchStart - contextChars);
        int leftStop = matchStart;
        while (leftStop > leftCap)
        {
            if (IsBoundary(text[leftStop - 1], boundary)) break;
            leftStop--;
        }

        int rightCap = Math.Min(text.Length, matchEnd + contextChars);
        int rightStop = matchEnd;
        while (rightStop < rightCap)
        {
            if (IsBoundary(text[rightStop], boundary)) break;
            rightStop++;
        }

        return (text.Substring(leftStop, matchStart - leftStop),
                text.Substring(matchEnd, rightStop - matchEnd));
    }

    private static bool IsBoundary(char c, ContextBoundary mode) => mode switch
    {
        ContextBoundary.Char => false,
        ContextBoundary.Bracket => c is '[' or ']',
        ContextBoundary.Sentence => c is '.' or '!' or '?' or ':' or ';',
        ContextBoundary.Comma => c is ',',
        _ => false,
    };

    private static bool ScopeMatches(string anchorScope, ProjectionScopes filter)
    {
        // Anchor scopes are strings ("body", "hdr1", "ftr2", "fn", "en", "cmt").
        // ProjectionScopes is a flags enum over the same categories.
        if (anchorScope == "body") return filter.HasFlag(ProjectionScopes.Body);
        if (anchorScope.StartsWith("hdr", StringComparison.Ordinal)) return filter.HasFlag(ProjectionScopes.Headers);
        if (anchorScope.StartsWith("ftr", StringComparison.Ordinal)) return filter.HasFlag(ProjectionScopes.Footers);
        if (anchorScope == "fn") return filter.HasFlag(ProjectionScopes.Footnotes);
        if (anchorScope == "en") return filter.HasFlag(ProjectionScopes.Endnotes);
        if (anchorScope == "cmt") return filter.HasFlag(ProjectionScopes.Comments);
        return false;
    }

    private OpenXmlPart? ResolvePart(string partUri) =>
        EnumerateProjectedParts().FirstOrDefault(p => p.Uri.ToString() == partUri)
        ?? RevisionStoryParts().Select(story => story.Part)
            .FirstOrDefault(part => part.Uri.ToString() == partUri);

    private static RunFormatting ExtractFormatting(XElement run, OpenXmlPart? ownerPart)
    {
        var rPr = run.Element(W.rPr);
        string? hyperlinkUrl = null;
        for (var p = run.Parent; p is not null; p = p.Parent)
        {
            if (p.Name == W.hyperlink)
            {
                var rid = (string?)p.Attribute(R.id);
                if (!string.IsNullOrEmpty(rid) && ownerPart is not null)
                {
                    var rel = ownerPart.HyperlinkRelationships.FirstOrDefault(x => x.Id == rid);
                    if (rel is not null) hyperlinkUrl = rel.Uri.ToString();
                }
                break;
            }
        }

        return new RunFormatting
        {
            Bold = rPr?.Element(W.b) is not null,
            Italic = rPr?.Element(W.i) is not null,
            Underline = rPr?.Element(W.u) is not null,
            Strike = rPr?.Element(W.strike) is not null,
            Code = string.Equals((string?)rPr?.Element(W.rStyle)?.Attribute(W.val), "Code", StringComparison.Ordinal),
            Color = (string?)rPr?.Element(W.color)?.Attribute(W.val),
            HyperlinkUrl = hyperlinkUrl,
            RunStyle = (string?)rPr?.Element(W.rStyle)?.Attribute(W.val),
        };
    }

    // ─── Tier A: text CRUD ────────────────────────────────────────────────

    public EditResult ReplaceText(string anchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceText requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline
            && element.Descendants().Any(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
                "tracked whole-paragraph replacement containing bookmark markers is unsupported; use a surgical span replacement or switch recording mode",
                anchorId);
        if (_trackedChanges == TrackedChangeMode.RenderInline
            && FirstUnrecordableInlineContent(new[] { element }) is { } unrecordable)
            return EditResult.Fail(EditErrorCode.IncompatibleElementType,
                $"Tracked whole-paragraph replacement does not support {unrecordable} inside a paragraph; no changes were made.",
                anchorId);
        // Strip a leading auto-number prefix from the payload before parsing. The
        // projector emits "## Fourth The total number…" — auto-number from numPr
        // plus a space separator plus the run text — so an agent that echoes the
        // visible heading back as its replacement payload otherwise gets the
        // prefix applied twice (Word renders the auto-number AND the run text now
        // begins with "Fourth"). See DS091.
        markdownPayload = StripResolvedAutoNumberPrefix(element, markdownPayload);
        markdownPayload = MaybeApplySmartQuotes(markdownPayload);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
        // The anchor addresses exactly one block. Truncating a multi-block payload to
        // its first block and reporting success was silent data loss (#570).
        if (parsed.Blocks.Count > 1)
            return EditResult.Fail(EditErrorCode.UnsupportedMarkdownSyntax,
                $"payload parses to {parsed.Blocks.Count} markdown blocks but ReplaceText replaces exactly one; " +
                "join lines with soft breaks (two trailing spaces before the newline), or add blocks with InsertParagraph",
                anchorId);
        if (ValidatePendingHyperlinks(parsed.Blocks.SelectMany(b => b.RunElements), anchorId) is { } linkError)
            return linkError;

        // An UNESCAPED block marker declares block semantics under the projector-symmetric
        // contract (the projector emits literal hashes escaped, as "\#\# x"), so honor it the
        // way InsertParagraph does instead of consuming the marker and applying nothing (#570).
        // A plain payload leaves the paragraph mark untouched, exactly as before.
        var declaredStyle = parsed.Blocks.Count == 1
            ? DeclaredBlockStyleId(parsed.Blocks[0].Kind)
            : null;
        // A list marker on a paragraph that is not yet a list item promotes it, exactly as
        // ApplyListFormat would; on an existing list item the marker is the projection's own
        // spelling echoed back and the numbering is left alone.
        var declaredList = parsed.Blocks.Count == 1
            && parsed.Blocks[0].Kind
                is Internal.ParserBlockKind.BulletItem or Internal.ParserBlockKind.OrderedItem
            && ((int?)element.Element(W.pPr)?.Element(W.numPr)?.Element(W.numId)?.Attribute(W.val) ?? 0) == 0
            && ResolveStyleNumbering(element).numId is null
            ? parsed.Blocks[0]
            : null;
        XElement? oldPPr = null;
        if (declaredStyle is not null || declaredList is not null)
        {
            if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;
            if (declaredStyle is not null
                && _trackedChanges == TrackedChangeMode.RenderInline
                && !Internal.StyleFactory.HasParagraphStyle(_doc!, declaredStyle))
                return TrackedStructureUnsupported(
                    $"ReplaceText payload requiring synthesis of style '{declaredStyle}'", anchorId);
        }

        var hyperlinkOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, element);
        var oldHyperlinkIds = element.Descendants(W.hyperlink)
            .Select(h => (string?)h.Attribute(R.id)).Where(id => !string.IsNullOrEmpty(id)).Cast<string>().ToList();

        // Snapshot before style synthesis, so undo/rollback restores the styles part too
        // (same ordering as SetParagraphStyle).
        var preOp = TakeSnapshot();
        if (declaredStyle is not null)
        {
            if (!Internal.StyleFactory.EnsureParagraphStyle(_doc!, declaredStyle))
                return EditResult.Fail(EditErrorCode.UnknownStyle, $"style id not found: {declaredStyle}", anchorId);
        }
        if (declaredStyle is not null || declaredList is not null)
            oldPPr = new XElement(element.Element(W.pPr) ?? new XElement(W.pPr));

        // A replacement can drop a note's only reference: an untracked one removes a run that
        // mixes text with the reference, a tracked one un-inserts such a run inside the author's
        // own insertion. Baseline the referenced notes so the orphaned definition goes too.
        var referencedNotesBefore = element.Descendants()
            .Any(d => d.Name == W.footnoteReference || d.Name == W.endnoteReference)
            ? ReferencedNoteIds()
            : ((HashSet<int> Footnotes, HashSet<int> Endnotes)?)null;
        _history.RecordPreOp(preOp);
        try
        {
            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                ApplyReplaceTextTracked(element, parsed.Blocks);
            }
            else
            {
                ApplyReplaceTextAccept(element, parsed.Blocks);
            }
            var removed = new List<Anchor>();
            if (referencedNotesBefore is { } notesBefore)
                AppendPrunedNoteAnchors(PruneOrphanedNotes(notesBefore), removed, new HashSet<string>(StringComparer.Ordinal));
            if (declaredStyle is not null)
            {
                var pPr = element.Element(W.pPr);
                if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
                pPr.Element(W.pStyle)?.Remove();
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, declaredStyle)));
                // Same suppressor rule as BuildParagraphFromParsedBlock, so a heading
                // authored through ReplaceText carries the identical paragraph mark (#572).
                if (HeadingNumberingSuppressor(declaredStyle) is { } suppressor)
                    SetPPrChildInOrder(pPr, suppressor);
                if (_trackedChanges == TrackedChangeMode.RenderInline)
                    TrackPropertyMutation(pPr, oldPPr!, W.pPrChange,
                        _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate(), W.rPr, W.sectPr);
            }
            if (declaredList is not null)
            {
                AssignPayloadListNumbering(element, declaredList, new PayloadListState(parsed.Blocks)
                {
                    PrecedingNeighbor = element.ElementsBeforeSelf().LastOrDefault(),
                    FollowingNeighbor = element.ElementsAfterSelf().FirstOrDefault(),
                });
                TrackListPropertyMutation(element, oldPPr!, insertedNumPr: true);
            }
            PromoteHyperlinkRelationships(element);
            if (hyperlinkOwner is { } owner)
            {
                foreach (var relationshipId in oldHyperlinkIds)
                    Internal.OwnedPartRelationships.DeleteReferenceRelationshipIfOrphaned(owner.Part, relationshipId, R.id);
                Internal.OwnedPartRelationships.SweepOrphanedImages(owner.Part);
            }

            if (target.Anchor.Scope == "cmt") CommentsVersion++;
            InvalidateProjectionCache();
            // A declared style or list can flip the anchor kind (p → h, p → li); report the
            // fresh identity.
            var updated = declaredStyle is not null || declaredList is not null
                ? AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor
                : target.Anchor;
            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Removed = removed,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// The paragraph style a parsed markdown block explicitly declares — the same mapping
    /// <see cref="BuildParagraphFromParsedBlock"/> writes for InsertParagraph. Null for plain
    /// paragraphs and for list items, whose declaration is numbering rather than a style
    /// (see <see cref="AssignPayloadListNumbering"/>).
    /// </summary>
    private static string? DeclaredBlockStyleId(Internal.ParserBlockKind kind) => kind switch
    {
        Internal.ParserBlockKind.Heading1 or Internal.ParserBlockKind.Heading2
            or Internal.ParserBlockKind.Heading3 or Internal.ParserBlockKind.Heading4
            or Internal.ParserBlockKind.Heading5 or Internal.ParserBlockKind.Heading6 =>
            $"Heading{(int)kind - (int)Internal.ParserBlockKind.Heading1 + 1}",
        Internal.ParserBlockKind.Quote => "Quote",
        Internal.ParserBlockKind.Code => "Code",
        _ => null,
    };

    public EditResult DeleteBlock(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li" or "tbl" or "fn" or "en" or "cmt"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteBlock requires a block-level/footnote/endnote/comment anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        // Word reserves the TYPED footnote/endnote definitions (separator, continuationSeparator,
        // continuationNotice) for page-rendering scaffolding; they carry no user content and
        // removing one corrupts the document. Same predicate the projector filters on, so the two
        // can't drift over which types count as reserved.
        if (target.Anchor.Kind is "fn" or "en" && WmlToMarkdownConverter.IsBoilerplateNote(element))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"cannot delete a Word-reserved {target.Anchor.Kind} of type='{(string?)element.Attribute(W.type)}'",
                anchorId);

        // Use the range deleter for exactly this block so recording includes paragraph
        // marks and table rows, preserves anchors, and shares its pre-mutation guards.
        // fn/en/cmt are definitions in their own parts and retain structural deletion.
        if (_trackedChanges == TrackedChangeMode.RenderInline
            && target.Anchor.Kind is "p" or "h" or "li" or "tbl")
            return DeleteSiblingRangeCore(target, element, element.ElementsAfterSelf().FirstOrDefault());

        if (ValidateBookmarkRemoval(new[] { element }, anchorId) is { } bookmarkError)
            return bookmarkError;

        var hyperlinkOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, element);
        var referencedNotesBefore = ReferencedNoteIds();
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // For fn/en/cmt: also remove every cross-reference (footnoteReference,
            // endnoteReference, commentReference/RangeStart/RangeEnd) anywhere in
            // the package that points at this definition's id. Otherwise Word
            // renders broken superscript references for the orphaned ids.
            if (target.Anchor.Kind is "fn" or "en" or "cmt")
            {
                var elementId = (string?)element.Attribute(W.id);
                if (!string.IsNullOrEmpty(elementId))
                    RemoveCrossReferences(target.Anchor.Kind, elementId);

                // For comments, also prune Word's threading metadata (commentsExtended /
                // commentsIds entries keyed by the definition paragraphs' w14:paraId) so a
                // removed comment leaves no dangling reply/resolve state. Lives here — not in
                // RemoveComment — so the generic DeleteBlock path gets it too.
                if (target.Anchor.Kind == "cmt")
                {
                    var paraIds = element.Elements(W.p)
                        .Select(p => (string?)p.Attribute(W14.paraId))
                        .Where(pid => !string.IsNullOrEmpty(pid))
                        .Select(pid => pid!)
                        .ToList();
                    Internal.CommentOps.PruneThreadingMetadata(_doc!, paraIds);
                }
            }

            // Collect descendant anchors before removal so the caller knows what's gone.
            var index = AnchorIndex();
            var removed = new List<Anchor> { target.Anchor };
            foreach (var d in element.Descendants())
            {
                var unid = (string?)d.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                foreach (var kv in index)
                {
                    if (kv.Value.Unid == unid && kv.Value.Unid != target.Unid)
                        removed.Add(kv.Value.Anchor);
                }
            }
            element.Remove();
            AppendPrunedNoteAnchors(
                PruneOrphanedNotes(referencedNotesBefore),
                removed,
                new HashSet<string>(removed.Select(a => a.Id), StringComparer.Ordinal));
            if (hyperlinkOwner is { } owner)
                SweepOrphanedStoryRelationships(owner.Part);
            if (target.Anchor.Kind == "cmt") CommentsVersion++;
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Removed = removed,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// Deletes every top-level block-level element between <paramref name="fromAnchorId"/>
    /// (inclusive) and <paramref name="toAnchorIdExclusive"/> (exclusive) in document order.
    /// Both anchors must be block-level kinds (<c>p</c>, <c>h</c>, <c>li</c>, <c>tbl</c>),
    /// live in the same package part, and share a direct parent (no spanning into table
    /// cells or other nested containers). Records a single undo snapshot so
    /// <see cref="Undo"/> restores the entire range together.
    /// </summary>
    /// <remarks>
    /// In <see cref="TrackedChangeMode.RenderInline"/>, each paragraph in the range has
    /// its runs wrapped in <c>w:del</c> and its paragraph-mark marked deleted via
    /// <c>w:pPr/w:rPr/w:del</c>; each table row gets a <c>w:trPr/w:del</c> marker with
    /// its cell paragraphs wrapped recursively. Anchors stay live (<see cref="EditResult.Modified"/>
    /// instead of <see cref="EditResult.Removed"/>) so callers can re-address the same
    /// blocks before changes are accepted. Block-level <c>w:sdt</c> content controls and
    /// <c>w:customXml</c> wrappers use paired <c>w:customXmlDelRangeStart</c>/<c>End</c>
    /// ranges for their envelopes plus recursively tracked payload blocks (issues #473,
    /// #764). Locked and data-bound controls use the same shape: their metadata — and a
    /// custom-XML wrapper's <c>w:customXmlPr</c> — remains untouched until the revision is
    /// resolved. A paragraph containing run-level <c>w:customXml</c> is rejected with
    /// <see cref="EditErrorCode.IncompatibleElementType"/> before mutation because the
    /// paragraph deleter marks only direct-child runs and would leave that wrapper's text
    /// undeleted. Any other structural fall-through is reported in
    /// <see cref="EditResult.Removed"/> rather than silently disappearing.
    /// </remarks>
    public EditResult DeleteRange(string fromAnchorId, string toAnchorIdExclusive)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var fromTarget = FindAnchor(fromAnchorId);
        if (fromTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"from anchor not found: {fromAnchorId}", fromAnchorId);
        var toTarget = FindAnchor(toAnchorIdExclusive);
        if (toTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"to anchor not found: {toAnchorIdExclusive}", toAnchorIdExclusive);

        // Scope (package-part) check first — different parts can't form a contiguous
        // sibling range under any circumstance, even if the kinds look block-level.
        if (fromTarget.Anchor.Scope != toTarget.Anchor.Scope)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                $"DeleteRange anchors must live in the same package part; from={fromTarget.Anchor.Scope} to={toTarget.Anchor.Scope}",
                fromAnchorId);

        if (fromTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteRange requires block-level anchors; from kind={fromTarget.Anchor.Kind}", fromAnchorId);
        if (toTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteRange requires block-level anchors; to kind={toTarget.Anchor.Kind}", toAnchorIdExclusive);

        var fromElement = fromTarget.Resolve(_doc!);
        var toElement = toTarget.Resolve(_doc!);
        if (fromElement is null || toElement is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", fromAnchorId);
        if (fromElement.Parent != toElement.Parent)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "DeleteRange anchors must share a direct parent (no spanning into nested containers)",
                fromAnchorId);

        return DeleteSiblingRangeCore(fromTarget, fromElement, toElement);
    }

    /// <summary>
    /// Deletes a heading and every block-level sibling under it, up to (but not including)
    /// the next heading at the same or higher level. If no such next heading exists, the
    /// section extends to the end of the parent (the heading and everything after it).
    /// </summary>
    /// <param name="headingAnchorId">Anchor id of the heading paragraph (kind must be <c>h</c>).</param>
    /// <remarks>
    /// "Level" is the same notion <see cref="WmlToMarkdownConverter"/> uses for the projection:
    /// <c>Heading1</c> = 1, <c>Heading2</c> = 2, etc.; <c>Title</c> = 1, <c>Subtitle</c> = 2.
    /// Tracked-change mode inherits <see cref="DeleteRange"/>'s behavior via the shared
    /// <c>DeleteSiblingRangeCore</c> helper, including native <c>w:sdt</c>/<c>w:customXml</c>
    /// envelope deletion, anchor accounting, and the pre-mutation refusal of run-level
    /// <c>w:customXml</c>.
    /// </remarks>
    public EditResult DeleteSection(string headingAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var headingTarget = FindAnchor(headingAnchorId);
        if (headingTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"heading anchor not found: {headingAnchorId}", headingAnchorId);
        if (headingTarget.Anchor.Kind != "h")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteSection requires a heading anchor (kind=h); got kind={headingTarget.Anchor.Kind}",
                headingAnchorId);

        var headingElement = headingTarget.Resolve(_doc!);
        if (headingElement is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "heading element resolved null", headingAnchorId);

        int level = WmlToMarkdownConverter.HeadingLevel(headingElement);

        // Scan forward siblings for the next heading at level <= ours. If none, toElement
        // stays null and DeleteSiblingRangeCore will delete to the end of the parent.
        XElement? toElement = null;
        foreach (var sibling in headingElement.ElementsAfterSelf())
        {
            if (sibling.Name == W.p && WmlToMarkdownConverter.IsHeading(sibling)
                && WmlToMarkdownConverter.HeadingLevel(sibling) <= level)
            {
                toElement = sibling;
                break;
            }
        }

        return DeleteSiblingRangeCore(headingTarget, headingElement, toElement);
    }

    /// <summary>
    /// Shared core for <see cref="DeleteRange"/>, <see cref="DeleteSection"/>, and
    /// tracked <see cref="DeleteBlock"/>.
    /// Takes resolved XElement endpoints — <paramref name="toElementExclusive"/> may be
    /// <c>null</c> to mean "delete to the end of the parent". Records one snapshot and
    /// returns a single <see cref="EditResult"/> aggregating every removed anchor.
    /// </summary>
    private EditResult DeleteSiblingRangeCore(
        AnchorTarget anchorForPatchScope,
        XElement fromElement,
        XElement? toElementExclusive)
    {
        // Walk siblings from `fromElement` forward, accumulating elements to remove.
        var toRemove = new List<XElement>();
        var current = (XElement?)fromElement;
        while (current is not null && current != toElementExclusive)
        {
            toRemove.Add(current);
            current = current.ElementsAfterSelf().FirstOrDefault();
        }
        if (toElementExclusive is not null && current != toElementExclusive)
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                "'to' anchor does not follow 'from' in document order",
                anchorForPatchScope.Anchor.Id);

        bool trackedChanges = _trackedChanges == TrackedChangeMode.RenderInline;
        if (trackedChanges && FirstUnrecordableInlineContent(toRemove) is { } unrecordable)
        {
            return EditResult.Fail(
                EditErrorCode.IncompatibleElementType,
                $"Tracked block deletion does not support {unrecordable} inside a paragraph; no changes were made.",
                anchorForPatchScope.Anchor.Id);
        }

        // A live non-body container must end in a paragraph (the rule InsertTable keeps). When
        // the deletion would leave it empty or ending in a table, record the content deletion
        // without deleting the final pilcrow, so every review engine keeps the same paragraph
        // and formatting rather than synthesizing one or emptying the story entirely.
        XElement? retainedParagraph = null;
        if (trackedChanges && fromElement.Parent is { } parent && parent.Name != W.body
            && toRemove.LastOrDefault(el => el.Name == W.p) is { } lastParagraph)
        {
            var lastSurvivor = parent.Elements().Except(toRemove)
                .LastOrDefault(IsSurvivingParagraphContainerBlock);
            if (lastSurvivor is null || (lastSurvivor.Name == W.tbl && lastSurvivor.IsBefore(lastParagraph)))
                retainedParagraph = lastParagraph;
        }

        // Paragraph markers normally migrate to a surviving following paragraph. Tables,
        // inline controls, and terminal paragraphs instead lose their bookmark endpoints.
        // References inside any selected block are being deleted too, even when its
        // paragraph shell survives, and must not cause a false BookmarkInUse refusal.
        var removedParagraphs = trackedChanges
            ? ParagraphsRemovedOnAcceptance(toRemove, retainedParagraph)
            : new HashSet<XElement>();
        var structuralRoots = trackedChanges
            ? toRemove.Where(el => el.Name != W.p || removedParagraphs.Contains(el))
                .Concat(removedParagraphs.Except(toRemove))
                .Concat(toRemove.Where(el => el.Name == W.p).SelectMany(el =>
                    el.Descendants(W.sdt).Where(IsInlineControl))).ToList()
            : toRemove;
        if (ValidateBookmarkRemoval(structuralRoots, anchorForPatchScope.Anchor.Id, toRemove) is { } bookmarkError)
            return bookmarkError;

        // Content the session author inserted is un-inserted outright rather than marked (see
        // DeleteInlineElementInPlace), which can orphan the notes, hyperlinks and images it
        // carried, so a deletion that removes own content sweeps like a structural removal.
        // A mere own paragraph mark or row mark removes nothing and must not trigger the
        // whole-part sweep that would take unrelated pre-existing orphan relationships with it.
        var author = _revisionAuthor ?? "docxodus";
        bool structurallyRemoves = !trackedChanges || toRemove.Any(el =>
            (el.Name != W.p && el.Name != W.tbl && el.Name != W.sdt && el.Name != W.customXml)
            || (el.Name == W.p && el != retainedParagraph && IsWhollyOwnInsertion(el, author))
            || HasOwnInsertedContent(el, author));
        var hyperlinkOwner = structurallyRemoves
            ? Internal.OwnedPartRelationships.FindOwner(_doc!, fromElement) : null;
        var referencedNotesBefore = structurallyRemoves ? ReferencedNoteIds()
            : ((HashSet<int> Footnotes, HashSet<int> Endnotes)?)null;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var index = AnchorIndex();
            if (trackedChanges)
            {
                // Tracked-change path: mark each block with w:del markup rather than
                // removing it. Anchors stay live in the document tree so callers can
                // re-address the same blocks before changes are accepted. Ordinary
                // paragraphs/tables retain DeleteBlock's single top-level-anchor
                // contract. Structured wrappers and every descendant anchor they keep
                // live are reported as Modified. A remaining
                // structural fall-through is a real removal and is reported as such.
                var modified = new List<Anchor>();
                var trackedRemoved = new List<Anchor>();
                var modifiedIds = new HashSet<string>(StringComparer.Ordinal);
                var trackedRemovedIds = new HashSet<string>(StringComparer.Ordinal);
                var stamp = NewRevisionStamp();
                foreach (var el in toRemove)
                {
                    if (el.Name == W.p)
                    {
                        if (el != retainedParagraph && IsWhollyOwnInsertion(el, stamp.Author))
                        {
                            // The author's own pending paragraph, text and pilcrow alike: Word
                            // removes it outright rather than recording a deletion of an
                            // insertion that nobody could reject into anything.
                            CollectAnchors(el, includeDescendants: true, index, anchorForPatchScope, trackedRemoved, trackedRemovedIds);
                            el.Remove();
                            continue;
                        }
                        CollectAnchors(el, includeDescendants: false, index, anchorForPatchScope, modified, modifiedIds);
                        MarkParagraphAsTrackedDeleted(el, stamp, preserveParagraphMark: el == retainedParagraph);
                    }
                    else if (el.Name == W.tbl)
                    {
                        CollectAnchors(el, includeDescendants: false, index, anchorForPatchScope, modified, modifiedIds);
                        MarkTableAsTrackedDeleted(el, stamp);
                    }
                    else if (el.Name == W.sdt || el.Name == W.customXml)
                    {
                        CollectAnchors(el, includeDescendants: true, index, anchorForPatchScope, modified, modifiedIds);
                        MarkStructuredBlockAsTrackedDeleted(el, stamp);
                    }
                    else
                    {
                        CollectAnchors(
                            el,
                            includeDescendants: true,
                            index,
                            anchorForPatchScope,
                            trackedRemoved,
                            trackedRemovedIds);
                        el.Remove();
                    }
                }
                if (referencedNotesBefore is { } before)
                    AppendPrunedNoteAnchors(PruneOrphanedNotes(before), trackedRemoved, trackedRemovedIds);
                if (hyperlinkOwner is { } trackedOwner)
                    SweepOrphanedStoryRelationships(trackedOwner.Part);
                InvalidateProjectionCache(sweepOrphanedImages: structurallyRemoves);
                return new EditResult
                {
                    Success = true,
                    Modified = modified,
                    Removed = trackedRemoved,
                    Patch = PatchFor(anchorForPatchScope),
                };
            }

            var removed = new List<Anchor>();
            var removedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var el in toRemove)
            {
                // Collect this element's anchor plus every descendant anchor.
                CollectAnchors(el, includeDescendants: true, index, anchorForPatchScope, removed, removedIds);
                el.Remove();
            }
            AppendPrunedNoteAnchors(PruneOrphanedNotes(referencedNotesBefore!.Value), removed, removedIds);
            if (hyperlinkOwner is { } owner)
                SweepOrphanedStoryRelationships(owner.Part);
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Removed = removed,
                Patch = PatchFor(anchorForPatchScope),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorForPatchScope.Anchor.Id);
        }
    }

    private static void CollectAnchors(
        XElement el,
        bool includeDescendants,
        IReadOnlyDictionary<string, AnchorTarget> index,
        AnchorTarget scope,
        List<Anchor> destination,
        HashSet<string> seenIds)
    {
        var candidates = includeDescendants ? el.DescendantsAndSelf() : new[] { el };
        foreach (var candidate in candidates)
        {
            var unid = (string?)candidate.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            // The known target needs no reverse scan. Every other lookup is part-scoped:
            // identical header/footer content can legitimately share the same Unid.
            var anchor = unid == scope.Unid ? scope.Anchor
                : index.Values.FirstOrDefault(target => target.PartUri == scope.PartUri && target.Unid == unid)?.Anchor;
            if (anchor is { } found && seenIds.Add(found.Id))
                destination.Add(found);
        }
    }

    /// <summary>A paragraph whose mark an earlier revision already deletes or moves away.</summary>
    private static bool IsTrackedDeletedParagraph(XElement paragraph) =>
        paragraph.Element(W.pPr)?.Element(W.rPr) is { } mark
        && (mark.Element(W.del) is not null || mark.Element(W.moveFrom) is not null);

    /// <summary>An inline content control: one inside a paragraph's own run content, as opposed
    /// to a block control reached through a text box anchored in that paragraph.</summary>
    private static bool IsInlineControl(XElement sdt) =>
        sdt.Ancestors().TakeWhile(a => a.Name != W.txbxContent).Any(a => a.Name == W.p);

    private static bool IsBlockSibling(XElement element) =>
        element.Name == W.p || element.Name == W.tbl || element.Name == W.sdt
        || element.Name == W.customXml || element.Name == W.altChunk;

    /// <summary>Whether a container child still holds a block once every pending deletion is
    /// accepted: a paragraph whose mark survives, a table with a surviving row, or a wrapper
    /// holding such a block — the removal both review engines perform.</summary>
    private static bool IsSurvivingParagraphContainerBlock(XElement element) =>
        element.Name == W.p ? !IsTrackedDeletedParagraph(element)
        : element.Name == W.tbl ? WordprocessingMLUtil.TableRows(element)
            .Any(row => row.Element(W.trPr)?.Element(W.del) is null)
        : element.Name == W.sdt ? element.Element(W.sdtContent)?.Elements()
            .Any(IsSurvivingParagraphContainerBlock) == true
        : element.Name == W.customXml ? element.Elements().Any(IsSurvivingParagraphContainerBlock)
        : element.Name == W.altChunk;

    /// <summary>Paragraph content the tracked deleter cannot mark: a run-level custom-XML
    /// wrapper (no reversible envelope), a simple field with no result run (nothing for a
    /// <c>w:del</c> to hold), or a subdocument reference. Text-box content rides inside its
    /// anchoring run and is exempt. Returns what to name in the refusal, or null.</summary>
    private static string? UnrecordableInlineContentKind(XElement element)
    {
        var kind = element.Name == W.customXml ? "run-level w:customXml"
            : element.Name == W.fldSimple && !element.Descendants(W.r).Any() ? "an empty w:fldSimple"
            : element.Name == W.subDoc ? "w:subDoc"
            : null;
        return kind is not null
            && element.Ancestors().Any(a => a.Name == W.p)
            && !element.Ancestors().Any(a => a.Name == W.txbxContent)
            ? kind : null;
    }

    /// <summary>The first content under <paramref name="roots"/> a tracked deletion cannot mark,
    /// named for the refusal, or null. Shared by block deletion and whole-paragraph replacement
    /// so both refuse the same shapes before mutating anything.</summary>
    private static string? FirstUnrecordableInlineContent(IEnumerable<XElement> roots) =>
        roots.SelectMany(root => root.DescendantsAndSelf())
            .Select(UnrecordableInlineContentKind)
            .FirstOrDefault(kind => kind is not null);

    /// <summary>
    /// The paragraphs a tracked deletion of <paramref name="range"/> removes outright on
    /// acceptance, whose bookmark endpoints therefore cannot migrate: every paragraph in the
    /// range other than <paramref name="retained"/> whose following block is not a paragraph
    /// that keeps its markers (live, or removed with markers migrating further — through
    /// paragraphs an earlier revision already deleted as well), plus the already-deleted
    /// paragraphs immediately before the range that were migrating into its first paragraph.
    /// One backward pass over the contiguous sibling range.
    /// </summary>
    private static HashSet<XElement> ParagraphsRemovedOnAcceptance(
        IReadOnlyList<XElement> range, XElement? retained)
    {
        var removed = new HashSet<XElement>();
        bool markersSurvive = false;
        for (var next = range[^1].ElementsAfterSelf().FirstOrDefault(IsBlockSibling);
             next is not null && next.Name == W.p;
             next = next.ElementsAfterSelf().FirstOrDefault(IsBlockSibling))
        {
            if (IsTrackedDeletedParagraph(next)) continue;
            markersSurvive = true;
            break;
        }

        var precedingDeleted = range[0].Name == W.p
            ? range[0].ElementsBeforeSelf().Reverse().Where(IsBlockSibling)
                .TakeWhile(sibling => sibling.Name == W.p && IsTrackedDeletedParagraph(sibling))
            : Enumerable.Empty<XElement>();
        foreach (var element in range.Reverse().Concat(precedingDeleted))
        {
            if (element.Name == W.p)
            {
                bool isRemoved = element != retained && !markersSurvive;
                if (isRemoved) removed.Add(element);
                markersSurvive = !isRemoved;
            }
            else if (IsBlockSibling(element)) markersSurvive = false;
        }
        return removed;
    }

    /// <summary>
    /// Strips every cross-reference pointing at the named footnote/endnote/comment id
    /// from every part of the package that can hold one. For footnotes/endnotes that's
    /// just <c>w:footnoteReference</c>/<c>w:endnoteReference</c>; for comments it's the
    /// triple <c>w:commentReference</c> + <c>w:commentRangeStart</c> + <c>w:commentRangeEnd</c>
    /// — leaving any of the three behind makes Word render a broken comment marker.
    /// </summary>
    private void RemoveCrossReferences(string kind, string elementId)
    {
        XName referenceName = kind switch
        {
            "fn" => W.footnoteReference,
            "en" => W.endnoteReference,
            "cmt" => W.commentReference,
            _ => null!,
        };
        if (referenceName is null) return;

        foreach (var part in EnumerateProjectedParts())
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            bool any = false;
            foreach (var refEl in root.Descendants(referenceName)
                .Where(r => (string?)r.Attribute(W.id) == elementId).ToList())
            {
                var parentRun = refEl.Parent;
                refEl.Remove();
                any = true;
                // The reference was the only meaningful child of its <w:r> wrapper:
                // strip the run too so we don't leave behind an empty <w:r> with a
                // FootnoteReference run style (which Word renders as an empty styled
                // span — invisible but untidy and confusing to downstream tooling).
                RemoveEmptyRunIfNeeded(parentRun);
            }
            if (kind == "cmt")
            {
                foreach (var rangeEl in root.Descendants(W.commentRangeStart)
                    .Concat(root.Descendants(W.commentRangeEnd))
                    .Where(r => (string?)r.Attribute(W.id) == elementId).ToList())
                {
                    rangeEl.Remove();
                    any = true;
                }
            }
            if (any) part.PutXDocument();
        }
    }

    /// <summary>
    /// If <paramref name="run"/> is a <c>&lt;w:r&gt;</c> whose only remaining children
    /// are properties (<c>w:rPr</c>) — no text, no breaks, no fields, no other content —
    /// remove the run. Avoids leaving orphaned styled-empty spans after the meaningful
    /// child (a footnote/endnote reference) was stripped.
    /// </summary>
    private static void RemoveEmptyRunIfNeeded(XElement? run)
    {
        if (run is null || run.Name != W.r) return;
        foreach (var child in run.Elements())
        {
            if (child.Name == W.rPr) continue;
            return; // has meaningful content — keep the run
        }
        run.Remove();
    }
}
