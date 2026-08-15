// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

public enum ContentControlType
{
    PlainText,
    RichText,
    Checkbox,
    Date,
    DropDownList,
    ComboBox,
    Picture,
    RepeatingSection,
    RepeatingSectionItem,
    Unsupported,
}

public enum ContentControlPlacement { Inline, Block, Row, Cell, Unknown }

public enum ContentControlBindingPolicy
{
    /// <summary>Never alter a binding. Bound controls fail closed.</summary>
    Preserve = 0,

    /// <summary>Remove only the selected control's own native data-binding element
    /// (w:dataBinding or w15:dataBinding) before filling it. A binding on any ancestor
    /// still fails closed.</summary>
    DetachTarget = 1,
}

public sealed record ContentControlFillOptions
{
    public ContentControlBindingPolicy BindingPolicy { get; init; } = ContentControlBindingPolicy.Preserve;
}

public sealed record ContentControlBindingInfo(
    string? StoreItemId, string? XPath, string? PrefixMappings);

/// <summary>A native Word structured-document tag in outer-before-inner story order.</summary>
public sealed record ContentControlInfo
{
    required public string AnchorId { get; init; }
    required public ContentControlType Type { get; init; }
    required public ContentControlPlacement Placement { get; init; }
    public string? NativeId { get; init; }
    public string? Tag { get; init; }
    public string? Alias { get; init; }
    public string? Lock { get; init; }
    public bool IsShowingPlaceholder { get; init; }
    public ContentControlBindingInfo? Binding { get; init; }
    public bool IsBound => Binding is not null;
    required public string OwningPartUri { get; init; }
    required public string Scope { get; init; }
    public string? ParentAnchorId { get; init; }
    public int Depth { get; init; }
    public bool HasValidNativeId { get; init; }
    public bool HasDuplicateNativeId { get; init; }
    public bool CanMutate { get; init; }
    public bool CanDetachTargetBinding { get; init; }
    public string? UnsupportedReason { get; init; }
    public string Text { get; init; } = string.Empty;
    public IReadOnlyList<string> ItemValues { get; init; } = Array.Empty<string>();
}

public sealed partial class DocxSession
{
    private static readonly XNamespace ContentControlW =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace ContentControlW14 =
        "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace ContentControlW15 =
        "http://schemas.microsoft.com/office/word/2012/wordml";

    private static readonly IReadOnlyDictionary<XName, ContentControlType>
        ContentControlFamilies = new Dictionary<XName, ContentControlType>
        {
            [ContentControlW14 + "checkbox"] = ContentControlType.Checkbox,
            [ContentControlW15 + "repeatingSection"] = ContentControlType.RepeatingSection,
            [ContentControlW15 + "repeatingSectionItem"] = ContentControlType.RepeatingSectionItem,
            [W.picture] = ContentControlType.Picture,
            [W.date] = ContentControlType.Date,
            [W.dropDownList] = ContentControlType.DropDownList,
            [W.comboBox] = ContentControlType.ComboBox,
            [W.text] = ContentControlType.PlainText,
            [ContentControlW + "richText"] = ContentControlType.RichText,
        };

    private static readonly HashSet<XName> ContentControlMetadata = new()
    {
        W.id, W.tag, W.alias, W.dataBinding, ContentControlW15 + "dataBinding",
        W.showingPlcHdr, ContentControlW + "lock", ContentControlW + "placeholder",
        ContentControlW + "temporary", ContentControlW15 + "appearance",
        ContentControlW15 + "color", W.rPr,
    };

    private sealed record ContentControlCandidate(
        OwnedPartRelationships.Owner Owner,
        XElement Element,
        ContentControlIdentity.Entry Identity,
        ContentControlInfo Info,
        string? MalformedReason,
        string? MalformedAncestorReason);

    private sealed record PictureContentControlTarget(
        ImageCandidate? Image,
        EditErrorCode? ErrorCode,
        string? Diagnostic);

    public IReadOnlyList<ContentControlInfo> ListContentControls(
        ProjectionScopes scopes = ProjectionScopes.All)
    {
        ThrowIfDisposed();
        return BuildContentControlRegistry(scopes).Select(candidate => candidate.Info).ToList();
    }

    public ContentControlInfo? GetContentControl(string anchorId)
    {
        ThrowIfDisposed();
        return BuildContentControlRegistry(ProjectionScopes.All)
            .FirstOrDefault(candidate => string.Equals(candidate.Info.AnchorId, anchorId,
                StringComparison.Ordinal))?.Info;
    }

    public EditResult FillContentControlText(string anchorId, string text,
        ContentControlFillOptions? options = null) =>
        FillTextualContentControl(anchorId, text, rich: false, options);

    public EditResult FillContentControlRichText(string anchorId, string markdown,
        ContentControlFillOptions? options = null) =>
        FillTextualContentControl(anchorId, markdown, rich: true, options);

    public EditResult SetContentControlChecked(string anchorId, bool isChecked,
        ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(anchorId, ContentControlType.Checkbox, options,
            out var candidate, out var error) is false) return error!;
        if (ContainsNestedContentControl(candidate!.Element))
            return NestedFillError(anchorId);

        var checkbox = candidate.Element.Element(W.sdtPr)?.Element(ContentControlW14 + "checkbox");
        if (checkbox is null)
            return EditResult.Fail(EditErrorCode.ContentControlMalformed,
                "checkbox content control has no w14:checkbox properties", anchorId);
        var checkedElement = checkbox.Element(ContentControlW14 + "checked");

        var stateElement = checkbox.Element(isChecked
            ? ContentControlW14 + "checkedState"
            : ContentControlW14 + "uncheckedState");
        var fallback = isChecked ? 0x2612 : 0x2610;
        var glyph = TryParseHexScalar((string?)stateElement?.Attribute(ContentControlW14 + "val"),
            out var scalar) ? char.ConvertFromUtf32(scalar) : char.ConvertFromUtf32(fallback);
        var stateFont = (string?)stateElement?.Attribute(ContentControlW14 + "font");
        if (ValidateWholeContentReplacement(candidate, replacement: null, anchorId) is { } replacementError)
            return replacementError;

        return MutateContentControl(candidate, options, () =>
        {
            if (checkedElement is null)
            {
                checkedElement = new XElement(ContentControlW14 + "checked");
                checkbox.AddFirst(checkedElement);
            }
            checkedElement.SetAttributeValue(ContentControlW14 + "val", isChecked ? "1" : "0");
            ReplaceControlWithPlainText(candidate.Element, glyph, stateFont);
        });
    }

    public EditResult SetContentControlDate(string anchorId, DateTimeOffset value,
        string? displayText = null, ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(anchorId, ContentControlType.Date, options,
            out var candidate, out var error) is false) return error!;
        if (ContainsNestedContentControl(candidate!.Element))
            return NestedFillError(anchorId);
        var date = candidate.Element.Element(W.sdtPr)?.Element(W.date);
        if (date is null)
            return EditResult.Fail(EditErrorCode.ContentControlMalformed,
                "date content control has no w:date properties", anchorId);
        var shown = displayText ?? value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        if (ValidateWholeContentReplacement(candidate, replacement: null, anchorId) is { } replacementError)
            return replacementError;
        return MutateContentControl(candidate, options, () =>
        {
            date.SetAttributeValue(W.fullDate, value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'",
                CultureInfo.InvariantCulture));
            ReplaceControlWithPlainText(candidate.Element, shown);
        });
    }

    public EditResult SelectContentControlItem(string anchorId, string value,
        ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(anchorId,
            new[] { ContentControlType.DropDownList, ContentControlType.ComboBox }, options,
            out var candidate, out var error) is false) return error!;
        if (ContainsNestedContentControl(candidate!.Element))
            return NestedFillError(anchorId);
        var props = candidate.Element.Element(W.sdtPr)!;
        var list = props.Element(W.dropDownList) ?? props.Element(W.comboBox)!;
        var isComboBox = list.Name == W.comboBox;
        var matches = list.Elements(W.listItem).Where(item =>
            string.Equals((string?)item.Attribute(ContentControlW + "value"), value, StringComparison.Ordinal)
            || string.Equals((string?)item.Attribute(W.displayText), value, StringComparison.Ordinal)).ToList();
        if (matches.Count > 1 || matches.Count == 0 && !isComboBox)
            return EditResult.Fail(EditErrorCode.InvalidContentControlValue,
                matches.Count == 0
                    ? $"content control has no list item matching '{value}'"
                    : $"content control has multiple list items matching '{value}'", anchorId);
        var selectedValue = matches.Count == 1
            ? (string?)matches[0].Attribute(ContentControlW + "value")
                ?? (string?)matches[0].Attribute(W.displayText) ?? string.Empty
            : value;
        var display = matches.Count == 1
            ? (string?)matches[0].Attribute(W.displayText) ?? selectedValue
            : value;
        if (ValidateWholeContentReplacement(candidate, replacement: null, anchorId) is { } replacementError)
            return replacementError;
        return MutateContentControl(candidate, options, () =>
        {
            list.SetAttributeValue(W.lastValue, selectedValue);
            ReplaceControlWithPlainText(candidate.Element, display);
        });
    }

    public EditResult FillContentControlPicture(string anchorId, byte[] imageBytes,
        ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(anchorId, ContentControlType.Picture, options,
            out var candidate, out var error) is false) return error!;
        if (ContainsNestedContentControl(candidate!.Element))
            return NestedFillError(anchorId);
        var binary = ValidateImageBytes(imageBytes, anchorId);
        if (binary.Error is not null) return binary.Error;
        var target = ResolvePictureContentControlTarget(candidate!.Element,
            EnumerateImageCandidates(ProjectionScopes.All));
        if (target.ErrorCode is { } errorCode)
            return EditResult.Fail(errorCode, target.Diagnostic!, anchorId);
        var image = target.Image!;
        var blip = image.Blip!;

        return MutateContentControl(candidate!, options, () =>
        {
            var relationship = OwnedPartRelationships.FindOrAddImagePart(_doc!, candidate!.Owner.Part,
                imageBytes, binary.ContentType!, binary.Format);
            blip.SetAttributeValue(ImageR + "embed", relationship.RelationshipId);
            candidate.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
            OwnedPartRelationships.SweepOrphanedImages(candidate.Owner.Part);
        });
    }

    /// <summary>Clone one direct repeating-section item. The new item is inserted after
    /// <paramref name="afterItemAnchorId"/>, or after the final item when omitted.</summary>
    public EditResult AddRepeatingSectionItem(string sectionAnchorId,
        string? afterItemAnchorId = null, ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(sectionAnchorId, ContentControlType.RepeatingSection,
            options, out var section, out var error) is false) return error!;
        var content = section!.Element.Element(W.sdtContent);
        if (content is null)
            return EditResult.Fail(EditErrorCode.ContentControlMalformed,
                "repeating section has no w:sdtContent", sectionAnchorId);
        var items = content.Elements(W.sdt).Where(IsRepeatingSectionItem).ToList();
        if (items.Count == 0 || content.Elements().Any(element => element.Name != W.sdt
                || !IsRepeatingSectionItem(element)))
            return EditResult.Fail(EditErrorCode.RepeatingSectionConstraint,
                "repeating section must contain only one or more direct repeating-section-item controls",
                sectionAnchorId);

        XElement template;
        if (afterItemAnchorId is null) template = items[^1];
        else
        {
            var after = BuildContentControlRegistry(ProjectionScopes.All).FirstOrDefault(value =>
                string.Equals(value.Info.AnchorId, afterItemAnchorId, StringComparison.Ordinal));
            if (after is null || !items.Any(item => ReferenceEquals(item, after.Element)))
                return EditResult.Fail(EditErrorCode.RepeatingSectionConstraint,
                    "afterItemAnchorId is not a direct item of the selected repeating section",
                    afterItemAnchorId);
            template = after.Element;
        }
        if (FindUnsafeRepeatingCloneCarrier(template) is { } unsafeCarrier)
            return EditResult.Fail(EditErrorCode.RepeatingSectionConstraint,
                $"repeating item contains clone-sensitive markup ({unsafeCarrier})",
                sectionAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            DetachTargetBindingIfRequested(section.Element, options);
            var clone = new XElement(template);
            foreach (var element in clone.DescendantsAndSelf())
                element.Attribute(PtOpenXml.Unid)?.Remove();
            AssignFreshContentControlIds(clone);
            UnidHelper.AssignToSelfAndDescendants(clone);
            AssignFreshDocumentPropertyIds(clone);
            AssignFreshParagraphIds(clone);
            template.AddAfterSelf(clone);
            ContentControlIdentity.AssignStableUnids(section.Owner.Part.GetXDocument().Root!);
            InvalidateProjectionCache();
            var createdUnid = (string)clone.Attribute(PtOpenXml.Unid)!;
            var created = new Anchor($"sdt:{section.Owner.Scope}:{createdUnid}", "sdt",
                section.Owner.Scope, createdUnid);
            return new EditResult { Success = true, Created = new[] { created },
                Modified = new[] { AnchorFromCandidate(section) } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, sectionAnchorId);
        }
    }

    public EditResult RemoveRepeatingSectionItem(string itemAnchorId)
    {
        if (ResolveContentControlForMutation(itemAnchorId, ContentControlType.RepeatingSectionItem,
            options: null, out var item, out var error, removingWrapper: true) is false) return error!;
        var outer = item!.Element.Parent?.Parent;
        if (outer is null || outer.Name != W.sdt || !IsRepeatingSection(outer)
            || item.Element.Parent?.Name != W.sdtContent)
            return EditResult.Fail(EditErrorCode.RepeatingSectionConstraint,
                "repeating-section item is not a direct child of a repeating section", itemAnchorId);
        var siblings = item.Element.Parent.Elements(W.sdt).Where(IsRepeatingSectionItem).ToList();
        if (siblings.Count <= 1)
            return EditResult.Fail(EditErrorCode.RepeatingSectionConstraint,
                "a repeating section must retain at least one item", itemAnchorId);
        var parentCandidate = BuildContentControlRegistry(ProjectionScopes.All).First(value =>
            ReferenceEquals(value.Element, outer));
        if (ValidateEffectiveLocks(parentCandidate, removingWrapper: false) is { } parentLock)
            return parentLock;
        if (ValidateBindingPolicy(parentCandidate, options: null) is { } bindingError)
            return bindingError;
        if (ValidateBookmarkRemoval(new[] { item.Element }, itemAnchorId) is { } bookmarkError)
            return bookmarkError;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var removed = AnchorFromCandidate(item);
            item.Element.Remove();
            SweepOrphanedStoryRelationships(item.Owner.Part);
            InvalidateProjectionCache();
            return new EditResult { Success = true, Removed = new[] { removed },
                Modified = new[] { AnchorFromCandidate(parentCandidate) } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, itemAnchorId);
        }
    }

    private EditResult FillTextualContentControl(string anchorId, string payload, bool rich,
        ContentControlFillOptions? options)
    {
        var expected = rich ? ContentControlType.RichText : ContentControlType.PlainText;
        if (ResolveContentControlForMutation(anchorId, expected, options,
            out var candidate, out var error) is false) return error!;
        if (ContainsNestedContentControl(candidate!.Element)) return NestedFillError(anchorId);

        if (!rich)
        {
            if (ValidateWholeContentReplacement(candidate, replacement: null, anchorId) is { } replacementError)
                return replacementError;
            return MutateContentControl(candidate, options,
                () => ReplaceControlWithPlainText(candidate.Element, payload));
        }

        var parsed = MarkdownPayloadParser.Parse(payload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
        if (parsed.Blocks.Count == 0)
            parsed = ParseResult.Ok(new[]
            {
                new ParsedBlock(ParserBlockKind.Paragraph, 0,
                    new[] { new XElement(W.r) }),
            });
        if (candidate.Info.Placement == ContentControlPlacement.Inline && parsed.Blocks.Count != 1)
            return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                "an inline rich-text control accepts exactly one markdown block", anchorId);
        if (candidate.Info.Placement is not (ContentControlPlacement.Inline or ContentControlPlacement.Block))
            return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                "rich-text fill supports only inline and block content controls", anchorId);
        var replacement = parsed.Blocks.SelectMany(block => block.RunElements).ToList();
        if (ValidateWholeContentReplacement(candidate, replacement, anchorId) is { } richReplacementError)
            return richReplacementError;

        return MutateContentControl(candidate, options, () =>
        {
            var content = candidate.Element.Element(W.sdtContent)!;
            if (candidate.Info.Placement == ContentControlPlacement.Inline)
            {
                var block = parsed.Blocks[0];
                content.ReplaceNodes(block.RunElements.Select(element => new XElement(element)));
            }
            else
            {
                var blocks = parsed.Blocks.Select(BuildParagraphFromParsedBlock).ToList();
                content.ReplaceNodes(blocks);
            }
            candidate.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
        });
    }

    private bool ResolveContentControlForMutation(string anchorId, ContentControlType expected,
        ContentControlFillOptions? options, out ContentControlCandidate? candidate,
        out EditResult? error, bool removingWrapper = false) =>
        ResolveContentControlForMutation(anchorId, new[] { expected }, options,
            out candidate, out error, removingWrapper);

    private bool ResolveContentControlForMutation(string anchorId,
        IReadOnlyCollection<ContentControlType> expected, ContentControlFillOptions? options,
        out ContentControlCandidate? candidate, out EditResult? error,
        bool removingWrapper = false)
    {
        candidate = null;
        error = null;
        if (_disposed)
        {
            error = EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
            return false;
        }
        if (TrackedContentControlBlocker() is { } trackedReason)
        {
            error = EditResult.Fail(EditErrorCode.TrackedOperationUnsupported, trackedReason, anchorId);
            return false;
        }
        candidate = BuildContentControlRegistry(ProjectionScopes.All).FirstOrDefault(value =>
            string.Equals(value.Info.AnchorId, anchorId, StringComparison.Ordinal));
        if (candidate is null)
        {
            error = EditResult.Fail(EditErrorCode.ContentControlNotFound,
                $"content control not found: {anchorId}", anchorId);
            return false;
        }
        if (candidate.MalformedReason is not null || candidate.MalformedAncestorReason is not null
            || !candidate.Identity.HasMutableIdentity)
        {
            error = EditResult.Fail(EditErrorCode.ContentControlMalformed,
                candidate.MalformedReason ?? candidate.MalformedAncestorReason
                    ?? candidate.Info.UnsupportedReason
                    ?? "content control has no unique valid native w:id", anchorId);
            return false;
        }
        if (candidate.Info.Type == ContentControlType.Unsupported)
        {
            error = EditResult.Fail(EditErrorCode.ContentControlUnsupported,
                candidate.Info.UnsupportedReason ?? "unsupported content-control family", anchorId);
            return false;
        }
        if (!expected.Contains(candidate.Info.Type))
        {
            error = EditResult.Fail(EditErrorCode.ContentControlWrongType,
                $"operation requires {string.Join(" or ", expected)} but target is {candidate.Info.Type}", anchorId);
            return false;
        }
        if (candidate.Info.Placement == ContentControlPlacement.Unknown)
        {
            error = EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                candidate.Info.UnsupportedReason ?? "unsupported content-control placement", anchorId);
            return false;
        }
        if (!IsMutationPlacementSupported(candidate.Info.Type, candidate.Info.Placement))
        {
            error = EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                candidate.Info.UnsupportedReason
                    ?? $"{candidate.Info.Type} mutation supports only inline and block content controls",
                anchorId);
            return false;
        }
        if (ValidateEffectiveLocks(candidate, removingWrapper) is { } lockError)
        {
            error = lockError;
            return false;
        }
        if (ValidateBindingPolicy(candidate, options) is { } bindingError)
        {
            error = bindingError;
            return false;
        }
        return true;
    }

    /// <summary>
    /// The session-mode gate every content-control mutation shares: a whole-control fill and a
    /// repeating-item add/remove both rewrite a payload wholesale, which has no faithful tracked
    /// representation. Discovery and mutation both read it, so the registry cannot advertise
    /// <c>canMutate</c> for an operation that is guaranteed to be refused.
    /// </summary>
    private string? TrackedContentControlBlocker() =>
        _trackedChanges == TrackedChangeMode.RenderInline
            ? "content-control mutations cannot be represented faithfully as tracked revisions; use surgical text operations inside the control or switch modes"
            : null;

    /// <summary>
    /// The bookmark consequences of an operation that discards the target's complete payload —
    /// the same gate <see cref="ValidateWholeContentReplacement"/> applies to a whole-control
    /// fill and <see cref="RemoveRepeatingSectionItem"/> applies to the item it removes.
    /// Evaluated once for discovery so a control whose fill is certain to fail is not reported
    /// mutable. Picture fill is excluded because it rewrites only the blip relationship and
    /// therefore takes no bookmark gate at mutation time.
    /// </summary>
    private string? WholeContentBookmarkBlocker(XElement element, ContentControlType type)
    {
        var removalRoot = type switch
        {
            ContentControlType.RepeatingSectionItem => element,
            _ when IsWholeContentReplacementType(type) => element.Element(W.sdtContent),
            _ => null,
        };
        return removalRoot is null
            ? null
            : ValidateBookmarkRemoval(new[] { removalRoot }, string.Empty)?.Error?.Message;
    }

    private EditResult? ValidateEffectiveLocks(ContentControlCandidate candidate, bool removingWrapper)
    {
        foreach (var control in candidate.Element.AncestorsAndSelf(W.sdt))
        {
            var token = (string?)control.Element(W.sdtPr)?.Element(ContentControlW + "lock")?.Attribute(W.val);
            if (token is "contentLocked" or "sdtContentLocked")
                return EditResult.Fail(EditErrorCode.ContentControlLocked,
                    "target content is locked by this control or an ancestor", candidate.Info.AnchorId);
            if (removingWrapper && ReferenceEquals(control, candidate.Element)
                && token is "sdtLocked" or "sdtContentLocked")
                return EditResult.Fail(EditErrorCode.ContentControlLocked,
                    "target content-control wrapper is locked", candidate.Info.AnchorId);
        }
        return null;
    }

    private EditResult? ValidateBindingPolicy(ContentControlCandidate candidate,
        ContentControlFillOptions? options)
    {
        var boundControls = candidate.Element.AncestorsAndSelf(W.sdt).Where(control =>
            FindDataBinding(control.Element(W.sdtPr)) is not null).ToList();
        if (boundControls.Count == 0) return null;
        var targetBound = boundControls.Any(control => ReferenceEquals(control, candidate.Element));
        var hasBoundAncestor = boundControls.Any(control => !ReferenceEquals(control, candidate.Element));
        if (hasBoundAncestor)
            return EditResult.Fail(EditErrorCode.ContentControlBound,
                "target is inside a data-bound ancestor; only the selected target's own binding may be detached",
                candidate.Info.AnchorId);
        if (!targetBound) return null;
        if (options?.BindingPolicy == ContentControlBindingPolicy.DetachTarget) return null;
        return EditResult.Fail(EditErrorCode.ContentControlBound,
            "target is data-bound; retry with bindingPolicy=detach_target to remove only its native data-binding element",
            candidate.Info.AnchorId);
    }

    private EditResult MutateContentControl(ContentControlCandidate candidate,
        ContentControlFillOptions? options, Action mutation)
    {
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            DetachTargetBindingIfRequested(candidate.Element, options);
            mutation();
            PromoteHyperlinkRelationships(candidate.Element);
            SweepOrphanedStoryRelationships(candidate.Owner.Part);
            UnidHelper.AssignToSelfAndDescendants(candidate.Element);
            ContentControlIdentity.AssignStableUnids(candidate.Owner.Part.GetXDocument().Root!);
            InvalidateProjectionCache();
            return new EditResult { Success = true,
                Modified = new[] { AnchorFromCandidate(candidate) } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, candidate.Info.AnchorId);
        }
    }

    private static void DetachTargetBindingIfRequested(XElement control,
        ContentControlFillOptions? options)
    {
        if (options?.BindingPolicy == ContentControlBindingPolicy.DetachTarget)
            foreach (var binding in FindDataBindings(control.Element(W.sdtPr)).ToList())
                binding.Remove();
    }

    private static void ReplaceControlWithPlainText(XElement control, string text,
        string? stateFont = null)
    {
        var content = control.Element(W.sdtContent)
            ?? throw new InvalidOperationException("content control has no w:sdtContent");
        var placement = DetectContentControlPlacement(control);
        var oldRunProperties = content.Descendants(W.r).Select(run => run.Element(W.rPr))
            .FirstOrDefault(value => value is not null);
        var run = new XElement(W.r,
            oldRunProperties is null ? null : new XElement(oldRunProperties),
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        if (!string.IsNullOrWhiteSpace(stateFont))
        {
            var runProperties = run.Element(W.rPr);
            if (runProperties is null)
            {
                runProperties = new XElement(W.rPr);
                run.AddFirst(runProperties);
            }
            var fonts = runProperties.Element(W.rFonts);
            if (fonts is null)
            {
                // CT_RPr is a strict sequence and the cloned rPr can already carry earlier
                // members (w:ins, w:del, w:rStyle, the move markers). Insert at the schema slot
                // rather than at position 0.
                fonts = new XElement(W.rFonts);
                WordprocessingMLUtil.InsertRPrChildInOrder(runProperties, fonts);
            }
            fonts.SetAttributeValue(W.ascii, stateFont);
            fonts.SetAttributeValue(W.hAnsi, stateFont);
            fonts.SetAttributeValue(W.eastAsia, stateFont);
            fonts.SetAttributeValue(W.cs, stateFont);
        }
        if (placement == ContentControlPlacement.Inline)
        {
            content.ReplaceNodes(run);
        }
        else if (placement == ContentControlPlacement.Block)
        {
            var oldParagraphProperties = content.Elements(W.p).Select(p => p.Element(W.pPr))
                .FirstOrDefault(value => value is not null);
            content.ReplaceNodes(new XElement(W.p,
                oldParagraphProperties is null ? null : new XElement(oldParagraphProperties), run));
        }
        else
        {
            throw new InvalidOperationException($"plain text cannot fill a {placement} content control");
        }
        control.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
    }

    /// <summary>Validate every consequence of replacing the target's complete payload before
    /// taking an undo snapshot. Existing bookmark ranges must be safe to remove, and hyperlinks
    /// in detached Markdown must still resolve after those ranges are gone.</summary>
    private EditResult? ValidateWholeContentReplacement(
        ContentControlCandidate candidate,
        IEnumerable<XElement>? replacement,
        string anchorId)
    {
        var content = candidate.Element.Element(W.sdtContent);
        if (content is null)
            return EditResult.Fail(EditErrorCode.ContentControlMalformed,
                "content control has no w:sdtContent", anchorId);
        if (ValidateBookmarkRemoval(new[] { content }, anchorId) is { } bookmarkError)
            return bookmarkError;
        return replacement is null
            ? null
            : ValidatePendingHyperlinks(replacement, anchorId, content);
    }

    private static bool ContainsNestedContentControl(XElement control) =>
        control.Element(W.sdtContent)?.Descendants(W.sdt).Any() == true;

    private static EditResult NestedFillError(string anchorId) =>
        EditResult.Fail(EditErrorCode.ContentControlNestedFillUnsupported,
            "whole-control fill is refused when the target contains nested controls; address the child control directly",
            anchorId);

    private IReadOnlyList<ContentControlCandidate> BuildContentControlRegistry(ProjectionScopes scopes)
    {
        var result = new List<ContentControlCandidate>();
        IReadOnlyList<ImageCandidate>? imageCandidates = null;
        var trackedBlocker = TrackedContentControlBlocker();
        var owners = OwnedPartRelationships.StoryParts(_doc!);
        var roots = owners.Select(owner => owner.Part.GetXDocument().Root)
            .Where(root => root is not null).Cast<XElement>().ToList();
        var identitiesByRoot = ContentControlIdentity.AssignStableUnids(roots, out _);
        foreach (var owner in owners)
        {
            if (!ScopeIncluded(owner.Scope, scopes)) continue;
            var root = owner.Part.GetXDocument().Root;
            if (root is null) continue;
            var identities = identitiesByRoot[root];
            var byElement = identities.ToDictionary(identity => identity.Element,
                identity => identity);
            var anchorByElement = identities.ToDictionary(identity => identity.Element,
                identity => $"sdt:{owner.Scope}:{identity.Unid}");
            foreach (var identity in identities)
            {
                var element = identity.Element;
                var malformed = ValidateContentControlStructure(element);
                var malformedAncestor = element.Ancestors(W.sdt)
                    .Select(ValidateContentControlStructure)
                    .FirstOrDefault(reason => reason is not null);
                var props = element.Element(W.sdtPr);
                var type = ClassifyContentControl(props);
                var placement = DetectContentControlPlacement(element);
                var binding = FindDataBinding(props);
                var parent = element.Ancestors(W.sdt).FirstOrDefault();
                var lockToken = (string?)props?.Element(ContentControlW + "lock")?.Attribute(W.val);
                // Discovery evaluates the mutation-time gates in the order
                // ResolveContentControlForMutation applies them, so the first reason an agent
                // reads here is the reason the mutation would actually return.
                string? unsupported = null;
                if (trackedBlocker is not null) unsupported = trackedBlocker;
                else if (malformed is not null) unsupported = malformed;
                else if (malformedAncestor is not null)
                    unsupported = $"ancestor content control is malformed: {malformedAncestor}";
                else if (!identity.HasValidNativeId) unsupported = "missing or invalid native w:sdtPr/w:id";
                else if (identity.IsDuplicateNativeId) unsupported = "duplicate native w:sdtPr/w:id in package";
                else if (placement == ContentControlPlacement.Unknown) unsupported = "unsupported or malformed OOXML placement";
                else if (type == ContentControlType.Unsupported) unsupported = "unsupported content-control family";

                bool targetBound = binding is not null;
                bool ancestorBound = element.Ancestors(W.sdt).Any(ancestor =>
                    FindDataBinding(ancestor.Element(W.sdtPr)) is not null);
                bool locked = element.AncestorsAndSelf(W.sdt).Any(control =>
                    (string?)control.Element(W.sdtPr)?.Element(ContentControlW + "lock")?.Attribute(W.val)
                        is "contentLocked" or "sdtContentLocked");
                bool wrapperLocked = type == ContentControlType.RepeatingSectionItem
                    && lockToken is "sdtLocked" or "sdtContentLocked";
                bool placementSupported = IsMutationPlacementSupported(type, placement);
                if (unsupported is null && !placementSupported)
                    unsupported = $"{type} mutation supports only inline and block content controls";
                if (unsupported is null && IsWholeControlFillType(type)
                    && ContainsNestedContentControl(element))
                    unsupported = "whole-control fill is unsupported when the target contains nested controls";
                if (unsupported is null)
                    unsupported = RepeatingMutationConstraint(element, type);
                if (unsupported is null && type == ContentControlType.Picture)
                {
                    imageCandidates ??= EnumerateImageCandidates(ProjectionScopes.All);
                    var pictureTarget = ResolvePictureContentControlTarget(element, imageCandidates);
                    if (pictureTarget.ErrorCode is not null)
                        unsupported = pictureTarget.Diagnostic;
                }
                if (unsupported is null)
                    unsupported = WholeContentBookmarkBlocker(element, type);
                bool defaultMutable = unsupported is null && !locked && !wrapperLocked
                    && !targetBound && !ancestorBound;

                var items = props?.Elements().FirstOrDefault(value =>
                        value.Name == W.dropDownList || value.Name == W.comboBox)
                    ?.Elements(W.listItem)
                    .Select(value => (string?)value.Attribute(ContentControlW + "value") ?? string.Empty).ToList()
                    ?? (IReadOnlyList<string>)Array.Empty<string>();
                var info = new ContentControlInfo
                {
                    AnchorId = anchorByElement[element],
                    Type = type,
                    Placement = placement,
                    NativeId = identity.NativeId,
                    Tag = (string?)props?.Element(W.tag)?.Attribute(W.val),
                    Alias = (string?)props?.Element(W.alias)?.Attribute(W.val),
                    Lock = lockToken,
                    IsShowingPlaceholder = props?.Element(W.showingPlcHdr) is not null,
                    Binding = binding is null ? null : new ContentControlBindingInfo(
                        (string?)binding.Attribute(W.storeItemID),
                        (string?)binding.Attribute(W.xpath),
                        (string?)binding.Attribute(W.prefixMappings)),
                    OwningPartUri = owner.PartUri,
                    Scope = owner.Scope,
                    ParentAnchorId = parent is not null && anchorByElement.TryGetValue(parent, out var parentId)
                        ? parentId : null,
                    Depth = element.Ancestors(W.sdt).Count(),
                    HasValidNativeId = identity.HasValidNativeId,
                    HasDuplicateNativeId = identity.IsDuplicateNativeId,
                    CanMutate = defaultMutable,
                    CanDetachTargetBinding = unsupported is null && targetBound && !ancestorBound
                        && !locked && !wrapperLocked,
                    UnsupportedReason = unsupported ?? (locked ? "content locked by target or ancestor"
                        : wrapperLocked ? "content-control wrapper is locked"
                        : ancestorBound ? "inside a data-bound ancestor"
                        : targetBound ? "target is data-bound; explicit detach_target is required" : null),
                    Text = string.Concat(element.Element(W.sdtContent)?.Descendants(W.t)
                        .Select(text => (string)text) ?? Enumerable.Empty<string>()),
                    ItemValues = items,
                };
                result.Add(new ContentControlCandidate(owner, element, byElement[element], info,
                    malformed, malformedAncestor));
            }
        }
        return result;
    }

    /// <summary>Apply the picture topology contract once for both discovery and mutation.
    /// A picture SDT is mutable only when it owns exactly one canonical embedded image.</summary>
    private static PictureContentControlTarget ResolvePictureContentControlTarget(
        XElement control, IReadOnlyList<ImageCandidate> imageCandidates)
    {
        var images = imageCandidates.Where(image =>
            ReferenceEquals(image.Outer, control)
            || image.Outer.Ancestors().Any(ancestor => ReferenceEquals(ancestor, control)))
            .ToList();
        if (images.Count != 1)
            return new PictureContentControlTarget(null, EditErrorCode.ContentControlMalformed,
                $"picture content control must contain exactly one mutable image; found {images.Count}");
        var image = images[0];
        if (image.Info.IsLinked)
            return new PictureContentControlTarget(null, EditErrorCode.LinkedImageReadOnly,
                "a linked picture content control is read-only");
        if (!image.Info.CanMutate || image.Blip is null)
            return new PictureContentControlTarget(null, EditErrorCode.UnsupportedImageMarkup,
                image.Info.UnsupportedReason ?? "picture content control uses unsupported image markup");
        return new PictureContentControlTarget(image, null, null);
    }

    private static bool ScopeIncluded(string scope, ProjectionScopes scopes) => scope switch
    {
        "body" => scopes.HasFlag(ProjectionScopes.Body),
        var value when value.StartsWith("hdr", StringComparison.Ordinal) => scopes.HasFlag(ProjectionScopes.Headers),
        var value when value.StartsWith("ftr", StringComparison.Ordinal) => scopes.HasFlag(ProjectionScopes.Footers),
        "fn" => scopes.HasFlag(ProjectionScopes.Footnotes),
        "en" => scopes.HasFlag(ProjectionScopes.Endnotes),
        "cmt" => scopes.HasFlag(ProjectionScopes.Comments),
        _ => false,
    };

    private static ContentControlType ClassifyContentControl(XElement? props)
    {
        if (props is null) return ContentControlType.Unsupported;
        var family = props.Elements().Where(element =>
            !ContentControlMetadata.Contains(element.Name)).ToList();
        if (family.Count == 0) return ContentControlType.RichText;
        return family.Count == 1 && ContentControlFamilies.TryGetValue(family[0].Name, out var type)
            ? type
            : ContentControlType.Unsupported;
    }

    private static string? ValidateContentControlStructure(XElement control)
    {
        var properties = control.Elements(W.sdtPr).ToList();
        if (properties.Count != 1)
            return $"content control must contain exactly one w:sdtPr; found {properties.Count}";
        var contents = control.Elements(W.sdtContent).ToList();
        if (contents.Count != 1)
            return $"content control must contain exactly one w:sdtContent; found {contents.Count}";
        var ids = properties[0].Elements(W.id).ToList();
        if (ids.Count != 1)
            return $"w:sdtPr must contain exactly one w:id; found {ids.Count}";
        if (!ContentControlIdentity.TryCanonicalizeNativeId(
                (string?)ids[0].Attribute(W.val), out _))
            return "w:sdtPr/w:id must have a signed 32-bit integer w:val";
        var locks = properties[0].Elements(ContentControlW + "lock").ToList();
        if (locks.Count > 1)
            return $"w:sdtPr must contain at most one w:lock; found {locks.Count}";
        if (locks.Count == 1 && (string?)locks[0].Attribute(W.val)
                is not ("unlocked" or "sdtLocked" or "contentLocked" or "sdtContentLocked"))
            return "w:sdtPr/w:lock must have a supported w:val";
        var family = properties[0].Elements().Where(element =>
            !ContentControlMetadata.Contains(element.Name)).ToList();
        if (family.Count > 1)
            return "w:sdtPr must contain at most one mutually exclusive content-control family marker";
        return null;
    }

    private static string? RepeatingMutationConstraint(XElement control, ContentControlType type)
    {
        if (type == ContentControlType.RepeatingSection)
        {
            var content = control.Element(W.sdtContent);
            var items = content?.Elements(W.sdt).Where(IsRepeatingSectionItem).ToList()
                ?? new List<XElement>();
            if (items.Count == 0 || content!.Elements().Any(element => element.Name != W.sdt
                    || !IsRepeatingSectionItem(element)))
                return "repeating section must contain only one or more direct repeating-section-item controls";
            if (FindUnsafeRepeatingCloneCarrier(items[^1]) is { } unsafeCarrier)
                return $"default repeating-item template contains clone-sensitive markup ({unsafeCarrier})";
            return null;
        }
        if (type != ContentControlType.RepeatingSectionItem) return null;
        var outer = control.Parent?.Parent;
        if (control.Parent?.Name != W.sdtContent || outer?.Name != W.sdt
            || !IsRepeatingSection(outer))
            return "repeating-section item is not a direct child of a repeating section";
        if (control.Parent.Elements(W.sdt).Count(IsRepeatingSectionItem) <= 1)
            return "a repeating section must retain at least one item";
        return null;
    }

    private static ContentControlPlacement DetectContentControlPlacement(XElement control)
    {
        var content = control.Element(W.sdtContent);
        if (content is null) return ContentControlPlacement.Unknown;
        var children = content.Elements().ToList();
        if (children.Count == 0)
            return DetectContentControlPlacementFromContext(control);
        // A nested SDT is valid in every placement grammar, so an sdt-only payload is
        // intrinsically ambiguous from children alone. Its parent context is authoritative.
        if (children.All(element => element.Name == W.sdt))
            return DetectContentControlPlacementFromContext(control);
        bool allInline = children.All(element => element.Name == W.r || element.Name == W.hyperlink
            || element.Name == W.fldSimple || element.Name == W.sdt || element.Name == W.smartTag
            || element.Name == W.bookmarkStart || element.Name == W.bookmarkEnd
            || element.Name == W.commentRangeStart || element.Name == W.commentRangeEnd);
        if (allInline && control.Ancestors(W.p).Any()) return ContentControlPlacement.Inline;
        if (children.All(element => element.Name == W.tr || element.Name == W.sdt))
            return ContentControlPlacement.Row;
        if (children.All(element => element.Name == W.tc || element.Name == W.sdt))
            return ContentControlPlacement.Cell;
        if (children.All(element => element.Name == W.p || element.Name == W.tbl
                || element.Name == W.sdt || element.Name == W.bookmarkStart || element.Name == W.bookmarkEnd))
            return ContentControlPlacement.Block;
        return ContentControlPlacement.Unknown;
    }

    /// <summary>An empty or nested-SDT-only sdtContent has no unambiguous child grammar from
    /// which to infer its typed SDT context. Use the nearest OOXML content-model boundary
    /// instead, walking transparently through nested SDTs and revision/custom-XML carriers.</summary>
    private static ContentControlPlacement DetectContentControlPlacementFromContext(XElement control)
    {
        foreach (var ancestor in control.Ancestors())
        {
            if (ancestor.Name == W.p) return ContentControlPlacement.Inline;
            if (ancestor.Name == W.tc || ancestor.Name == W.body || ancestor.Name == W.hdr
                || ancestor.Name == W.ftr || ancestor.Name == W.footnote
                || ancestor.Name == W.endnote || ancestor.Name == W.comment
                || ancestor.Name == W.txbxContent)
                return ContentControlPlacement.Block;
            if (ancestor.Name == W.tr) return ContentControlPlacement.Cell;
            if (ancestor.Name == W.tbl) return ContentControlPlacement.Row;
        }
        return ContentControlPlacement.Unknown;
    }

    private static bool IsRepeatingSection(XElement control) =>
        control.Element(W.sdtPr)?.Element(ContentControlW15 + "repeatingSection") is not null;

    private static bool IsRepeatingSectionItem(XElement control) =>
        control.Element(W.sdtPr)?.Element(ContentControlW15 + "repeatingSectionItem") is not null;

    private void AssignFreshContentControlIds(XElement root)
    {
        var used = new HashSet<int>();
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
        foreach (var control in owner.Part.GetXDocument().Descendants(W.sdt))
        {
            var raw = (string?)control.Element(W.sdtPr)?.Element(W.id)?.Attribute(W.val);
            if (int.TryParse(raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var id))
                used.Add(id);
        }
        int next = 1;
        foreach (var control in root.DescendantsAndSelf(W.sdt))
        {
            while (used.Contains(next) && next < int.MaxValue) next++;
            if (used.Contains(next)) throw new InvalidOperationException("no unused content-control id remains");
            used.Add(next);
            var props = control.Elements(W.sdtPr).Single();
            var id = props.Elements(W.id).Single();
            id.SetAttributeValue(W.val, next.ToString(CultureInfo.InvariantCulture));
            next++;
        }
    }

    /// <summary>
    /// Freshen the identity half of Word's paragraph identity pair on a clone. Word 2013+ writes
    /// <c>w14:paraId</c> on essentially every <c>w:p</c>, so refusing to clone a paragraph that
    /// carries one would make repeating sections inert on real templates; a paraId is
    /// package-unique, so the clone gets fresh values from the shared allocator instead.
    /// </summary>
    /// <remarks>
    /// <c>w14:textId</c> is deliberately left verbatim: it is a hash of the paragraph's text
    /// rather than an identity, and Word itself emits the same value for two paragraphs with the
    /// same content — which is exactly what a clone is. The clone gate still refuses items
    /// carrying markup whose identity <em>is</em> semantic (bookmarks, comment and note
    /// references, permissions, custom-XML and tracked-revision ranges).
    /// </remarks>
    private void AssignFreshParagraphIds(XElement root)
    {
        var carriers = root.DescendantsAndSelf().Attributes(W14.paraId).ToList();
        if (carriers.Count == 0) return;
        var allocator = new CommentOps.ParaIdAllocator(_doc!.MainDocumentPart!);
        foreach (var attribute in carriers) attribute.SetValue(allocator.Next());
    }

    private void AssignFreshDocumentPropertyIds(XElement root)
    {
        var used = OwnedPartRelationships.StoryParts(_doc!)
            .SelectMany(owner => owner.Part.GetXDocument().Descendants(WP.docPr))
            .Select(element => uint.TryParse((string?)element.Attribute("id"),
                NumberStyles.None, CultureInfo.InvariantCulture, out var id) ? id : 0)
            .Where(id => id != 0).ToHashSet();
        uint next = 1;
        foreach (var docPr in root.Descendants(WP.docPr))
        {
            while (next != 0 && used.Contains(next)) next++;
            if (next == 0)
                throw new InvalidOperationException("no globally available wp:docPr id remains");
            docPr.SetAttributeValue("id", next.ToString(CultureInfo.InvariantCulture));
            used.Add(next++);
        }
    }

    private static string? FindUnsafeRepeatingCloneCarrier(XElement item)
    {
        foreach (var control in item.DescendantsAndSelf(W.sdt))
            if (ValidateContentControlStructure(control) is { } malformed)
                return $"malformed content control: {malformed}";
        var revision = item.Descendants().FirstOrDefault(RevisionOps.IsRecognizedRevisionMarker);
        if (revision is not null) return $"tracked revision {revision.Name.LocalName}";
        var unsafeNames = new HashSet<XName>
        {
            W.bookmarkStart, W.bookmarkEnd, W.commentRangeStart, W.commentRangeEnd,
            W.commentReference, W.footnoteReference, W.endnoteReference,
            ContentControlW + "permStart", ContentControlW + "permEnd",
            ContentControlW + "customXml",
            ContentControlW + "customXmlInsRangeStart", ContentControlW + "customXmlInsRangeEnd",
            ContentControlW + "customXmlDelRangeStart", ContentControlW + "customXmlDelRangeEnd",
            ContentControlW + "customXmlMoveFromRangeStart", ContentControlW + "customXmlMoveFromRangeEnd",
            ContentControlW + "customXmlMoveToRangeStart", ContentControlW + "customXmlMoveToRangeEnd",
            ContentControlW + "moveFromRangeStart", ContentControlW + "moveFromRangeEnd",
            ContentControlW + "moveToRangeStart", ContentControlW + "moveToRangeEnd",
            ContentControlW + "moveFrom", ContentControlW + "moveTo",
        };
        var unsafeElement = item.Descendants().FirstOrDefault(element => unsafeNames.Contains(element.Name));
        return unsafeElement?.Name.LocalName;
    }

    private static IEnumerable<XElement> FindDataBindings(XElement? properties) =>
        properties?.Elements().Where(element =>
            element.Name == W.dataBinding || element.Name == ContentControlW15 + "dataBinding")
        ?? Enumerable.Empty<XElement>();

    private static XElement? FindDataBinding(XElement? properties) =>
        FindDataBindings(properties).FirstOrDefault();

    private static bool IsMutationPlacementSupported(ContentControlType type,
        ContentControlPlacement placement) => type switch
    {
        ContentControlType.PlainText or ContentControlType.RichText
            or ContentControlType.Checkbox or ContentControlType.Date
            or ContentControlType.DropDownList or ContentControlType.ComboBox =>
            placement is ContentControlPlacement.Inline or ContentControlPlacement.Block,
        _ => placement != ContentControlPlacement.Unknown,
    };

    /// <summary>The families whose fill discards and rebuilds the whole <c>w:sdtContent</c>
    /// payload, and therefore takes the bookmark-removal gate.</summary>
    private static bool IsWholeContentReplacementType(ContentControlType type) => type is
        ContentControlType.PlainText or ContentControlType.RichText
        or ContentControlType.Checkbox or ContentControlType.Date
        or ContentControlType.DropDownList or ContentControlType.ComboBox;

    /// <summary>Every family filled as a unit, adding picture — whose fill retargets only the
    /// blip relationship and so leaves existing payload markup in place.</summary>
    private static bool IsWholeControlFillType(ContentControlType type) =>
        IsWholeContentReplacementType(type) || type == ContentControlType.Picture;

    private static bool TryParseHexScalar(string? value, out int scalar)
    {
        scalar = 0;
        return !string.IsNullOrEmpty(value)
            && int.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out scalar)
            && scalar is >= 0 and <= 0x10ffff && (scalar < 0xd800 || scalar > 0xdfff);
    }

    private static Anchor AnchorFromCandidate(ContentControlCandidate candidate) =>
        new(candidate.Info.AnchorId, "sdt", candidate.Info.Scope, candidate.Identity.Unid);
}
