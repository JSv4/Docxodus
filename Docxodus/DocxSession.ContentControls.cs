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

    /// <summary>Remove only the selected control's own w:dataBinding before filling it.
    /// A binding on any ancestor still fails closed.</summary>
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

    private sealed record ContentControlCandidate(
        OwnedPartRelationships.Owner Owner,
        XElement Element,
        ContentControlIdentity.Entry Identity,
        ContentControlInfo Info);

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
        if (checkedElement is null)
        {
            checkedElement = new XElement(ContentControlW14 + "checked");
            checkbox.AddFirst(checkedElement);
        }

        var stateElement = checkbox.Element(isChecked
            ? ContentControlW14 + "checkedState"
            : ContentControlW14 + "uncheckedState");
        var fallback = isChecked ? 0x2612 : 0x2610;
        var glyph = TryParseHexScalar((string?)stateElement?.Attribute(ContentControlW14 + "val"),
            out var scalar) ? char.ConvertFromUtf32(scalar) : char.ConvertFromUtf32(fallback);

        return MutateContentControl(candidate, options, () =>
        {
            checkedElement.SetAttributeValue(ContentControlW14 + "val", isChecked ? "1" : "0");
            ReplaceControlWithPlainText(candidate.Element, glyph);
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
        var matches = list.Elements(W.listItem).Where(item =>
            string.Equals((string?)item.Attribute(ContentControlW + "value"), value, StringComparison.Ordinal)
            || string.Equals((string?)item.Attribute(W.displayText), value, StringComparison.Ordinal)).ToList();
        if (matches.Count != 1)
            return EditResult.Fail(EditErrorCode.InvalidContentControlValue,
                matches.Count == 0
                    ? $"content control has no list item matching '{value}'"
                    : $"content control has multiple list items matching '{value}'", anchorId);
        var display = (string?)matches[0].Attribute(W.displayText)
            ?? (string?)matches[0].Attribute(ContentControlW + "value") ?? string.Empty;
        return MutateContentControl(candidate, options,
            () => ReplaceControlWithPlainText(candidate.Element, display));
    }

    public EditResult FillContentControlPicture(string anchorId, byte[] imageBytes,
        ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(anchorId, ContentControlType.Picture, options,
            out var candidate, out var error) is false) return error!;
        var binary = ValidateImageBytes(imageBytes, anchorId);
        if (binary.Error is not null) return binary.Error;
        var images = EnumerateImageCandidates(ProjectionScopes.All).Where(image =>
            ReferenceEquals(image.Outer, candidate!.Element)
            || image.Outer.Ancestors().Any(ancestor => ReferenceEquals(ancestor, candidate!.Element)))
            .ToList();
        if (images.Count != 1)
            return EditResult.Fail(EditErrorCode.ContentControlMalformed,
                $"picture content control must contain exactly one mutable image; found {images.Count}", anchorId);
        var image = images[0];
        if (image.Info.IsLinked)
            return EditResult.Fail(EditErrorCode.LinkedImageReadOnly,
                "a linked picture content control is read-only", anchorId);
        if (!image.Info.CanMutate || image.Blip is null)
            return EditResult.Fail(EditErrorCode.UnsupportedImageMarkup,
                image.Info.UnsupportedReason ?? "picture content control uses unsupported image markup", anchorId);

        return MutateContentControl(candidate!, options, () =>
        {
            var relationship = OwnedPartRelationships.FindOrAddImagePart(_doc!, candidate!.Owner.Part,
                imageBytes, binary.ContentType!, binary.Format);
            image.Blip.SetAttributeValue(ImageR + "embed", relationship.RelationshipId);
            candidate.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
            OwnedPartRelationships.SweepOrphanedImages(candidate.Owner.Part, ImageR + "embed", ImageR + "link");
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
                $"repeating item contains clone-sensitive markup ({unsafeCarrier.Name.LocalName})",
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
            foreach (var docPr in clone.Descendants(WP.docPr))
                docPr.SetAttributeValue("id", NextDocumentPropertyId().ToString(CultureInfo.InvariantCulture));
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
            return MutateContentControl(candidate, options,
                () => ReplaceControlWithPlainText(candidate.Element, payload));

        var parsed = MarkdownPayloadParser.Parse(payload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
        if (parsed.Blocks.Count == 0)
            parsed = MarkdownPayloadParser.Parse("");
        if (candidate.Info.Placement == ContentControlPlacement.Inline && parsed.Blocks.Count != 1)
            return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                "an inline rich-text control accepts exactly one markdown block", anchorId);
        if (candidate.Info.Placement is not (ContentControlPlacement.Inline or ContentControlPlacement.Block))
            return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                "rich-text fill supports only inline and block content controls", anchorId);

        return MutateContentControl(candidate, options, () =>
        {
            var content = candidate.Element.Element(W.sdtContent)!;
            if (candidate.Info.Placement == ContentControlPlacement.Inline)
            {
                var block = parsed.Blocks.Count == 0
                    ? new ParsedBlock(ParserBlockKind.Paragraph, 0, Array.Empty<XElement>())
                    : parsed.Blocks[0];
                content.ReplaceNodes(block.RunElements.Select(element => new XElement(element)));
            }
            else
            {
                var blocks = parsed.Blocks.Select(BuildParagraphFromParsedBlock).ToList();
                if (blocks.Count == 0) blocks.Add(new XElement(W.p));
                content.ReplaceNodes(blocks);
            }
            candidate.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
            PromoteHyperlinkRelationships(candidate.Element);
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
        if (_trackedChanges == TrackedChangeMode.RenderInline)
        {
            error = EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
                "whole content-control fills cannot be represented faithfully as tracked revisions; use surgical text operations inside the control or switch modes",
                anchorId);
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
        if (!candidate.Identity.HasMutableIdentity)
        {
            error = EditResult.Fail(EditErrorCode.ContentControlMalformed,
                candidate.Info.UnsupportedReason ?? "content control has no unique valid native w:id", anchorId);
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
            control.Element(W.sdtPr)?.Element(W.dataBinding) is not null).ToList();
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
            "target is data-bound; retry with bindingPolicy=detach_target to remove only its w:dataBinding",
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
            control.Element(W.sdtPr)?.Element(W.dataBinding)?.Remove();
    }

    private static void ReplaceControlWithPlainText(XElement control, string text)
    {
        var content = control.Element(W.sdtContent)
            ?? throw new InvalidOperationException("content control has no w:sdtContent");
        var placement = DetectContentControlPlacement(control);
        var oldRunProperties = content.Descendants(W.r).Select(run => run.Element(W.rPr))
            .FirstOrDefault(value => value is not null);
        var run = new XElement(W.r,
            oldRunProperties is null ? null : new XElement(oldRunProperties),
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
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

    private static bool ContainsNestedContentControl(XElement control) =>
        control.Element(W.sdtContent)?.Descendants(W.sdt).Any() == true;

    private static EditResult NestedFillError(string anchorId) =>
        EditResult.Fail(EditErrorCode.ContentControlNestedFillUnsupported,
            "whole-control fill is refused when the target contains nested controls; address the child control directly",
            anchorId);

    private IReadOnlyList<ContentControlCandidate> BuildContentControlRegistry(ProjectionScopes scopes)
    {
        var result = new List<ContentControlCandidate>();
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
        {
            if (!ScopeIncluded(owner.Scope, scopes)) continue;
            var root = owner.Part.GetXDocument().Root;
            if (root is null) continue;
            var identities = ContentControlIdentity.AssignStableUnids(root);
            var byElement = identities.ToDictionary(identity => identity.Element,
                identity => identity);
            var anchorByElement = identities.ToDictionary(identity => identity.Element,
                identity => $"sdt:{owner.Scope}:{identity.Unid}");
            foreach (var identity in identities)
            {
                var element = identity.Element;
                var props = element.Element(W.sdtPr);
                var type = ClassifyContentControl(props);
                var placement = DetectContentControlPlacement(element);
                var binding = props?.Element(W.dataBinding);
                var parent = element.Ancestors(W.sdt).FirstOrDefault();
                var lockToken = (string?)props?.Element(ContentControlW + "lock")?.Attribute(W.val);
                string? unsupported = null;
                if (!identity.HasValidNativeId) unsupported = "missing or invalid native w:sdtPr/w:id";
                else if (identity.IsDuplicateNativeId) unsupported = "duplicate native w:sdtPr/w:id in owning story";
                else if (placement == ContentControlPlacement.Unknown) unsupported = "unsupported or malformed OOXML placement";
                else if (type == ContentControlType.Unsupported) unsupported = "unsupported content-control family";

                bool targetBound = binding is not null;
                bool ancestorBound = element.Ancestors(W.sdt).Any(ancestor =>
                    ancestor.Element(W.sdtPr)?.Element(W.dataBinding) is not null);
                bool locked = element.AncestorsAndSelf(W.sdt).Any(control =>
                    (string?)control.Element(W.sdtPr)?.Element(ContentControlW + "lock")?.Attribute(W.val)
                        is "contentLocked" or "sdtContentLocked");
                bool defaultMutable = unsupported is null && !locked && !targetBound && !ancestorBound;

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
                    CanDetachTargetBinding = unsupported is null && targetBound && !ancestorBound && !locked,
                    UnsupportedReason = unsupported ?? (locked ? "content locked by target or ancestor"
                        : ancestorBound ? "inside a data-bound ancestor"
                        : targetBound ? "target is data-bound; explicit detach_target is required" : null),
                    Text = string.Concat(element.Element(W.sdtContent)?.Descendants(W.t)
                        .Select(text => (string)text) ?? Enumerable.Empty<string>()),
                    ItemValues = items,
                };
                result.Add(new ContentControlCandidate(owner, element, byElement[element], info));
            }
        }
        return result;
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
        if (props.Element(ContentControlW14 + "checkbox") is not null) return ContentControlType.Checkbox;
        if (props.Element(ContentControlW15 + "repeatingSection") is not null) return ContentControlType.RepeatingSection;
        if (props.Element(ContentControlW15 + "repeatingSectionItem") is not null) return ContentControlType.RepeatingSectionItem;
        if (props.Element(W.picture) is not null) return ContentControlType.Picture;
        if (props.Element(W.date) is not null) return ContentControlType.Date;
        if (props.Element(W.dropDownList) is not null) return ContentControlType.DropDownList;
        if (props.Element(W.comboBox) is not null) return ContentControlType.ComboBox;
        if (props.Element(W.text) is not null) return ContentControlType.PlainText;
        if (props.Element(ContentControlW + "richText") is not null) return ContentControlType.RichText;
        var knownMetadata = new HashSet<XName> { W.id, W.tag, W.alias, W.dataBinding,
            W.showingPlcHdr, ContentControlW + "lock", ContentControlW + "placeholder",
            ContentControlW + "temporary", ContentControlW + "appearance",
            ContentControlW + "color" };
        return props.Elements().All(element => knownMetadata.Contains(element.Name))
            ? ContentControlType.RichText
            : ContentControlType.Unsupported;
    }

    private static ContentControlPlacement DetectContentControlPlacement(XElement control)
    {
        var content = control.Element(W.sdtContent);
        if (content is null) return ContentControlPlacement.Unknown;
        var children = content.Elements().ToList();
        if (children.Count == 0)
        {
            if (control.Ancestors(W.p).Any()) return ContentControlPlacement.Inline;
            return ContentControlPlacement.Block;
        }
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
            var props = control.Element(W.sdtPr);
            if (props is null)
            {
                props = new XElement(W.sdtPr);
                control.AddFirst(props);
            }
            var id = props.Element(W.id);
            if (id is null)
            {
                id = new XElement(W.id);
                props.Add(id);
            }
            id.SetAttributeValue(W.val, next.ToString(CultureInfo.InvariantCulture));
            next++;
        }
    }

    private static XElement? FindUnsafeRepeatingCloneCarrier(XElement item)
    {
        var unsafeNames = new HashSet<XName>
        {
            W.bookmarkStart, W.bookmarkEnd, W.commentRangeStart, W.commentRangeEnd,
            W.commentReference, W.footnoteReference, W.endnoteReference,
            ContentControlW + "permStart", ContentControlW + "permEnd",
            ContentControlW + "customXmlInsRangeStart", ContentControlW + "customXmlInsRangeEnd",
            ContentControlW + "customXmlDelRangeStart", ContentControlW + "customXmlDelRangeEnd",
        };
        return item.Descendants().FirstOrDefault(element => unsafeNames.Contains(element.Name));
    }

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
