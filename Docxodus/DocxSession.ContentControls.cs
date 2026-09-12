// Copyright (c) John Scrudato IV. All rights reserved.
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

/// <summary>How a whole-control fill treats the nested controls inside its target (issue #763).</summary>
public enum ContentControlNestedPolicy
{
    /// <summary>Refuse the fill when the target contains nested controls (default).</summary>
    Refuse = 0,

    /// <summary>Keep every nested control in place; the payload replaces only the content
    /// outside them. <see cref="ContentControlFillOptions.ChildFills"/> may fill named
    /// textual children in the same operation.</summary>
    Preserve = 1,

    /// <summary>Replace the whole payload, nested controls included. Each removed control
    /// is reported in <c>Removed</c>; a locked or data-bound nested control refuses.</summary>
    Replace = 2,
}

public sealed record ContentControlFillOptions
{
    public ContentControlBindingPolicy BindingPolicy { get; init; } = ContentControlBindingPolicy.Preserve;

    public ContentControlNestedPolicy NestedControls { get; init; } = ContentControlNestedPolicy.Refuse;

    /// <summary>
    /// With <see cref="ContentControlNestedPolicy.Preserve"/>: plain-text fills for nested
    /// text or rich-text controls of the target, keyed by their <c>sdt</c> anchor. Every key
    /// must be a nested textual control of the target and must pass its own gates; otherwise
    /// the whole operation fails without mutating.
    /// </summary>
    public IReadOnlyDictionary<string, string>? ChildFills { get; init; }
}

/// <summary>
/// Whether one operation would succeed on a control right now, evaluated by the same gates
/// the operation applies. <see cref="NestedControls"/> names the nested policy the entry
/// describes when the target contains nested controls; null otherwise.
/// </summary>
public sealed record ContentControlOperationSupport(
    string Operation,
    string? NestedControls,
    bool CanMutate,
    string? Reason);

/// <summary>The content-control operations, one per public mutation method.</summary>
internal enum ContentControlOperation
{
    FillText,
    FillRichText,
    SetChecked,
    SetDate,
    SelectItem,
    FillPicture,
    AddRepeatingItem,
    RemoveRepeatingItem,
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

    /// <summary>Anchors of the controls nested anywhere inside this control's payload, in story order.</summary>
    public IReadOnlyList<string> NestedControlAnchorIds { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Per-operation support for this control's family, including the nested-policy variants
    /// of a fill when the target contains nested controls and the session's tracked-change
    /// mode. <see cref="CanMutate"/> is the entry for the family's default operation and options.
    /// </summary>
    public IReadOnlyList<ContentControlOperationSupport> Operations { get; init; } =
        Array.Empty<ContentControlOperationSupport>();
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
        if (ResolveContentControlForMutation(anchorId, ContentControlOperation.SetChecked, options,
            out var candidate, out var error) is false) return error!;

        var checkbox = candidate!.Element.Element(W.sdtPr)?.Element(ContentControlW14 + "checkbox");
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
        if (ResolveContentControlForMutation(anchorId, ContentControlOperation.SetDate, options,
            out var candidate, out var error) is false) return error!;
        var date = candidate!.Element.Element(W.sdtPr)?.Element(W.date);
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
        if (ResolveContentControlForMutation(anchorId, ContentControlOperation.SelectItem, options,
            out var candidate, out var error) is false) return error!;
        var props = candidate!.Element.Element(W.sdtPr)!;
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
        return MutateContentControl(candidate, options, () =>
        {
            list.SetAttributeValue(W.lastValue, selectedValue);
            ReplaceControlWithPlainText(candidate.Element, display);
        });
    }

    public EditResult FillContentControlPicture(string anchorId, byte[] imageBytes,
        ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(anchorId, ContentControlOperation.FillPicture, options,
            out var candidate, out var error) is false) return error!;
        var binary = ValidateImageBytes(imageBytes, anchorId);
        if (binary.Error is not null) return binary.Error;
        var target = ResolvePictureContentControlTarget(candidate!.Element,
            EnumerateImageCandidates(ProjectionScopes.All));
        if (target.ErrorCode is { } errorCode)
            return EditResult.Fail(errorCode, target.Diagnostic!, anchorId);
        var image = target.Image!;
        var blip = image.Blip!;
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;

        return MutateContentControl(candidate!, options, () =>
        {
            var relationship = OwnedPartRelationships.FindOrAddImagePart(_doc!, candidate!.Owner.Part,
                imageBytes, binary.ContentType!, binary.Format);
            if (tracked)
            {
                // Word's own shape for a tracked picture swap: the old run (drawing included)
                // becomes a deletion and a fresh run carrying the new blip an insertion, so
                // accept keeps only the new image and reject only the original — both media
                // parts and their relationships stay referenced until the revision resolves.
                var run = image.Outer.Parent!;
                var replacement = new XElement(run);
                foreach (var element in replacement.DescendantsAndSelf())
                    element.Attribute(PtOpenXml.Unid)?.Remove();
                replacement.Descendants(A.blip).Single()
                    .SetAttributeValue(ImageR + "embed", relationship.RelationshipId);
                AssignFreshDocumentPropertyIds(replacement);
                var stamp = NewRevisionStamp();
                var deletion = CreateRevisionEnvelope(W.del, stamp);
                run.ReplaceWith(deletion);
                deletion.Add(run);
                deletion.AddAfterSelf(CreateRevisionEnvelope(W.ins, stamp, replacement));
            }
            else
            {
                blip.SetAttributeValue(ImageR + "embed", relationship.RelationshipId);
                OwnedPartRelationships.SweepOrphanedImages(candidate.Owner.Part);
            }
            candidate.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
        });
    }

    /// <summary>Clone one direct repeating-section item. The new item is inserted after
    /// <paramref name="afterItemAnchorId"/>, or after the final item when omitted. Under
    /// <c>render_inline</c> the clone is a tracked content-control insertion: two paired
    /// custom-XML insertion ranges cross its tags and its paragraphs are inserted content.</summary>
    public EditResult AddRepeatingSectionItem(string sectionAnchorId,
        string? afterItemAnchorId = null, ContentControlFillOptions? options = null)
    {
        if (ResolveContentControlForMutation(sectionAnchorId, ContentControlOperation.AddRepeatingItem,
            options, out var section, out var error) is false) return error!;
        var content = section!.Element.Element(W.sdtContent)!;
        var items = content.Elements(W.sdt).Where(IsRepeatingSectionItem).ToList();

        XElement template;
        if (afterItemAnchorId is null) template = DefaultRepeatingTemplate(items);
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
        if (TrackedRepeatingInsertBlocker(template) is { } trackedReason)
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported, trackedReason, sectionAnchorId);

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
            if (_trackedChanges == TrackedChangeMode.RenderInline)
                MarkStructuredBlockAsTrackedInserted(clone, NewRevisionStamp());
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

    /// <summary>Remove one direct repeating-section item. Under <c>render_inline</c> the item is
    /// a tracked content-control deletion (the shape <c>DeleteRange</c> uses), so it stays live
    /// until the revision resolves and is reported in <c>Modified</c> rather than <c>Removed</c>.</summary>
    public EditResult RemoveRepeatingSectionItem(string itemAnchorId)
    {
        if (ResolveContentControlForMutation(itemAnchorId, ContentControlOperation.RemoveRepeatingItem,
            options: null, out var item, out var error) is false) return error!;
        var outer = item!.Element.Parent!.Parent!;
        var parentCandidate = BuildContentControlRegistry(ProjectionScopes.All).First(value =>
            ReferenceEquals(value.Element, outer));
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var itemAnchor = AnchorFromCandidate(item);
            if (tracked)
            {
                MarkStructuredBlockAsTrackedDeleted(item.Element, NewRevisionStamp());
                InvalidateProjectionCache();
                return new EditResult { Success = true,
                    Modified = new[] { itemAnchor, AnchorFromCandidate(parentCandidate) } };
            }
            item.Element.Remove();
            SweepOrphanedStoryRelationships(item.Owner.Part);
            InvalidateProjectionCache();
            return new EditResult { Success = true, Removed = new[] { itemAnchor },
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
        options ??= new ContentControlFillOptions();
        var operation = rich ? ContentControlOperation.FillRichText : ContentControlOperation.FillText;
        if (ResolveContentControlForMutation(anchorId, operation, options,
            out var candidate, out var error) is false) return error!;
        var content = candidate!.Element.Element(W.sdtContent)!;
        var placement = candidate.Info.Placement;
        var nested = NestedContentControls(candidate.Element);
        var policy = nested.Count == 0 ? ContentControlNestedPolicy.Refuse : options.NestedControls;
        var preserveNested = policy == ContentControlNestedPolicy.Preserve;
        var replaced = ReplacedPayloadNodes(content, preserveNested);

        IReadOnlyList<XElement> payloadNodes;
        if (!rich)
        {
            payloadNodes = new[] { PlainTextPayloadNode(placement, replaced, payload) };
        }
        else
        {
            var parsed = MarkdownPayloadParser.Parse(payload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
            if (parsed.Blocks.Count == 0)
                parsed = ParseResult.Ok(new[]
                {
                    new ParsedBlock(ParserBlockKind.Paragraph, 0,
                        new[] { new XElement(W.r) }),
                });
            if (placement == ContentControlPlacement.Inline && parsed.Blocks.Count != 1)
                return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                    "an inline rich-text control accepts exactly one markdown block", anchorId);
            var runs = parsed.Blocks.SelectMany(block => block.RunElements).ToList();
            if (ValidatePendingHyperlinks(runs, anchorId, content) is { } hyperlinkError)
                return hyperlinkError;
            payloadNodes = placement == ContentControlPlacement.Inline
                ? parsed.Blocks[0].RunElements.Select(element => new XElement(element)).ToList()
                : parsed.Blocks.Select(BuildParagraphFromParsedBlock).ToList();
        }

        // Child fills were validated by the gate; resolve their targets before the snapshot.
        var registry = BuildContentControlRegistry(ProjectionScopes.All);
        var childFills = (options.ChildFills ?? new Dictionary<string, string>())
            .Select(pair => (Child: registry.First(value =>
                string.Equals(value.Info.AnchorId, pair.Key, StringComparison.Ordinal)), Text: pair.Value))
            .ToList();
        var nestedAnchors = nested
            .Select(control => registry.FirstOrDefault(value => ReferenceEquals(value.Element, control)))
            .Where(value => value is not null)
            .Select(value => AnchorFromCandidate(value!))
            .ToList();
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;

        return MutateContentControl(candidate, options, () =>
        {
            ReplacePayload(content, placement, preserveNested, payloadNodes, tracked);
            candidate.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
            foreach (var (child, text) in childFills)
            {
                var childContent = child.Element.Element(W.sdtContent)!;
                var childReplaced = ReplacedPayloadNodes(childContent, preserveNested: false);
                ReplacePayload(childContent, child.Info.Placement, preserveNested: false,
                    new[] { PlainTextPayloadNode(child.Info.Placement, childReplaced, text) }, tracked);
                child.Element.Element(W.sdtPr)?.Element(W.showingPlcHdr)?.Remove();
            }
        },
        removed: policy == ContentControlNestedPolicy.Replace && !tracked ? nestedAnchors : null,
        alsoModified: policy == ContentControlNestedPolicy.Replace && tracked
            ? nestedAnchors
            : childFills.Select(pair => AnchorFromCandidate(pair.Child)).ToList());
    }

    /// <summary>The plain-text payload the fill writes: one run (inline) or one paragraph (block),
    /// carrying the first replaced run's properties — and, for a block, the first replaced
    /// paragraph's properties — so a fill keeps the control's existing formatting.</summary>
    private static XElement PlainTextPayloadNode(ContentControlPlacement placement,
        IReadOnlyList<XElement> replaced, string text)
    {
        var oldRunProperties = replaced.SelectMany(node => node.DescendantsAndSelf(W.r))
            .Select(run => run.Element(W.rPr)).FirstOrDefault(value => value is not null);
        var run = new XElement(W.r,
            oldRunProperties is null ? null : new XElement(oldRunProperties),
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        if (placement == ContentControlPlacement.Inline) return run;
        var oldParagraphProperties = replaced.Where(node => node.Name == W.p)
            .Select(paragraph => paragraph.Element(W.pPr)).FirstOrDefault(value => value is not null);
        return new XElement(W.p,
            oldParagraphProperties is null ? null : new XElement(oldParagraphProperties), run);
    }

    /// <summary>
    /// Replace the target's own content with <paramref name="payload"/>. Every direct payload
    /// child is replaced, except — when <paramref name="preserveNested"/> — a nested control,
    /// a table that holds one, and the part of a paragraph that is one: such a paragraph loses
    /// only its own inline content and stays as the nested control's container. The first
    /// paragraph with replaced content hosts the payload's first block (keeping its paragraph
    /// properties); further blocks follow it as new paragraphs; with no host the payload goes
    /// at the end. Under <c>render_inline</c> nothing is discarded: replaced runs become
    /// <c>w:del</c>, replaced paragraphs/tables/nested wrappers take their ordinary deletion
    /// markup, and the payload is inserted content — the host paragraph carrying both the
    /// deletion and the first inserted block, Word's own shape for retyping a control — so
    /// accept yields exactly the payload and reject exactly the original.
    /// </summary>
    private void ReplacePayload(XElement content, ContentControlPlacement placement,
        bool preserveNested, IReadOnlyList<XElement> payload, bool tracked)
    {
        var stamp = tracked ? NewRevisionStamp() : default;
        if (placement == ContentControlPlacement.Inline)
        {
            var replaced = content.Elements()
                .Where(node => !IsPreservedPayloadNode(node, preserveNested) && !IsRangeMarker(node))
                .ToList();
            ReplaceInlineNodes(content, replaced, payload, tracked, stamp);
            return;
        }

        XElement? host = null;
        var remaining = payload.ToList();
        foreach (var child in content.Elements().ToList())
        {
            if (IsPreservedPayloadNode(child, preserveNested) || IsRangeMarker(child)) continue;
            if (child.Name == W.p)
            {
                var hostsNested = preserveNested && child.Descendants(W.sdt).Any();
                var inline = child.Elements()
                    .Where(node => node.Name != W.pPr && !IsRangeMarker(node)
                        && !(hostsNested && (node.Name == W.sdt || node.Descendants(W.sdt).Any())))
                    .ToList();
                if (host is null && remaining.Count > 0)
                {
                    host = child;
                    var first = remaining[0];
                    remaining.RemoveAt(0);
                    ReplaceInlineNodes(child, inline, first.Elements()
                        .Where(node => node.Name != W.pPr).Select(node => new XElement(node)).ToList(),
                        tracked, stamp);
                    continue;
                }
                if (!hostsNested)
                {
                    if (tracked) MarkParagraphAsTrackedDeleted(child, stamp);
                    else child.Remove();
                }
                else
                {
                    ReplaceInlineNodes(child, inline, Array.Empty<XElement>(), tracked, stamp);
                }
                continue;
            }
            if (tracked) MarkTrackedStructuredContentChild(child, stamp);
            else child.Remove();
        }

        if (remaining.Count == 0) return;
        if (tracked)
            foreach (var paragraph in remaining)
                MarkParagraphContentAndMark(paragraph, W.ins, stamp.Author, stamp.Date);
        if (host is not null) host.AddAfterSelf(remaining);
        else content.Add(remaining);
    }

    /// <summary>Replace <paramref name="replaced"/> (inline children of <paramref name="container"/>)
    /// with <paramref name="payload"/> at the first replaced position, or at the end when nothing
    /// is replaced; tracked, the replaced runs become one <c>w:del</c> and the payload one
    /// <c>w:ins</c> right after it.</summary>
    private void ReplaceInlineNodes(XElement container, IReadOnlyList<XElement> replaced,
        IReadOnlyList<XElement> payload, bool tracked, RevisionStamp stamp)
    {
        if (!tracked)
        {
            if (payload.Count > 0)
            {
                if (replaced.Count > 0) replaced[0].AddBeforeSelf(payload);
                else container.Add(payload);
            }
            foreach (var node in replaced) node.Remove();
            return;
        }

        // Un-inserting the author's own earlier payload detaches those nodes, so the slot the
        // insertion goes into is remembered as the node before the replaced range, not as one
        // of the replaced nodes themselves.
        var slot = replaced.Count > 0 ? replaced[0].PreviousNode : null;
        XElement? deletion = null;
        XElement? lastKept = null;
        foreach (var node in replaced)
        {
            if (node.Name == W.r)
            {
                if (deletion is null)
                {
                    deletion = CreateRevisionEnvelope(W.del, stamp);
                    node.AddBeforeSelf(deletion);
                }
                node.Remove();
                ConvertTextToDeletedText(node);
                deletion.Add(node);
            }
            else
            {
                WrapDescendantRunsInDel(node, stamp);
                if (node.Parent is not null) lastKept = node;
            }
        }
        if (payload.Count == 0) return;
        var insertion = CreateRevisionEnvelope(W.ins, stamp, payload);
        if (deletion is not null) deletion.AddAfterSelf(insertion);
        else if (lastKept is not null) lastKept.AddAfterSelf(insertion);
        else if (slot is not null) slot.AddAfterSelf(insertion);
        else if (replaced.Count > 0) container.AddFirst(insertion);
        else container.Add(insertion);
    }

    private static bool IsPreservedPayloadNode(XElement node, bool preserveNested) =>
        preserveNested && (node.Name == W.sdt || (node.Name != W.p && node.Descendants(W.sdt).Any()));

    private static bool IsRangeMarker(XElement node) =>
        node.Name == W.bookmarkStart || node.Name == W.bookmarkEnd
        || node.Name == W.commentRangeStart || node.Name == W.commentRangeEnd;

    /// <summary>The nodes a fill discards (or, tracked, deletes): every direct payload child
    /// that is not preserved, and the own inline content of a paragraph that hosts a preserved
    /// nested control.</summary>
    private static IReadOnlyList<XElement> ReplacedPayloadNodes(XElement content, bool preserveNested)
    {
        var replaced = new List<XElement>();
        foreach (var child in content.Elements())
        {
            if (IsPreservedPayloadNode(child, preserveNested) || IsRangeMarker(child)) continue;
            if (child.Name == W.p && preserveNested && child.Descendants(W.sdt).Any())
            {
                replaced.AddRange(child.Elements().Where(node => node.Name != W.pPr
                    && !IsRangeMarker(node) && node.Name != W.sdt && !node.Descendants(W.sdt).Any()));
                continue;
            }
            replaced.Add(child);
        }
        return replaced;
    }

    /// <summary>Delete every run under <paramref name="container"/> the way Word does, in place,
    /// keeping hyperlink and field containers: an ordinary run becomes <c>w:del</c>; a run already
    /// deleted or moved away stays as it is; a run inside the session author's own insertion is
    /// simply un-inserted (retyping your own tracked text leaves no trace of the first attempt);
    /// a run inside another author's insertion is deleted inside that insertion.</summary>
    private void WrapDescendantRunsInDel(XElement container, RevisionStamp stamp)
    {
        foreach (var run in container.Descendants(W.r).ToList())
        {
            if (run.Ancestors().Any(ancestor => ancestor.Name == W.del || ancestor.Name == W.moveFrom))
                continue;
            var insertion = run.Ancestors().FirstOrDefault(ancestor =>
                ancestor.Name == W.ins || ancestor.Name == W.moveTo);
            if (insertion is not null
                && string.Equals((string?)insertion.Attribute(W.author), stamp.Author, StringComparison.Ordinal))
            {
                run.Remove();
                if (!insertion.HasElements) insertion.Remove();
                continue;
            }
            var envelope = CreateRevisionEnvelope(W.del, stamp);
            run.ReplaceWith(envelope);
            envelope.Add(run);
            ConvertTextToDeletedText(run);
        }
    }

    /// <summary>
    /// The insertion mirror of <c>MarkStructuredBlockAsTrackedDeleted</c>: two paired custom-XML
    /// insertion ranges cross the wrapper's tags and every payload paragraph (and nested block
    /// wrapper) is inserted content, so reject removes the wrapper and payload together and
    /// accept keeps them with the markers stripped.
    /// </summary>
    private void MarkStructuredBlockAsTrackedInserted(XElement wrapper, RevisionStamp stamp)
    {
        var content = wrapper.Element(W.sdtContent)
            ?? throw new InvalidOperationException("block w:sdt has no w:sdtContent");
        foreach (var child in content.Elements().ToList())
        {
            if (child.Name == W.p) MarkParagraphContentAndMark(child, W.ins, stamp.Author, stamp.Date);
            else if (child.Name == W.sdt) MarkStructuredBlockAsTrackedInserted(child, stamp);
            else throw new InvalidOperationException(
                $"tracked insertion of a repeating item cannot represent {child.Name.LocalName}");
        }
        var boundaries = Internal.StructuredRevisionOps.AddCrossBoundaryMarkers(
            content, W.customXmlInsRangeStart, W.customXmlInsRangeEnd,
            name => CreateRevisionEnvelope(name, stamp));
        wrapper.AddBeforeSelf(boundaries.Before);
        wrapper.AddAfterSelf(boundaries.After);
    }

    private bool ResolveContentControlForMutation(string anchorId,
        ContentControlOperation operation, ContentControlFillOptions? options,
        out ContentControlCandidate? candidate, out EditResult? error)
    {
        candidate = null;
        error = null;
        if (_disposed)
        {
            error = EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
            return false;
        }
        var registry = BuildContentControlRegistry(ProjectionScopes.All);
        candidate = registry.FirstOrDefault(value =>
            string.Equals(value.Info.AnchorId, anchorId, StringComparison.Ordinal));
        if (candidate is null)
        {
            error = EditResult.Fail(EditErrorCode.ContentControlNotFound,
                $"content control not found: {anchorId}", anchorId);
            return false;
        }
        IReadOnlyList<ImageCandidate>? imageCandidates = null;
        error = GateContentControlOperation(candidate, operation,
            options ?? new ContentControlFillOptions(), registry, ref imageCandidates);
        return error is null;
    }

    /// <summary>
    /// The one gate for every content-control operation, applied in the order a mutation would
    /// report it. Discovery evaluates it per operation (and per nested policy) to build
    /// <see cref="ContentControlInfo.Operations"/>, and every mutation evaluates it once before
    /// taking an undo snapshot, so the registry can never advertise a mutation the session is
    /// guaranteed to refuse.
    /// </summary>
    private EditResult? GateContentControlOperation(
        ContentControlCandidate candidate,
        ContentControlOperation operation,
        ContentControlFillOptions options,
        IReadOnlyList<ContentControlCandidate> registry,
        ref IReadOnlyList<ImageCandidate>? imageCandidates)
    {
        var anchorId = candidate.Info.AnchorId;
        var type = candidate.Info.Type;
        if (candidate.MalformedReason is not null || candidate.MalformedAncestorReason is not null
            || !candidate.Identity.HasMutableIdentity)
            return EditResult.Fail(EditErrorCode.ContentControlMalformed,
                candidate.MalformedReason ?? candidate.MalformedAncestorReason
                    ?? (candidate.Identity.HasMutableIdentity ? null : candidate.Info.UnsupportedReason)
                    ?? "content control has no unique valid native w:id", anchorId);
        if (type == ContentControlType.Unsupported)
            return EditResult.Fail(EditErrorCode.ContentControlUnsupported,
                "unsupported content-control family", anchorId);
        var expected = ExpectedTypes(operation);
        if (!expected.Contains(type))
            return EditResult.Fail(EditErrorCode.ContentControlWrongType,
                $"operation requires {string.Join(" or ", expected)} but target is {type}", anchorId);
        if (candidate.Info.Placement == ContentControlPlacement.Unknown)
            return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                "unsupported or malformed OOXML placement", anchorId);
        if (!IsMutationPlacementSupported(type, candidate.Info.Placement))
            return EditResult.Fail(EditErrorCode.ContentControlPlacementUnsupported,
                $"{type} mutation supports only inline and block content controls", anchorId);
        if (ValidateEffectiveLocks(candidate,
                removingWrapper: operation == ContentControlOperation.RemoveRepeatingItem) is { } lockError)
            return lockError;
        if (ValidateBindingPolicy(candidate, options) is { } bindingError)
            return bindingError;
        if (NestedControlsGate(candidate, operation, options, registry) is { } nestedError)
            return nestedError;
        if (TrackedOperationBlocker(candidate.Element, type, operation, options, ref imageCandidates)
            is { } trackedReason)
            return EditResult.Fail(EditErrorCode.TrackedOperationUnsupported, trackedReason, anchorId);
        if (operation is ContentControlOperation.AddRepeatingItem or ContentControlOperation.RemoveRepeatingItem
            && RepeatingMutationConstraint(candidate.Element, type) is { } repeatingReason)
            return EditResult.Fail(EditErrorCode.RepeatingSectionConstraint, repeatingReason, anchorId);
        if (operation == ContentControlOperation.FillPicture)
        {
            imageCandidates ??= EnumerateImageCandidates(ProjectionScopes.All);
            var target = ResolvePictureContentControlTarget(candidate.Element, imageCandidates);
            if (target.ErrorCode is { } code)
                return EditResult.Fail(code, target.Diagnostic!, anchorId);
        }
        if (_trackedChanges != TrackedChangeMode.RenderInline
            && BookmarkRemovalRoots(candidate, operation, options) is { } roots
            && ValidateBookmarkRemoval(roots, anchorId) is { } bookmarkError)
            return bookmarkError;
        return null;
    }

    /// <summary>
    /// Nested-control semantics of a whole-control fill (issue #763). Only text and rich-text
    /// fills can preserve or replace nested controls; every other family's payload is atomic.
    /// Replacing refuses when a nested control is locked or data-bound, so no protected control
    /// is ever discarded by implication. Preserving may fill named textual children, each of
    /// which must pass its own gates; a key that is not a nested control of the target fails.
    /// </summary>
    private EditResult? NestedControlsGate(ContentControlCandidate candidate,
        ContentControlOperation operation, ContentControlFillOptions options,
        IReadOnlyList<ContentControlCandidate> registry)
    {
        var anchorId = candidate.Info.AnchorId;
        var childFills = options.ChildFills ?? new Dictionary<string, string>();
        if (childFills.Count > 0 && options.NestedControls != ContentControlNestedPolicy.Preserve)
            return EditResult.Fail(EditErrorCode.ContentControlNestedFillUnsupported,
                "childFills requires nestedControls=preserve", anchorId);
        if (!IsFillOperation(operation)) return null;
        var nested = NestedContentControls(candidate.Element);
        if (nested.Count == 0)
            return childFills.Count == 0
                ? null
                : EditResult.Fail(EditErrorCode.ContentControlNotFound,
                    "childFills names controls, but the target contains no nested controls", anchorId);
        if (!IsTextualFill(operation) || options.NestedControls == ContentControlNestedPolicy.Refuse)
            return NestedFillError(anchorId, IsTextualFill(operation));
        var byElement = registry.ToDictionary(value => value.Element, value => value);
        if (options.NestedControls == ContentControlNestedPolicy.Replace)
        {
            foreach (var control in nested)
            {
                var nestedAnchor = byElement.TryGetValue(control, out var nestedCandidate)
                    ? nestedCandidate.Info.AnchorId : "nested control";
                var props = control.Element(W.sdtPr);
                if ((string?)props?.Element(ContentControlW + "lock")?.Attribute(W.val)
                        is "sdtLocked" or "contentLocked" or "sdtContentLocked")
                    return EditResult.Fail(EditErrorCode.ContentControlLocked,
                        $"nested control {nestedAnchor} is locked; replacing the payload would remove it", anchorId);
                if (FindDataBinding(props) is not null)
                    return EditResult.Fail(EditErrorCode.ContentControlBound,
                        $"nested control {nestedAnchor} is data-bound; replacing the payload would remove it", anchorId);
            }
            return null;
        }
        foreach (var childAnchor in childFills.Keys)
        {
            var child = registry.FirstOrDefault(value =>
                string.Equals(value.Info.AnchorId, childAnchor, StringComparison.Ordinal));
            if (child is null || !child.Element.Ancestors().Any(ancestor => ReferenceEquals(ancestor, candidate.Element)))
                return EditResult.Fail(EditErrorCode.ContentControlNotFound,
                    $"childFills target {childAnchor} is not a control nested in {anchorId}", childAnchor);
            if (child.Info.Type is not (ContentControlType.PlainText or ContentControlType.RichText))
                return EditResult.Fail(EditErrorCode.ContentControlWrongType,
                    $"childFills target {childAnchor} is {child.Info.Type}; only text and rich-text children can be filled", childAnchor);
            IReadOnlyList<ImageCandidate>? images = null;
            if (GateContentControlOperation(child, ContentControlOperation.FillText,
                    new ContentControlFillOptions(), registry, ref images) is { } childError)
                return childError;
        }
        return null;
    }

    /// <summary>
    /// The operations that have a faithful native tracked representation under
    /// <c>render_inline</c>, and why the others do not. Text and rich-text fills are run-level
    /// deletions and insertions inside the wrapper; a picture fill is a deleted run and an
    /// inserted run; repeating items use the paired custom-XML range envelopes; a checkbox
    /// state, date value or list selection lives in <c>w:sdtPr</c>, which no revision covers.
    /// </summary>
    private string? TrackedOperationBlocker(XElement element, ContentControlType type,
        ContentControlOperation operation, ContentControlFillOptions options,
        ref IReadOnlyList<ImageCandidate>? imageCandidates)
    {
        if (_trackedChanges != TrackedChangeMode.RenderInline) return null;
        switch (operation)
        {
            case ContentControlOperation.FillText:
            case ContentControlOperation.FillRichText:
                if (options.NestedControls != ContentControlNestedPolicy.Replace) return null;
                var content = element.Element(W.sdtContent);
                return NestedContentControls(element).Any(nested =>
                        !ReferenceEquals(nested.Parent, content)
                        || DetectContentControlPlacement(nested) != ContentControlPlacement.Block)
                    ? "a tracked fill cannot remove a nested control that is not a block-level child of the target; keep it with nestedControls=preserve or switch modes"
                    : null;
            case ContentControlOperation.SetChecked:
                return "a checkbox state (w14:checked) has no tracked representation; switch modes to change it";
            case ContentControlOperation.SetDate:
                return "a date value (w:date) has no tracked representation; switch modes to change it";
            case ContentControlOperation.SelectItem:
                return "a list selection (w:lastValue) has no tracked representation; switch modes to change it";
            case ContentControlOperation.FillPicture:
                imageCandidates ??= EnumerateImageCandidates(ProjectionScopes.All);
                var target = ResolvePictureContentControlTarget(element, imageCandidates);
                return target.Image is not null && target.Image.Outer.Parent?.Name != W.r
                    ? "a tracked picture fill needs the picture in a plain run"
                    : null;
            case ContentControlOperation.AddRepeatingItem:
                var items = element.Element(W.sdtContent)?.Elements(W.sdt).Where(IsRepeatingSectionItem).ToList();
                return items is { Count: > 0 } ? TrackedRepeatingInsertBlocker(DefaultRepeatingTemplate(items)) : null;
            default:
                return null;
        }
    }

    /// <summary>The item a default add clones: the last one that can be cloned safely. An item
    /// that is itself a tracked insertion carries revision ids and range markers no clone may
    /// repeat, so the pristine item before it stays the template until the revision resolves.</summary>
    private static XElement DefaultRepeatingTemplate(IReadOnlyList<XElement> items) =>
        items.LastOrDefault(item => FindUnsafeRepeatingCloneCarrier(item) is null) ?? items[^1];

    private static string? TrackedRepeatingInsertBlocker(XElement template) =>
        template.Descendants(W.tbl).Any()
            || template.Descendants(W.p).Any(paragraph => paragraph.Descendants(W.sdt).Any())
            ? "tracked insertion of a repeating item containing a table or an inline nested control is unsupported; switch modes to add it"
            : null;

    /// <summary>The payload nodes an untracked operation discards, for the bookmark gate; null
    /// when the operation discards nothing.</summary>
    private static IReadOnlyList<XElement>? BookmarkRemovalRoots(ContentControlCandidate candidate,
        ContentControlOperation operation, ContentControlFillOptions options)
    {
        if (operation == ContentControlOperation.RemoveRepeatingItem)
            return new[] { candidate.Element };
        if (!IsFillOperation(operation) || operation == ContentControlOperation.FillPicture)
            return null;
        var content = candidate.Element.Element(W.sdtContent);
        if (content is null) return null;
        return options.NestedControls == ContentControlNestedPolicy.Preserve
            && NestedContentControls(candidate.Element).Count > 0
            ? ReplacedPayloadNodes(content, preserveNested: true)
            : new[] { content };
    }

    private static IReadOnlyList<XElement> NestedContentControls(XElement control) =>
        control.Element(W.sdtContent)?.Descendants(W.sdt).ToList() ?? new List<XElement>();

    private static bool IsTextualFill(ContentControlOperation operation) =>
        operation is ContentControlOperation.FillText or ContentControlOperation.FillRichText;

    private static bool IsFillOperation(ContentControlOperation operation) =>
        operation is ContentControlOperation.FillText or ContentControlOperation.FillRichText
            or ContentControlOperation.SetChecked or ContentControlOperation.SetDate
            or ContentControlOperation.SelectItem or ContentControlOperation.FillPicture;

    private static IReadOnlyList<ContentControlType> ExpectedTypes(ContentControlOperation operation) =>
        operation switch
        {
            ContentControlOperation.FillText => new[] { ContentControlType.PlainText, ContentControlType.RichText },
            ContentControlOperation.FillRichText => new[] { ContentControlType.RichText },
            ContentControlOperation.SetChecked => new[] { ContentControlType.Checkbox },
            ContentControlOperation.SetDate => new[] { ContentControlType.Date },
            ContentControlOperation.SelectItem => new[] { ContentControlType.DropDownList, ContentControlType.ComboBox },
            ContentControlOperation.FillPicture => new[] { ContentControlType.Picture },
            ContentControlOperation.AddRepeatingItem => new[] { ContentControlType.RepeatingSection },
            ContentControlOperation.RemoveRepeatingItem => new[] { ContentControlType.RepeatingSectionItem },
            _ => Array.Empty<ContentControlType>(),
        };

    private static IReadOnlyList<ContentControlOperation> OperationsForFamily(ContentControlType type) =>
        type switch
        {
            ContentControlType.PlainText => new[] { ContentControlOperation.FillText },
            ContentControlType.RichText => new[] { ContentControlOperation.FillText, ContentControlOperation.FillRichText },
            ContentControlType.Checkbox => new[] { ContentControlOperation.SetChecked },
            ContentControlType.Date => new[] { ContentControlOperation.SetDate },
            ContentControlType.DropDownList or ContentControlType.ComboBox => new[] { ContentControlOperation.SelectItem },
            ContentControlType.Picture => new[] { ContentControlOperation.FillPicture },
            ContentControlType.RepeatingSection => new[] { ContentControlOperation.AddRepeatingItem },
            ContentControlType.RepeatingSectionItem => new[] { ContentControlOperation.RemoveRepeatingItem },
            _ => Array.Empty<ContentControlOperation>(),
        };

    internal static string OperationName(ContentControlOperation operation) => operation switch
    {
        ContentControlOperation.FillText => "fill_text",
        ContentControlOperation.FillRichText => "fill_rich_text",
        ContentControlOperation.SetChecked => "set_checked",
        ContentControlOperation.SetDate => "set_date",
        ContentControlOperation.SelectItem => "select_item",
        ContentControlOperation.FillPicture => "fill_picture",
        ContentControlOperation.AddRepeatingItem => "add_repeating_item",
        ContentControlOperation.RemoveRepeatingItem => "remove_repeating_item",
        _ => throw new ArgumentOutOfRangeException(nameof(operation)),
    };

    internal static string NestedPolicyName(ContentControlNestedPolicy policy) => policy switch
    {
        ContentControlNestedPolicy.Refuse => "refuse",
        ContentControlNestedPolicy.Preserve => "preserve",
        ContentControlNestedPolicy.Replace => "replace",
        _ => throw new ArgumentOutOfRangeException(nameof(policy)),
    };

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
        ContentControlFillOptions? options, Action mutation,
        IReadOnlyList<Anchor>? removed = null, IReadOnlyList<Anchor>? alsoModified = null)
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
            var modified = new List<Anchor> { AnchorFromCandidate(candidate) };
            if (alsoModified is not null) modified.AddRange(alsoModified);
            return new EditResult { Success = true,
                Removed = removed ?? Array.Empty<Anchor>(),
                Modified = modified };
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

    private static EditResult NestedFillError(string anchorId, bool textual) =>
        EditResult.Fail(EditErrorCode.ContentControlNestedFillUnsupported,
            textual
                ? "whole-control fill is refused when the target contains nested controls; pass nestedControls=preserve or replace, or address the child control directly"
                : "whole-control fill is refused when the target contains nested controls; address the child control directly",
            anchorId);

    private IReadOnlyList<ContentControlCandidate> BuildContentControlRegistry(ProjectionScopes scopes)
    {
        var result = new List<ContentControlCandidate>();
        IReadOnlyList<ImageCandidate>? imageCandidates = null;
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
                var defaultOperation = OperationsForFamily(type).FirstOrDefault();
                if (OperationsForFamily(type).Count > 0
                    && TrackedOperationBlocker(element, type, defaultOperation, new ContentControlFillOptions(),
                        ref imageCandidates) is { } trackedReason)
                    unsupported = trackedReason;
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

        // Per-operation support is the same gate every mutation applies, evaluated for each
        // operation of the family — and, when the target contains nested controls, for each
        // nested policy — so an agent planning off the registry sees exactly what will apply.
        var byElementAll = result.ToDictionary(candidate => candidate.Element, candidate => candidate);
        for (int i = 0; i < result.Count; i++)
        {
            var candidate = result[i];
            var nestedAnchors = NestedContentControls(candidate.Element)
                .Select(control => byElementAll.TryGetValue(control, out var nested) ? nested.Info.AnchorId : null)
                .Where(anchor => anchor is not null)
                .Select(anchor => anchor!)
                .ToList();
            var operations = new List<ContentControlOperationSupport>();
            foreach (var operation in OperationsForFamily(candidate.Info.Type))
            {
                if (IsTextualFill(operation) && nestedAnchors.Count > 0)
                {
                    foreach (var policy in new[]
                             {
                                 ContentControlNestedPolicy.Refuse,
                                 ContentControlNestedPolicy.Preserve,
                                 ContentControlNestedPolicy.Replace,
                             })
                    {
                        var error = GateContentControlOperation(candidate, operation,
                            new ContentControlFillOptions { NestedControls = policy }, result, ref imageCandidates);
                        operations.Add(new ContentControlOperationSupport(
                            OperationName(operation), NestedPolicyName(policy), error is null, error?.Error?.Message));
                    }
                }
                else
                {
                    var error = GateContentControlOperation(candidate, operation,
                        new ContentControlFillOptions(), result, ref imageCandidates);
                    operations.Add(new ContentControlOperationSupport(
                        OperationName(operation), null, error is null, error?.Error?.Message));
                }
            }
            result[i] = candidate with
            {
                Info = candidate.Info with
                {
                    NestedControlAnchorIds = nestedAnchors,
                    Operations = operations,
                },
            };
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
            if (items.Count == 0 || content!.Elements().Any(element => !IsRevisionRangeMarker(element)
                    && (element.Name != W.sdt || !IsRepeatingSectionItem(element))))
                return "repeating section must contain only one or more direct repeating-section-item controls";
            if (FindUnsafeRepeatingCloneCarrier(DefaultRepeatingTemplate(items)) is { } unsafeCarrier)
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
        // Revision carriers are transparent to placement: a custom-XML revision range marker says
        // nothing about the grammar around it, and an inline w:ins/w:del stands for the runs it
        // wraps — exactly what a tracked fill or a tracked repeating-item edit leaves behind, so
        // a control must keep reading as inline or block through its own revisions.
        var children = content.Elements().Where(element => !IsRevisionRangeMarker(element)).ToList();
        if (children.Count == 0)
            return DetectContentControlPlacementFromContext(control);
        // A nested SDT is valid in every placement grammar, so an sdt-only payload is
        // intrinsically ambiguous from children alone. Its parent context is authoritative.
        if (children.All(element => element.Name == W.sdt))
            return DetectContentControlPlacementFromContext(control);
        bool allInline = children.All(IsInlineSdtContent);
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

    private static bool IsInlineSdtContent(XElement element) =>
        element.Name == W.r || element.Name == W.hyperlink
        || element.Name == W.fldSimple || element.Name == W.sdt || element.Name == W.smartTag
        || element.Name == W.bookmarkStart || element.Name == W.bookmarkEnd
        || element.Name == W.commentRangeStart || element.Name == W.commentRangeEnd
        || ((element.Name == W.ins || element.Name == W.del
                || element.Name == W.moveFrom || element.Name == W.moveTo)
            && element.Elements().All(IsInlineSdtContent));

    private static bool IsRevisionRangeMarker(XElement element) =>
        element.Name == W.customXmlInsRangeStart || element.Name == W.customXmlInsRangeEnd
        || element.Name == W.customXmlDelRangeStart || element.Name == W.customXmlDelRangeEnd
        || element.Name == W.customXmlMoveFromRangeStart || element.Name == W.customXmlMoveFromRangeEnd
        || element.Name == W.customXmlMoveToRangeStart || element.Name == W.customXmlMoveToRangeEnd;

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
