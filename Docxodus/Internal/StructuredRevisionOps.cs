#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Builds the paired cross-boundary range topology Word uses to track the existence of
/// structured OOXML wrappers such as <c>w:sdt</c> and <c>w:customXml</c>.
/// </summary>
internal static class StructuredRevisionOps
{
    /// <summary>
    /// Adds the two inner markers for a wrapper whose opening and closing tags are tracked,
    /// and returns the two outer markers for the caller to place around the wrapper.
    /// </summary>
    /// <remarks>
    /// Range A starts before the wrapper and ends at the start of its content. Range B starts
    /// at the end of the content and ends after the wrapper. The range ids must be distinct;
    /// <see cref="RevisionProcessor"/> recognizes the wrapper by intersecting the elements
    /// crossed by both ranges.
    /// </remarks>
    internal static (XElement Before, XElement After) AddCrossBoundaryMarkers(
        XElement contentContainer,
        XName startName,
        XName endName,
        Func<XName, XElement> createRangeStart)
    {
        ArgumentNullException.ThrowIfNull(contentContainer);
        ArgumentNullException.ThrowIfNull(startName);
        ArgumentNullException.ThrowIfNull(endName);
        ArgumentNullException.ThrowIfNull(createRangeStart);

        var before = createRangeStart(startName);
        var beforeId = RequiredRangeId(before);
        var openingEnd = new XElement(endName, new XAttribute(W.id, beforeId));

        // A w:customXml wrapper is its own content container and keeps its properties as
        // its first child; the schema orders them ahead of any range markup.
        if (contentContainer.Elements().FirstOrDefault() is { } first
            && first.Name == W.customXmlPr)
            first.AddAfterSelf(openingEnd);
        else
            contentContainer.AddFirst(openingEnd);

        var closingStart = createRangeStart(startName);
        var afterId = RequiredRangeId(closingStart);
        contentContainer.Add(closingStart);
        var after = new XElement(endName, new XAttribute(W.id, afterId));
        return (before, after);
    }

    /// <summary>
    /// The revision side of a wrapper whose existence <see cref="AddCrossBoundaryMarkers"/> tracked, read back
    /// from the opening range's end marker it placed first in <paramref name="contentContainer"/> (after any
    /// <c>w:customXmlPr</c>): true for a deleted wrapper, false for an inserted one, and null when the wrapper is
    /// not tracked. Everything inside such a wrapper — including block-level markers that carry no
    /// <c>w:ins</c>/<c>w:del</c> of their own — exists in only that one revision view.
    /// </summary>
    internal static bool? TrackedWrapperIsDeleted(XElement contentContainer)
    {
        ArgumentNullException.ThrowIfNull(contentContainer);
        var first = contentContainer.Elements().FirstOrDefault(e => e.Name != W.customXmlPr);
        if (first == null)
            return null;
        if (first.Name == W.customXmlDelRangeEnd)
            return true;
        if (first.Name == W.customXmlInsRangeEnd)
            return false;
        return null;
    }

    private static string RequiredRangeId(XElement start) =>
        (string?)start.Attribute(W.id)
        ?? throw new InvalidOperationException("structured revision range start has no w:id");
}
