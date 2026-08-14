#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Builds the paired cross-boundary range topology Word uses to track the existence of
/// structured OOXML wrappers such as <c>w:sdt</c>.
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

        contentContainer.AddFirst(openingEnd);

        var closingStart = createRangeStart(startName);
        var afterId = RequiredRangeId(closingStart);
        contentContainer.Add(closingStart);
        var after = new XElement(endName, new XAttribute(W.id, afterId));
        return (before, after);
    }

    private static string RequiredRangeId(XElement start) =>
        (string?)start.Attribute(W.id)
        ?? throw new InvalidOperationException("structured revision range start has no w:id");
}
