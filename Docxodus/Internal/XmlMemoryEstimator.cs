#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Approximate retained-heap cost of a LINQ-to-XML tree, used to bound the undo ring.
///
/// <para><b>Approximate on purpose.</b> The consumer is a memory BUDGET, which needs an
/// estimate that is cheap, monotonic in document size, and never wildly low. It does not need
/// accounting precision, and paying for precision would be self-defeating: the estimator runs
/// on every mutation, next to the deep clone it is measuring.</para>
///
/// <para>Serialized length is deliberately NOT the measure. An <see cref="XElement"/> costs far
/// more in memory than its angle brackets on disk — a document whose XML is 5 MB routinely
/// occupies 40-60 MB as a live tree, because every element, attribute and text node is a heap
/// object with headers, parent/sibling links and an interned <see cref="XName"/>. Budgeting
/// against serialized size would under-count by roughly an order of magnitude, which is the
/// error that matters here.</para>
///
/// <para>Constants are 64-bit CLR object sizes rounded up to the allocation granularity, and
/// erring high. They are a model, not a measurement: treat the output as "same order of
/// magnitude as the real retained set", which is all a budget requires.</para>
/// </summary>
internal static class XmlMemoryEstimator
{
    /// <summary>XElement: object header + XName ref + parent/next links + annotations slot + content ref.</summary>
    private const long ElementBytes = 80;

    /// <summary>XAttribute: object header + XName ref + value ref + parent/next links.</summary>
    private const long AttributeBytes = 64;

    /// <summary>XText / XComment / XProcessingInstruction: object header + value ref + links.</summary>
    private const long TextNodeBytes = 48;

    /// <summary>Any other XNode shape (XDocumentType, XCData is an XText subclass, …).</summary>
    private const long OtherNodeBytes = 48;

    /// <summary>String object header + length/hash fields, before the UTF-16 payload.</summary>
    private const long StringHeaderBytes = 26;

    /// <summary>
    /// Approximate retained bytes of <paramref name="doc"/>, including its element/attribute/text
    /// objects and their string payloads. Names are excluded: <see cref="XName"/> instances are
    /// interned per namespace+localname, so a document with 100,000 <c>w:p</c> elements holds ONE
    /// "w:p" name. Counting them per node would inflate the estimate by the document's size.
    /// </summary>
    /// <remarks>
    /// Walks the raw <c>FirstNode</c>/<c>NextNode</c> and <c>FirstAttribute</c>/<c>NextAttribute</c>
    /// links rather than <c>DescendantNodes()</c>/<c>Attributes()</c>. Those LINQ helpers allocate an
    /// iterator per element and run a state machine per node, which measured at ~1.7x the cost of the
    /// deep clone this runs beside — unacceptable next to a per-mutation snapshot. The manual walk is
    /// allocation-free and costs a small fraction of the clone. An explicit stack (not recursion)
    /// keeps a pathologically nested document from overflowing.
    /// </remarks>
    public static long Estimate(XDocument doc)
    {
        long total = 0;
        var stack = new Stack<XElement>();

        for (var node = doc.FirstNode; node is not null; node = node.NextNode)
        {
            if (node is XElement root) stack.Push(root);
            else total += NodeCost(node);
        }

        while (stack.Count > 0)
        {
            var element = stack.Pop();
            total += ElementBytes;

            for (var attribute = element.FirstAttribute; attribute is not null; attribute = attribute.NextAttribute)
                total += AttributeBytes + StringCost(attribute.Value);

            for (var node = element.FirstNode; node is not null; node = node.NextNode)
            {
                if (node is XElement child) stack.Push(child);
                else total += NodeCost(node);
            }
        }

        return total;
    }

    private static long NodeCost(XNode node) => node switch
    {
        XText text => TextNodeBytes + StringCost(text.Value),
        XComment comment => TextNodeBytes + StringCost(comment.Value),
        XProcessingInstruction pi => TextNodeBytes + StringCost(pi.Data),
        _ => OtherNodeBytes,
    };

    /// <summary>Header plus the UTF-16 payload (2 bytes per char).</summary>
    private static long StringCost(string? value) =>
        value is null ? 0 : StringHeaderBytes + (2L * value.Length);
}
