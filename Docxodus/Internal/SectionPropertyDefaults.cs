// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// The section-property values whose omission the specification defines (ECMA-376 Part 1, §17.6), used in
/// both directions. A section-property revision writes them out on whichever side leaves them out, because a
/// desktop word processor reads a value one side of a <c>w:sectPrChange</c> omits as unchanged rather than as
/// its default (issue #1775). And every comparison of section properties drops them first, so that a stated
/// default and an omitted one compare equal, and a resolved redline compares equal to its input.
/// </summary>
internal static class SectionPropertyDefaults
{
    /// <summary>The <c>CT_OnOff</c> section properties that an omission turns off. <c>w:formProt</c> is left
    /// out: whether an omitted one protects its section under forms protection is not settled here.</summary>
    private static readonly XName[] Toggles = { W.noEndnote, W.titlePg, W.bidi, W.rtlGutter };

    /// <summary>Attributes whose omission reads as a value.</summary>
    private static readonly Dictionary<(XName Element, XName Attribute), string> AttributeDefaults = new()
    {
        [(W.type, W.val)] = "nextPage",
        [(W.pgSz, W.orient)] = "portrait",
        [(W.pgNumType, W.fmt)] = "decimal",
        [(W.cols, W.num)] = "1",
        [(W.cols, W.sep)] = "0",
        [(W.vAlign, W.val)] = "top",
        [(W.textDirection, W.val)] = "lrTb",
        [(W.docGrid, W.type)] = "default",
    };

    /// <summary>Elements whose omission reads as the element holding only the attribute defaults above. An
    /// omitted page size, margin set, line numbering or page border has no such reading.</summary>
    private static readonly XName[] ElementsWithDefaults = { W.type, W.pgNumType, W.cols, W.vAlign, W.textDirection, W.docGrid };

    /// <summary>
    /// Make <paramref name="target"/>, one side of a section-property revision, state every value that
    /// <paramref name="other"/> states where leaving it out means a value defined above: a toggle
    /// <paramref name="other"/> has on is written <c>w:val="0"</c>, and an attribute (or the element holding
    /// it) left at its default is written out. Values with no specified default are left as written.
    /// </summary>
    public static void StateOmitted(XElement target, XElement other)
    {
        foreach (var source in other.Elements().ToList())
        {
            var existing = target.Element(source.Name);
            if (Toggles.Contains(source.Name))
            {
                if (existing is null && OnOff(source) == true)
                    WordprocessingMLUtil.InsertSectPrChildInOrder(target, new XElement(source.Name, new XAttribute(W.val, "0")));
                continue;
            }
            if (existing is null && !ElementsWithDefaults.Contains(source.Name))
                continue;
            var stated = existing ?? new XElement(source.Name);
            foreach (var attribute in source.Attributes())
            {
                // Explicit columns set their own count; a stated count of one beside them would contradict them.
                if (source.Name == W.cols && attribute.Name == W.num &&
                    (source.Elements(W.col).Any() || stated.Elements(W.col).Any()))
                    continue;
                if (stated.Attribute(attribute.Name) is null &&
                    AttributeDefaults.TryGetValue((source.Name, attribute.Name), out var value))
                    stated.SetAttributeValue(attribute.Name, value);
            }
            if (existing is null && stated.HasAttributes)
                WordprocessingMLUtil.InsertSectPrChildInOrder(target, stated);
        }
    }

    /// <summary>
    /// A copy of <paramref name="sectPr"/> without the values that only restate a default: a toggle stated
    /// off, an attribute equal to its default, and an element with a default left holding nothing. Compare
    /// section properties through this, so that a stated default and an omitted one are the same property.
    /// </summary>
    public static XElement WithoutStatedDefaults(XElement sectPr)
    {
        var copy = new XElement(sectPr);
        foreach (var child in copy.Elements().ToList())
        {
            if (Toggles.Contains(child.Name))
            {
                // Every spelling of off is the omitted toggle and every spelling of on the bare element; a value
                // that is neither stays as written, so it still differs from both.
                var state = OnOff(child);
                if (state == false)
                    child.Remove();
                else if (state == true)
                    child.Attribute(W.val)?.Remove();
                continue;
            }
            foreach (var attribute in child.Attributes().ToList())
                if (AttributeDefaults.TryGetValue((child.Name, attribute.Name), out var value) &&
                    (attribute.Value == value || OnOffAttributes.Contains((child.Name, attribute.Name)) && OnOff(attribute.Value) == false))
                    attribute.Remove();
            // Bookkeeping attributes (the reader's unids) say nothing about the property.
            if (ElementsWithDefaults.Contains(child.Name) && !child.HasElements &&
                child.Attributes().All(attribute => attribute.Name.Namespace == PtOpenXml.pt))
                child.Remove();
        }
        return copy;
    }

    /// <summary>The attribute defaults above that are on/off values, whose off reads in any spelling.</summary>
    private static readonly (XName Element, XName Attribute)[] OnOffAttributes = { (W.cols, W.sep) };

    /// <summary>A toggle's state: on when it has no <c>w:val</c>.</summary>
    private static bool? OnOff(XElement toggle) =>
        toggle.Attribute(W.val) is { } value ? OnOff(value.Value) : true;

    /// <summary>An on/off value, read without regard to case as Docxodus's readers read it; null when it is
    /// neither, so an unreadable value is never taken for a default.</summary>
    private static bool? OnOff(string value) => WordprocessingMLUtil.OnOffValue(value);
}
