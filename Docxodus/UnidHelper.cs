#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Shared helpers for the <c>PtOpenXml.Unid</c> stable-id attribute. The Unid is a 32-char
/// hex string derived from a Guid; once assigned to an element it remains stable across
/// reformatting and reordering, which is the foundation for both <see cref="WmlComparer"/>'s
/// element matching and the <see cref="WmlToMarkdownConverter"/>'s anchor scheme.
/// </summary>
internal static class UnidHelper
{
    internal static string GenerateUnid() => Guid.NewGuid().ToString().Replace("-", "");

    /// <summary>
    /// Assigns a <c>PtOpenXml.Unid</c> attribute to <paramref name="contentParent"/> (if it is a
    /// footnote/endnote root) and to every descendant that does not already have one.
    /// </summary>
    internal static void AssignToAllElements(XElement contentParent)
    {
        if (contentParent.Name == W.footnote || contentParent.Name == W.endnote)
        {
            if (contentParent.Attribute(PtOpenXml.Unid) == null)
            {
                contentParent.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
            }
        }

        foreach (var d in contentParent.Descendants())
        {
            if (d.Attribute(PtOpenXml.Unid) == null)
            {
                d.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
            }
        }
    }
}
