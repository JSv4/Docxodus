#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus;

/// <summary>
/// Raw OOXML escape hatch for cases the markdown subset can't express
/// (complex tables, math, drawings, content controls, custom XML).
/// Accessed via <see cref="DocxSession.Raw"/>.
/// </summary>
public sealed class RawDocxOps
{
    private readonly DocxSession _session;

    internal RawDocxOps(DocxSession session) => _session = session;

    /// <summary>Returns the OOXML for the element the anchor names, useful as a template.</summary>
    public string GetXml(string anchorId) => _session.RawGetXmlInternal(anchorId);

    /// <summary>Inserts a block-level XML fragment before/after the anchor; new Unids auto-assigned.</summary>
    public EditResult InsertXml(string anchorId, Position pos, string xml) =>
        _session.RawInsertXmlInternal(anchorId, pos, xml);

    /// <summary>Replaces the anchored element with the supplied fragment.</summary>
    public EditResult ReplaceXml(string anchorId, string xml) =>
        _session.RawReplaceXmlInternal(anchorId, xml);
}
