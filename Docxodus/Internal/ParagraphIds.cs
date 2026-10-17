// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Makes every <c>w14:paraId</c> unique in a package whose story parts hold content from more than one
/// source document (issue #1170).
/// <para>A paraId must be unique across the whole package. A comparison keeps a paragraph it replaces whole
/// as a deleted copy of the original beside an inserted copy of the revision, and both carry the source's
/// paraId; a comment the revision added brings its definition's paraId from the other package, where it
/// may name a paragraph of the original. <see cref="MakeUnique(WordprocessingDocument)"/> renames each
/// later duplicate.</para>
/// </summary>
internal static class ParagraphIds
{
    /// <summary>
    /// Give every paragraph in the story parts of <paramref name="document"/> (comments, body, headers,
    /// footers, footnotes, endnotes and the glossary document) its own <c>w14:paraId</c>. Comment
    /// definitions are read first, and the keys of the <c>commentsExtended</c>/<c>commentsIds</c> threading
    /// entries next, so a comment keeps the paraId its threading is keyed on and a story paragraph never
    /// takes over a threading key; nothing refers to a story paragraph's paraId, so a later duplicate is
    /// renamed to a fresh value from <see cref="CommentOps.ParaIdAllocator"/>. Writes back each part it
    /// changed; one scan and no writes when every id is already unique.
    /// </summary>
    internal static void MakeUnique(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart;
        if (main is null)
            return;

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<(OpenXmlPart Part, XAttribute Id)>? duplicates = null;
        void Scan(OpenXmlPart? part)
        {
            var root = part?.GetXDocument().Root;
            if (root is null)
                return;
            foreach (var id in root.Descendants().Attributes(W14.paraId))
                if (!seen.Add(id.Value))
                    (duplicates ??= new()).Add((part!, id));
        }

        Scan(main.WordprocessingCommentsPart);
        foreach (var key in CommentOps.ThreadingParaIds(main))
            seen.Add(key);
        Scan(main);
        foreach (var part in main.HeaderParts) Scan(part);
        foreach (var part in main.FooterParts) Scan(part);
        Scan(main.FootnotesPart);
        Scan(main.EndnotesPart);
        Scan(main.GlossaryDocumentPart);
        if (duplicates is null)
            return;

        var allocator = new CommentOps.ParaIdAllocator(main);
        foreach (var (_, id) in duplicates)
            id.Value = allocator.Next();
        foreach (var part in duplicates.Select(d => d.Part).Distinct())
            part.PutXDocument();
    }
}
