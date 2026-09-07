// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Resolves the auto-numbering prefix Word would render for a paragraph that
/// carries <c>w:numPr</c> — e.g. <c>"1."</c>, <c>"1.1"</c>, <c>"First"</c>.
/// Shared by <see cref="WmlToMarkdownConverter"/> (which embeds the prefix
/// inline in the markdown projection) and <see cref="DocxSession"/> (which
/// strips a matching prefix from <c>ReplaceText</c> payloads so an agent that
/// echoes back what it sees doesn't render the prefix twice).
/// </summary>
internal static class ListNumberResolver
{
    /// <summary>
    /// Returns the trimmed auto-number prefix for <paramref name="paragraph"/>, or
    /// <c>null</c> when the paragraph has no <c>w:numPr</c>, the retriever can't
    /// resolve it, or the resolved marker is a single non-alphanumeric glyph
    /// (bullets, which aren't meaningful as a heading prefix).
    /// </summary>
    public static string? Resolve(
        XElement paragraph,
        WordprocessingDocument doc,
        ListItemRetrieverSettings? settings = null)
    {
        // Don't short-circuit on inline `w:numPr`: numbering can also come from a
        // paragraph's style (the Heading1 style declaring its own numPr). The legal
        // NVCA Model COI does exactly that for its top-level "First Article"/"Second
        // Article" headings — they have <w:pStyle val="Heading1"/> only, no inline
        // numPr — and the HTML converter resolves them via ListItemRetriever just
        // fine. Letting ListItemRetriever decide (it returns null for non-list-items)
        // keeps the projector aligned with the HTML converter — see issue #141.
        try
        {
            var resolved = ListItemRetriever.RetrieveListItem(doc, paragraph, settings ?? new ListItemRetrieverSettings());
            if (string.IsNullOrWhiteSpace(resolved)) return null;
            var trimmed = resolved.TrimEnd();
            if (trimmed.Length == 1 && !char.IsLetterOrDigit(trimmed, 0)) return null;
            return trimmed;
        }
        catch
        {
            return null;
        }
    }
}
