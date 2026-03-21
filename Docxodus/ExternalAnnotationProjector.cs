#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Projects external annotations onto HTML content.
/// Wraps annotated text ranges with styled spans.
/// </summary>
public static class ExternalAnnotationProjector
{
    private static readonly XNamespace Xhtml = "http://www.w3.org/1999/xhtml";

    /// <summary>
    /// Project external annotations onto an HTML document.
    /// This post-processes the HTML to wrap annotated text with styled spans.
    /// </summary>
    /// <param name="htmlDocument">The HTML document (as XElement).</param>
    /// <param name="annotationSet">The external annotation set to project.</param>
    /// <param name="settings">Projection settings.</param>
    /// <returns>Modified HTML with annotations projected.</returns>
    public static XElement ProjectAnnotations(
        XElement htmlDocument,
        ExternalAnnotationSet annotationSet,
        ExternalAnnotationProjectionSettings settings)
    {
        if (htmlDocument == null) throw new ArgumentNullException(nameof(htmlDocument));
        if (annotationSet == null) throw new ArgumentNullException(nameof(annotationSet));
        settings ??= new ExternalAnnotationProjectionSettings();

        // Clone the document to avoid modifying the original
        var result = new XElement(htmlDocument);

        // Sort annotations by start offset for correct nesting
        var sortedAnnotations = annotationSet.LabelledText
            .Where(a => !a.Structural && a.AnnotationJson is TextSpan)
            .Select(a => (Annotation: a, Span: (TextSpan)a.AnnotationJson!))
            .OrderBy(x => x.Span.Start)
            .ThenByDescending(x => x.Span.End) // Longer spans first for nesting
            .ToList();

        // Project each annotation using text search (not offsets).
        // We rebuild the text map each iteration because projecting an annotation
        // modifies the tree (adds wrapper + label spans), which shifts offsets.
        // GetTextNodes skips already-projected annotation wrappers so their label
        // text doesn't pollute the offset calculation.
        foreach (var (annotation, span) in sortedAnnotations)
        {
            var label = annotationSet.TextLabels.TryGetValue(annotation.AnnotationLabel, out var l)
                ? l : null;

            // Use the annotation's raw text to find it in the HTML
            var searchText = span.Text ?? annotation.RawText;
            if (string.IsNullOrEmpty(searchText)) continue;

            // Rebuild text map from current tree state (skipping already-projected spans)
            var textMap = BuildTextMap(result);
            var htmlText = GetHtmlText(textMap);

            // Find this text in the HTML
            var htmlLocation = FindTextInHtml(htmlText, searchText, new HashSet<int>());
            if (htmlLocation == null) continue;

            // Create a synthetic span with HTML-space offsets
            var htmlSpan = new TextSpan
            {
                Id = span.Id,
                Start = htmlLocation.Value.start,
                End = htmlLocation.Value.end,
                Text = searchText
            };

            ProjectSingleAnnotation(result, textMap, annotation, htmlSpan, label, settings);
        }

        // Add CSS for annotations
        AddAnnotationCss(result, annotationSet.TextLabels, settings);

        return result;
    }

    /// <summary>
    /// Convert a document to HTML with external annotations projected.
    /// This combines WmlToHtmlConverter with annotation projection.
    /// </summary>
    /// <param name="doc">The source document.</param>
    /// <param name="annotationSet">The external annotation set to project.</param>
    /// <param name="htmlSettings">HTML conversion settings.</param>
    /// <param name="projectionSettings">Annotation projection settings.</param>
    /// <returns>HTML string with annotations projected.</returns>
    public static string ConvertWithAnnotations(
        WmlDocument doc,
        ExternalAnnotationSet annotationSet,
        WmlToHtmlConverterSettings? htmlSettings = null,
        ExternalAnnotationProjectionSettings? projectionSettings = null)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (annotationSet == null) throw new ArgumentNullException(nameof(annotationSet));

        htmlSettings ??= new WmlToHtmlConverterSettings();
        projectionSettings ??= new ExternalAnnotationProjectionSettings();

        // Convert document to HTML
        var html = WmlToHtmlConverter.ConvertToHtml(doc, htmlSettings);

        // Project annotations
        var annotatedHtml = ProjectAnnotations(html, annotationSet, projectionSettings);

        return annotatedHtml.ToString();
    }

    #region Text Mapping

    private class TextMapEntry
    {
        public XText TextNode { get; set; } = null!;
        public int StartOffset { get; set; }
        public int EndOffset { get; set; }
    }

    private static List<TextMapEntry> BuildTextMap(XElement html)
    {
        var entries = new List<TextMapEntry>();
        var offset = 0;

        // Find body element
        var body = html.Descendants()
            .FirstOrDefault(e => e.Name.LocalName.Equals("body", StringComparison.OrdinalIgnoreCase));

        if (body == null)
        {
            body = html;
        }

        // Collect all text nodes with their offsets
        foreach (var textNode in GetTextNodes(body))
        {
            var text = textNode.Value;
            entries.Add(new TextMapEntry
            {
                TextNode = textNode,
                StartOffset = offset,
                EndOffset = offset + text.Length
            });
            offset += text.Length;
        }

        return entries;
    }

    /// <summary>
    /// Get the concatenated text from all text nodes in the HTML body.
    /// This represents the "HTML text" which may differ from document source text.
    /// </summary>
    private static string GetHtmlText(List<TextMapEntry> textMap)
    {
        var sb = new StringBuilder();
        foreach (var entry in textMap)
        {
            sb.Append(entry.TextNode.Value);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find the text in HTML by searching for the annotation's raw text.
    /// Returns (htmlStart, htmlEnd) offsets in HTML text space, or null if not found.
    /// </summary>
    private static (int start, int end)? FindTextInHtml(
        string htmlText,
        string searchText,
        HashSet<int> usedOffsets)
    {
        if (string.IsNullOrEmpty(searchText)) return null;

        // Find all occurrences and pick the first unused one
        var index = 0;
        while (index < htmlText.Length)
        {
            index = htmlText.IndexOf(searchText, index, StringComparison.Ordinal);
            if (index < 0) break;

            if (!usedOffsets.Contains(index))
            {
                usedOffsets.Add(index);
                return (index, index + searchText.Length);
            }
            index++;
        }

        return null;
    }

    private static IEnumerable<XText> GetTextNodes(XElement element)
    {
        foreach (var node in element.Nodes())
        {
            if (node is XText text)
            {
                yield return text;
            }
            else if (node is XElement child)
            {
                // Skip script and style elements
                var name = child.Name.LocalName.ToLowerInvariant();
                if (name == "script" || name == "style")
                    continue;

                // Skip already-projected annotation wrappers so their label
                // text doesn't shift offsets during subsequent projections
                if (child.Attribute("data-annotation-id") != null)
                    continue;

                foreach (var childText in GetTextNodes(child))
                {
                    yield return childText;
                }
            }
        }
    }

    #endregion

    #region Annotation Projection

    private static void ProjectSingleAnnotation(
        XElement html,
        List<TextMapEntry> textMap,
        OpenContractsAnnotation annotation,
        TextSpan span,
        AnnotationLabel? label,
        ExternalAnnotationProjectionSettings settings)
    {
        // Find text nodes that overlap with this annotation
        var overlappingEntries = textMap
            .Where(e => e.EndOffset > span.Start && e.StartOffset < span.End)
            .ToList();

        if (overlappingEntries.Count == 0) return;

        var isFirst = true;
        var isLast = false;

        for (int i = 0; i < overlappingEntries.Count; i++)
        {
            isLast = (i == overlappingEntries.Count - 1);
            var entry = overlappingEntries[i];

            WrapTextNode(entry, span, annotation, label, settings, isFirst, isLast);
            isFirst = false;
        }
    }

    private static void WrapTextNode(
        TextMapEntry entry,
        TextSpan span,
        OpenContractsAnnotation annotation,
        AnnotationLabel? label,
        ExternalAnnotationProjectionSettings settings,
        bool isFirst,
        bool isLast)
    {
        var textNode = entry.TextNode;
        var text = textNode.Value;
        var parent = textNode.Parent;
        if (parent == null) return;

        // Calculate the portion of text to wrap
        var textStart = Math.Max(0, span.Start - entry.StartOffset);
        var textEnd = Math.Min(text.Length, span.End - entry.StartOffset);

        if (textStart >= textEnd) return;

        // Split the text into before/annotated/after parts
        var beforeText = text.Substring(0, textStart);
        var annotatedText = text.Substring(textStart, textEnd - textStart);
        var afterText = text.Substring(textEnd);

        // Create the wrapper span
        var wrapper = CreateAnnotationWrapper(annotation, label, settings, isFirst, isLast);
        wrapper.Add(new XText(annotatedText));

        // Build the replacement nodes
        var replacements = new List<XNode>();
        if (beforeText.Length > 0)
        {
            replacements.Add(new XText(beforeText));
        }
        replacements.Add(wrapper);
        if (afterText.Length > 0)
        {
            replacements.Add(new XText(afterText));
        }

        // Replace the text node
        textNode.ReplaceWith(replacements.ToArray());
    }

    private static XElement CreateAnnotationWrapper(
        OpenContractsAnnotation annotation,
        AnnotationLabel? label,
        ExternalAnnotationProjectionSettings settings,
        bool isFirst,
        bool isLast)
    {
        var prefix = settings.CssClassPrefix;
        var cssClasses = new List<string> { $"{prefix}highlight" };

        if (isFirst && isLast)
        {
            cssClasses.Add($"{prefix}single");
        }
        else if (isFirst)
        {
            cssClasses.Add($"{prefix}start");
        }
        else if (isLast)
        {
            cssClasses.Add($"{prefix}end");
        }
        else
        {
            cssClasses.Add($"{prefix}continuation");
        }

        var wrapper = new XElement("span",
            new XAttribute("class", string.Join(" ", cssClasses)));

        // Add data attributes
        if (settings.IncludeMetadata)
        {
            wrapper.Add(new XAttribute("data-annotation-id", annotation.Id ?? ""));
            wrapper.Add(new XAttribute("data-label-id", annotation.AnnotationLabel ?? ""));

            if (label != null)
            {
                wrapper.Add(new XAttribute("data-label", label.Text));
                wrapper.Add(new XAttribute("style", $"--annot-color: {label.Color};"));
            }
        }

        // Add label element if this is the first segment
        if (isFirst && settings.LabelMode != AnnotationLabelMode.None && label != null)
        {
            var labelText = label.Text;
            if (!string.IsNullOrEmpty(labelText))
            {
                switch (settings.LabelMode)
                {
                    case AnnotationLabelMode.Above:
                    case AnnotationLabelMode.Inline:
                        var labelSpan = new XElement("span",
                            new XAttribute("class", $"{prefix}label"),
                            new XText(labelText));
                        wrapper.AddFirst(labelSpan);
                        break;
                    case AnnotationLabelMode.Tooltip:
                        wrapper.Add(new XAttribute("title", labelText));
                        break;
                }
            }
        }

        return wrapper;
    }

    #endregion

    #region Incremental Annotation API

    /// <summary>
    /// Project annotations onto an HTML string (already converted from DOCX).
    /// This avoids re-converting the DOCX when only annotations change.
    /// </summary>
    /// <param name="html">HTML string (previously converted via WmlToHtmlConverter).</param>
    /// <param name="annotationSet">The external annotation set to project.</param>
    /// <param name="settings">Projection settings.</param>
    /// <returns>HTML string with annotations projected.</returns>
    public static string ProjectAnnotationsOntoHtml(
        string html,
        ExternalAnnotationSet annotationSet,
        ExternalAnnotationProjectionSettings? settings = null)
    {
        if (string.IsNullOrEmpty(html)) throw new ArgumentNullException(nameof(html));
        if (annotationSet == null) throw new ArgumentNullException(nameof(annotationSet));
        settings ??= new ExternalAnnotationProjectionSettings();

        var htmlDoc = ParseHtmlString(html, out var wasWrapped);
        var result = ProjectAnnotations(htmlDoc, annotationSet, settings);
        return SerializeHtmlString(result, wasWrapped);
    }

    /// <summary>
    /// Add a single annotation to existing HTML without re-converting the document.
    /// The HTML should already be converted (with or without other annotations).
    /// </summary>
    /// <param name="html">HTML string.</param>
    /// <param name="annotation">The annotation to add.</param>
    /// <param name="label">Label definition for the annotation.</param>
    /// <param name="settings">Projection settings.</param>
    /// <returns>HTML string with the annotation added.</returns>
    public static string AddAnnotationToHtml(
        string html,
        OpenContractsAnnotation annotation,
        AnnotationLabel? label,
        ExternalAnnotationProjectionSettings? settings = null)
    {
        if (string.IsNullOrEmpty(html)) throw new ArgumentNullException(nameof(html));
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        settings ??= new ExternalAnnotationProjectionSettings();

        var htmlDoc = ParseHtmlString(html, out var wasWrapped);

        // Build text map and find annotation location
        var textMap = BuildTextMap(htmlDoc);
        var htmlText = GetHtmlText(textMap);
        var usedOffsets = new HashSet<int>();

        if (annotation.AnnotationJson is TextSpan span)
        {
            var searchText = span.Text ?? annotation.RawText;
            if (!string.IsNullOrEmpty(searchText))
            {
                var htmlLocation = FindTextInHtml(htmlText, searchText, usedOffsets);
                if (htmlLocation != null)
                {
                    var htmlSpan = new TextSpan
                    {
                        Id = span.Id,
                        Start = htmlLocation.Value.start,
                        End = htmlLocation.Value.end,
                        Text = searchText
                    };

                    textMap = BuildTextMap(htmlDoc);
                    ProjectSingleAnnotation(htmlDoc, textMap, annotation, htmlSpan, label, settings);
                }
            }
        }

        // Add per-annotation CSS (label color class)
        if (label != null)
        {
            AddSingleAnnotationCss(htmlDoc, annotation, label, settings);
        }

        return SerializeHtmlString(htmlDoc, wasWrapped);
    }

    /// <summary>
    /// Remove a single annotation from HTML by annotation ID.
    /// Unwraps annotation spans back to plain text.
    /// </summary>
    /// <param name="html">HTML string with annotations.</param>
    /// <param name="annotationId">ID of the annotation to remove.</param>
    /// <param name="cssClassPrefix">CSS class prefix used for annotations (default: "ext-annot-").</param>
    /// <returns>HTML string with the annotation removed.</returns>
    public static string RemoveAnnotationFromHtml(
        string html,
        string annotationId,
        string cssClassPrefix = "ext-annot-")
    {
        if (string.IsNullOrEmpty(html)) throw new ArgumentNullException(nameof(html));
        if (string.IsNullOrEmpty(annotationId)) throw new ArgumentNullException(nameof(annotationId));

        var htmlDoc = ParseHtmlString(html, out var wasWrapped);

        // Find all spans with data-annotation-id matching
        var annotationSpans = htmlDoc.Descendants("span")
            .Where(e => (string?)e.Attribute("data-annotation-id") == annotationId)
            .ToList();

        foreach (var span in annotationSpans)
        {
            // Remove label child spans
            var labelSpans = span.Elements("span")
                .Where(e =>
                {
                    var cls = (string?)e.Attribute("class") ?? "";
                    return cls.Contains($"{cssClassPrefix}label");
                })
                .ToList();

            foreach (var labelSpan in labelSpans)
            {
                labelSpan.Remove();
            }

            // Replace the annotation span with its remaining content (unwrap)
            var parent = span.Parent;
            if (parent != null)
            {
                var nodes = span.Nodes().ToList();
                foreach (var node in nodes)
                {
                    span.AddBeforeSelf(node);
                }
                span.Remove();
            }
        }

        return SerializeHtmlString(htmlDoc, wasWrapped);
    }

    /// <summary>
    /// Generate CSS to hide annotations with specific label IDs.
    /// This enables CSS-based label filtering without re-rendering.
    /// </summary>
    /// <param name="hiddenLabelIds">Label IDs to hide.</param>
    /// <param name="cssClassPrefix">CSS class prefix (default: "ext-annot-").</param>
    /// <returns>CSS string that hides the specified labels.</returns>
    public static string GenerateVisibilityCss(
        IEnumerable<string> hiddenLabelIds,
        string cssClassPrefix = "ext-annot-")
    {
        if (hiddenLabelIds == null) throw new ArgumentNullException(nameof(hiddenLabelIds));

        var css = new StringBuilder();
        css.AppendLine("/* Annotation Visibility Overrides */");

        foreach (var labelId in hiddenLabelIds)
        {
            var safeId = labelId.Replace(" ", "-").Replace(".", "-");
            // Hide the highlight styling but keep the text visible
            css.AppendLine($".{cssClassPrefix}highlight[data-label-id=\"{safeId}\"] {{");
            css.AppendLine("  background-color: transparent !important;");
            css.AppendLine("  border-bottom: none !important;");
            css.AppendLine("}");
            // Hide the label text
            css.AppendLine($".{cssClassPrefix}highlight[data-label-id=\"{safeId}\"] .{cssClassPrefix}label {{");
            css.AppendLine("  display: none !important;");
            css.AppendLine("}");
        }

        return css.ToString();
    }

    /// <summary>
    /// Generate annotation CSS for a set of labels.
    /// Useful when you need the CSS separately from the HTML (e.g., for incremental updates).
    /// </summary>
    /// <param name="labels">Label definitions.</param>
    /// <param name="settings">Projection settings.</param>
    /// <returns>CSS string for the given labels and settings.</returns>
    public static string GenerateAnnotationCssString(
        Dictionary<string, AnnotationLabel> labels,
        ExternalAnnotationProjectionSettings? settings = null)
    {
        if (labels == null) throw new ArgumentNullException(nameof(labels));
        settings ??= new ExternalAnnotationProjectionSettings();
        return BuildAnnotationCssString(labels, settings);
    }

    /// <summary>
    /// Add CSS for a single annotation to existing HTML.
    /// Used by AddAnnotationToHtml to inject per-label color classes.
    /// </summary>
    private static void AddSingleAnnotationCss(
        XElement html,
        OpenContractsAnnotation annotation,
        AnnotationLabel label,
        ExternalAnnotationProjectionSettings settings)
    {
        var prefix = settings.CssClassPrefix;
        var safeId = (annotation.AnnotationLabel ?? "").Replace(" ", "-").Replace(".", "-");

        var css = new StringBuilder();
        css.AppendLine();
        css.AppendLine($"/* Annotation label: {safeId} */");
        css.AppendLine($".{prefix}label-{safeId} {{");
        css.AppendLine($"  --annot-color: {label.Color};");
        css.AppendLine("}");

        var head = html.Descendants()
            .FirstOrDefault(e => e.Name.LocalName.Equals("head", StringComparison.OrdinalIgnoreCase));

        if (head != null)
        {
            var style = new XElement("style",
                new XAttribute("type", "text/css"),
                new XText(css.ToString()));
            head.Add(style);
        }
    }

    #endregion

    #region HTML Fragment Parsing

    // Synthetic wrapper element used to handle HTML fragments with multiple root elements.
    // XElement.Parse() requires a single root, but sanitized HTML (e.g., DOMPurify output)
    // often has multiple top-level elements like <style>...<div>...
    private const string FragmentWrapper = "docxodus-fragment-root";

    /// <summary>
    /// Common HTML named entities mapped to their numeric XML equivalents.
    /// XML only supports &amp; &lt; &gt; &quot; &apos; natively.
    /// </summary>
    private static readonly Dictionary<string, string> HtmlEntities = new(StringComparer.Ordinal)
    {
        { "&nbsp;", "&#160;" },
        { "&ndash;", "&#8211;" },
        { "&mdash;", "&#8212;" },
        { "&lsquo;", "&#8216;" },
        { "&rsquo;", "&#8217;" },
        { "&ldquo;", "&#8220;" },
        { "&rdquo;", "&#8221;" },
        { "&bull;", "&#8226;" },
        { "&hellip;", "&#8230;" },
        { "&trade;", "&#8482;" },
        { "&copy;", "&#169;" },
        { "&reg;", "&#174;" },
        { "&deg;", "&#176;" },
        { "&plusmn;", "&#177;" },
        { "&times;", "&#215;" },
        { "&divide;", "&#247;" },
        { "&laquo;", "&#171;" },
        { "&raquo;", "&#187;" },
        { "&cent;", "&#162;" },
        { "&pound;", "&#163;" },
        { "&euro;", "&#8364;" },
        { "&sect;", "&#167;" },
        { "&para;", "&#182;" },
        { "&micro;", "&#181;" },
        { "&frac12;", "&#189;" },
        { "&frac14;", "&#188;" },
        { "&frac34;", "&#190;" },
    };

    /// <summary>
    /// Parse an HTML string that may contain multiple root elements or HTML named entities.
    /// Wraps in a synthetic root and replaces HTML entities with numeric equivalents.
    /// </summary>
    /// <param name="html">HTML string to parse.</param>
    /// <param name="wasWrapped">True if a synthetic wrapper was added (i.e., input had multiple roots).</param>
    /// <returns>Parsed XElement.</returns>
    private static XElement ParseHtmlString(string html, out bool wasWrapped)
    {
        // Replace HTML named entities with numeric equivalents for XML compatibility
        var xmlSafe = html;
        foreach (var (entity, numeric) in HtmlEntities)
        {
            if (xmlSafe.Contains(entity))
                xmlSafe = xmlSafe.Replace(entity, numeric);
        }

        // Try parsing as-is first (single root element)
        try
        {
            var result = XElement.Parse(xmlSafe);
            wasWrapped = false;
            return result;
        }
        catch (System.Xml.XmlException)
        {
            // Multiple roots or other XML issue - wrap in synthetic root
            wasWrapped = true;
            return XElement.Parse($"<{FragmentWrapper}>{xmlSafe}</{FragmentWrapper}>");
        }
    }

    /// <summary>
    /// Serialize an XElement back to string, removing the synthetic wrapper if one was added.
    /// </summary>
    private static string SerializeHtmlString(XElement element, bool wasWrapped)
    {
        if (!wasWrapped)
            return element.ToString();

        // Remove the synthetic wrapper - return just the inner content
        var sb = new StringBuilder();
        foreach (var node in element.Nodes())
        {
            sb.Append(node.ToString());
        }
        return sb.ToString();
    }

    #endregion

    #region CSS Generation

    private static string BuildAnnotationCssString(
        Dictionary<string, AnnotationLabel> labels,
        ExternalAnnotationProjectionSettings settings)
    {
        var prefix = settings.CssClassPrefix;

        var css = new StringBuilder();
        css.AppendLine();
        css.AppendLine("/* External Annotation Styles */");
        css.AppendLine($".{prefix}highlight {{");
        css.AppendLine("  background-color: color-mix(in srgb, var(--annot-color, #FFEB3B) 30%, transparent);");
        css.AppendLine("  border-bottom: 2px solid var(--annot-color, #FFEB3B);");
        css.AppendLine("  display: inline;");
        css.AppendLine("  padding: 0.1em 0;");
        css.AppendLine("}");

        // Both Above and Inline use inline display - Above just has a superscript style
        if (settings.LabelMode == AnnotationLabelMode.Above)
        {
            css.AppendLine($".{prefix}label {{");
            css.AppendLine("  font-size: 0.65em;");
            css.AppendLine("  background-color: var(--annot-color, #FFEB3B);");
            css.AppendLine("  color: white;");
            css.AppendLine("  padding: 0.1em 0.3em;");
            css.AppendLine("  border-radius: 3px;");
            css.AppendLine("  vertical-align: super;");
            css.AppendLine("  margin-right: 0.2em;");
            css.AppendLine("}");
        }
        else if (settings.LabelMode == AnnotationLabelMode.Inline)
        {
            css.AppendLine($".{prefix}label {{");
            css.AppendLine("  font-size: 0.65em;");
            css.AppendLine("  background-color: var(--annot-color, #FFEB3B);");
            css.AppendLine("  color: white;");
            css.AppendLine("  padding: 0.1em 0.3em;");
            css.AppendLine("  border-radius: 3px;");
            css.AppendLine("  margin-right: 0.2em;");
            css.AppendLine("}");
        }

        // Add per-label color classes
        foreach (var (id, label) in labels)
        {
            var safeId = id.Replace(" ", "-").Replace(".", "-");
            css.AppendLine($".{prefix}label-{safeId} {{");
            css.AppendLine($"  --annot-color: {label.Color};");
            css.AppendLine("}");
        }

        return css.ToString();
    }

    private static void AddAnnotationCss(
        XElement html,
        Dictionary<string, AnnotationLabel> labels,
        ExternalAnnotationProjectionSettings settings)
    {
        var css = BuildAnnotationCssString(labels, settings);

        // Find or create head element
        var head = html.Descendants()
            .FirstOrDefault(e => e.Name.LocalName.Equals("head", StringComparison.OrdinalIgnoreCase));

        if (head != null)
        {
            var style = new XElement("style",
                new XAttribute("type", "text/css"),
                new XText(css));
            head.Add(style);
        }
    }

    #endregion
}
