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

        // Build the text-to-element mapping
        var textMap = BuildTextMap(result);
        var htmlText = GetHtmlText(textMap);

        // Track which offsets we've used (for handling multiple occurrences)
        var usedOffsets = new HashSet<int>();

        // Sort annotations by start offset for correct nesting
        var sortedAnnotations = annotationSet.LabelledText
            .Where(a => !a.Structural && a.AnnotationJson is TextSpan)
            .Select(a => (Annotation: a, Span: (TextSpan)a.AnnotationJson!))
            .OrderBy(x => x.Span.Start)
            .ThenByDescending(x => x.Span.End) // Longer spans first for nesting
            .ToList();

        // Project each annotation using text search (not offsets)
        foreach (var (annotation, span) in sortedAnnotations)
        {
            var label = annotationSet.TextLabels.TryGetValue(annotation.AnnotationLabel, out var l)
                ? l : null;

            // Use the annotation's raw text to find it in the HTML
            var searchText = span.Text ?? annotation.RawText;
            if (string.IsNullOrEmpty(searchText)) continue;

            // Find this text in the HTML
            var htmlLocation = FindTextInHtml(htmlText, searchText, usedOffsets);
            if (htmlLocation == null) continue;

            // Create a synthetic span with HTML-space offsets
            var htmlSpan = new TextSpan
            {
                Id = span.Id,
                Start = htmlLocation.Value.start,
                End = htmlLocation.Value.end,
                Text = searchText
            };

            // Rebuild text map since we may have modified it in previous iteration
            textMap = BuildTextMap(result);

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
                if (name != "script" && name != "style")
                {
                    foreach (var childText in GetTextNodes(child))
                    {
                        yield return childText;
                    }
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

    #region CSS Generation

    private static void AddAnnotationCss(
        XElement html,
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

        // Find or create head element
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
}
