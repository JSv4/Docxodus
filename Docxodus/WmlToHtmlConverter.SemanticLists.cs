// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus
{
    /// <summary>
    /// Semantic lists (issue #895): with <see cref="WmlToHtmlConverterSettings.SemanticLists"/> on,
    /// Word list paragraphs come out as <c>ol</c>/<c>ul</c>/<c>li</c> rather than as paragraphs
    /// carrying a generated marker. Formatting assembly stamps each list paragraph with its list,
    /// level, counter value and number format; <see cref="ConvertParagraph"/> carries them to the
    /// converted element as a <see cref="SemanticListItem"/>; and <see cref="GroupSemanticLists"/>
    /// then groups adjacent items, after conversion and before styles become classes.
    /// </summary>
    public static partial class WmlToHtmlConverter
    {
        /// <summary>The list facts of one converted list paragraph.</summary>
        private sealed record SemanticListItem(int? NumId, int Level, int? Value, string? Format, bool IsBidi)
        {
            public bool Ordered => Format != "bullet";
        }

        /// <summary>Paginated output measures and splits paragraphs, so it never groups them.</summary>
        private static bool EmitsSemanticLists(WmlToHtmlConverterSettings settings) =>
            settings.SemanticLists && settings.RenderPagination != PaginationMode.Paginated;

        /// <summary>Carries a list paragraph's stamped list facts to its converted element. A
        /// numbered heading stays a heading.</summary>
        private static void AnnotateSemanticListItem(XElement converted, XElement paragraph, XName elementName, bool isBidi)
        {
            if (elementName != Xhtml.p || paragraph.Attribute(PtOpenXml.ListLevel) is not { } level)
                return;
            converted.AddAnnotation(new SemanticListItem(
                (int?)paragraph.Attribute(PtOpenXml.ListNumId),
                (int)level,
                (int?)paragraph.Attribute(PtOpenXml.ListValue),
                (string?)paragraph.Attribute(PtOpenXml.ListFormat),
                isBidi));
        }

        /// <summary>
        /// Turns every run of adjacent list paragraphs into lists. Items of one list at one level
        /// share an <c>ol</c> (numbered) or <c>ul</c> (bulleted); a deeper level opens a list inside
        /// the preceding item; a change of list, or of numbered to bulleted, at the same level starts
        /// a sibling list. Anything else between two items (a plain paragraph, a table) ends the run,
        /// so a list never crosses a cell, a note or a section.
        /// </summary>
        private static void GroupSemanticLists(XElement root)
        {
            var parents = root.DescendantsAndSelf()
                .Where(e => e.Elements().Any(c => c.Annotation<SemanticListItem>() != null))
                .ToList();
            foreach (var parent in parents)
            {
                var run = new List<XElement>();
                foreach (var node in parent.Nodes().ToList())
                {
                    if (node is XElement element && element.Annotation<SemanticListItem>() != null)
                        run.Add(element);
                    else if (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))
                        BuildLists(run);
                }
                BuildLists(run);
            }
        }

        /// <summary>A list being built: its element, the item that opened it, and the item it is nested in.</summary>
        private sealed class OpenList(XElement list, SemanticListItem first, XElement? parentItem)
        {
            public XElement List { get; } = list;
            public SemanticListItem First { get; } = first;
            public XElement? ParentItem { get; } = parentItem;
            public List<(XElement Item, SemanticListItem Info)> Items { get; } = new();
        }

        /// <summary>Replaces one run of adjacent list paragraphs with the lists they form, and empties the run.</summary>
        private static void BuildLists(List<XElement> run)
        {
            if (run.Count == 0)
                return;

            var stack = new List<OpenList>();
            var all = new List<OpenList>();
            var roots = new List<XElement>();
            foreach (var paragraph in run)
            {
                var info = paragraph.Annotation<SemanticListItem>()!;
                while (stack.Count > 0 && stack[^1].First.Level > info.Level)
                    stack.RemoveAt(stack.Count - 1);
                if (stack.Count > 0 && stack[^1].First.Level == info.Level &&
                    (stack[^1].First.NumId != info.NumId || stack[^1].First.Ordered != info.Ordered))
                    stack.RemoveAt(stack.Count - 1);

                if (stack.Count == 0 || stack[^1].First.Level < info.Level)
                {
                    var parentItem = stack.Count > 0 ? stack[^1].Items[^1].Item : null;
                    var open = new OpenList(new XElement(info.Ordered ? Xhtml.ol : Xhtml.ul), info, parentItem);
                    if (parentItem != null)
                        parentItem.Add(open.List);
                    else
                        roots.Add(open.List);
                    stack.Add(open);
                    all.Add(open);
                }

                stack[^1].Items.Add((paragraph, info));
            }

            run[0].AddBeforeSelf(roots);
            foreach (var paragraph in run)
            {
                paragraph.Remove();
                paragraph.Name = Xhtml.li;
            }
            foreach (var open in all)
                foreach (var (item, _) in open.Items)
                    open.List.Add(item);
            // Outer lists first, so a nested list measures its items from its parent item's final indent.
            foreach (var open in all)
                FinishList(open);
            run.Clear();
        }

        /// <summary>
        /// Styles one list and its items. Each item keeps its own leading indent, measured from the
        /// item it is nested in (or from the container at the top level), so indents never compound.
        /// CSS draws the markers only when it reproduces every item's marker text exactly and each
        /// marker hangs in its item's hanging indent before a tab; otherwise every item keeps its
        /// generated marker spans and the list draws none.
        /// </summary>
        private static void FinishList(OpenList open)
        {
            var listStyle = Style(open.List);
            listStyle["margin-top"] = "0";
            listStyle["margin-bottom"] = "0";
            listStyle["padding"] = "0";
            listStyle["text-indent"] = "0";
            if (open.ParentItem != null)
            {
                // A nested list sits inside its parent item's box: without these, its items would
                // inherit the parent paragraph's alignment and line height where their own paragraphs
                // set none (text-indent is reset above for the same reason).
                listStyle["text-align"] = "start";
                listStyle["line-height"] = "normal";
            }

            var container = open.ParentItem is { } parent ? IndentsInches(parent) : (Leading: 0m, Trailing: 0m);
            var types = open.Items.Select(i => CssMarkerType(i.Item, i.Info)).Distinct().ToList();
            var cssType = types.Count == 1 ? types[0] : null;
            listStyle["list-style-type"] = cssType ?? "none";

            int? previous = null;
            for (var index = 0; index < open.Items.Count; index++)
            {
                var (item, info) = open.Items[index];
                var style = Style(item);
                var (leadingSide, trailingSide) = info.IsBidi ? ("margin-right", "margin-left") : ("margin-left", "margin-right");
                var indents = IndentsInches(item);
                style[leadingSide] = FormatInches(indents.Leading - container.Leading);
                style[trailingSide] = FormatInches(indents.Trailing - container.Trailing);
                // Remember the item's absolute indents for any list nested inside it.
                item.AddAnnotation(new AbsoluteIndents(indents.Leading, indents.Trailing));

                // An item that keeps its own marker lays out as the paragraph it was. A list item
                // box would also differ from it in a quirks-mode page, where Chromium sizes the line
                // of a list item differently.
                if (cssType == null)
                    style["display"] = "block";
                else
                {
                    item.Elements().First(IsMarkerWrapper).Remove();
                    style["text-indent"] = "0";
                    if (open.First.Ordered && info.Value is { } value)
                    {
                        if (index == 0 && value != 1)
                            open.List.SetAttributeValue("start", value);
                        else if (index > 0 && previous is { } before && value != before + 1)
                            item.SetAttributeValue("value", value);
                        previous = value;
                    }
                }
            }

            // An item's space after belongs below its own text, which is now above the nested list.
            if (open.ParentItem is { } owner && Style(owner).Remove("margin-bottom", out var spaceAfter))
                listStyle["margin-top"] = spaceAfter;
        }

        private sealed record AbsoluteIndents(decimal Leading, decimal Trailing);

        private static Dictionary<string, string> Style(XElement element)
        {
            var style = element.Annotation<Dictionary<string, string>>();
            if (style == null)
            {
                style = new Dictionary<string, string>();
                element.AddAnnotation(style);
            }
            return style;
        }

        /// <summary>The item's leading and trailing indents in inches from its container, as the
        /// paragraph conversion wrote them (<c>0</c> or <c>N.NNin</c>).</summary>
        private static (decimal Leading, decimal Trailing) IndentsInches(XElement item)
        {
            if (item.Annotation<AbsoluteIndents>() is { } known)
                return (known.Leading, known.Trailing);
            var bidi = item.Annotation<SemanticListItem>()?.IsBidi == true;
            var style = item.Annotation<Dictionary<string, string>>();
            decimal Inches(string property) =>
                style != null && style.TryGetValue(property, out var value) && value.EndsWith("in", StringComparison.Ordinal) &&
                decimal.TryParse(value[..^2], NumberStyles.Number, CultureInfo.InvariantCulture, out var inches)
                    ? inches
                    : 0m;
            return bidi ? (Inches("margin-right"), Inches("margin-left")) : (Inches("margin-left"), Inches("margin-right"));
        }

        private static string FormatInches(decimal inches) =>
            inches == 0m ? "0" : string.Format(NumberFormatInfo.InvariantInfo, "{0:0.00}in", inches);

        /// <summary>The generated marker: a span tagged as a list marker that holds the marker text
        /// and the tab that follows it.</summary>
        private static bool IsMarkerWrapper(XElement element) =>
            element.Name == Xhtml.span && (string?)element.Attribute("data-list-marker") == "true" &&
            element.Descendants(Xhtml.span).Any(s => s.Attribute("data-docx-tab") != null);

        /// <summary>
        /// The <c>list-style-type</c> that draws this item's marker in Word's place, or null. CSS can
        /// stand in only when the marker is followed by a tab, hangs in a hanging indent (a negative
        /// first-line indent), and its text is exactly what the type generates for the item's value:
        /// a number in one of CSS's counter styles followed by a period, or a bullet glyph CSS draws.
        /// </summary>
        private static string? CssMarkerType(XElement item, SemanticListItem info)
        {
            var wrapper = item.Elements().FirstOrDefault();
            if (wrapper == null || !IsMarkerWrapper(wrapper))
                return null;
            if (item.Annotation<Dictionary<string, string>>() is not { } style ||
                !style.TryGetValue("text-indent", out var textIndent) || !textIndent.StartsWith("-", StringComparison.Ordinal))
                return null;
            // A ::marker takes the item's own formatting, so the marker run must already look like
            // the item: same family, size, weight, style, colour and decoration. And where the marker
            // run's font differs from the text's, its line box can be the tallest on the first line
            // (Word counts it too), so dropping the span would shorten the line.
            var markerFormat = EffectiveFormat(MarkerTextSpan(wrapper), style);
            var textFormat = EffectiveFormat(item.Elements().Skip(1).FirstOrDefault(e => e.Name == Xhtml.span && e.Value.Trim().Length > 0), style);
            if (markerFormat != EffectiveFormat(null, style) ||
                markerFormat.Family != textFormat.Family || markerFormat.Size != textFormat.Size)
                return null;
            var marker = wrapper.Value.Trim().Trim('\u200e', '\u200f');

            if (!info.Ordered)
                return marker switch
                {
                    "\u2022" => "disc",
                    "\u25e6" => "circle",
                    "\u25aa" or "\u25a0" => "square",
                    _ => null,
                };

            if (info.Value is not { } value || value < 1)
                return null;
            var (type, text) = info.Format switch
            {
                "decimal" => ("decimal", value.ToString(CultureInfo.InvariantCulture)),
                "decimalZero" => ("decimal-leading-zero", value.ToString("00", CultureInfo.InvariantCulture)),
                "lowerLetter" => ("lower-alpha", Alphabetic(value)),
                "upperLetter" => ("upper-alpha", Alphabetic(value).ToUpperInvariant()),
                "lowerRoman" => ("lower-roman", Roman(value)?.ToLowerInvariant()),
                "upperRoman" => ("upper-roman", Roman(value)),
                _ => ((string?)null, (string?)null),
            };
            return text != null && marker == text + "." ? type : null;
        }

        /// <summary>The innermost marker span: the one holding the marker text.</summary>
        private static XElement? MarkerTextSpan(XElement wrapper) =>
            wrapper.Descendants(Xhtml.span).LastOrDefault(s => (string?)s.Attribute("data-list-marker") == "true");

        /// <summary>A span's text formatting, each property falling back to the item's own and then
        /// to the CSS initial value; with no span, the item's own formatting.</summary>
        private static (string? Family, string? Size, string Weight, string Style, string? Color, string? Decoration) EffectiveFormat(
            XElement? span, Dictionary<string, string> itemStyle)
        {
            var style = span?.Annotation<Dictionary<string, string>>();
            string? Get(string property) =>
                style != null && style.TryGetValue(property, out var value) ? value
                : itemStyle.TryGetValue(property, out var inherited) ? inherited : null;
            return (Get("font-family"), Get("font-size"), Get("font-weight") ?? "normal", Get("font-style") ?? "normal",
                Get("color"), Get("text-decoration"));
        }

        /// <summary>CSS <c>lower-alpha</c>: a..z, aa, ab, … (bijective base 26). Word repeats the
        /// letter after z (aa, bb), so its 27th marker differs and keeps its span.</summary>
        private static string Alphabetic(int value)
        {
            var text = new StringBuilder();
            for (var n = value; n > 0; n = (n - 1) / 26)
                text.Insert(0, (char)('a' + (n - 1) % 26));
            return text.ToString();
        }

        /// <summary>CSS <c>upper-roman</c>, defined for 1–3999.</summary>
        private static string? Roman(int value)
        {
            if (value > 3999)
                return null;
            var text = new StringBuilder();
            foreach (var (amount, numeral) in RomanNumerals)
                for (; value >= amount; value -= amount)
                    text.Append(numeral);
            return text.ToString();
        }

        private static readonly (int Amount, string Numeral)[] RomanNumerals =
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
        };
    }
}
