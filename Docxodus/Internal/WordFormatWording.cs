using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// What Word prints in an All Markup balloon for a tracked formatting change (issue #1300): the text
/// after "Formatted:", naming each property that differs between the two sides by its new value,
/// for example <c>Font: Bold</c> or <c>Indent: Left: 0.5"</c>.
///
/// <see cref="Stamp"/> runs before formatting assembly, while each change still holds direct
/// properties. Both sides go through the assembler's own read-only resolvers, so whatever styles
/// and document defaults give both sides alike drops out of the comparison. The results are
/// written as attributes in <see cref="BalloonData"/> onto the change element, which assembly
/// carries to the converter; the converter publishes each as <c>data-</c> plus its local name.
/// </summary>
internal static class WordFormatWording
{
    /// <summary>The namespace of the balloon data stamped on change elements.</summary>
    internal static readonly XNamespace BalloonData = "urn:docxodus:rendering:balloon";

    /// <summary>
    /// Stamps every <c>w:rPrChange</c> and <c>w:pPrChange</c> of the content phrases with Word's
    /// wording (<c>format-description</c>) and its author's colour (<c>author-color</c>, from
    /// <paramref name="authorColor"/>). A run's change also gets <c>format-key</c>: changes that
    /// recorded the same properties share it, which is how runs of one revision are recognised.
    /// </summary>
    internal static void Stamp(WordprocessingDocument wDoc, Func<string?, string?> authorColor)
    {
        var recordedSets = new Dictionary<string, int>(StringComparer.Ordinal);
        // Runs repeat the same change many times over; resolving styles is the costly part.
        var runWordings = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var part in wDoc.ContentParts())
        {
            foreach (var change in part.GetXDocument().Descendants()
                .Where(e => e.Name == W.rPrChange || e.Name == W.pPrChange))
            {
                var properties = change.Parent;
                var paragraph = change.Ancestors(W.p).FirstOrDefault();
                if (properties == null || paragraph == null)
                    continue;
                string wording;
                if (change.Name == W.pPrChange)
                    wording = DescribeParagraph(wDoc, paragraph, change);
                else
                {
                    // The wording depends only on the paragraph's properties and the run's two sides.
                    string Text(IEnumerable<XElement>? elements) =>
                        string.Concat((elements ?? Enumerable.Empty<XElement>()).Select(e => e.ToString(SaveOptions.DisableFormatting)));
                    var signature = string.Join("\u0001", Text(paragraph.Element(W.pPr)?.Elements().Where(e => e.Name != W.pPrChange)),
                        Text(properties.Elements().Where(e => e.Name != W.rPrChange)), Text(change.Element(W.rPr)?.Elements()));
                    if (!runWordings.TryGetValue(signature, out var known))
                        runWordings[signature] = known = DescribeRun(EffectiveRun(wDoc, paragraph, properties),
                            EffectiveRun(wDoc, paragraph, change.Element(W.rPr)));
                    wording = known;
                    if (properties.Parent?.Name == W.r)
                    {
                        var recorded = RecordedSet(change.Element(W.rPr));
                        if (!recordedSets.TryGetValue(recorded, out var ordinal))
                            recordedSets[recorded] = ordinal = recordedSets.Count;
                        change.SetAttributeValue(BalloonData + "format-key", ordinal);
                    }
                }
                change.SetAttributeValue(BalloonData + "format-description", wording);
                change.SetAttributeValue(BalloonData + "author-color", authorColor((string?)change.Attribute(W.author)));
            }
        }
    }

    /// <summary>
    /// The text of a change's balloon, "Formatted:" and its wording (or a bare "Formatted"), which the
    /// Word presentation also uses as the change's tooltip. Null when <see cref="Stamp"/> did not run.
    /// </summary>
    internal static string? BalloonText(XElement change) =>
        change.Attribute(BalloonData + "format-description")?.Value is { } wording
            ? wording.Length == 0 ? "Formatted" : $"Formatted: {wording}"
            : null;

    /// <summary>
    /// The family and size of Word's balloon text: the document's Balloon Text paragraph style,
    /// or Word's own default, Segoe UI 9 pt.
    /// </summary>
    internal static (string Family, double SizePt) BalloonTextFont(WordprocessingDocument wDoc)
    {
        var hasStyle = wDoc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root?.Elements(W.style)
            .Any(s => (string?)s.Attribute(W.type) == "paragraph" && (string?)s.Attribute(W.styleId) == "BalloonText") == true;
        if (!hasStyle)
            return ("Segoe UI", 9);
        var probe = new XElement(W.p, new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, "BalloonText"))));
        var effective = EffectiveRun(wDoc, probe, null);
        var halfPoints = Number(Get(effective, W.sz));
        return (Get(effective, W.rFonts, W.ascii) ?? "Segoe UI", halfPoints > 0 ? halfPoints / 2 : 9);
    }

    /// <summary>The font family of a comment's text: that of its first run with text.</summary>
    internal static string? CommentTextFont(WordprocessingDocument wDoc, XElement comment)
    {
        var run = comment.Descendants(W.t).Select(t => t.Parent).FirstOrDefault(r => r?.Element(W.annotationRef) == null);
        var paragraph = run?.Ancestors(W.p).FirstOrDefault();
        return paragraph == null ? null : Get(EffectiveRun(wDoc, paragraph, run!.Element(W.rPr)), W.rFonts, W.ascii);
    }

    private static readonly (XName Property, string Name)[] FontToggles =
    {
        (W.b, "Bold"), (W.i, "Italic"), (W.caps, "All caps"), (W.smallCaps, "Small caps"), (W.vanish, "Hidden"),
    };

    internal static string DescribeRun(XElement now, XElement was)
    {
        var fontPhrases = new List<string>();
        Changed(fontPhrases, now, was, p => Get(p, W.rFonts, W.ascii), v => $"(Default) {v}");
        Changed(fontPhrases, now, was, p => Get(p, W.rFonts, W.eastAsia), v => $"(Asian) {v}");
        Changed(fontPhrases, now, was, p => Get(p, W.sz), v => Points(Number(v) / 2));
        foreach (var (property, name) in FontToggles)
            Toggled(fontPhrases, now, was, property, name, "Not " + name);
        Changed(fontPhrases, now, was, p => Get(p, W.vertAlign) ?? "baseline", v => v switch
        {
            "superscript" => "Superscript",
            "subscript" => "Subscript",
            _ => "Not Superscript/ Subscript",
        });

        var phrases = fontPhrases.Count == 0 ? new List<string>() : new List<string> { "Font: " + string.Join(", ", fontPhrases) };
        Changed(phrases, now, was, ColorName, v => $"Font color: {v}");
        Changed(phrases, now, was, p => Get(p, W.u) ?? "none",
            v => v == "none" ? "No underline" : v == "double" ? "Double underline" : "Underline");
        Toggled(phrases, now, was, W.strike, "Strikethrough", "Not Strikethrough");
        Toggled(phrases, now, was, W.dstrike, "Double strikethrough", "Not Double strikethrough");
        Changed(phrases, now, was, p => Get(p, W.highlight) is null or "none" ? "Not Highlight" : "Highlight", v => v);
        return string.Join(", ", phrases);
    }

    private static string DescribeParagraph(WordprocessingDocument wDoc, XElement paragraph, XElement change)
    {
        var now = EffectiveParagraph(wDoc, paragraph.Element(W.pPr));
        var was = EffectiveParagraph(wDoc, change.Element(W.pPr));
        var styles = wDoc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root?.Elements(W.style)
            .Where(s => (string?)s.Attribute(W.type) == "paragraph").ToList() ?? new List<XElement>();
        var defaultStyle = (string?)styles.FirstOrDefault(s => s.Attribute(W._default).ToBoolean() == true)?.Attribute(W.styleId);
        string? StyleName(XElement pPr)
        {
            var id = Get(pPr, W.pStyle) ?? defaultStyle;
            return id == null ? null
                : (string?)styles.FirstOrDefault(s => (string?)s.Attribute(W.styleId) == id)?.Element(W.name)?.Attribute(W.val) ?? id;
        }

        var phrases = new List<string>(4);
        Changed(phrases, now, was, StyleName, v => v);
        Changed(phrases, now, was, p => Alignments.GetValueOrDefault(Get(p, W.jc) ?? "left", "Left"), v => v);

        var nowList = ListLevel(now);
        var wasList = ListLevel(was);
        // New numbering, or a new level, brings its own indentation; Word names only the numbering.
        var indentComesWithList = nowList != null && nowList != wasList && paragraph.Element(W.pPr)?.Element(W.ind) == null;
        if (!indentComesWithList)
            Changed(phrases, now, was, Indent, v => $"Indent: {v}");

        var spacing = new List<string>();
        Changed(spacing, now, was, p => Get(p, W.spacing, W.before) ?? "0", v => $"Before: {Points(Number(v) / 20)}");
        Changed(spacing, now, was, p => Get(p, W.spacing, W.after) ?? "0", v => $"After: {Points(Number(v) / 20)}");
        if (spacing.Count > 0)
            phrases.Add($"Space {string.Join(", ", spacing)}");
        Changed(phrases, now, was, LineSpacing, v => $"Line spacing: {v}");

        if (nowList == null && wasList != null)
            phrases.Add("No bullets or numbering");
        else if (nowList is { } list && list.NumId != wasList?.NumId)
            phrases.Add(DescribeList(wDoc, paragraph, list));

        Toggled(phrases, now, was, W.contextualSpacing, "Don't add space between paragraphs of the same style",
            "Add space between paragraphs of the same style");
        Toggled(phrases, now, was, W.keepNext, "Keep with next", "Don't keep with next");
        Toggled(phrases, now, was, W.keepLines, "Keep lines together", "Don't keep lines together");
        Toggled(phrases, now, was, W.pageBreakBefore, "Page break before", "No page break before");
        return string.Join(", ", phrases);
    }

    private static readonly Dictionary<string, string> Alignments = new(StringComparer.Ordinal)
    {
        ["left"] = "Left", ["start"] = "Left", ["center"] = "Centered", ["right"] = "Right", ["end"] = "Right",
        ["both"] = "Justified", ["distribute"] = "Distributed",
    };

    /// <summary>Word gives the whole new indent; its Left is where a hanging first line starts.</summary>
    private static string Indent(XElement pPr)
    {
        double Twips(params XName[] names) => Number(names.Select(n => Get(pPr, W.ind, n)).FirstOrDefault(v => v != null));
        var (left, hanging, firstLine, right) = (Twips(W.left, W.start), Twips(W.hanging), Twips(W.firstLine), Twips(W.right, W.end));
        var text = $"Left: {Inches(left - hanging)}";
        text += hanging != 0 ? $", Hanging: {Inches(hanging)}" : firstLine != 0 ? $", First line: {Inches(firstLine)}" : "";
        return right != 0 ? $"{text}, Right: {Inches(right)}" : text;
    }

    private static string LineSpacing(XElement pPr)
    {
        var line = Number(Get(pPr, W.spacing, W.line) ?? "240");
        return Get(pPr, W.spacing, W.lineRule) switch
        {
            "atLeast" => $"At least {Points(line / 20)}",
            "exact" => $"Exactly {Points(line / 20)}",
            _ when line == 240 => "single",
            _ when line == 360 => "1.5 lines",
            _ when line == 480 => "Double",
            _ => $"Multiple {Rounded(line / 240, "0.##")} li",
        };
    }

    private static (string NumId, int Level)? ListLevel(XElement pPr)
    {
        var numId = Get(pPr.Element(W.numPr), W.numId);
        return numId is null or "0" ? null : (numId, (int)Number(Get(pPr.Element(W.numPr), W.ilvl)));
    }

    private static readonly Dictionary<string, string> NumberingStyles = new(StringComparer.Ordinal)
    {
        ["decimal"] = "1, 2, 3, …", ["lowerLetter"] = "a, b, c, …", ["upperLetter"] = "A, B, C, …",
        ["lowerRoman"] = "i, ii, iii, …", ["upperRoman"] = "I, II, III, …",
    };

    /// <summary>A list level as Word describes it, found through the list-item resolver on a probe paragraph.</summary>
    private static string DescribeList(WordprocessingDocument wDoc, XElement paragraph, (string NumId, int Level) list)
    {
        var level = $"Level: {list.Level + 1}";
        var numbering = wDoc.MainDocumentPart?.NumberingDefinitionsPart?.GetXDocument();
        var styles = wDoc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument();
        var probe = new XElement(W.p, new XElement(W.pPr, paragraph.Element(W.pPr)?.Element(W.pStyle),
            new XElement(W.numPr, new XElement(W.ilvl, new XAttribute(W.val, list.Level)),
                new XElement(W.numId, new XAttribute(W.val, list.NumId)))));
        if (numbering != null && styles != null)
            ListItemRetriever.InitListItemInfo(numbering, styles, probe);
        var lvl = probe.Annotation<ListItemRetriever.ListItemInfo>() is { IsListItem: true } info ? info.Lvl(list.Level) : null;
        if (lvl == null)
            return $"Numbered + {level}";

        var ind = lvl.Element(W.pPr)?.Element(W.ind);
        double Twips(XName name) => Number((string?)ind?.Attribute(name));
        var indentAt = ind?.Attribute(W.left) != null ? Twips(W.left) : Twips(W.start);
        var placement = $"Aligned at: {Inches(indentAt - Twips(W.hanging) + Twips(W.firstLine))} + Indent at: {Inches(indentAt)}";
        var format = Get(lvl, W.numFmt) ?? "decimal";
        return format == "bullet"
            ? $"Bulleted + {level} + {placement}"
            : $"Numbered + {level} + Numbering Style: {NumberingStyles.GetValueOrDefault(format, format)}" +
              $" + Start at: {Get(lvl, W.start) ?? "1"} + Alignment: {Alignments.GetValueOrDefault(Get(lvl, W.lvlJc) ?? "left", "Left")}" +
              $" + {placement}";
    }

    private static readonly Dictionary<string, string> ThemeColors = new(StringComparer.Ordinal)
    {
        ["text1"] = "Text 1", ["dark1"] = "Text 1", ["text2"] = "Text 2", ["dark2"] = "Text 2",
        ["background1"] = "Background 1", ["light1"] = "Background 1", ["background2"] = "Background 2",
        ["light2"] = "Background 2", ["hyperlink"] = "Hyperlink", ["followedHyperlink"] = "Followed Hyperlink",
    };

    /// <summary>Word's names for its standard colours.</summary>
    private static readonly Dictionary<string, string> StandardColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["auto"] = "Auto", ["000000"] = "Black", ["FFFFFF"] = "White", ["C00000"] = "Dark Red", ["FF0000"] = "Red",
        ["FFC000"] = "Orange", ["FFFF00"] = "Yellow", ["92D050"] = "Light Green", ["00B050"] = "Green",
        ["00B0F0"] = "Light Blue", ["0070C0"] = "Blue", ["002060"] = "Dark Blue", ["7030A0"] = "Purple",
    };

    private static string ColorName(XElement rPr)
    {
        if (Get(rPr, W.color, W.themeColor) is { } theme)
        {
            var themeName = ThemeColors.TryGetValue(theme, out var known) ? known
                : theme.StartsWith("accent", StringComparison.Ordinal) ? $"Accent {theme.Substring(6)}" : theme;
            // A shade or tint keeps that fraction of the colour (hex 00-FF): Word names the rest.
            int Percent(string hex) => int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var kept)
                ? (int)Math.Round(100 - kept / 2.55)
                : 0;
            if (Get(rPr, W.color, W.themeShade) is { Length: 2 } shade && Percent(shade) > 0)
                return $"{themeName}, Darker {Percent(shade)}%";
            if (Get(rPr, W.color, W.themeTint) is { Length: 2 } tint && Percent(tint) > 0)
                return $"{themeName}, Lighter {Percent(tint)}%";
            return themeName;
        }
        var value = Get(rPr, W.color) ?? "auto";
        if (StandardColors.TryGetValue(value, out var name))
            return name;
        return value.Length == 6 && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)
            ? $"Custom Color(RGB({(rgb >> 16) & 255},{(rgb >> 8) & 255},{rgb & 255}))"
            : value;
    }

    /// <summary>
    /// The effective properties of a run carrying <paramref name="rPr"/>, resolved inside a detached
    /// paragraph with <paramref name="paragraph"/>'s properties so that its paragraph style applies.
    /// </summary>
    private static XElement EffectiveRun(WordprocessingDocument wDoc, XElement paragraph, XElement? rPr)
    {
        var context = new XElement(W.p,
            new XElement(W.pPr, paragraph.Element(W.pPr)?.Elements().Where(e => e.Name != W.pPrChange)),
            new XElement(W.r, new XElement(W.rPr, rPr?.Elements().Where(e => e.Name != W.rPrChange)), new XElement(W.t, "x")));
        return FormattingAssembler.ResolveEffectiveRunProperties(wDoc, context.Element(W.r)!);
    }

    private static XElement EffectiveParagraph(WordprocessingDocument wDoc, XElement? pPr) =>
        FormattingAssembler.ResolveEffectiveParagraphProperties(wDoc,
            new XElement(W.p, new XElement(W.pPr, pPr?.Elements().Where(e => e.Name != W.pPrChange))));

    /// <summary>Properties Word records per script or proofing pass; they never split a revision.</summary>
    private static readonly HashSet<XName> PerScript = new() { W.szCs, W.bCs, W.iCs, W.lang, W.noProof, W.rtl };

    /// <summary>The recorded properties in a canonical text form: attribute order, rsids and per-script properties do not count.</summary>
    private static string RecordedSet(XElement? rPr)
    {
        var canonical = (rPr?.Elements() ?? Enumerable.Empty<XElement>())
            .Where(property => !PerScript.Contains(property.Name))
            .Select(property => new XElement(property.Name.LocalName, property.Attributes()
                .Where(a => a.Name.Namespace == W.w)
                .Where(a => !a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal))
                .Select(a => new XAttribute(a.Name.LocalName, a.Value))
                .OrderBy(a => a.Name.LocalName, StringComparer.Ordinal)).ToString(SaveOptions.DisableFormatting))
            .ToList();
        canonical.Sort(StringComparer.Ordinal);
        return string.Concat(canonical);
    }

    private static void Changed(List<string> phrases, XElement now, XElement was, Func<XElement, string?> read, Func<string, string> word)
    {
        var value = read(now);
        if (value != null && value != read(was))
            phrases.Add(word(value));
    }

    private static void Toggled(List<string> phrases, XElement now, XElement was, XName property, string on, string off)
    {
        bool IsOn(XElement p) => p.Element(property) is { } e && e.Attribute(W.val).ToBoolean() != false;
        if (IsOn(now) != IsOn(was))
            phrases.Add(IsOn(now) ? on : off);
    }

    private static string? Get(XElement? properties, XName element, XName? attribute = null) =>
        properties?.Element(element)?.Attribute(attribute ?? W.val)?.Value;

    private static double Number(string? value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : 0;

    /// <summary>Formats a value; adding 0.0 turns a negative zero into a plain "0".</summary>
    private static string Rounded(double value, string format) => (value + 0.0).ToString(format, CultureInfo.InvariantCulture);

    private static string Inches(double twips) => Rounded(Math.Round(twips / 1440, 2), "0.##") + "\"";

    private static string Points(double points) => Rounded(points, "0.#") + " pt";
}
