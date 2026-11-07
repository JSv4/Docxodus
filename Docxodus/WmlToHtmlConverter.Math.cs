using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

public static partial class WmlToHtmlConverter
{
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    private sealed record MathRenderingContext(bool Supported, string FontFamily, WordprocessingDocument Document)
    {
        public Dictionary<string, decimal?> FontSizes { get; } = new();
    }

    private static bool HasUnknownMathAttributes(XElement element, XName? allowed = null) => element.Attributes()
        .Any(a => !a.IsNamespaceDeclaration && a.Name != PtOpenXml.Unid && a.Name != allowed);

    private static MathRenderingContext MathContext(WordprocessingDocument document)
    {
        var part = document.MainDocumentPart!;
        if (part.Annotation<MathRenderingContext>() is { } cached) return cached;
        var properties = part.DocumentSettingsPart?.GetXDocument().Root?.Element(M.mathPr);
        var fontElement = properties?.Element(M.mathFont);
        var font = (string?)fontElement?.Attribute(M.val);
        var supported = properties == null || (!HasUnknownMathAttributes(properties)
            && properties.Elements().All(e => e.Name == M.mathFont)
            && properties.Elements(M.mathFont).Count() <= 1
            && (fontElement == null || (!fontElement.HasElements && !string.IsNullOrWhiteSpace(font)
                && !HasUnknownMathAttributes(fontElement, M.val))));
        // An absent OMML mathFont permits font substitution. Let the browser's math family and
        // the existing export font policy expose that choice; never silently use the body font.
        var family = font == null ? "math" : "'" + string.Concat(font.Select(c =>
            c is '\\' or '\'' || char.IsControl(c) ? "\\" + ((int)c).ToString("X", CultureInfo.InvariantCulture) + " " : c.ToString())) + "', math";
        var context = new MathRenderingContext(supported, family, document);
        part.AddAnnotation(context);
        return context;
    }

    private static decimal? MathFontSize(XElement source, MathRenderingContext context)
    {
        var paragraph = source.Ancestors(W.p).FirstOrDefault();
        // The shared introspection resolver omits table-style layers. Keep that context
        // diagnosed until math participates in the full table/run formatting cascade.
        if (paragraph == null || source.Ancestors(W.tbl).Any()) return null;
        var style = paragraph.Element(W.pPr)?.Element(W.pStyle);
        var key = (string?)style?.Attribute(W.val) ?? "";
        if (context.FontSizes.TryGetValue(key, out var cached)) return cached;
        var probe = new XElement(W.r);
        _ = new XElement(W.p, new XElement(W.pPr, style), probe);
        // Resolve an ordinary unstyled run in the paragraph's style context. A paragraph
        // mark's direct rPr and a large neighboring run do not format the equation.
        var properties = FormattingAssembler.ResolveEffectiveRunProperties(context.Document, probe);
        var size = properties.Elements().Any(e => e.Name != W.rFonts && e.Name != W.sz
            && e.Name != W.szCs && e.Name != W.lang) ? null : (decimal?)properties.Element(W.sz)?.Attribute(W.val);
        return context.FontSizes[key] = size > 0 ? size / 2 : null;
    }

    // Project only complete expressions. Dropping an unknown operator or property would change
    // the formula, so an unsupported descendant keeps the whole equation on the diagnostic path.
    private static XElement? ProcessMath(XElement source, MathRenderingContext context)
    {
        if (!context.Supported || HasUnknownMathAttributes(source)) return null;
        var display = source.Name == M.oMathPara;
        if (display && (source.Elements().Count() != 1 || source.Element(M.oMath) == null)) return null;
        // OMML display fractions retain full-size nested operands by default; MathML shrinks
        // them. Keep that unmodeled display rule explicit rather than silently changing size.
        if (display && source.Descendants(M.f).Any(f => f.Ancestors(M.f).Any())) return null;
        var size = MathFontSize(source, context);
        if (size == null) return null;
        var expression = MathExpression(display ? source.Element(M.oMath)! : source);
        return expression == null || expression.Value.Length == 0 ? null : new XElement(MathMl + "math",
            new XAttribute("display", display ? "block" : "inline"),
            new XAttribute("style", "font-family: " + context.FontFamily + ";font-size: "
                + size.Value.ToString(CultureInfo.InvariantCulture) + "pt;"), expression);
    }

    private static XElement? MathExpression(XElement source)
    {
        if (HasUnknownMathAttributes(source)) return null;
        var name = source.Name;
        if (name == M.r)
        {
            if (source.Elements().Any(e => e.Name != M.t) || !source.HasElements) return null;
            var tokens = new List<XElement>();
            foreach (var text in source.Elements())
            {
                if (text.HasElements || HasUnknownMathAttributes(text, XNamespace.Xml + "space")) return null;
                foreach (Match part in Regex.Matches(text.Value, @"[0-9]+(?:\.[0-9]+)?|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\s\S]", RegexOptions.CultureInvariant))
                {
                    if (part.Length > 1 && char.IsDigit(part.Value[0]))
                    {
                        tokens.Add(new XElement(MathMl + "mn", part.Value));
                        continue;
                    }
                    if (!Rune.TryGetRuneAt(part.Value, 0, out var rune)) return null;
                    if (Rune.IsWhiteSpace(rune) || Rune.GetUnicodeCategory(rune) is UnicodeCategory.NonSpacingMark
                        or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format
                        or UnicodeCategory.Control or UnicodeCategory.Surrogate || rune.Value == '.') return null;
                    var token = new XElement(MathMl + (Rune.IsLetter(rune) ? "mi" : Rune.IsDigit(rune) ? "mn" : "mo"), rune.ToString());
                    if (token.Name.LocalName == "mo")
                        token.Add(new XAttribute("stretchy", "false"), new XAttribute("largeop", "false"));
                    tokens.Add(token);
                }
            }
            return new XElement(MathMl + "mrow", tokens);
        }
        if (name == M.oMath || name == M.num || name == M.den || name == M.e || name == M.deg)
        {
            var children = source.Elements().Select(MathExpression).ToArray();
            return children.Any(e => e == null) ? null : new XElement(MathMl + "mrow", children);
        }
        if (name == M.f)
        {
            // The default stacked fraction has a bar. Other fraction types retain diagnostics.
            if (source.Elements().Any(e => e.Name != M.num && e.Name != M.den)
                || source.Elements(M.num).Count() != 1 || source.Elements(M.den).Count() != 1) return null;
            var numerator = MathExpression(source.Element(M.num)!);
            var denominator = MathExpression(source.Element(M.den)!);
            return numerator == null || denominator == null ? null : new XElement(MathMl + "mfrac", numerator, denominator);
        }
        if (name == M.rad)
        {
            if (source.Elements().Any(e => e.Name != M.radPr && e.Name != M.deg && e.Name != M.e)
                || source.Elements(M.radPr).Count() > 1 || source.Elements(M.deg).Count() != 1
                || source.Elements(M.e).Count() != 1) return null;
            var properties = source.Element(M.radPr);
            if (properties != null && (HasUnknownMathAttributes(properties) || properties.Elements().Any(e => e.Name != M.degHide)
                || properties.Elements(M.degHide).Count() > 1)) return null;
            var hidden = properties?.Element(M.degHide);
            if (hidden != null && (hidden.HasElements || HasUnknownMathAttributes(hidden, M.val))) return null;
            var hide = hidden == null ? false : (string?)hidden.Attribute(M.val) switch
            {
                null or "1" or "true" or "on" => true,
                "0" or "false" or "off" => false,
                _ => (bool?)null,
            };
            if (hide == null) return null;
            var operand = MathExpression(source.Element(M.e)!);
            var degree = MathExpression(source.Element(M.deg)!);
            if (operand == null || degree == null) return null;
            return hide.Value ? new XElement(MathMl + "msqrt", operand)
                : new XElement(MathMl + "mroot", operand, degree);
        }
        return null;
    }
}
