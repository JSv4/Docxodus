using System;
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

        /// <summary>The document's declared <c>m:mathFont</c>, or null when it declares none.</summary>
        public string? FontName { get; init; }

        /// <summary>Where integral limits go when an <c>m:nary</c> omits <c>m:limLoc</c> (ECMA-376 §22.1.2.49).</summary>
        public string IntegralLimits { get; init; } = "subSup";

        /// <summary>Where other n-ary limits go when an <c>m:nary</c> omits <c>m:limLoc</c> (§22.1.2.71).</summary>
        public string NaryLimits { get; init; } = "undOvr";
    }

    /// <summary>One equation's projection state: the document context plus whether it is display math.</summary>
    private readonly record struct MathScope(MathRenderingContext Context, bool Display);

    private static bool HasUnknownMathAttributes(XElement element, XName? allowed = null) => element.Attributes()
        .Any(a => !a.IsNamespaceDeclaration && a.Name != PtOpenXml.Unid && a.Name != allowed);

    /// <summary>
    /// The document math settings (<c>m:mathPr</c>, ECMA-376 §22.1.2.62) the projection draws as Word does, each
    /// with the values it does so for. A setting that is not listed, or a listed one at another value, keeps
    /// every equation of the document on the placeholder. Word 2024 writes most of these into every document
    /// it saves, at the value the setting has when absent (issue #1710).
    /// </summary>
    private static readonly Dictionary<XName, Func<XElement, bool>> MathSettingsDrawn = new()
    {
        [M.mathFont] = setting => !string.IsNullOrWhiteSpace((string?)setting.Attribute(M.val)),
        // The limit placement an m:nary without m:limLoc takes; MathContext reads the value.
        [M.intLim] = setting => MathSettingIs(setting, "subSup", "undOvr"),
        [M.naryLim] = setting => MathSettingIs(setting, "subSup", "undOvr"),
        // Display equations are laid out as display math, between the text margins.
        [M.smallFrac] = setting => MathOnOff(setting, absent: false) == false,
        [M.dispDef] = setting => MathOnOff(setting, absent: true) == true,
        [M.lMargin] = setting => MathSettingIs(setting, "0"),
        [M.rMargin] = setting => MathSettingIs(setting, "0"),
        // Centred. A display paragraph is projected only when it holds one equation, and a group of one
        // centred as a group is that equation centred.
        [M.defJc] = setting => MathSettingIs(setting, "centerGroup", "center"),
        // How Word breaks an equation too long for its line, and where it sets the lines after the first.
        // The projection never breaks an equation, with these settings or without them, so none of their
        // values changes what it draws.
        [M.brkBin] = setting => MathSettingIs(setting, "before", "after", "repeat"),
        [M.brkBinSub] = setting => MathSettingIs(setting, "--", "-+", "+-"),
        [M.wrapIndent] = setting => uint.TryParse((string?)setting.Attribute(M.val), NumberStyles.None, CultureInfo.InvariantCulture, out _),
        [M.wrapRight] = setting => MathOnOff(setting, absent: false) != null,
    };

    private static bool MathSettingIs(XElement setting, params string[] values) =>
        values.Contains((string?)setting.Attribute(M.val));

    private static MathRenderingContext MathContext(WordprocessingDocument document)
    {
        var part = document.MainDocumentPart!;
        if (part.Annotation<MathRenderingContext>() is { } cached) return cached;
        var properties = part.DocumentSettingsPart?.GetXDocument().Root?.Element(M.mathPr);
        var font = (string?)properties?.Element(M.mathFont)?.Attribute(M.val);
        // Each setting once, carrying only its value, and drawn as Word draws it.
        var supported = properties == null || (!HasUnknownMathAttributes(properties)
            && properties.Elements().All(setting => properties.Elements(setting.Name).Count() == 1
                && !setting.HasElements && !HasUnknownMathAttributes(setting, M.val)
                && MathSettingsDrawn.TryGetValue(setting.Name, out var drawn) && drawn(setting)));
        // An absent OMML mathFont permits font substitution. Let the browser's math family and
        // the existing export font policy expose that choice; never silently use the body font.
        var family = font == null ? "math" : "'" + string.Concat(font.Select(c =>
            c is '\\' or '\'' || char.IsControl(c) ? "\\" + ((int)c).ToString("X", CultureInfo.InvariantCulture) + " " : c.ToString())) + "', math";
        var context = new MathRenderingContext(supported, family, document)
        {
            FontName = font,
            IntegralLimits = (string?)properties?.Element(M.intLim)?.Attribute(M.val) ?? "subSup",
            NaryLimits = (string?)properties?.Element(M.naryLim)?.Attribute(M.val) ?? "undOvr",
        };
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
        // Kerning and ligatures are text typography, which Word does not apply to a math zone; Word 2024
        // writes both into the run defaults of every new document.
        var size = properties.Elements().Any(e => e.Name != W.rFonts && e.Name != W.sz && e.Name != W.szCs
            && e.Name != W.lang && e.Name != W.kern && e.Name != W14.ligatures)
            ? null : (decimal?)properties.Element(W.sz)?.Attribute(W.val);
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
        // them, and so do MathML table cells. Keep that unmodeled display rule explicit rather
        // than silently changing the size of a nested fraction or large operator.
        if (display && source.Descendants().Any(e => (e.Name == M.f || e.Name == M.nary)
            && e.Ancestors().Any(a => a.Name == M.f || a.Name == M._m))) return null;
        var size = MathFontSize(source, context);
        if (size == null) return null;
        var expression = MathExpression(display ? source.Element(M.oMath)! : source, new MathScope(context, display));
        if (expression == null || expression.Value.Length == 0) return null;
        var math = new XElement(MathMl + "math",
            new XAttribute("display", display ? "block" : "inline"),
            new XAttribute("style", "font-family: " + context.FontFamily + ";font-size: "
                + size.Value.ToString(CultureInfo.InvariantCulture) + "pt;"), expression);
        return display ? math : new XElement(Xhtml.span, new XAttribute("data-docx-math", "true"), math);
    }

    private static XElement Mml(string name, params object?[] content) => new(MathMl + name, content);

    private static XElement MathOperator(string text, params (string Name, string Value)[] attributes) =>
        Mml("mo", attributes.Select(a => new XAttribute(a.Name, a.Value)), text);

    /// <summary>An OMML on/off property (ST_OnOff): absent takes <paramref name="absent"/>, a bare
    /// element is on, and anything other than the six spellings is unprojectable (null).</summary>
    private static bool? MathOnOff(XElement? property, bool absent)
    {
        if (property == null) return absent;
        if (property.HasElements || HasUnknownMathAttributes(property, M.val)) return null;
        return (string?)property.Attribute(M.val) switch
        {
            null or "1" or "true" or "on" => true,
            "0" or "false" or "off" => false,
            _ => null,
        };
    }

    /// <summary>The <c>m:val</c> of a property, <paramref name="absent"/> when the property is
    /// missing, and null when the property is present without a value.</summary>
    private static string? MathValue(XElement? properties, XName name, string absent) =>
        properties?.Element(name) is { } property ? (string?)property.Attribute(M.val) : absent;

    /// <summary>One printable character, as OMML's ST_Char properties carry. An empty value means no
    /// character; combining, control and format characters have no standalone glyph to draw.</summary>
    private static bool IsMathCharacter(string? value, bool allowEmpty) =>
        value != null && (value.Length == 0 ? allowEmpty
            : Rune.TryGetRuneAt(value, 0, out var rune) && rune.Utf16SequenceLength == value.Length
                && !Rune.IsWhiteSpace(rune) && Rune.GetUnicodeCategory(rune) is not (UnicodeCategory.NonSpacingMark
                    or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format
                    or UnicodeCategory.Control or UnicodeCategory.Surrogate or UnicodeCategory.PrivateUse));

    /// <summary>Whether an <c>m:*Pr</c> element holds only <paramref name="allowed"/> properties, each
    /// at most once and carrying only <c>m:val</c>. <c>m:ctrlPr</c> formats the structure's own glyphs
    /// and passes only when it changes nothing the projection draws. A matrix's
    /// <c>m:mcs</c> has children of its own, which the matrix validates.</summary>
    private static bool MathPropertiesValid(XElement? properties, MathScope scope, params XName[] allowed)
    {
        if (properties == null) return true;
        if (HasUnknownMathAttributes(properties)) return false;
        foreach (var property in properties.Elements())
        {
            if (properties.Elements(property.Name).Count() > 1 || !allowed.Contains(property.Name)) return false;
            if (property.Name == M.ctrlPr)
            {
                if (HasUnknownMathAttributes(property) || property.Elements().Any(e => e.Name != W.rPr)
                    || property.Elements().Count() > 1
                    || (property.Element(W.rPr) is { } format
                        && MathRunColor(format, MathControlCharacters(properties), scope, controlCharacter: true) == null)) return false;
            }
            else if (property.Name != M.mcs && (property.HasElements || HasUnknownMathAttributes(property, M.val))) return false;
        }
        return true;
    }

    /// <summary>The characters a structure's control properties format: its explicit operator, fence
    /// and separator characters plus every default one, so that each font slot they could use is checked.</summary>
    private static string MathControlCharacters(XElement properties) =>
        "()|/√∫ˆ⏟" + string.Concat(properties.Elements()
            .Where(e => e.Name == M.chr || e.Name == M.begChr || e.Name == M.sepChr || e.Name == M.endChr)
            .Select(e => (string?)e.Attribute(M.val)));

    /// <summary>The CSS color a math run's <c>w:rPr</c> paints, "" for none, or null when it carries
    /// formatting the projection does not draw. Fonts must name the document's math font, which the
    /// equation already uses, for every font slot <paramref name="text"/> draws with. A control
    /// character's properties may also carry <c>w:i</c>, which Word writes on them by default; OMML
    /// takes math italics from <c>m:sty</c>, not from it.</summary>
    private static string? MathRunColor(XElement properties, string text, MathScope scope, bool controlCharacter)
    {
        if (HasUnknownMathAttributes(properties)) return null;
        var color = "";
        foreach (var property in properties.Elements())
        {
            if (properties.Elements(property.Name).Count() > 1) return null;
            if (property.Name == W.rFonts)
            {
                if (!MathFontsMatch(property, text, scope)) return null;
            }
            else if (property.Name == W.i && controlCharacter)
            {
                if (property.HasElements || HasUnknownMathAttributes(property, W.val)) return null;
            }
            else if (property.Name == W.color && !controlCharacter)
            {
                var value = (string?)property.Attribute(W.val);
                if (property.HasElements || HasUnknownMathAttributes(property, W.val) || value == null) return null;
                if (value == "auto") continue;
                if (value.Length != 6 || !value.All(char.IsAsciiHexDigit)) return null;
                color = "#" + value.ToUpperInvariant();
            }
            else return null;
        }
        return color;
    }

    private static readonly (XName Font, XName Theme)[] MathFontSlots =
    {
        (W.ascii, W.asciiTheme), (W.hAnsi, W.hAnsiTheme), (W.eastAsia, W.eastAsiaTheme), (W.cs, W.cstheme),
    };

    /// <summary>Whether a math run's <c>w:rFonts</c> names the document math font in every slot its
    /// characters use. Word gives each character one of four font slots (Latin, high ANSI, East Asian,
    /// complex script; [MS-OI29500] §2.1.87), so a slot no character uses does not change the drawing,
    /// as with the themed East Asian font Word writes on Latin math runs. A theme font resolves through
    /// the theme's font scheme, and a slot the run leaves unset keeps the math font.</summary>
    private static bool MathFontsMatch(XElement fonts, string text, MathScope scope)
    {
        var name = scope.Context.FontName;
        if (name == null || fonts.HasElements || fonts.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != PtOpenXml.Unid
            && a.Name != W.hint && !MathFontSlots.Any(s => s.Font == a.Name || s.Theme == a.Name))) return false;
        var resolved = new XElement(W.rFonts, fonts.Attribute(W.hint));
        foreach (var (font, theme) in MathFontSlots)
        {
            // A theme font takes precedence over the slot's literal font (§17.3.2.26).
            var value = fonts.Attribute(theme) is { } themed ? MathThemeFont(scope.Context.Document, themed.Value) ?? "" : (string?)fonts.Attribute(font);
            if (value != null) resolved.SetAttributeValue(font, value);
        }
        var slots = new FormattingAssembler.CharStyleAttributes(fonts.Ancestors(W.p).FirstOrDefault()?.Element(W.pPr), new XElement(W.rPr, resolved));
        foreach (var rune in text.EnumerateRunes())
        {
            // The slot table covers the Basic Multilingual Plane; supplementary ideographs are East Asian.
            var slot = rune.Value is >= 0x20000 and <= 0x3FFFF ? FormattingAssembler.FontType.EastAsia
                : rune.IsBmp ? FormattingAssembler.DetermineFontTypeFromCharacter((char)rune.Value, slots) : FormattingAssembler.FontType.HAnsi;
            var font = slot switch
            {
                FormattingAssembler.FontType.Ascii => W.ascii,
                FormattingAssembler.FontType.HAnsi => W.hAnsi,
                FormattingAssembler.FontType.EastAsia => W.eastAsia,
                _ => W.cs,
            };
            if ((string?)resolved.Attribute(font) is { } face && face != name) return false;
        }
        return true;
    }

    /// <summary>The typeface a theme font reference (ST_Theme, such as <c>minorEastAsia</c>) names in
    /// the document theme, or null when the theme leaves it to a per-script font.</summary>
    private static string? MathThemeFont(WordprocessingDocument document, string theme)
    {
        var scheme = document.MainDocumentPart?.ThemePart?.GetXDocument().Root?.Element(A.themeElements)?.Element(A.fontScheme);
        var group = theme.StartsWith("major", System.StringComparison.Ordinal) ? scheme?.Element(A.majorFont)
            : theme.StartsWith("minor", System.StringComparison.Ordinal) ? scheme?.Element(A.minorFont) : null;
        var script = theme.Length > 5 ? theme[5..] switch
        {
            "Ascii" or "HAnsi" => A.latin,
            "EastAsia" => A.ea,
            "Bidi" => A.cs,
            _ => null,
        } : null;
        var face = script == null ? null : (string?)group?.Element(script)?.Attribute("typeface");
        return string.IsNullOrEmpty(face) ? null : face;
    }

    /// <summary>At most one properties element and exactly one of each named argument.</summary>
    private static bool MathShape(XElement source, XName properties, params XName[] arguments) =>
        source.Elements().All(e => e.Name == properties || arguments.Contains(e.Name))
        && source.Elements(properties).Count() <= 1
        && arguments.All(a => source.Elements(a).Count() == 1);

    private static readonly XName[] MathControlOnly = { M.ctrlPr };

    // Project only complete expressions. Dropping an unknown operator or property would change
    // the formula, so an unsupported descendant keeps the whole equation on the diagnostic path.
    private static XElement? MathExpression(XElement source, MathScope scope)
    {
        if (HasUnknownMathAttributes(source)) return null;
        var name = source.Name;
        if (name == M.r)
        {
            var tokens = MathRunTokens(source, scope, alignment: false);
            return tokens == null ? null : Mml("mrow", tokens);
        }
        if (name == M.oMath || name == M.num || name == M.den || name == M.e || name == M.deg
            || name == M.sub || name == M.sup || name == M.lim || name == M.fName)
        {
            var children = MathArgumentChildren(source, scope)?.Select(child => MathExpression(child, scope)).ToArray();
            return children == null || children.Any(e => e == null) ? null : Mml("mrow", children);
        }
        XElement? Arg(XName argument) => MathExpression(source.Element(argument)!, scope);
        if (name == M.f)
        {
            if (!MathShape(source, M.fPr, M.num, M.den)) return null;
            var properties = source.Element(M.fPr);
            if (!MathPropertiesValid(properties, scope, M.type, M.ctrlPr)) return null;
            var numerator = Arg(M.num);
            var denominator = Arg(M.den);
            if (numerator == null || denominator == null) return null;
            // A skewed fraction has no MathML Core form; it keeps the diagnostic.
            return MathValue(properties, M.type, "bar") switch
            {
                "bar" => Mml("mfrac", numerator, denominator),
                "noBar" => Mml("mfrac", new XAttribute("linethickness", "0"), numerator, denominator),
                // Word sets a linear fraction's slash tight against both operands.
                "lin" => Mml("mrow", numerator, MathOperator("/", ("stretchy", "false"), ("largeop", "false"),
                    ("lspace", "0"), ("rspace", "0")), denominator),
                _ => null,
            };
        }
        if (name == M.rad)
        {
            if (!MathShape(source, M.radPr, M.deg, M.e)) return null;
            var properties = source.Element(M.radPr);
            if (!MathPropertiesValid(properties, scope, M.degHide, M.ctrlPr)) return null;
            var hide = MathOnOff(properties?.Element(M.degHide), false);
            if (hide == null) return null;
            var operand = Arg(M.e);
            var degree = Arg(M.deg);
            if (operand == null || degree == null) return null;
            return hide.Value ? Mml("msqrt", operand) : Mml("mroot", operand, degree);
        }
        if (name == M.sSup || name == M.sSub || name == M.sSubSup || name == M.sPre)
        {
            var properties = name == M.sSup ? M.sSupPr : name == M.sSub ? M.sSubPr : name == M.sSubSup ? M.sSubSupPr : M.sPrePr;
            var arguments = name == M.sSup ? new[] { M.e, M.sup } : name == M.sSub ? new[] { M.e, M.sub } : new[] { M.e, M.sub, M.sup };
            if (!MathShape(source, properties, arguments)) return null;
            var pr = source.Element(properties);
            // Aligned scripts (m:alnScr) stack sub- and superscript flush; only the default is drawn.
            if (!MathPropertiesValid(pr, scope, name == M.sSubSup ? new[] { M.alnScr, M.ctrlPr } : MathControlOnly)
                || MathOnOff(pr?.Element(M.alnScr), false) != false) return null;
            var parts = arguments.Select(Arg).ToArray();
            if (parts.Any(p => p == null)) return null;
            return name == M.sSup ? Mml("msup", parts) : name == M.sSub ? Mml("msub", parts)
                : name == M.sSubSup ? Mml("msubsup", parts)
                : Mml("mmultiscripts", parts[0], Mml("mprescripts"), parts[1], parts[2]);
        }
        if (name == M.nary) return MathNary(source, scope);
        if (name == M.d) return MathDelimiter(source, scope);
        if (name == M._m) return MathMatrix(source, scope);
        if (name == M.eqArr) return MathEquationArray(source, scope);
        if (name == M.func)
        {
            if (!MathShape(source, M.funcPr, M.fName, M.e) || !MathPropertiesValid(source.Element(M.funcPr), scope, MathControlOnly)) return null;
            var function = Arg(M.fName);
            var argument = Arg(M.e);
            // Word sets a thin space between a function name and its argument. U+2061 would carry
            // the same meaning but draws nothing: MathML's dictionary gives it no spacing.
            return function == null || argument == null ? null
                : Mml("mrow", function, Mml("mspace", new XAttribute("width", "0.1667em")), argument);
        }
        if (name == M.acc)
        {
            if (!MathShape(source, M.accPr, M.e)) return null;
            var properties = source.Element(M.accPr);
            if (!MathPropertiesValid(properties, scope, M.chr, M.ctrlPr)) return null;
            var accent = MathAccent(MathValue(properties, M.chr, "\u0302"));
            var operand = Arg(M.e);
            return accent == null || operand == null ? null
                : Mml("mover", new XAttribute("accent", "true"), operand, MathOperator(accent));
        }
        if (name == M.bar)
        {
            if (!MathShape(source, M.barPr, M.e)) return null;
            var properties = source.Element(M.barPr);
            if (!MathPropertiesValid(properties, scope, M.pos, M.ctrlPr)) return null;
            var operand = Arg(M.e);
            // A rule along the operand's edge, like a border box's side. A stretchy overline or low line
            // would need the math font to carry stretch variants for it; Cambria Math has none, and
            // the bar then stays one character wide.
            return operand == null ? null : MathValue(properties, M.pos, "bot") switch
            {
                "top" => Mml("mrow", new XAttribute("style", "padding-top: 0.1em;border-top: 0.06em solid currentColor;"), operand),
                "bot" => Mml("mrow", new XAttribute("style", "padding-bottom: 0.1em;border-bottom: 0.06em solid currentColor;"), operand),
                _ => null,
            };
        }
        if (name == M.limLow || name == M.limUpp)
        {
            var properties = name == M.limLow ? M.limLowPr : M.limUppPr;
            if (!MathShape(source, properties, M.e, M.lim) || !MathPropertiesValid(source.Element(properties), scope, MathControlOnly)) return null;
            var operand = Arg(M.e);
            var limit = Arg(M.lim);
            return operand == null || limit == null ? null : Mml(name == M.limLow ? "munder" : "mover", operand, limit);
        }
        if (name == M.groupChr)
        {
            if (!MathShape(source, M.groupChrPr, M.e)) return null;
            var properties = source.Element(M.groupChrPr);
            if (!MathPropertiesValid(properties, scope, M.chr, M.pos, M.vertJc, M.ctrlPr)) return null;
            var character = MathValue(properties, M.chr, "\u23DF");
            var position = MathValue(properties, M.pos, "bot");
            // ECMA-376 §22.1.2.119: an absent m:vertJc is top and a valueless one is bot. MathML keeps
            // the operand on the baseline, which is the layout of a bottom group justified to the top
            // and of a top group justified to the bottom: Word's underbrace and overbrace.
            var justification = properties?.Element(M.vertJc) is { } jc ? (string?)jc.Attribute(M.val) ?? "bot" : "top";
            var operand = Arg(M.e);
            if (operand == null || !IsMathCharacter(character, allowEmpty: false)) return null;
            return (position, justification) switch
            {
                ("bot", "top") => Mml("munder", operand, MathOperator(character!, ("stretchy", "true"))),
                ("top", "bot") => Mml("mover", operand, MathOperator(character!, ("stretchy", "true"))),
                _ => null,
            };
        }
        if (name == M.box)
        {
            // Box properties other than the no-break flag change operator spacing or line breaking.
            if (!MathShape(source, M.boxPr, M.e)) return null;
            var properties = source.Element(M.boxPr);
            if (!MathPropertiesValid(properties, scope, M.noBreak, M.ctrlPr) || MathOnOff(properties?.Element(M.noBreak), true) == null) return null;
            return Arg(M.e);
        }
        if (name == M.borderBox)
        {
            if (!MathShape(source, M.borderBoxPr, M.e)) return null;
            var properties = source.Element(M.borderBoxPr);
            var sides = new[] { M.hideTop, M.hideRight, M.hideBot, M.hideLeft };
            var strikes = new[] { M.strikeH, M.strikeV, M.strikeBLTR, M.strikeTLBR };
            if (!MathPropertiesValid(properties, scope, sides.Concat(strikes).Append(M.ctrlPr).ToArray())) return null;
            // MathML Core has no menclose, so strike-through lines keep the diagnostic.
            if (strikes.Any(s => MathOnOff(properties?.Element(s), false) != false)) return null;
            var hidden = sides.Select(s => MathOnOff(properties?.Element(s), false)).ToArray();
            if (hidden.Any(h => h == null)) return null;
            var operand = Arg(M.e);
            if (operand == null) return null;
            var style = "padding: 0.1em;border: 0.06em solid currentColor;" + string.Concat(
                new[] { "top", "right", "bottom", "left" }.Where((_, i) => hidden[i] == true).Select(side => "border-" + side + "-style: none;"));
            return Mml("mrow", new XAttribute("style", style), operand);
        }
        return null;
    }

    /// <summary>The content of an OMML argument (CT_OMathArg). Its trailing <c>m:ctrlPr</c> formats
    /// only an empty argument's placeholder; <c>m:argPr</c> changes the argument's size and is not drawn.</summary>
    private static IEnumerable<XElement>? MathArgumentChildren(XElement argument, MathScope scope)
    {
        var control = argument.Elements(M.ctrlPr).ToArray();
        if (control.Length > 1 || (control.Length == 1 && (argument.Elements().Last() != control[0]
            || !MathPropertiesValid(new XElement(M.argPr, control[0]), scope, MathControlOnly)))) return null;
        return argument.Elements().Where(e => e.Name != M.ctrlPr);
    }

    private const string IntegralOperators = "\u222B\u222C\u222D\u222E\u222F\u2230\u2231\u2232\u2233\u2A0C";

    private static XElement? MathNary(XElement source, MathScope scope)
    {
        if (!MathShape(source, M.naryPr, M.sub, M.sup, M.e)) return null;
        var properties = source.Element(M.naryPr);
        if (!MathPropertiesValid(properties, scope, M.chr, M.limLoc, M.grow, M.subHide, M.supHide, M.ctrlPr)) return null;
        var character = MathValue(properties, M.chr, "\u222B");
        if (!IsMathCharacter(character, allowEmpty: false)) return null;
        var location = MathValue(properties, M.limLoc,
            IntegralOperators.Contains(character!, System.StringComparison.Ordinal) ? scope.Context.IntegralLimits : scope.Context.NaryLimits);
        var grow = MathOnOff(properties?.Element(M.grow), false);
        var subHide = MathOnOff(properties?.Element(M.subHide), false);
        var supHide = MathOnOff(properties?.Element(M.supHide), false);
        if (location is not ("undOvr" or "subSup") || grow == null || subHide == null || supHide == null) return null;
        var under = location == "undOvr";
        var operand = MathExpression(source.Element(M.e)!, scope);
        // A hidden limit is not drawn, so its content needs no projection.
        var lower = subHide.Value ? null : MathExpression(source.Element(M.sub)!, scope);
        var upper = supHide.Value ? null : MathExpression(source.Element(M.sup)!, scope);
        if (operand == null || (!subHide.Value && lower == null) || (!supHide.Value && upper == null)) return null;
        // Limits placed under and over move beside the operator in inline (compact) math, as in
        // Word's inline n-ary operators; limits placed beside never move.
        var op = MathOperator(character!, ("largeop", "true"), ("movablelimits", under ? "true" : "false"),
            ("stretchy", grow.Value ? "true" : "false"));
        var limited = (lower, upper) switch
        {
            (null, null) => op,
            (not null, null) => Mml(under ? "munder" : "msub", op, lower),
            (null, not null) => Mml(under ? "mover" : "msup", op, upper),
            _ => Mml(under ? "munderover" : "msubsup", op, lower, upper),
        };
        return Mml("mrow", limited, operand);
    }

    private static XElement? MathDelimiter(XElement source, MathScope scope)
    {
        if (source.Elements().Any(e => e.Name != M.dPr && e.Name != M.e) || source.Elements(M.dPr).Count() > 1
            || !source.Elements(M.e).Any()) return null;
        var properties = source.Element(M.dPr);
        if (!MathPropertiesValid(properties, scope, M.begChr, M.sepChr, M.endChr, M.grow, M.shp, M.ctrlPr)) return null;
        var begin = MathValue(properties, M.begChr, "(");
        var separator = MathValue(properties, M.sepChr, "|");
        var end = MathValue(properties, M.endChr, ")");
        var grow = MathOnOff(properties?.Element(M.grow), true);
        // shp=centered keeps a grown delimiter symmetric about the math axis; match follows the content.
        var shape = MathValue(properties, M.shp, "centered");
        if (!IsMathCharacter(begin, true) || !IsMathCharacter(separator, true) || !IsMathCharacter(end, true)
            || grow == null || shape is not ("centered" or "match")) return null;
        (string, string)[] Fence(string form, string role) => new[]
        {
            (role, "true"), ("form", form), ("stretchy", grow.Value ? "true" : "false"), ("symmetric", shape == "centered" ? "true" : "false"),
        };
        var content = new List<XElement>();
        if (begin!.Length > 0) content.Add(MathOperator(begin, Fence("prefix", "fence")));
        var first = true;
        foreach (var operand in source.Elements(M.e))
        {
            if (!first && separator!.Length > 0) content.Add(MathOperator(separator, Fence("infix", "separator")));
            first = false;
            var projected = MathExpression(operand, scope);
            if (projected == null) return null;
            content.Add(projected);
        }
        if (end!.Length > 0) content.Add(MathOperator(end, Fence("postfix", "fence")));
        return Mml("mrow", content);
    }

    private static XElement? MathMatrix(XElement source, MathScope scope)
    {
        if (source.Elements().Any(e => e.Name != M.mPr && e.Name != M.mr) || source.Elements(M.mPr).Count() > 1) return null;
        var properties = source.Element(M.mPr);
        // Placeholder hiding only affects Word's editing aids for empty cells. Spacing and base
        // justification properties change the layout and keep the diagnostic.
        if (!MathPropertiesValid(properties, scope, M.mcs, M.plcHide, M.ctrlPr)
            || MathOnOff(properties?.Element(M.plcHide), false) == null) return null;
        var alignments = new List<string>();
        if (properties?.Element(M.mcs) is { } columns)
        {
            if (HasUnknownMathAttributes(columns) || !columns.HasElements || columns.Elements().Any(e => e.Name != M.mc)) return null;
            foreach (var column in columns.Elements())
            {
                if (HasUnknownMathAttributes(column) || column.Elements().Count() != 1 || column.Element(M.mcPr) is not { } pr
                    || !MathPropertiesValid(pr, scope, M.count, M.mcJc)) return null;
                var count = MathValue(pr, M.count, "1");
                var justification = MathValue(pr, M.mcJc, "center");
                if (!int.TryParse(count, NumberStyles.None, CultureInfo.InvariantCulture, out var n) || n is < 1 or > MaxMathColumns
                    || justification is not ("center" or "left" or "right")) return null;
                alignments.AddRange(Enumerable.Repeat(justification, n));
            }
        }
        var rows = source.Elements(M.mr).ToArray();
        if (rows.Length == 0 || rows.Any(r => HasUnknownMathAttributes(r) || !r.HasElements || r.Elements().Any(e => e.Name != M.e))
            || rows.Select(r => r.Elements().Count()).Distinct().Count() != 1) return null;
        var table = Mml("mtable");
        foreach (var row in rows)
        {
            var line = Mml("mtr");
            var column = 0;
            foreach (var cell in row.Elements())
            {
                var projected = MathExpression(cell, scope);
                if (projected == null) return null;
                var alignment = column < alignments.Count ? alignments[column] : "center";
                column++;
                line.Add(Mml("mtd", alignment == "center" ? null : new XAttribute("style", MathCellAlignment(alignment)), projected));
            }
            table.Add(line);
        }
        return table;
    }

    /// <summary>The CSS that sets a table cell's math to one side. Chromium moves a cell's math only for
    /// the <c>-webkit-</c> keyword (its own centring uses <c>-webkit-center</c>); an engine without that
    /// keyword drops the second declaration and keeps the standard one.</summary>
    private static string MathCellAlignment(string side) => "text-align: " + side + ";text-align: -webkit-" + side + ";";

    private static readonly XName MathAlignmentPoint = "alignment-point";

    /// <summary>The most columns a matrix column group or an equation-array row may span. Both tables
    /// repeat a per-column cost on every row, so an unbounded count would square a small input.</summary>
    private const int MaxMathColumns = 64;

    private static XElement? MathEquationArray(XElement source, MathScope scope)
    {
        // Row spacing, alignment and base justification properties change layout and keep the diagnostic.
        if (source.Elements().Any(e => e.Name != M.eqArrPr && e.Name != M.e) || source.Elements(M.eqArrPr).Count() > 1
            || !source.Elements(M.e).Any() || !MathPropertiesValid(source.Element(M.eqArrPr), scope, MathControlOnly)) return null;
        var rows = new List<List<List<XElement>>>();
        foreach (var row in source.Elements(M.e))
        {
            if (HasUnknownMathAttributes(row)) return null;
            var children = MathArgumentChildren(row, scope);
            if (children == null) return null;
            // An ampersand in a row's own runs marks an alignment point rather than drawing a glyph.
            var cells = new List<List<XElement>> { new() };
            foreach (var child in children)
            {
                var parts = child.Name == M.r ? MathRunTokens(child, scope, alignment: true)
                    : MathExpression(child, scope) is { } projected ? new List<XElement> { projected } : null;
                if (parts == null) return null;
                foreach (var part in parts)
                {
                    if (part.Name != MathAlignmentPoint) cells[^1].Add(part);
                    else if (cells.Count == MaxMathColumns) return null;
                    else cells.Add(new List<XElement>());
                }
            }
            rows.Add(cells);
        }
        var width = rows.Max(r => r.Count);
        // Alignment points pair up as right- then left-aligned columns, so the rows' points line up.
        // Without one, each row is centered.
        var table = Mml("mtable", scope.Display ? new XAttribute("displaystyle", "true") : null);
        foreach (var row in rows)
        {
            table.Add(Mml("mtr", Enumerable.Range(0, width).Select(i => Mml("mtd",
                width == 1 ? null : new XAttribute("style", MathCellAlignment(i % 2 == 0 ? "right" : "left")
                    + "padding-left: 0;padding-right: 0;"),
                Mml("mrow", i < row.Count ? row[i] : null)))));
        }
        return table;
    }

    /// <summary>A math run's text as MathML tokens, styled by its <c>m:rPr</c> and <c>w:rPr</c>. With
    /// <paramref name="alignment"/>, an ampersand becomes an alignment-point marker.</summary>
    private static List<XElement>? MathRunTokens(XElement source, MathScope scope, bool alignment)
    {
        if (!source.Elements(M.t).Any() || source.Elements().Any(e => e.Name != M.t && e.Name != M.rPr && e.Name != W.rPr)
            || source.Elements(M.rPr).Count() > 1 || source.Elements(W.rPr).Count() > 1) return null;
        var mathProperties = source.Element(M.rPr);
        // m:nor (ordinary text), m:brk and m:aln are not drawn and keep the diagnostic.
        if (!MathPropertiesValid(mathProperties, scope, M.sty, M.scr, M.lit)) return null;
        var color = source.Element(W.rPr) is { } format
            ? MathRunColor(format, string.Concat(source.Elements(M.t).Select(t => t.Value)), scope, controlCharacter: false) : "";
        var literal = MathOnOff(mathProperties?.Element(M.lit), false);
        var variant = MathVariant(MathValue(mathProperties, M.scr, "roman"), MathValue(mathProperties, M.sty, ""));
        if (color == null || literal == null || variant == null) return null;
        var tokens = new List<XElement>();
        StringBuilder? upright = null;
        void EndIdentifier()
        {
            if (upright != null) tokens.Add(Mml("mi", upright.ToString()));
            upright = null;
        }
        foreach (var text in source.Elements(M.t))
        {
            if (text.HasElements || HasUnknownMathAttributes(text, XNamespace.Xml + "space")) return null;
            foreach (Match part in Regex.Matches(text.Value, @"[0-9]+(?:\.[0-9]+)?|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\s\S]", RegexOptions.CultureInvariant))
            {
                if (part.Length > 1 && char.IsDigit(part.Value[0]))
                {
                    EndIdentifier();
                    tokens.Add(Mml("mn", MathStyledDigits(part.Value, variant)));
                    continue;
                }
                if (!Rune.TryGetRuneAt(part.Value, 0, out var rune)) return null;
                if (Rune.IsWhiteSpace(rune) || Rune.GetUnicodeCategory(rune) is UnicodeCategory.NonSpacingMark
                    or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format
                    or UnicodeCategory.Control or UnicodeCategory.Surrogate || rune.Value == '.') return null;
                if (alignment && rune.Value == '&')
                {
                    // A literal run's ampersand has no settled meaning in an equation array.
                    if (literal.Value) return null;
                    EndIdentifier();
                    tokens.Add(new XElement(MathAlignmentPoint));
                    continue;
                }
                if (Rune.IsLetter(rune) && variant == "upright")
                {
                    // A run's adjacent upright letters read as one identifier, like a function name.
                    (upright ??= new StringBuilder()).Append(rune.ToString());
                    continue;
                }
                EndIdentifier();
                if (Rune.IsLetter(rune))
                {
                    var styled = variant == "italic" ? rune.ToString() : MathStyledLetter(rune, variant);
                    if (styled == null) return null;
                    tokens.Add(Mml("mi", styled));
                }
                else if (Rune.IsDigit(rune))
                {
                    // The styled alphabets hold only 0-9; another script's digit would lose the style.
                    var digit = MathStyledDigits(rune.ToString(), variant);
                    if (digit == null) return null;
                    tokens.Add(Mml("mn", digit));
                }
                // OMML's bold styles have no drawable projection for operators and punctuation here.
                else if (variant.Contains("bold", System.StringComparison.Ordinal)) return null;
                else tokens.Add(MathOperator(rune.ToString(), ("stretchy", "false"), ("largeop", "false")));
            }
        }
        EndIdentifier();
        foreach (var token in tokens.Where(t => t.Name.Namespace == MathMl))
        {
            // A one-letter identifier is italic by default in MathML; longer ones are already upright.
            if (variant == "upright" && token.Name == MathMl + "mi" && token.Value.EnumerateRunes().Count() == 1)
                token.Add(new XAttribute("mathvariant", "normal"));
            if (color.Length > 0) token.Add(new XAttribute("style", "color: " + color + ";"));
        }
        return tokens;
    }

    /// <summary>The Unicode mathematical alphabet that an OMML script (<c>m:scr</c>) and style
    /// (<c>m:sty</c>) select for letters. "italic" is MathML's own default for a single letter;
    /// "upright" is plain text. Bold double-struck and monospace have no Unicode alphabet.</summary>
    private static string? MathVariant(string? script, string? style) => (script, style) switch
    {
        ("roman", "" or "i") => "italic",
        ("roman", "p") => "upright",
        ("roman", "b") => "bold",
        ("roman", "bi") => "bold-italic",
        ("script", "" or "p" or "i") => "script",
        ("script", "b" or "bi") => "bold-script",
        ("fraktur", "" or "p" or "i") => "fraktur",
        ("fraktur", "b" or "bi") => "bold-fraktur",
        ("double-struck", "" or "p" or "i") => "double-struck",
        ("sans-serif", "p") => "sans-serif",
        ("sans-serif", "" or "i") => "sans-serif-italic",
        ("sans-serif", "b") => "sans-serif-bold",
        ("sans-serif", "bi") => "sans-serif-bold-italic",
        ("monospace", "" or "p" or "i") => "monospace",
        _ => null,
    };

    // Unicode's Mathematical Alphanumeric Symbols block: the first capital Latin letter of each
    // alphabet, its first capital Greek letter where it has Greek, and its first digit (0 for none).
    private static readonly Dictionary<string, (int Latin, int Greek, int Digit)> MathAlphabets = new()
    {
        ["bold"] = (0x1D400, 0x1D6A8, 0x1D7CE),
        ["bold-italic"] = (0x1D468, 0x1D71C, 0x1D7CE),
        ["script"] = (0x1D49C, 0, 0),
        ["bold-script"] = (0x1D4D0, 0, 0x1D7CE),
        ["fraktur"] = (0x1D504, 0, 0),
        ["double-struck"] = (0x1D538, 0, 0x1D7D8),
        ["bold-fraktur"] = (0x1D56C, 0, 0x1D7CE),
        ["sans-serif"] = (0x1D5A0, 0, 0x1D7E2),
        ["sans-serif-italic"] = (0x1D608, 0, 0x1D7E2),
        ["sans-serif-bold"] = (0x1D5D4, 0x1D756, 0x1D7EC),
        ["sans-serif-bold-italic"] = (0x1D63C, 0x1D790, 0x1D7EC),
        ["monospace"] = (0x1D670, 0, 0x1D7F6),
    };

    // Letters Unicode had already encoded in Letterlike Symbols leave holes in those alphabets.
    private static readonly Dictionary<(string, char), int> MathAlphabetHoles = new()
    {
        [("script", 'B')] = 0x212C, [("script", 'E')] = 0x2130, [("script", 'F')] = 0x2131, [("script", 'H')] = 0x210B,
        [("script", 'I')] = 0x2110, [("script", 'L')] = 0x2112, [("script", 'M')] = 0x2133, [("script", 'R')] = 0x211B,
        [("script", 'e')] = 0x212F, [("script", 'g')] = 0x210A, [("script", 'o')] = 0x2134,
        [("fraktur", 'C')] = 0x212D, [("fraktur", 'H')] = 0x210C, [("fraktur", 'I')] = 0x2111, [("fraktur", 'R')] = 0x211C,
        [("fraktur", 'Z')] = 0x2128,
        [("double-struck", 'C')] = 0x2102, [("double-struck", 'H')] = 0x210D, [("double-struck", 'N')] = 0x2115,
        [("double-struck", 'P')] = 0x2119, [("double-struck", 'Q')] = 0x211A, [("double-struck", 'R')] = 0x211D,
        [("double-struck", 'Z')] = 0x2124,
    };

    private static string? MathStyledLetter(Rune letter, string variant)
    {
        var alphabet = MathAlphabets[variant];
        var value = letter.Value;
        int? point = value switch
        {
            (>= 'A' and <= 'Z') or (>= 'a' and <= 'z') => MathAlphabetHoles.TryGetValue((variant, (char)value), out var hole) ? hole
                : alphabet.Latin + (value <= 'Z' ? value - 'A' : 26 + value - 'a'),
            // Each Greek alphabet runs Α–Ω (ϴ fills the gap at U+03A2), ∇, then α–ω.
            >= 0x391 and <= 0x3A9 and not 0x3A2 when alphabet.Greek != 0 => alphabet.Greek + value - 0x391,
            >= 0x3B1 and <= 0x3C9 when alphabet.Greek != 0 => alphabet.Greek + 26 + value - 0x3B1,
            _ => null,
        };
        return point == null ? null : char.ConvertFromUtf32(point.Value);
    }

    /// <summary>The digits in the variant's alphabet, a decimal point kept as is; null when the
    /// alphabet styles digits and one of them is outside 0-9.</summary>
    private static string? MathStyledDigits(string digits, string variant)
    {
        var start = MathAlphabets.TryGetValue(variant, out var alphabet) ? alphabet.Digit : 0;
        if (start == 0) return digits;
        return digits.All(c => c is '.' or (>= '0' and <= '9'))
            ? string.Concat(digits.Select(c => c == '.' ? "." : char.ConvertFromUtf32(start + c - '0'))) : null;
    }

    /// <summary>The spacing form MathML draws over an operand for an OMML accent character. Word
    /// stores combining marks, which have no standalone glyph; other characters draw as themselves.</summary>
    private static string? MathAccent(string? character) => character switch
    {
        "\u0300" => "`",
        "\u0301" => "\u00B4",
        "\u0302" => "\u02C6",
        "\u0303" => "\u02DC",
        "\u0304" or "\u0305" => "\u00AF",
        "\u0306" => "\u02D8",
        "\u0307" => "\u02D9",
        "\u0308" => "\u00A8",
        "\u030A" => "\u02DA",
        "\u030C" => "\u02C7",
        "\u20D0" => "\u21BC",
        "\u20D1" => "\u21C0",
        "\u20D6" => "\u2190",
        "\u20D7" => "\u2192",
        "\u20E1" => "\u2194",
        _ => IsMathCharacter(character, allowEmpty: false) ? character : null,
    };
}

