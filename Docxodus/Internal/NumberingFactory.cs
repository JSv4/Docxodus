#nullable enable

using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Synthesizes reusable numbering definitions (bullet, decimal, letter, roman — plain or
/// parenthesized) so a plain paragraph can be promoted to a real list item.
/// <see cref="DocxSession.ApplyListFormat"/> / <see cref="DocxSession.ApplyListFormatRange"/>
/// use this when no suitable numbering exists. Definitions are tagged with a fixed marker
/// <c>w:nsid</c> per format and resolved find-or-create, so the op is idempotent across calls,
/// save/reopen, and undo/redo. Session snapshots cover both the numbering part and paragraph
/// <c>w:numPr</c> references.
/// </summary>
internal static class NumberingFactory
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Stable per-format marker (8-hex <c>w:nsid</c> value) used to find-or-create our own
    /// definition. Saved documents already carry these — a value, once shipped, is frozen.
    /// </summary>
    private static string NsidFor(ListFormat fmt) => fmt switch
    {
        ListFormat.Bullet => "0D0CB001",
        ListFormat.Decimal => "0D0CD001",
        ListFormat.DecimalParenthesis => "0D0CD002",
        ListFormat.LowerLetter => "0D0CA001",
        ListFormat.LowerLetterParenthesis => "0D0CA002",
        ListFormat.UpperLetter => "0D0CA101",
        ListFormat.UpperLetterParenthesis => "0D0CA102",
        ListFormat.LowerRoman => "0D0CC001",
        ListFormat.LowerRomanParenthesis => "0D0CC002",
        ListFormat.UpperRoman => "0D0CC101",
        ListFormat.UpperRomanParenthesis => "0D0CC102",
        _ => throw new ArgumentOutOfRangeException(nameof(fmt), fmt, "no numbering definition for this format"),
    };

    internal static bool IsDocxodusDefinition(XElement abstractNum)
    {
        var nsid = (string?)abstractNum.Element(W + "nsid")?.Attribute(W + "val");
        return Enum.GetValues<ListFormat>()
            .Where(format => format != ListFormat.None)
            .Any(format => string.Equals(nsid, NsidFor(format), StringComparison.OrdinalIgnoreCase));
    }

    // Standard Word bullet cycle (•, o, ▪) for synthesized nested levels — same glyph/font set
    // BuildAbstractNum emits for our own multi-level lists, so source and synthesized lists nest
    // identically.
    private static readonly string[] SynthBulletGlyphs = { "•", "o", "▪" };
    private static readonly string[] SynthBulletFonts = { "Symbol", "Courier New", "Wingdings" };

    /// <summary>
    /// Ensure a numbering definition for <paramref name="fmt"/> exists and return a numId
    /// pointing at it. <see cref="ListFormat.None"/> is not a numbering definition and throws.
    /// </summary>
    public static int EnsureNumbering(WordprocessingDocument doc, ListFormat fmt)
    {
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("no MainDocumentPart");
        var part = main.NumberingDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<NumberingDefinitionsPart>();
            part.PutXDocument(new XDocument(
                new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName))));
        }

        var root = part.GetXDocument().Root!;
        string nsid = NsidFor(fmt);

        // Find our previously-synthesized abstractNum (by marker nsid), or build one.
        var abstractNum = root.Elements(W + "abstractNum")
            .FirstOrDefault(a => (string?)a.Element(W + "nsid")?.Attribute(W + "val") == nsid);
        if (abstractNum is null)
        {
            int absId = NextId(root, "abstractNum", "abstractNumId");
            abstractNum = BuildAbstractNum(fmt, absId, nsid);
            WordprocessingMLUtil.InsertNumberingChildInOrder(root, abstractNum);
        }

        var abstractId = (string)abstractNum.Attribute(W + "abstractNumId")!;

        // Reuse an existing w:num pointing at our abstractNum, or create one.
        var num = root.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Element(W + "abstractNumId")?.Attribute(W + "val") == abstractId);
        if (num is null)
        {
            int numId = NextId(root, "num", "numId");
            num = new XElement(W + "num",
                new XAttribute(W + "numId", numId),
                new XElement(W + "abstractNumId", new XAttribute(W + "val", abstractId)));
            WordprocessingMLUtil.InsertNumberingChildInOrder(root, num);
        }

        // Flush the numbering part to its stream — the session's Save only persists the
        // projected parts (body/headers/...), not the numbering part we just mutated.
        part.PutXDocument();
        return (int)num.Attribute(W + "numId")!;
    }

    private static int NextId(XElement root, string elemLocalName, string idAttrLocalName)
    {
        int max = 0;
        foreach (var e in root.Elements(W + elemLocalName))
        {
            if (int.TryParse((string?)e.Attribute(W + idAttrLocalName), out var v))
                max = Math.Max(max, v);
        }
        return max + 1;
    }

    /// <summary>The level text for a numbered level: <c>(%N)</c> when parenthesized, else <c>%N.</c></summary>
    private static string NumberedLvlText(bool paren, int lvl) =>
        paren ? $"(%{lvl + 1})" : $"%{lvl + 1}.";

    /// <summary>Build a spec-valid 9-level abstractNum for <paramref name="fmt"/>.</summary>
    private static XElement BuildAbstractNum(ListFormat fmt, int absId, string nsid)
    {
        var (baseFormat, paren) = NumberFormats.FromListFormat(fmt);
        bool bullet = baseFormat == NumberFormat.Bullet;
        string numFmtToken = NumberFormats.ToOoxml(baseFormat);

        var an = new XElement(W + "abstractNum",
            new XAttribute(W + "abstractNumId", absId),
            new XElement(W + "nsid", new XAttribute(W + "val", nsid)),
            new XElement(W + "multiLevelType", new XAttribute(W + "val", "hybridMultilevel")));

        // Bullet glyphs cycle (•, o, ▪) using Symbol / Courier New / Wingdings, like Word.
        var bulletGlyphs = new[] { "", "o", "" };
        var bulletFonts = new[] { "Symbol", "Courier New", "Wingdings" };

        for (int lvl = 0; lvl < 9; lvl++)
        {
            int indentLeft = 720 * (lvl + 1);
            var pPr = new XElement(W + "pPr",
                new XElement(W + "ind",
                    new XAttribute(W + "left", indentLeft),
                    new XAttribute(W + "hanging", 360)));

            XElement lvl_;
            if (bullet)
            {
                lvl_ = new XElement(W + "lvl",
                    new XAttribute(W + "ilvl", lvl),
                    new XElement(W + "start", new XAttribute(W + "val", 1)),
                    new XElement(W + "numFmt", new XAttribute(W + "val", "bullet")),
                    new XElement(W + "lvlText", new XAttribute(W + "val", bulletGlyphs[lvl % 3])),
                    new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                    pPr,
                    new XElement(W + "rPr",
                        new XElement(W + "rFonts",
                            new XAttribute(W + "ascii", bulletFonts[lvl % 3]),
                            new XAttribute(W + "hAnsi", bulletFonts[lvl % 3]),
                            new XAttribute(W + "hint", "default"))));
            }
            else
            {
                lvl_ = new XElement(W + "lvl",
                    new XAttribute(W + "ilvl", lvl),
                    new XElement(W + "start", new XAttribute(W + "val", 1)),
                    new XElement(W + "numFmt", new XAttribute(W + "val", numFmtToken)),
                    new XElement(W + "lvlText", new XAttribute(W + "val", NumberedLvlText(paren, lvl))),
                    new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                    pPr);
            }
            an.Add(lvl_);
        }

        return an;
    }

    /// <summary>Build one spec-valid <c>w:lvl</c> at level <paramref name="lvl"/>. A
    /// <paramref name="numFmtToken"/> of <c>bullet</c> uses the glyph/font pair; any other token
    /// gets a numbered level text (parenthesized when <paramref name="paren"/>).</summary>
    private static XElement BuildLevel(string numFmtToken, bool paren, int lvl, string glyph, string font)
    {
        var pPr = new XElement(W + "pPr",
            new XElement(W + "ind",
                new XAttribute(W + "left", 720 * (lvl + 1)),
                new XAttribute(W + "hanging", 360)));

        if (numFmtToken == "bullet")
            return new XElement(W + "lvl",
                new XAttribute(W + "ilvl", lvl),
                new XElement(W + "start", new XAttribute(W + "val", 1)),
                new XElement(W + "numFmt", new XAttribute(W + "val", "bullet")),
                new XElement(W + "lvlText", new XAttribute(W + "val", glyph)),
                new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                pPr,
                new XElement(W + "rPr",
                    new XElement(W + "rFonts",
                        new XAttribute(W + "ascii", font),
                        new XAttribute(W + "hAnsi", font),
                        new XAttribute(W + "hint", "default"))));

        return new XElement(W + "lvl",
            new XAttribute(W + "ilvl", lvl),
            new XElement(W + "start", new XAttribute(W + "val", 1)),
            new XElement(W + "numFmt", new XAttribute(W + "val", numFmtToken)),
            new XElement(W + "lvlText", new XAttribute(W + "val", NumberedLvlText(paren, lvl))),
            new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
            pPr);
    }

    /// <summary>
    /// Read the <c>w:startOverride</c> value at <paramref name="ilvl"/> on the <c>w:num</c>
    /// behind <paramref name="numId"/>, or null when the num, the <c>w:lvlOverride</c>, or the
    /// <c>w:startOverride</c> is absent.
    /// </summary>
    public static int? GetStartOverride(WordprocessingDocument doc, int numId, int ilvl)
    {
        var root = doc.MainDocumentPart?.NumberingDefinitionsPart?.GetXDocument().Root;
        var num = root?.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Attribute(W + "numId") == numId.ToString());
        var ovr = num?.Elements(W + "lvlOverride")
            .FirstOrDefault(o => (string?)o.Attribute(W + "ilvl") == ilvl.ToString());
        return int.TryParse((string?)ovr?.Element(W + "startOverride")?.Attribute(W + "val"), out var v)
            ? v : (int?)null;
    }

    /// <summary>
    /// Clone the <c>w:num</c> behind <paramref name="numId"/> into a NEW <c>w:num</c> (fresh
    /// numId, same abstractNumId, existing lvlOverrides copied verbatim) whose
    /// <c>w:lvlOverride[@w:ilvl]/w:startOverride</c> at <paramref name="ilvl"/> is set to
    /// <paramref name="value"/> — or removed, when <paramref name="value"/> is null. Returns the
    /// new numId, or null when <paramref name="numId"/> resolves to no <c>w:num</c>.
    /// </summary>
    /// <remarks>
    /// Additive-only ON PURPOSE: the source num may be shared by paragraphs outside the requested
    /// sequence, so the caller repoints only the affected paragraphs' <c>w:numPr</c> at a clone.
    /// Session snapshots restore both the paragraph references and the numbering definition.
    /// </remarks>
    public static int? CloneNumWithStartOverride(WordprocessingDocument doc, int numId, int ilvl, int? value)
    {
        var part = doc.MainDocumentPart?.NumberingDefinitionsPart;
        var root = part?.GetXDocument().Root;
        var num = root?.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Attribute(W + "numId") == numId.ToString());
        if (part is null || root is null || num is null) return null;

        static int OvrIlvl(XElement o) =>
            int.TryParse((string?)o.Attribute(W + "ilvl"), out var v) ? v : -1;

        int newNumId = NextId(root, "num", "numId");
        var clone = new XElement(num);
        clone.SetAttributeValue(W + "numId", newNumId);

        var ovr = clone.Elements(W + "lvlOverride").FirstOrDefault(o => OvrIlvl(o) == ilvl);
        if (value is { } v)
        {
            var startOverride = new XElement(W + "startOverride", new XAttribute(W + "val", v));
            if (ovr is not null)
            {
                // CT_NumLvl sequence: startOverride, then lvl — replace at the front.
                ovr.Element(W + "startOverride")?.Remove();
                ovr.AddFirst(startOverride);
            }
            else
            {
                ovr = new XElement(W + "lvlOverride", new XAttribute(W + "ilvl", ilvl), startOverride);
                // Keep lvlOverrides in ascending ilvl order (Word's shape); they are the last
                // children of CT_Num, so a plain Add appends correctly when none follows.
                var next = clone.Elements(W + "lvlOverride").FirstOrDefault(o => OvrIlvl(o) > ilvl);
                if (next is not null) next.AddBeforeSelf(ovr); else clone.Add(ovr);
            }
        }
        else if (ovr is not null)
        {
            ovr.Element(W + "startOverride")?.Remove();
            if (!ovr.Elements().Any()) ovr.Remove();
        }

        WordprocessingMLUtil.InsertNumberingChildInOrder(root, clone);
        part.PutXDocument();
        return newNumId;
    }

    /// <summary>
    /// Ensure the abstractNum behind <paramref name="numId"/> defines a <c>w:lvl</c> for every
    /// level up to <paramref name="targetIlvl"/>. Many real-world documents (notably python-docx's
    /// default "List Bullet"/"List Number") define ONLY level 0, so nesting — bumping <c>w:ilvl</c>
    /// past the defined levels — would point at an undefined level and render with no marker/indent
    /// change. This synthesizes the missing level definitions (bullet glyph cycle or decimal,
    /// matching the numbering's existing format) so nesting works on ANY list. Idempotent; mutates
    /// and flushes the numbering part only when a level is actually added. Returns true if it did.
    /// </summary>
    public static bool EnsureLevelDefined(WordprocessingDocument doc, int numId, int targetIlvl)
    {
        if (targetIlvl < 0 || targetIlvl > 8) return false;
        var part = doc.MainDocumentPart?.NumberingDefinitionsPart;
        if (part is null) return false;
        var root = part.GetXDocument().Root;
        if (root is null) return false;

        var num = root.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Attribute(W + "numId") == numId.ToString());
        var absId = (string?)num?.Element(W + "abstractNumId")?.Attribute(W + "val");
        if (absId is null) return false;
        var abstractNum = root.Elements(W + "abstractNum")
            .FirstOrDefault(a => (string?)a.Attribute(W + "abstractNumId") == absId);
        if (abstractNum is null) return false;

        static int LvlOf(XElement e) =>
            int.TryParse((string?)e.Attribute(W + "ilvl"), out var v) ? v : -1;
        bool Defines(int l) => abstractNum.Elements(W + "lvl").Any(e => LvlOf(e) == l);
        if (Defines(targetIlvl)) return false;

        // Synthesize missing levels in the numbering's own format, read off the deepest
        // already-defined level (default: bullet). Parenthesized level text carries down too,
        // so nesting a "(a)" list yields "(a)" sub-levels rather than reverting to "a.".
        var deepest = abstractNum.Elements(W + "lvl")
            .Where(e => LvlOf(e) >= 0).OrderByDescending(LvlOf).FirstOrDefault();
        string numFmtToken = deepest is null
            ? "bullet"
            : (string?)deepest.Element(W + "numFmt")?.Attribute(W + "val") ?? "bullet";
        bool paren = numFmtToken != "bullet"
            && ((string?)deepest?.Element(W + "lvlText")?.Attribute(W + "val"))?.StartsWith("(", StringComparison.Ordinal) == true;

        bool mutated = false;
        for (int l = 0; l <= targetIlvl; l++)
        {
            if (Defines(l)) continue;
            var lvlEl = BuildLevel(numFmtToken, paren, l, SynthBulletGlyphs[l % 3], SynthBulletFonts[l % 3]);
            // w:lvl children must be in ilvl order; insert after the nearest lower level, or before
            // the nearest higher one, else append (lvl is the last child in CT_AbstractNum).
            var prevLvl = abstractNum.Elements(W + "lvl")
                .Where(e => LvlOf(e) >= 0 && LvlOf(e) < l).OrderByDescending(LvlOf).FirstOrDefault();
            if (prevLvl is not null) prevLvl.AddAfterSelf(lvlEl);
            else
            {
                var nextLvl = abstractNum.Elements(W + "lvl")
                    .Where(e => LvlOf(e) > l).OrderBy(LvlOf).FirstOrDefault();
                if (nextLvl is not null) nextLvl.AddBeforeSelf(lvlEl);
                else abstractNum.Add(lvlEl);
            }
            mutated = true;
        }

        if (mutated)
        {
            // A list that defined only level 0 is typically marked singleLevel. WmlToHtmlConverter
            // (ListItemRetriever) FORCES ilvl=0 for singleLevel numbering, so without this upgrade
            // the deeper levels we just added would never render (the nest would still show flat).
            var mlt = abstractNum.Element(W + "multiLevelType");
            if (mlt is null)
            {
                var mltEl = new XElement(W + "multiLevelType", new XAttribute(W + "val", "hybridMultilevel"));
                var nsid = abstractNum.Element(W + "nsid");
                if (nsid is not null) nsid.AddAfterSelf(mltEl); else abstractNum.AddFirst(mltEl);
            }
            else if ((string?)mlt.Attribute(W + "val") == "singleLevel")
            {
                mlt.SetAttributeValue(W + "val", "hybridMultilevel");
            }
            part.PutXDocument();
        }
        return mutated;
    }
}
