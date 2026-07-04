#nullable enable

using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Synthesizes reusable bullet / decimal numbering definitions so a plain paragraph can be
/// promoted to a real list item. <see cref="DocxSession.ApplyListFormat"/> uses this when no
/// suitable numbering exists. Definitions are tagged with a fixed marker <c>w:nsid</c> per
/// format and resolved find-or-create, so the op is idempotent across calls, save/reopen, and
/// undo (the numbering part is not snapshotted; the paragraph's <c>w:numPr</c> is).
/// </summary>
internal static class NumberingFactory
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Stable per-format markers (8-hex nsid values) used to find-or-create our own definition.
    private const string BulletNsid = "0D0CB001";
    private const string DecimalNsid = "0D0CD001";

    // Standard Word bullet cycle (•, o, ▪) for synthesized nested levels — same glyph/font set
    // BuildAbstractNum emits for our own multi-level lists, so source and synthesized lists nest
    // identically.
    private static readonly string[] SynthBulletGlyphs = { "•", "o", "▪" };
    private static readonly string[] SynthBulletFonts = { "Symbol", "Courier New", "Wingdings" };

    /// <summary>
    /// Ensure a bullet or decimal numbering definition exists and return a numId pointing at it.
    /// Only NumberFormat.Bullet and NumberFormat.Decimal are supported here.
    /// </summary>
    public static int EnsureNumbering(WordprocessingDocument doc, NumberFormat fmt)
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
        bool bullet = fmt == NumberFormat.Bullet;
        string nsid = bullet ? BulletNsid : DecimalNsid;

        // Find our previously-synthesized abstractNum (by marker nsid), or build one.
        var abstractNum = root.Elements(W + "abstractNum")
            .FirstOrDefault(a => (string?)a.Element(W + "nsid")?.Attribute(W + "val") == nsid);
        if (abstractNum is null)
        {
            int absId = NextId(root, "abstractNum", "abstractNumId");
            abstractNum = BuildAbstractNum(bullet, absId, nsid);
            // CT_Numbering order: numPicBullet*, abstractNum*, num* — keep abstractNums grouped.
            var lastAbstract = root.Elements(W + "abstractNum").LastOrDefault();
            if (lastAbstract is not null) lastAbstract.AddAfterSelf(abstractNum);
            else
            {
                var firstNum = root.Elements(W + "num").FirstOrDefault();
                if (firstNum is not null) firstNum.AddBeforeSelf(abstractNum);
                else root.Add(abstractNum);
            }
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
            root.Add(num); // nums come after abstractNums
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

    /// <summary>Build a spec-valid 9-level bullet or decimal abstractNum.</summary>
    private static XElement BuildAbstractNum(bool bullet, int absId, string nsid)
    {
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
                    new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
                    new XElement(W + "lvlText", new XAttribute(W + "val", $"%{lvl + 1}.")),
                    new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                    pPr);
            }
            an.Add(lvl_);
        }

        return an;
    }

    /// <summary>Build one spec-valid <c>w:lvl</c> (bullet or decimal) at level <paramref name="lvl"/>.</summary>
    private static XElement BuildLevel(bool bullet, int lvl, string glyph, string font)
    {
        var pPr = new XElement(W + "pPr",
            new XElement(W + "ind",
                new XAttribute(W + "left", 720 * (lvl + 1)),
                new XAttribute(W + "hanging", 360)));

        if (bullet)
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
            new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
            new XElement(W + "lvlText", new XAttribute(W + "val", $"%{lvl + 1}.")),
            new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
            pPr);
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

        // Detect bullet vs numbered from the deepest already-defined level (default: bullet).
        var deepest = abstractNum.Elements(W + "lvl")
            .Where(e => LvlOf(e) >= 0).OrderByDescending(LvlOf).FirstOrDefault();
        bool bullet = deepest is null
            || (string?)deepest.Element(W + "numFmt")?.Attribute(W + "val") == "bullet";

        bool mutated = false;
        for (int l = 0; l <= targetIlvl; l++)
        {
            if (Defines(l)) continue;
            var lvlEl = BuildLevel(bullet, l, SynthBulletGlyphs[l % 3], SynthBulletFonts[l % 3]);
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
