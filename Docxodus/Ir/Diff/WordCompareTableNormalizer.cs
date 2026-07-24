#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Replicates one piece of Microsoft Word's compare-output table normalization: for a FIXED-WIDTH
/// table (<c>w:tblW w:type="dxa"</c>) that declares no explicit cell margins, Word materializes a
/// hairline <c>w:tblCellMar</c> (left/right) plus a matching <c>w:tblInd</c>. Without this, a renderer
/// applies its OWN default cell margin (LibreOffice ≈ 108 twips) which overflows the fixed column
/// widths and shifts every cell's text horizontally versus Word's output — the whole table "ghosts".
///
/// The inset equals the table's border width: a 0.5pt (<c>w:sz="4"</c>) border ⇒ 10 twips, matching the
/// value Word wrote verbatim across every fixed-width table in the compare corpus (border <c>w:sz</c> is
/// in eighths of a point; 1pt = 20 twips ⇒ twips = <c>sz × 2.5</c>). A table with no border gives no
/// width to derive from, so it is left alone (we do not guess); an AUTO-width table is left alone (Word
/// applies no such normalization there); a table that already declares <c>w:tblCellMar</c> is left alone.
///
/// This mirrors the docDefaults backfill (<see cref="WordStockDocDefaults"/>) — the whole engine goal is
/// to reproduce what Word's compare emits, and a rendered redline is only faithful if its tables land
/// where Word's do. Applied as a single-owner post-pass over the assembled body blocks so every table
/// (top-level or nested, on any render path) is normalized identically.
/// </summary>
internal static class WordCompareTableNormalizer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>CT_TblPrBase child order, enough of it to place <c>w:tblInd</c> and <c>w:tblCellMar</c>.</summary>
    private static readonly XName[] TblPrOrder =
    {
        W + "tblStyle", W + "tblpPr", W + "tblOverlap", W + "bidiVisual",
        W + "tblStyleRowBandSize", W + "tblStyleColBandSize", W + "tblW", W + "jc",
        W + "tblCellSpacing", W + "tblInd", W + "tblBorders", W + "shd", W + "tblLayout",
        W + "tblCellMar", W + "tblLook", W + "tblCaption", W + "tblDescription", W + "tblPrChange",
    };

    /// <summary>Normalize every fixed-width table reachable from these body blocks, in place.</summary>
    internal static void NormalizeAll(IEnumerable<XElement> blocks)
    {
        foreach (var block in blocks)
            foreach (var tbl in DescendantTables(block))
                NormalizeTable(tbl);
    }

    private static IEnumerable<XElement> DescendantTables(XElement block)
        => block.Name == W + "tbl"
            ? new[] { block }.Concat(block.DescendantsAndSelf(W + "tbl").Skip(1))
            : block.Descendants(W + "tbl");

    private static void NormalizeTable(XElement tbl)
    {
        var tblPr = tbl.Element(W + "tblPr");
        if (tblPr is null)
            return;

        // Only fixed-width (dxa) tables — auto/pct layout is not normalized by Word here.
        var tblW = tblPr.Element(W + "tblW");
        if ((string?)tblW?.Attribute(W + "type") != "dxa")
            return;

        // Already-declared cell margins are authoritative — never overwrite them.
        if (tblPr.Element(W + "tblCellMar") is not null)
            return;

        // Derive the hairline inset from the table's border width; with no border there is nothing to
        // derive and Word's behavior is unattested, so leave the table untouched.
        int? insetTwips = BorderInsetTwips(tblPr);
        if (insetTwips is not { } inset)
            return;

        var cellMar = new XElement(
            W + "tblCellMar",
            new XElement(W + "left", new XAttribute(W + "w", inset), new XAttribute(W + "type", "dxa")),
            new XElement(W + "right", new XAttribute(W + "w", inset), new XAttribute(W + "type", "dxa")));
        InsertInOrder(tblPr, cellMar);

        // Word pairs the inset with a matching table indent, but only when none is already declared.
        if (tblPr.Element(W + "tblInd") is null)
            InsertInOrder(tblPr,
                new XElement(W + "tblInd", new XAttribute(W + "w", inset), new XAttribute(W + "type", "dxa")));
    }

    /// <summary>The table's vertical border width in twips (left, else insideV, else any side), or null
    /// when no single-line border width can be read.</summary>
    private static int? BorderInsetTwips(XElement tblPr)
    {
        var borders = tblPr.Element(W + "tblBorders");
        if (borders is null)
            return null;

        var edge = borders.Element(W + "left")
            ?? borders.Element(W + "insideV")
            ?? borders.Elements().FirstOrDefault(e => e.Attribute(W + "sz") is not null);
        var szAttr = edge?.Attribute(W + "sz");
        if (szAttr is null || !int.TryParse(szAttr.Value, out var szEighthsPt) || szEighthsPt <= 0)
            return null;

        // sz is eighths of a point; 1pt = 20 twips ⇒ twips = sz / 8 * 20 = sz * 2.5.
        return (int)Math.Round(szEighthsPt * 2.5, MidpointRounding.AwayFromZero);
    }

    private static void InsertInOrder(XElement tblPr, XElement child)
    {
        int rank = Array.IndexOf(TblPrOrder, child.Name);
        XElement? after = null;
        foreach (var existing in tblPr.Elements())
        {
            int r = Array.IndexOf(TblPrOrder, existing.Name);
            if (r >= 0 && r < rank)
                after = existing;
        }
        if (after is null)
            tblPr.AddFirst(child);
        else
            after.AddAfterSelf(child);
    }
}
