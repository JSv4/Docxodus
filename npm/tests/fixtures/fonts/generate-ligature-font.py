#!/usr/bin/env python3
"""Regenerate the tiny ligature test font used by tests/export-ligatures.spec.ts; written to ligature/ (#1562).

Original rectangle outlines: A-Z at 625 units and space at 250 in a 1000-unit em, whole pixels at
12pt (16px), so hinted layouts round nothing. One standard ligature: F F becomes a 750-unit F_F
glyph, 500 units narrower than the pair. No host font or third-party outline is used; the
repository MIT license applies.
"""

from pathlib import Path

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen

ROOT = Path(__file__).resolve().parent / "ligature"
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def box(width):
    pen = TTGlyphPen(None)
    if width:
        pen.moveTo((60, 0))
        pen.lineTo((width - 60, 0))
        pen.lineTo((width - 60, 700))
        pen.lineTo((60, 700))
        pen.closePath()
    return pen.glyph()


order = [".notdef", "space", *LETTERS, "F_F"]
widths = {".notdef": 625, "space": 250, "F_F": 750, **{letter: 625 for letter in LETTERS}}
builder = FontBuilder(1000, isTTF=True)
builder.setupGlyphOrder(order)
builder.setupCharacterMap({32: "space", **{ord(letter): letter for letter in LETTERS}})
builder.setupGlyf({name: box(0 if name == "space" else widths[name]) for name in order})
builder.setupHorizontalMetrics({name: (widths[name], 0) for name in order})
builder.setupHorizontalHeader(ascent=800, descent=-200)
builder.setupNameTable({
    "familyName": "Docxodus Ligature Test",
    "styleName": "Regular",
    "uniqueFontIdentifier": "Docxodus tests: DocxodusLigatureTest-Regular: 1.000",
    "fullName": "Docxodus Ligature Test Regular",
    "psName": "DocxodusLigatureTest-Regular",
    "version": "Version 1.000",
})
builder.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200,
                 fsType=0, usWeightClass=400, usWidthClass=5)
builder.setupPost()
builder.addOpenTypeFeatures("languagesystem DFLT dflt; languagesystem latn dflt; feature liga { sub F F by F_F; } liga;")
# Fixed timestamps keep regeneration byte-identical.
builder.font["head"].created = builder.font["head"].modified = 3_786_825_600
builder.save(ROOT / "docxodus-ligature-test.ttf")
