#!/usr/bin/env python3
"""Regenerate tiny deterministic fonts used only by the #442 security tests."""

from pathlib import Path
import struct

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont


ROOT = Path(__file__).resolve().parent
CHARACTERS = " ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def glyph_name(character: str) -> str:
    return "space" if character == " " else character


def build_ttf(path: Path, family: str, postscript_name: str, advance_scale: float = 1.0) -> None:
    order = [".notdef", *(glyph_name(character) for character in CHARACTERS)]
    glyphs = {}
    metrics = {}
    for index, name in enumerate(order):
        pen = TTGlyphPen(None)
        base_width = 600 if name == "space" else 540 + (index % 4) * 30
        width = round(base_width * advance_scale)
        if name != "space":
            pen.moveTo((60, 0))
            pen.lineTo((width - 60, 0))
            pen.lineTo((width - 60, 700))
            pen.lineTo((60, 700))
            pen.closePath()
        glyphs[name] = pen.glyph()
        metrics[name] = (width, 0)

    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap({ord(character): glyph_name(character) for character in CHARACTERS})
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics(metrics)
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({
        "familyName": family,
        "styleName": "Regular",
        "uniqueFontIdentifier": f"Docxodus tests: {postscript_name}: 1.000",
        "fullName": f"{family} Regular",
        "psName": postscript_name,
        "version": "Version 1.000",
    })
    builder.setupOS2(
        sTypoAscender=800,
        sTypoDescender=-200,
        usWinAscent=800,
        usWinDescent=200,
        fsType=0,
        usWeightClass=400,
        usWidthClass=5,
    )
    builder.setupPost()
    builder.setupMaxp()
    builder.font["head"].created = 2_082_844_800
    builder.font["head"].modified = 2_082_844_800
    builder.save(path)


def woff_from_ttf(source: Path, destination: Path) -> None:
    font = TTFont(source, recalcTimestamp=False)
    font.flavor = "woff"
    font.save(destination, reorderTables=False)


def load_failure_from_ttf(source: Path, destination: Path) -> None:
    data = bytearray(source.read_bytes())
    table_count = struct.unpack_from(">H", data, 4)[0]
    for index in range(table_count):
        record = 12 + index * 16
        if data[record:record + 4] == b"glyf":
            # Metadata/cmap remain readable to fontkit, while Chromium's
            # OpenType sanitizer rejects the impossible outline table.
            struct.pack_into(">I", data, record + 12, 1)
            destination.write_bytes(data)
            return
    raise RuntimeError("generated font has no glyf table")


def main() -> None:
    carlito = ROOT / "synthetic-carlito.ttf"
    policy = ROOT / "docxodus-policy-base.ttf"
    metric = ROOT / "docxodus-metric-test.ttf"
    build_ttf(carlito, "Carlito", "SyntheticCarlito-Regular")
    build_ttf(ROOT / "synthetic-carlito-narrow.ttf", "Carlito", "SyntheticCarlitoNarrow-Regular", 0.5)
    build_ttf(ROOT / "synthetic-carlito-wide.ttf", "Carlito", "SyntheticCarlitoWide-Regular", 1.6)
    build_ttf(policy, "Docxodus Policy Test", "DocxodusPolicyTest-Regular")
    build_ttf(metric, "Docxodus Metric Test", "DocxodusMetricTest-Regular")
    woff_from_ttf(metric, ROOT / "docxodus-metric-test.woff")
    load_base = ROOT / "docxodus-load-base.ttf"
    build_ttf(load_base, "Docxodus Load Failure", "DocxodusLoadFailure-Regular")
    load_failure_from_ttf(load_base, ROOT / "docxodus-load-failure.ttf")
    load_base.unlink()


if __name__ == "__main__":
    main()
