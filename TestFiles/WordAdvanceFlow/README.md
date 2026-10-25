# Native Word word-advance controls

These original 11pt and 12pt inputs isolate glyph advances, natural word gaps, final-line spacing, repeated spaces and wrapping. Both have zero Open XML SDK findings. The raw Word PDF and returned DOCX remain in the immutable issue1124 fixture archive; `word-observations.json` pins their identities and preserves native word coordinates. Word did not expose its build number.

The browser tests use independently read Liberation Serif font-table advances for the same natural `pa` words. They cover four alignments and three compatibility settings while preserving inline geometry and complete source text. Native Times New Roman has the same relevant advances. PDF placement quantization remains an explicit residual; no arbitrary pixel grid or offset is fitted. The full public-export comparison separately retains native line and page inventories.
