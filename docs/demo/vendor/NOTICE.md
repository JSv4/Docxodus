# Third-party notices — `docs/demo/vendor/`

**Docxodus is MIT-licensed (see the root `LICENSE`). Nothing in this directory
is.** Everything here is third-party software and third-party game data,
vendored so the DOCX Arcade's Doom cartridge can run with no build step and no
external CDN, and each piece keeps the license it arrived with.

Nothing in this directory is compiled into, linked against, or shipped in the
`Docxodus` NuGet package or the `docxodus` npm package. It is loaded only by
`docs/demo/doom-cart.js` — a demo page on the GitHub Pages site — and the npm
package boundary check (`npm/scripts/check-package-boundary.mjs`) fails the
build if any of it ever reaches the published tarball.

---

## `doomgeneric/doomgeneric_module.js` — GNU General Public License, version 2

id Software's Doom source, by way of
[doomgeneric](https://github.com/ozkl/doomgeneric) and the JavaScript port
[doomgenericjs](https://github.com/grubbyplaya/doomgenericjs). id released the
Doom source under the GPL in 1997; this is a build of it, so it is GPL-2.0 and
so is anything combined with it.

| | |
|---|---|
| Upstream | `https://github.com/grubbyplaya/doomgenericjs` |
| Commit | `99d7a55651b5f774e9b8911ef96e91a3652ef85f` ("Fix ES6 module") |
| Path | `doomgeneric/doomgeneric_module.js` |
| SHA-256 | `efd7c34714e3753a84cf6873f2c2ae0e1a41bde1f10aaa2b0a95cb6935e8cce9` |
| Size | 3,097,628 bytes |
| License | GPL-2.0 — full text in `doomgeneric/COPYING` |

The file is vendored **byte-for-byte as upstream built it**; the hash above is
the check. It is the `makefile.module` target — Emscripten with
`-sMODULARIZE=1 -sEXPORT_ES6=1 -sSINGLE_FILE=1 -sWASM=0`, which is why it is a
single 3 MB ES module with the runtime inlined and no separate `.wasm` to
fetch. Corresponding source for this binary is the upstream repository at the
commit named above, which carries the complete Doom C sources it was built
from.

### What this means for the license of the code that calls it

`docs/demo/doom-cart.js` is written against this engine and combined with it at
runtime, so **that file is offered under GPL-2.0-or-later**, and says so in its
own header. It is the only GPL file in the Docxodus tree. Every other file in
`docs/demo/` — including `ascii-arcade.js`, which imports the cartridge — stays
MIT, which is GPL-compatible: the *combination* that a visitor's browser
assembles when they select the Doom cartridge is GPL-2.0, while each MIT file
remains separately available under MIT.

## `freedoom/freedoom1.wad.gz` — BSD 3-Clause

Doom is an engine; it needs an IWAD to have anything to render. Doom's *source*
is free software but its *game data* never was, so the arcade runs on
[Freedoom](https://freedoom.github.io/) Phase 1 — a complete, freely licensed
replacement IWAD. doomgeneric's own IWAD table (`d_iwad.c`) already recognises
`freedoom1.wad` and treats it as a retail Doom, so no configuration is needed
beyond the filename.

| | |
|---|---|
| Upstream | `https://github.com/freedoom/freedoom/releases/tag/v0.13.0` |
| Release asset | `freedoom-0.13.0.zip` → `freedoom-0.13.0/freedoom1.wad` |
| SHA-256 (`freedoom1.wad`, uncompressed) | `7323bcc168c5a45ff10749b339960e98314740a734c30d4b9f3337001f9e703d` |
| SHA-256 (`freedoom1.wad.gz`, as shipped) | `8762911b2a372fa5c9fc6e725b36eec939a55434e9d2c61a2f9ad6596c2e95c0` |
| Size | 28,795,076 bytes → 10,278,537 gzipped |
| License | BSD 3-Clause — full text in `freedoom/COPYING.txt`, verbatim from the release |

The only change from the release asset is gzip: GitHub Pages will not
content-encode an `application/octet-stream`, so compressing the WAD in the
repository is the only way the visitor downloads 10 MB instead of 29. The
browser inflates it with `DecompressionStream('gzip')` — no library — and
`doom-cart.js` checks the `IWAD` magic before handing it to the engine.

Reproduce the shipped file with:

```bash
curl -LO https://github.com/freedoom/freedoom/releases/download/v0.13.0/freedoom-0.13.0.zip
unzip -j freedoom-0.13.0.zip freedoom-0.13.0/freedoom1.wad freedoom-0.13.0/COPYING.txt
sha256sum freedoom1.wad   # must be 7323bcc1…
gzip -9 -c freedoom1.wad > freedoom1.wad.gz
```

### Why not the shareware `doom1.wad`

doomgenericjs vendors id Software's shareware `doom1.wad`, and at 4.2 MB it is
a third the size. It is not used here. id's shareware terms cover redistributing
the complete, unmodified shareware *package*; lifting the IWAD out of it and
committing it to an unrelated public repository is not that, and the data is
copyrighted commercial game content either way. Freedoom exists precisely so
that a project can ship a playable Doom without that question, and this
repository already made that choice once — the cartridge this one replaced was
built from Freedoom's E1M1.

Nothing in the code prefers one IWAD over the other: `doomCart({ wadUrl })`
takes an IWAD URL and `arcade.html?wad=…` passes one through, so anyone holding
a licensed `doom.wad` or `doom2.wad` can point the cartridge at their own copy
and play the retail maps. The URL must be same-origin — host the WAD beside the
page — and the filename has to be one Doom's own IWAD table recognises.
