# Third-party notices — the DOCX Arcade's Doom cartridge

**Docxodus is MIT-licensed (see the root `LICENSE`). The two things this
directory describes are not.** Neither of them is *in* this repository: both
are pinned CDN URLs, and what lives here is the license text and the
provenance record for them.

They are not vendored because 13 MB of binary in a docs directory shows up in
every clone forever, never diffs usefully, and — in the engine's case — is a
verbatim copy of something upstream already publishes. jsDelivr serves both
with `access-control-allow-origin: *` and an immutable cache, and resolves a
40-hex commit to exactly those bytes, so a pin cannot change underneath us.

Nothing described here is compiled into, linked against, or shipped in the
`Docxodus` NuGet package or the `docxodus` npm package. Both are loaded only by
`docs/demo/doom-cart.js`, a demo page on the GitHub Pages site.

---

## The engine — GNU General Public License, version 2

id Software's Doom source, by way of
[doomgeneric](https://github.com/ozkl/doomgeneric) and the JavaScript port
[doomgenericjs](https://github.com/grubbyplaya/doomgenericjs). id released the
Doom source under the GPL in 1997; this is a build of it, so it is GPL-2.0, and
so is anything combined with it.

| | |
|---|---|
| URL | `https://cdn.jsdelivr.net/gh/grubbyplaya/doomgenericjs@99d7a55651b5f774e9b8911ef96e91a3652ef85f/doomgeneric/doomgeneric_module.js` |
| Upstream | `https://github.com/grubbyplaya/doomgenericjs` |
| Commit | `99d7a55651b5f774e9b8911ef96e91a3652ef85f` ("Fix ES6 module") |
| SHA-256 | `efd7c34714e3753a84cf6873f2c2ae0e1a41bde1f10aaa2b0a95cb6935e8cce9` |
| Size | 3,097,628 bytes |
| License | GPL-2.0 — full text in `doomgeneric/COPYING` |

This is upstream's own published build, served from upstream's own tree — the
pin is by commit, so it is the same artifact a `git checkout` of that SHA would
produce. It is the `makefile.module` target: Emscripten with
`-sMODULARIZE=1 -sEXPORT_ES6=1 -sSINGLE_FILE=1 -sWASM=0`, which is why it is a
single 3 MB ES module with the runtime inlined and no separate `.wasm` to
fetch. Corresponding source for it is that repository at that commit, which
carries the complete Doom C sources it was built from.

Verify the pin at any time with:

```bash
curl -sL 'https://cdn.jsdelivr.net/gh/grubbyplaya/doomgenericjs@99d7a55651b5f774e9b8911ef96e91a3652ef85f/doomgeneric/doomgeneric_module.js' | sha256sum
```

### What this means for the license of the code that calls it

`docs/demo/doom-cart.js` is written against this engine and combined with it at
runtime, so **that file is offered under GPL-2.0-or-later**, and says so in its
own header. It is the only GPL file in the Docxodus tree — which is why
`doomgeneric/COPYING` still ships here even though the engine binary does not.
Every other file in `docs/demo/` — including `ascii-arcade.js`, which imports
the cartridge — stays MIT, which is GPL-compatible: each file remains
separately available under its own terms, and only the combination a visitor's
browser assembles when the Doom cartridge is selected is GPL-2.0.

## The game data — BSD 3-Clause

Doom is an engine; it needs an IWAD to have anything to render. Doom's *source*
is free software but its *game data* never was, so the arcade runs on
[Freedoom](https://freedoom.github.io/) Phase 1, a complete and freely licensed
replacement IWAD. doomgeneric's own IWAD table (`d_iwad.c`) recognises
`freedoom1.wad` and treats it as a retail Doom, so the filename is the whole
configuration.

| | |
|---|---|
| URL | `https://cdn.jsdelivr.net/gh/JSv4/freedoom-iwad@70ee6ec942d090b4dd7ba04927f09ac79c8dc085/freedoom1.wad.gz` |
| Republished in | `https://github.com/JSv4/freedoom-iwad` at `70ee6ec942d090b4dd7ba04927f09ac79c8dc085` |
| Upstream | `https://github.com/freedoom/freedoom/releases/tag/v0.13.0` |
| Release asset | `freedoom-0.13.0.zip` → `freedoom-0.13.0/freedoom1.wad` |
| SHA-256 (uncompressed) | `7323bcc168c5a45ff10749b339960e98314740a734c30d4b9f3337001f9e703d` |
| Size | 28,795,076 bytes, ~10 MB gzipped |
| License | BSD 3-Clause — full text in `freedoom/COPYING.txt`, verbatim from the release |

The recorded digest is of the **uncompressed** WAD, deliberately: that is the
artifact Freedoom published and the only hash that is stable, since two
correct gzip implementations at the same level need not agree byte-for-byte.
`doom-cart.js` inflates the file with `DecompressionStream('gzip')` and checks
the `IWAD` magic before handing it to the engine.

### Why the IWAD needs a repository of its own

GitHub serves release assets **without CORS**, so a browser cannot fetch
`freedoom1.wad` from the Freedoom release directly — only a server-side tool
can. jsDelivr will serve any file in a public repository's tree with
`access-control-allow-origin: *`, but Freedoom's WADs are built artifacts and
are not in its tree. So the file is republished in a small sibling repository
whose only purpose is to be CDN-addressable.

[`JSv4/freedoom-iwad`](https://github.com/JSv4/freedoom-iwad) contains exactly:

- `freedoom1.wad.gz` — `gzip -9` of the release asset above;
- `COPYING.txt` — Freedoom's license, verbatim (BSD 3-Clause requires the
  notice travel with a redistribution, and that repository *is* a
  redistribution);
- a `README.md` pointing back here.

Rebuild its contents with:

```bash
curl -LO https://github.com/freedoom/freedoom/releases/download/v0.13.0/freedoom-0.13.0.zip
unzip -j freedoom-0.13.0.zip freedoom-0.13.0/freedoom1.wad freedoom-0.13.0/COPYING.txt
sha256sum freedoom1.wad   # must be 7323bcc1…
gzip -9 freedoom1.wad
```

**The pin is a commit, not a tag** — the same shape as the engine's. jsDelivr
resolves `@<40-hex>` to exactly those bytes, so unlike a tag it cannot be moved
under us. Changing the IWAD therefore means a new commit there and a new SHA
here; nothing needs re-tagging, and nothing can silently change beneath a
version that is already published.

### The tests do not use the CDN

`npm/scripts/fetch-doom-iwad.mjs` pulls the release asset straight from
Freedoom (server-side, where CORS does not apply), verifies the SHA-256 above,
and writes a gzipped copy into the Playwright webroot. The specs then pass
`?wad=./vendor/freedoom1.wad.gz`, which is same-origin. A browser suite that
depended on a CDN being up, and on a sibling repository being reachable, would
fail for reasons that have nothing to do with the change under test.

### Why not the shareware `doom1.wad`

doomgenericjs carries id Software's shareware `doom1.wad` in its tree, so it is
already CDN-addressable and is a third the size. It is not used here. id's
shareware terms cover redistributing the complete, unmodified shareware
*package*; building a product's default experience on an IWAD lifted out of it
is not that, and the data is copyrighted commercial game content either way.
Freedoom exists precisely so a project can ship a playable Doom without that
question.

Nothing in the code prefers one IWAD over the other. `doomCart({ wadUrl })`
takes an IWAD URL and `arcade.html?wad=…` passes one through, so anyone holding
a licensed `doom.wad` or `doom2.wad` can point the cartridge at their own copy
and play the retail maps. The URL must be same-origin — host the WAD beside the
page — and the filename has to be one Doom's own IWAD table recognises.

There is deliberately no equivalent override for the engine. `import()`
executes what it fetches, on the page's origin and with its privileges, so a
URL parameter naming the module would be remote code execution from a crafted
link. The engine is only ever the pin above; the cartridge's URL gate allows
exactly that pin plus same-origin, and nothing else.
