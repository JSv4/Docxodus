# @docxodus/export

Supported Node.js and command-line export of DOCX files to standalone paginated HTML and PDF.
The package launches pinned Chromium around the public `docxodus/export-browser` materializer; it
does not contain a second converter or pagination engine.

## Install

Install matching package versions. The companion installs its tested Chromium revision during
deployment and never downloads a browser while converting a document.

```console
npm install docxodus@9.9.0 @docxodus/export@9.9.0
```

Node.js 22.13 or newer is required. Environments that omit the bundled browser can pass an explicit
`browserExecutablePath`; the CLI also reads `DOCXODUS_CHROMIUM_PATH`. No PATH-wide browser guessing
occurs.

## Node API

```js
import { readFile, writeFile } from "node:fs/promises";
import { convertDocxToPdf } from "@docxodus/export";

const source = new Uint8Array(await readFile("contract.docx"));
const result = await convertDocxToPdf(source, {
  reviewProfile: "final",
  commentProfile: "endnotes",
  documentVersion: 12,
  timeoutMs: 120_000,
});

await writeFile("contract.pdf", result.pdf);
console.log(result.pageCount, result.pageMap, result.renderReport);
```

`renderDocxArtifacts()` produces HTML and PDF from one browser materialization. The file API adds
stable source reads and atomic no-replace destination commits:

```js
import { renderDocxFile } from "@docxodus/export";

await renderDocxFile("contract.docx", {
  pdfPath: "contract.pdf",
  pageMapPath: "contract.pages.json",
  reportPath: "contract.render.json",
}, {
  reviewProfile: "markup",
  commentProfile: "margin",
});
```

Caller `Uint8Array` values are synchronously copied. A caller-supplied Playwright Chromium
`Browser` remains caller-owned; Docxodus closes only the fresh context it creates. Every runtime
asset is length/hash checked against the public manifest and served at a routed `.invalid` origin.
Any request outside that closed graph fails the operation.

PDF printing enables backgrounds, uses CSS page sizes, applies zero browser margins, preserves DOM
text/links/vector content, and verifies every page MediaBox/CropBox with a real PDF parser. Reports
record the exact PDF SHA-256 and Chromium's volatile metadata; PDF byte identity is intentionally
not claimed across runs.

## CLI

Profiles are explicit:

```console
docxodus convert contract.docx --to pdf --output contract.pdf \
  --review-profile final --comments endnotes \
  --timeout 120000 --report contract.render.json \
  --page-map contract.pages.json
```

Additional flags include `--document-version`, `--expected-source-digest`, `--title`,
`--unsupported-content`, `--strict-fonts`, `--browser-executable`, repeatable `--limit
name=integer`, repeatable `--font-directory`, `--font-license-attestations`, and
`--environment-attestation`. Artifact bytes are never written to stdout. Existing destinations,
input aliases, and duplicate destinations are rejected; the CLI never overwrites a file.

## Runtime boundaries

- Explicit font-directory loading is accepted by the public contract but fails with
  `unsupported_runtime` until issue #442 supplies the verified pre-layout font hook. Browser-observed
  fonts remain reported honestly; `strictFonts` therefore fails closed.
- The `original` review profile remains fail-closed until issue #444 completes its projection.
- Mixed-section and broader fidelity ratchets are extended by issues #440 and #443.

Failures are `DocxodusExportError` objects with stable code, phase, remediation, safe detail, and a
structured failed report when materialization had begun.

## Framed host

`docxodus-export-host` is the non-shell integration boundary used by delivery adapters. It accepts
exactly one length-prefixed JSON envelope (four-byte, unsigned big-endian length), returns one
length-prefixed response, rejects duplicate batch IDs and unknown fields, and encodes PDF payloads
as canonical base64. One host-owned Chromium browser is reused across the envelope while every
batch receives a fresh isolated context. Executable and font-directory authority is deliberately
process-owned: a deployment may set `DOCXODUS_CHROMIUM_PATH`, but a framed request cannot select a
local executable or filesystem directory.
