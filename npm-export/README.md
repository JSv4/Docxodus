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
stable source reads and a staged no-replace publication transaction. All payloads are fsynced
before the first destination becomes visible; a later commit failure rolls back every destination
that is still owned and unmodified, and reports any path that could not safely be rolled back:

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

PDF printing reopens the finalized snapshot in a second script-disabled isolated context, enables
backgrounds, uses CSS page sizes, applies zero browser margins, preserves DOM text/links/vector
content, and verifies the exact PageMap plus every page's effective inherited MediaBox/CropBox,
rotation, and `UserUnit` with a real PDF parser. Reports record the exact PDF SHA-256 and Chromium's
volatile metadata; PDF byte identity is intentionally not claimed across runs. Mixed
portrait/landscape and Letter/A4 sections retain their per-page CSS size; screen zoom/transforms
and their compensation margins are removed by the print contract and cannot change physical PDF
geometry or content placement.

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

- Repeatable `--font-directory` (API: `fontDirectories`) supplies verified fonts: each file is
  discovered, licence-checked, hashed, and matched inside the sandboxed page through a resolver
  binding — font bytes cross the boundary, filesystem paths never do. The render report records
  every requested family's resolution (`resolved`, `substituted`, `missing`, `load_failed`) with
  file digests, and `strictFonts` turns unresolved or unverified families into failures. Without
  font directories the render falls back to browser-observed fonts, reported honestly as
  unverified. See "Reproducible font configuration" below.
- All three review profiles (`final`, `original`, `markup`) are supported; `final` and `original`
  derive their projection out-of-place and never mutate the source bytes.
- Generated PDFs are covered by the same visual-fidelity ratchet as the HTML renderer.
- Chromium keeps its process sandbox, so the render host has to permit unprivileged user
  namespaces. Ubuntu 23.10 and later restrict them through AppArmor by default; check with
  `unshare --user --map-root-user true` and permit them with
  `sysctl -w kernel.apparmor_restrict_unprivileged_userns=0`. A launch that fails this way is
  reported as its own condition rather than as a suspect executable.

Failures are `DocxodusExportError` objects with stable code, phase, remediation, safe detail, and a
structured failed report when materialization had begun. The CLI additionally writes the underlying
cause chain to stderr, which is where a Chromium launch diagnostic becomes readable.

## Reproducible font configuration

Layout depends on fonts, so a reproducible render pins them. The supported configuration is the
license-safe metric-substitute set the visual-parity gates run under (the shared contract in
`docxodus`'s `font-contract` module; rationale in `npm/tests/visual-parity/README.md`):

| Declared family | Substitute | Debian package | Metric-compatible |
|---|---|---|---|
| Calibri | Carlito | fonts-crosextra-carlito | yes |
| Calibri Light | Carlito | fonts-crosextra-carlito | no — documented approximation |
| Cambria | Caladea | fonts-crosextra-caladea | yes |
| Times New Roman | Liberation Serif | fonts-liberation2 | yes |
| Arial | Liberation Sans | fonts-liberation2 | yes |
| Courier New | Liberation Mono | fonts-liberation2 | yes |

Install the packages and hand the exporter their directories:

```console
apt-get install fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation2
docxodus convert contract.docx --to pdf --output contract.pdf \
  --font-directory /usr/share/fonts/truetype/crosextra \
  --font-directory /usr/share/fonts/truetype/liberation
```

The render report then records each family's resolution with the exact file SHA-256, and the
renderer fingerprint binds the Chromium identity to the font-resolution digest, so two hosts
running this configuration produce comparable fingerprints. TTF and OTF files are admitted on
their embedded licensing flags alone; WOFF and WOFF2 files additionally require an explicit
embedding attestation (`--font-license-attestations`). A font whose embedded license forbids
embedding is never used — the export fails rather than silently substituting.

## Framed host

`docxodus-export-host` is the non-shell integration boundary used by delivery adapters. Protocol v1
is a bounded frame sequence:

1. A strict UTF-8 JSON control frame (four-byte unsigned big-endian byte length) declares unique
   sources as `{ id, byteLength, sha256, mediaType }` and batches as
   `{ id, sourceId, artifactRequestIds, options }`.
2. One exact-length raw DOCX frame follows for each source, in declaration order. Sources reused by
   several batches cross the pipe once. Digest, length, ids, counts, aggregate bytes, canonical
   ordering, unknown fields, duplicate JSON properties, and trailing input are verified before a
   browser starts.
3. A successful response begins with a bounded control frame containing keyed batch metadata and
   digest/length/media-type artifact descriptors, followed by raw HTML, PDF, PageMap, and report
   frames in descriptor order. Canonical JSON artifacts contain no newline or frame prefix.

Any batch or cleanup failure makes the logical request fatal; earlier successful payloads are not
returned as nominal artifacts. A safely retained failed report may follow only as a declared
diagnostic artifact. One host-owned Chromium browser is reused while every batch receives a fresh
isolated materialization context and a separate PDF context. Executable and font-directory
authority is process-owned: a deployment may set `DOCXODUS_CHROMIUM_PATH`, but a framed request
cannot select a local executable or filesystem directory.
