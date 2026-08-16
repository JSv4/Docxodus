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

## Review and comment profiles

`reviewProfile` selects one deterministic view of native Word revisions:

| Value | Rendered result |
|---|---|
| `final` | Accepted result: insertions and move destinations use final formatting; deletions and move sources are hidden. |
| `original` | Rejected result: deletions and move sources use original formatting; insertions and move destinations are hidden. |
| `markup` | Review result: supported insertions, deletions, moves, and formatting changes remain visibly marked with their available author/date metadata. |

`commentProfile` is orthogonal to that view:

| Value | Rendered result |
|---|---|
| `hidden` | Comment highlights, markers, and bodies are omitted without hiding the commented document text. |
| `inline` | A print-visible comment thread is placed beside its first range/reference in story order. |
| `endnotes` | Print-visible comment threads are collected in an ordered document-end section with reference links. |
| `margin` | Print-visible comment threads are placed in the owning page's margin and associated with their ranges. |

Visible modes retain each range and body, ordered replies, author and date when present, and a
printable `open`, `resolved`, or `unknown` state. Missing extended-comment metadata is represented
as `unknown`, never inferred. This contract applies in the body, headers, footers, footnotes, and
endnotes and is shared by the browser API, Node API, framed host, and CLI. Orphaned and cyclic
reply-parent chains remain visible as auditable independent roots, carry `data-comment-topology`,
and produce `comment_parent_orphaned` or `comment_parent_cycle` warnings; strict policy rejects
the same malformed topology. Hidden presentation still emits the complete report diagnostics while
omitting comment bodies and markers from the published HTML/PDF.

The input snapshot is immutable. `renderReport.source` always records its exact raw-package digest,
byte length, and document version. `final` and `original` project an isolated package and record the
projected digest and length as `derivedProfileSource`; `markup` renders the unchanged source and
omits that field. When an upstream policy engine already owns the exact final/original package, set
`reviewProfileAlreadyApplied: true` to preserve it byte-for-byte. Export then verifies that no
tracked revision remains, omits `derivedProfileSource`, and rejects the option for `markup`. The
requested profiles participate in the layout digest and renderer fingerprint.

The default `unsupportedContent: "warn"` policy records a structured diagnostic naming every
unsupported revision, comment, or story family and its owning package part, and continues only when
the limitation remains explicit. `"strict"` rejects before publishing output. Under either policy,
`final` and `original` fail if any tracked-change marker remains after projection; neither profile
silently accepts, rejects, removes, or relabels an unsupported edit.

Caller `Uint8Array` values are synchronously copied. A caller-supplied Playwright Chromium
`Browser` remains caller-owned; Docxodus closes only the fresh context it creates. Every runtime
asset is length/hash checked against the public manifest and served at a routed `.invalid` origin.
Any request outside that closed graph fails the operation.

PDF printing enables backgrounds, uses CSS page sizes, applies zero browser margins, preserves DOM
text/links/vector content, and verifies every page MediaBox/CropBox origin and dimensions with a
real PDF parser. Mixed portrait/landscape and Letter/A4 sections retain their per-page CSS size;
screen zoom/transforms are removed by the print contract and cannot change physical PDF geometry.
Reports record the exact PDF SHA-256 and Chromium's volatile metadata; PDF byte identity is
intentionally not claimed across runs.

Before printing, the browser waits for explicit font, image, chart/SVG, pagination, and stable-page
tree signals. It repeats the barrier after reopening the serialized standalone document so the
checks apply to the exact DOM Chromium prints. One total deadline bounds the operation; structured
failures identify the incomplete phase and current pending resources. The end-to-end test output
includes `test-artifacts/view-artifacts.html`, which links the successful PDF/HTML plus readiness,
geometry, font-resolution, request, and failure evidence even when a later test fails. Its profile
matrix covers `final`/`hidden`, `original`/`inline`, `markup`/`endnotes`, and `markup`/`margin`.
Successful rows retain the source DOCX, standalone HTML, PDF, screenshot, PageMap, render report,
HTML/PDF extracted text, and profile-comparison summary; the strict-policy row retains its request
and structured failed report. CI uploads that directory with `if: always()`, so a later test failure
does not hide completed evidence.

## Deterministic fonts

Pass deployment-controlled font roots through `fontDirectories`; the CLI exposes the same option
as repeatable `--font-directory` flags. Roots are resolved once in caller order and scanned in
stable lexical order. An exact requested family wins, followed by the shared Docxodus substitution
contract. Earlier roots win across root boundaries, while conflicting files for the same family,
style, weight, and stretch inside one root are rejected. Byte-identical files are deduplicated.
Symlinks and non-regular files are never followed.

```js
const result = await convertDocxToPdf(source, {
  reviewProfile: "final",
  commentProfile: "hidden",
  fontDirectories: ["/opt/contract-fonts", "/opt/fallback-fonts"],
  fontLicenseAttestations: [{
    schemaVersion: 1,
    usage: "standalone-document-font-embedding",
    fileSha256: "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
    embeddingPermitted: true,
    basis: "Vendor webfont license, order 2026-1042",
    attester: "release-engineering@example.invalid",
  }],
  strictFonts: true,
});
```

TTF/OTF embedding rights are read from `OS/2.fsType`. An explicit restricted or bitmap-only value
always fails closed. WOFF/WOFF2, or any file whose rights cannot be derived, requires an affirmative
schema-v1 attestation with usage `standalone-document-font-embedding`, bound to the exact lowercase
SHA-256; an attestation cannot override an explicit restriction. `basis` and optional `attester`
are bounded printable evidence strings. No subsetting is performed. OOXML-obfuscated embedded fonts
remain unsupported and are reported from package preflight rather than silently substituted.

Default mode returns structured warnings for substitutions, missing/load-failed fonts, metric
mismatch, partial glyph coverage, synthesized faces, and unattested browser fallback.
`strictFonts: true` fails the `font_loading` phase for every non-exact, partial, synthesized, or
unverified outcome. Missing legal embedding evidence fails regardless of strictness.

Reports expose requested stacks and face attributes, selected family/face, status, source, format,
file digest/version, coverage, metric compatibility, browser-fallback availability, and a canonical
license-evidence identity. A configured resolver miss remains `missing` even when Chromium can paint
the sample through an unidentified generic fallback; that fallback observation is recorded
separately and never upgrades the resolution status.
The accompanying `fontIdentity` binds the resolver contract, substitution contract, and complete
resolution decision set. Absolute font paths are excluded from generated CSS, standalone metadata,
reports, diagnostics, and renderer fingerprints. Font bytes appear only in the standalone
document's generated data-webfont rules; base64 bytes never enter reports or fingerprints. Use the
reported digests to correlate deployment files without disclosing their locations.

## CLI

Profiles are explicit:

```console
docxodus convert contract.docx --to pdf --output contract.pdf \
  --review-profile final --comments endnotes \
  --timeout 120000 --report contract.render.json \
  --page-map contract.pages.json
```

The same vocabulary produces a rejected-result PDF or a review PDF without a second CLI policy:

```console
docxodus convert contract.docx --to pdf --output contract-original.pdf \
  --review-profile original --comments inline \
  --report contract-original.render.json --page-map contract-original.pages.json

docxodus convert contract.docx --to pdf --output contract-review.pdf \
  --review-profile markup --comments margin --unsupported-content strict \
  --report contract-review.render.json --page-map contract-review.pages.json
```

Additional flags include `--document-version`, `--expected-source-digest`,
`--review-profile-already-applied`, `--title`,
`--unsupported-content`, `--strict-fonts`, `--browser-executable`, repeatable `--limit
name=integer`, repeatable `--font-directory`, `--font-license-attestations`, and
`--environment-attestation`. Artifact bytes are never written to stdout. Existing destinations,
input aliases, and duplicate destinations are rejected; the CLI never overwrites a file.

## Runtime boundaries

- Browser-only callers may supply their own asynchronous resolver through
  `docxodus/export-browser`; `@docxodus/export` owns filesystem discovery, licensing policy, and the
  Playwright bridge. Unattested system fonts remain `browserObserved` and cannot satisfy strict
  mode. A caller-owned browser is at most `callerAttested`, never `nodeVerified`.
- The broader generated-PDF fidelity ratchet is extended by issue #443.

Failures are `DocxodusExportError` objects with stable code, phase, remediation, safe detail, and a
structured failed report when materialization had begun.

## Framed host

`docxodus-export-host` is the non-shell integration boundary used by delivery adapters. It accepts
exactly one length-prefixed JSON envelope (four-byte, unsigned big-endian length), returns one
length-prefixed response, rejects duplicate batch IDs and unknown fields, and encodes PDF payloads
as canonical base64. One host-owned Chromium browser is reused across the envelope while every
batch receives a fresh isolated context. Executable and font-directory authority is deliberately
process-owned: a deployment may set `DOCXODUS_CHROMIUM_PATH` and
`DOCXODUS_FONT_POLICY_PATH`, but a framed request cannot select a local executable, filesystem
directory, or embedding-rights attestation. The font-policy file is a bounded schema-v1 JSON object
with `fontDirectories` and `fontLicenseAttestations` arrays; relative roots resolve from the policy
file's directory. The host reads it once and applies the same process-owned policy configuration to
every batch; each batch snapshots the selected font bytes before rendering.

```json
{
  "schemaVersion": 1,
  "fontDirectories": ["fonts/contract", "fonts/fallback"],
  "fontLicenseAttestations": []
}
```
