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

Before printing, the browser waits for explicit font, image, chart/SVG, pagination, and stable-page
tree signals. It repeats the barrier after reopening the serialized standalone document so the
checks apply to the exact DOM Chromium prints. One total deadline bounds the operation; structured
failures identify the incomplete phase and current pending resources. The end-to-end test output
includes `test-artifacts/view-artifacts.html`, which links the successful PDF/HTML plus readiness,
geometry, font-resolution, request, and failure evidence even when a later test fails. CI uploads
that directory with `if: always()`, so a later test failure does not hide the completed evidence.

## Deterministic fonts

Pass deployment-controlled font roots through `fontDirectories`; the CLI exposes the same option
as repeatable `--font-directory` flags. Roots are resolved once in caller order. Entries are
code-unit sorted and scanned in a deterministic files-before-descendants order. An exact requested
family wins, followed by the shared Docxodus substitution
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
    permittedOutputs: ["html", "pdf"],
    subsettingPermitted: true,
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
are bounded printable evidence strings. Each attestation names its permitted HTML/PDF outputs and
whether subsetting is permitted. HTML carries the verified complete data-webfont bytes; PDF fails
closed for a no-subsetting face because Chromium does not expose full-program embedding proof.
OOXML-obfuscated embedded fonts remain unsupported and are reported from package preflight rather
than silently substituted.

Default mode returns structured warnings for substitutions, missing/load-failed fonts, metric
mismatch, partial glyph coverage, synthesized faces, and unattested browser fallback.
`strictFonts: true` fails the `font_loading` phase for every non-exact, partial, synthesized, or
unverified outcome. Missing legal embedding evidence fails regardless of strictness.

Reports expose requested stacks and face attributes, selected family/face, status, source, format,
file digest/version, coverage, metric compatibility, and a canonical license-evidence identity.
The accompanying `fontIdentity` binds the resolver contract, substitution contract, and complete
resolution decision set. A separate `fontReadiness` array records the exact final CSS request key,
bounded sample commitment, and availability used for offline reopen and JS-disabled PDF parity.
Absolute font paths are excluded from generated CSS, standalone metadata,
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

Additional flags include `--document-version`, `--expected-source-digest`, `--title`,
`--unsupported-content`, `--strict-fonts`, `--browser-executable`, repeatable `--limit
name=integer`, repeatable `--font-directory`, `--font-license-attestations`, and
`--environment-attestation`. Artifact bytes are never written to stdout. Existing destinations,
input aliases, and duplicate destinations are rejected; the CLI never overwrites a file.

## Runtime boundaries

- Browser-only callers may supply their own asynchronous resolver through
  `docxodus/export-browser`; `@docxodus/export` owns filesystem discovery, licensing policy, and the
  Playwright bridge. Unattested system fonts remain `browserObserved` and cannot satisfy strict
  mode. A caller-owned browser is at most `callerAttested`, never `nodeVerified`.
- The `original` review profile remains fail-closed until issue #444 completes its projection.
- The broader generated-PDF fidelity ratchet is extended by issue #443.

Failures are `DocxodusExportError` objects with stable code, phase, remediation, safe detail, and a
structured failed report when materialization had begun.

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
cannot select a local executable, filesystem directory, or embedding-rights attestation. A
deployment may set `DOCXODUS_FONT_POLICY_PATH` to a bounded schema-v1 JSON object with
`fontDirectories` and `fontLicenseAttestations` arrays. Relative roots resolve from the policy
file's directory. The host reads the policy once and applies the same process-owned configuration
to every batch; each batch snapshots the selected font bytes before rendering.

```json
{
  "schemaVersion": 1,
  "fontDirectories": ["fonts/contract", "fonts/fallback"],
  "fontLicenseAttestations": []
}
```
