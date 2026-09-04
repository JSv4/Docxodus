# Editor screenshots

Regenerates the pictures under `docs/images/editor/` — the README's editor images and every
screenshot in `docs/architecture/editor_ui_surface.md` — from the **shipped** surface
(`mountRibbon`, hosted by `npm/examples/editor.html`) with a real document. The document is
the NVCA Model Certificate of Incorporation (`TestFiles/NVCA-Model-COI.docx`), as it has been
for every editor screenshot, so the pictures stay comparable across releases.

From the repository root:

```bash
cd npm
npm run build && npm run pretest             # dist/wasm/ plus the staged example host
cp ../TestFiles/NVCA-Model-COI.docx dist/wasm/sample.docx
python3 -m http.server 8082 --directory dist/wasm &
node ../tools/screenshots/editor/capture.mjs http://localhost:8082/editor.html ../docs/images/editor
```

Captured, in order: `editor-overview`, `ribbon-home`, `ribbon-insert`, `ribbon-layout`,
`ribbon-review`, `ribbon-view`, `table-picker`, `ribbon-table-contextual`, `comments`,
`header-band`, `paginated`, `footnotes`, `ribbon-compact`. The script drives the surface through
the same `window.__demo` / `window.__selectTab` handles the Playwright specs use, so a capture
that fails is a surface regression, not a tooling one.

Uses the Chromium that `npx playwright install chromium` put in place for the npm tests; set
`PLAYWRIGHT_MODULE` to point at another Playwright install if needed.
