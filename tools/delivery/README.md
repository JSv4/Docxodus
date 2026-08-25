# docxodus-deliver

`docxodus-deliver` builds a hash-addressed, independently verifiable delivery directory from a
named baseline DOCX and working DOCX. It is a thin CLI over the same `DeliveryBundleService` used
by the .NET programmatic API and `docxodus_deliver` MCP tool.

```console
docxodus-deliver baseline.docx working.docx delivery \
  --baseline-version=0 --final-version=1 --final-name=final \
  --pre-existing=preserve --generated=accept \
  --artifact=final:final-docx:required \
  --artifact=semantic:semantic-delta:required \
  --artifact=validation:validation-report:required
```

Each repeatable `--artifact` value has this form:

```text
id:kind:requiredness[:review-profile:comment-profile]
```

Render artifacts require the final two profile fields. Production HTML/PDF/PageMap/report output
uses the epic #434 framed host and must be configured with absolute paths:

```console
docxodus-deliver baseline.docx working.docx delivery \
  ... \
  --node-executable=/opt/node/bin/node \
  --export-host=/opt/docxodus/export/dist/host.js \
  --chromium-executable=/opt/chromium/chrome \
  --artifact=html:standalone-html:required:final:endnotes
```

`DOCXODUS_NODE_PATH`, `DOCXODUS_EXPORT_HOST_PATH`, and optional `DOCXODUS_CHROMIUM_PATH`
provide process-owned equivalents. The adapter never searches PATH or invokes a shell. Each exact
source/review/comment cohort is materialized once, with PageMap and render-report sidecars added
automatically. `--render-timeout=MS`, `--unsupported-content=warn|strict`, and `--strict-fonts`
control the bounded render. Without renderer configuration, render outputs are explicitly
unavailable. The CLI cannot synthesize authoritative mutation history, so change receipts must be
built through the programmatic API with a `DeliveryReceiptContext`.

The CLI snapshots each input twice through one read-only handle and rejects a changing file. The
output path must not exist. Artifacts are written into a private staging directory, verified, and
then atomically renamed; the manifest is written last inside the stage. Incomplete output is
retained only when `--return-incomplete` explicitly requests diagnostic publication.
