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

Render artifacts require the final two profile fields. The delivery core does not silently launch
an evaluation renderer: standalone HTML/PDF stays explicitly unavailable until an epic #434
renderer adapter is supplied. The CLI likewise cannot synthesize authoritative mutation history,
so change receipts must be built through the programmatic API with a `DeliveryReceiptContext`.

The output path must not exist. Artifacts are written into a private staging directory, verified,
and then atomically renamed; the manifest is written last inside the stage.
