import assert from "node:assert/strict";
import { test } from "node:test";
import { convertDocxToPdf, DocxodusExportError } from "../dist/index.js";

// Invalid CSS options must fail before package parsing or browser launch.
for (const additionalCss of [null, 42, {}, "\ud800"]) {
  test(`rejects invalid additionalCss ${JSON.stringify(additionalCss)} before transfer`, async () => {
    await assert.rejects(convertDocxToPdf(Uint8Array.of(1), {
      reviewProfile: "final", commentProfile: "hidden", additionalCss,
    }), error => error instanceof DocxodusExportError
      && error.code === "invalid_argument" && error.phase === "input_validation"
      && /additionalCss/.test(error.message));
  });
}
