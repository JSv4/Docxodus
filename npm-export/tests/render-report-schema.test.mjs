import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { before, test } from "node:test";
import Ajv2020 from "ajv/dist/2020.js";
import { DocxodusExportError, renderDocxArtifacts } from "../dist/index.js";
import { generateBrowserCleanupDocx } from "./mixed-section-fixture.mjs";

const schema = JSON.parse(await readFile(
  new URL("../../docs/schemas/render-report-v2.schema.json", import.meta.url), "utf8"));
const validate = new Ajv2020({ strict: false, allErrors: true }).compile(schema);
const reports = {};

before(async () => {
  const source = generateBrowserCleanupDocx();
  const options = {
    reviewProfile: "markup", commentProfile: "hidden", timeoutMs: 60_000,
    ...(process.env.DOCXODUS_CHROMIUM_PATH
      ? { browserExecutablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
  };
  reports.complete = (await renderDocxArtifacts(source, {
    ...options, outputs: ["html"],
  })).renderReport;
  await assert.rejects(renderDocxArtifacts(source, {
    ...options, outputs: ["pdf"], limits: { pdfOutputBytes: 1 },
  }), error => {
    assert.ok(error instanceof DocxodusExportError);
    assert.equal(error.code, "resource_limit");
    assert.equal(error.phase, "pdf_print");
    assert.equal(error.report?.status, "failed");
    reports.failed = error.report;
    return true;
  });
}, { timeout: 120_000 });

for (const status of ["complete", "failed"]) {
  test(`v2 schema binds ${status} font observations and rejects invented file evidence`, () => {
    const report = structuredClone(reports[status]);
    report.fontObservation = { status: 'observed', method: 'chromium-cdp', media: 'print',
      fileIdentity: 'unavailable', uses: [{ textDigest: 'a'.repeat(64), textCodePointCount: 1,
        faces: [{ family: 'Actual Face', isCustomFont: false, glyphCount: 1 }] }] };
    report.fontObservationDigest = 'b'.repeat(64);
    assert.equal(validate(report), true, JSON.stringify(validate.errors));
    for (const mutate of [
      r => { delete r.fontObservationDigest; },
      r => { delete r.fontObservation; },
      r => { r.fontObservation.uses[0].faces[0].fileSha256 = 'c'.repeat(64); },
      r => { r.fontObservation.uses[0].faces[0].glyphCount = 0; },
      r => { r.fontObservation.uses[0].faces[0].family = '\u0000bad'; },
    ]) {
      const invalid = structuredClone(report); mutate(invalid);
      assert.equal(validate(invalid), false);
    }
    report.fontObservation = { status: 'unavailable', reason: 'browser_api_unavailable' };
    delete report.fontObservationDigest;
    assert.equal(validate(report), true);
    report.fontObservation = { status: 'unavailable', reason: 'diagnostic_limit' };
    report.fontObservationDigest = 'b'.repeat(64);
    assert.equal(validate(report), true, JSON.stringify(validate.errors));
    delete report.fontObservationDigest;
    delete report.fontObservation;
    assert.equal(validate(report), true, 'legacy reports remain valid');
  });

  test(`published v2 schema accepts an actual ${status} report`, () => {
    assert.equal(reports[status].schema, "https://docxodus.dev/schemas/render/render-report/v2");
    assert.equal(reports[status].schemaVersion, 2);
    assert.equal(validate(reports[status]), true, JSON.stringify(validate.errors?.slice(0, 5)));
  });

  for (const [uriVersion, numberVersion] of [[1, 1], [1, 2], [2, 1]]) {
    test(`v2 schema rejects ${status} report identity v${uriVersion}/${numberVersion}`, () => {
      const report = structuredClone(reports[status]);
      report.schema = `https://docxodus.dev/schemas/render/render-report/v${uriVersion}`;
      report.schemaVersion = numberVersion;
      assert.equal(validate(report), false);
    });
  }
}
