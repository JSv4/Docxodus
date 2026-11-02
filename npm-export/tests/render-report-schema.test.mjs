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
