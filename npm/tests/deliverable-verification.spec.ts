import { expect, test } from "@playwright/test";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import type { DeliverableVerificationResult } from "../src/index.js";

const testDirectory = dirname(fileURLToPath(import.meta.url));

async function waitForDocxodus(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout: 30_000,
  });
}

test.describe("deliverable verification transports", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/test-harness.html");
    await waitForDocxodus(page);
  });

  test("trimmed WASM stateless and typed-session exports return canonical reports", async ({ page }) => {
    const result = await page.evaluate(async () => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const bytes = bridge.CreateBlankDocx() as Uint8Array;
      const sha256 = async (value: Uint8Array): Promise<string> => {
        const digest = await crypto.subtle.digest(
          "SHA-256",
          new Uint8Array(value).buffer,
        );
        return Array.from(new Uint8Array(digest), (byte) =>
          byte.toString(16).padStart(2, "0")
        ).join("");
      };
      const expectedDigest = await sha256(bytes);

      const directJson = (window as any).Docxodus.DocumentConverter
        .VerifyDeliverable(bytes) as string;
      const directAgain = (window as any).Docxodus.DocumentConverter
        .VerifyDeliverable(bytes) as string;
      const direct = JSON.parse(directJson);
      const compared = JSON.parse(
        (window as any).Docxodus.DocumentConverter
          .VerifyDeliverableWithBaseline(bytes, bytes) as string
      );

      const session = (window as any).Docxodus.openTypedSession(bytes, "");
      try {
        const checkpointDigest = await sha256(session.save());
        const versionBefore = session.getVersion();
        const sessionReport = session.verifyDeliverable();
        const sessionAgain = session.verifyDeliverable();
        return {
          direct,
          compared,
          directIsCanonical: directJson === directAgain,
          expectedDigest,
          sessionReport,
          checkpointDigest,
          sessionIsCanonical:
            JSON.stringify(sessionReport) === JSON.stringify(sessionAgain),
          versionBefore,
          versionAfter: session.getVersion(),
        };
      } finally {
        session.close();
      }
    });

    const direct: DeliverableVerificationResult = result.direct;
    const compared: DeliverableVerificationResult = result.compared;
    const sessionReport: DeliverableVerificationResult = result.sessionReport;

    expect(direct.schema).toBe(
      "https://docxodus.dev/schemas/verification/deliverable-verification/v1",
    );
    expect(direct.schemaVersion).toBe(1);
    expect(direct.mode).toBe("standard");
    expect(direct.decision).toMatch(/^[a-z]/);
    expect(direct.baselineCompared).toBe(false);
    expect(direct.baselinePackage).toBeNull();
    expect(direct.deliverablePackage.rawPackageBytesDigest.value).toBe(
      result.expectedDigest,
    );
    expect(result.directIsCanonical).toBe(true);

    expect(compared.baselineCompared).toBe(true);
    expect(compared.baselinePackage?.rawPackageBytesDigest.value).toBe(
      result.expectedDigest,
    );
    expect(compared.deliverablePackage.rawPackageBytesDigest.value).toBe(
      result.expectedDigest,
    );

    expect(sessionReport.schema).toBe(direct.schema);
    expect(sessionReport.schemaVersion).toBe(1);
    expect(sessionReport.baselineCompared).toBe(true);
    expect(sessionReport.baselinePackage).not.toBeNull();
    expect(sessionReport.baselinePackage?.rawPackageBytesDigest.value).toBe(
      result.expectedDigest,
    );
    expect(sessionReport.deliverablePackage.rawPackageBytesDigest.value).toBe(
      result.checkpointDigest,
    );
    expect(result.sessionIsCanonical).toBe(true);
    expect(result.versionAfter).toBe(result.versionBefore);
  });
});

test("public npm stateless API accepts an optional exact baseline", async ({ page }) => {
  const bytes = new Uint8Array(
    readFileSync(resolve(testDirectory, "../../TestFiles/HC006-Test-01.docx")),
  );
  const expectedDigest = createHash("sha256").update(bytes).digest("hex");

  await page.goto("http://localhost:8083/");
  const result = await page.evaluate(async (byteValues) => {
    // The self-contained CDN bundle re-exports the package's main API without
    // requiring a browser import map for npm's bare editor dependencies.
    const moduleUrl = "http://localhost:8083/embed.bundle.js";
    const api = await import(moduleUrl);
    await api.initialize("http://localhost:8083/wasm/");
    const input = new Uint8Array(byteValues);
    const before = Array.from(input);
    const withoutBaseline = await api.verifyDeliverable(input);
    const withBaseline = await api.verifyDeliverable(input, input);
    return {
      withoutBaseline,
      withBaseline,
      inputUnchanged: Array.from(input).every(
        (value, index) => value === before[index],
      ),
    };
  }, Array.from(bytes));

  const withoutBaseline: DeliverableVerificationResult = result.withoutBaseline;
  const withBaseline: DeliverableVerificationResult = result.withBaseline;
  expect(withoutBaseline.mode).toBe("standard");
  expect(withoutBaseline.baselineCompared).toBe(false);
  expect(withoutBaseline.deliverablePackage.rawPackageBytesDigest.value).toBe(
    expectedDigest,
  );
  expect(withBaseline.baselineCompared).toBe(true);
  expect(withBaseline.baselinePackage?.rawPackageBytesDigest.value).toBe(
    expectedDigest,
  );
  expect(withBaseline.deliverablePackage.rawPackageBytesDigest.value).toBe(
    expectedDigest,
  );
  expect(result.inputUnchanged).toBe(true);
});
