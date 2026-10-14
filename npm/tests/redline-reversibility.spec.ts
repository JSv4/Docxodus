import { expect, test, Page } from "@playwright/test";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import type { RedlineReversibilityProof } from "../src/index.js";

// Transport-seam coverage for the redline reversibility proof. The proof engine is
// covered by the .NET suite; this spec asserts the trimmed WASM export and the worker
// path hand JavaScript callers the same canonical schema-v1 document.

const testDirectory = dirname(fileURLToPath(import.meta.url));
const TEST_FILES_DIR = join(testDirectory, "../../TestFiles");

const SCHEMA =
  "https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1";

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(readFileSync(join(TEST_FILES_DIR, relativePath)));
}

function sha256(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

async function waitForDocxodus(page: Page): Promise<void> {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout: 30_000,
  });
}

test.describe("redline reversibility proof transports", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/test-harness.html");
    await waitForDocxodus(page);
  });

  test("the trimmed WASM export returns a canonical proof bound to its three inputs", async ({
    page,
  }) => {
    const baseline = readTestFile("WC/WC001-Digits.docx");
    const intendedFinal = readTestFile("WC/WC001-Digits-Mod.docx");

    const result = await page.evaluate(
      ([b, f]) => {
        const baselineBytes = new Uint8Array(b);
        const finalBytes = new Uint8Array(f);
        const converter = (window as any).Docxodus.DocumentConverter;
        const redline = (window as any).DocxodusTests.docxDiffCompare(
          baselineBytes,
          finalBytes
        ).docxBytes as Uint8Array;

        const json = converter.ProveRedlineReversibility(
          baselineBytes,
          finalBytes,
          redline
        ) as string;
        const again = converter.ProveRedlineReversibility(
          baselineBytes,
          finalBytes,
          redline
        ) as string;

        return { json, deterministic: json === again, redline: Array.from(redline) };
      },
      [Array.from(baseline), Array.from(intendedFinal)]
    );

    expect(result.deterministic).toBe(true);

    const proof = JSON.parse(result.json) as RedlineReversibilityProof;
    expect(proof.schema).toBe(SCHEMA);
    expect(proof.schemaVersion).toBe(1);
    expect(proof.baselinePackage.rawPackageBytesDigest.value).toBe(sha256(baseline));
    expect(proof.intendedFinalPackage.rawPackageBytesDigest.value).toBe(
      sha256(intendedFinal)
    );
    expect(proof.redlinePackage.rawPackageBytesDigest.value).toBe(
      sha256(new Uint8Array(result.redline))
    );
    expect(proof.revisionClassifications.length).toBeGreaterThan(0);
    // Nothing can be conflicted: there is no pre-existing review state to conflict with.
    expect(
      proof.revisionClassifications.every((item) => item.disposition !== "conflicted")
    ).toBe(true);
    expect(proof.acceptToFinal).not.toBeNull();
    expect(proof.rejectToBaseline).not.toBeNull();
    expect(proof.acceptToFinal!.direction).toBe("acceptToFinal");
    expect(proof.rejectToBaseline!.direction).toBe("rejectToBaseline");
    // Each path states the document it must reproduce.
    expect(proof.acceptToFinal!.expectedPackage.rawPackageBytesDigest.value).toBe(
      sha256(intendedFinal)
    );
    expect(proof.rejectToBaseline!.expectedPackage.rawPackageBytesDigest.value).toBe(
      sha256(baseline)
    );
  });

  test("a malformed package is a typed finding, and neither path is attempted", async ({
    page,
  }) => {
    const baseline = readTestFile("WC/WC001-Digits.docx");

    const json = await page.evaluate(([b]) => {
      const bytes = new Uint8Array(b);
      return (window as any).Docxodus.DocumentConverter.ProveRedlineReversibility(
        bytes,
        bytes,
        new Uint8Array([1, 2, 3])
      ) as string;
    }, [Array.from(baseline)]);

    const proof = JSON.parse(json) as RedlineReversibilityProof;
    expect(proof.success).toBe(false);
    expect(proof.findings.some((finding) => finding.severity === "error")).toBe(true);
    // Fail-closed: no partial path result can be misread as evidence.
    expect(proof.acceptToFinal).toBeNull();
    expect(proof.rejectToBaseline).toBeNull();
  });
});

test.describe("redline reversibility proof over the worker", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/worker-test-harness.html");
    await page.waitForFunction(
      () => (window as any).DocxodusWorkerTests !== undefined,
      { timeout: 10_000 }
    );
  });

  test("the worker path hands back the same proof document as the direct export", async ({
    page,
  }) => {
    const baseline = readTestFile("WC/WC001-Digits.docx");
    const intendedFinal = readTestFile("WC/WC001-Digits-Mod.docx");

    const result = await page.evaluate(
      async ([b, f]) => {
        await (window as any).createDocxodusWorker();
        const worker = (window as any).DocxodusWorker;
        const baselineBytes = new Uint8Array(b);
        const finalBytes = new Uint8Array(f);
        const redline = (await worker.compareDocuments(
          baselineBytes,
          finalBytes
        )) as Uint8Array;

        // The proxy clones before transfer, so the same arrays are reusable.
        const proof = await worker.proveRedlineReversibility(
          baselineBytes,
          finalBytes,
          redline
        );
        const again = await worker.proveRedlineReversibility(
          baselineBytes,
          finalBytes,
          redline
        );

        return {
          proof,
          deterministic: JSON.stringify(proof) === JSON.stringify(again),
          redline: Array.from(redline),
        };
      },
      [Array.from(baseline), Array.from(intendedFinal)]
    );

    expect(result.deterministic).toBe(true);

    const proof = result.proof as RedlineReversibilityProof;
    expect(proof.schema).toBe(SCHEMA);
    expect(proof.schemaVersion).toBe(1);
    expect(proof.baselinePackage.rawPackageBytesDigest.value).toBe(sha256(baseline));
    expect(proof.intendedFinalPackage.rawPackageBytesDigest.value).toBe(
      sha256(intendedFinal)
    );
    expect(proof.redlinePackage.rawPackageBytesDigest.value).toBe(
      sha256(new Uint8Array(result.redline))
    );
    expect(proof.revisionClassifications.length).toBeGreaterThan(0);
    expect(proof.acceptToFinal).not.toBeNull();
    expect(proof.rejectToBaseline).not.toBeNull();
    expect(proof.acceptToFinal!.expectedPackage.rawPackageBytesDigest.value).toBe(
      sha256(intendedFinal)
    );
    expect(proof.rejectToBaseline!.expectedPackage.rawPackageBytesDigest.value).toBe(
      sha256(baseline)
    );
  });
});
