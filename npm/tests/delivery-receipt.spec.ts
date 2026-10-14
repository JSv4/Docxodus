import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import type { DeliveryReceiptVerificationResult } from "../src/index.js";

// Portable delivery-receipt verification through the browser transport (issue #520).
// Uses the vendored cross-language fixture TestFiles/Delivery/DR001-* — the same files
// the C# DCR055 pin and the Python transport test verify, so a canonical-format drift
// is caught on every side of the wire.

const testDirectory = dirname(fileURLToPath(import.meta.url));

function fixture(): { receiptJson: string; artifacts: Record<string, number[]> } {
  const testFiles = resolve(testDirectory, "../../TestFiles");
  return {
    receiptJson: readFileSync(
      resolve(testFiles, "Delivery/DR001-Receipt.json"),
      "utf8",
    ),
    artifacts: {
      "clean-docx": Array.from(
        readFileSync(resolve(testFiles, "HC001-5DayTourPlanTemplate.docx")),
      ),
      "semantic-source-to-delivered": Array.from(
        readFileSync(resolve(testFiles, "Delivery/DR001-Semantic.json")),
      ),
    },
  };
}

test("public npm API verifies the vendored receipt and detects tamper", async ({ page }) => {
  await page.goto("http://localhost:8083/");
  const result = await page.evaluate(async (input) => {
    const moduleUrl = "http://localhost:8083/embed.bundle.js";
    const api = await import(moduleUrl);
    await api.initialize("http://localhost:8083/wasm/");
    const artifacts = Object.fromEntries(
      Object.entries(input.artifacts).map(([id, bytes]) => [
        id,
        new Uint8Array(bytes as number[]),
      ]),
    );
    const verified = await api.verifyDeliveryReceipt(input.receiptJson, artifacts);

    const tampered = { ...artifacts };
    const cloned = new Uint8Array(tampered["clean-docx"]);
    cloned[cloned.length - 1] ^= 0xff;
    tampered["clean-docx"] = cloned;
    const rejected = await api.verifyDeliveryReceipt(input.receiptJson, tampered);

    const bare = await api.verifyDeliveryReceipt(input.receiptJson);
    const malformed = await api.verifyDeliveryReceipt('{"nope": true}');
    return { verified, rejected, bare, malformed };
  }, fixture());

  const verified: DeliveryReceiptVerificationResult = result.verified;
  expect(verified.isValid).toBe(true);
  expect(verified.receiptDigestValid).toBe(true);
  expect(verified.contractValid).toBe(true);
  expect(
    Object.fromEntries(verified.artifacts.map((a) => [a.artifactId, a.status])),
  ).toEqual({
    "clean-docx": "verified",
    "semantic-source-to-delivered": "verified",
  });

  const rejected: DeliveryReceiptVerificationResult = result.rejected;
  expect(rejected.isValid).toBe(false);
  expect(rejected.receiptDigestValid).toBe(true);
  expect(
    rejected.artifacts.find((a) => a.artifactId === "clean-docx")?.status,
  ).toBe("digest_mismatch");

  const bare: DeliveryReceiptVerificationResult = result.bare;
  expect(bare.isValid).toBe(false);
  expect(bare.artifacts.every((a) => a.status === "missing")).toBe(true);

  const malformed: DeliveryReceiptVerificationResult = result.malformed;
  expect(malformed.isValid).toBe(false);
  expect(malformed.receiptDigestValid).toBe(false);
  expect(malformed.findings.length).toBeGreaterThan(0);
});
