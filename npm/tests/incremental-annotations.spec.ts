import { test, expect, Page } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, "../../TestFiles");

function readTestFile(filename: string): Uint8Array {
  const filePath = path.join(TEST_FILES_DIR, filename);
  if (!fs.existsSync(filePath)) {
    throw new Error(`Test file not found: ${filePath}`);
  }
  return new Uint8Array(fs.readFileSync(filePath));
}

async function waitForWasm(page: Page, timeout = 30000): Promise<void> {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout,
  });
}

function formatMs(ms: number): string {
  if (ms < 1) return `${(ms * 1000).toFixed(1)}μs`;
  if (ms < 1000) return `${ms.toFixed(2)}ms`;
  return `${(ms / 1000).toFixed(2)}s`;
}

test.describe("Incremental Annotation Performance", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto("/test-harness.html");
    await waitForWasm(page);
  });

  test("benchmark incremental annotation projection vs full DOCX re-conversion", async ({
    page,
  }) => {
    const bytes = readTestFile("HC007-Test-02.docx");

    // Send doc bytes to browser
    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const results = await page.evaluate(async () => {
      const bytes = (window as any).testDocxBytes;
      const Docxodus = (window as any).Docxodus;
      const Tests = (window as any).DocxodusTests;
      const ITERATIONS = 5;

      // Step 1: Create an external annotation set from the document
      const setResult = Tests.createExternalAnnotationSet(bytes, "bench-doc");
      if (setResult.error) throw new Error(JSON.stringify(setResult.error));
      const annotationSet = setResult.annotationSet;

      // Step 2: Add several annotations to the set via text search
      // C# JSON serializer uses PascalCase; handle both casings
      const docContent =
        annotationSet.Content || annotationSet.content || "";
      const labelledText =
        annotationSet.LabelledText || annotationSet.labelledText || [];
      const searchTerms = ["the", "and", "of", "to", "in"];
      const labels: Record<string, any> = {};

      for (let i = 0; i < searchTerms.length; i++) {
        const labelId = `LABEL_${i}`;
        labels[labelId] = {
          id: labelId,
          text: searchTerms[i].toUpperCase(),
          color: `hsl(${i * 72}, 70%, 50%)`,
        };

        // Find the first occurrence via the client-side helper
        const annResult = Tests.createAnnotationFromSearch(
          `ann-${i}`,
          labelId,
          docContent,
          searchTerms[i],
          1
        );
        if (annResult.annotation) {
          labelledText.push(annResult.annotation);
        }
      }
      // Set fields in both casings so serialization works either way
      annotationSet.textLabels = labels;
      annotationSet.TextLabels = labels;
      annotationSet.labelledText = labelledText;
      annotationSet.LabelledText = labelledText;

      const annotationSetJson = JSON.stringify(annotationSet);

      // ──────────────────────────────────────────────────────────
      // Benchmark A: Full DOCX → HTML with external annotations
      //   (re-reads & re-converts the DOCX every time)
      // ──────────────────────────────────────────────────────────
      const fullConversionTimes: number[] = [];
      for (let i = 0; i < ITERATIONS; i++) {
        const start = performance.now();
        const result =
          Docxodus.DocumentConverter.ConvertDocxToHtmlWithExternalAnnotations(
            bytes,
            annotationSetJson,
            "Document",
            "docx-",
            true,
            "",
            "ext-annot-",
            0
          );
        fullConversionTimes.push(performance.now() - start);

        // Verify it produced valid output
        if (i === 0) {
          const parsed = JSON.parse(result);
          const html = parsed.Html || parsed.html;
          if (!html || html.length < 100) {
            throw new Error("Full conversion produced no HTML");
          }
        }
      }

      // ──────────────────────────────────────────────────────────
      // Benchmark B: Convert DOCX once, then project annotations
      //   onto the cached HTML string (the incremental approach)
      // ──────────────────────────────────────────────────────────

      // One-time base conversion
      const baseConvStart = performance.now();
      const baseHtml = Docxodus.DocumentConverter.ConvertDocxToHtml(bytes);
      const baseConvTime = performance.now() - baseConvStart;

      const incrementalTimes: number[] = [];
      for (let i = 0; i < ITERATIONS; i++) {
        const start = performance.now();
        const result = Docxodus.DocumentConverter.ProjectAnnotationsOntoHtml(
          baseHtml,
          annotationSetJson,
          "ext-annot-",
          0 // Above label mode
        );
        incrementalTimes.push(performance.now() - start);

        // Verify it produced valid output
        if (i === 0) {
          const parsed = JSON.parse(result);
          const html = parsed.Html || parsed.html;
          if (!html || html.length < 100) {
            throw new Error("Incremental projection produced no HTML");
          }
        }
      }

      // ──────────────────────────────────────────────────────────
      // Benchmark C: Single annotation add/remove (hot path)
      // ──────────────────────────────────────────────────────────

      // Get annotated HTML to work with
      const projResult = Docxodus.DocumentConverter.ProjectAnnotationsOntoHtml(
        baseHtml,
        annotationSetJson,
        "ext-annot-",
        0
      );
      const projParsed = JSON.parse(projResult);
      const annotatedHtml = projParsed.Html || projParsed.html;

      // Add single annotation
      const newAnn = Tests.createAnnotationFromSearch(
        "ann-new",
        "LABEL_0",
        docContent,
        "document",
        1
      );
      const newAnnJson = JSON.stringify(newAnn.annotation);
      const newLabelJson = JSON.stringify(labels["LABEL_0"]);

      const addTimes: number[] = [];
      for (let i = 0; i < ITERATIONS; i++) {
        const start = performance.now();
        Docxodus.DocumentConverter.AddAnnotationToHtml(
          annotatedHtml,
          newAnnJson,
          newLabelJson,
          "ext-annot-",
          0
        );
        addTimes.push(performance.now() - start);
      }

      // Remove single annotation
      const removeTimes: number[] = [];
      for (let i = 0; i < ITERATIONS; i++) {
        const start = performance.now();
        Docxodus.DocumentConverter.RemoveAnnotationFromHtml(
          annotatedHtml,
          "ann-0",
          "ext-annot-"
        );
        removeTimes.push(performance.now() - start);
      }

      const median = (arr: number[]) => {
        const sorted = [...arr].sort((a, b) => a - b);
        return sorted[Math.floor(sorted.length / 2)];
      };

      return {
        annotations: annotationSet.labelledText.length,
        fullConversion: {
          medianMs: median(fullConversionTimes),
          allMs: fullConversionTimes,
        },
        incrementalProjection: {
          baseConversionMs: baseConvTime,
          medianMs: median(incrementalTimes),
          allMs: incrementalTimes,
        },
        singleAdd: {
          medianMs: median(addTimes),
          allMs: addTimes,
        },
        singleRemove: {
          medianMs: median(removeTimes),
          allMs: removeTimes,
        },
      };
    });

    // ── Print results ──
    console.log("\n" + "═".repeat(70));
    console.log("INCREMENTAL ANNOTATION PERFORMANCE BENCHMARK");
    console.log("═".repeat(70));
    console.log(`Annotations:           ${results.annotations}`);
    console.log(`Iterations per test:   5`);
    console.log("─".repeat(70));

    console.log(
      `Full DOCX re-conversion:  ${formatMs(results.fullConversion.medianMs)} (median)`
    );
    console.log(
      `  runs: ${results.fullConversion.allMs.map(formatMs).join(", ")}`
    );
    console.log("");

    console.log(
      `Base HTML conversion:     ${formatMs(results.incrementalProjection.baseConversionMs)} (one-time cost)`
    );
    console.log(
      `Incremental projection:   ${formatMs(results.incrementalProjection.medianMs)} (median)`
    );
    console.log(
      `  runs: ${results.incrementalProjection.allMs.map(formatMs).join(", ")}`
    );
    console.log("");

    console.log(
      `Single add annotation:    ${formatMs(results.singleAdd.medianMs)} (median)`
    );
    console.log(
      `  runs: ${results.singleAdd.allMs.map(formatMs).join(", ")}`
    );
    console.log(
      `Single remove annotation: ${formatMs(results.singleRemove.medianMs)} (median)`
    );
    console.log(
      `  runs: ${results.singleRemove.allMs.map(formatMs).join(", ")}`
    );

    const speedup =
      results.fullConversion.medianMs /
      results.incrementalProjection.medianMs;
    const addSpeedup =
      results.fullConversion.medianMs / results.singleAdd.medianMs;
    const removeSpeedup =
      results.fullConversion.medianMs / results.singleRemove.medianMs;

    console.log("─".repeat(70));
    console.log(`VERDICT:`);
    console.log(
      `  Projection vs full re-conversion: ${speedup.toFixed(1)}x ${speedup > 1 ? "FASTER" : "SLOWER"}`
    );
    console.log(
      `  Single add vs full re-conversion:  ${addSpeedup.toFixed(1)}x ${addSpeedup > 1 ? "FASTER" : "SLOWER"}`
    );
    console.log(
      `  Single remove vs full:             ${removeSpeedup.toFixed(1)}x ${removeSpeedup > 1 ? "FASTER" : "SLOWER"}`
    );
    console.log("═".repeat(70));

    // The incremental approach must be faster than re-converting the whole DOCX
    expect(results.incrementalProjection.medianMs).toBeLessThan(
      results.fullConversion.medianMs
    );
    expect(results.singleAdd.medianMs).toBeLessThan(
      results.fullConversion.medianMs
    );
    expect(results.singleRemove.medianMs).toBeLessThan(
      results.fullConversion.medianMs
    );
  });
});
