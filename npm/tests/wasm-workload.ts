// The representative browser workload behind two build-configuration specs:
//   - wasm-steady-state.spec.ts times it (median/p95 after warmup) against the bundle in
//     dist/wasm, so every runtime configuration (interpreter, jiterpreter knobs, AOT) is
//     scored the same way;
//   - aot-profile-record.spec.ts drives it under the Mono AOT profiler, so the recorded
//     profile (wasm/DocxodusWasm/docxodus.aotprofile) is by construction the code these
//     operations execute.
// One definition, so what is measured is what the profile covers. It exercises the three
// things the browser package is used for: DocxDiff compare (a small pair and a heavyweight
// legal document against an edited variant of itself), DOCX→HTML conversion, and the
// per-mutation editor refresh (ReplaceText + single-block re-render on a live session).
//
// runWasmWorkload executes INSIDE the page (page.evaluate serialises it), so it must stay
// self-contained: no captured module-scope values, only the window.Docxodus bridge.
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const TEST_FILES_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../TestFiles');

/** Fixtures, relative to TestFiles/. */
export const WORKLOAD_FIXTURES = {
  /** Two 11 KB revisions of one document — the "typical" compare. */
  smallLeft: 'WC/WC001-Digits.docx',
  smallRight: 'WC/WC001-Digits-Mod.docx',
  /** A 147 KB legal form (footnotes, ~390 bookmarks, cross-reference fields): the heavy end. */
  heavy: 'NVCA-Model-COI.docx',
  /** The editor benchmark document (editor-latency-bench.spec.ts). */
  editorDoc: 'HC031-Complicated-Document.docx',
} as const;

export interface WorkloadIterations {
  compareSmall: number;
  compareHeavy: number;
  convert: number;
  convertHeavy: number;
  sessionRefresh: number;
}

export interface WorkloadInput {
  smallLeft: number[];
  smallRight: number[];
  heavy: number[];
  editorDoc: number[];
  /** Untimed runs of each operation before the timed ones. */
  warmup: number;
  iterations: WorkloadIterations;
  /** Also return the generated heavy-variant bytes (for an out-of-browser counterpart run). */
  returnVariant?: boolean;
}

export interface OpTiming {
  op: string;
  n: number;
  minMs: number;
  medianMs: number;
  p95Ms: number;
  meanMs: number;
}

export interface WorkloadResult {
  timings: OpTiming[];
  /** Size of the last output of each op; 0 means the bridge reported an error. */
  outputSizes: Record<string, number>;
  variantEdits: { replaced: number; inserted: number; deleted: number; failed: number };
  variant?: number[];
}

export function loadWorkloadInput(
  options: { warmup: number; iterations: WorkloadIterations; returnVariant?: boolean },
): WorkloadInput {
  const read = (rel: string) => Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, rel)));
  return {
    smallLeft: read(WORKLOAD_FIXTURES.smallLeft),
    smallRight: read(WORKLOAD_FIXTURES.smallRight),
    heavy: read(WORKLOAD_FIXTURES.heavy),
    editorDoc: read(WORKLOAD_FIXTURES.editorDoc),
    ...options,
  };
}

/** Runs in the browser. */
export function runWasmWorkload(input: WorkloadInput): WorkloadResult {
  const D = (window as any).Docxodus;
  const S = D.DocxSessionBridge;
  const toBytes = (a: number[]) => new Uint8Array(a);
  const smallLeft = toBytes(input.smallLeft);
  const smallRight = toBytes(input.smallRight);
  const heavy = toBytes(input.heavy);
  const editorDoc = toBytes(input.editorDoc);

  const ok = (json: string) => {
    try { return JSON.parse(json).success === true; } catch { return false; }
  };
  // Body paragraphs in document order (the anchorIndex is keyed by id, not ordered).
  const bodyParagraphs = (handle: number): { id: string; kind: string }[] => {
    const proj = JSON.parse(S.Project(handle));
    return (Object.entries(proj.anchorIndex) as [string, any][])
      .filter(([, t]) => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
      .map(([id, t]) => ({ id, kind: t.kind as string, idx: proj.markdown.indexOf('{#' + id + '}') }))
      .filter((x) => x.idx >= 0)
      .sort((a, b) => a.idx - b.idx)
      .map(({ id, kind }) => ({ id, kind }));
  };

  // The heavy pair: the legal document against a deterministic edit of itself — text
  // replaced in every 10th paragraph, a paragraph inserted after every 17th, every 23rd
  // deleted — so the compare has real alignment work at every level, not an identity
  // fast path. Individual edits may be refused (a paragraph that is only a field, say);
  // that is counted, not fatal.
  const variantEdits = { replaced: 0, inserted: 0, deleted: 0, failed: 0 };
  let variant: Uint8Array;
  {
    const h = S.OpenSession(heavy, '');
    try {
      const apply = (r: string, bucket: 'replaced' | 'inserted' | 'deleted') => {
        if (ok(r)) variantEdits[bucket]++; else variantEdits.failed++;
      };
      bodyParagraphs(h).forEach(({ id }, i) => {
        if (i % 10 === 3) {
          apply(S.ReplaceText(h, id, 'Revised paragraph ' + i + ' for the steady-state workload.'), 'replaced');
        } else if (i % 17 === 5) {
          apply(S.InsertParagraph(h, id, 'after', 'Inserted paragraph ' + i + ' for the steady-state workload.'), 'inserted');
        } else if (i % 23 === 7) {
          apply(S.DeleteBlock(h, id), 'deleted');
        }
      });
      variant = S.Save(h);
    } finally {
      S.CloseSession(h);
    }
  }

  const timings: OpTiming[] = [];
  const outputSizes: Record<string, number> = {};
  const measure = (op: string, n: number, fn: () => number) => {
    for (let i = 0; i < input.warmup; i++) fn();
    const samples: number[] = [];
    for (let i = 0; i < n; i++) {
      const t0 = performance.now();
      outputSizes[op] = fn();
      samples.push(performance.now() - t0);
    }
    samples.sort((a, b) => a - b);
    const rank = (p: number) => samples[Math.max(0, Math.ceil(p * samples.length) - 1)];
    timings.push({
      op, n,
      minMs: samples[0],
      medianMs: rank(0.5),
      p95Ms: rank(0.95),
      meanMs: samples.reduce((a, b) => a + b, 0) / samples.length,
    });
  };

  const it = input.iterations;
  measure('diff.compare.small', it.compareSmall,
    () => D.DocxDiffBridge.Compare(smallLeft, smallRight, '').length);
  measure('diff.compare.heavy', it.compareHeavy,
    () => D.DocxDiffBridge.Compare(heavy, variant, '').length);
  measure('html.convert', it.convert,
    () => D.DocumentConverter.ConvertDocxToHtml(editorDoc).length);
  measure('html.convert.heavy', it.convertHeavy,
    () => D.DocumentConverter.ConvertDocxToHtml(heavy).length);
  {
    const h = S.OpenSession(editorDoc, '');
    try {
      const targets = bodyParagraphs(h).filter((p) => p.kind === 'p').slice(0, 5).map((p) => p.id);
      let i = 0;
      measure('session.refresh', it.sessionRefresh, () => {
        const a = targets[i++ % targets.length];
        const r = S.ReplaceText(h, a, 'Steady-state text ' + i);
        if (!ok(r)) throw new Error('ReplaceText failed: ' + r);
        const html: string = S.RenderBlockHtml(h, a, 'docx-', true);
        return html.startsWith('<') ? html.length : 0;
      });
    } finally {
      S.CloseSession(h);
    }
  }

  return {
    timings,
    outputSizes,
    variantEdits,
    variant: input.returnVariant ? Array.from(variant) : undefined,
  };
}
