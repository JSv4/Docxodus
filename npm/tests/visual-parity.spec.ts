import { test, expect, type Page } from '@playwright/test';
import { execFileSync, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  GATING_DISPOSITION_KINDS,
  REQUIRED_VISUAL_CATEGORIES,
  VISUAL_DISPOSITION_KINDS,
  VISUAL_PARITY_CORPUS,
  type VisualCorpusEntry,
  type VisualDisposition,
} from './visual-parity/corpus.js';
import { compareImages, VISUAL_THRESHOLDS, type PageMetrics, type VisualSeverity } from './visual-parity/metrics.js';
import {
  FONT_CONTRACT,
  FONT_CONTRACT_FILE,
  PROBE_TEXT,
  assertFontContract,
  fontContractReport,
  generateFontProbeDocx,
  probeLineCountsFromPdfText,
  probeMarker,
} from './visual-parity/font-contract.js';
import {
  assertLibreOfficeContract,
  popplerVersionOutput,
} from './visual-parity/environment-contract.js';
import { decodePng, encodePng } from './visual-parity/png.js';
import {
  RATCHET_RECORD_FILE,
  assertRecordUpdateProvenance,
  buildRecord,
  compareToRecord,
  readRecord,
  serializeRecord,
} from './visual-parity/ratchet.js';
import { readWordReference } from './visual-parity/word-reference.js';

test.skip(process.env.DOCXODUS_VISUAL_PARITY !== '1',
  'set DOCXODUS_VISUAL_PARITY=1 on a host with libreoffice and pdftoppm');

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '../..');
const severityOrder: VisualSeverity[] = ['close', 'minor', 'major', 'severe'];

interface PageResult extends PageMetrics {
  page: number;
  artifacts: { docxodus: string; libreoffice: string; overlay: string };
  artifactSha256: { docxodus: string; libreoffice: string; overlay: string };
}

interface CaseResult {
  id: string;
  path: string;
  categories: string[];
  rationale: string;
  disposition: VisualDisposition;
  docxodusPages: number;
  libreofficePages: number;
  pageCountDelta: number;
  severity: VisualSeverity;
  pages: PageResult[];
  error?: string;
}

/**
 * A severe result gates a strict run only when its attribution says the renderer owns it:
 * an established renderer bug, an untriaged discrepancy, or a case that failed to convert
 * at all. Environment deltas and documented LibreOffice deviations are tracked but do not
 * fail the gate — that separation is the point of the disposition field.
 */
function gatesStrictRun(result: CaseResult): boolean {
  if (result.severity !== 'severe') return false;
  return result.error !== undefined || GATING_DISPOSITION_KINDS.includes(result.disposition.kind);
}

function commandVersion(command: string, args: string[]): string {
  const result = spawnSync(command, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  if (result.error) return String(result.error);
  return `${result.stdout ?? ''}\n${result.stderr ?? ''}`.trim();
}

function sha256(path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function assertTrackedCorpus(entries: VisualCorpusEntry[]): void {
  const covered = new Set(entries.flatMap(entry => entry.categories));
  const missing = REQUIRED_VISUAL_CATEGORIES.filter(category => !covered.has(category));
  if (missing.length) throw new Error(`Visual corpus misses required categories: ${missing.join(', ')}`);

  for (const entry of entries) {
    if (!VISUAL_DISPOSITION_KINDS.includes(entry.disposition.kind)) {
      throw new Error(`${entry.id} has an unknown disposition kind: ${entry.disposition.kind}`);
    }
    if (!entry.disposition.rationale.trim()) {
      throw new Error(`${entry.id} disposition needs a rationale: a disposition is a reviewed claim`);
    }
    if (isAbsolute(entry.path) || relative(repoRoot, resolve(repoRoot, entry.path)).startsWith('..')) {
      throw new Error(`${entry.id} escapes the repository: ${entry.path}`);
    }
    execFileSync('git', ['ls-files', '--error-unmatch', entry.path], {
      cwd: repoRoot,
      stdio: 'pipe',
    });
    const fixturePath = resolve(repoRoot, entry.path);
    if (!existsSync(fixturePath)) {
      throw new Error(`${entry.id} references a missing tracked fixture: ${entry.path}`);
    }
    if (!lstatSync(fixturePath).isFile()) {
      throw new Error(`${entry.id} fixture must be a regular file, not a symlink: ${entry.path}`);
    }

    // A tracked pathname can still contain locally replaced bytes. Compare the worktree blob to
    // HEAD so an ignored/untracked corpus cannot be smuggled in behind a known filename.
    const committedBlob = execFileSync('git', ['rev-parse', `HEAD:${entry.path}`], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    }).trim();
    const worktreeBlob = execFileSync('git', ['hash-object', fixturePath], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    }).trim();
    if (worktreeBlob !== committedBlob) {
      throw new Error(`${entry.id} fixture differs from HEAD: ${entry.path}`);
    }
  }
}

function selectedCorpus(): VisualCorpusEntry[] {
  const filter = process.env.DOCXODUS_VISUAL_PARITY_FILTER;
  if (!filter) return VISUAL_PARITY_CORPUS;
  const requested = new Set(filter.split(',').map(value => value.trim()).filter(Boolean));
  const selected = VISUAL_PARITY_CORPUS.filter(entry => requested.has(entry.id));
  const unknown = [...requested].filter(id => !VISUAL_PARITY_CORPUS.some(entry => entry.id === id));
  if (unknown.length) throw new Error(`Unknown visual parity case(s): ${unknown.join(', ')}`);
  if (!selected.length) throw new Error('Visual parity filter selected no cases');
  return selected;
}

function numericPngs(directory: string, prefix: string): string[] {
  const pattern = new RegExp(`^${prefix}-(\\d+)\\.png$`);
  return readdirSync(directory)
    .map(name => ({ name, match: name.match(pattern) }))
    .filter((item): item is { name: string; match: RegExpMatchArray } => item.match !== null)
    .sort((a, b) => Number(a.match[1]) - Number(b.match[1]))
    .map(item => join(directory, item.name));
}

function libreofficeEnv(work: string): NodeJS.ProcessEnv {
  const runtimeDir = join(work, 'runtime');
  const homeDir = join(work, 'home');
  for (const directory of [runtimeDir, homeDir]) mkdirSync(directory, { recursive: true, mode: 0o700 });
  return {
    ...process.env,
    HOME: homeDir,
    XDG_RUNTIME_DIR: runtimeDir,
    LANG: 'C.UTF-8',
    LC_ALL: 'C.UTF-8',
    TZ: 'UTC',
    // The same font-substitution contract Chromium runs under (playwright.config.ts).
    FONTCONFIG_FILE: FONT_CONTRACT_FILE,
  };
}

function libreofficePdf(docxPath: string, work: string): string {
  const pdfDir = join(work, 'pdf');
  const profileDir = join(work, 'profile');
  for (const directory of [pdfDir, profileDir]) mkdirSync(directory, { recursive: true, mode: 0o700 });

  execFileSync('libreoffice', [
    `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
    '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
    '--convert-to', 'pdf', '--outdir', pdfDir, docxPath,
  ], { env: libreofficeEnv(work), stdio: 'pipe', timeout: 120000 });

  const pdfPath = join(pdfDir, `${docxPath.split('/').pop()!.replace(/\.docx$/i, '')}.pdf`);
  if (!existsSync(pdfPath)) throw new Error(`LibreOffice did not produce ${pdfPath}`);
  return pdfPath;
}

function renderLibreOffice(docxPath: string, work: string): string[] {
  const pdfPath = libreofficePdf(docxPath, work);
  const prefix = 'libreoffice-page';
  execFileSync('pdftoppm', [
    '-r', '96', '-png', pdfPath, join(work, prefix),
  ], { env: libreofficeEnv(work), stdio: 'pipe', timeout: 120000 });
  const pages = numericPngs(work, prefix);
  if (!pages.length) throw new Error(`pdftoppm produced no pages for ${docxPath}`);
  return pages;
}

/**
 * Chromium must actually be running under the contract, not merely have it on disk — the config
 * injects FONTCONFIG_FILE only at browser launch. Calibri Light is the discriminator: without
 * the contract a host resolves it to whatever it has (Noto Sans, DejaVu Sans), whose advance
 * widths differ from the substitute's.
 */
async function assertBrowserFontContract(page: Page): Promise<void> {
  const mismatches = await page.evaluate((entries) => {
    const context = document.createElement('canvas').getContext('2d')!;
    const width = (family: string) => {
      context.font = `16px "${family}"`;
      return context.measureText('Sphinx of black quartz, judge my vow 0123456789').width;
    };
    return entries
      .filter(([family, substitute]) => Math.abs(width(family) - width(substitute)) > 0.5)
      .map(([family, substitute]) => `${family} does not measure as ${substitute}`);
  }, FONT_CONTRACT.map(entry => [entry.family, entry.substitute] as [string, string]));
  if (mismatches.length) {
    throw new Error(`Chromium is not applying the font contract (${FONT_CONTRACT_FILE}):\n  ` +
      mismatches.join('\n  '));
  }
}

async function materializeComparisonFixture(
  page: Page,
  entry: VisualCorpusEntry,
  work: string,
): Promise<string> {
  const sourcePath = resolve(repoRoot, entry.path);
  if ((entry.revisionMode ?? 'source') === 'source') return sourcePath;

  // LibreOffice's headless PDF conversion follows the document's saved redline-display state and
  // exposes no filter option for selecting Word's final view. Materialize one accepted-revision
  // copy outside the checkout, then give those identical bytes to both engines. This keeps the
  // tracked-change case apples-to-apples without importing or mutating a corpus fixture.
  const accepted = await page.evaluate((input) => {
    const bytes = (window as any).Docxodus.DocxDiffBridge.AcceptRevisions(new Uint8Array(input));
    return Array.from(bytes as Uint8Array);
  }, Array.from(new Uint8Array(readFileSync(sourcePath))));
  if (!accepted.length) throw new Error(`Revision acceptance produced no bytes for ${entry.path}`);

  const acceptedPath = join(work, 'accepted-revisions.docx');
  writeFileSync(acceptedPath, Buffer.from(accepted));
  return acceptedPath;
}

async function waitForStableRendering(page: Page): Promise<void> {
  await page.evaluate(async () => {
    await document.fonts.ready;
    const images = Array.from(document.images);
    await Promise.all(images.map(image => image.complete
      ? Promise.resolve()
      : new Promise<void>(resolveImage => {
          image.addEventListener('load', () => resolveImage(), { once: true });
          image.addEventListener('error', () => resolveImage(), { once: true });
        })));
    await new Promise<void>(resolveFrame => requestAnimationFrame(() =>
      requestAnimationFrame(() => resolveFrame())));
  });
}

async function renderDocxodus(page: Page, docxPath: string, output: string): Promise<string[]> {
  const bytes = new Uint8Array(readFileSync(docxPath));
  const result = await page.evaluate((input) => {
    try {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(input),
        'Document', 'docx-', true, '', -1, 'comment-',
        1, 1, 'page-',
        false, 0, 'annot-',
        true, true,
        false, false, false,
      );
      if (html.startsWith('{') && html.includes('"Error"')) return { error: html };
      return { html };
    } catch (error) {
      return { error: String(error) };
    }
  }, Array.from(bytes));
  if (!result.html) throw new Error(`Docxodus conversion failed: ${result.error}`);

  await page.evaluate((html) => {
    document.documentElement.style.background = 'white';
    document.body.innerHTML = '<main id="visual-parity-root"></main>';
    document.body.style.cssText = 'margin:0;padding:0;background:white;';
    const root = document.getElementById('visual-parity-root')!;
    (window as any).DocxodusPagination.paginateHtml(html, root, {
      scale: 1,
      showPageNumbers: false,
      pageGap: 0,
      fragmentParagraphs: false,
    });
    const style = document.createElement('style');
    style.textContent = `
      *, *::before, *::after {
        animation: none !important;
        caret-color: transparent !important;
        transition: none !important;
      }
      #pagination-container, .page-container { margin: 0 !important; padding: 0 !important; }
      .page-box { box-shadow: none !important; margin: 0 !important; }
    `;
    document.head.appendChild(style);
  }, result.html);
  await waitForStableRendering(page);

  const pages = page.locator('#visual-parity-root .page-box');
  const count = await pages.count();
  if (!count) throw new Error('Docxodus pagination produced no page boxes');
  const paths: string[] = [];
  for (let index = 0; index < count; index++) {
    const path = join(output, `docxodus-${index + 1}.png`);
    await pages.nth(index).screenshot({ path, animations: 'disabled', caret: 'hide' });
    paths.push(path);
  }
  return paths;
}

function worse(a: VisualSeverity, b: VisualSeverity): VisualSeverity {
  return severityOrder.indexOf(a) >= severityOrder.indexOf(b) ? a : b;
}

function summarizeCases(cases: CaseResult[]) {
  const counts = Object.fromEntries(severityOrder.map(level =>
    [level, cases.filter(result => result.severity === level).length]));
  const severeByDisposition = Object.fromEntries(VISUAL_DISPOSITION_KINDS.map(kind => [
    kind,
    cases.filter(result => result.severity === 'severe' && result.disposition.kind === kind).length,
  ]));
  return {
    cases: cases.length,
    pagesCompared: cases.reduce((sum, result) => sum + result.pages.length, 0),
    pageCountMismatches: cases.filter(result => result.pageCountDelta !== 0).length,
    errors: cases.filter(result => result.error !== undefined).length,
    severityCounts: counts,
    severeByDisposition,
    strictGatingCases: cases.filter(gatesStrictRun).map(result => result.id),
    meanSsim: cases.flatMap(result => result.pages).reduce((sum, page) => sum + page.ssim, 0) /
      Math.max(1, cases.flatMap(result => result.pages).length),
    meanInkF1: cases.flatMap(result => result.pages).reduce((sum, page) => sum + page.tolerantInkF1, 0) /
      Math.max(1, cases.flatMap(result => result.pages).length),
  };
}

test('stratified tracked corpus matches LibreOffice at pixel level', async ({ page, browserName }, testInfo) => {
  test.setTimeout(20 * 60 * 1000);
  assertTrackedCorpus(VISUAL_PARITY_CORPUS);
  const fontContract = fontContractReport(repoRoot); // throws if the host misses the contract
  // The reference-version contract (issue #403): fail in the first second with install guidance
  // rather than after twenty minutes at the ratchet's fingerprint check.
  const libreofficeVersion = assertLibreOfficeContract();
  const corpus = selectedCorpus();
  const configuredOutputRoot = resolve(process.env.DOCXODUS_VISUAL_PARITY_OUTPUT ??
    join(tmpdir(), 'docxodus-visual-parity'));
  const outputRelativeToRepo = relative(repoRoot, configuredOutputRoot);
  if (outputRelativeToRepo === '' ||
      (!outputRelativeToRepo.startsWith('..') && !isAbsolute(outputRelativeToRepo))) {
    throw new Error(`Visual parity artifacts must stay outside the repository: ${configuredOutputRoot}`);
  }
  const outputRoot = testInfo.retry === 0
    ? configuredOutputRoot
    : join(configuredOutputRoot, `retry-${testInfo.retry}`);
  if (existsSync(outputRoot) && readdirSync(outputRoot).length) {
    throw new Error(`Visual parity output must be empty to prevent stale artifacts: ${outputRoot}`);
  }
  mkdirSync(outputRoot, { recursive: true, mode: 0o700 });

  await page.setViewportSize({ width: 1400, height: 1000 });
  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 90000 });
  await assertBrowserFontContract(page);
  await page.addScriptTag({ url: '/pagination.bundle.js' });

  const cases: CaseResult[] = [];
  for (const [caseIndex, entry] of corpus.entries()) {
    const caseOutput = join(outputRoot, entry.id);
    mkdirSync(caseOutput, { recursive: true });
    const work = mkdtempSync(join(tmpdir(), `docxodus-visual-${entry.id}-`));
    let docxodusPageCount = 0;
    let libreofficePageCount = 0;
    try {
      const comparisonPath = await materializeComparisonFixture(page, entry, work);
      const libreofficeTempPages = renderLibreOffice(comparisonPath, work);
      libreofficePageCount = libreofficeTempPages.length;
      const docxodusPages = await renderDocxodus(page, comparisonPath, caseOutput);
      docxodusPageCount = docxodusPages.length;
      const libreofficePages = libreofficeTempPages.map((source, index) => {
        const destination = join(caseOutput, `libreoffice-${index + 1}.png`);
        writeFileSync(destination, readFileSync(source));
        return destination;
      });

      const pageResults: PageResult[] = [];
      const comparablePages = Math.min(docxodusPages.length, libreofficePages.length);
      let caseSeverity: VisualSeverity = docxodusPages.length === libreofficePages.length ? 'close' : 'severe';
      for (let index = 0; index < comparablePages; index++) {
        const comparison = compareImages(
          decodePng(readFileSync(docxodusPages[index])),
          decodePng(readFileSync(libreofficePages[index])),
        );
        const overlayPath = join(caseOutput, `overlay-${index + 1}.png`);
        writeFileSync(overlayPath, encodePng(comparison.overlay));
        caseSeverity = worse(caseSeverity, comparison.metrics.severity);
        pageResults.push({
          page: index + 1,
          ...comparison.metrics,
          artifacts: {
            docxodus: relative(outputRoot, docxodusPages[index]),
            libreoffice: relative(outputRoot, libreofficePages[index]),
            overlay: relative(outputRoot, overlayPath),
          },
          artifactSha256: {
            docxodus: sha256(docxodusPages[index]),
            libreoffice: sha256(libreofficePages[index]),
            overlay: sha256(overlayPath),
          },
        });
      }

      const result: CaseResult = {
        ...entry,
        docxodusPages: docxodusPages.length,
        libreofficePages: libreofficePages.length,
        pageCountDelta: docxodusPages.length - libreofficePages.length,
        severity: caseSeverity,
        pages: pageResults,
      };
      cases.push(result);
      writeFileSync(join(caseOutput, 'metrics.json'), `${JSON.stringify(result, null, 2)}\n`);
      console.log(`[${caseIndex + 1}/${corpus.length}] ${entry.id}: ` +
        `${result.docxodusPages}/${result.libreofficePages} pages, ${result.severity} ` +
        `(${result.disposition.kind})`);
    } catch (error) {
      const result: CaseResult = {
        ...entry,
        docxodusPages: docxodusPageCount,
        libreofficePages: libreofficePageCount,
        pageCountDelta: docxodusPageCount - libreofficePageCount,
        severity: 'severe',
        pages: [],
        error: error instanceof Error ? error.message : String(error),
      };
      cases.push(result);
      writeFileSync(join(caseOutput, 'metrics.json'), `${JSON.stringify(result, null, 2)}\n`);
      console.error(`[${caseIndex + 1}/${corpus.length}] ${entry.id}: ERROR ${result.error}`);
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
  }

  const summary = {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    gitCommit: commandVersion('git', ['rev-parse', 'HEAD']),
    workingTreeDirty: commandVersion('git', ['status', '--porcelain']).length > 0,
    revisionMode: 'final',
    revisionPreprocessing: 'tracked-change cases are accepted once outside the checkout and the identical bytes are rendered by both engines',
    dpi: 96,
    deviceScaleFactor: 1,
    masks: [],
    thresholds: VISUAL_THRESHOLDS,
    environment: {
      browserName,
      chromium: page.context().browser()?.version() ?? 'unknown',
      libreoffice: libreofficeVersion,
      pdftoppm: popplerVersionOutput(),
      fontContract,
      locale: 'C.UTF-8',
      timezone: 'UTC',
    },
    coverage: Object.fromEntries(REQUIRED_VISUAL_CATEGORIES.map(category => [
      category,
      VISUAL_PARITY_CORPUS.filter(entry => entry.categories.includes(category)).map(entry => entry.id),
    ])),
    // Word-evidence coverage (issue #402): which cases have recorded Word measurements backing
    // their dispositions, and which are still pending capture.
    wordReference: (() => {
      const record = readWordReference();
      if (!record) return { measured: [], pending: VISUAL_PARITY_CORPUS.map(entry => entry.id) };
      return {
        measured: record.cases.filter(entry => entry.status === 'measured').map(entry => entry.id),
        pending: record.cases.filter(entry => entry.status === 'pending').map(entry => entry.id),
      };
    })(),
    aggregate: summarizeCases(cases),
    cases,
  };
  const summaryPath = join(outputRoot, 'summary.json');
  writeFileSync(summaryPath, `${JSON.stringify(summary, null, 2)}\n`);
  console.log(`Visual parity report: ${summaryPath}`);

  // The regression ratchet (issue #395). Broader than strict mode: it covers every case at every
  // severity, so a `close` case sliding to `minor` is caught rather than merely archived in an
  // artifact that expires in 14 days.
  const complete = !process.env.DOCXODUS_VISUAL_PARITY_FILTER;
  if (process.env.DOCXODUS_VISUAL_PARITY_UPDATE_RECORD === '1') {
    if (!complete) {
      throw new Error('Refusing to write a partial ratchet record: unset ' +
        'DOCXODUS_VISUAL_PARITY_FILTER so every case is measured in the same run.');
    }
    assertRecordUpdateProvenance(summary);
    const record = buildRecord(summary, new Date().toISOString().slice(0, 10));
    writeFileSync(RATCHET_RECORD_FILE, serializeRecord(record));
    console.log(`Ratchet record refreshed: ${RATCHET_RECORD_FILE}`);
  } else {
    const record = readRecord();
    if (!record) {
      console.log('No ratchet record yet; create one with DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1.');
    } else {
      const comparison = compareToRecord(record, summary, { expectComplete: complete });
      console.log(`Ratchet [${comparison.status}]: ${comparison.message}`);
      // Opt-out rather than opt-in: a ratchet nobody enables is the artifact nobody downloaded.
      if (process.env.DOCXODUS_VISUAL_PARITY_RATCHET !== '0') {
        expect(comparison.status, comparison.message).toBe('ok');
      }
    }
  }

  if (process.env.DOCXODUS_VISUAL_PARITY_STRICT === '1') {
    expect(cases.filter(gatesStrictRun).map(result =>
      ({ id: result.id, severity: result.severity, disposition: result.disposition.kind, error: result.error })),
      'strict mode rejects severe cases attributed to the renderer (renderer-bug or unattributed) ' +
      'and conversion errors; environment and reference-deviation severes are reported, not gated',
    ).toEqual([]);
  }
});

/**
 * The contract drift probe (issue #379): one generated paragraph per declared family, wrapped in
 * a 6.5in column by both renderers. Wrapping is the metric-sensitive observable — if either
 * engine resolves a family to a different font, advance widths change and the paragraph wraps to
 * a different number of lines. fc-match proves what fontconfig WOULD resolve; this proves what
 * the two renderers actually DID.
 */
test('declared font families wrap identically in Chromium and LibreOffice', async ({ page }) => {
  test.setTimeout(5 * 60 * 1000);
  assertFontContract();
  assertLibreOfficeContract();

  const work = mkdtempSync(join(tmpdir(), 'docxodus-font-probe-'));
  try {
    const docxPath = join(work, 'font-probe.docx');
    writeFileSync(docxPath, Buffer.from(generateFontProbeDocx()));
    const pdfText = execFileSync('pdftotext', ['-layout', libreofficePdf(docxPath, work), '-'], {
      encoding: 'utf8',
      env: libreofficeEnv(work),
      timeout: 120000,
    });
    const libreofficeLines = probeLineCountsFromPdfText(pdfText);

    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 90000 });
    await assertBrowserFontContract(page);
    await page.addScriptTag({ url: '/pagination.bundle.js' });

    const chromiumLines = await page.evaluate(
      ({ bytes, markers }: { bytes: number[]; markers: string[] }) => {
        const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
          new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
          1, 1, 'page-', false, 0, 'annot-', true, true, false, false, false,
        );
        if (html.startsWith('{') && html.includes('"Error"')) throw new Error(html);
        document.body.innerHTML = '<main id="font-probe-root"></main>';
        (window as any).DocxodusPagination.paginateHtml(
          html, document.getElementById('font-probe-root'),
          { scale: 1, showPageNumbers: false, pageGap: 0, fragmentParagraphs: false });
        return document.fonts.ready.then(() => markers.map(marker => {
          // Only rendered pages — the hidden pagination staging copy has no client rects.
          const paragraph = Array.from(document.querySelectorAll('#font-probe-root .page-box p'))
            .find(candidate => (candidate.textContent || '').includes(marker));
          if (!paragraph) return -1;
          const range = document.createRange();
          range.selectNodeContents(paragraph);
          const lineTops = new Set(Array.from(range.getClientRects())
            .filter(rect => rect.width > 1 && rect.height > 1)
            .map(rect => Math.round(rect.top)));
          return lineTops.size;
        }));
      },
      {
        bytes: Array.from(generateFontProbeDocx()),
        markers: FONT_CONTRACT.map((_, index) => probeMarker(index)),
      });

    for (const [index, entry] of FONT_CONTRACT.entries()) {
      expect(chromiumLines[index], `${entry.family} paragraph missing from Chromium render`)
        .toBeGreaterThan(1);
      expect(libreofficeLines[index], `${entry.family} paragraph missing from LibreOffice render`)
        .toBeGreaterThan(1);
      expect(
        chromiumLines[index],
        `${entry.family} (contract: ${entry.substitute}) wraps to ${chromiumLines[index]} lines ` +
        `in Chromium but ${libreofficeLines[index]} in LibreOffice — the substitution contract ` +
        `has drifted between the renderers. Probe text: "${PROBE_TEXT.slice(0, 40)}…"`,
      ).toBe(libreofficeLines[index]);
    }
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
});
