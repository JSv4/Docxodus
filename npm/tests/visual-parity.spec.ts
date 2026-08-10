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
import { REQUIRED_VISUAL_CATEGORIES, VISUAL_PARITY_CORPUS, type VisualCorpusEntry } from './visual-parity/corpus.js';
import { compareImages, VISUAL_THRESHOLDS, type PageMetrics, type VisualSeverity } from './visual-parity/metrics.js';
import { decodePng, encodePng } from './visual-parity/png.js';
import {
  FONT_CONTRACT_PACKAGES,
  FONTCONFIG_FRAGMENT,
  resolveFontContract,
  writeFontconfigRoot,
  type FontContractStatus,
} from './visual-parity/fonts.js';
import {
  compareProbeLines,
  generateFontProbeDocx,
  inkLines,
  PROBE_FAMILIES,
  type FontProbeResult,
} from './visual-parity/font-probe.js';

test.skip(process.env.DOCXODUS_VISUAL_PARITY !== '1',
  'set DOCXODUS_VISUAL_PARITY=1 on a host with libreoffice and pdftoppm');

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '../..');
const severityOrder: VisualSeverity[] = ['close', 'minor', 'major', 'severe'];

/**
 * The font-substitution contract has to be in force before EITHER engine starts, and Chromium's
 * fontconfig is read when the browser process launches — which happens when the first test
 * requests a page, after this module is evaluated. So the layered fontconfig root is written
 * here, at import time, and handed to the browser through `test.use` below and to every
 * LibreOffice subprocess through its env.
 *
 * `DOCXODUS_VISUAL_PARITY_HOST_FONTS=1` opts out and measures the host as it is, which is how you
 * reproduce a report from a machine that installed the fragment permanently instead.
 */
const useHostFonts = process.env.DOCXODUS_VISUAL_PARITY_HOST_FONTS === '1';
const fontconfigRoot = useHostFonts || process.env.DOCXODUS_VISUAL_PARITY !== '1'
  ? undefined
  : writeFontconfigRoot(mkdtempSync(join(tmpdir(), 'docxodus-visual-fontconfig-')));
const fontEnv: NodeJS.ProcessEnv = fontconfigRoot
  ? { ...process.env, FONTCONFIG_FILE: fontconfigRoot }
  : process.env;

if (fontconfigRoot) {
  test.use({ launchOptions: { env: { ...process.env, FONTCONFIG_FILE: fontconfigRoot } } });
}

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
  docxodusPages: number;
  libreofficePages: number;
  pageCountDelta: number;
  severity: VisualSeverity;
  pages: PageResult[];
  error?: string;
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

function renderLibreOffice(docxPath: string, work: string): string[] {
  const pdfDir = join(work, 'pdf');
  const profileDir = join(work, 'profile');
  const runtimeDir = join(work, 'runtime');
  const homeDir = join(work, 'home');
  for (const directory of [pdfDir, profileDir, runtimeDir, homeDir]) mkdirSync(directory, { mode: 0o700 });

  const deterministicEnv = {
    ...fontEnv,
    HOME: homeDir,
    XDG_RUNTIME_DIR: runtimeDir,
    LANG: 'C.UTF-8',
    LC_ALL: 'C.UTF-8',
    TZ: 'UTC',
  };
  execFileSync('libreoffice', [
    `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
    '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
    '--convert-to', 'pdf', '--outdir', pdfDir, docxPath,
  ], { env: deterministicEnv, stdio: 'pipe', timeout: 120000 });

  const pdfPath = join(pdfDir, `${docxPath.split('/').pop()!.replace(/\.docx$/i, '')}.pdf`);
  if (!existsSync(pdfPath)) throw new Error(`LibreOffice did not produce ${pdfPath}`);
  const prefix = 'libreoffice-page';
  execFileSync('pdftoppm', [
    '-r', '96', '-png', pdfPath, join(work, prefix),
  ], { env: deterministicEnv, stdio: 'pipe', timeout: 120000 });
  const pages = numericPngs(work, prefix);
  if (!pages.length) throw new Error(`pdftoppm produced no pages for ${docxPath}`);
  return pages;
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

/**
 * Renders the synthetic wrapping probe through both engines and reports whether they agree.
 *
 * This is the signal that separates a renderer regression from a font-environment change: the
 * probe's only variable is which face each engine resolved, so when its lines stop matching, the
 * corpus numbers that moved with them did not move because of this repository.
 */
async function runFontProbe(page: Page, outputRoot: string): Promise<FontProbeResult & {
  families: string[];
  artifacts: { docxodus: string; libreoffice: string };
}> {
  const probeOutput = join(outputRoot, 'font-probe');
  mkdirSync(probeOutput, { recursive: true });
  const work = mkdtempSync(join(tmpdir(), 'docxodus-font-probe-'));
  try {
    const probePath = join(work, 'font-probe.docx');
    writeFileSync(probePath, generateFontProbeDocx());

    const libreofficePages = renderLibreOffice(probePath, work);
    const libreofficePng = join(probeOutput, 'libreoffice-1.png');
    writeFileSync(libreofficePng, readFileSync(libreofficePages[0]));
    const docxodusPages = await renderDocxodus(page, probePath, probeOutput);

    const comparison = compareProbeLines(
      inkLines(decodePng(readFileSync(docxodusPages[0]))),
      inkLines(decodePng(readFileSync(libreofficePng))),
    );
    return {
      ...comparison,
      families: PROBE_FAMILIES,
      artifacts: {
        docxodus: relative(outputRoot, docxodusPages[0]),
        libreoffice: relative(outputRoot, libreofficePng),
      },
    };
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

function worse(a: VisualSeverity, b: VisualSeverity): VisualSeverity {
  return severityOrder.indexOf(a) >= severityOrder.indexOf(b) ? a : b;
}

function summarizeCases(cases: CaseResult[]) {
  const counts = Object.fromEntries(severityOrder.map(level =>
    [level, cases.filter(result => result.severity === level).length]));
  return {
    cases: cases.length,
    pagesCompared: cases.reduce((sum, result) => sum + result.pages.length, 0),
    pageCountMismatches: cases.filter(result => result.pageCountDelta !== 0).length,
    errors: cases.filter(result => result.error !== undefined).length,
    severityCounts: counts,
    meanSsim: cases.flatMap(result => result.pages).reduce((sum, page) => sum + page.ssim, 0) /
      Math.max(1, cases.flatMap(result => result.pages).length),
    meanInkF1: cases.flatMap(result => result.pages).reduce((sum, page) => sum + page.tolerantInkF1, 0) /
      Math.max(1, cases.flatMap(result => result.pages).length),
  };
}

test('stratified tracked corpus matches LibreOffice at pixel level', async ({ page, browserName }) => {
  test.setTimeout(20 * 60 * 1000);
  assertTrackedCorpus(VISUAL_PARITY_CORPUS);
  const corpus = selectedCorpus();
  const outputRoot = resolve(process.env.DOCXODUS_VISUAL_PARITY_OUTPUT ??
    join(tmpdir(), 'docxodus-visual-parity'));
  const outputRelativeToRepo = relative(repoRoot, outputRoot);
  if (outputRelativeToRepo === '' ||
      (!outputRelativeToRepo.startsWith('..') && !isAbsolute(outputRelativeToRepo))) {
    throw new Error(`Visual parity artifacts must stay outside the repository: ${outputRoot}`);
  }
  if (existsSync(outputRoot) && readdirSync(outputRoot).length) {
    throw new Error(`Visual parity output must be empty to prevent stale artifacts: ${outputRoot}`);
  }
  mkdirSync(outputRoot, { recursive: true, mode: 0o700 });

  // The font contract gates the whole run: without it the two engines can be set in different
  // faces, and every number below would be measuring the host rather than the renderer.
  const fontContract: FontContractStatus = resolveFontContract(fontEnv);
  if (!fontContract.satisfied) {
    const message = `${fontContract.problem}\nRequired packages: ${FONT_CONTRACT_PACKAGES.join(' ')}`;
    if (process.env.DOCXODUS_VISUAL_PARITY_STRICT === '1') throw new Error(message);
    test.skip(true, message);
  }

  await page.setViewportSize({ width: 1400, height: 1000 });
  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 90000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });

  const fontProbe = await runFontProbe(page, outputRoot);
  console.log(fontProbe.agreed
    ? `Font probe: engines agree on ${fontProbe.docxodusLines} lines ` +
      `(worst advance delta ${fontProbe.maxAdvanceDeltaPx} px)`
    : `Font probe: FONT ENVIRONMENT DRIFT — ${fontProbe.problem}`);

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
        `${result.docxodusPages}/${result.libreofficePages} pages, ${result.severity}`);
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
      libreoffice: commandVersion('libreoffice', ['--version']),
      pdftoppm: commandVersion('pdftoppm', ['-v']).split('\n')[0],
      locale: 'C.UTF-8',
      timezone: 'UTC',
    },
    // The exact faces both engines rendered with, so a rerun can tell a renderer change from a
    // font-environment change instead of guessing.
    fontContract: {
      ...fontContract,
      fragment: relative(repoRoot, FONTCONFIG_FRAGMENT),
      source: useHostFonts ? 'host fontconfig (DOCXODUS_VISUAL_PARITY_HOST_FONTS=1)' : 'repository fragment',
    },
    fontProbe,
    coverage: Object.fromEntries(REQUIRED_VISUAL_CATEGORIES.map(category => [
      category,
      VISUAL_PARITY_CORPUS.filter(entry => entry.categories.includes(category)).map(entry => entry.id),
    ])),
    aggregate: summarizeCases(cases),
    cases,
  };
  const summaryPath = join(outputRoot, 'summary.json');
  writeFileSync(summaryPath, `${JSON.stringify(summary, null, 2)}\n`);
  console.log(`Visual parity report: ${summaryPath}`);

  // Reported before the corpus verdict on purpose: when the probe fails, the corpus numbers are
  // not evidence about the renderer, and saying so first keeps the two from being confused.
  expect(fontProbe.agreed, `font environment drift, not a renderer regression: ${fontProbe.problem}`)
    .toBe(true);

  if (process.env.DOCXODUS_VISUAL_PARITY_STRICT === '1') {
    expect(cases.filter(result => result.severity === 'severe'),
      'strict mode rejects severe page-count, geometry, or visual discrepancies').toEqual([]);
  }
});
