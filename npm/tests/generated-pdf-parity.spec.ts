import { expect, test, type Page } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import {
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  GATING_DISPOSITION_KINDS,
  type VisualDisposition,
} from './visual-parity/corpus.js';
import {
  PDF_PARITY_CORPUS,
  REQUIRED_PDF_PARITY_CATEGORIES,
  type PdfParityCorpusEntry,
} from './visual-parity/pdf-corpus.js';
import {
  compareImages,
  VISUAL_THRESHOLDS,
  type PageMetrics,
  type VisualSeverity,
} from './visual-parity/metrics.js';
import {
  PDF_RASTER_CONTRACT,
  PDF_RASTER_CONTRACT_SHA256,
  PDF_PARITY_LIMITS,
  inspectPdf,
  rasterizePdf,
  sha256File,
  type PdfBox,
  type PdfInspection,
} from './visual-parity/pdf.js';
import { decodePng, encodePng } from './visual-parity/png.js';
import {
  FONT_CONTRACT_FILE,
  fontContractReport,
} from './visual-parity/font-contract.js';
import {
  assertLibreOfficeContract,
} from './visual-parity/environment-contract.js';
import {
  assertRecordUpdateProvenance,
  buildRecord,
  compareToRecord,
  readRecord,
  serializeRecord,
} from './visual-parity/ratchet.js';
import {
  assertSafeCaseId,
  prepareExternalOutputRoot,
  resolveTrackedRegularFile,
} from './visual-parity/benchmark-paths.js';
import {
  commandVersion,
  pinExecutable,
  resolveExecutable,
} from './visual-parity/toolchain.js';
import {
  assertBuildOwningLifecycle,
  captureGeneratedPdfBuildEvidence,
} from './visual-parity/build-provenance.js';
import { assertSupportedPdfResult } from './visual-parity/pdf-result.js';
import { exactLinkEvidence, type LinkEvidence } from './visual-parity/pdf-links.js';

assertBuildOwningLifecycle(
  process.env.DOCXODUS_GENERATED_PDF_PARITY === '1',
  process.env.npm_lifecycle_event,
);

test.skip(process.env.DOCXODUS_GENERATED_PDF_PARITY !== '1',
  'set DOCXODUS_GENERATED_PDF_PARITY=1 on a host with LibreOffice and Poppler');

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '../..');
const severityOrder: readonly VisualSeverity[] = ['close', 'minor', 'major', 'severe'];
const GENERATED_PDF_RATCHET_RECORD_FILE = resolve(
  __dirname,
  'visual-parity/generated-pdf-ratchet.json',
);
const PHYSICAL_GEOMETRY_TOLERANCE_POINTS = 0.5;
const BOOTSTRAP_ARTIFACTS = new Set(['ci-context.json', 'index.html']);

interface ExportModule {
  convertDocxToPdf(document: Uint8Array, options: Record<string, unknown>): Promise<{
    pdf: Uint8Array;
    pageCount: number;
    pageMap: unknown;
    renderReport: {
      status: 'complete';
      environment: { rendererFingerprint: string; verification: string };
      bindings: { pdfDigest?: string };
      source: { rawPackageBytesDigest: string };
      options: { reviewProfile: string; commentProfile: string };
      [key: string]: unknown;
    };
    rendererFingerprint: string;
    warnings: unknown[];
  }>;
}

interface GeometryDelta {
  x: number;
  y: number;
  width: number;
  height: number;
  maximumAbsoluteDelta: number;
}

interface PageGeometryComparison {
  page: number;
  candidate: { mediaBox: PdfBox; cropBox: PdfBox };
  reference: { mediaBox: PdfBox; cropBox: PdfBox };
  mediaBoxDelta: GeometryDelta;
  cropBoxDelta: GeometryDelta;
  passed: boolean;
}

interface TextExtractorEvidence {
  characterCount: number;
  sha256: string;
  requiredText: string[];
  missingRequiredText: string[];
  forbiddenText: string[];
  observedForbiddenText: string[];
  passed: boolean;
}

interface SemanticEvidence {
  candidate: {
    pdfjs: TextExtractorEvidence;
    pdftotext: TextExtractorEvidence;
    links: LinkEvidence;
  };
  reference: {
    pdfjs: TextExtractorEvidence;
    pdftotext: TextExtractorEvidence;
  };
  /** Candidate-side verdict. This, and only this, gates the run. */
  passed: boolean;
  /**
   * Whether the REFERENCE PDF also yielded the expected text. A failure here is a property of
   * LibreOffice, Poppler, or the font contract — not of Docxodus — so it is reported rather than
   * gated, exactly as a raster environment change is. Folding it into `passed` would fail an
   * "unconditional" contract that names Docxodus for someone else's release.
   */
  referenceExtractionHealthy: boolean;
}

interface PdfParityPageResult extends PageMetrics {
  page: number;
  physicalGeometry: PageGeometryComparison;
  environmentFingerprint: string;
  rendererFingerprint: string;
  artifacts: { candidate: string; reference: string; overlay: string };
  artifactSha256: { candidate: string; reference: string; overlay: string };
}

interface PdfParityCaseResult {
  id: string;
  path: string;
  categories: readonly string[];
  rationale: string;
  disposition: VisualDisposition;
  profiles: PdfParityCorpusEntry['profiles'];
  provenance: PdfParityCorpusEntry['source']['provenance'];
  sourceSha256: string;
  sourceGitBlob: string;
  referenceSourceSha256: string;
  docxodusPages: number;
  libreofficePages: number;
  pageCountDelta: number;
  physicalGeometryPassed: boolean;
  semanticChecksPassed: boolean;
  vectorContentPassed: boolean;
  vectorPathOperations?: { candidate: number; reference: number };
  severity: VisualSeverity;
  rendererFingerprint?: string;
  environmentFingerprint?: string;
  artifacts: Record<string, string>;
  artifactSha256: Record<string, string>;
  semantic?: SemanticEvidence;
  pages: PdfParityPageResult[];
  error?: string;
}

interface RunState {
  schemaVersion: 1;
  status: 'running' | 'passed' | 'failed';
  phase: string;
  startedAt: string;
  updatedAt: string;
  outputRoot: string;
  casesCompleted: number;
  casesSelected: number;
  failure?: { phase: string; message: string };
}

function sha256(bytes: Uint8Array | string): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function writeJsonAtomic(path: string, value: unknown): void {
  writeTextAtomic(path, `${JSON.stringify(value, null, 2)}\n`);
}

function writeTextAtomic(path: string, value: string): void {
  const temporary = `${path}.${process.pid}.${randomUUID()}.tmp`;
  try {
    writeFileSync(temporary, value, { flag: 'wx', mode: 0o600 });
    renameSync(temporary, path);
  } catch (error) {
    // The record path stages inside the repository, so a surviving .tmp would make the NEXT
    // run's `git status --porcelain` dirty and refuse the refresh, blaming the implementation
    // for this benchmark's own leftover file.
    rmSync(temporary, { force: true });
    throw error;
  }
}

function escapeHtml(value: unknown): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function writeViewer(outputRoot: string, state: RunState, cases: PdfParityCaseResult[]): void {
  const caseSections = cases.map((entry) => {
    const pageRows = entry.pages.map((page) => `
      <tr><th>Page ${page.page}</th>
        <td><a href="${escapeHtml(page.artifacts.candidate)}"><img src="${escapeHtml(page.artifacts.candidate)}" alt="Docxodus page ${page.page}"></a></td>
        <td><a href="${escapeHtml(page.artifacts.reference)}"><img src="${escapeHtml(page.artifacts.reference)}" alt="Reference page ${page.page}"></a></td>
        <td><a href="${escapeHtml(page.artifacts.overlay)}"><img src="${escapeHtml(page.artifacts.overlay)}" alt="Difference overlay ${page.page}"></a></td>
      </tr>`).join('');
    const links = Object.entries(entry.artifacts)
      .map(([label, path]) => `<li><a href="${escapeHtml(path)}">${escapeHtml(label)}</a></li>`)
      .join('');
    return `<section>
      <h2>${escapeHtml(entry.id)} — ${escapeHtml(entry.severity)}</h2>
      <p>${escapeHtml(entry.rationale)}</p>
      ${entry.error ? `<p class="failure"><strong>Failure:</strong> ${escapeHtml(entry.error)}</p>` : ''}
      <ul>${links}</ul>
      ${pageRows ? `<table><thead><tr><th></th><th>Docxodus PDF</th><th>LibreOffice PDF</th><th>Overlay</th></tr></thead><tbody>${pageRows}</tbody></table>` : ''}
    </section>`;
  }).join('\n');
  const failureLink = existsSync(join(outputRoot, 'failure.json'))
    ? '<li><a href="failure.json">Failure details</a></li>'
    : '';
  const summaryLink = existsSync(join(outputRoot, 'summary.json'))
    ? '<li><a href="summary.json">Complete summary</a></li>'
    : '<li><a href="summary.partial.json">Partial summary</a></li>';
  const ciLink = existsSync(join(outputRoot, 'ci-context.json'))
    ? '<li><a href="ci-context.json">CI context</a></li>'
    : '';
  writeTextAtomic(join(outputRoot, 'index.html'), `<!doctype html>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Docxodus generated-PDF parity evidence</title>
<style>
body{font:15px/1.45 system-ui;margin:2rem;max-width:110rem}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:.5rem;vertical-align:top}img{display:block;width:100%;min-width:14rem;height:auto}.failure{color:#a00}code{word-break:break-all}section{margin:2rem 0 4rem}
</style>
<h1>Docxodus generated-PDF parity evidence</h1>
<p>Status: <strong>${escapeHtml(state.status)}</strong>; phase: <code>${escapeHtml(state.phase)}</code>; ${state.casesCompleted}/${state.casesSelected} cases materialized.</p>
<ul><li><a href="run.json">Run state</a></li>${summaryLink}${failureLink}${ciLink}</ul>
${caseSections || '<p>No case artifacts have materialized yet. Inspect the run/failure context above.</p>'}
`);
}

function initializeOutputRoot(retry: number): string {
  const configuredOutputRoot = resolve(process.env.DOCXODUS_GENERATED_PDF_PARITY_OUTPUT
    ?? join(tmpdir(), 'docxodus-generated-pdf-parity'));
  return prepareExternalOutputRoot(repoRoot, configuredOutputRoot, retry, BOOTSTRAP_ARTIFACTS);
}

function selectedCorpus(): PdfParityCorpusEntry[] {
  const filter = process.env.DOCXODUS_GENERATED_PDF_PARITY_FILTER;
  if (!filter) return [...PDF_PARITY_CORPUS.cases];
  const requested = new Set(filter.split(',').map((value) => value.trim()).filter(Boolean));
  const selected = PDF_PARITY_CORPUS.cases.filter((entry) => requested.has(entry.id));
  const unknown = [...requested].filter((id) => !PDF_PARITY_CORPUS.cases.some((entry) => entry.id === id));
  if (unknown.length) throw new Error(`Unknown generated-PDF parity case(s): ${unknown.join(', ')}`);
  if (!selected.length) throw new Error('Generated-PDF parity filter selected no cases.');
  return selected;
}

function assertPinnedSource(entry: PdfParityCorpusEntry, gitExecutable: string): Uint8Array {
  assertSafeCaseId(entry.id);
  const path = resolveTrackedRegularFile(repoRoot, entry.source.path);
  execFileSync(gitExecutable, ['ls-files', '--error-unmatch', entry.source.path], {
    cwd: repoRoot,
    stdio: 'pipe',
  });
  const bytes = new Uint8Array(readFileSync(path));
  const digest = sha256(bytes);
  if (digest !== entry.source.sha256) {
    throw new Error(`${entry.id} source SHA-256 changed (${digest} != ${entry.source.sha256}).`);
  }
  const blob = execFileSync(gitExecutable, ['hash-object', entry.source.path], {
    cwd: repoRoot,
    encoding: 'utf8',
  }).trim();
  if (blob !== entry.source.gitBlob) {
    throw new Error(`${entry.id} source Git blob changed (${blob} != ${entry.source.gitBlob}).`);
  }
  const headBlob = execFileSync(gitExecutable, ['rev-parse', `HEAD:${entry.source.path}`], {
    cwd: repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  }).trim();
  if (headBlob !== entry.source.gitBlob) {
    throw new Error(`${entry.id} manifest does not bind the fixture at HEAD (${headBlob} != ${entry.source.gitBlob}).`);
  }
  return bytes;
}

/**
 * Poppler resolves fonts through fontconfig for any face the PDF does not embed, so it has to run
 * under the SAME contract both renderers used. The runner process does not inherit it — the
 * Playwright config injects FONTCONFIG_FILE into `launchOptions.env` (the browser), not here — so
 * omitting this leaves pdftoppm/pdftotext on host fonts while `summary.json` reports the pinned
 * contract as though it applied, and a distro font update then reads as a renderer regression.
 */
function popplerEnv(): NodeJS.ProcessEnv {
  return {
    ...process.env,
    LANG: 'C.UTF-8',
    LC_ALL: 'C.UTF-8',
    TZ: 'UTC',
    FONTCONFIG_FILE: FONT_CONTRACT_FILE,
  };
}

function libreofficeEnv(work: string): NodeJS.ProcessEnv {
  const runtimeDir = join(work, 'runtime');
  const homeDir = join(work, 'home');
  for (const directory of [runtimeDir, homeDir]) {
    mkdirSync(directory, { recursive: true, mode: 0o700 });
  }
  return {
    ...process.env,
    HOME: homeDir,
    XDG_RUNTIME_DIR: runtimeDir,
    LANG: 'C.UTF-8',
    LC_ALL: 'C.UTF-8',
    TZ: 'UTC',
    FONTCONFIG_FILE: FONT_CONTRACT_FILE,
  };
}

async function ensureWasmBridge(page: Page): Promise<void> {
  if (page.url().endsWith('/test-harness.html')) return;
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 90_000 });
}

async function referenceSource(
  page: Page,
  entry: PdfParityCorpusEntry,
  source: Uint8Array,
  caseOutput: string,
): Promise<{ path: string; sha256: string }> {
  if (entry.profiles.reference.revisionProjection === 'source') {
    const path = join(caseOutput, 'reference-source.docx');
    writeFileSync(path, source);
    return { path, sha256: sha256(source) };
  }
  await ensureWasmBridge(page);
  const accepted = await page.evaluate((input) => {
    const bytes = (window as any).Docxodus.DocxDiffBridge.AcceptRevisions(new Uint8Array(input));
    return Array.from(bytes as Uint8Array);
  }, Array.from(source));
  if (!accepted.length) throw new Error(`${entry.id}: final-view reference projection returned no bytes.`);
  const bytes = new Uint8Array(accepted);
  const path = join(caseOutput, 'reference-source.docx');
  writeFileSync(path, bytes);
  return { path, sha256: sha256(bytes) };
}

function renderReferencePdf(
  libreofficeExecutable: string,
  sourcePath: string,
  work: string,
  destination: string,
): Uint8Array {
  const pdfDirectory = join(work, 'reference-pdf');
  const profileDirectory = join(work, 'libreoffice-profile');
  mkdirSync(pdfDirectory, { recursive: true, mode: 0o700 });
  mkdirSync(profileDirectory, { recursive: true, mode: 0o700 });
  execFileSync(libreofficeExecutable, [
    `-env:UserInstallation=${pathToFileURL(profileDirectory).href}`,
    '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
    '--convert-to', 'pdf:writer_pdf_Export', '--outdir', pdfDirectory, sourcePath,
  ], {
    env: libreofficeEnv(work),
    stdio: 'pipe',
    timeout: 120_000,
  });
  const produced = join(pdfDirectory, `${basename(sourcePath).replace(/\.docx$/i, '')}.pdf`);
  if (!existsSync(produced)) throw new Error(`LibreOffice did not produce ${produced}.`);
  const producedMetadata = lstatSync(produced);
  if (!producedMetadata.isFile() || producedMetadata.isSymbolicLink()
    || producedMetadata.size > PDF_PARITY_LIMITS.maximumPdfBytes) {
    throw new Error(`LibreOffice PDF is not a bounded regular non-symlink file: ${produced}`);
  }
  const bytes = new Uint8Array(readFileSync(produced));
  writeFileSync(destination, bytes);
  return bytes;
}

function normalizeText(value: string): string {
  return value.normalize('NFC').replace(/\s+/g, ' ').trim();
}

function textEvidence(text: string, entry: PdfParityCorpusEntry): TextExtractorEvidence {
  const normalized = normalizeText(text);
  const requiredText = entry.semantics.requiredText.map(normalizeText);
  const forbiddenText = (entry.semantics.forbiddenText ?? []).map(normalizeText);
  const missingRequiredText = requiredText.filter((fragment) => !normalized.includes(fragment));
  const observedForbiddenText = forbiddenText.filter((fragment) => normalized.includes(fragment));
  return {
    characterCount: normalized.length,
    sha256: sha256(normalized),
    requiredText,
    missingRequiredText,
    forbiddenText,
    observedForbiddenText,
    passed: normalized.length > 0
      && missingRequiredText.length === 0
      && observedForbiddenText.length === 0,
  };
}

function popplerText(pdftotextExecutable: string, path: string): string {
  return execFileSync(pdftotextExecutable, ['-enc', 'UTF-8', '-nopgbrk', path, '-'], {
    encoding: 'utf8',
    env: popplerEnv(),
    timeout: 120_000,
    maxBuffer: 6 * 1024 * 1024,
  });
}

function semanticEvidence(
  pdftotextExecutable: string,
  candidatePath: string,
  candidateInspection: PdfInspection,
  referencePath: string,
  referenceInspection: PdfInspection,
  entry: PdfParityCorpusEntry,
): SemanticEvidence {
  const candidatePdfjs = textEvidence(candidateInspection.searchableText, entry);
  const candidatePoppler = textEvidence(popplerText(pdftotextExecutable, candidatePath), entry);
  const candidateLinks = exactLinkEvidence(candidateInspection, entry);
  const referencePdfjs = textEvidence(referenceInspection.searchableText, entry);
  const referencePoppler = textEvidence(popplerText(pdftotextExecutable, referencePath), entry);
  return {
    candidate: {
      pdfjs: candidatePdfjs,
      pdftotext: candidatePoppler,
      links: candidateLinks,
    },
    reference: {
      pdfjs: referencePdfjs,
      pdftotext: referencePoppler,
    },
    passed: candidatePdfjs.passed && candidatePoppler.passed && candidateLinks.passed,
    referenceExtractionHealthy: referencePdfjs.passed && referencePoppler.passed,
  };
}

function boxDelta(candidate: PdfBox, reference: PdfBox): GeometryDelta {
  const delta = {
    x: candidate.x - reference.x,
    y: candidate.y - reference.y,
    width: candidate.width - reference.width,
    height: candidate.height - reference.height,
  };
  return {
    ...delta,
    maximumAbsoluteDelta: Math.max(...Object.values(delta).map(Math.abs)),
  };
}

function compareGeometry(
  candidate: PdfInspection,
  reference: PdfInspection,
): PageGeometryComparison[] {
  const count = Math.min(candidate.pages.length, reference.pages.length);
  return Array.from({ length: count }, (_, index) => {
    const candidatePage = candidate.pages[index];
    const referencePage = reference.pages[index];
    const mediaBoxDelta = boxDelta(candidatePage.mediaBox, referencePage.mediaBox);
    const cropBoxDelta = boxDelta(candidatePage.cropBox, referencePage.cropBox);
    return {
      page: index + 1,
      candidate: { mediaBox: candidatePage.mediaBox, cropBox: candidatePage.cropBox },
      reference: { mediaBox: referencePage.mediaBox, cropBox: referencePage.cropBox },
      mediaBoxDelta,
      cropBoxDelta,
      passed: mediaBoxDelta.maximumAbsoluteDelta <= PHYSICAL_GEOMETRY_TOLERANCE_POINTS
        && cropBoxDelta.maximumAbsoluteDelta <= PHYSICAL_GEOMETRY_TOLERANCE_POINTS,
    };
  });
}

function worse(left: VisualSeverity, right: VisualSeverity): VisualSeverity {
  return severityOrder.indexOf(left) >= severityOrder.indexOf(right) ? left : right;
}

function hasHardParityFailure(result: PdfParityCaseResult): boolean {
  return result.error !== undefined
    || result.pageCountDelta !== 0
    || !result.physicalGeometryPassed
    || !result.semanticChecksPassed
    || !result.vectorContentPassed;
}

function gatesStrictRasterRun(result: PdfParityCaseResult): boolean {
  return result.severity === 'severe'
    && GATING_DISPOSITION_KINDS.includes(result.disposition.kind);
}

function summarizeCases(cases: PdfParityCaseResult[]) {
  const pages = cases.flatMap((entry) => entry.pages);
  return {
    cases: cases.length,
    pagesCompared: pages.length,
    pageCountMismatches: cases.filter((entry) => entry.pageCountDelta !== 0).length,
    physicalGeometryFailures: cases.filter((entry) => !entry.physicalGeometryPassed).length,
    semanticFailures: cases.filter((entry) => !entry.semanticChecksPassed).length,
    // Reported, never gated: see SemanticEvidence.referenceExtractionHealthy.
    referenceExtractionFailures: cases.filter((entry) =>
      entry.semantic !== undefined && !entry.semantic.referenceExtractionHealthy).length,
    vectorFailures: cases.filter((entry) => !entry.vectorContentPassed).length,
    errors: cases.filter((entry) => entry.error !== undefined).length,
    severityCounts: Object.fromEntries(severityOrder.map((level) => [
      level,
      cases.filter((entry) => entry.severity === level).length,
    ])),
    hardGatingCases: cases.filter(hasHardParityFailure).map((entry) => entry.id),
    strictRasterGatingCases: cases.filter(gatesStrictRasterRun).map((entry) => entry.id),
    meanSsim: pages.reduce((sum, page) => sum + page.ssim, 0) / Math.max(1, pages.length),
    meanInkPrecision: pages.reduce((sum, page) => sum + page.tolerantInkPrecision, 0)
      / Math.max(1, pages.length),
    meanInkRecall: pages.reduce((sum, page) => sum + page.tolerantInkRecall, 0)
      / Math.max(1, pages.length),
    meanInkF1: pages.reduce((sum, page) => sum + page.tolerantInkF1, 0)
      / Math.max(1, pages.length),
  };
}

test('supported generated PDFs match reference PDFs through the fidelity ratchet', async ({ page, browserName }, testInfo) => {
  test.setTimeout(30 * 60 * 1000);

  const outputRoot = initializeOutputRoot(testInfo.retry);
  const corpus = selectedCorpus();
  const startedAt = new Date().toISOString();
  // Capture the source identity before any benchmark-owned write. Record updates intentionally
  // dirty the repository at the end of a successful run; recomputing status while writing the
  // final artifact would therefore mislabel a verified clean source as dirty.
  const git = pinExecutable('git', ['--version']);
  const sourceCommit = commandVersion(git.path, ['rev-parse', 'HEAD']);
  const sourceTree = commandVersion(git.path, ['rev-parse', 'HEAD^{tree}']);
  const sourceWorkingTreeStatus = commandVersion(git.path, ['status', '--porcelain']);
  const sourceWorkingTreeDirty = sourceWorkingTreeStatus.length > 0;
  const cases: PdfParityCaseResult[] = [];
  let phase = 'artifact-initialization';
  let environment: Record<string, unknown> = {};
  const state: RunState = {
    schemaVersion: 1,
    status: 'running',
    phase,
    startedAt,
    updatedAt: startedAt,
    outputRoot,
    casesCompleted: 0,
    casesSelected: corpus.length,
  };

  const persist = (complete = false): void => {
    state.phase = phase;
    state.updatedAt = new Date().toISOString();
    state.casesCompleted = cases.length;
    writeJsonAtomic(join(outputRoot, 'run.json'), state);
    const summary = {
      schemaVersion: 1,
      measurementPipeline: 'generated-pdf-v1',
      generatedAt: state.updatedAt,
      gitCommit: sourceCommit,
      gitTree: sourceTree,
      workingTreeDirty: sourceWorkingTreeDirty,
      rasterContract: {
        ...PDF_RASTER_CONTRACT,
        sha256: PDF_RASTER_CONTRACT_SHA256,
      },
      physicalGeometryTolerancePoints: PHYSICAL_GEOMETRY_TOLERANCE_POINTS,
      thresholds: VISUAL_THRESHOLDS,
      environment,
      coverage: Object.fromEntries(REQUIRED_PDF_PARITY_CATEGORIES.map((category) => [
        category,
        PDF_PARITY_CORPUS.cases.filter((entry) =>
          (entry.categories as readonly string[]).includes(category))
          .map((entry) => entry.id),
      ])),
      aggregate: summarizeCases(cases),
      cases,
    };
    writeJsonAtomic(join(outputRoot, complete ? 'summary.json' : 'summary.partial.json'), summary);
    writeViewer(outputRoot, state, cases);
  };

  persist();
  try {
    phase = 'environment-preflight';
    const tools = {
      fcMatch: pinExecutable('fc-match', ['--version']),
      libreoffice: pinExecutable('libreoffice', ['--version']),
      pdftoppm: pinExecutable('pdftoppm', ['-v']),
      pdftotext: pinExecutable('pdftotext', ['-v']),
    };
    const fontContract = fontContractReport(repoRoot, tools.fcMatch.path);
    const libreofficeVersion = assertLibreOfficeContract(tools.libreoffice.evidence.version);
    const poppler = tools.pdftoppm.evidence.version.split('\n')[0].trim();
    const pdftotext = tools.pdftotext.evidence.version;
    const build = captureGeneratedPdfBuildEvidence(repoRoot);
    const exportEntry = resolveTrackedRegularFile(repoRoot, 'npm-export/dist/index.js');
    const exporter = await import(pathToFileURL(exportEntry).href) as ExportModule;
    if (typeof exporter.convertDocxToPdf !== 'function') {
      throw new Error('The supported @docxodus/export entry point does not export convertDocxToPdf.');
    }
    const playwrightCoreEntry = resolveTrackedRegularFile(
      repoRoot,
      'npm-export/node_modules/playwright-core/index.mjs',
    );
    const exporterPlaywright = await import(pathToFileURL(playwrightCoreEntry).href) as {
      chromium: { executablePath(): string; launch(options: Record<string, unknown>): Promise<{
        version(): string;
        close(): Promise<void>;
      }> };
    };
    const chromiumExecutablePath = resolveExecutable(exporterPlaywright.chromium.executablePath());
    const chromiumProbe = await exporterPlaywright.chromium.launch({
      headless: true,
      executablePath: chromiumExecutablePath,
    });
    const chromiumVersion = chromiumProbe.version();
    await chromiumProbe.close();
    environment = {
      browserName,
      chromium: chromiumVersion,
      chromiumExecutable: {
        executable: basename(chromiumExecutablePath),
        sha256: sha256File(chromiumExecutablePath),
      },
      node: process.version,
      source: { commit: sourceCommit, tree: sourceTree, workingTreeDirty: sourceWorkingTreeDirty },
      build,
      moduleEntries: {
        exporter: sha256File(exportEntry),
        playwrightCore: sha256File(playwrightCoreEntry),
      },
      libreoffice: libreofficeVersion,
      pdftoppm: poppler,
      pdftotext,
      tools: {
        git: git.evidence,
        fcMatch: tools.fcMatch.evidence,
        libreoffice: tools.libreoffice.evidence,
        pdftoppm: tools.pdftoppm.evidence,
        pdftotext: tools.pdftotext.evidence,
      },
      fontContract,
      locale: 'C.UTF-8',
      timezone: 'UTC',
      rasterContractSha256: PDF_RASTER_CONTRACT_SHA256,
    };
    persist();

    for (const [caseIndex, entry] of corpus.entries()) {
      phase = `case:${entry.id}`;
      assertSafeCaseId(entry.id);
      const caseOutput = join(outputRoot, entry.id);
      mkdirSync(caseOutput, { recursive: true, mode: 0o700 });
      const work = mkdtempSync(join(tmpdir(), `docxodus-generated-pdf-${entry.id}-`));
      const artifacts: Record<string, string> = {
        source: `${entry.id}/source.docx`,
        referenceSource: `${entry.id}/reference-source.docx`,
        metrics: `${entry.id}/metrics.json`,
      };
      let sourceSha256 = entry.source.sha256;
      let referenceSourceSha256 = entry.source.sha256;
      let docxodusPages = 0;
      let libreofficePages = 0;
      try {
        const source = assertPinnedSource(entry, git.path);
        sourceSha256 = sha256(source);
        writeFileSync(join(caseOutput, 'source.docx'), source);
        const projectedReference = await referenceSource(page, entry, source, caseOutput);
        referenceSourceSha256 = projectedReference.sha256;

        const candidate = await exporter.convertDocxToPdf(source, {
          ...entry.profiles.candidate,
          expectedSourceDigest: sourceSha256,
          unsupportedContent: 'warn',
          timeoutMs: 120_000,
        });
        const verifiedCandidate = assertSupportedPdfResult(candidate, {
          sourceSha256,
          reviewProfile: entry.profiles.candidate.reviewProfile,
          commentProfile: entry.profiles.candidate.commentProfile,
        });
        docxodusPages = verifiedCandidate.pageCount;
        const candidatePdfPath = join(caseOutput, 'docxodus.pdf');
        writeFileSync(candidatePdfPath, candidate.pdf);
        writeJsonAtomic(join(caseOutput, 'render-report.json'), candidate.renderReport);
        writeJsonAtomic(join(caseOutput, 'page-map.json'), candidate.pageMap);
        artifacts.candidatePdf = `${entry.id}/docxodus.pdf`;
        artifacts.renderReport = `${entry.id}/render-report.json`;
        artifacts.pageMap = `${entry.id}/page-map.json`;

        if (candidate.renderReport.source.rawPackageBytesDigest !== sourceSha256) {
          throw new Error(`${entry.id}: render report source digest does not bind the pinned source.`);
        }
        if (candidate.renderReport.options.reviewProfile !== entry.profiles.candidate.reviewProfile
          || candidate.renderReport.options.commentProfile !== entry.profiles.candidate.commentProfile) {
          throw new Error(`${entry.id}: render report profile differs from the corpus contract.`);
        }
        if (candidate.renderReport.bindings.pdfDigest !== sha256(candidate.pdf)) {
          throw new Error(`${entry.id}: render report PDF digest does not bind the returned artifact.`);
        }

        const referencePdfPath = join(caseOutput, 'reference.pdf');
        const referencePdf = renderReferencePdf(
          tools.libreoffice.path,
          projectedReference.path,
          work,
          referencePdfPath,
        );
        artifacts.referencePdf = `${entry.id}/reference.pdf`;

        const candidateInspection = await inspectPdf(candidate.pdf);
        const referenceInspection = await inspectPdf(referencePdf);
        if (candidateInspection.pageCount !== verifiedCandidate.pageCount) {
          throw new Error(`${entry.id}: export API and independent PDF parser disagree on page count.`);
        }
        for (const [pageIndex, pageMapPage] of verifiedCandidate.pages.entries()) {
          const pdfPage = candidateInspection.pages[pageIndex];
          if (!pdfPage
            || Math.abs(pageMapPage.width - pdfPage.mediaBox.width) > PHYSICAL_GEOMETRY_TOLERANCE_POINTS
            || Math.abs(pageMapPage.height - pdfPage.mediaBox.height) > PHYSICAL_GEOMETRY_TOLERANCE_POINTS) {
            throw new Error(`${entry.id}: PageMap page ${pageIndex + 1} does not bind the PDF MediaBox.`);
          }
        }
        docxodusPages = candidateInspection.pageCount;
        libreofficePages = referenceInspection.pageCount;
        writeJsonAtomic(join(caseOutput, 'candidate-inspection.json'), candidateInspection);
        writeJsonAtomic(join(caseOutput, 'reference-inspection.json'), referenceInspection);
        artifacts.candidateInspection = `${entry.id}/candidate-inspection.json`;
        artifacts.referenceInspection = `${entry.id}/reference-inspection.json`;

        const semantic = semanticEvidence(
          tools.pdftotext.path,
          candidatePdfPath,
          candidateInspection,
          referencePdfPath,
          referenceInspection,
          entry,
        );
        writeJsonAtomic(join(caseOutput, 'semantic.json'), semantic);
        artifacts.semantic = `${entry.id}/semantic.json`;
        const vectorContentPassed = !(entry.categories as readonly string[]).includes('charts')
          || candidateInspection.vectorPathOperations > 0;
        writeJsonAtomic(join(caseOutput, 'vector-content.json'), {
          required: (entry.categories as readonly string[]).includes('charts'),
          passed: vectorContentPassed,
          candidateConstructPathOperations: candidateInspection.vectorPathOperations,
          referenceConstructPathOperations: referenceInspection.vectorPathOperations,
        });
        artifacts.vectorContent = `${entry.id}/vector-content.json`;

        const geometry = compareGeometry(candidateInspection, referenceInspection);
        const physicalGeometryPassed = candidateInspection.pageCount === referenceInspection.pageCount
          && geometry.every((pageGeometry) => pageGeometry.passed);
        writeJsonAtomic(join(caseOutput, 'physical-geometry.json'), {
          tolerancePoints: PHYSICAL_GEOMETRY_TOLERANCE_POINTS,
          pageCountDelta: candidateInspection.pageCount - referenceInspection.pageCount,
          passed: physicalGeometryPassed,
          pages: geometry,
        });
        artifacts.physicalGeometry = `${entry.id}/physical-geometry.json`;

        const candidateRaster = rasterizePdf(candidatePdfPath, join(caseOutput, 'docxodus'), {
          executablePath: tools.pdftoppm.path,
          env: popplerEnv(),
        });
        const referenceRaster = rasterizePdf(referencePdfPath, join(caseOutput, 'libreoffice'), {
          executablePath: tools.pdftoppm.path,
          env: popplerEnv(),
        });
        if (candidateRaster.contractSha256 !== referenceRaster.contractSha256
          || candidateRaster.contractSha256 !== PDF_RASTER_CONTRACT_SHA256) {
          throw new Error(`${entry.id}: candidate/reference raster contracts differ.`);
        }
        if (candidateRaster.pages.length !== candidateInspection.pageCount
          || referenceRaster.pages.length !== referenceInspection.pageCount) {
          throw new Error(`${entry.id}: PDF parser and rasterizer disagree on page count.`);
        }

        const environmentFingerprint = sha256(JSON.stringify({
          schemaVersion: 1,
          externalEnvironment: environment,
          rendererFingerprint: candidate.rendererFingerprint,
          sourceSha256,
          referenceSourceSha256,
          rasterContractSha256: PDF_RASTER_CONTRACT_SHA256,
        }));
        const comparablePages = Math.min(candidateRaster.pages.length, referenceRaster.pages.length);
        let severity: VisualSeverity = candidateRaster.pages.length === referenceRaster.pages.length
          && physicalGeometryPassed
          && semantic.passed
          ? 'close'
          : 'severe';
        const pages: PdfParityPageResult[] = [];
        for (let index = 0; index < comparablePages; index++) {
          const candidatePage = candidateRaster.pages[index];
          const referencePage = referenceRaster.pages[index];
          const comparison = compareImages(
            decodePng(readFileSync(candidatePage.path)),
            decodePng(readFileSync(referencePage.path)),
          );
          const overlayPath = join(caseOutput, `overlay-${index + 1}.png`);
          writeFileSync(overlayPath, encodePng(comparison.overlay));
          severity = worse(severity, comparison.metrics.severity);
          pages.push({
            page: index + 1,
            ...comparison.metrics,
            physicalGeometry: geometry[index],
            environmentFingerprint,
            rendererFingerprint: candidate.rendererFingerprint,
            artifacts: {
              candidate: relative(outputRoot, candidatePage.path),
              reference: relative(outputRoot, referencePage.path),
              overlay: relative(outputRoot, overlayPath),
            },
            artifactSha256: {
              candidate: candidatePage.sha256,
              reference: referencePage.sha256,
              overlay: sha256File(overlayPath),
            },
          });
        }

        const result: PdfParityCaseResult = {
          id: entry.id,
          path: entry.source.path,
          categories: entry.categories,
          rationale: entry.rationale,
          disposition: entry.disposition,
          profiles: entry.profiles,
          provenance: entry.source.provenance,
          sourceSha256,
          sourceGitBlob: entry.source.gitBlob,
          referenceSourceSha256,
          docxodusPages: candidateRaster.pages.length,
          libreofficePages: referenceRaster.pages.length,
          pageCountDelta: candidateRaster.pages.length - referenceRaster.pages.length,
          physicalGeometryPassed,
          semanticChecksPassed: semantic.passed,
          vectorContentPassed,
          vectorPathOperations: {
            candidate: candidateInspection.vectorPathOperations,
            reference: referenceInspection.vectorPathOperations,
          },
          severity,
          rendererFingerprint: candidate.rendererFingerprint,
          environmentFingerprint,
          artifacts,
          artifactSha256: {
            source: sourceSha256,
            referenceSource: referenceSourceSha256,
            candidatePdf: candidateRaster.pdfSha256,
            referencePdf: referenceRaster.pdfSha256,
          },
          semantic,
          pages,
        };
        cases.push(result);
        writeJsonAtomic(join(caseOutput, 'metrics.json'), result);
        console.log(`[${caseIndex + 1}/${corpus.length}] ${entry.id}: `
          + `${result.docxodusPages}/${result.libreofficePages} pages, ${result.severity} `
          + `(${result.disposition.kind}), semantics ${semantic.passed ? 'pass' : 'FAIL'}, `
          + `geometry ${physicalGeometryPassed ? 'pass' : 'FAIL'}`
          + (semantic.referenceExtractionHealthy
            ? ''
            : ' [reference extraction degraded — environment, not gated]'));
      } catch (error) {
        const failedReport = error && typeof error === 'object'
          && 'report' in error ? (error as { report?: unknown }).report : undefined;
        if (failedReport !== undefined) {
          writeJsonAtomic(join(caseOutput, 'failed-render-report.json'), failedReport);
          artifacts.failedRenderReport = `${entry.id}/failed-render-report.json`;
        }
        const result: PdfParityCaseResult = {
          id: entry.id,
          path: entry.source.path,
          categories: entry.categories,
          rationale: entry.rationale,
          disposition: entry.disposition,
          profiles: entry.profiles,
          provenance: entry.source.provenance,
          sourceSha256,
          sourceGitBlob: entry.source.gitBlob,
          referenceSourceSha256,
          docxodusPages,
          libreofficePages,
          pageCountDelta: docxodusPages - libreofficePages,
          physicalGeometryPassed: false,
          semanticChecksPassed: false,
          vectorContentPassed: false,
          severity: 'severe',
          artifacts,
          artifactSha256: {},
          pages: [],
          error: error instanceof Error ? error.message : String(error),
        };
        cases.push(result);
        writeJsonAtomic(join(caseOutput, 'metrics.json'), result);
        console.error(`[${caseIndex + 1}/${corpus.length}] ${entry.id}: ERROR ${result.error}`);
      } finally {
        rmSync(work, { recursive: true, force: true });
      }
      persist();
    }

    phase = 'summary';
    persist(true);
    const summaryPath = join(outputRoot, 'summary.json');
    const summary = JSON.parse(readFileSync(summaryPath, 'utf8'));

    phase = 'hard-gate';
    expect(cases.filter(hasHardParityFailure).map((entry) => ({
      id: entry.id,
      pageCountDelta: entry.pageCountDelta,
      physicalGeometryPassed: entry.physicalGeometryPassed,
      semanticChecksPassed: entry.semanticChecksPassed,
      vectorContentPassed: entry.vectorContentPassed,
      error: entry.error,
    })), 'generated-PDF conversion, page-count, physical-geometry, semantic, and chart-vector contracts '
      + 'are unconditional; raster severity and disposition cannot waive them').toEqual([]);

    phase = 'ratchet-comparison';
    const complete = !process.env.DOCXODUS_GENERATED_PDF_PARITY_FILTER;
    const updateRecord = process.env.DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD === '1';
    if (updateRecord && !complete) {
      throw new Error('Refusing to write a partial generated-PDF ratchet record; unset the filter.');
    }
    const existingRecord = readRecord(GENERATED_PDF_RATCHET_RECORD_FILE);
    if (!existingRecord) {
      throw new Error('Generated-PDF ratchet record is missing. Run a complete benchmark with '
        + 'DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD=1 and commit the numbers-only record.');
    }
    const comparison = compareToRecord(existingRecord, summary, {
      expectComplete: complete,
      updateRecordEnv: 'DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD',
    });
    writeJsonAtomic(join(outputRoot, 'ratchet-comparison.json'), comparison);
    console.log(`Generated-PDF ratchet [${comparison.status}]: ${comparison.message}`);
    if (!updateRecord && process.env.DOCXODUS_VISUAL_PARITY_RATCHET !== '0') {
      expect(comparison.status, comparison.message).toBe('ok');
    }

    phase = 'strict-gate';
    if (process.env.DOCXODUS_VISUAL_PARITY_STRICT === '1') {
      expect(cases.filter(gatesStrictRasterRun).map((entry) => ({
        id: entry.id,
        severity: entry.severity,
        disposition: entry.disposition.kind,
      })), 'strict generated-PDF parity additionally rejects renderer-owned severe raster '
        + 'differences').toEqual([]);
    }

    phase = 'evidence-stability';
    const finalSourceCommit = commandVersion(git.path, ['rev-parse', 'HEAD']);
    const finalSourceTree = commandVersion(git.path, ['rev-parse', 'HEAD^{tree}']);
    const finalWorkingTreeStatus = commandVersion(git.path, ['status', '--porcelain']);
    if (finalSourceCommit !== sourceCommit || finalSourceTree !== sourceTree
      || finalWorkingTreeStatus !== sourceWorkingTreeStatus) {
      throw new Error('Source commit, tree, or working-tree state changed during the benchmark.');
    }
    if (JSON.stringify(captureGeneratedPdfBuildEvidence(repoRoot)) !== JSON.stringify(build)
      || JSON.stringify(fontContractReport(repoRoot, tools.fcMatch.path)) !== JSON.stringify(fontContract)
      || sha256File(exportEntry) !== (environment.moduleEntries as Record<string, string>).exporter
      || sha256File(playwrightCoreEntry)
        !== (environment.moduleEntries as Record<string, string>).playwrightCore
      || sha256File(chromiumExecutablePath)
        !== (environment.chromiumExecutable as Record<string, string>).sha256) {
      throw new Error('The built exporter, browser assets, module entries, or Chromium changed during the run.');
    }
    for (const tool of [git, tools.fcMatch, tools.libreoffice, tools.pdftoppm, tools.pdftotext]) {
      if (sha256File(tool.path) !== tool.evidence.executableSha256) {
        throw new Error(`Pinned executable changed during the run: ${tool.evidence.command}`);
      }
    }

    if (updateRecord) {
      phase = 'record-update';
      assertRecordUpdateProvenance(summary);
      assertRecordUpdateProvenance({
        ...summary,
        gitCommit: finalSourceCommit,
        workingTreeDirty: finalWorkingTreeStatus.length > 0,
      });
      const record = buildRecord(summary, new Date().toISOString().slice(0, 10));
      record.description = 'Generated-PDF visual-fidelity regression ratchet (issue #443). '
        + 'Numbers only; complete PDFs, rasters, semantic evidence, geometry, hashes, and '
        + 'fingerprints live in the uploaded artifact. See npm/tests/visual-parity/README.md.';
      writeTextAtomic(GENERATED_PDF_RATCHET_RECORD_FILE, serializeRecord(record));
      console.log(`Generated-PDF ratchet record refreshed: ${GENERATED_PDF_RATCHET_RECORD_FILE}`);
    }

    phase = 'complete';
    state.status = 'passed';
    persist(true);
    console.log(`Generated-PDF parity viewer: ${join(outputRoot, 'index.html')}`);
  } catch (error) {
    const message = error instanceof Error ? error.stack ?? error.message : String(error);
    state.status = 'failed';
    state.failure = { phase, message };
    writeJsonAtomic(join(outputRoot, 'failure.json'), {
      schemaVersion: 1,
      failedAt: new Date().toISOString(),
      phase,
      message,
      remediation: 'Open index.html, inspect the retained per-case PDFs/rasters/reports, and rerun '
        + 'the documented generated-PDF parity command after addressing the named phase.',
      environment,
    });
    persist(cases.length === corpus.length);
    throw error;
  }
});
