import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page, type TestInfo } from '@playwright/test';
import {
  EXPECTED_PROFILE_COMMENTS,
  COMMENT_REVISION_TOKENS,
  generateLongMarginCommentDocx,
  generateReviewCommentOverlapDocx,
  generateReviewCommentStoriesDocx,
  REVIEW_PROFILE_TOKENS,
  STORY_PROFILE_TOKENS,
  type ExpectedComment,
} from './review-comment-profile-fixture.js';

type ReviewProfile = 'final' | 'original' | 'markup';
type CommentProfile = 'hidden' | 'inline' | 'endnotes' | 'margin';

interface RenderReport {
  status: 'complete';
  source: { rawPackageBytesDigest: string; byteLength: number };
  derivedProfileSource?: { rawPackageBytesDigest: string; byteLength: number };
  options: { reviewProfile: ReviewProfile; commentProfile: CommentProfile };
  warnings: Array<{ code: string; severity: string; phase: string; partUri?: string }>;
}

interface BrowserResult {
  html: string;
  pageCount: number;
  pageMap: {
    pages: unknown[];
    fragments: Array<{ anchorId?: string }>;
  };
  renderReport: RenderReport;
  warnings: Array<{ code: string }>;
}

interface BrowserFailure {
  name?: string;
  message: string;
  stack?: string;
  report?: unknown;
  [key: string]: unknown;
}

type ConversionOutcome =
  | { ok: true; result: BrowserResult }
  | { ok: false; error: BrowserFailure };

interface DomAudit {
  visibleText: string;
  paragraphs: string[];
  stories: Record<'body' | 'header' | 'footer' | 'footnote' | 'endnote', string>;
  revisions: Array<{
    tag: string;
    classes: string[];
    text: string;
    author: string | null;
    date: string | null;
    moveId: string | null;
    title: string | null;
  }>;
  comments: Array<{
    id: string;
    nodeId: string;
    classes: string[];
    author: string | null;
    date: string | null;
    parentId: string | null;
    resolved: string | null;
    text: string;
  }>;
  commentOrder: string[];
  rangeTextByComment: Record<string, string>;
  storyCommentIds: Record<'body' | 'header' | 'footer' | 'footnote' | 'endnote', string[]>;
  markers: number;
  highlights: number;
  commentsSections: number;
  marginNotes: number;
  inlineThreadCounts: Record<string, number>;
  format: { fontWeight: string; color: string; classes: string[] } | null;
}

const reviewProfiles: readonly ReviewProfile[] = ['final', 'original', 'markup'];
const commentProfiles: readonly CommentProfile[] = ['hidden', 'inline', 'endnotes', 'margin'];
const revisionAuthor = 'Profile Reviewer';
const revisionDate = '2026-08-16T12:34:56Z';
const __dirname = dirname(fileURLToPath(import.meta.url));
const matrixArtifactRoot = resolve(__dirname, '../test-artifacts/review-comment-profiles');

interface MatrixCaseRecord {
  id: string;
  reviewProfile: ReviewProfile;
  commentProfile: CommentProfile;
  status: 'pending' | 'passed' | 'failed';
  sourceDigest?: string;
  htmlDigest?: string;
  pageCount?: number;
  warningCodes?: string[];
  artifacts: Record<string, string>;
  failure?: { message: string; stack?: string };
}

function digest(value: Uint8Array | string): string {
  return createHash('sha256').update(value).digest('hex');
}

function normalize(value: string): string {
  return value.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

async function ready(page: Page): Promise<void> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
}

async function convert(
  page: Page,
  source: Uint8Array,
  reviewProfile: ReviewProfile,
  commentProfile: CommentProfile,
): Promise<ConversionOutcome> {
  return page.evaluate(async ({ bytes, options }) => {
    const api = (window as any).DocxodusStandalone;
    try {
      const result = await api.convertAfterCallerMutation(bytes, {
        ...options,
        documentVersion: 444,
        unsupportedContent: 'warn',
      });
      return { ok: true, result };
    } catch (caught) {
      const error = caught as any;
      let serialized: Record<string, unknown>;
      try {
        serialized = typeof error?.toJSON === 'function'
          ? error.toJSON()
          : {
              name: error?.name,
              message: error?.message ?? String(error),
              stack: error?.stack,
            };
      } catch {
        serialized = { message: String(error) };
      }
      return { ok: false, error: serialized };
    }
  }, {
    bytes: Array.from(source),
    options: { reviewProfile, commentProfile },
  }) as Promise<ConversionOutcome>;
}

function writeJson(path: string, value: unknown): void {
  writeText(path, `${JSON.stringify(value, null, 2)}\n`);
}

function writeText(path: string, value: string): void {
  const temporary = `${path}.tmp`;
  writeFileSync(temporary, value);
  renameSync(temporary, path);
}

function escapeHtml(value: unknown): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

async function attachPaths(
  testInfo: TestInfo,
  caseId: string,
  paths: Array<{ name: string; path: string; contentType: string }>,
): Promise<void> {
  for (const artifact of paths) {
    await testInfo.attach(`${caseId}-${artifact.name}`, {
      path: artifact.path,
      contentType: artifact.contentType,
    });
  }
}

function writeViewer(directory: string, caseId: string, files: readonly string[]): string {
  const items = files.map((name) => `<li><a href="${escapeHtml(name)}">${escapeHtml(name)}</a></li>`).join('\n');
  const viewer = `<!doctype html><meta charset="utf-8"><title>${escapeHtml(caseId)} review/comment evidence</title>
<style>body{font:15px system-ui;margin:2rem;max-width:76rem}li{margin:.45rem 0}iframe,img{width:100%;border:1px solid #bbb}iframe{height:70vh}</style>
<h1>${escapeHtml(caseId)}</h1><p>Generated before semantic assertions so a failing case remains inspectable.</p>
<ul>${items}</ul>
  ${files.includes('screenshot.png') ? '<h2>Rendered pages</h2><img src="screenshot.png" alt="Rendered pages">' : ''}
  ${files.includes('standalone.html') ? '<h2>Standalone HTML</h2><iframe sandbox src="standalone.html"></iframe>' : ''}`;
  const path = join(directory, 'view-artifacts.html');
  writeText(path, viewer);
  return path;
}

function persistMatrixGallery(cases: readonly MatrixCaseRecord[], status: 'running' | 'passed' | 'failed'): void {
  const expectedCaseIds = reviewProfiles.flatMap((reviewProfile) =>
    commentProfiles.map((commentProfile) => `${reviewProfile}-${commentProfile}`));
  const manifest = {
    schemaVersion: 1,
    status,
    expectedCaseIds,
    cases,
  };
  writeJson(join(matrixArtifactRoot, 'run.json'), {
    schemaVersion: 1,
    status,
    completed: cases.filter((entry) => entry.status !== 'pending').length,
    expected: expectedCaseIds.length,
  });
  writeJson(join(matrixArtifactRoot, 'matrix.json'), manifest);

  const byId = new Map(cases.map((entry) => [entry.id, entry]));
  const rows = expectedCaseIds.map((id) => {
    const entry = byId.get(id);
    const links = entry
      ? Object.entries(entry.artifacts).map(([label, path]) =>
          `<a href="${escapeHtml(path)}">${escapeHtml(label)}</a>`).join(' · ')
      : '';
    return `<tr><th>${escapeHtml(id)}</th><td>${escapeHtml(entry?.status ?? 'pending')}</td><td>${links}</td><td>${escapeHtml(entry?.failure?.message ?? '')}</td></tr>`;
  }).join('\n');
  writeText(join(matrixArtifactRoot, 'index.html'), `<!doctype html>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Review/comment profile matrix</title>
<style>body{font:15px/1.45 system-ui;margin:2rem;max-width:100rem}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:.5rem;text-align:left;vertical-align:top}.failed{color:#a00}</style>
<h1>Review/comment profile matrix</h1>
<p>Status: <strong class="${status === 'failed' ? 'failed' : ''}">${escapeHtml(status)}</strong>. Artifacts are written before assertions and paths are portable.</p>
<p><a href="matrix.json">Matrix manifest</a> · <a href="run.json">Run state</a></p>
<table><thead><tr><th>Case</th><th>Status</th><th>Artifacts</th><th>Failure</th></tr></thead><tbody>${rows}</tbody></table>`);
}

async function auditRenderedPage(page: Page, html: string): Promise<DomAudit> {
  await page.setContent(html, { waitUntil: 'load' });
  await page.evaluate(async () => { await document.fonts.ready; });
  return page.evaluate((formatToken) => {
    const clean = (value: string | null | undefined) =>
      (value ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    const text = (selector: string) => clean(Array.from(
      document.querySelectorAll<HTMLElement>(selector),
      (node) => node.innerText,
    ).join(' '));
    const storyFor = (node: Element): 'body' | 'header' | 'footer' | 'footnote' | 'endnote' => {
      if (node.closest('.page-header')) return 'header';
      if (node.closest('.page-footer')) return 'footer';
      if (node.closest('.page-footnotes')) return 'footnote';
      if (node.closest('[data-source-anchor-id^="p:en:"], [data-source-anchor-id^="en:en:"]'))
        return 'endnote';
      return 'body';
    };
    const revisionNodes = Array.from(document.querySelectorAll<HTMLElement>(
      'ins, del, .rev-format-change, .rev-row-ins, .rev-row-del, .rev-cell-ins, .rev-cell-del, .rev-cell-merge',
    ));
    const comments = Array.from(document.querySelectorAll<HTMLElement>(
      '[data-comment-node-id], [data-comment-id], [id^="comment-"]',
    )).map((node) => {
      const id = node.dataset.commentNodeId
        ?? node.dataset.commentId
        ?? node.id.match(/^comment-(?!ref-)(.+)$/)?.[1]
        ?? node.id.match(/^comment-ref-(.+)$/)?.[1]
        ?? '';
      return {
        id,
        nodeId: node.id,
        classes: Array.from(node.classList),
        author: node.dataset.author ?? null,
        date: node.dataset.date ?? null,
        parentId: node.dataset.commentParentId ?? null,
        resolved: node.dataset.commentStatus === 'resolved'
          ? 'true'
          : node.dataset.commentStatus === 'open'
            ? 'false'
            : node.dataset.commentStatus ?? null,
        text: clean(node.innerText),
      };
    }).filter((entry) => entry.id.length > 0);
    const commentOrder: string[] = [];
    for (const comment of comments) {
      if (!commentOrder.includes(comment.id)) commentOrder.push(comment.id);
    }
    const rangeTextByComment: Record<string, string> = {};
    for (const node of Array.from(document.querySelectorAll<HTMLElement>('.comment-highlight[data-comment-id]'))) {
      const id = node.dataset.commentId!;
      rangeTextByComment[id] = clean(`${rangeTextByComment[id] ?? ''} ${node.innerText}`);
    }
    const storyCommentSets = {
      body: new Set<string>(), header: new Set<string>(), footer: new Set<string>(),
      footnote: new Set<string>(), endnote: new Set<string>(),
    };
    for (const node of Array.from(document.querySelectorAll<HTMLElement>('[data-comment-id]'))) {
      const id = node.dataset.commentId;
      if (id) storyCommentSets[storyFor(node)].add(id);
    }

    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let formatElement: HTMLElement | null = null;
    while (walker.nextNode()) {
      if (walker.currentNode.textContent?.includes(formatToken)) {
        formatElement = walker.currentNode.parentElement;
        break;
      }
    }
    const formatStyle = formatElement ? getComputedStyle(formatElement) : null;
    return {
      visibleText: clean(document.body.innerText),
      paragraphs: Array.from(document.querySelectorAll<HTMLElement>('.page-box p'), (node) => clean(node.innerText)),
      stories: {
        body: text('.page-content'),
        header: text('.page-header'),
        footer: text('.page-footer'),
        footnote: text('.page-footnotes'),
        endnote: text('[data-source-anchor-id^="p:en:"], [data-source-anchor-id^="en:en:"]'),
      },
      revisions: revisionNodes.map((node) => ({
        tag: node.tagName.toLowerCase(),
        classes: Array.from(node.classList),
        text: clean(node.innerText),
        author: node.dataset.author ?? null,
        date: node.dataset.date ?? null,
        moveId: node.dataset.moveId ?? null,
        title: node.getAttribute('title'),
      })),
      comments,
      commentOrder,
      rangeTextByComment,
      storyCommentIds: Object.fromEntries(Object.entries(storyCommentSets)
        .map(([story, ids]) => [story, Array.from(ids)])),
      markers: document.querySelectorAll('.comment-marker').length,
      highlights: document.querySelectorAll('.comment-highlight').length,
      commentsSections: document.querySelectorAll('.comments-section').length,
      marginNotes: document.querySelectorAll('.page-comment-margin .comment-margin-note').length,
      inlineThreadCounts: Object.fromEntries(Array.from(
        document.querySelectorAll<HTMLElement>('.comment-inline-thread [data-comment-node-id]'),
      ).reduce((counts, node) => {
        const id = node.dataset.commentNodeId;
        if (id) counts.set(id, (counts.get(id) ?? 0) + 1);
        return counts;
      }, new Map<string, number>())),
      format: formatElement && formatStyle ? {
        fontWeight: formatStyle.fontWeight,
        color: formatStyle.color,
        classes: Array.from(formatElement.classList),
      } : null,
    };
  }, REVIEW_PROFILE_TOKENS.format) as Promise<DomAudit>;
}

function selectedTokens(profile: ReviewProfile): { present: string[]; absent: string[] } {
  const original = [
    REVIEW_PROFILE_TOKENS.original,
    REVIEW_PROFILE_TOKENS.deletedCommentTarget,
    ...Object.values(STORY_PROFILE_TOKENS).map((tokens) => tokens.original),
  ];
  const final = [
    REVIEW_PROFILE_TOKENS.final,
    REVIEW_PROFILE_TOKENS.insertedCommentTarget,
    ...Object.values(STORY_PROFILE_TOKENS).map((tokens) => tokens.final),
  ];
  if (profile === 'final') return { present: final, absent: original };
  if (profile === 'original') return { present: original, absent: final };
  return { present: [...original, ...final], absent: [] };
}

function assertReviewProfile(profile: ReviewProfile, audit: DomAudit): void {
  const tokens = selectedTokens(profile);
  expect(audit.visibleText).toContain(REVIEW_PROFILE_TOKENS.stable);
  for (const token of tokens.present) expect(audit.visibleText, `${profile} should contain ${token}`).toContain(token);
  for (const token of tokens.absent) expect(audit.visibleText, `${profile} should omit ${token}`).not.toContain(token);

  const sourceParagraph = audit.paragraphs.find((paragraph) =>
    paragraph.includes(REVIEW_PROFILE_TOKENS.moveSourceContext)) ?? '';
  const destinationParagraph = audit.paragraphs.find((paragraph) =>
    paragraph.includes(REVIEW_PROFILE_TOKENS.moveDestinationContext)) ?? '';
  expect(sourceParagraph.length).toBeGreaterThan(0);
  expect(destinationParagraph.length).toBeGreaterThan(0);
  expect(sourceParagraph.includes(REVIEW_PROFILE_TOKENS.move)).toBe(profile !== 'final');
  expect(destinationParagraph.includes(REVIEW_PROFILE_TOKENS.move)).toBe(profile !== 'original');

  expect(audit.stories.header).toContain(
    profile === 'original' ? STORY_PROFILE_TOKENS.header.original : STORY_PROFILE_TOKENS.header.final,
  );
  expect(audit.stories.footer).toContain(
    profile === 'original' ? STORY_PROFILE_TOKENS.footer.original : STORY_PROFILE_TOKENS.footer.final,
  );
  expect(audit.stories.footnote).toContain(
    profile === 'original' ? STORY_PROFILE_TOKENS.footnote.original : STORY_PROFILE_TOKENS.footnote.final,
  );
  expect(audit.stories.endnote).toContain(
    profile === 'original' ? STORY_PROFILE_TOKENS.endnote.original : STORY_PROFILE_TOKENS.endnote.final,
  );

  expect(audit.format).not.toBeNull();
  const numericWeight = Number.parseInt(audit.format!.fontWeight, 10);
  if (profile === 'original') {
    expect(numericWeight).toBeLessThan(600);
    expect(audit.format!.color).toBe('rgb(255, 0, 0)');
  } else {
    expect(numericWeight).toBeGreaterThanOrEqual(600);
    expect(audit.format!.color).toBe('rgb(0, 0, 255)');
  }

  if (profile !== 'markup') {
    expect(audit.revisions).toEqual([]);
    return;
  }
  const revision = (text: string) => audit.revisions.find((entry) => entry.text.includes(text));
  for (const token of [
    REVIEW_PROFILE_TOKENS.original,
    REVIEW_PROFILE_TOKENS.final,
    REVIEW_PROFILE_TOKENS.insertedCommentTarget,
    REVIEW_PROFILE_TOKENS.deletedCommentTarget,
    ...Object.values(STORY_PROFILE_TOKENS).flatMap(({ original, final }) => [original, final]),
  ]) {
    const entry = revision(token);
    expect(entry, `missing markup element for ${token}`).toBeDefined();
    expect(entry!.author).toBe(revisionAuthor);
    expect(entry!.date).toBe(revisionDate);
  }
  const moveFrom = revision(REVIEW_PROFILE_TOKENS.move);
  const moveEntries = audit.revisions.filter((entry) => entry.text.includes(REVIEW_PROFILE_TOKENS.move));
  expect(moveFrom).toBeDefined();
  expect(moveEntries).toHaveLength(2);
  expect(moveEntries).toContainEqual(expect.objectContaining({ moveId: '21' }));
  expect(moveEntries).toContainEqual(expect.objectContaining({ moveId: '23' }));
  expect(moveEntries.every((entry) => entry.author === revisionAuthor && entry.date === revisionDate)).toBe(true);
  const formatChange = audit.revisions.find((entry) => entry.classes.includes('rev-format-change'));
  expect(formatChange).toEqual(expect.objectContaining({ author: revisionAuthor, date: revisionDate }));
  expect(formatChange?.title).toContain('Bold added');
  expect(formatChange?.title).toContain('Color changed');
}

function metadataFor(audit: DomAudit, comment: ExpectedComment): DomAudit['comments'] {
  return audit.comments.filter((entry) => entry.id === comment.id);
}

function assertCommentProfile(
  profile: CommentProfile,
  reviewProfile: ReviewProfile,
  html: string,
  audit: DomAudit,
): void {
  if (profile === 'hidden') {
    expect(audit.markers).toBe(0);
    expect(audit.highlights).toBe(0);
    expect(audit.commentsSections).toBe(0);
    expect(audit.marginNotes).toBe(0);
    expect(audit.comments).toEqual([]);
    for (const comment of EXPECTED_PROFILE_COMMENTS) {
      expect(html).not.toContain(comment.body);
      expect(audit.visibleText).not.toContain(comment.body);
    }
    return;
  }

  expect(audit.markers).toBeGreaterThan(0);
  expect(audit.highlights).toBeGreaterThan(0);
  for (const comment of EXPECTED_PROFILE_COMMENTS) {
    expect(html, `${profile} must retain comment body ${comment.id}`).toContain(comment.body);
    const nodes = metadataFor(audit, comment);
    expect(nodes.length, `${profile} must expose comment ${comment.id}`).toBeGreaterThan(0);
    expect(nodes.some((node) => node.author === comment.author), `comment ${comment.id} author`).toBe(true);
    expect(nodes.some((node) => node.date === comment.date), `comment ${comment.id} date`).toBe(true);
    expect(nodes.some((node) => node.resolved === comment.resolved), `comment ${comment.id} resolved state`).toBe(true);
    if (comment.parentId) {
      expect(nodes.some((node) => node.parentId === comment.parentId), `comment ${comment.id} parent`).toBe(true);
    }
  }
  expect(audit.commentOrder.indexOf('0')).toBeLessThan(audit.commentOrder.indexOf('1'));

  const presentCommentRevision = reviewProfile === 'original'
    ? COMMENT_REVISION_TOKENS.original
    : COMMENT_REVISION_TOKENS.final;
  const absentCommentRevision = reviewProfile === 'final'
    ? COMMENT_REVISION_TOKENS.original
    : COMMENT_REVISION_TOKENS.final;
  expect(audit.visibleText).toContain(presentCommentRevision);
  if (reviewProfile === 'markup') {
    expect(audit.visibleText).toContain(COMMENT_REVISION_TOKENS.original);
    expect(audit.visibleText).toContain(COMMENT_REVISION_TOKENS.final);
    for (const token of Object.values(COMMENT_REVISION_TOKENS)) {
      const revision = audit.revisions.find((entry) => entry.text.includes(token));
      expect(revision, `missing comment-body markup for ${token}`).toEqual(expect.objectContaining({
        author: revisionAuthor,
        date: revisionDate,
      }));
    }
  } else {
    expect(audit.visibleText).not.toContain(absentCommentRevision);
  }

  if (profile === 'inline') {
    expect(audit.inlineThreadCounts).toEqual(Object.fromEntries(
      EXPECTED_PROFILE_COMMENTS.map((comment) => [comment.id, 1]),
    ));
  }

  const overlapText = audit.rangeTextByComment['0'] ?? '';
  expect(overlapText).toContain(
    reviewProfile === 'original' ? REVIEW_PROFILE_TOKENS.original : REVIEW_PROFILE_TOKENS.final,
  );
  if (reviewProfile === 'markup') expect(overlapText).toContain(REVIEW_PROFILE_TOKENS.original);
  else expect(overlapText).not.toContain(
    reviewProfile === 'final' ? REVIEW_PROFILE_TOKENS.original : REVIEW_PROFILE_TOKENS.final,
  );
  const insertionRange = audit.rangeTextByComment['2'] ?? '';
  const deletionRange = audit.rangeTextByComment['3'] ?? '';
  expect(insertionRange.includes(REVIEW_PROFILE_TOKENS.insertedCommentTarget)).toBe(reviewProfile !== 'original');
  expect(deletionRange.includes(REVIEW_PROFILE_TOKENS.deletedCommentTarget)).toBe(reviewProfile !== 'final');

  for (const [story, id] of [
    ['body', '10'], ['header', '11'], ['footer', '12'], ['footnote', '13'], ['endnote', '14'],
  ] as const) {
    expect(audit.storyCommentIds[story], `${profile} ${story} comment range`).toContain(id);
  }

  if (profile === 'inline') {
    expect(audit.commentsSections).toBe(0);
    expect(audit.marginNotes).toBe(0);
  } else if (profile === 'endnotes') {
    expect(audit.commentsSections).toBeGreaterThan(0);
    expect(audit.marginNotes).toBe(0);
    for (const comment of EXPECTED_PROFILE_COMMENTS) expect(audit.visibleText).toContain(comment.body);
  } else {
    expect(audit.commentsSections).toBe(0);
    expect(audit.marginNotes).toBeGreaterThanOrEqual(EXPECTED_PROFILE_COMMENTS.length);
    for (const comment of EXPECTED_PROFILE_COMMENTS) expect(audit.visibleText).toContain(comment.body);
  }
}

test('review/comment fixtures are byte deterministic', async ({}, testInfo) => {
  const overlap = generateReviewCommentOverlapDocx();
  const stories = generateReviewCommentStoriesDocx();
  expect(generateReviewCommentOverlapDocx()).toEqual(overlap);
  expect(generateReviewCommentStoriesDocx()).toEqual(stories);
  expect(digest(overlap)).not.toBe(digest(stories));
  const directory = testInfo.outputPath('review-comment-fixtures');
  mkdirSync(directory, { recursive: true });
  const overlapPath = join(directory, 'overlap.docx');
  const storiesPath = join(directory, 'all-stories.docx');
  writeFileSync(overlapPath, overlap);
  writeFileSync(storiesPath, stories);
  await attachPaths(testInfo, 'fixtures', [
    { name: 'overlap.docx', path: overlapPath, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    { name: 'all-stories.docx', path: storiesPath, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  ]);
});

test.describe('standalone review/comment acceptance matrix', () => {
  test('retains all 12 profile combinations in one portable evidence gallery', async ({ page }, testInfo) => {
    test.setTimeout(10 * 60 * 1000);
    rmSync(matrixArtifactRoot, { recursive: true, force: true });
    mkdirSync(matrixArtifactRoot, { recursive: true });

    const cases: MatrixCaseRecord[] = reviewProfiles.flatMap((reviewProfile) =>
      commentProfiles.map((commentProfile) => ({
        id: `${reviewProfile}-${commentProfile}`,
        reviewProfile,
        commentProfile,
        status: 'pending' as const,
        artifacts: {},
      })));
    const failures: Array<{ id: string; message: string }> = [];
    persistMatrixGallery(cases, 'running');

    for (const entry of cases) {
      const { id: caseId, reviewProfile, commentProfile } = entry;
      const directory = join(matrixArtifactRoot, caseId);
      const caseFiles = ['source.docx'];
      mkdirSync(directory, { recursive: true });
      const source = generateReviewCommentStoriesDocx();
      const sourceSnapshot = new Uint8Array(source);
      const sourceDigest = digest(source);
      writeFileSync(join(directory, 'source.docx'), source);
      entry.sourceDigest = sourceDigest;
      entry.artifacts.source = `${caseId}/source.docx`;
      persistMatrixGallery(cases, 'running');

      try {
        await ready(page);
        const outcome = await convert(page, source, reviewProfile, commentProfile);
        if (!outcome.ok) {
          writeJson(join(directory, 'conversion-error.json'), outcome.error);
          caseFiles.push('conversion-error.json');
          entry.artifacts.conversionError = `${caseId}/conversion-error.json`;
          throw new Error(`conversion failed: ${JSON.stringify(outcome.error)}`);
        }

        const { result } = outcome;
        writeFileSync(join(directory, 'standalone.html'), result.html);
        writeJson(join(directory, 'page-map.json'), result.pageMap);
        writeJson(join(directory, 'render-report.json'), result.renderReport);
        caseFiles.push('standalone.html', 'page-map.json', 'render-report.json');
        Object.assign(entry.artifacts, {
          html: `${caseId}/standalone.html`,
          pageMap: `${caseId}/page-map.json`,
          renderReport: `${caseId}/render-report.json`,
        });
        entry.htmlDigest = digest(result.html);
        entry.pageCount = result.pageCount;
        entry.warningCodes = result.renderReport.warnings.map((warning) => warning.code);

        const audit = await auditRenderedPage(page, result.html);
        writeJson(join(directory, 'semantic-audit.json'), audit);
        await page.screenshot({ path: join(directory, 'screenshot.png'), fullPage: true });
        caseFiles.push('semantic-audit.json', 'screenshot.png');
        Object.assign(entry.artifacts, {
          semanticAudit: `${caseId}/semantic-audit.json`,
          screenshot: `${caseId}/screenshot.png`,
        });
        writeViewer(directory, caseId, caseFiles);
        entry.artifacts.viewer = `${caseId}/view-artifacts.html`;

        // Assertions deliberately follow durable publication of the full case evidence.
        expect(source, 'Node-owned source fixture changed during browser conversion')
          .toEqual(sourceSnapshot);
        expect(result.pageCount).toBeGreaterThan(2);
        expect(result.pageMap.pages).toHaveLength(result.pageCount);
        expect(result.renderReport.status).toBe('complete');
        expect(result.renderReport.options)
          .toEqual(expect.objectContaining({ reviewProfile, commentProfile }));
        expect(result.renderReport.source).toEqual({
          rawPackageBytesDigest: sourceDigest,
          byteLength: source.byteLength,
          documentVersion: 444,
        });
        expect(result.renderReport.warnings.filter((warning) =>
          warning.code.includes('revision')
          || warning.code === 'fragment_target_unavailable'
          || warning.severity === 'error')).toEqual([]);
        if (reviewProfile === 'markup') {
          expect(result.renderReport.derivedProfileSource).toBeUndefined();
        } else {
          expect(result.renderReport.derivedProfileSource).toEqual({
            rawPackageBytesDigest: expect.stringMatching(/^[0-9a-f]{64}$/),
            byteLength: expect.any(Number),
          });
          expect(result.renderReport.derivedProfileSource!.byteLength).toBeGreaterThan(0);
        }

        assertReviewProfile(reviewProfile, audit);
        assertCommentProfile(commentProfile, reviewProfile, result.html, audit);
        entry.status = 'passed';
      } catch (caught) {
        const failure = {
          message: caught instanceof Error ? caught.message : String(caught),
          ...(caught instanceof Error && caught.stack ? { stack: caught.stack } : {}),
        };
        entry.status = 'failed';
        entry.failure = failure;
        failures.push({ id: caseId, message: failure.message });
        writeJson(join(directory, 'failure.json'), failure);
        if (!caseFiles.includes('failure.json')) caseFiles.push('failure.json');
        entry.artifacts.failure = `${caseId}/failure.json`;
        writeViewer(directory, caseId, caseFiles);
        entry.artifacts.viewer = `${caseId}/view-artifacts.html`;
      }
      persistMatrixGallery(cases, 'running');
    }

    const finalStatus = failures.length === 0 ? 'passed' : 'failed';
    persistMatrixGallery(cases, finalStatus);
    await attachPaths(testInfo, 'review-comment-profile-matrix', [
      { name: 'index.html', path: join(matrixArtifactRoot, 'index.html'), contentType: 'text/html' },
      { name: 'matrix.json', path: join(matrixArtifactRoot, 'matrix.json'), contentType: 'application/json' },
    ]);

    expect(cases).toHaveLength(reviewProfiles.length * commentProfiles.length);
    expect(new Set(cases.map((entry) => entry.id)).size).toBe(cases.length);
    expect(cases.every((entry) => entry.status !== 'pending')).toBe(true);
    expect(cases.filter((entry) => existsSync(
      join(matrixArtifactRoot, entry.id, 'view-artifacts.html')))).toHaveLength(cases.length);
    for (const entry of cases) {
      for (const path of Object.values(entry.artifacts)) {
        expect(isAbsolute(path), `${entry.id} artifact path must be portable`).toBe(false);
        const resolvedPath = resolve(matrixArtifactRoot, path);
        const local = relative(matrixArtifactRoot, resolvedPath);
        expect(local.startsWith('..') || isAbsolute(local), `${entry.id} artifact escaped root`)
          .toBe(false);
        expect(existsSync(resolvedPath), `${entry.id} artifact missing: ${path}`).toBe(true);
      }
    }
    expect(failures, failures.map((failure) => `${failure.id}: ${failure.message}`).join('\n'))
      .toEqual([]);
  });
});

test('oversized margin threads fail closed with retained evidence', async ({ page }, testInfo) => {
  const directory = resolve(__dirname, '../test-artifacts/review-comment-profile-failures/margin-overflow');
  rmSync(directory, { recursive: true, force: true });
  mkdirSync(directory, { recursive: true });
  const source = generateLongMarginCommentDocx();
  writeFileSync(join(directory, 'source.docx'), source);
  await ready(page);
  const outcome = await convert(page, source, 'markup', 'margin');
  writeJson(join(directory, outcome.ok ? 'unexpected-success.json' : 'expected-error.json'), outcome);
  writeViewer(directory, 'margin-overflow', [
    'source.docx', outcome.ok ? 'unexpected-success.json' : 'expected-error.json',
  ]);
  await attachPaths(testInfo, 'margin-overflow', [
    { name: 'viewer.html', path: join(directory, 'view-artifacts.html'), contentType: 'text/html' },
    { name: 'source.docx', path: join(directory, 'source.docx'), contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  ]);
  expect(outcome.ok).toBe(false);
  if (!outcome.ok) {
    expect(outcome.error.message).toContain('comment margin is clipped');
    expect(outcome.error).toEqual(expect.objectContaining({
      code: 'pagination_failure',
      phase: 'running_story_placement',
    }));
  }
});
