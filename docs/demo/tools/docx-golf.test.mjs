// Headless logic checks for DOCX GOLF (docs/demo/docx-golf.js) — the pure,
// DOM-free parts: course shape validation, the stroke counter, golf score
// naming, and the caddie's revision-list formatting. Everything that needs
// the engine (DocxDiff scoring, hole honesty, the reference solutions) is
// proven in the browser by npm/tests/demo-golf.spec.ts.
import assert from 'node:assert/strict';
import test from 'node:test';

import {
  COURSE,
  validateCourse,
  createStrokeCounter,
  scoreName,
  caddieLines,
} from '../docx-golf.js';

// ─── scoreName ────────────────────────────────────────────────────────

test('scoreName follows golf naming relative to par', () => {
  assert.equal(scoreName(1, 3), 'hole in one');
  assert.equal(scoreName(1, 1), 'hole in one');
  assert.equal(scoreName(2, 4), 'eagle');
  assert.equal(scoreName(3, 4), 'birdie');
  assert.equal(scoreName(4, 4), 'par');
  assert.equal(scoreName(5, 4), 'bogey');
  assert.equal(scoreName(6, 4), 'double bogey');
  assert.equal(scoreName(9, 4), '+5');
});

// ─── createStrokeCounter ──────────────────────────────────────────────

test('stroke counter counts state-change observations, not state deltas', () => {
  const counter = createStrokeCounter(5);
  assert.equal(counter.strokes(), 0);
  counter.observe(5); // unchanged — no stroke
  assert.equal(counter.strokes(), 0);
  counter.observe(6); // one committed edit
  assert.equal(counter.strokes(), 1);
  counter.observe(6); // still settled
  assert.equal(counter.strokes(), 1);
  // A burst of edits between two observations is one swing, however many
  // document states it passed through (a multi-block ribbon op is one gesture).
  counter.observe(9);
  assert.equal(counter.strokes(), 2);
  // Undo changes the document too — in golf every swing counts.
  counter.observe(8);
  assert.equal(counter.strokes(), 3);
});

// ─── caddieLines ──────────────────────────────────────────────────────

test('caddie phrases revisions as the work remaining, player → target', () => {
  const lines = caddieLines([
    { revisionType: 'Inserted', text: 'the Company' },
    { revisionType: 'Deleted', text: 'Acme' },
  ]);
  assert.deepEqual(lines, ['add "the Company"', 'remove "Acme"']);
});

test('caddie collapses a move pair into one hint', () => {
  const lines = caddieLines([
    { revisionType: 'Moved', text: 'Governing Law', moveGroupId: 1, isMoveSource: true },
    { revisionType: 'Moved', text: 'Governing Law', moveGroupId: 1, isMoveSource: false },
    { revisionType: 'Inserted', text: 'notice' },
  ]);
  assert.deepEqual(lines, ['move "Governing Law"', 'add "notice"']);
});

test('caddie names the changed properties of a format revision', () => {
  const lines = caddieLines([
    {
      revisionType: 'FormatChanged',
      text: 'Recitals',
      formatChange: { changedPropertyNames: ['StyleId', 'Bold'] },
    },
  ]);
  assert.deepEqual(lines, ['reformat "Recitals" (StyleId, Bold)']);
});

test('caddie truncates long revision text and returns [] when nothing remains', () => {
  const long = 'x'.repeat(80);
  const [line] = caddieLines([{ revisionType: 'Inserted', text: long }]);
  assert.ok(line.length < 60, `hint should be short, got ${line.length} chars`);
  assert.ok(line.includes('…'), 'truncation should be visible');
  assert.deepEqual(caddieLines([]), []);
});

// ─── validateCourse ───────────────────────────────────────────────────

const validHole = (over = {}) => ({
  id: 'h1',
  title: 'First tee',
  par: 1,
  brief: 'Fix the word.',
  start: ['Hello wrold.'],
  target: ['Hello world.'],
  solve: () => {},
  ...over,
});

test('a well-formed course validates clean', () => {
  assert.deepEqual(validateCourse([validHole()]), []);
});

test('validateCourse flags duplicate ids, bad par, empty brief, and missing solve', () => {
  const problems = validateCourse([
    validHole(),
    validHole({ par: 0, brief: '', solve: undefined }),
  ]);
  assert.ok(problems.some((p) => p.includes('duplicate')), `expected duplicate id problem in ${problems}`);
  assert.ok(problems.some((p) => p.includes('par')), `expected par problem in ${problems}`);
  assert.ok(problems.some((p) => p.includes('brief')), `expected brief problem in ${problems}`);
  assert.ok(problems.some((p) => p.includes('solve')), `expected solve problem in ${problems}`);
});

test('validateCourse rejects a hole whose start already equals its target', () => {
  const problems = validateCourse([validHole({ target: ['Hello wrold.'] })]);
  assert.ok(problems.some((p) => p.includes('already')), `expected already-solved problem in ${problems}`);
});

test('validateCourse accepts builder functions in place of paragraph lists', () => {
  const problems = validateCourse([
    validHole({ start: () => {}, target: () => {} }),
  ]);
  assert.deepEqual(problems, []);
});

// ─── The shipped course ───────────────────────────────────────────────

test('the shipped course validates clean and escalates across the surface', () => {
  assert.deepEqual(validateCourse(COURSE), []);
  assert.ok(COURSE.length >= 4, 'course should have at least four holes');
  assert.equal(COURSE[0].par, 1, 'hole 1 is the one-stroke teaching hole');
  // At least one hole must exercise heading styles, and one must be built
  // programmatically (the table hole) — the surface breadth is the point.
  assert.ok(
    COURSE.some((h) => Array.isArray(h.target) && h.target.some((p) => /^#{1,3} /.test(p))),
    'a hole should target markdown heading styles',
  );
  assert.ok(
    COURSE.some((h) => typeof h.start === 'function'),
    'a hole should build its documents programmatically (tables)',
  );
});
