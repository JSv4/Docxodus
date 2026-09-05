// Headless logic checks for DIFF STRESS (docs/demo/diff-stress.js) — the pure,
// DOM-free parts: the frame-time meter, the sparkline geometry, the pipeline
// table, and the two ratios the mode exists to report. The loop itself needs the
// comparison engine and is proven in the browser by npm/tests/demo-redline.spec.ts.
import assert from 'node:assert/strict';
import test from 'node:test';

import {
  CLAUSE_COUNT,
  PIPELINES,
  REDLINE_HTML_OPTIONS,
  clauseFor,
  costRatio,
  createFpsMeter,
  frameBudget,
  msPerKb,
  pipelineById,
  sparklinePath,
} from '../diff-stress.js';

// ─── The meter ────────────────────────────────────────────────────────

test('fps comes from the MEDIAN frame, so one hitch does not halve the reading', () => {
  const meter = createFpsMeter(10);
  for (let i = 0; i < 9; i++) meter.record(100); // a steady 10fps
  assert.equal(Math.round(meter.fps()), 10);
  meter.record(900); // one stall
  // A mean would report ~5.6fps; the median still says the run is a 10fps run.
  assert.equal(Math.round(meter.fps()), 10);
  // …but the instantaneous readout does show the stall.
  assert.ok(meter.instantFps() < 2);
  assert.equal(meter.lastMs(), 900);
});

test('the window slides but the total and count do not', () => {
  const meter = createFpsMeter(3);
  for (const ms of [10, 20, 30, 40, 50]) meter.record(ms);
  assert.deepEqual(meter.frames(), [30, 40, 50], 'only the window is kept');
  assert.equal(meter.count(), 5, 'every frame is still counted');
  assert.equal(meter.totalMs(), 150, 'and still totalled');
});

test('reset makes a run its own experiment', () => {
  // Without this, switching pipeline depth reports a p50 describing neither one.
  const meter = createFpsMeter(30);
  for (let i = 0; i < 5; i++) meter.record(200);
  meter.reset();
  assert.equal(meter.count(), 0);
  assert.equal(meter.fps(), 0);
  assert.deepEqual(meter.frames(), []);
  meter.record(100);
  assert.equal(Math.round(meter.fps()), 10);
});

test('an empty meter reports zeroes rather than NaN or Infinity', () => {
  const meter = createFpsMeter();
  assert.equal(meter.fps(), 0);
  assert.equal(meter.instantFps(), 0);
  assert.equal(meter.lastMs(), 0);
  assert.equal(meter.percentile(50), 0);
});

test('percentiles are nearest-rank over the window', () => {
  const meter = createFpsMeter(10);
  for (const ms of [50, 10, 40, 20, 30]) meter.record(ms);
  assert.equal(meter.percentile(50), 30);
  assert.equal(meter.percentile(100), 50);
});

// ─── Budget and ratios ────────────────────────────────────────────────

test('frameBudget names what the viewer actually perceives', () => {
  assert.equal(frameBudget(0).id, 'idle');
  assert.equal(frameBudget(40).id, 'fluid');
  assert.equal(frameBudget(180).id, 'brisk');
  assert.equal(frameBudget(515).id, 'chunky');
  assert.equal(frameBudget(1200).id, 'slideshow');
});

test('costRatio is the number the whole mode exists to report', () => {
  // Measured shape: a ~2ms mutation against a ~324ms computed redline.
  assert.equal(Math.round(costRatio(324, 2)), 162);
  // A zero denominator reports 0 rather than Infinity — the panel renders it.
  assert.equal(costRatio(324, 0), 0);
  assert.equal(costRatio(324, undefined), 0);
});

test('msPerKb normalises frame cost against document size', () => {
  assert.equal(msPerKb(200, 1024), 200);
  assert.equal(msPerKb(200, 2048), 100);
  assert.equal(msPerKb(200, 0), 0);
});

// ─── Sparkline geometry ───────────────────────────────────────────────

test('sparklinePath spans the full width and scales to its own peak', () => {
  const d = sparklinePath([100, 200, 100], 300, 46);
  assert.ok(d.startsWith('M0.0,'), `starts at x=0: ${d}`);
  const points = d.split(/[ML]/).filter(Boolean).map((p) => p.split(','));
  assert.equal(points.length, 3);
  assert.equal(Number(points[2][0]), 300, 'last point reaches the full width');
  // The peak sits at the top of the box, the troughs below it.
  const ys = points.map((p) => Number(p[1]));
  assert.ok(ys[1] < ys[0], 'the tallest frame is drawn highest');
  assert.ok(ys.every((y) => y >= 0 && y <= 46), `inside the viewBox: ${ys}`);
});

test('sparklinePath declines to draw a line through fewer than two points', () => {
  assert.equal(sparklinePath([], 300, 46), '');
  assert.equal(sparklinePath([100], 300, 46), '');
  assert.equal(sparklinePath(undefined, 300, 46), '');
});

test('a flat trace still renders inside the box', () => {
  // All-equal values must not divide by a zero range.
  const d = sparklinePath([200, 200, 200], 300, 46);
  const ys = d.split(/[ML]/).filter(Boolean).map((p) => Number(p.split(',')[1]));
  assert.ok(ys.every((y) => Number.isFinite(y) && y >= 0 && y <= 46), `finite: ${ys}`);
});

// ─── Pipelines ────────────────────────────────────────────────────────

test('the three depths are ordered cheapest-first and name their engine calls', () => {
  assert.deepEqual(PIPELINES.map((p) => p.id), ['revisions', 'redline', 'full']);
  for (const pipeline of PIPELINES) {
    assert.ok(pipeline.label, `${pipeline.id} needs a label`);
    assert.ok(pipeline.title, `${pipeline.id} needs a tooltip`);
    assert.ok(pipeline.calls.length >= 1, `${pipeline.id} must name its engine calls`);
    for (const call of pipeline.calls) {
      assert.match(call, /^(docxDiff|convertDocx)/, `${call} is not an engine call`);
    }
  }
  // The full pipeline is the redline pipeline plus a render, not a different one.
  assert.ok(PIPELINES[2].calls.includes('convertDocxToHtml'));
  assert.ok(PIPELINES[2].calls.length > PIPELINES[1].calls.length);
});

test('an unknown depth falls back to the cheapest rather than throwing', () => {
  assert.equal(pipelineById('nope').id, 'revisions');
  assert.equal(pipelineById(undefined).id, 'revisions');
  assert.equal(pipelineById('full').id, 'full');
});

test('the redline render asks for the markup explicitly', () => {
  // convertDocxToHtml ACCEPTS revisions by default (it runs RevisionAccepter), so
  // rendering a redline with default options shows a clean document and no
  // redline at all — the exact trap docs/demo/docx-golf.js documents.
  assert.equal(REDLINE_HTML_OPTIONS.renderTrackedChanges, true);
  assert.equal(REDLINE_HTML_OPTIONS.showDeletedContent, true);
});

// ─── The appended clauses ─────────────────────────────────────────────

test('every appended clause is distinct, so alignment is not a freebie', () => {
  // The diff engine aligns text; appending the same paragraph repeatedly would
  // be an unrealistically easy alignment problem and would flatter the numbers.
  const seen = new Set();
  for (let i = 0; i < CLAUSE_COUNT * 3; i++) {
    const clause = clauseFor(i);
    assert.ok(!seen.has(clause), `clause ${i} repeats: ${clause}`);
    seen.add(clause);
    assert.ok(clause.length > 40, 'a clause should be a real sentence');
  }
});

test('clauses are numbered from their frame index', () => {
  assert.ok(clauseFor(0).startsWith('1.'));
  assert.ok(clauseFor(9).startsWith('10.'));
});
