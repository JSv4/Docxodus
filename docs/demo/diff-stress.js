// DIFF STRESS — run the comparison engine flat out, and report what it actually does.
//
// The theater beside this file produces its redline by RECORDING: each MCP call
// writes `w:ins`/`w:del` into the live package as it lands, which is a few
// milliseconds of in-memory OOXML work. This mode does the other thing — it
// COMPUTES the redline from scratch after every single edit, baseline against
// current, and measures how fast that can possibly go.
//
// The honest headline is that it is not 60fps and it was never going to be.
// Measured on the theater's own ~3KB agreement in a browser: a mutation is ~4ms,
// `docxDiffGetRevisions` is ~160ms, `docxDiffCompare` ~280ms, and adding the
// tracked-changes HTML render puts a full frame around ~420ms. So the comparison
// engine runs roughly 40x the cost of the mutation path, and a redline-per-edit
// loop tops out somewhere between 2 and 6 frames per second.
//
// That gap IS the demo. Both pipelines emit the same native markup, and the
// stress meter shows you what each one costs to get there — which is the whole
// argument for recording a redline when you are the one making the edits, and
// for computing one when you are handed a document somebody else changed.
//
// The document GROWS while the loop runs (each frame adds a clause), so the
// frame time climbs visibly with document size instead of sitting at whatever
// one fixed input happened to cost. A stress test on a fixed 3KB input would
// tell you almost nothing.
//
// Import-safe under Node: the meter, the sparkline, the pipeline table and the
// budget classifier are pure and tested by docs/demo/tools/diff-stress.test.mjs.
// The loop needs the engine and is proven by npm/tests/demo-redline.spec.ts.

// ─── The pipelines (pure data) ────────────────────────────────────────
//
// Three depths, because "how fast is the diff engine" has three different
// honest answers depending on what you need out of it. Each names the engine
// calls it makes so the readout can say what it just timed.

export const PIPELINES = [
  {
    id: 'revisions',
    label: 'revisions',
    title: 'docxDiffGetRevisions — the anchor-addressed revision list, no package built',
    calls: ['docxDiffGetRevisions'],
  },
  {
    id: 'redline',
    label: 'redline',
    title: 'docxDiffCompareProducts — one memoized pass yielding the redline package '
      + 'AND its revision list',
    calls: ['docxDiffCompareProducts'],
  },
  {
    id: 'full',
    label: 'full + HTML',
    title: 'The whole pipeline: compare, then render the tracked changes to HTML',
    calls: ['docxDiffCompareProducts', 'convertDocxToHtml'],
  },
];

export function pipelineById(id) {
  return PIPELINES.find((p) => p.id === id) ?? PIPELINES[0];
}

/** Options passed to the HTML render. The converter ACCEPTS revisions by default
 *  (it runs RevisionAccepter), so rendering a redline with defaults shows a clean
 *  document and no redline at all. Ask for the markup explicitly. */
export const REDLINE_HTML_OPTIONS = {
  renderTrackedChanges: true,
  showDeletedContent: true,
  renderMoveOperations: true,
};

// ─── The meter (pure) ─────────────────────────────────────────────────

/**
 * Frame-time meter over a rolling window.
 *
 * `fps` is derived from the MEDIAN frame in the window rather than the mean:
 * one 900ms hitch should not halve a reading that is otherwise steady, and the
 * median is what a viewer perceives as the rate. `instantFps` is the last frame
 * alone, for the twitchy readout.
 */
export function createFpsMeter(windowSize = 30) {
  let frames = [];
  let total = 0;
  let count = 0;
  return {
    /** Clear the window. Each run is its own experiment: carrying frames across
     *  a depth change would report a p50 that describes neither pipeline. */
    reset() {
      frames = [];
      total = 0;
      count = 0;
    },
    record(ms) {
      frames.push(ms);
      if (frames.length > windowSize) frames.shift();
      total += ms;
      count++;
    },
    frames: () => [...frames],
    count: () => count,
    totalMs: () => total,
    instantFps: () => (frames.length ? 1000 / frames[frames.length - 1] : 0),
    lastMs: () => (frames.length ? frames[frames.length - 1] : 0),
    fps() {
      if (!frames.length) return 0;
      const sorted = [...frames].sort((a, b) => a - b);
      return 1000 / sorted[Math.floor(sorted.length / 2)];
    },
    percentile(p) {
      if (!frames.length) return 0;
      const sorted = [...frames].sort((a, b) => a - b);
      return sorted[Math.max(0, Math.ceil((p / 100) * sorted.length) - 1)];
    },
  };
}

/** Classify a frame time the way a viewer experiences it. The thresholds are
 *  about perception, not about a 16ms budget nobody is going to hit here. */
export function frameBudget(ms) {
  if (ms <= 0) return { id: 'idle', label: '—' };
  if (ms < 100) return { id: 'fluid', label: 'fluid' };
  if (ms < 250) return { id: 'brisk', label: 'brisk' };
  if (ms < 600) return { id: 'chunky', label: 'chunky' };
  return { id: 'slideshow', label: 'slideshow' };
}

/** An SVG polyline path for the frame-time trace, scaled to the tallest frame
 *  in the window so the shape of the trend survives any absolute magnitude. */
export function sparklinePath(values, width, height, pad = 1) {
  if (!values || values.length < 2) return '';
  const peak = Math.max(...values, 1);
  const span = Math.max(1, values.length - 1);
  const usable = Math.max(1, height - pad * 2);
  return values
    .map((value, i) => {
      const x = (i / span) * width;
      const y = height - pad - (value / peak) * usable;
      return `${i === 0 ? 'M' : 'L'}${x.toFixed(1)},${y.toFixed(1)}`;
    })
    .join(' ');
}

/** Cost per kilobyte of document — the number that says whether the engine is
 *  scaling with the input or falling over. */
export function msPerKb(frameMs, bytes) {
  const kb = (bytes ?? 0) / 1024;
  if (kb <= 0) return 0;
  return frameMs / kb;
}

/** The comparison the whole mode exists to make: how many times more expensive
 *  is computing the redline than recording it? */
export function costRatio(diffMs, mutateMs) {
  if (!mutateMs || mutateMs <= 0) return 0;
  return diffMs / mutateMs;
}

// ─── The clauses the loop appends ─────────────────────────────────────
//
// Real contract sentences rather than lorem, because the diff engine aligns
// text and repeated identical paragraphs would be an unrealistically easy
// alignment problem. Each frame takes the next one and stamps its index, so no
// two inserted paragraphs are identical.

const CLAUSE_BODIES = [
  'Supplier shall maintain commercial general liability insurance of not less than '
    + '$5,000,000 per occurrence throughout the term.',
  'Neither party may assign this Agreement without the prior written consent of the other, '
    + 'except to a successor in interest by merger.',
  'Customer shall provide Supplier with reasonable access to systems and personnel '
    + 'necessary for performance of the Services.',
  'Supplier shall retain records relating to the Services for a period of seven (7) years '
    + 'following termination.',
  'Any dispute arising under this Agreement shall first be referred to the parties’ '
    + 'respective executive sponsors for resolution.',
  'The parties shall review the fee schedule annually and may adjust it by mutual '
    + 'written agreement.',
  'Supplier warrants that the Deliverables will conform in all material respects to the '
    + 'specifications for ninety (90) days after acceptance.',
  'Each party shall comply with all applicable anti-bribery and anti-corruption laws in '
    + 'the performance of this Agreement.',
];

/** The clause text for frame `n` — cycled and numbered, so the document grows
 *  with plausible, non-identical content. */
export function clauseFor(n) {
  const body = CLAUSE_BODIES[n % CLAUSE_BODIES.length];
  return `${n + 1}.${(n % 9) + 1} ${body}`;
}

export const CLAUSE_COUNT = CLAUSE_BODIES.length;

// ─── The panel (page chrome) ──────────────────────────────────────────

const STRESS_CSS = `
.dxs { display: flex; flex-direction: column; gap: 9px; padding: 10px 12px; min-height: 0; }
.dxs * { box-sizing: border-box; }
.dxs-row { display: flex; flex-direction: column; gap: 6px; }
.dxs-go { width: 100%; font: 600 12px/1 system-ui, sans-serif; padding: 9px 12px;
  border: 1px solid #dc2626; border-radius: 8px; background: #dc2626; color: #fff;
  cursor: pointer; white-space: nowrap; }
.dxs-go[data-running="true"] { background: #1e293b; border-color: #475569; }
.dxs-depths { display: flex; gap: 4px; width: 100%; }
.dxs-depths button { flex: 1; min-width: 0; font: 600 10.5px/1 "SF Mono", Consolas, monospace;
  padding: 7px 4px; border: 1px solid #334155; border-radius: 7px; background: #1e293b;
  color: #cbd5e1; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.dxs-depths button[aria-pressed="true"] { background: #f87171; border-color: #f87171; color: #0f172a; }

.dxs-big { display: grid; grid-template-columns: 1fr 1fr; gap: 1px; background: #1e293b;
  border: 1px solid #1e293b; border-radius: 9px; overflow: hidden; }
.dxs-big > div { background: #0b1220; padding: 10px 8px; text-align: center; }
.dxs-big b { display: block; font: 700 22px/1.1 "SF Mono", Consolas, monospace; color: #f8fafc; }
.dxs-big span { font-size: 9.5px; text-transform: uppercase; letter-spacing: .07em; color: #64748b; }
.dxs-big .dxs-verdict { font: 700 10px/1 "SF Mono", Consolas, monospace; margin-top: 3px; }
.dxs-verdict[data-b="fluid"] { color: #4ade80; }
.dxs-verdict[data-b="brisk"] { color: #a3e635; }
.dxs-verdict[data-b="chunky"] { color: #fbbf24; }
.dxs-verdict[data-b="slideshow"] { color: #f87171; }
.dxs-verdict[data-b="idle"] { color: #475569; }

.dxs-spark { background: #020617; border: 1px solid #1e293b; border-radius: 9px; padding: 6px 8px; }
.dxs-spark svg { display: block; width: 100%; height: 46px; }
.dxs-spark .dxs-cap { display: flex; justify-content: space-between;
  font: 9.5px/1.4 "SF Mono", Consolas, monospace; color: #64748b; margin-top: 2px; }

.dxs-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 1px;
  background: #1e293b; border: 1px solid #1e293b; border-radius: 9px; overflow: hidden; }
.dxs-grid > div { background: #0f172a; padding: 7px 9px; }
.dxs-grid b { display: block; font: 700 13px/1.2 "SF Mono", Consolas, monospace; color: #e2e8f0; }
.dxs-grid span { font-size: 9.5px; text-transform: uppercase; letter-spacing: .06em; color: #64748b; }

.dxs-ratio { border: 1px solid #7c2d12; background: #1c1008; border-radius: 9px;
  padding: 9px 11px; font-size: 11.5px; color: #fdba74; }
.dxs-ratio b { color: #fb923c; font: 700 15px/1 "SF Mono", Consolas, monospace; }
.dxs-note { font-size: 10.5px; color: #64748b; line-height: 1.45; }
.dxs-note code { color: #94a3b8; }
`;

/** Build the stress panel inside `root`. Pure DOM construction. */
export function mountStressPanel(root) {
  if (!root.ownerDocument.getElementById('dxs-style')) {
    const style = root.ownerDocument.createElement('style');
    style.id = 'dxs-style';
    style.textContent = STRESS_CSS;
    root.ownerDocument.head.appendChild(style);
  }

  root.classList.add('dxs');
  root.innerHTML = `
    <div class="dxs-row">
      <button class="dxs-go" data-dxs="go">▶ Run diff stress</button>
      <div class="dxs-depths" data-dxs="depths" role="group" aria-label="Pipeline depth"></div>
    </div>
    <div class="dxs-big">
      <div><b data-dxs="fps">–</b><span>diffs / sec</span>
        <div class="dxs-verdict" data-dxs="verdict" data-b="idle">—</div></div>
      <div><b data-dxs="frame">–</b><span>ms / diff</span>
        <div class="dxs-verdict" data-dxs="depthlabel">revisions</div></div>
    </div>
    <div class="dxs-spark">
      <svg data-dxs="svg" viewBox="0 0 300 46" preserveAspectRatio="none" aria-hidden="true">
        <path data-dxs="path" fill="none" stroke="#f87171" stroke-width="1.5"
          stroke-linejoin="round" stroke-linecap="round" />
      </svg>
      <div class="dxs-cap"><span data-dxs="sparkleft">frame time</span>
        <span data-dxs="sparkright">peak –</span></div>
    </div>
    <div class="dxs-grid">
      <div><b data-dxs="frames">0</b><span>diffs run</span></div>
      <div><b data-dxs="revisions">0</b><span>revisions found</span></div>
      <div><b data-dxs="size">–</b><span>document KB</span></div>
      <div><b data-dxs="perkb">–</b><span>ms per KB</span></div>
    </div>
    <div class="dxs-ratio" data-dxs="ratio">
      Computing a redline costs <b data-dxs="ratiox">–</b> what recording one does.
    </div>
    <p class="dxs-note">Every frame appends a clause, then compares the live document
      against the untouched baseline from scratch — so the input grows under the engine and
      the frame time climbs with it. The editor keeps showing the <em>recorded</em> markup;
      this meter times the <em>computed</em> one.</p>`;

  const grab = (name) => root.querySelector(`[data-dxs="${name}"]`);
  return {
    ui: {
      panel: root, go: grab('go'), depths: grab('depths'),
      fps: grab('fps'), verdict: grab('verdict'), frame: grab('frame'),
      depthlabel: grab('depthlabel'), path: grab('path'),
      sparkleft: grab('sparkleft'), sparkright: grab('sparkright'),
      frames: grab('frames'), revisions: grab('revisions'),
      size: grab('size'), perkb: grab('perkb'),
      ratio: grab('ratio'), ratiox: grab('ratiox'),
    },
  };
}

// ─── The loop ─────────────────────────────────────────────────────────

const SPARK_W = 300;
const SPARK_H = 46;

/**
 * Run the comparison engine once against the live document, at `depth`.
 *
 * Returns the measured cost and what the pass produced. Split out from the loop
 * so the spec can time a single frame without driving the UI.
 */
export async function runDiffFrame({ engine, baseline, current, depth }) {
  const pipeline = pipelineById(depth);
  const started = performance.now();
  let revisions = [];
  let html = null;

  if (pipeline.id === 'revisions') {
    revisions = await engine.docxDiffGetRevisions(baseline, current);
  } else {
    // One memoized alignment pass yielding both products, rather than paying
    // for the alignment twice (issue #594) — measured at ~20% cheaper than
    // docxDiffCompare + docxDiffGetRevisions as separate calls.
    const products = await engine.docxDiffCompareProducts(
      baseline, current, undefined, ['redline', 'revisions']);
    revisions = products.revisions ?? [];
    if (pipeline.id === 'full' && products.redline) {
      html = await engine.convertDocxToHtml(products.redline, REDLINE_HTML_OPTIONS);
    }
  }

  return { ms: performance.now() - started, revisions, html, pipeline };
}

/**
 * Drive the stress loop against the live session.
 *
 * `applyEdit` appends one clause and returns the milliseconds that mutation
 * cost — the theater passes a function that dispatches a real MCP call, so the
 * mutation half of the comparison is measured through the same path the rest of
 * the demo uses rather than through a shortcut.
 *
 * `onFrame` fires after each measured frame so the host can repaint the editor
 * and, at `full` depth, show the computed redline.
 */
export function createStressRunner({ engine, session, ui, applyEdit, onFrame, maxFrames = 40 }) {
  const meter = createFpsMeter(30);
  let depth = PIPELINES[0].id;
  let running = false;
  let stopRequested = false;
  let baseline = null;
  let mutateMsTotal = 0;
  let mutateCount = 0;
  let lastRevisions = 0;
  let lastSize = 0;
  let startBytes = 0;

  const depthButtons = PIPELINES.map((pipeline) => {
    const b = ui.panel.ownerDocument.createElement('button');
    b.textContent = pipeline.label;
    b.title = pipeline.title;
    b.dataset.depth = pipeline.id;
    b.setAttribute('aria-pressed', String(pipeline.id === depth));
    b.addEventListener('click', () => setDepth(pipeline.id));
    ui.depths.appendChild(b);
    return b;
  });

  function setDepth(id) {
    depth = pipelineById(id).id;
    for (const b of depthButtons) {
      b.setAttribute('aria-pressed', String(b.dataset.depth === depth));
    }
    ui.depthlabel.textContent = pipelineById(depth).label;
  }

  function paint() {
    const fps = meter.fps();
    const last = meter.lastMs();
    ui.fps.textContent = fps > 0 ? fps.toFixed(1) : '–';
    ui.frame.textContent = last > 0 ? Math.round(last) : '–';
    const budget = frameBudget(last);
    ui.verdict.textContent = budget.label;
    ui.verdict.dataset.b = budget.id;
    ui.frames.textContent = String(meter.count());
    ui.revisions.textContent = String(lastRevisions);
    ui.size.textContent = lastSize ? (lastSize / 1024).toFixed(1) : '–';
    ui.perkb.textContent = lastSize ? msPerKb(last, lastSize).toFixed(1) : '–';

    const frames = meter.frames();
    ui.path.setAttribute('d', sparklinePath(frames, SPARK_W, SPARK_H));
    ui.sparkright.textContent = frames.length ? `peak ${Math.round(Math.max(...frames))}ms` : 'peak –';

    const avgMutate = mutateCount ? mutateMsTotal / mutateCount : 0;
    const ratio = costRatio(last, avgMutate);
    ui.ratiox.textContent = ratio > 0 ? `${ratio.toFixed(0)}×` : '–';
  }

  async function loop() {
    running = true;
    stopRequested = false;
    ui.go.textContent = '■ Stop';
    ui.go.dataset.running = 'true';
    // A run is a self-contained experiment: the meter, the mutation average and
    // the baseline all start fresh, so the readout describes THIS pipeline over
    // THIS document rather than a blend of every run since the page loaded.
    meter.reset();
    mutateMsTotal = 0;
    mutateCount = 0;
    try {
      baseline = session.save();
      startBytes = baseline.length;
      lastSize = baseline.length;
      for (let frame = 0; frame < maxFrames && !stopRequested; frame++) {
        // The mutation half, timed through the same path the theater uses.
        const mutateMs = await applyEdit(frame);
        mutateMsTotal += mutateMs;
        mutateCount++;

        const current = session.save();
        lastSize = current.length;
        const result = await runDiffFrame({ engine, baseline, current, depth });
        meter.record(result.ms);
        lastRevisions = result.revisions.length;
        paint();
        if (onFrame) onFrame(result);
        // Yield so the editor can repaint and the Stop button can be pressed.
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    } finally {
      running = false;
      ui.go.textContent = '▶ Run diff stress';
      ui.go.dataset.running = 'false';
      paint();
    }
  }

  ui.go.addEventListener('click', () => {
    if (running) { stopRequested = true; return; }
    void loop();
  });

  setDepth(depth);
  paint();

  return {
    run: loop,
    stop: () => { stopRequested = true; },
    setDepth,
    isRunning: () => running,
    stats: () => ({
      frames: meter.count(),
      fps: meter.fps(),
      lastMs: meter.lastMs(),
      p50: meter.percentile(50),
      p95: meter.percentile(95),
      totalMs: meter.totalMs(),
      revisions: lastRevisions,
      documentBytes: lastSize,
      startBytes,
      grewBy: lastSize - startBytes,
      depth,
      avgMutateMs: mutateCount ? mutateMsTotal / mutateCount : 0,
      ratio: costRatio(meter.percentile(50), mutateCount ? mutateMsTotal / mutateCount : 0),
    }),
  };
}
