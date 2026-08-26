// DOCX GOLF — course play on a live Word document.
//
// Sibling of ascii-arcade.js, but the inverse bet: the Arcade painted frames
// INTO one paragraph through `raw.replaceXml` (the escape hatch), proving the
// addressing model while touching almost none of the editing surface. Golf
// makes the surface itself the game. Every hole is a pair of documents: the
// one on your tee (loaded into the live ribbon editor) and the target. You
// play with the real clubs — typing into blocks, the Style dropdown, the
// table toolbar, Delete block, undo — and the referee is the product's own
// comparison engine: a hole is CLEARED when `docxDiffGetRevisions` between
// your document and the target returns zero revisions. Exact, unarguable,
// and the caddie panel phrases whatever revisions remain as the work left.
//
// Strokes are counted from the document, not from the toolbar: the driver
// polls `session.save()` and fingerprints the bytes, so one committed burst
// of editing — however it was made — is one swing. That also makes the
// driver honest about the one thing it must never do: mutate your ball.
// Scoring only ever reads.
//
// This file's home is docs/demo/ for the same reason ascii-arcade.js lives
// there: GitHub Pages deploys docs/ verbatim with no build step, and the npm
// Playwright webroot gets a pretest copy. It is demo content, not library
// machinery, and is deliberately NOT shipped in the npm package. The pure
// parts (course validation, stroke counter, score naming, caddie phrasing)
// are import-safe under Node for docs/demo/tools/docx-golf.test.mjs; the
// engine-dependent parts run only in the browser and are proven by
// npm/tests/demo-golf.spec.ts.

// ─── Golf arithmetic ──────────────────────────────────────────────────

/** Classic golf naming for a finished hole. */
export function scoreName(strokes, par) {
  if (strokes <= 1) return 'hole in one';
  const d = strokes - par;
  if (d <= -2) return 'eagle';
  if (d === -1) return 'birdie';
  if (d === 0) return 'par';
  if (d === 1) return 'bogey';
  if (d === 2) return 'double bogey';
  return `+${d}`;
}

/** Count swings as observed changes of an opaque state token (here: the
 *  fingerprint of the saved document). A burst of edits between two
 *  observations is ONE swing however many intermediate states it passed
 *  through, and any change — undo included — counts. */
export function createStrokeCounter(initial) {
  let last = initial;
  let strokes = 0;
  return {
    observe(token) {
      if (token !== last) {
        last = token;
        strokes++;
      }
      return strokes;
    },
    strokes: () => strokes,
  };
}

/** FNV-1a over the saved bytes — the stroke counter's state token. Cheap
 *  enough to run on a poll tick, and collision-safe enough for "did the
 *  document change since last look". */
export function bytesDigest(bytes) {
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) {
    h ^= bytes[i];
    h = Math.imul(h, 0x01000193);
  }
  return h >>> 0;
}

// ─── The caddie's phrasing ────────────────────────────────────────────

const quote = (text) => {
  const t = String(text ?? '').replace(/\s+/g, ' ').trim();
  return `"${t.length > 38 ? t.slice(0, 37) + '…' : t}"`;
};

/** Phrase a DocxDiff revision list (player → target) as the work remaining.
 *  Inserted = text you still need to add; Deleted = text you still need to
 *  remove; a Moved pair collapses to one hint. */
export function caddieLines(revisions) {
  const seenMoves = new Set();
  const lines = [];
  for (const rev of revisions) {
    const kind = String(rev.revisionType ?? '');
    if (kind === 'Moved') {
      const key = rev.moveGroupId ?? `~${rev.text}`;
      if (seenMoves.has(key)) continue;
      seenMoves.add(key);
      lines.push(`move ${quote(rev.text)}`);
    } else if (kind === 'Inserted') {
      lines.push(`add ${quote(rev.text)}`);
    } else if (kind === 'Deleted') {
      lines.push(`remove ${quote(rev.text)}`);
    } else if (kind === 'FormatChanged') {
      const props = rev.formatChange?.changedPropertyNames ?? [];
      lines.push(`reformat ${quote(rev.text)}${props.length ? ` (${props.join(', ')})` : ''}`);
    } else {
      lines.push(`${kind.toLowerCase() || 'change'} ${quote(rev.text)}`);
    }
  }
  return lines;
}

// ─── Course validation ────────────────────────────────────────────────

/** Static shape checks a level must pass before it is worth booting an
 *  engine for. The engine-level honesty checks (the start does not already
 *  score as solved; the reference solves within par) run per hole load in
 *  the driver and per hole in the Playwright spec. */
export function validateCourse(course) {
  const problems = [];
  const ids = new Set();
  if (!Array.isArray(course) || course.length === 0) return ['course must be a non-empty array'];
  course.forEach((hole, i) => {
    const at = `hole ${i + 1} (${hole?.id ?? '?'})`;
    if (!hole || typeof hole !== 'object') { problems.push(`${at}: not an object`); return; }
    if (typeof hole.id !== 'string' || !hole.id) problems.push(`${at}: missing id`);
    else if (ids.has(hole.id)) problems.push(`${at}: duplicate id`);
    else ids.add(hole.id);
    if (typeof hole.title !== 'string' || !hole.title) problems.push(`${at}: missing title`);
    if (!Number.isInteger(hole.par) || hole.par < 1) problems.push(`${at}: par must be a positive integer`);
    if (typeof hole.brief !== 'string' || !hole.brief.trim()) problems.push(`${at}: missing brief`);
    if (typeof hole.solve !== 'function') problems.push(`${at}: missing solve reference solution`);
    for (const side of ['start', 'target']) {
      const doc = hole[side];
      const ok = typeof doc === 'function' ||
        (Array.isArray(doc) && doc.length > 0 && doc.every((p) => typeof p === 'string'));
      if (!ok) problems.push(`${at}: ${side} must be a markdown paragraph list or a builder function`);
    }
    if (Array.isArray(hole.start) && Array.isArray(hole.target) &&
        JSON.stringify(hole.start) === JSON.stringify(hole.target)) {
      problems.push(`${at}: start already equals target — the hole would begin solved`);
    }
  });
  return problems;
}

// ─── The course ───────────────────────────────────────────────────────
// Five holes, each spotlighting a different stretch of the editing surface.
// Paragraphs are markdown (the editor's own projection: `#`/`##` are the
// Heading styles the ribbon's Style dropdown writes, so par is reachable
// through the UI). `solve` is the content-addressed reference solution —
// it never names an anchor id, because a player has to find things too.

const H2_NOTICES_FIXED =
  'Any notice to the Company must be delivered to the address set out on the signature page.';

const H5_FEE_INTRO = 'Rates below are fixed for the term of the engagement.';

function buildFeeSchedule(session, build, { discoveryRate, duplicateRow }) {
  build.paragraphs(['# Fee Schedule', H5_FEE_INTRO]);
  const intro = session.findAllByText(H5_FEE_INTRO)[0];
  const cells = [
    'Service', 'Hourly rate', 'Unit',
    'Discovery review', discoveryRate, 'per hour',
    'Contract drafting', '$300', 'per hour',
  ];
  if (duplicateRow) cells.push('Contract drafting', '$300', 'per hour');
  const rows = cells.length / 3;
  build.check(
    session.insertTable(intro.id, 'after', rows, 3, { cellContents: cells }),
    'fee schedule table',
  );
}

export const COURSE = [
  {
    id: 'first-tee',
    title: 'First tee',
    par: 1,
    surface: 'text',
    brief:
      'Warm up your grip: the preamble misspells "Purchaser". Click the paragraph, ' +
      'fix the word, then click anywhere outside it — edits count when they commit. ' +
      'One clean stroke.',
    start: [
      '# Master Services Agreement',
      'This Agreement is entered into by Acme Industries, Inc. and the Purchasr identified on the signature page.',
    ],
    target: [
      '# Master Services Agreement',
      'This Agreement is entered into by Acme Industries, Inc. and the Purchaser identified on the signature page.',
    ],
    solve(session) {
      const hit = session.findAllByText('Purchasr')[0];
      return [session.replaceText(hit.id, this.target[1])];
    },
  },
  {
    id: 'clause-order',
    title: 'Order in the house',
    par: 4,
    surface: 'structure',
    brief:
      'The closing clauses were pasted back in the wrong order, and one defined term ' +
      'was never conformed. Get Governing Law — heading AND body — ahead of ' +
      'Indemnification, and make the Notices clause say "the Company". Delete block ' +
      'and some retyping will get you there. Mind the Recitals: the definition of ' +
      '"the Company" must not change.',
    start: [
      '# Master Services Agreement',
      '## Recitals',
      'Acme Industries, Inc. ("the Company") and the Counterparty enter into this Agreement as of the date last signed below.',
      '## Indemnification',
      'The Company shall indemnify and hold harmless the Counterparty against any third-party claim arising out of the Services.',
      '## Governing Law',
      'This Agreement is governed by the laws of the State of Delaware, without regard to its conflict-of-laws principles.',
      '## Notices',
      'Any notice to Acme must be delivered to the address set out on the signature page.',
    ],
    target: [
      '# Master Services Agreement',
      '## Recitals',
      'Acme Industries, Inc. ("the Company") and the Counterparty enter into this Agreement as of the date last signed below.',
      '## Governing Law',
      'This Agreement is governed by the laws of the State of Delaware, without regard to its conflict-of-laws principles.',
      '## Indemnification',
      'The Company shall indemnify and hold harmless the Counterparty against any third-party claim arising out of the Services.',
      '## Notices',
      'Any notice to the Company must be delivered to the address set out on the signature page.',
    ],
    solve(session) {
      const heading = session.findAllByText('Governing Law')[0];
      const body = session.findAllByText('governed by the laws of the State of Delaware')[0];
      const indemnification = session.findAllByText('Indemnification')[0];
      const notices = session.findAllByText('Any notice to Acme')[0];
      return [
        session.moveBlock(heading.id, indemnification.id, 'before'),
        session.moveBlock(body.id, indemnification.id, 'before'),
        session.replaceText(notices.id, H2_NOTICES_FIXED),
      ];
    },
  },
  {
    id: 'defined-term',
    title: 'Conform the term',
    par: 3,
    surface: 'precision',
    brief:
      'The Recitals define "the Consultant" — then the drafter kept typing the full ' +
      'company name anyway. Conform the three later clauses to the defined term ' +
      'WITHOUT touching the definition itself. A document-wide find-and-replace ' +
      'would break the Recital; golf rewards the scoped shot.',
    start: [
      '# Consulting Agreement',
      'Blue Ridge Analytics LLC ("the Consultant") will provide the services described in the Statement of Work.',
      'Blue Ridge Analytics LLC shall invoice the Client monthly, with payment due within thirty days.',
      'All work product prepared by Blue Ridge Analytics LLC is assigned to the Client upon payment.',
      'Blue Ridge Analytics LLC warrants that the services will be performed in a professional and workmanlike manner.',
    ],
    target: [
      '# Consulting Agreement',
      'Blue Ridge Analytics LLC ("the Consultant") will provide the services described in the Statement of Work.',
      'The Consultant shall invoice the Client monthly, with payment due within thirty days.',
      'All work product prepared by the Consultant is assigned to the Client upon payment.',
      'The Consultant warrants that the services will be performed in a professional and workmanlike manner.',
    ],
    solve(session) {
      return [
        ['shall invoice the Client', this.target[2]],
        ['All work product', this.target[3]],
        ['warrants that the services', this.target[4]],
      ].map(([needle, text]) =>
        session.replaceText(session.findAllByText(needle)[0].id, text));
    },
  },
  {
    id: 'heading-day',
    title: 'Heading day',
    par: 3,
    surface: 'styles',
    brief:
      'This agreement arrived flat — every paragraph is body text. Style the title ' +
      'as Heading 1 and the two section titles as Heading 2 (the Style dropdown is ' +
      'your club). The referee cares about styles, not just words: the text is ' +
      'already identical.',
    start: [
      'Employment Agreement',
      'Duties',
      'The Executive shall serve as Chief Operating Officer and shall devote substantially all business time to the Company.',
      'Compensation',
      "Base salary accrues from the Start Date and is payable in accordance with the Company's normal payroll practices.",
    ],
    target: [
      '# Employment Agreement',
      '## Duties',
      'The Executive shall serve as Chief Operating Officer and shall devote substantially all business time to the Company.',
      '## Compensation',
      "Base salary accrues from the Start Date and is payable in accordance with the Company's normal payroll practices.",
    ],
    solve(session) {
      return [
        ['Employment Agreement', 'Heading1'],
        ['Duties', 'Heading2'],
        ['Compensation', 'Heading2'],
      ].map(([needle, style]) =>
        session.setParagraphStyle(session.findAllByText(needle)[0].id, style));
    },
  },
  {
    id: 'table-stakes',
    title: 'Table stakes',
    par: 2,
    surface: 'tables',
    brief:
      'The fee schedule picked up a duplicate row, and Discovery review is still ' +
      'billed at last year\'s rate. Make it $400 and delete the extra row — the ' +
      'table toolbar appears when your caret is inside the table.',
    start: (session, build) =>
      buildFeeSchedule(session, build, { discoveryRate: '$450', duplicateRow: true }),
    target: (session, build) =>
      buildFeeSchedule(session, build, { discoveryRate: '$400', duplicateRow: false }),
    solve(session) {
      const rate = session.findAllByText('$450')[0];
      const results = [session.replaceText(rate.id, '$400')];
      // The duplicate is the last row: cells are projected row-major, so with
      // 4 rows × 3 columns its first cell is index 9. Current engines address
      // row ops by the canonical `tc` anchor; engines from before table-
      // addressing was canonicalized took the cell's paragraph anchor instead,
      // so fall back for the pinned CDN engine.
      let removed = session.deleteTableRow(session.findByKind('tc', 'body')[9].id);
      if (!removed.success) {
        removed = session.deleteTableRow(session.findAllByText('Contract drafting')[1].id);
      }
      results.push(removed);
      return results;
    },
  },
];

// ─── Document building (driver-side, engine required) ─────────────────

function check(result, what) {
  if (!result?.success) {
    throw new Error(`${what} failed: ${result?.error?.code} ${result?.error?.message ?? ''}`);
  }
  return result;
}

/** Fill a one-paragraph body with a paragraph list. A `#`/`##`/`###` prefix
 *  declares a heading — applied as `setParagraphStyle`, the exact shape the
 *  ribbon's Style dropdown writes (`w:pStyle` only). The markdown block
 *  syntax would NOT be equivalent: `insertParagraph("## x")` attaches
 *  outline numbering the dropdown never writes, and a hole built that way
 *  could not be cleared with the clubs a player actually holds. */
function buildParagraphs(session, paragraphs) {
  const first = session.findByKind('p', 'body')[0];
  if (!first) throw new Error('expected a body paragraph to build on');
  const heading = (line) => {
    const m = /^(#{1,3}) (.*)$/.exec(line);
    return m ? { style: `Heading${m[1].length}`, text: m[2] } : { style: null, text: line };
  };
  const seed = heading(paragraphs[0]);
  let prev = check(session.replaceText(first.id, seed.text), 'seed paragraph')
    .modified?.[0]?.id ?? first.id;
  if (seed.style) {
    prev = check(session.setParagraphStyle(prev, seed.style), 'seed style')
      .modified?.[0]?.id ?? prev;
  }
  for (let i = 1; i < paragraphs.length; i++) {
    const { style, text } = heading(paragraphs[i]);
    let id = check(session.insertParagraph(prev, 'after', text), `paragraph ${i + 1}`)
      .created?.[0]?.id;
    if (style) {
      id = check(session.setParagraphStyle(id, style), `paragraph ${i + 1} style`)
        .modified?.[0]?.id ?? id;
    }
    prev = id ?? prev;
  }
}

/** Build a hole document (list or builder function) into a session whose body
 *  is a single blank-ish paragraph. */
function buildDoc(session, spec) {
  const build = {
    paragraphs: (list) => buildParagraphs(session, list),
    check,
  };
  if (typeof spec === 'function') spec(session, build);
  else buildParagraphs(session, spec);
}

/** Clear the live body back to one seed paragraph, whatever the last hole
 *  left in it. Sweep every block kind a hole can create — headings are kind
 *  'h', not 'p', and a missed kind here haunts every later hole as ghost
 *  content only the diff engine can see. */
function resetBody(session) {
  const KINDS = ['p', 'h', 'li', 'tbl'];
  const blocks = () => KINDS.flatMap((kind) => session.findByKind(kind, 'body'));
  const first = blocks()[0];
  if (!first) throw new Error('document has no body blocks');
  const seed = check(session.insertParagraph(first.id, 'before', '(teeing up…)'), 'tee seed');
  const seedId = seed.created[0].id;
  for (const block of blocks()) {
    if (block.id === seedId) continue;
    check(session.deleteBlock(block.id), `clearing ${block.id}`);
  }
}

// ─── The caddie panel (page chrome) ───────────────────────────────────

const PANEL_CSS = `
.dxg { display: flex; flex-direction: column; height: 100%; min-height: 0;
  font: 13px/1.5 system-ui, sans-serif; color: #1c2733; background: #f4f7f5;
  border-left: 1px solid #d7dfd9; }
.dxg * { box-sizing: border-box; }
.dxg-head { padding: 12px 14px 10px; border-bottom: 1px solid #d7dfd9; }
.dxg-brand { font: 700 15px/1 "SF Mono", Consolas, monospace; letter-spacing: .08em; color: #14532d; }
.dxg-holes { display: flex; gap: 5px; margin-top: 10px; flex-wrap: wrap; }
.dxg-holes button { font: 600 12px/1 "SF Mono", Consolas, monospace; padding: 6px 0;
  width: 34px; border: 1px solid #c3cfc6; border-radius: 7px; background: #fff;
  color: #33413a; cursor: pointer; }
.dxg-holes button[aria-pressed="true"] { background: #14532d; color: #fff; border-color: #14532d; }
.dxg-holes button[data-cleared="true"] { border-color: #16a34a; color: #16a34a; }
.dxg-holes button[data-cleared="true"][aria-pressed="true"] { color: #fff; }
.dxg-body { flex: 1; min-height: 0; overflow: auto; padding: 12px 14px; }
.dxg-title { font-weight: 700; font-size: 14px; }
.dxg-chips { display: flex; gap: 6px; margin: 6px 0 8px; }
.dxg-chip { font: 600 10.5px/1 "SF Mono", Consolas, monospace; letter-spacing: .06em;
  padding: 4px 8px; border-radius: 999px; background: #e2ebe4; color: #33513e; }
.dxg-brief { color: #45524b; margin: 0 0 12px; }
.dxg-score { display: flex; gap: 14px; padding: 9px 12px; border: 1px solid #d7dfd9;
  border-radius: 10px; background: #fff; margin-bottom: 10px; }
.dxg-score b { font: 700 17px/1.1 "SF Mono", Consolas, monospace; display: block; }
.dxg-score span { font-size: 10.5px; text-transform: uppercase; letter-spacing: .07em; color: #6b7a71; }
.dxg-score .dxg-diffs b { color: #b45309; }
.dxg-score[data-cleared="true"] .dxg-diffs b { color: #16a34a; }
.dxg-tabs { display: flex; gap: 4px; margin-bottom: 8px; }
.dxg-tabs button { flex: 1; font: 600 11.5px/1 system-ui, sans-serif; padding: 7px 0;
  border: 1px solid #c3cfc6; border-radius: 8px; background: #fff; color: #33413a; cursor: pointer; }
.dxg-tabs button[aria-pressed="true"] { background: #1c2733; border-color: #1c2733; color: #fff; }
.dxg-view { border: 1px solid #d7dfd9; border-radius: 10px; background: #fff;
  min-height: 180px; max-height: 44vh; overflow: auto; }
.dxg-hints { margin: 0; padding: 10px 12px 10px 28px; }
.dxg-hints li { margin: 4px 0; font: 12.5px/1.45 "SF Mono", Consolas, monospace; color: #7c2d12; }
.dxg-hints .dxg-done { list-style: none; margin-left: -16px; color: #16a34a; font-weight: 700; }
.dxg-view iframe { display: block; width: 100%; height: 42vh; border: 0; }
.dxg-note { color: #8a978e; font-size: 11.5px; margin: 8px 2px 0; }
.dxg-foot { padding: 10px 14px; border-top: 1px solid #d7dfd9; display: flex; gap: 8px; }
.dxg-foot button { font: 600 12.5px/1 system-ui, sans-serif; padding: 9px 12px;
  border: 1px solid #c3cfc6; border-radius: 9px; background: #fff; color: #33413a; cursor: pointer; }
.dxg-foot .dxg-next { flex: 1; background: #14532d; border-color: #14532d; color: #fff; display: none; }
.dxg-foot .dxg-next[data-on="true"] { display: block; }
.dxg-banner { display: none; margin: 0 0 10px; padding: 12px 14px; border-radius: 10px;
  background: #dcfce7; border: 1px solid #86efac; color: #14532d; font-weight: 600; }
.dxg-banner[data-on="true"] { display: block; }
.dxg-banner small { display: block; font-weight: 400; color: #3f6212; margin-top: 3px; }
.dxg-error { display: none; margin: 0 0 10px; padding: 10px 12px; border-radius: 10px;
  background: #fee2e2; border: 1px solid #fca5a5; color: #7f1d1d; font-size: 12px; }
.dxg-error[data-on="true"] { display: block; }
`;

/** Build the caddie panel inside `root` and return the element refs the
 *  driver wires. Pure DOM construction — no engine, no game state. */
export function mountGolfPanel(root) {
  const style = document.createElement('style');
  style.textContent = PANEL_CSS;
  document.head.appendChild(style);

  root.classList.add('dxg');
  root.innerHTML = `
    <div class="dxg-head">
      <div class="dxg-brand">⛳ DOCX GOLF</div>
      <div class="dxg-holes" data-dxg="holes"></div>
    </div>
    <div class="dxg-body">
      <div class="dxg-banner" data-dxg="banner"></div>
      <div class="dxg-error" data-dxg="error"></div>
      <div class="dxg-title" data-dxg="title"></div>
      <div class="dxg-chips">
        <span class="dxg-chip" data-dxg="parchip"></span>
        <span class="dxg-chip" data-dxg="surface"></span>
      </div>
      <p class="dxg-brief" data-dxg="brief"></p>
      <div class="dxg-score" data-dxg="score">
        <div><b data-dxg="strokes">0</b><span>strokes</span></div>
        <div><b data-dxg="par">–</b><span>par</span></div>
        <div class="dxg-diffs"><b data-dxg="diffs">–</b><span>diffs left</span></div>
      </div>
      <div class="dxg-tabs" data-dxg="tabs">
        <button data-view="caddie" aria-pressed="true">Caddie</button>
        <button data-view="target" aria-pressed="false">Target</button>
        <button data-view="redline" aria-pressed="false">Redline</button>
      </div>
      <div class="dxg-view" data-dxg="view"><ul class="dxg-hints" data-dxg="hints"></ul></div>
      <p class="dxg-note">Edits score when they commit — click outside a paragraph to bank a stroke.
        The redline is your document compared against the target.</p>
    </div>
    <div class="dxg-foot">
      <button data-dxg="reset">↻ Reset hole</button>
      <button class="dxg-next" data-dxg="next">Next hole ▶</button>
    </div>`;

  const grab = (name) => root.querySelector(`[data-dxg="${name}"]`);
  return {
    ui: {
      holes: grab('holes'), banner: grab('banner'), error: grab('error'),
      title: grab('title'), parchip: grab('parchip'), surface: grab('surface'),
      brief: grab('brief'), score: grab('score'), strokes: grab('strokes'),
      par: grab('par'), diffs: grab('diffs'), tabs: grab('tabs'),
      view: grab('view'), hints: grab('hints'),
      reset: grab('reset'), next: grab('next'),
    },
  };
}

// ─── The driver ───────────────────────────────────────────────────────

const POLL_MS = 700;
const SCORE_DEBOUNCE_MS = 350;

/**
 * Run the course against a ribbon-hosted editor.
 *
 * `engine` is the docxodus module namespace (the page passes its pinned
 * import), of which the driver uses: `docxDiffGetRevisions` (the referee),
 * `docxDiffCompare` + `convertDocxToHtml` (the redline and target views),
 * and `openDocxSession` + `createBlankDocx` (building each hole's target).
 * The driver mutates the live session ONLY while loading a hole; scoring
 * reads `session.save()` and never writes.
 *
 * Returns the controller the host page publishes as `window.__golf`.
 */
export function startGolf({ editor, session, engine, ui, course = COURSE }) {
  const problems = validateCourse(course);
  if (problems.length) throw new Error(`course invalid: ${problems.join('; ')}`);

  let holeIndex = -1;
  let counter = createStrokeCounter(0);
  let targetBytes = null;
  let targetHtml = null; // lazy per hole
  let cleared = false;
  let revisionsLeft = -1;
  let loading = false;
  let scoring = false;
  let scoreQueued = false;
  let scoreTimer = 0;
  let view = 'caddie';
  const scorecard = new Array(course.length).fill(null);

  // ── panel rendering ──────────────────────────────────────────────
  const holeButtons = [];
  course.forEach((hole, i) => {
    const b = document.createElement('button');
    b.textContent = String(i + 1);
    b.title = hole.title;
    b.addEventListener('click', () => { void loadHole(i); });
    holeButtons.push(b);
    ui.holes.appendChild(b);
  });

  function paintChrome() {
    const hole = course[holeIndex];
    holeButtons.forEach((b, i) => {
      b.setAttribute('aria-pressed', String(i === holeIndex));
      b.dataset.cleared = String(scorecard[i] != null);
    });
    ui.title.textContent = `Hole ${holeIndex + 1} — ${hole.title}`;
    ui.parchip.textContent = `par ${hole.par}`;
    ui.surface.textContent = hole.surface ?? 'document';
    ui.brief.textContent = hole.brief;
    ui.par.textContent = String(hole.par);
    ui.strokes.textContent = String(counter.strokes());
    ui.score.dataset.cleared = String(cleared);
    ui.next.dataset.on = String(cleared);
  }

  function paintScore(revisions) {
    revisionsLeft = revisions.length;
    ui.diffs.textContent = String(revisionsLeft);
    ui.strokes.textContent = String(counter.strokes());
    ui.hints.innerHTML = '';
    if (revisionsLeft === 0) {
      const li = document.createElement('li');
      li.className = 'dxg-done';
      li.textContent = '✔ your document matches the target';
      ui.hints.appendChild(li);
    } else {
      for (const line of caddieLines(revisions)) {
        const li = document.createElement('li');
        li.textContent = line;
        ui.hints.appendChild(li);
      }
    }
  }

  function banner(text, sub) {
    ui.banner.dataset.on = String(Boolean(text));
    ui.banner.innerHTML = '';
    if (!text) return;
    ui.banner.append(text);
    if (sub) {
      const s = document.createElement('small');
      s.textContent = sub;
      ui.banner.appendChild(s);
    }
  }

  function fail(err) {
    ui.error.dataset.on = 'true';
    ui.error.textContent = String(err?.message ?? err).slice(0, 300);
    throw err instanceof Error ? err : new Error(String(err));
  }

  const iframeView = (html) => {
    ui.view.innerHTML = '';
    const frame = document.createElement('iframe');
    frame.setAttribute('sandbox', ''); // static preview: no scripts, no navigation
    frame.srcdoc = html;
    ui.view.appendChild(frame);
  };

  async function renderView() {
    if (view === 'caddie') {
      ui.view.innerHTML = '';
      ui.view.appendChild(ui.hints);
      return;
    }
    if (view === 'target') {
      targetHtml ??= await engine.convertDocxToHtml(targetBytes);
      iframeView(targetHtml);
      return;
    }
    const redline = await engine.docxDiffCompare(session.save(), targetBytes);
    iframeView(await engine.convertDocxToHtml(redline));
  }

  for (const b of ui.tabs.querySelectorAll('button')) {
    b.addEventListener('click', () => {
      view = b.dataset.view;
      for (const other of ui.tabs.querySelectorAll('button')) {
        other.setAttribute('aria-pressed', String(other === b));
      }
      renderView().catch(fail);
    });
  }

  // ── scoring ──────────────────────────────────────────────────────
  async function scoreNow() {
    const bytes = session.save();
    const revisions = await engine.docxDiffGetRevisions(bytes, targetBytes);
    paintScore(revisions);
    if (view === 'redline') await renderView();
    if (revisions.length === 0 && !cleared) holeClear();
    return revisions;
  }

  async function runScore() {
    if (scoring) { scoreQueued = true; return; }
    scoring = true;
    try {
      await scoreNow();
    } catch (err) {
      fail(err);
    } finally {
      scoring = false;
      if (scoreQueued) { scoreQueued = false; void runScore(); }
    }
  }

  function scheduleScore() {
    clearTimeout(scoreTimer);
    scoreTimer = setTimeout(() => { void runScore(); }, SCORE_DEBOUNCE_MS);
  }

  const pollTimer = setInterval(() => {
    if (loading || cleared || holeIndex < 0) return;
    const before = counter.strokes();
    let digest;
    try {
      digest = bytesDigest(session.save());
    } catch {
      return; // a structural edit is mid-flight; the next tick will see it
    }
    if (counter.observe(digest) !== before) {
      ui.strokes.textContent = String(counter.strokes());
      scheduleScore();
    }
  }, POLL_MS);

  function holeClear() {
    cleared = true;
    const strokes = Math.max(1, counter.strokes());
    const hole = course[holeIndex];
    scorecard[holeIndex] = { strokes, par: hole.par, name: scoreName(strokes, hole.par) };
    const done = scorecard.filter(Boolean);
    const total = done.reduce((n, s) => n + s.strokes - s.par, 0);
    banner(
      `⛳ HOLE CLEAR — ${scoreName(strokes, hole.par)} (${strokes}/${hole.par})`,
      done.length === course.length
        ? `Round complete: ${total > 0 ? '+' + total : total === 0 ? 'even par' : total} for the course. New round: pick any hole.`
        : `Running total ${total > 0 ? '+' + total : total === 0 ? 'even' : total} · next hole when you are ready.`,
    );
    paintChrome();
  }

  // ── hole lifecycle ───────────────────────────────────────────────
  async function loadHole(i) {
    if (loading || i < 0 || i >= course.length) return;
    loading = true;
    ui.error.dataset.on = 'false';
    banner(null);
    try {
      const hole = course[i];
      holeIndex = i;
      cleared = false;
      scorecard[i] = null;
      targetHtml = null;

      // The tee: rebuild the live document to the hole's start…
      resetBody(session);
      buildDoc(session, hole.start);
      editor.refresh();

      // …and the pin: build the target in a throwaway session.
      const shadow = engine.openDocxSession(engine.createBlankDocx());
      try {
        buildDoc(shadow, hole.target);
        targetBytes = shadow.save();
      } finally {
        shadow.close();
      }

      counter = createStrokeCounter(bytesDigest(session.save()));
      paintChrome();

      // Honesty check, every load: a hole must not begin solved.
      const revisions = await scoreNow();
      if (revisions.length === 0) {
        fail(new Error(`course defect: hole ${i + 1} starts already solved`));
      }
      if (view !== 'caddie') await renderView();
    } catch (err) {
      fail(err);
    } finally {
      loading = false;
    }
  }

  ui.reset.addEventListener('click', () => { void loadHole(holeIndex); });
  ui.next.addEventListener('click', () => {
    void loadHole(Math.min(course.length - 1, holeIndex + 1));
  });

  const controller = {
    course,
    holeIndex: () => holeIndex,
    hole: () => course[holeIndex],
    strokes: () => counter.strokes(),
    revisionsLeft: () => revisionsLeft,
    cleared: () => cleared,
    scorecard: () => scorecard.slice(),
    loadHole,
    next: () => loadHole(holeIndex + 1),
    /** Score immediately (reads only). Returns the revision count. */
    check: async () => (await scoreNow()).length,
    /** Run the current hole's reference solution — the spec's honesty probe.
     *  Returns the ops it spent and whether they all succeeded. */
    playReference: async () => {
      const hole = course[holeIndex];
      const results = hole.solve(session);
      editor.refresh();
      const revisions = await scoreNow();
      return {
        ops: results.length,
        allOk: results.every((r) => r?.success),
        results,
        revisionsLeft: revisions.length,
      };
    },
    dispose: () => {
      clearInterval(pollTimer);
      clearTimeout(scoreTimer);
    },
  };

  void loadHole(0);
  return controller;
}
