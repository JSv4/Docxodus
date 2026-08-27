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

/** The course score so far: strokes-vs-par summed over the holes the player
 *  actually scored. A caddie-assisted hole is cleared but not played — it
 *  contributes nothing to the total (#584). */
export function runningTotal(scorecard) {
  let total = 0;
  let played = 0;
  let assisted = 0;
  for (const entry of scorecard) {
    if (!entry) continue;
    if (entry.assisted) { assisted++; continue; }
    played++;
    total += entry.strokes - entry.par;
  }
  return { total, played, assisted };
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

const H6_BODY = 'The Lender may rely on the opinions set out in the Master Agreement.';
const H6_NOTE = 'As defined in the Master Agreement dated January 5, 2026.';

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
      'Indemnification, and make the Notices clause say "the Company". The ⠿ handle ' +
      'beside each block moves it (Move up / Move down). Mind the Recitals: the ' +
      'definition of "the Company" must not change.',
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
  {
    id: 'footnote-drill',
    title: 'Footnote drill',
    par: 2,
    surface: 'notes',
    brief:
      'The reliance clause needs its citation. Put your caret at the very END of the ' +
      'body paragraph, use Insert → Footnote, and make the note read exactly: ' +
      '"As defined in the Master Agreement dated January 5, 2026." Then click back ' +
      'into the body — notes are part of the document, and the referee reads them too.',
    start: ['# Reliance Letter', H6_BODY],
    target: (session, build) => {
      build.paragraphs(['# Reliance Letter', H6_BODY]);
      const body = session.findAllByText('The Lender may rely')[0];
      build.check(session.insertFootnote(body.id, H6_BODY.length, H6_NOTE), 'target footnote');
    },
    solve(session) {
      const body = session.findAllByText('The Lender may rely')[0];
      return [session.insertFootnote(body.id, H6_BODY.length, H6_NOTE)];
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
  // Note definitions live in their own parts and outlive the body paragraphs
  // whose markers cited them — sweep them too, or the footnote hole haunts
  // every later hole the way ghost headings once did.
  for (const kind of ['fn', 'en']) {
    for (const note of session.findByKind(kind)) {
      check(session.deleteBlock(note.id), `clearing ${kind} ${note.id}`);
    }
  }
}

// ─── The caddie panel (page chrome) ───────────────────────────────────

const PANEL_CSS = `
.dxg { display: flex; flex-direction: column; height: 100%; min-height: 0;
  font: 13px/1.5 system-ui, sans-serif; color: #1c2733; background: #f4f7f5;
  border-left: 1px solid #d7dfd9; }
.dxg * { box-sizing: border-box; }
.dxg-head { flex: none; padding: 12px 14px 10px; border-bottom: 1px solid #d7dfd9; }
.dxg-headrow { display: flex; align-items: center; gap: 10px; min-width: 0; }
.dxg-brand { font: 700 15px/1 "SF Mono", Consolas, monospace; letter-spacing: .08em;
  color: #14532d; white-space: nowrap; }
.dxg-mini { display: none; font: 600 11.5px/1 "SF Mono", Consolas, monospace; color: #45524b;
  margin-left: auto; white-space: nowrap; min-width: 0; overflow: hidden; text-overflow: ellipsis; }
.dxg-toggle { display: none; font: 600 11.5px/1 system-ui, sans-serif; padding: 7px 12px;
  border: 1px solid #14532d; border-radius: 8px; background: #14532d; color: #fff;
  cursor: pointer; white-space: nowrap; }
.dxg-toggle[aria-expanded="true"] { background: #fff; color: #14532d; }
/* The sheet is an invisible wrapper on a desk (holes, body, foot flow as one
   column) and becomes the phone bottom sheet in compact mode below. */
.dxg-sheet { display: flex; flex-direction: column; flex: 1; min-height: 0; }
.dxg-minibrief, .dxg-grab, .dxg-scrim { display: none; }
.dxg-holesrow { display: flex; align-items: baseline; gap: 8px; flex: none;
  padding: 10px 14px; border-bottom: 1px solid #d7dfd9; }

/* ── Compact (phone) mode ─────────────────────────────────────────────────────
   The caddie docks as a fixed scorecard bar along the bottom edge — brand,
   live mini score, a one-line hole brief, and the Caddie button — so the
   document keeps the whole screen and the objective stays readable without
   opening anything. The toggle raises the full panel as a bottom sheet OVER
   the document (grab handle, rounded top, scrim behind), the pattern a thumb
   already knows, instead of squeezing the editor into a sliver. Driver sets
   data-compact from a media query and data-open from the toggle/scrim/grab;
   it also opens the sheet itself when a hole clears, so the banner and the
   unlocked Next hole are never celebrated off-screen. */
.dxg[data-compact="true"] { position: fixed; left: 0; right: 0; bottom: 0; top: auto;
  z-index: 60; height: auto;
  max-height: calc(100vh - 16px); max-height: calc(100dvh - 16px);
  background: transparent; border-left: 0; }
.dxg[data-compact="true"] .dxg-mini { display: inline; }
.dxg[data-compact="true"] .dxg-toggle { display: inline-block; }
.dxg[data-compact="true"] .dxg-brand { font-size: 13px; }
.dxg[data-compact="true"] .dxg-head { order: 2; position: relative; z-index: 1;
  padding: 9px 12px calc(9px + env(safe-area-inset-bottom));
  background: rgba(244, 247, 245, .96);
  -webkit-backdrop-filter: blur(8px); backdrop-filter: blur(8px);
  border-top: 1px solid #d7dfd9; border-bottom: 0;
  box-shadow: 0 -6px 18px rgba(16, 42, 26, .08); }
.dxg[data-compact="true"]:not([data-open="true"]) .dxg-minibrief { display: block;
  width: 100%; margin: 7px 0 0; padding: 0; border: 0; background: none;
  cursor: pointer; text-align: left; font: 12px/1.4 system-ui, sans-serif;
  color: #45524b; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.dxg-minibrief b { color: #14532d; font-weight: 700; }
.dxg[data-compact="true"] .dxg-sheet { display: none; order: 1; position: relative;
  z-index: 1; min-height: 0; margin-bottom: -1px; max-height: 76vh; max-height: 76dvh;
  background: #f4f7f5; border: 1px solid #d7dfd9; border-bottom: 0;
  border-radius: 16px 16px 0 0; overflow: hidden;
  box-shadow: 0 -18px 44px rgba(16, 42, 26, .2); }
.dxg[data-compact="true"][data-open="true"] .dxg-sheet { display: flex;
  animation: dxg-rise .28s cubic-bezier(.32, .72, .28, 1); }
.dxg[data-compact="true"][data-open="true"] .dxg-scrim { display: block;
  position: fixed; inset: 0; z-index: 0; background: rgba(16, 42, 26, .35);
  -webkit-tap-highlight-color: transparent; }
.dxg[data-compact="true"] .dxg-grab { display: grid; place-items: center; flex: none;
  width: 100%; height: 24px; margin: 4px 0 0; padding: 0; border: 0;
  background: none; cursor: pointer; }
.dxg-grab i { display: block; width: 44px; height: 5px; border-radius: 999px;
  background: #c3cfc6; }
.dxg[data-compact="true"] .dxg-holesrow { padding: 4px 14px 10px; }
.dxg[data-compact="true"] .dxg-view { max-height: 38vh; max-height: 38dvh; }
.dxg[data-compact="true"] .dxg-view iframe { height: 36vh; height: 36dvh; }
@keyframes dxg-rise { from { transform: translateY(26px); opacity: .4; }
  to { transform: translateY(0); opacity: 1; } }
@media (prefers-reduced-motion: reduce) {
  .dxg[data-compact="true"][data-open="true"] .dxg-sheet { animation: none; } }
.dxg-holeslabel { font: 600 10.5px/1 "SF Mono", Consolas, monospace; letter-spacing: .07em;
  text-transform: uppercase; color: #6b7a71; }
.dxg-holes { display: flex; gap: 5px; flex-wrap: wrap; }
.dxg-holes button { font: 600 12px/1 "SF Mono", Consolas, monospace; padding: 6px 0;
  width: 34px; border: 1px solid #c3cfc6; border-radius: 7px; background: #fff;
  color: #33413a; cursor: pointer; }
.dxg-holes button[aria-pressed="true"] { background: #14532d; color: #fff; border-color: #14532d; }
.dxg-holes button[data-cleared="true"] { border-color: #16a34a; color: #16a34a; }
.dxg-holes button[data-cleared="true"][aria-pressed="true"] { color: #fff; }
.dxg-body { flex: 1; min-height: 0; overflow: auto; padding: 12px 14px; }
.dxg-howto { border: 1px solid #cfe3d4; border-radius: 10px; background: #eef6f0;
  margin: 0 0 12px; padding: 8px 12px; }
.dxg-howto summary { cursor: pointer; font-weight: 700; font-size: 12.5px; color: #14532d; }
.dxg-howto ol { margin: 8px 0 4px; padding-left: 20px; }
.dxg-howto li { margin: 4px 0; color: #33413a; }
.dxg-howto p { margin: 8px 0 2px; color: #45524b; }
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
.dxg-busy { margin: 0; padding: 14px 12px; color: #6b7a71; font-size: 12px; }
.dxg-note { color: #8a978e; font-size: 11.5px; margin: 8px 2px 0; }
.dxg-foot { padding: 10px 14px; border-top: 1px solid #d7dfd9; display: flex; gap: 8px; }
.dxg-foot button { font: 600 12.5px/1 system-ui, sans-serif; padding: 9px 12px;
  border: 1px solid #c3cfc6; border-radius: 9px; background: #fff; color: #33413a; cursor: pointer; }
/* Next is always visible so the player knows advancing exists — it just stays
   locked until the referee reads zero diffs. */
.dxg-foot .dxg-next { flex: 1; background: #14532d; border-color: #14532d; color: #fff; }
.dxg-foot .dxg-next[disabled] { background: #eef2ef; border-color: #d7dfd9; color: #9aa79f;
  cursor: default; }
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
    <div class="dxg-scrim" data-dxg="scrim" aria-hidden="true"></div>
    <div class="dxg-head">
      <div class="dxg-headrow">
        <div class="dxg-brand">⛳ DOCX GOLF</div>
        <span class="dxg-mini" data-dxg="mini" title="Strokes / par · diffs left"></span>
        <button class="dxg-toggle" data-dxg="toggle" aria-expanded="false"
          title="Show the caddie">☰ Caddie</button>
      </div>
      <button class="dxg-minibrief" data-dxg="minibrief"
        title="Open the caddie"></button>
    </div>
    <div class="dxg-sheet" data-dxg="sheet">
    <button class="dxg-grab" data-dxg="grab" aria-label="Close the caddie"
      title="Close the caddie"><i></i></button>
    <div class="dxg-holesrow">
      <span class="dxg-holeslabel">Holes</span>
      <div class="dxg-holes" data-dxg="holes" role="group" aria-label="Pick a hole"></div>
    </div>
    <div class="dxg-body">
      <div class="dxg-banner" data-dxg="banner"></div>
      <div class="dxg-error" data-dxg="error"></div>
      <details class="dxg-howto" data-dxg="howto" open>
        <summary>How to play</summary>
        <ol>
          <li><b>Edit the document</b> on the left — it is a real Word file, and it is your ball.</li>
          <li><b>Click outside the paragraph</b> to commit the edit. Each committed burst of
            edits is one stroke, undo included.</li>
          <li>Match the target document. When <b>diffs left</b> reaches <b>0</b> the hole is
            cleared and <b>Next hole</b> unlocks.</li>
        </ol>
        <p>Stuck? <b>Target</b> shows the document you are aiming for, <b>Redline</b> shows
          your differences as tracked changes, and <b>Show me</b> concedes the hole to the
          caddie (cleared, but not scored).</p>
      </details>
      <div class="dxg-title" data-dxg="title"></div>
      <div class="dxg-chips">
        <span class="dxg-chip" data-dxg="parchip" title="The reference solution clears the hole in this many strokes"></span>
        <span class="dxg-chip" data-dxg="surface" title="The part of the editing surface this hole plays"></span>
      </div>
      <p class="dxg-brief" data-dxg="brief"></p>
      <div class="dxg-score" data-dxg="score">
        <div title="Committed edits so far — undo counts too"><b data-dxg="strokes">0</b><span>strokes</span></div>
        <div title="The reference solution's stroke count"><b data-dxg="par">–</b><span>par</span></div>
        <div class="dxg-diffs" title="Differences between your document and the target — zero clears the hole"><b data-dxg="diffs">–</b><span>diffs left</span></div>
      </div>
      <div class="dxg-tabs" data-dxg="tabs">
        <button data-view="caddie" aria-pressed="true"
          title="The caddie's list of the work remaining">Caddie</button>
        <button data-view="target" aria-pressed="false"
          title="The document you are trying to match">Target</button>
        <button data-view="redline" aria-pressed="false"
          title="Your document compared against the target, as tracked changes">Redline</button>
      </div>
      <div class="dxg-view" data-dxg="view"><ul class="dxg-hints" data-dxg="hints"></ul></div>
      <p class="dxg-note">Edits score when they commit — click outside a paragraph to bank a stroke.</p>
    </div>
    <div class="dxg-foot">
      <button data-dxg="reset" title="Re-tee this hole: the document and your strokes go back to the start">↻ Reset</button>
      <button data-dxg="showme" title="Concede: the caddie makes the edits for you — the hole clears but is not scored">🏳 Show me</button>
      <button class="dxg-next" data-dxg="next" disabled title="Clear the hole to unlock">Next hole ▶</button>
    </div>
    </div>`;

  const grab = (name) => root.querySelector(`[data-dxg="${name}"]`);
  return {
    ui: {
      panel: root,
      holes: grab('holes'), banner: grab('banner'), error: grab('error'),
      title: grab('title'), parchip: grab('parchip'), surface: grab('surface'),
      brief: grab('brief'), score: grab('score'), strokes: grab('strokes'),
      par: grab('par'), diffs: grab('diffs'), tabs: grab('tabs'),
      view: grab('view'), hints: grab('hints'),
      reset: grab('reset'), next: grab('next'),
      showme: grab('showme'), toggle: grab('toggle'), mini: grab('mini'),
      howto: grab('howto'), sheet: grab('sheet'), scrim: grab('scrim'),
      grabber: grab('grab'), minibrief: grab('minibrief'),
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
  let assisted = false;
  let revisionsLeft = -1;
  let loading = false;
  let pendingHole = null;
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
    b.title = `Hole ${i + 1} — ${hole.title} (par ${hole.par})`;
    b.setAttribute('aria-label', b.title);
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
    // The bar's one-line brief: on a phone the objective must be readable
    // without opening the sheet, or the course reads as a bare document.
    ui.minibrief.replaceChildren();
    const miniTitle = document.createElement('b');
    miniTitle.textContent = `Hole ${holeIndex + 1} · ${hole.title}`;
    ui.minibrief.append(miniTitle, ` — ${hole.brief}`);
    ui.par.textContent = String(hole.par);
    ui.strokes.textContent = String(counter.strokes());
    ui.score.dataset.cleared = String(cleared);
    // Next stays on screen so the player knows advancing exists; it unlocks
    // when the hole clears. On the last hole the clear banner offers the new
    // round instead, so the button stays locked.
    const lastHole = holeIndex === course.length - 1;
    ui.next.dataset.on = String(cleared);
    ui.next.disabled = !cleared || lastHole;
    ui.next.title = !cleared
      ? 'Clear the hole to unlock'
      : lastHole
        ? 'End of the course — pick any hole above to play again'
        : 'Tee up the next hole';
  }

  function paintScore(revisions) {
    revisionsLeft = revisions.length;
    ui.diffs.textContent = String(revisionsLeft);
    ui.strokes.textContent = String(counter.strokes());
    // Scorecard notation, same as the clear banner: strokes/par, then work left.
    ui.mini.textContent =
      `${counter.strokes()}/${course[holeIndex]?.par ?? '–'} · ${revisionsLeft} left`;
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

  // The redline DOCX is a tracked-changes document, and `convertDocxToHtml`
  // ACCEPTS revisions by default (the converter runs RevisionAccepter) — so
  // rendering it with default options shows the clean target and no redline
  // at all. Ask for the markup explicitly.
  const REDLINE_HTML_OPTIONS = {
    renderTrackedChanges: true,
    showDeletedContent: true,
    renderMoveOperations: true,
  };

  const busyView = (label) => {
    ui.view.innerHTML = '';
    const p = document.createElement('p');
    p.className = 'dxg-busy';
    p.textContent = label;
    ui.view.appendChild(p);
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
    iframeView(await engine.convertDocxToHtml(redline, REDLINE_HTML_OPTIONS));
  }

  for (const b of ui.tabs.querySelectorAll('button')) {
    b.addEventListener('click', () => {
      view = b.dataset.view;
      for (const other of ui.tabs.querySelectorAll('button')) {
        other.setAttribute('aria-pressed', String(other === b));
      }
      // Target and redline renders go through the engine and can take a
      // beat — say so instead of leaving the last view frozen in place.
      if (view === 'target' && !targetHtml) busyView('Rendering the target…');
      else if (view === 'redline') busyView('Comparing your document against the target…');
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

  let howtoFolded = false;

  function holeClear() {
    cleared = true;
    // The player has proven they know how to play — fold the guide away
    // (once; if they re-open it later it stays open).
    if (!howtoFolded) { howtoFolded = true; ui.howto.open = false; }
    const strokes = Math.max(1, counter.strokes());
    const hole = course[holeIndex];
    const name = assisted ? 'caddie-assisted' : scoreName(strokes, hole.par);
    scorecard[holeIndex] = { strokes, par: hole.par, name, assisted };
    const done = scorecard.filter(Boolean);
    const score = runningTotal(scorecard);
    const totalText = score.total > 0 ? '+' + score.total : score.total === 0 ? 'even' : String(score.total);
    const assistedNote = score.assisted > 0
      ? ` (${score.assisted} hole${score.assisted === 1 ? '' : 's'} shown by the caddie)`
      : '';
    banner(
      assisted
        ? `🏳 The caddie played the line — hole cleared, no score`
        : `⛳ HOLE CLEAR — ${name} (${strokes}/${hole.par})`,
      done.length === course.length
        ? `Round complete: ${totalText} over ${score.played} scored hole${score.played === 1 ? '' : 's'}${assistedNote}. New round: pick any hole.`
        : `Running total ${totalText}${assistedNote} · next hole when you are ready.`,
    );
    paintChrome();
    // On a phone the banner and the unlocked Next hole live in the sheet —
    // raise it so the clear is never celebrated off-screen.
    if (isCompact()) setOpen(true);
  }

  // ── hole lifecycle ───────────────────────────────────────────────
  async function loadHole(i) {
    if (i < 0 || i >= course.length) return;
    // Last click wins: a request arriving while a load is in flight is queued
    // (replacing any earlier queued request) and honored when the load settles,
    // instead of being silently dropped (#588).
    if (loading) { pendingHole = i; return; }
    loading = true;
    ui.error.dataset.on = 'false';
    banner(null);
    try {
      const hole = course[i];
      holeIndex = i;
      cleared = false;
      assisted = false;
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
      if (pendingHole !== null && pendingHole !== holeIndex) {
        const next = pendingHole;
        pendingHole = null;
        void loadHole(next);
      } else {
        pendingHole = null;
      }
    }
  }

  ui.reset.addEventListener('click', () => { void loadHole(holeIndex); });
  ui.next.addEventListener('click', () => {
    void loadHole(Math.min(course.length - 1, holeIndex + 1));
  });
  // Concede: the caddie plays the reference line on the live document. The
  // hole clears, but the scorecard marks it assisted rather than scoring it.
  ui.showme.addEventListener('click', () => {
    if (loading || cleared || holeIndex < 0) return;
    try {
      assisted = true;
      course[holeIndex].solve(session);
      editor.refresh();
      void runScore();
    } catch (err) {
      fail(err);
    }
  });

  // Compact (phone) mode: the caddie docks as a bottom scorecard bar and the
  // toggle raises it as a bottom sheet over the document. The scrim, the grab
  // handle, Escape, and the toggle itself all lower it again.
  const compactQuery = typeof window.matchMedia === 'function'
    ? window.matchMedia('(max-width: 640px)')
    : null;
  const isCompact = () => ui.panel.dataset.compact === 'true';
  const setOpen = (open) => {
    ui.panel.dataset.open = String(open);
    ui.toggle.setAttribute('aria-expanded', String(open));
    ui.toggle.textContent = open ? '✕ Close' : '☰ Caddie';
    ui.toggle.title = open ? 'Close the caddie' : 'Show the caddie';
  };
  const applyCompact = () => {
    const compact = Boolean(compactQuery?.matches);
    ui.panel.dataset.compact = String(compact);
    setOpen(false); // a density flip never strands an open sheet
  };
  compactQuery?.addEventListener('change', applyCompact);
  applyCompact();
  ui.toggle.addEventListener('click', () => setOpen(ui.panel.dataset.open !== 'true'));
  ui.minibrief.addEventListener('click', () => setOpen(true));
  ui.scrim.addEventListener('click', () => setOpen(false));
  ui.grabber.addEventListener('click', () => setOpen(false));
  const onKeydown = (event) => {
    if (event.key === 'Escape' && isCompact() && ui.panel.dataset.open === 'true') setOpen(false);
  };
  document.addEventListener('keydown', onKeydown);

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
      document.removeEventListener('keydown', onKeydown);
    },
  };

  void loadHole(0);
  return controller;
}
