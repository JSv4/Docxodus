// REDLINE THEATER — three counsel negotiate a contract through MCP tool calls,
// live, at a speed you can watch.
//
// The third demo in the series, and the third bet. The Arcade painted frames INTO
// one paragraph through `raw.replaceXml`, proving the addressing model while
// touching almost none of the editing surface. Golf made the editing surface the
// game and the comparison engine the referee. Theater makes the AGENT PROTOCOL the
// show: every edit you watch land is dispatched from a real JSON-RPC 2.0
// `tools/call` frame, in the exact shape `docxodus-mcp` accepts over stdio, and
// every one is recorded as native Word tracked-change markup as it lands.
//
// Nothing here is a rendering of a redline. The document IS the redline while you
// watch it being written: the session runs in `render_inline` recording mode, so
// each `docxodus_edit` writes `w:ins`/`w:del` into the live package, the editor
// repaints just the block that changed, and what you are looking at is the file
// that downloads. The three counsel are three values of the session's revision
// author, so the result is genuinely author-attributed — the same markup Word's
// Reviewing pane groups by reviewer.
//
// The proof at the end is the point of the whole thing. `proveRedlineReversibility`
// rebuilds two packages from the redline — accept-all must reproduce the intended
// final, reject-all must reproduce the baseline the show started from — and reports
// the divergences. A redline you cannot reject back to the baseline is a rewrite
// wearing tracked changes; this one is checked, on screen, every run.
//
// ON SPEED. The pacing is a display choice, not an engine limit, and the HUD says
// so with measured numbers: mutations are single-digit milliseconds because they
// are in-memory OOXML edits, and the repaint is frame-dropped (ops keep flowing
// while a repaint is in flight; the next animation frame renders whatever state
// the document has reached). MAX removes the pacing delay entirely and the
// throughput readout becomes the engine's actual rate. That is the honest answer
// to "is it fast enough to look like animation" — you can watch it and read the
// number at the same time.
//
// This file's home is docs/demo/ for the same reason its two siblings live there:
// GitHub Pages deploys docs/ verbatim with no build step, and the npm Playwright
// webroot gets a pretest copy. It is demo content, not library machinery, and is
// deliberately NOT shipped in the npm package. The pure parts (script validation,
// binding substitution, telemetry, attribution roll-up) are import-safe under Node
// for docs/demo/tools/redline-theater.test.mjs; the engine-dependent parts run only
// in the browser and are proven by npm/tests/demo-redline.spec.ts.

import {
  createBrowserMcpEndpoint,
  createLatencyStats,
  frameLabel,
  summarizeArgs,
  IMPLEMENTED_TOOLS,
} from './mcp-wire.js';

import { mountStressPanel, createStressRunner, clauseFor } from './diff-stress.js';

// ─── Binding substitution (pure) ──────────────────────────────────────
//
// A real agent does not know anchor ids up front: it searches, reads an id out of
// the result, and puts that id in the next call. The script models exactly that.
// A step may `bind` a name to the anchor its result yields, and any later argument
// written `@name` is substituted with that anchor id BEFORE the frame is minted —
// so the frame on the wire carries the real id, the way the agent's would.

/** Pull the reusable anchor id out of a search result: text/regex matches carry it
 *  at `enclosingAnchor.id`, anchor-only results at `id`. Mutation results bind to
 *  what they touched — `modified` first, then `created`. */
export function bindFromResult(result) {
  if (!result || typeof result !== 'object') return null;
  if (Array.isArray(result.matches) && result.matches.length) {
    return result.matches[0].enclosingAnchor?.id ?? null;
  }
  if (Array.isArray(result.anchors) && result.anchors.length) {
    return result.anchors[0].id ?? null;
  }
  if (Array.isArray(result.modified) && result.modified.length) return result.modified[0].id;
  if (Array.isArray(result.created) && result.created.length) return result.created[0].id;
  return null;
}

/** The character offset just past a text match, for ops that address a point inside
 *  a paragraph rather than the paragraph itself — `insert_footnote`'s
 *  `characterOffset` is the only one the script uses. An agent gets this the same
 *  way: it is arithmetic on the span the search already returned, not a second
 *  lookup. Returns null when the result carries no match to measure. */
export function offsetAfterMatch(result) {
  const span = result?.matches?.[0]?.span;
  if (!span || typeof span.start !== 'number' || typeof span.length !== 'number') return null;
  return span.start + span.length;
}

/** Substitute `@name` references from `bindings`, deeply. An unbound reference is
 *  an error rather than a silent literal `"@name"` — a mis-scripted step must fail
 *  loudly, not quietly edit the wrong block. */
export function resolveArgs(args, bindings) {
  if (typeof args === 'string') {
    if (args.startsWith('@')) {
      const key = args.slice(1);
      if (!(key in bindings) || bindings[key] == null) {
        throw new Error(`unbound script reference: @${key}`);
      }
      return bindings[key];
    }
    return args;
  }
  if (Array.isArray(args)) return args.map((v) => resolveArgs(v, bindings));
  if (args && typeof args === 'object') {
    const out = {};
    for (const [k, v] of Object.entries(args)) out[k] = resolveArgs(v, bindings);
    return out;
  }
  return args;
}

/** Collect every `@name` reference in a step's arguments. */
export function referencesIn(args, found = []) {
  if (typeof args === 'string') {
    if (args.startsWith('@')) found.push(args.slice(1));
  } else if (Array.isArray(args)) {
    for (const v of args) referencesIn(v, found);
  } else if (args && typeof args === 'object') {
    for (const v of Object.values(args)) referencesIn(v, found);
  }
  return found;
}

// ─── Script validation (pure) ─────────────────────────────────────────

/** Static checks the script must pass before it is worth booting an engine for:
 *  every step names a tool and action this endpoint actually routes, and every
 *  `@reference` is bound by an EARLIER step (so the run cannot deadlock on an id
 *  that is never produced). The contract test additionally proves each pair
 *  against the real ToolCatalog.cs. */
export function validateScript(script, implemented = IMPLEMENTED_TOOLS) {
  const problems = [];
  if (!Array.isArray(script) || script.length === 0) return ['script must be a non-empty array'];
  const bound = new Set();
  script.forEach((act, ai) => {
    const at = `act ${ai + 1} (${act?.id ?? '?'})`;
    if (!act || typeof act !== 'object') { problems.push(`${at}: not an object`); return; }
    if (!act.id) problems.push(`${at}: missing id`);
    if (!act.title) problems.push(`${at}: missing title`);
    if (!act.counsel?.name) problems.push(`${at}: missing counsel`);
    if (!Array.isArray(act.steps) || act.steps.length === 0) {
      problems.push(`${at}: missing steps`);
      return;
    }
    act.steps.forEach((step, si) => {
      const where = `${at} step ${si + 1}`;
      const spec = implemented[step?.tool];
      if (!spec) { problems.push(`${where}: unknown tool ${step?.tool}`); return; }
      const action = step.args?.action;
      if (spec.actions !== null) {
        if (!action) problems.push(`${where}: ${step.tool} requires an action`);
        else if (!spec.actions.includes(action)) {
          problems.push(`${where}: ${step.tool} has no action ${action}`);
        }
      }
      if (spec.modes && step.args?.mode && !spec.modes.includes(step.args.mode)) {
        problems.push(`${where}: ${step.tool} has no mode ${step.args.mode}`);
      }
      for (const ref of referencesIn(step.args ?? {})) {
        if (!bound.has(ref)) problems.push(`${where}: @${ref} is used before it is bound`);
      }
      if (step.bind) bound.add(step.bind);
      if (step.bindOffset) bound.add(step.bindOffset);
    });
  });
  return problems;
}

// ─── The baseline document ────────────────────────────────────────────
//
// Built clean, with recording OFF, before the show begins — it is the thing every
// mark is measured against, and the thing reject-all has to reproduce. Two
// deliberate defects are planted for Act III to find: the agreement defines
// "Supplier" and then says "Vendor" twice.

const DEFINED_TERM_DEFECT = 'The Vendor shall not subcontract any portion of the Services ' +
  'without the prior written consent of Customer.';

export const BASELINE = [
  '# Master Services Agreement',
  'This Master Services Agreement (the "Agreement") is entered into as of March 3, 2026 by and ' +
    'between Northwind Logistics, Inc. ("Customer") and Arden Systems LLC ("Supplier").',
  '## Section 1 — Definitions',
  '"Confidential Information" means any non-public information disclosed by either party to the ' +
    'other, whether or not marked as confidential.',
  '"Services" means the professional services described in each Statement of Work.',
  '"Deliverable" means any work product Supplier furnishes to Customer under a Statement of Work.',
  '## Section 2 — Services and Statements of Work',
  'Supplier shall perform the Services described in each Statement of Work executed by the parties.',
  'Each Statement of Work shall identify the Deliverables, the fee basis, and the acceptance criteria.',
  DEFINED_TERM_DEFECT,
  '## Section 3 — Fees and Payment',
  'Customer shall pay the fees set out in the applicable Statement of Work within sixty (60) days ' +
    'of the date of invoice.',
  'The Vendor shall issue invoices monthly in arrears.',
  '## Section 4 — Confidentiality',
  "Each party shall protect the other party's Confidential Information using the same degree of " +
    'care it uses to protect its own, and in no event less than a reasonable degree of care.',
  '## Section 5 — Limitation of Liability',
  "Supplier's total liability under this Agreement shall not exceed the fees paid in the twelve " +
    '(12) months preceding the claim.',
  '## Section 6 — Term and Termination',
  'This Agreement commences on the Effective Date and continues for three (3) years unless ' +
    'terminated earlier in accordance with this Section 6.',
  'Either party may terminate this Agreement for material breach on thirty (30) days written notice.',
  '## Section 7 — Notices',
  'Any notice under this Agreement must be delivered to the address set out on the signature page.',
  '## Section 8 — Governing Law',
  'This Agreement is governed by the laws of the State of Delaware, without regard to its ' +
    'conflict of laws principles.',
];

/** The fee table, inserted after the payment paragraph. Built with the same
 *  `insertTable` call the ribbon's table button makes. */
const FEE_TABLE = {
  rows: 4,
  columns: 3,
  cellContents: [
    'Service', 'Rate', 'Unit',
    'Implementation', '$280', 'per hour',
    'Managed support', '$9,500', 'per month',
    'Out-of-scope change', '$340', 'per hour',
  ],
};

// ─── The counsel ──────────────────────────────────────────────────────

export const COUNSEL = {
  customer: {
    id: 'customer',
    name: 'Dana Whitfield',
    role: "Customer's counsel",
    initials: 'DW',
    color: '#1d4ed8',
  },
  supplier: {
    id: 'supplier',
    name: 'Marcus Oyelaran',
    role: "Supplier's counsel",
    initials: 'MO',
    color: '#b45309',
  },
  compliance: {
    id: 'compliance',
    name: 'Priya Raghunathan',
    role: 'Compliance review',
    initials: 'PR',
    color: '#7e22ce',
  },
};

// ─── The negotiation ──────────────────────────────────────────────────
//
// Three acts. Each step is one MCP tool call; `bind` names the anchor its result
// yields, `@name` spends it. `note` is the caption the stage shows while the call
// is in flight — the human sentence behind the frame.

export const SCRIPT = [
  {
    id: 'act-1',
    title: 'Act I — Customer tightens the terms',
    counsel: COUNSEL.customer,
    synopsis:
      'Payment terms pulled in, the liability cap doubled, and carve-outs added under it. ' +
      'Every mark is attributed to Dana Whitfield.',
    steps: [
      {
        note: 'Find the payment-terms sentence',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'within sixty (60) days' },
        bind: 'payment',
      },
      {
        note: 'Net 60 → Net 30',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@payment',
          find: 'sixty (60) days', replace: 'thirty (30) days',
        },
      },
      {
        note: 'Find the liability cap',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'total liability under this Agreement' },
        bind: 'cap',
      },
      {
        note: 'Raise the cap to 2x trailing fees',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@cap',
          find: 'shall not exceed the fees paid',
          replace: 'shall not exceed two times (2x) the fees paid',
        },
      },
      {
        note: 'Open a carve-out list under the cap',
        tool: 'docxodus_create',
        args: {
          action: 'insert_paragraph', anchorId: '@cap', position: 'after',
          markdown: 'The foregoing limitation shall not apply to:',
        },
        bind: 'carveIntro',
      },
      {
        note: 'Carve-out (a): confidentiality',
        tool: 'docxodus_create',
        args: {
          action: 'insert_paragraph', anchorId: '@carveIntro', position: 'after',
          markdown: '(a) breach of Section 4 (Confidentiality);',
        },
        bind: 'carveA',
      },
      {
        note: 'Carve-out (b): indemnities',
        tool: 'docxodus_create',
        args: {
          action: 'insert_paragraph', anchorId: '@carveA', position: 'after',
          markdown: "(b) a party's indemnification obligations; and",
        },
        bind: 'carveB',
      },
      {
        note: 'Carve-out (c): wilful misconduct',
        tool: 'docxodus_create',
        args: {
          action: 'insert_paragraph', anchorId: '@carveB', position: 'after',
          markdown: '(c) fraud, gross negligence, or wilful misconduct.',
        },
        bind: 'carveC',
      },
      {
        // Deliberate, and the most interesting frame in the show. Converting
        // those three paragraphs into a Word-native auto-numbered list means
        // writing `w:numPr`, and there is no tracked-change markup that
        // records "this paragraph joined a numbering instance" reversibly. So
        // the engine REFUSES it — nothing is written, rather than writing a
        // change that reject-all could not undo. That refusal is the same
        // guarantee the proof at the end verifies, enforced up front.
        note: 'Try to make them a Word-native (a)(b)(c) list — the engine should refuse',
        tool: 'docxodus_list',
        args: {
          action: 'apply_format_range',
          firstAnchorId: '@carveA', lastAnchorId: '@carveC',
          listFormat: 'lowerLetterParenthesis',
        },
        expectRefusal: true,
        refusalNote:
          'Refused, by design: no reversible tracked-change encoding for list membership. '
          + 'Nothing was written — an irreversible mark is worse than a missing one.',
      },
      {
        note: 'Leave a comment on the cap',
        tool: 'docxodus_comment',
        args: {
          action: 'add', anchorId: '@cap', author: COUNSEL.customer.name, initials: 'DW',
          markdown: 'Cap moved to 2x with standard carve-outs. This is the market position for ' +
            'a deal this size — happy to discuss the multiple, not the carve-outs.',
        },
      },
    ],
  },
  {
    id: 'act-2',
    title: 'Act II — Supplier pushes back',
    counsel: COUNSEL.supplier,
    synopsis:
      'The cap comes down to 1.5x, the limitation is made mutual, and a side-letter footnote ' +
      'records the trade. Marcus Oyelaran signs these.',
    steps: [
      {
        note: 'Find the cap again — it has moved',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'two times (2x)' },
        bind: 'cap2',
      },
      {
        note: 'Counter at 1.5x',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@cap2',
          find: 'two times (2x)', replace: 'one and one-half times (1.5x)',
        },
      },
      {
        note: 'Make the limitation mutual',
        tool: 'docxodus_create',
        args: {
          action: 'insert_paragraph', anchorId: '@cap2', position: 'before',
          markdown: 'This Section 5 applies equally to both parties.',
        },
        bind: 'mutual',
      },
      {
        note: 'Bold the mutuality, so it is not missed on a read-through',
        tool: 'docxodus_format',
        args: {
          action: 'apply_format_by_substring', anchorId: '@mutual',
          substring: 'applies equally to both parties', format: { bold: true },
        },
      },
      {
        // The marker goes at the END of the clause, not against the figure it
        // qualifies, and that is the engine's constraint rather than a style
        // choice: the "1.5x" text was written moments ago as a tracked insertion,
        // and citing into it puts the reference run inside a w:ins — a revision
        // nested in a revision, with no independently meaningful resolution.
        // InsertFootnote refuses that outright ("offset falls inside a revision
        // or unsupported inline container"). So the script searches the clause's
        // UNTOUCHED tail and cites past it, which is also where Word convention
        // puts a marker: after the closing period.
        note: 'Find the end of the clause, outside the marks, for the marker',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'months preceding the claim.' },
        bind: 'cap3',
        bindOffset: 'capMark',
      },
      {
        // A footnote is the beat this act wants, and until #625 it was the one
        // beat the demo could not have: `insert_footnote` under `render_inline`
        // wrote a definition into /word/footnotes.xml that reject-all left
        // behind, so the finale's reversibility proof failed on it (#614). The
        // citation is now the reversible unit — rejecting removes the reference
        // run, that leaves the note uncited, and the note-lifecycle rule prunes
        // the definition in the same resolve. The proof at the end of this run
        // is what re-checks that on every load.
        note: 'Footnote the trade that produced this number',
        tool: 'docxodus_create',
        args: {
          action: 'insert_footnote', anchorId: '@cap3', characterOffset: '@capMark',
          markdown: 'The parties agreed this multiple on March 11, 2026; it is recorded in the '
            + 'side letter of even date.',
        },
      },
      {
        note: 'Payment terms back out to Net 45',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'thirty (30) days of the date of invoice' },
        bind: 'payment2',
      },
      {
        note: 'Net 30 → Net 45',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@payment2',
          find: 'thirty (30) days', replace: 'forty-five (45) days',
        },
      },
      {
        note: 'Answer the comment on the record',
        tool: 'docxodus_comment',
        args: {
          action: 'add', anchorId: '@cap2', author: COUNSEL.supplier.name, initials: 'MO',
          markdown: '1.5x is our ceiling at this price point, and the limitation has to run ' +
            'both ways. Carve-outs (a) and (c) accepted as drafted.',
        },
      },
    ],
  },
  {
    id: 'act-3',
    title: 'Act III — Compliance sweeps the document',
    counsel: COUNSEL.compliance,
    synopsis:
      'Defined-term conformance ("Vendor" is not a defined term), a new data-protection ' +
      'section, and Governing Law moved to the end where it belongs.',
    steps: [
      {
        note: 'Find the first undefined "Vendor"',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'The Vendor shall not subcontract' },
        bind: 'vendor1',
      },
      {
        note: 'Vendor → Supplier (the defined term)',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@vendor1',
          find: 'The Vendor', replace: 'Supplier',
        },
      },
      {
        note: 'Find the second one',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'The Vendor shall issue invoices' },
        bind: 'vendor2',
      },
      {
        note: 'Conform it too',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@vendor2',
          find: 'The Vendor', replace: 'Supplier',
        },
      },
      {
        note: 'Find Governing Law — the new section goes before it',
        tool: 'docxodus_search',
        args: { mode: 'text', query: 'Section 8 — Governing Law' },
        bind: 'govHeading',
      },
      {
        note: 'Add a Data Protection heading',
        tool: 'docxodus_create',
        args: {
          action: 'insert_heading', anchorId: '@govHeading', position: 'before',
          text: 'Section 8 — Data Protection', level: 2,
        },
        bind: 'dpHeading',
      },
      {
        note: 'And the clause under it',
        tool: 'docxodus_create',
        args: {
          action: 'insert_paragraph', anchorId: '@dpHeading', position: 'after',
          markdown: 'Where Supplier processes personal data on behalf of Customer, the parties ' +
            'shall execute the Data Processing Addendum attached as Schedule 2, which is ' +
            'incorporated into this Agreement by reference.',
        },
        bind: 'dpClause',
      },
      {
        note: 'Renumber Governing Law to 9',
        tool: 'docxodus_edit',
        args: {
          action: 'replace_text_range', anchorId: '@govHeading',
          find: 'Section 8 — Governing Law', replace: 'Section 9 — Governing Law',
        },
      },
      {
        note: 'Emphasise the incorporation-by-reference',
        tool: 'docxodus_format',
        args: {
          action: 'apply_format_by_substring', anchorId: '@dpClause',
          substring: 'incorporated into this Agreement by reference',
          format: { italic: true },
        },
      },
      {
        note: 'File the compliance note',
        tool: 'docxodus_comment',
        args: {
          action: 'add', anchorId: '@dpClause', author: COUNSEL.compliance.name, initials: 'PR',
          markdown: 'DPA reference added — required before this can go to signature. ' +
            '"Vendor" conformed to the defined term "Supplier" in two places.',
        },
      },
    ],
  },
];

// ─── Telemetry (pure) ─────────────────────────────────────────────────

/** Roll tracked-change entries up by author, in the order the counsel appear.
 *  `listRevisions` reports the author stamped on the markup, which is what Word's
 *  Reviewing pane groups by — so this is the same breakdown a reviewer sees. */
export function attributionRollup(revisions, counsel = Object.values(COUNSEL)) {
  const counts = new Map();
  for (const revision of revisions ?? []) {
    const author = revision.author ?? '(unattributed)';
    counts.set(author, (counts.get(author) ?? 0) + 1);
  }
  const rows = counsel
    .map((c) => ({ name: c.name, role: c.role, color: c.color, count: counts.get(c.name) ?? 0 }))
    .filter((row) => row.count > 0);
  for (const [author, count] of counts) {
    if (!counsel.some((c) => c.name === author)) {
      rows.push({ name: author, role: 'other', color: '#64748b', count });
    }
  }
  return rows;
}

/** Throughput over a run: calls per second, from the wall clock the run actually
 *  took (pacing delays included, because that is what the viewer watched). */
export function throughput(callCount, elapsedMs) {
  if (!elapsedMs || elapsedMs <= 0) return 0;
  return (callCount / elapsedMs) * 1000;
}

/**
 * The exact set of package parts that leaving a review comment touches.
 *
 * Rejecting every tracked change cannot empty these, because a comment is not a
 * tracked change: `comments.xml` holds the comment bodies, `document.xml` carries
 * the `w:commentRangeStart`/`End` markers that anchor them, `styles.xml` gains the
 * `CommentText`/`CommentReference` styles, and the content-types map, the
 * document rels and settings record the new part. So a reject-path divergence
 * confined to these is expected — and a divergence ANYWHERE ELSE is not, which is
 * the point of writing the set out rather than matching a pattern. Being a closed
 * set is what caught `insert_footnote` leaving an unremovable `/word/footnotes.xml`
 * behind (issue #614, fixed in #625); a pattern like /comments|_rels/ would have
 * excused the note part too and the demo would have reported the redline clean.
 */
export const COMMENT_ATTRIBUTABLE_PARTS = new Set([
  '/[Content_Types].xml',
  '/word/_rels/document.xml.rels',
  '/word/comments.xml',
  '/word/document.xml',
  '/word/settings.xml',
  '/word/styles.xml',
]);

/** Split a reject-path divergence list into the parts comments explain and the
 *  parts nothing explains. Only the second kind impeaches a redline. With no
 *  comments in the document nothing is excusable, so everything is unexplained. */
export function classifyDivergences(divergences, commentCount = 0) {
  const explained = [];
  const unexplained = [];
  for (const divergence of divergences ?? []) {
    const excusable = commentCount > 0
      && COMMENT_ATTRIBUTABLE_PARTS.has(divergence.partUri ?? '');
    (excusable ? explained : unexplained).push(divergence);
  }
  return { explained, unexplained };
}

/**
 * The finale's verdict, from two independent engines.
 *
 * `proof` is `proveRedlineReversibility`, which rebuilds packages and compares
 * them part by part. `residual` is the DocxDiff revision list between the ORIGINAL
 * baseline and the document you get by rejecting every mark — the content-level
 * question, and the one that decides the headline: zero revisions means the text
 * came back exactly.
 *
 * The two disagree in a way worth showing rather than hiding. The package proof
 * reports the reject path as divergent whenever the reviewers left comments,
 * because `/word/comments.xml` is in the redline and not in the baseline and
 * rejecting cannot remove it. That is correct and expected: a comment is not a
 * revision. So the verdict is driven by content, and the comment parts are named
 * as what they are.
 */
export function proofVerdict(proof, residual, commentCount = 0) {
  if (!proof) return { ok: false, label: 'not run', detail: 'the proof has not been run yet' };

  const accepted = proof.acceptToFinal?.equivalent === true;
  const residualCount = residual?.length ?? 0;
  const { unexplained } = classifyDivergences(
    proof.rejectToBaseline?.divergences ?? proof.rejectToBaseline?.divergence,
    commentCount);

  if (accepted && residualCount === 0 && unexplained.length === 0) {
    const aside = commentCount > 0
      ? ` The ${commentCount} reviewer comment${commentCount === 1 ? '' : 's'} survive${
        commentCount === 1 ? 's' : ''} rejection — a comment is not a tracked change.`
      : '';
    return {
      ok: true,
      label: 'REVERSIBLE',
      detail: 'Accept-all reproduces the negotiated final. Reject-all restores the baseline '
        + `with zero content differences, confirmed independently by the diff engine.${aside}`,
    };
  }

  const reasons = [];
  if (!accepted) reasons.push('accept-all did not reproduce the negotiated final');
  if (residualCount > 0) {
    reasons.push(`${residualCount} content difference${residualCount === 1 ? '' : 's'} `
      + 'remain after reject-all');
  }
  if (unexplained.length > 0) {
    reasons.push(`${unexplained.length} unexplained package divergence${
      unexplained.length === 1 ? '' : 's'} (${
      unexplained.slice(0, 2).map((d) => d.partUri).join(', ')})`);
  }
  for (const finding of proof.findings ?? []) {
    reasons.push(String(finding.code ?? finding.message ?? 'finding'));
  }
  return {
    ok: false,
    label: 'NOT REVERSIBLE',
    detail: reasons.join('; ') || 'the proof reported failure without findings',
  };
}

/** The unid an editor block carries in `data-anchor`, from a `kind:scope:unid`
 *  anchor id (or from a bare unid, unchanged). */
export function unidOf(anchorId) {
  if (typeof anchorId !== 'string') return null;
  const parts = anchorId.split(':');
  return parts[parts.length - 1] || null;
}

// ─── Speeds ───────────────────────────────────────────────────────────

/** Pacing presets. `delayMs` is the gap between tool calls; MAX removes it and
 *  the run becomes engine-bound, which is the point of offering it. */
export const SPEEDS = [
  { id: 'read', label: '1×', delayMs: 260, hint: 'readable — one call at a time' },
  { id: 'brisk', label: '4×', delayMs: 65, hint: 'brisk — the intended tempo' },
  { id: 'max', label: 'MAX', delayMs: 0, hint: 'no pacing — engine-bound throughput' },
];

// ─── The stage panel (page chrome) ────────────────────────────────────

const PANEL_CSS = `
.dxt { display: flex; flex-direction: column; height: 100%; min-height: 0;
  font: 13px/1.5 system-ui, sans-serif; color: #e2e8f0; background: #0f172a;
  border-left: 1px solid #1e293b; }
.dxt * { box-sizing: border-box; }
.dxt-head { flex: none; padding: 12px 14px 10px; border-bottom: 1px solid #1e293b; }
.dxt-headrow { display: flex; align-items: center; gap: 10px; }
.dxt-brand { font: 700 14px/1 "SF Mono", Consolas, monospace; letter-spacing: .1em;
  color: #f87171; white-space: nowrap; }
.dxt-sub { font-size: 11px; color: #64748b; margin-top: 5px; }
.dxt-sub code { color: #94a3b8; font-size: 10.5px; }

.dxt-controls { display: flex; gap: 6px; padding: 10px 14px; border-bottom: 1px solid #1e293b;
  align-items: center; flex: none; flex-wrap: wrap; }
.dxt-controls button { font: 600 12px/1 system-ui, sans-serif; padding: 8px 12px;
  border: 1px solid #334155; border-radius: 8px; background: #1e293b; color: #e2e8f0;
  cursor: pointer; }
.dxt-controls button:hover:not([disabled]) { border-color: #475569; background: #273449; }
.dxt-controls button[disabled] { opacity: .45; cursor: default; }
.dxt-run { background: #dc2626 !important; border-color: #dc2626 !important; color: #fff !important;
  flex: 1; min-width: 120px; }
.dxt-run[disabled] { background: #334155 !important; border-color: #334155 !important; }
.dxt-speeds { display: flex; gap: 4px; margin-left: auto; }
.dxt-speeds button { padding: 7px 10px; font: 600 11px/1 "SF Mono", Consolas, monospace; }
.dxt-speeds button[aria-pressed="true"] { background: #f87171; border-color: #f87171; color: #0f172a; }

.dxt-hud { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1px; flex: none;
  background: #1e293b; border-bottom: 1px solid #1e293b; }
.dxt-hud div { background: #0f172a; padding: 9px 6px; text-align: center; }
.dxt-hud b { display: block; font: 700 16px/1.15 "SF Mono", Consolas, monospace; color: #f8fafc; }
.dxt-hud span { font-size: 9.5px; text-transform: uppercase; letter-spacing: .06em; color: #64748b; }

.dxt-modes { display: flex; gap: 4px; padding: 0 14px 10px; flex: none; }
.dxt-modes button { flex: 1; font: 600 11px/1 system-ui, sans-serif; padding: 7px 0;
  border: 1px solid #334155; border-radius: 8px; background: #1e293b; color: #cbd5e1;
  cursor: pointer; }
.dxt-modes button[aria-pressed="true"] { background: #f87171; border-color: #f87171;
  color: #0f172a; }
.dxt-pane { display: none; flex: 1; min-height: 0; overflow-y: auto; }
/* The proof and author roll-up describe the NEGOTIATION. While the stress meter
   is up they are stale context competing for the same column, so the foot
   collapses to the download button until the wire comes back. */
.dxt[data-mode="stress"] .dxt-proof,
.dxt[data-mode="stress"] .dxt-authors { display: none; }
.dxt-pane[data-on="true"] { display: block; }
.dxt-pane.dxt-wire[data-on="true"] { display: block; }

.dxt-now { flex: none; padding: 9px 14px; border-bottom: 1px solid #1e293b; min-height: 46px; }
.dxt-act { font: 600 11px/1.3 system-ui, sans-serif; }
.dxt-note { font-size: 12px; color: #cbd5e1; margin-top: 2px; }

.dxt-wire { flex: 1; min-height: 0; overflow-y: auto; overflow-x: hidden; padding: 8px 0;
  font: 11px/1.5 "SF Mono", Consolas, "Courier New", monospace;
  background: #020617; }
.dxt-frame { padding: 3px 12px; border-left: 2px solid transparent; word-break: break-word; }
.dxt-frame.req { color: #7dd3fc; }
.dxt-frame.res { color: #4ade80; padding-top: 0; padding-bottom: 7px; }
.dxt-frame.res.err { color: #fca5a5; }
.dxt-frame.res.refused { color: #fbbf24; }
.dxt-frame .dxt-id { color: #475569; }
.dxt-frame .dxt-method { color: #f8fafc; font-weight: 700; }
.dxt-frame .dxt-args { color: #64748b; }
.dxt-frame .dxt-ms { color: #fbbf24; }
.dxt-frame.live { border-left-color: #f87171; background: #0b1220; }

.dxt-foot { flex: none; padding: 10px 14px; border-top: 1px solid #1e293b; }
.dxt-proof { display: none; border-radius: 9px; padding: 10px 12px; margin-bottom: 9px;
  border: 1px solid #14532d; background: #052e16; }
.dxt-proof[data-on="true"] { display: block; }
.dxt-proof[data-ok="false"] { border-color: #7f1d1d; background: #300c0c; }
.dxt-proof b { font: 700 12px/1.3 "SF Mono", Consolas, monospace; color: #4ade80;
  letter-spacing: .05em; }
.dxt-proof[data-ok="false"] b { color: #fca5a5; }
.dxt-proof p { margin: 4px 0 0; font-size: 11.5px; color: #a7c8b4; }
.dxt-proof[data-ok="false"] p { color: #e2b4b4; }
.dxt-authors { display: flex; flex-direction: column; gap: 5px; margin-bottom: 9px; }
.dxt-author { display: flex; align-items: center; gap: 8px; font-size: 11.5px; }
.dxt-swatch { width: 9px; height: 9px; border-radius: 2px; flex: none; }
.dxt-author-n { color: #e2e8f0; }
.dxt-author-r { color: #64748b; font-size: 10.5px; }
.dxt-author-c { margin-left: auto; font: 700 12px/1 "SF Mono", Consolas, monospace; color: #f8fafc; }
.dxt-save { width: 100%; font: 600 12.5px/1 system-ui, sans-serif; padding: 9px 12px;
  border: 1px solid #334155; border-radius: 9px; background: #1e293b; color: #e2e8f0;
  cursor: pointer; }
.dxt-save[disabled] { opacity: .45; cursor: default; }
.dxt-error { display: none; margin-bottom: 9px; padding: 9px 11px; border-radius: 9px;
  background: #300c0c; border: 1px solid #7f1d1d; color: #fca5a5; font-size: 11.5px; }
.dxt-error[data-on="true"] { display: block; }

@media (max-width: 900px) {
  .dxt { border-left: 0; border-top: 1px solid #1e293b; }
  .dxt-hud b { font-size: 14px; }
  .dxt-wire { font-size: 10.5px; }
}
`;

/** Build the stage panel inside `root` and return the refs the driver wires.
 *  Pure DOM construction — no engine, no run state. */
export function mountTheaterPanel(root) {
  const style = document.createElement('style');
  style.textContent = PANEL_CSS;
  document.head.appendChild(style);

  root.classList.add('dxt');
  root.innerHTML = `
    <div class="dxt-head">
      <div class="dxt-headrow"><div class="dxt-brand">⚖ REDLINE THEATER</div></div>
      <div class="dxt-sub">Live <code>JSON-RPC 2.0</code> · the <code>docxodus-mcp</code> tool
        contract · native tracked changes, written as you watch</div>
    </div>
    <div class="dxt-controls">
      <button class="dxt-run" data-dxt="run">▶ Run the negotiation</button>
      <button data-dxt="reset" title="Rebuild the baseline and clear the wire">↻</button>
      <div class="dxt-speeds" data-dxt="speeds" role="group" aria-label="Playback speed"></div>
    </div>
    <div class="dxt-hud">
      <div title="MCP tool calls dispatched"><b data-dxt="calls">0</b><span>calls</span></div>
      <div title="Tracked changes written into the document"><b data-dxt="revisions">0</b><span>revisions</span></div>
      <div title="Median tool-call duration, measured in this tab"><b data-dxt="p50">–</b><span>p50 ms</span></div>
      <div title="Tool calls per second over the run, pacing included"><b data-dxt="rate">–</b><span>calls/s</span></div>
    </div>
    <div class="dxt-now">
      <div class="dxt-act" data-dxt="act">Idle — press Run</div>
      <div class="dxt-note" data-dxt="note">The document below is clean. Nothing is recorded yet.</div>
    </div>
    <div class="dxt-modes" data-dxt="modes" role="group" aria-label="Panel mode">
      <button data-mode="wire" aria-pressed="true"
        title="The MCP frames the negotiation dispatches">MCP wire</button>
      <button data-mode="stress" aria-pressed="false"
        title="Run the comparison engine flat out and measure it">Diff stress</button>
    </div>
    <div class="dxt-wire dxt-pane" data-dxt="wire" data-on="true" role="log" aria-label="MCP wire"></div>
    <div class="dxt-pane" data-dxt="stress"></div>
    <div class="dxt-foot">
      <div class="dxt-error" data-dxt="error"></div>
      <div class="dxt-proof" data-dxt="proof"></div>
      <div class="dxt-authors" data-dxt="authors"></div>
      <button class="dxt-save" data-dxt="save" disabled
        title="Download the redline — a real .docx with native tracked changes">⬇ Download the redline (.docx)</button>
    </div>`;

  const grab = (name) => root.querySelector(`[data-dxt="${name}"]`);
  return {
    ui: {
      panel: root,
      run: grab('run'), reset: grab('reset'), speeds: grab('speeds'),
      calls: grab('calls'), revisions: grab('revisions'), p50: grab('p50'), rate: grab('rate'),
      act: grab('act'), note: grab('note'), wire: grab('wire'),
      modes: grab('modes'), stress: grab('stress'),
      error: grab('error'), proof: grab('proof'), authors: grab('authors'), save: grab('save'),
    },
  };
}

// ─── The driver ───────────────────────────────────────────────────────

const check = (result, what) => {
  if (!result?.success) {
    throw new Error(`${what} failed: ${result?.error?.code ?? ''} ${result?.error?.message ?? ''}`);
  }
  return result;
};

/** Build the baseline into a session whose body is one blank paragraph. Headings
 *  are applied as `setParagraphStyle` — the exact shape the ribbon's Style
 *  dropdown writes — rather than markdown ATX, which would also attach outline
 *  numbering and make the baseline something the UI could not have produced. */
function buildBaseline(session) {
  const first = session.findByKind('p', 'body')[0];
  if (!first) throw new Error('expected a body paragraph to build on');
  const heading = (line) => {
    const m = /^(#{1,3}) (.*)$/.exec(line);
    return m ? { style: `Heading${m[1].length}`, text: m[2] } : { style: null, text: line };
  };
  const seed = heading(BASELINE[0]);
  let prev = check(session.replaceText(first.id, seed.text), 'baseline seed').modified?.[0]?.id
    ?? first.id;
  if (seed.style) {
    prev = check(session.setParagraphStyle(prev, seed.style), 'seed style').modified?.[0]?.id ?? prev;
  }
  let paymentAnchor = null;
  for (let i = 1; i < BASELINE.length; i++) {
    const { style, text } = heading(BASELINE[i]);
    let id = check(session.insertParagraph(prev, 'after', text), `baseline ¶${i + 1}`)
      .created?.[0]?.id;
    if (style) {
      id = check(session.setParagraphStyle(id, style), `baseline ¶${i + 1} style`)
        .modified?.[0]?.id ?? id;
    }
    if (text.startsWith('Customer shall pay the fees')) paymentAnchor = id;
    prev = id ?? prev;
  }
  // The fee table sits under the payment paragraph — a real table, so the
  // negotiation runs over a document with more than paragraphs in it.
  if (paymentAnchor) {
    check(
      session.insertTable(paymentAnchor, 'after', FEE_TABLE.rows, FEE_TABLE.columns,
        { cellContents: FEE_TABLE.cellContents }),
      'fee table',
    );
  }
}

/** Clear the body back to one seed paragraph, sweeping every kind an act can
 *  create — headings are kind `h`, notes live in their own parts, and a missed
 *  kind haunts the next run as content only the diff engine can see. */
function resetBody(session) {
  const blocks = () => ['p', 'h', 'li', 'tbl'].flatMap((kind) => session.findByKind(kind, 'body'));
  const first = blocks()[0];
  if (!first) throw new Error('document has no body blocks');
  const seed = check(session.insertParagraph(first.id, 'before', '(setting the stage…)'), 'seed');
  const seedId = seed.created[0].id;
  for (const block of blocks()) {
    if (block.id === seedId) continue;
    check(session.deleteBlock(block.id), `clearing ${block.id}`);
  }
  for (const kind of ['fn', 'en']) {
    for (const note of session.findByKind(kind)) {
      check(session.deleteBlock(note.id), `clearing ${kind} ${note.id}`);
    }
  }
  for (const comment of session.listComments()) {
    const id = comment.anchorId ?? comment.id;
    if (id) session.removeComment(id);
  }
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const macrotask = () => new Promise((resolve) => setTimeout(resolve, 0));

/**
 * Run the negotiation against a ribbon-hosted editor.
 *
 * `engine` is the docxodus module namespace (the page passes its pinned import),
 * of which the driver uses `TrackedChangeMode` (the recording mode the show turns
 * on) and `proveRedlineReversibility` (the finale). Everything between those two
 * goes through the MCP endpoint, which goes through `DocxSession` — the driver
 * never calls a mutation method directly once the show has started.
 *
 * Returns the controller the host page publishes as `window.__theater`.
 */
export function startTheater({ editor, session, engine, ui, script = SCRIPT, autoRun = false }) {
  const problems = validateScript(script);
  if (problems.length) throw new Error(`script invalid: ${problems.join('; ')}`);

  const endpoint = createBrowserMcpEndpoint({
    session,
    sessionId: 'sess_theater_01',
    // No `onMutate` here on purpose: `dispatchStep` schedules the repaint
    // itself, because it also knows WHICH anchor the call touched and the
    // camera follows it. Wiring both would schedule every repaint twice.
    // Inject the real enum rather than leaning on the wire module's copy of its
    // ordinals — the engine is right here, and the recording mode is the one
    // setting the whole show depends on being correct.
    trackedChangeModes: {
      accept: engine.TrackedChangeMode.Accept,
      render_inline: engine.TrackedChangeMode.RenderInline,
      strip_deletions: engine.TrackedChangeMode.StripDeletions,
    },
  });

  const renderStats = createLatencyStats();
  let speed = SPEEDS[1];
  let running = false;
  let cancelled = false;
  let baselineBytes = null;
  let redlineBytes = null;
  let callCount = 0;
  let refusals = 0;
  let runStartedAt = 0;
  const bindings = Object.create(null);

  // ── the frame-dropped repaint ────────────────────────────────────
  // Ops keep flowing while a repaint is in flight; the next animation frame
  // renders whatever state the document has reached by then. Several ops
  // between two frames coalesce into ONE refresh, which is the batching
  // contract `DocxEditor.refresh()` documents for a host driving the session.
  let repaintQueued = false;
  let spotlightAnchor = null;

  function scheduleRepaint(anchorId) {
    if (anchorId) spotlightAnchor = anchorId;
    if (repaintQueued) return;
    repaintQueued = true;
    requestAnimationFrame(() => {
      repaintQueued = false;
      const target = spotlightAnchor;
      spotlightAnchor = null;
      const started = performance.now();
      try {
        editor.refresh();
      } catch (err) {
        // A repaint that loses a race with a structural edit is not fatal: the
        // next frame renders the settled document. Losing the RUN to it would be.
        console.warn('[theater] repaint skipped', err);
        return;
      }
      renderStats.record('refresh', performance.now() - started);
      if (target) spotlight(target);
    });
  }

  /** Scroll the changed block into view and pulse it — the camera. The editor
   *  stamps each block's bare unid as `data-anchor`, so the anchor an EditResult
   *  reports addresses the element directly. */
  function spotlight(anchorId) {
    const unid = unidOf(anchorId);
    if (!unid) return;
    const surface = ui.panel.ownerDocument.querySelector('[data-dxr-surface]');
    const el = surface?.querySelector(`[data-anchor="${unid}"]`)
      ?? surface?.querySelector(`[data-anchor$="${unid}"]`);
    if (!el) return;
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    if (typeof el.animate !== 'function') return;
    if (window.matchMedia?.('(prefers-reduced-motion: reduce)')?.matches) return;
    el.animate(
      [{ background: 'rgba(248,113,113,.28)' }, { background: 'rgba(248,113,113,0)' }],
      { duration: 620, easing: 'ease-out' },
    );
  }

  // ── the wire console ─────────────────────────────────────────────
  const MAX_WIRE_ROWS = 400;

  function pushFrame(kind, html, live) {
    const row = document.createElement('div');
    row.className = `dxt-frame ${kind}${live ? ' live' : ''}`;
    row.innerHTML = html;
    ui.wire.appendChild(row);
    while (ui.wire.childElementCount > MAX_WIRE_ROWS) ui.wire.firstElementChild.remove();
    ui.wire.scrollTop = ui.wire.scrollHeight;
    return row;
  }

  const esc = (text) => String(text).replace(/[&<>]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

  function logCall({ request, response, ms, isError, result }, expectedRefusal = false) {
    pushFrame('req',
      `<span class="dxt-id">→ #${request.id}</span> `
      + `<span class="dxt-method">${esc(frameLabel(request))}</span> `
      + `<span class="dxt-args">${esc(summarizeArgs(request.params.arguments))}</span>`,
      true);
    // An expected refusal is neither a success nor a fault — it is the engine
    // declining to write something it could not take back, so it gets its own
    // colour rather than the red an unplanned failure earns.
    const kind = isError ? (expectedRefusal ? 'refused' : 'err') : '';
    const glyph = isError ? (expectedRefusal ? '⊘' : '✕') : '✓';
    const summary = isError
      ? esc(result?.error?.message ?? result?.error ?? 'error')
      : resultSummary(result);
    pushFrame(`res${kind ? ' ' + kind : ''}`,
      `<span class="dxt-id">← #${request.id}</span> `
      + `${glyph} ${summary} <span class="dxt-ms">${ms.toFixed(1)}ms</span>`);
  }

  /** One line about what a tool result actually did — the counts a reviewer
   *  would care about, not the whole payload. */
  function resultSummary(result) {
    if (!result || typeof result !== 'object') return 'ok';
    if (Array.isArray(result.matches)) return `${result.matches.length} match(es)`;
    if (Array.isArray(result.anchors)) return `${result.anchors.length} anchor(s)`;
    if (Array.isArray(result.revisions)) return `${result.revisions.length} revision(s)`;
    if (Array.isArray(result.comments)) return `${result.comments.length} comment(s)`;
    const bits = [];
    if (result.created?.length) bits.push(`+${result.created.length} created`);
    if (result.modified?.length) bits.push(`~${result.modified.length} modified`);
    if (result.removed?.length) bits.push(`-${result.removed.length} removed`);
    return bits.length ? bits.join(', ') : 'ok';
  }

  // ── HUD ──────────────────────────────────────────────────────────
  function paintHud() {
    ui.calls.textContent = String(callCount);
    ui.p50.textContent = endpoint.stats.count()
      ? endpoint.stats.percentile(50).toFixed(1) : '–';
    const elapsed = runStartedAt ? performance.now() - runStartedAt : 0;
    ui.rate.textContent = elapsed > 0 ? throughput(callCount, elapsed).toFixed(1) : '–';
    try {
      ui.revisions.textContent = String(session.listRevisions().length);
    } catch { /* mid-edit; the next paint catches up */ }
  }

  function setStage(act, note) {
    if (act) {
      ui.act.innerHTML = act.counsel
        ? `<span style="color:${act.counsel.color}">●</span> ${esc(act.title)} `
          + `<span style="color:#64748b">· ${esc(act.counsel.name)}, ${esc(act.counsel.role)}</span>`
        : esc(act.title);
    }
    if (note !== undefined) ui.note.textContent = note;
  }

  function fail(err) {
    ui.error.dataset.on = 'true';
    ui.error.textContent = String(err?.message ?? err).slice(0, 300);
    console.error('[theater]', err);
  }

  // ── speeds ───────────────────────────────────────────────────────
  const speedButtons = SPEEDS.map((preset) => {
    const b = document.createElement('button');
    b.textContent = preset.label;
    b.title = preset.hint;
    b.setAttribute('aria-pressed', String(preset === speed));
    b.addEventListener('click', () => {
      speed = preset;
      for (const other of speedButtons) {
        other.setAttribute('aria-pressed', String(other === b));
      }
    });
    ui.speeds.appendChild(b);
    return b;
  });

  // ── the run ──────────────────────────────────────────────────────

  /** Dispatch one scripted step as a real MCP frame. */
  function dispatchStep(step) {
    const args = resolveArgs(step.args ?? {}, bindings);
    const outcome = endpoint.call(step.tool, args);
    callCount++;
    logCall(outcome, step.expectRefusal);
    if (outcome.isError) {
      // A step marked `expectRefusal` is demonstrating the engine's fail-closed
      // guarantee, so its refusal is the POINT, not a fault. Anything else is.
      if (!step.expectRefusal) {
        throw new Error(
          `${frameLabel(outcome.request)} failed: `
          + `${outcome.result?.error?.message ?? outcome.result?.error ?? 'unknown'}`);
      }
      refusals++;
      setStage(null, step.refusalNote ?? step.note);
      return outcome;
    }
    if (step.bind) {
      const bound = bindFromResult(outcome.result);
      if (!bound) throw new Error(`step "${step.note}" bound nothing to @${step.bind}`);
      bindings[step.bind] = bound;
    }
    if (step.bindOffset) {
      const offset = offsetAfterMatch(outcome.result);
      if (offset === null) {
        throw new Error(`step "${step.note}" bound no offset to @${step.bindOffset}`);
      }
      bindings[step.bindOffset] = offset;
    }
    // Spotlight what this call touched; a search touches nothing.
    const touched = bindFromResult(outcome.result);
    if (touched && step.tool !== 'docxodus_search') scheduleRepaint(touched);
    return outcome;
  }

  async function runScript() {
    if (running) return;
    running = true;
    cancelled = false;
    ui.error.dataset.on = 'false';
    ui.proof.dataset.on = 'false';
    ui.save.disabled = true;
    ui.run.textContent = '■ Stop';
    try {
      await stage();
      runStartedAt = performance.now();
      for (const act of script) {
        if (cancelled) break;
        setStage(act, act.synopsis);
        // Attribution is a SESSION setting, not a tool: the server takes
        // `revisionAuthor` at `docxodus_open`, so three counsel in one session is
        // a host-side switch. Every mark written after this line carries this name.
        session.setRevisionAuthor(act.counsel.name);
        for (const step of act.steps) {
          if (cancelled) break;
          setStage(null, step.note);
          dispatchStep(step);
          paintHud();
          if (speed.delayMs > 0) await sleep(speed.delayMs);
          else await macrotask(); // yield so the repaint frame can run
        }
      }
      if (!cancelled) await finale();
    } catch (err) {
      fail(err);
    } finally {
      running = false;
      ui.run.textContent = '▶ Run the negotiation';
      paintHud();
    }
  }

  /** Rebuild the baseline: recording OFF while the clean document is authored,
   *  then ON for the show. Both mode switches go over the wire, because
   *  `docxodus_track_changes set_mode` is a real tool — it is how an agent that
   *  wants a clean setup pass followed by a recorded pass does it. */
  async function stage() {
    setStage({ title: 'Setting the stage', counsel: null }, 'Authoring the clean baseline…');
    ui.wire.innerHTML = '';
    callCount = 0;
    refusals = 0;
    runStartedAt = 0;
    redlineBytes = null;
    for (const key of Object.keys(bindings)) delete bindings[key];

    logCall(endpoint.call('docxodus_track_changes',
      { action: 'set_mode', mode: 'accept' }));
    callCount++;

    session.setRevisionAuthor(null);
    resetBody(session);
    buildBaseline(session);
    editor.refresh();
    baselineBytes = session.save();

    logCall(endpoint.call('docxodus_track_changes',
      { action: 'set_mode', mode: 'render_inline' }));
    callCount++;
    paintHud();
    await macrotask();
  }

  /** The point of the exercise: the redline is checked, not asserted. */
  async function finale() {
    setStage({ title: 'Proving the redline', counsel: null },
      'Rebuilding two packages from the redline — accept-all and reject-all…');
    redlineBytes = session.save();

    const revisions = session.listRevisions();
    paintAuthors(attributionRollup(revisions));

    // The intended final is what accepting every mark produces. Building it in a
    // throwaway session keeps the live document — the thing on screen and the
    // thing that downloads — untouched by the proof.
    const shadow = engine.openDocxSession(redlineBytes);
    let intendedFinal;
    try {
      shadow.acceptAllRevisions();
      intendedFinal = shadow.save();
    } finally {
      shadow.close();
    }

    // And the reject path: what the document becomes when every mark is refused.
    const rejectShadow = engine.openDocxSession(redlineBytes);
    let rejected;
    try {
      rejectShadow.rejectAllRevisions();
      rejected = rejectShadow.save();
    } finally {
      rejectShadow.close();
    }

    // Two engines, two questions. The package proof rebuilds and compares parts;
    // the diff engine answers the content question the headline turns on — does
    // rejecting every mark put the text back exactly as it was?
    setStage(null, 'Comparing the rejected document against the baseline…');
    const [proof, residual] = await Promise.all([
      engine.proveRedlineReversibility(baselineBytes, intendedFinal, redlineBytes),
      engine.docxDiffGetRevisions(baselineBytes, rejected),
    ]);

    const comments = session.listComments().length;
    const verdict = proofVerdict(proof, residual, comments);
    ui.proof.dataset.on = 'true';
    ui.proof.dataset.ok = String(verdict.ok);
    ui.proof.innerHTML =
      `<b>${esc(verdict.label)}</b><p>${esc(verdict.detail)}</p>`
      + `<p>${revisions.length} tracked change(s) across ${
        attributionRollup(revisions).length} author(s)`
      + `${refusals ? `, ${refusals} refused as irreversible` : ''}`
      + `, in ${endpoint.stats.total().toFixed(0)}ms of tool time.</p>`;

    ui.save.disabled = false;
    setStage({ title: 'Negotiation complete', counsel: null },
      verdict.ok
        ? 'Every mark rejects back to the baseline. Download it and open it in Word.'
        : 'The proof did not pass — see the panel below.');
    paintHud();
    return { proof, residual, verdict };
  }

  function paintAuthors(rows) {
    ui.authors.innerHTML = '';
    for (const row of rows) {
      const el = document.createElement('div');
      el.className = 'dxt-author';
      el.innerHTML =
        `<i class="dxt-swatch" style="background:${row.color}"></i>`
        + `<span class="dxt-author-n">${esc(row.name)}</span>`
        + `<span class="dxt-author-r">${esc(row.role)}</span>`
        + `<span class="dxt-author-c">${row.count}</span>`;
      ui.authors.appendChild(el);
    }
  }

  // ── the diff-stress mode ─────────────────────────────────────────
  // The theater RECORDS its redline; this COMPUTES one from scratch after every
  // edit and times it. Both halves go through the same MCP endpoint, so the
  // mutation cost the meter compares against is the real one, not a shortcut.
  const { ui: stressUi } = mountStressPanel(ui.stress);
  let stressClause = null;

  const stress = createStressRunner({
    engine,
    session,
    ui: stressUi,
    /** Append one clause through a real `docxodus_create` frame and return what
     *  the mutation cost. Each frame lands after the previous one, so the
     *  document grows downward the way a document actually does. */
    applyEdit: async (frame) => {
      const anchorId = stressClause
        ?? session.findByKind('p', 'body').slice(-1)[0]?.id;
      if (!anchorId) throw new Error('stress: no body paragraph to append to');
      const outcome = endpoint.call('docxodus_create', {
        action: 'insert_paragraph',
        anchorId,
        position: 'after',
        markdown: clauseFor(frame),
      });
      callCount++;
      if (outcome.isError) {
        throw new Error(`stress frame ${frame} failed: `
          + `${outcome.result?.error?.message ?? 'unknown'}`);
      }
      stressClause = bindFromResult(outcome.result) ?? anchorId;
      scheduleRepaint(stressClause);
      paintHud();
      return outcome.ms;
    },
    onFrame: (result) => {
      // At `full` depth the computed redline is rendered as HTML — show it, so
      // the mode is not only numbers. The editor beside it keeps showing the
      // RECORDED markup, which is the contrast worth seeing.
      if (result.html) showComputedRedline(result.html);
    },
  });

  /** Render the computed redline into the stage caption area's sibling frame.
   *  Kept sandboxed: it is generated HTML, shown as a static preview. */
  let computedFrame = null;
  function showComputedRedline(html) {
    if (!computedFrame) {
      computedFrame = ui.panel.ownerDocument.createElement('iframe');
      computedFrame.setAttribute('sandbox', '');
      computedFrame.style.cssText =
        'display:block;width:100%;height:210px;border:1px solid #1e293b;'
        + 'border-radius:9px;background:#fff;margin-top:9px';
      stressUi.panel.appendChild(computedFrame);
    }
    computedFrame.srcdoc = html;
  }

  const modeButtons = [...ui.modes.querySelectorAll('button')];
  function setMode(mode) {
    for (const b of modeButtons) b.setAttribute('aria-pressed', String(b.dataset.mode === mode));
    ui.wire.dataset.on = String(mode === 'wire');
    ui.stress.dataset.on = String(mode === 'stress');
    ui.panel.dataset.mode = mode;
  }
  for (const b of modeButtons) {
    b.addEventListener('click', () => {
      // Never leave a stress loop running behind a hidden pane.
      if (b.dataset.mode !== 'stress') stress.stop();
      setMode(b.dataset.mode);
    });
  }
  setMode('wire');

  // ── wiring ───────────────────────────────────────────────────────
  ui.run.addEventListener('click', () => {
    if (running) { cancelled = true; return; }
    void runScript();
  });
  ui.reset.addEventListener('click', () => {
    if (running) cancelled = true;
    stress.stop();
    stressClause = null;
    ui.proof.dataset.on = 'false';
    ui.authors.innerHTML = '';
    ui.save.disabled = true;
    void stage().then(() => setStage({ title: 'Idle — press Run', counsel: null },
      'The document below is clean. Nothing is recorded yet.')).catch(fail);
  });
  ui.save.addEventListener('click', () => {
    const bytes = redlineBytes ?? session.save();
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'msa-redline.docx';
    a.click();
    URL.revokeObjectURL(url);
  });

  const ready = stage()
    .then(() => {
      setStage({ title: 'Idle — press Run', counsel: null },
        'The document below is clean. Nothing is recorded yet.');
      if (autoRun) return runScript();
      return undefined;
    })
    .catch(fail);

  return {
    ready,
    run: runScript,
    stage,
    stop: () => { cancelled = true; },
    setSpeed: (id) => {
      const preset = SPEEDS.find((s) => s.id === id);
      if (preset) {
        speed = preset;
        speedButtons.forEach((b, i) => b.setAttribute('aria-pressed', String(SPEEDS[i] === preset)));
      }
    },
    endpoint,
    stress,
    setMode,
    stats: () => ({
      calls: callCount,
      refusals,
      p50: endpoint.stats.percentile(50),
      p95: endpoint.stats.percentile(95),
      maxMs: endpoint.stats.max(),
      toolMs: endpoint.stats.total(),
      renderCount: renderStats.count(),
      renderP50: renderStats.percentile(50),
      revisions: session.listRevisions().length,
      perTool: endpoint.stats.perTool(),
    }),
    baseline: () => baselineBytes,
    redline: () => redlineBytes,
    isRunning: () => running,
  };
}
