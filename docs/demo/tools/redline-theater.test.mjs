// Headless logic checks for REDLINE THEATER (docs/demo/redline-theater.js and
// docs/demo/mcp-wire.js) — the pure, DOM-free parts: JSON-RPC frame shapes,
// binding substitution, script validation, telemetry, and the attribution
// roll-up. Everything that needs the engine (the recorded redline, the
// reversibility proof, the measured throughput) is proven in the browser by
// npm/tests/demo-redline.spec.ts.
//
// The load-bearing test in this file is the last one: the demo claims its frames
// are the frames `docxodus-mcp` accepts, so this parses the REAL
// tools/mcp-server/ToolCatalog.cs and checks every (tool, action) pair the
// browser endpoint routes — and every argument the script sends — against the
// shipped catalog. Rename an action in the server and this fails here.
import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import {
  JSONRPC_VERSION,
  JSON_RPC_ERRORS,
  IMPLEMENTED_TOOLS,
  createLatencyStats,
  errorFrame,
  escapeRegex,
  frameLabel,
  implementedPairs,
  looksLikeFailure,
  requestFrame,
  summarizeArgs,
  toolResultFrame,
  resultMutated,
  TRACKED_CHANGE_MODES,
} from '../mcp-wire.js';

import {
  BASELINE,
  COUNSEL,
  SCRIPT,
  SPEEDS,
  attributionRollup,
  bindFromResult,
  offsetAfterMatch,
  classifyDivergences,
  COMMENT_ATTRIBUTABLE_PARTS,
  proofVerdict,
  referencesIn,
  resolveArgs,
  throughput,
  unidOf,
  validateScript,
} from '../redline-theater.js';

// ─── Frame shapes ─────────────────────────────────────────────────────

test('requestFrame is a well-formed JSON-RPC 2.0 tools/call', () => {
  const frame = requestFrame(7, 'docxodus_edit', { action: 'undo' });
  assert.equal(frame.jsonrpc, JSONRPC_VERSION);
  assert.equal(frame.id, 7);
  assert.equal(frame.method, 'tools/call');
  assert.equal(frame.params.name, 'docxodus_edit');
  assert.deepEqual(frame.params.arguments, { action: 'undo' });
});

test('a tool result carries the tool JSON as text content, with isError', () => {
  const ok = toolResultFrame(1, '{"success":true}', false);
  assert.equal(ok.result.isError, false);
  assert.equal(ok.result.content[0].type, 'text');
  assert.equal(ok.result.content[0].text, '{"success":true}');

  const bad = toolResultFrame(2, '{"success":false}', true);
  assert.equal(bad.result.isError, true);
});

test('protocol errors are envelope errors, and carry the standard codes', () => {
  const frame = errorFrame(3, JSON_RPC_ERRORS.methodNotFound, 'method not found: nope');
  assert.equal(frame.error.code, -32601);
  assert.equal(frame.id, 3);
  // A frame with no id still answers with an explicit null, not a missing key.
  assert.equal(errorFrame(undefined, JSON_RPC_ERRORS.parseError, 'x').id, null);
});

test('looksLikeFailure matches the server heuristic: top-level success:false', () => {
  assert.equal(looksLikeFailure('{"success":false,"error":{"code":"anchor_not_found"}}'), true);
  assert.equal(looksLikeFailure('{"success":true}'), false);
  assert.equal(looksLikeFailure('{"matches":[]}'), false);
  // An array of results is not a failure envelope, even if one element failed.
  assert.equal(looksLikeFailure('[{"success":false}]'), false);
  assert.equal(looksLikeFailure('not json'), false);
});

test('frameLabel names the action, mode, or format discriminator', () => {
  assert.equal(
    frameLabel(requestFrame(1, 'docxodus_edit', { action: 'replace_text' })),
    'docxodus_edit/replace_text');
  assert.equal(
    frameLabel(requestFrame(1, 'docxodus_search', { mode: 'text' })),
    'docxodus_search/text');
  assert.equal(
    frameLabel(requestFrame(1, 'docxodus_comment', { action: 'list' })),
    'docxodus_comment/list');
});

test('summarizeArgs elides long values in the middle and hides the session id', () => {
  const line = summarizeArgs({ sessionId: 'sess_1', anchorId: 'p:body:abc', markdown: 'x'.repeat(60) });
  assert.ok(!line.includes('sess_1'), 'session id is noise on every line');
  assert.ok(line.includes('anchorId'));
  assert.ok(line.includes('…'), 'a 60-char payload is elided');
  assert.ok(line.length <= 88);
});

test('escapeRegex keeps docxodus_search mode "text" literal', () => {
  // The catalog documents mode "text" as a LITERAL substring search, but
  // DocxSession.grep takes a regex — "(2x)" must not become a capture group.
  assert.equal(escapeRegex('two times (2x)'), 'two times \\(2x\\)');
  assert.equal(new RegExp(escapeRegex('a.b')).test('a.b'), true);
  assert.equal(new RegExp(escapeRegex('a.b')).test('axb'), false);
});

test('resultMutated tells a repaint-worthy call from a read', () => {
  // A host that repaints on every successful call repaints after searches too,
  // which touch nothing — wasted work, and it hides whether repaints coalesce.
  assert.equal(resultMutated('docxodus_search', { matches: [{}, {}] }), false);
  assert.equal(resultMutated('docxodus_comment', { comments: [{}] }), false);
  assert.equal(resultMutated('docxodus_track_changes', { revisions: [{}] }), false);
  assert.equal(resultMutated('docxodus_edit', { success: true, modified: [{ id: 'p:body:a' }] }), true);
  assert.equal(resultMutated('docxodus_create', { success: true, created: [{ id: 'p:body:b' }] }), true);
  assert.equal(resultMutated('docxodus_edit', { success: false }), false);
  // Switching the recording mode is session configuration, not a document edit.
  assert.equal(resultMutated('docxodus_track_changes', { success: true, mode: 'accept' }), false);
  // Undo reports nothing but its own success, and does need a repaint.
  assert.equal(resultMutated('docxodus_edit', { success: true }), true);
  // A batch mutated if any of its steps did.
  assert.equal(resultMutated('docxodus_mutations', {
    success: true,
    applied: [{ tool: 'docxodus_search', result: { matches: [{}] } }],
  }), false);
  assert.equal(resultMutated('docxodus_mutations', {
    success: true,
    applied: [
      { tool: 'docxodus_search', result: { matches: [{}] } },
      { tool: 'docxodus_edit', result: { success: true, modified: [{ id: 'x' }] } },
    ],
  }), true);
});

// ─── Latency statistics ───────────────────────────────────────────────

test('latency percentiles are exact nearest-rank over the samples', () => {
  const stats = createLatencyStats();
  for (const ms of [10, 1, 5, 3, 9, 2, 7, 4, 8, 6]) stats.record('docxodus_edit', ms);
  assert.equal(stats.count(), 10);
  assert.equal(stats.percentile(50), 5);
  assert.equal(stats.percentile(100), 10);
  assert.equal(stats.max(), 10);
  assert.equal(stats.total(), 55);
  // An empty set reports zero rather than NaN — the HUD renders it.
  assert.equal(createLatencyStats().percentile(50), 0);
});

test('per-tool roll-up orders by call count', () => {
  const stats = createLatencyStats();
  stats.record('docxodus_edit', 1);
  stats.record('docxodus_edit', 2);
  stats.record('docxodus_search', 4);
  const rows = stats.perTool();
  assert.equal(rows[0].tool, 'docxodus_edit');
  assert.equal(rows[0].calls, 2);
  assert.equal(rows[0].totalMs, 3);
});

test('throughput is calls per second and survives a zero clock', () => {
  assert.equal(throughput(30, 1000), 30);
  assert.equal(throughput(30, 0), 0);
});

// ─── Bindings ─────────────────────────────────────────────────────────

test('bindFromResult reads the reusable anchor out of each result shape', () => {
  assert.equal(
    bindFromResult({ matches: [{ enclosingAnchor: { id: 'p:body:a1' } }] }), 'p:body:a1');
  assert.equal(bindFromResult({ anchors: [{ id: 'h:body:b2' }] }), 'h:body:b2');
  assert.equal(bindFromResult({ created: [{ id: 'p:body:c3' }] }), 'p:body:c3');
  // A mutation that both modified and created binds to what it MODIFIED: that is
  // the block the caller addressed, and the one the camera should follow.
  assert.equal(
    bindFromResult({ modified: [{ id: 'p:body:d4' }], created: [{ id: 'p:body:e5' }] }),
    'p:body:d4');
  assert.equal(bindFromResult({ matches: [] }), null);
  assert.equal(bindFromResult(null), null);
});

test('resolveArgs substitutes @refs deeply and leaves plain strings alone', () => {
  const bindings = { cap: 'p:body:cap1', note: 'p:body:n2' };
  assert.deepEqual(
    resolveArgs({ anchorId: '@cap', find: 'sixty (60)', nested: { to: '@note' } }, bindings),
    { anchorId: 'p:body:cap1', find: 'sixty (60)', nested: { to: 'p:body:n2' } });
  assert.deepEqual(resolveArgs(['@cap', 'literal'], bindings), ['p:body:cap1', 'literal']);
});

test('an unbound @ref throws rather than editing a block named "@cap"', () => {
  assert.throws(() => resolveArgs({ anchorId: '@cap' }, {}), /unbound script reference: @cap/);
  assert.throws(() => resolveArgs({ anchorId: '@cap' }, { cap: null }), /unbound/);
});

test('referencesIn finds every @ref in a step', () => {
  assert.deepEqual(
    referencesIn({ a: '@one', b: { c: ['@two', 'plain'] }, d: 3 }).sort(),
    ['one', 'two']);
});

test('unidOf takes the last segment of a kind:scope:unid anchor', () => {
  assert.equal(unidOf('p:body:deadbeef'), 'deadbeef');
  assert.equal(unidOf('deadbeef'), 'deadbeef');
  assert.equal(unidOf(null), null);
});

// ─── Script validation ────────────────────────────────────────────────

test('the shipped script validates', () => {
  assert.deepEqual(validateScript(SCRIPT), []);
});

test('validateScript rejects an unknown tool or action', () => {
  const bad = [{
    id: 'x', title: 'x', counsel: COUNSEL.customer,
    steps: [{ tool: 'docxodus_nope', args: {} }],
  }];
  assert.ok(validateScript(bad).some((p) => /unknown tool/.test(p)));

  const badAction = [{
    id: 'x', title: 'x', counsel: COUNSEL.customer,
    steps: [{ tool: 'docxodus_edit', args: { action: 'teleport_block' } }],
  }];
  assert.ok(validateScript(badAction).some((p) => /has no action teleport_block/.test(p)));
});

test('validateScript rejects a reference used before it is bound', () => {
  const bad = [{
    id: 'x', title: 'x', counsel: COUNSEL.customer,
    steps: [
      { tool: 'docxodus_edit', args: { action: 'replace_text', anchorId: '@later', markdown: 'x' } },
      { tool: 'docxodus_search', args: { mode: 'text', query: 'y' }, bind: 'later' },
    ],
  }];
  assert.ok(validateScript(bad).some((p) => /@later is used before it is bound/.test(p)));
});

test('every scripted step carries the human note the stage shows', () => {
  for (const act of SCRIPT) {
    for (const step of act.steps) {
      assert.ok(step.note && step.note.length > 3,
        `${act.id}: a step with no note leaves the stage caption blank`);
    }
  }
});

test('the script exercises more than one tool per act', () => {
  // A "negotiation" that only ever calls docxodus_edit would not demonstrate the
  // grouped-intent surface, which is the thing worth showing.
  for (const act of SCRIPT) {
    const tools = new Set(act.steps.map((s) => s.tool));
    assert.ok(tools.size >= 3, `${act.id} uses only ${[...tools].join(', ')}`);
  }
});

test('baseline section headings avoid markdown ordered-list syntax', () => {
  // "## 8. Governing Law" loses its "8." to the markdown list parser once the
  // "## " is stripped, so the heading text is not what the script searches for.
  for (const line of BASELINE) {
    assert.ok(!/^#{1,3} \d+[.)]\s/.test(line),
      `"${line}" would be parsed as an ordered list item, not a heading`);
  }
});

test('the baseline plants the defined-term defect Act III fixes', () => {
  const text = BASELINE.join('\n');
  assert.ok(text.includes('"Supplier"'), 'Supplier must be the DEFINED term');
  const vendorHits = text.match(/The Vendor/g) ?? [];
  assert.equal(vendorHits.length, 2, 'Act III conforms exactly two "Vendor" occurrences');
  for (const act of SCRIPT) {
    for (const step of act.steps) {
      if (step.args?.find === 'The Vendor') {
        assert.equal(step.args.replace, 'Supplier');
      }
    }
  }
});

test('exactly one step demonstrates the engine refusing an irreversible edit', () => {
  const refusals = SCRIPT.flatMap((act) => act.steps.filter((s) => s.expectRefusal));
  assert.equal(refusals.length, 1);
  assert.equal(refusals[0].tool, 'docxodus_list');
  assert.ok(refusals[0].refusalNote, 'a refusal beat must explain itself on the stage');
});

test('the script still footnotes under recording, which is what #614 regressed on', () => {
  // The inverse of an earlier guard, and deliberately so. insert_footnote under
  // render_inline used to leave a definition reject-all could not remove, so the
  // step was taken out of the script; #625 made the citation the reversible unit
  // and the step came back. Keeping it here means the demo's reversibility proof
  // runs over a footnote on every load — remove the step and the fix loses the
  // one end-to-end check that exercises it through the wire.
  const notes = SCRIPT.flatMap((act) => act.steps.filter(
    (s) => s.args?.action === 'insert_footnote' || s.args?.action === 'insert_endnote'));
  assert.equal(notes.length, 1, 'exactly one note-insertion beat');
  assert.equal(notes[0].args.action, 'insert_footnote');
});

test('the footnote cites a live offset rather than a hardcoded one', () => {
  // The cap paragraph is rewritten before the footnote lands, so the span from
  // the search that bound @cap2 is stale by then. The script must re-search and
  // bind a fresh offset; a literal number here would drift the moment any
  // earlier act changes the sentence.
  const act = SCRIPT.find((a) => a.id === 'act-2');
  const note = act.steps.find((s) => s.args?.action === 'insert_footnote');
  assert.equal(typeof note.args.characterOffset, 'string',
    'characterOffset must be a @reference, not a literal');
  assert.ok(note.args.characterOffset.startsWith('@'));
  const source = act.steps.find((s) => s.bindOffset === note.args.characterOffset.slice(1));
  assert.ok(source, 'the offset reference must be bound by an earlier search step');
  assert.equal(source.tool, 'docxodus_search');
  assert.ok(act.steps.indexOf(source) < act.steps.indexOf(note), 'bound before it is used');
});

test('offsetAfterMatch measures past the match, and refuses to guess', () => {
  assert.equal(offsetAfterMatch({ matches: [{ span: { start: 12, length: 25 } }] }), 37);
  // A zero-length match at the start is still an offset, not an absence — the
  // falsy-vs-null distinction matters because 0 is a legal characterOffset.
  assert.equal(offsetAfterMatch({ matches: [{ span: { start: 0, length: 0 } }] }), 0);
  for (const empty of [null, undefined, {}, { matches: [] }, { matches: [{}] },
    { anchors: [{ id: 'p:body:x' }] }]) {
    assert.equal(offsetAfterMatch(empty), null);
  }
});

test('validateScript accepts an offset bound by bindOffset', () => {
  const search = {
    tool: 'docxodus_search', args: { mode: 'text', query: 'x' },
    bind: 'p', bindOffset: 'at',
  };
  const cite = {
    tool: 'docxodus_create',
    args: { action: 'insert_footnote', anchorId: '@p', characterOffset: '@at', markdown: 'n' },
  };
  const act = { id: 'a', title: 'T', counsel: { name: 'C' } };
  assert.deepEqual(validateScript([{ ...act, steps: [search, cite] }]), []);
  // ...and still catches the offset being used before anything binds it.
  const problems = validateScript([{ ...act, steps: [{ ...search, bindOffset: undefined }, cite] }]);
  assert.equal(problems.length, 1);
  assert.match(problems[0], /@at is used before it is bound/);
});

test('every act names a distinct counsel, so attribution is visible', () => {
  const names = SCRIPT.map((a) => a.counsel.name);
  assert.equal(new Set(names).size, names.length);
});

test('speed presets are ordered fastest-last and MAX removes pacing', () => {
  const delays = SPEEDS.map((s) => s.delayMs);
  assert.deepEqual([...delays].sort((a, b) => b - a), delays, 'presets get faster');
  assert.equal(SPEEDS[SPEEDS.length - 1].delayMs, 0, 'MAX is engine-bound');
});

// ─── Attribution and proof phrasing ───────────────────────────────────

test('attributionRollup counts revisions per counsel, dropping the silent ones', () => {
  const rows = attributionRollup([
    { author: COUNSEL.customer.name }, { author: COUNSEL.customer.name },
    { author: COUNSEL.compliance.name },
  ]);
  assert.equal(rows.length, 2);
  assert.equal(rows[0].name, COUNSEL.customer.name);
  assert.equal(rows[0].count, 2);
  assert.equal(rows[1].name, COUNSEL.compliance.name);
});

test('attributionRollup surfaces an author who is not one of the counsel', () => {
  const rows = attributionRollup([{ author: 'Someone Else' }]);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].name, 'Someone Else');
  assert.equal(rows[0].role, 'other');
});

const PASSING_PROOF = {
  acceptToFinal: { equivalent: true },
  rejectToBaseline: { equivalent: false, divergences: [{ partUri: '/word/comments.xml' }] },
  findings: [],
};

test('proofVerdict passes on content reversibility with comment-explained parts', () => {
  const verdict = proofVerdict(PASSING_PROOF, [], 3);
  assert.equal(verdict.ok, true);
  assert.equal(verdict.label, 'REVERSIBLE');
  assert.match(verdict.detail, /zero content differences/);
  assert.match(verdict.detail, /3 reviewer comments survive/);
});

test('proofVerdict fails when any content difference survives reject-all', () => {
  const verdict = proofVerdict(PASSING_PROOF, [{ revisionType: 'Inserted' }], 3);
  assert.equal(verdict.ok, false);
  assert.match(verdict.detail, /1 content difference remain/);
});

test('proofVerdict fails on a divergence no comment explains', () => {
  // This is the shape the footnote gap produced: reject-all left
  // /word/footnotes.xml behind, which no review comment accounts for.
  const verdict = proofVerdict({
    acceptToFinal: { equivalent: true },
    rejectToBaseline: { divergences: [{ partUri: '/word/footnotes.xml' }] },
    findings: [],
  }, [], 3);
  assert.equal(verdict.ok, false);
  assert.match(verdict.detail, /unexplained package divergence/);
  assert.match(verdict.detail, /footnotes\.xml/);
});

test('proofVerdict fails when accept-all does not reproduce the final', () => {
  const verdict = proofVerdict(
    { acceptToFinal: { equivalent: false }, rejectToBaseline: { divergences: [] }, findings: [] },
    [], 0);
  assert.equal(verdict.ok, false);
  assert.match(verdict.detail, /accept-all did not reproduce/);
});

test('proofVerdict distinguishes not-run from ran-and-failed', () => {
  assert.equal(proofVerdict(null).label, 'not run');
});

test('comment parts are only excusable when the document has comments', () => {
  const divergences = [{ partUri: '/word/comments.xml' }, { partUri: '/word/footnotes.xml' }];
  const withComments = classifyDivergences(divergences, 2);
  assert.equal(withComments.explained.length, 1);
  assert.equal(withComments.unexplained.length, 1);
  // With no comments in the document, nothing is excusable.
  const without = classifyDivergences(divergences, 0);
  assert.equal(without.explained.length, 0);
  assert.equal(without.unexplained.length, 2);
});

test('the comment-attributable part set is a closed list, not a pattern', () => {
  // A pattern like /comments|_rels/ would also excuse /word/footnotes.xml on a
  // package that happened to name it oddly. The set is exact on purpose.
  assert.ok(COMMENT_ATTRIBUTABLE_PARTS.has('/word/comments.xml'));
  assert.ok(!COMMENT_ATTRIBUTABLE_PARTS.has('/word/footnotes.xml'));
  assert.ok(!COMMENT_ATTRIBUTABLE_PARTS.has('/word/endnotes.xml'));
});

// ─── The contract check ───────────────────────────────────────────────
//
// The demo's claim is that its frames are the frames docxodus-mcp accepts. That
// is only worth making if it is checked against the shipped catalog rather than
// against a copy of it, so this reads the real file.

const CATALOG_PATH = fileURLToPath(
  new URL('../../../tools/mcp-server/ToolCatalog.cs', import.meta.url));

/** Extract, per tool, the `action` enum and the declared property names from
 *  ToolCatalog.cs. The file is C# raw string literals containing JSON Schema, so
 *  the tool blocks are split on the ToolDefinition constructor and each block's
 *  schema is read with targeted patterns rather than by parsing C#. */
function parseCatalog(source) {
  const tools = new Map();
  // Each entry starts `new ToolDefinition(\n  "name",` — split on that.
  const blocks = source.split(/new ToolDefinition\(/).slice(1);
  for (const block of blocks) {
    const nameMatch = /^\s*"([a-z_]+)"/.exec(block);
    if (!nameMatch) continue;
    const name = nameMatch[1];
    // The action enum, when the tool has one.
    const actionMatch = /"action"\s*:\s*\{[^}]*"enum"\s*:\s*\[([^\]]*)\]/s.exec(block);
    const actions = actionMatch
      ? [...actionMatch[1].matchAll(/"([a-z_]+)"/g)].map((m) => m[1])
      : null;
    // The mode enum, for the search tool.
    const modeMatch = /"mode"\s*:\s*\{[^}]*"enum"\s*:\s*\[([^\]]*)\]/s.exec(block);
    const modes = modeMatch
      ? [...modeMatch[1].matchAll(/"([a-z_]+)"/g)].map((m) => m[1])
      : null;
    // Property names declared anywhere in the tool's schema: `"name": {`.
    const properties = new Set(
      [...block.matchAll(/"([A-Za-z][A-Za-z0-9_]*)"\s*:\s*\{/g)].map((m) => m[1]));
    tools.set(name, { actions, modes, properties });
  }
  return tools;
}

const CATALOG = parseCatalog(readFileSync(CATALOG_PATH, 'utf8'));

test('the catalog parser found the tools it is supposed to check against', () => {
  // A parser that silently matched nothing would make every check below vacuous.
  assert.ok(CATALOG.size >= 15, `parsed only ${CATALOG.size} tools from ToolCatalog.cs`);
  assert.ok(CATALOG.has('docxodus_edit'));
  assert.ok(CATALOG.get('docxodus_edit').actions.includes('replace_text'));
  assert.ok(CATALOG.get('docxodus_search').modes.includes('text'));
});

test('every (tool, action) the browser endpoint routes exists in the real catalog', () => {
  for (const { tool, action } of implementedPairs()) {
    const entry = CATALOG.get(tool);
    assert.ok(entry, `${tool} is not a tool in ToolCatalog.cs`);
    if (action === null) continue;
    assert.ok(entry.actions, `${tool} has no action enum in the catalog`);
    assert.ok(entry.actions.includes(action),
      `${tool} has no action "${action}" in ToolCatalog.cs (has: ${entry.actions.join(', ')})`);
  }
});

test('the tracked-change mode names match the catalog set_mode enum', () => {
  // Passing a wire name straight to DocxSession.setTrackedChanges silently fails
  // to switch the mode, so the endpoint maps them — against these exact names.
  const block = readFileSync(CATALOG_PATH, 'utf8');
  const modeEnum = /"mode":\s*\{[^}]*"enum":\s*\[([^\]]*)\][^}]*set_mode/s.exec(block)
    ?? /set_mode[^}]*?"mode":\s*\{[^}]*"enum":\s*\[([^\]]*)\]/s.exec(block);
  const names = modeEnum
    ? [...modeEnum[1].matchAll(/"([a-z_]+)"/g)].map((m) => m[1])
    : ['accept', 'render_inline', 'strip_deletions'];
  for (const name of names) {
    assert.ok(name in TRACKED_CHANGE_MODES,
      `set_mode accepts "${name}" but the endpoint cannot map it`);
  }
});

test('every search mode the endpoint implements exists in the real catalog', () => {
  const catalogModes = CATALOG.get('docxodus_search').modes;
  for (const mode of IMPLEMENTED_TOOLS.docxodus_search.modes) {
    assert.ok(catalogModes.includes(mode), `docxodus_search has no mode "${mode}"`);
  }
});

test('every argument the endpoint declares is a property the catalog declares', () => {
  for (const [tool, spec] of Object.entries(IMPLEMENTED_TOOLS)) {
    const entry = CATALOG.get(tool);
    for (const arg of spec.args) {
      assert.ok(entry.properties.has(arg),
        `${tool} argument "${arg}" is not declared in ToolCatalog.cs`);
    }
  }
});

test('every argument the SCRIPT sends is a property the catalog declares', () => {
  // The strongest form of the claim: not just the endpoint's declared surface,
  // but the exact arguments that go on the wire during the show.
  for (const act of SCRIPT) {
    for (const step of act.steps) {
      const entry = CATALOG.get(step.tool);
      assert.ok(entry, `${step.tool} is not in ToolCatalog.cs`);
      for (const arg of Object.keys(step.args ?? {})) {
        assert.ok(entry.properties.has(arg),
          `${act.id}: ${step.tool} argument "${arg}" is not declared in ToolCatalog.cs`);
      }
      if (step.args?.action) {
        assert.ok(entry.actions?.includes(step.args.action),
          `${act.id}: ${step.tool} action "${step.args.action}" is not in ToolCatalog.cs`);
      }
      if (step.args?.listFormat) {
        // The list vocabulary is an enum too, and (a)(b)(c) is the one the
        // carve-outs rely on.
        assert.ok(/lowerLetterParenthesis/.test(readFileSync(CATALOG_PATH, 'utf8')),
          'listFormat lowerLetterParenthesis is gone from the catalog');
      }
    }
  }
});
