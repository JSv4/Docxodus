// A browser-side MCP endpoint: the docxodus-mcp tool contract, spoken over real
// JSON-RPC 2.0 frames, dispatched against a live in-tab DocxSession.
//
// WHAT THIS IS, EXACTLY. `tools/mcp-server` is a .NET process that speaks
// newline-delimited JSON-RPC 2.0 on stdin/stdout and routes each `tools/call` by
// (tool, action) onto `DocxSessionOps`. A browser cannot subprocess it. But the
// thing an agent actually sends is a JSON frame, and the thing that executes it is
// a facade the WASM build carries too — so this module reimplements the SERVER'S
// FRONT HALF (envelope parsing, the (tool, action) routing table, the MCP result
// shape) over `npm/src/session.ts` instead of over stdio. The frames on the wire
// console of the demo page are therefore the same frames `docxodus-mcp` accepts;
// only the transport under them differs.
//
// That claim is worth something only if it is checked, so it is:
// `docs/demo/tools/mcp-wire.test.mjs` parses the REAL `tools/mcp-server/ToolCatalog.cs`
// and asserts every (tool, action) pair this module implements exists in the
// shipped catalog's action enum, with the arguments the catalog marks required.
// Rename an action in the server and that test fails here.
//
// What this module deliberately does NOT reimplement: the document store and its
// containment checks (there is no filesystem in a tab — sessions are opened by the
// host page, not by `docxodus_open` with a path), the retry/transaction journal,
// the delivery bundle host, and the MCP Apps UI extension. Those are the server's,
// and the demo does not pretend otherwise.
//
// Import-safe under Node: nothing here touches `document` or `window` at module
// scope, and everything above the endpoint factory is pure.

// ─── Protocol constants ───────────────────────────────────────────────

export const JSONRPC_VERSION = '2.0';

/** The baseline revision `docxodus-mcp` reports when a client names none
 *  (`Program.cs`'s `ProtocolVersion`). */
export const MCP_PROTOCOL_VERSION = '2024-11-05';

export const SERVER_INFO = { name: 'docxodus-mcp', version: '1.0.0' };

/** Transport/protocol failures only. A tool that fails for a business reason —
 *  bad anchor, unknown action, a refused mutation — is reported as a RESULT with
 *  `isError: true`, per the MCP convention the server follows (`Program.cs`). */
export const JSON_RPC_ERRORS = {
  parseError: -32700,
  invalidRequest: -32600,
  methodNotFound: -32601,
  invalidParams: -32602,
  internalError: -32603,
};

// ─── Frame construction (pure) ────────────────────────────────────────

/** A `tools/call` request frame — what an agent puts on the wire. */
export function requestFrame(id, tool, args) {
  return {
    jsonrpc: JSONRPC_VERSION,
    id,
    method: 'tools/call',
    params: { name: tool, arguments: args ?? {} },
  };
}

/** An MCP tool result frame. `resultJson` is the tool's own JSON, carried as text
 *  content exactly as the server carries it. */
export function toolResultFrame(id, resultJson, isError) {
  return {
    jsonrpc: JSONRPC_VERSION,
    id,
    result: {
      content: [{ type: 'text', text: resultJson }],
      isError: Boolean(isError),
    },
  };
}

/** A JSON-RPC protocol error frame — reserved for envelope problems. */
export function errorFrame(id, code, message) {
  return { jsonrpc: JSONRPC_VERSION, id: id ?? null, error: { code, message } };
}

/** The server's heuristic: a top-level `{"success": false}` — the shape every
 *  Docxodus `EditResult` serializes to — surfaces as a tool error so the calling
 *  agent notices without parsing the text content first. */
export function looksLikeFailure(resultJson) {
  try {
    const parsed = JSON.parse(resultJson);
    return Boolean(parsed) && typeof parsed === 'object' && !Array.isArray(parsed) &&
      parsed.success === false;
  } catch {
    return false;
  }
}

// ─── The routing table (pure data) ────────────────────────────────────
//
// (tool → actions) this endpoint implements, plus the arguments each action
// needs. The contract test cross-checks both against ToolCatalog.cs: every tool
// name must exist there, every action must be in that tool's `action` enum, and
// every argument named here must be a property the catalog declares.

export const IMPLEMENTED_TOOLS = {
  docxodus_search: {
    actions: null, // discriminated by `mode`, not `action`
    modes: ['text', 'kind'],
    args: ['sessionId', 'mode', 'query', 'caseSensitive', 'maxResults'],
  },
  docxodus_edit: {
    actions: ['replace_text', 'replace_text_range', 'insert_paragraph', 'delete_block',
      'move_block', 'split_paragraph', 'merge_paragraphs', 'undo', 'redo'],
    args: ['sessionId', 'action', 'anchorId', 'position', 'markdown', 'find', 'replace',
      'caseSensitive', 'sourceAnchorId', 'targetAnchorId', 'secondAnchorId', 'characterOffset'],
  },
  docxodus_format: {
    actions: ['apply_format', 'apply_format_by_substring', 'set_paragraph_style',
      'set_paragraph_format'],
    args: ['sessionId', 'action', 'anchorId', 'span', 'substring', 'format', 'styleId',
      'paragraphFormat'],
  },
  docxodus_create: {
    actions: ['insert_paragraph', 'insert_heading', 'insert_table', 'insert_footnote',
      'insert_horizontal_rule'],
    args: ['sessionId', 'action', 'anchorId', 'position', 'text', 'level', 'markdown',
      'rows', 'columns', 'cellContents', 'characterOffset', 'ruleStyle'],
  },
  docxodus_list: {
    actions: ['apply_format', 'apply_format_range'],
    args: ['sessionId', 'action', 'anchorId', 'listFormat', 'firstAnchorId', 'lastAnchorId'],
  },
  docxodus_comment: {
    actions: ['add', 'list'],
    args: ['sessionId', 'action', 'anchorId', 'span', 'author', 'markdown', 'initials'],
  },
  docxodus_track_changes: {
    actions: ['list', 'accept', 'reject', 'accept_all', 'reject_all', 'set_mode'],
    args: ['sessionId', 'action', 'revisionId', 'mode'],
  },
  docxodus_get_content: {
    actions: null, // discriminated by `format`
    args: ['sessionId', 'format'],
  },
  docxodus_mutations: {
    actions: null, // a batch of steps, each naming its own tool + action
    args: ['sessionId', 'steps', 'mode'],
  },
};

/** Every (tool, action) pair this endpoint routes, flattened — the contract
 *  test's input. Tools discriminated by something other than `action` report a
 *  null action and are checked for tool existence only. */
export function implementedPairs() {
  const pairs = [];
  for (const [tool, spec] of Object.entries(IMPLEMENTED_TOOLS)) {
    if (spec.actions === null) pairs.push({ tool, action: null });
    else for (const action of spec.actions) pairs.push({ tool, action });
  }
  return pairs;
}

// ─── Wire-console formatting (pure) ───────────────────────────────────

/** Render a frame's arguments as the one-line summary the wire console shows.
 *  Long strings are elided in the MIDDLE: the head of a markdown payload and its
 *  tail are both more informative than twice as much of its head. */
export function summarizeArgs(args, maxLength = 88) {
  if (!args || typeof args !== 'object') return '';
  const parts = [];
  for (const [key, value] of Object.entries(args)) {
    if (key === 'sessionId' || value === undefined) continue;
    parts.push(`${key}: ${summarizeValue(value)}`);
  }
  const line = parts.join(', ');
  return line.length > maxLength ? line.slice(0, maxLength - 1) + '…' : line;
}

function summarizeValue(value) {
  if (typeof value === 'string') {
    return value.length > 34 ? `"${value.slice(0, 20)}…${value.slice(-8)}"` : `"${value}"`;
  }
  if (Array.isArray(value)) return `[${value.length}]`;
  if (value && typeof value === 'object') return `{${Object.keys(value).join(',')}}`;
  return String(value);
}

/** The method label the console prints for a frame: `docxodus_edit/replace_text`,
 *  or just the tool when it has no `action` discriminator. */
export function frameLabel(frame) {
  const name = frame?.params?.name ?? '?';
  const args = frame?.params?.arguments ?? {};
  const discriminator = args.action ?? args.mode ?? args.format;
  return discriminator ? `${name}/${discriminator}` : name;
}

// ─── Latency statistics (pure) ────────────────────────────────────────

/** Rolling call statistics. Keeps every sample (a run is bounded by the script,
 *  not by time) so the percentiles are exact rather than estimated. */
export function createLatencyStats() {
  const samples = [];
  const byTool = new Map();
  return {
    record(tool, ms) {
      samples.push(ms);
      const list = byTool.get(tool) ?? [];
      list.push(ms);
      byTool.set(tool, list);
    },
    count: () => samples.length,
    total: () => samples.reduce((a, b) => a + b, 0),
    percentile(p) {
      if (samples.length === 0) return 0;
      const sorted = [...samples].sort((a, b) => a - b);
      // Nearest-rank: the smallest sample at or above the p-th position.
      const rank = Math.max(1, Math.ceil((p / 100) * sorted.length));
      return sorted[rank - 1];
    },
    max: () => (samples.length ? Math.max(...samples) : 0),
    perTool() {
      const rows = [];
      for (const [tool, list] of byTool) {
        rows.push({
          tool,
          calls: list.length,
          totalMs: list.reduce((a, b) => a + b, 0),
        });
      }
      return rows.sort((a, b) => b.calls - a.calls);
    },
  };
}

// ─── The endpoint (engine-dependent) ──────────────────────────────────

/** Catalog mode name → `TrackedChangeMode` enum value. These mirror the .NET
 *  enum's ordinals, which the TypeScript `TrackedChangeMode` re-declares; a host
 *  that has the engine module should inject the real enum rather than rely on
 *  these, and `redline-theater.js` does. */
export const TRACKED_CHANGE_MODES = {
  accept: 0,
  render_inline: 1,
  strip_deletions: 2,
};

/** Thrown for a business-level tool failure. Mirrors the server's
 *  `McpToolException`: caught at the dispatch seam and reported as a tool RESULT
 *  with `isError: true`, never as a JSON-RPC error. */
export class McpToolError extends Error {}

const need = (args, key) => {
  const value = args?.[key];
  if (value === undefined || value === null || value === '') {
    throw new McpToolError(`missing required argument: ${key}`);
  }
  return value;
};

/**
 * Build an endpoint over a live `DocxSession`.
 *
 * `sessionId` is the opaque handle the frames carry, so the console shows the
 * same shape a real agent would hold. The host page owns opening and closing the
 * session — `docxodus_open`/`docxodus_save` are the server's filesystem-scoped
 * lifecycle and have no meaning in a tab.
 *
 * `onMutate` fires after any call that reported a successful mutation, so the
 * host can repaint. It is called once per CALL (not once per underlying op), which
 * is the batching contract `DocxEditor.refresh()` documents.
 */
export function createBrowserMcpEndpoint({
  session,
  sessionId = 'sess_browser_01',
  onMutate,
  trackedChangeModes = TRACKED_CHANGE_MODES,
}) {
  if (!session) throw new Error('createBrowserMcpEndpoint requires a session');
  let nextId = 1;
  const stats = createLatencyStats();

  /** Route (tool, action) → DocxSession, returning the tool's own result JSON.
   *  Mirrors `Dispatcher.Call`'s switch, including its "unknown action" message. */
  function dispatch(tool, args) {
    switch (tool) {
      case 'docxodus_search': return search(args);
      case 'docxodus_edit': return edit(args);
      case 'docxodus_format': return format(args);
      case 'docxodus_create': return create(args);
      case 'docxodus_list': return listTool(args);
      case 'docxodus_comment': return comment(args);
      case 'docxodus_track_changes': return trackChanges(args);
      case 'docxodus_get_content': return getContent(args);
      case 'docxodus_mutations': return mutations(args);
      default: throw new McpToolError(`unknown tool: ${tool}`);
    }
  }

  const unknownAction = (tool, action) =>
    new McpToolError(`unknown action for ${tool}: ${action}`);

  function search(args) {
    const mode = need(args, 'mode');
    const query = need(args, 'query');
    if (mode === 'text') {
      const matches = session.grep(escapeRegex(query), {
        caseSensitive: args.caseSensitive ?? false,
        maxResults: args.maxResults,
      });
      return { matches: matches.map((m) => ({ enclosingAnchor: m.enclosingAnchor, span: m.span, text: m.text })) };
    }
    if (mode === 'kind') {
      return { anchors: session.findByKind(query, args.scope ?? 'body') };
    }
    throw new McpToolError(`unsupported search mode in the browser endpoint: ${mode}`);
  }

  function edit(args) {
    const action = need(args, 'action');
    switch (action) {
      case 'replace_text':
        return session.replaceText(need(args, 'anchorId'), need(args, 'markdown'));
      case 'replace_text_range': {
        const results = session.replaceTextRange(
          need(args, 'anchorId'), need(args, 'find'), args.replace ?? '',
          { caseSensitive: args.caseSensitive ?? false },
        );
        // The session returns one EditResult per attempted match; the tool result
        // collapses them the way the server does, so a caller sees one outcome.
        const failed = results.find((r) => !r.success);
        return failed ?? { success: true, results };
      }
      case 'insert_paragraph':
        return session.insertParagraph(
          need(args, 'anchorId'), need(args, 'position'), need(args, 'markdown'));
      case 'delete_block':
        return session.deleteBlock(need(args, 'anchorId'));
      case 'move_block':
        return session.moveBlock(
          need(args, 'sourceAnchorId'), need(args, 'targetAnchorId'), need(args, 'position'));
      case 'split_paragraph':
        return session.splitParagraph(need(args, 'anchorId'), need(args, 'characterOffset'));
      case 'merge_paragraphs':
        return session.mergeParagraphs(need(args, 'anchorId'), need(args, 'secondAnchorId'));
      case 'undo': return { success: session.undo() };
      case 'redo': return { success: session.redo() };
      default: throw unknownAction('docxodus_edit', action);
    }
  }

  function format(args) {
    const action = need(args, 'action');
    const anchorId = need(args, 'anchorId');
    switch (action) {
      case 'apply_format':
        return session.applyFormat(anchorId, args.span ?? null, need(args, 'format'));
      case 'apply_format_by_substring':
        return session.applyFormatBySubstring(
          anchorId, need(args, 'substring'), need(args, 'format'));
      case 'set_paragraph_style':
        return session.setParagraphStyle(anchorId, need(args, 'styleId'));
      case 'set_paragraph_format':
        return session.setParagraphFormat(anchorId, need(args, 'paragraphFormat'));
      default: throw unknownAction('docxodus_format', action);
    }
  }

  function create(args) {
    const action = need(args, 'action');
    switch (action) {
      case 'insert_paragraph':
        return session.insertParagraph(
          need(args, 'anchorId'), need(args, 'position'), need(args, 'markdown'));
      case 'insert_heading': {
        // The server has no heading primitive either: it composes an ATX payload
        // and calls InsertParagraph. Same composition here, same result.
        const level = Number(need(args, 'level'));
        const markdown = `${'#'.repeat(level)} ${need(args, 'text')}`;
        return session.insertParagraph(need(args, 'anchorId'), need(args, 'position'), markdown);
      }
      case 'insert_table':
        return session.insertTable(
          need(args, 'anchorId'), need(args, 'position'),
          Number(need(args, 'rows')), Number(need(args, 'columns')),
          { cellContents: args.cellContents },
        );
      case 'insert_footnote':
        return session.insertFootnote(
          need(args, 'anchorId'), Number(args.characterOffset ?? 0), need(args, 'markdown'));
      case 'insert_horizontal_rule':
        // The catalog takes a `ruleStyle` name; the session takes a border-edge
        // object. The server does the same widening.
        return session.insertHorizontalRule(
          need(args, 'anchorId'), need(args, 'position'),
          args.ruleStyle ? { style: args.ruleStyle } : undefined);
      default: throw unknownAction('docxodus_create', action);
    }
  }

  function listTool(args) {
    const action = need(args, 'action');
    switch (action) {
      case 'apply_format':
        return session.applyListFormat(need(args, 'anchorId'), need(args, 'listFormat'));
      case 'apply_format_range':
        return session.applyListFormatRange(
          need(args, 'firstAnchorId'), need(args, 'lastAnchorId'), need(args, 'listFormat'));
      default: throw unknownAction('docxodus_list', action);
    }
  }

  function comment(args) {
    const action = need(args, 'action');
    switch (action) {
      case 'add':
        return session.addComment(
          need(args, 'anchorId'), args.span ?? null, need(args, 'author'),
          need(args, 'markdown'), { initials: args.initials });
      case 'list': return { comments: session.listComments() };
      default: throw unknownAction('docxodus_comment', action);
    }
  }

  function trackChanges(args) {
    const action = need(args, 'action');
    switch (action) {
      case 'list': return { revisions: session.listRevisions() };
      case 'accept': return session.acceptRevision(need(args, 'revisionId'));
      case 'reject': return session.rejectRevision(need(args, 'revisionId'));
      case 'accept_all': return session.acceptAllRevisions();
      case 'reject_all': return session.rejectAllRevisions();
      case 'set_mode': {
        // The wire speaks the catalog's names (`accept`/`render_inline`/
        // `strip_deletions`); `DocxSession.setTrackedChanges` takes the
        // TrackedChangeMode enum. The server does the same widening in
        // `Dispatcher.Open`. Passing the wire name straight through silently
        // fails to switch the mode, which is why this mapping is explicit.
        const wire = need(args, 'mode');
        if (!(wire in trackedChangeModes)) {
          throw new McpToolError(`unknown tracked-change mode: ${wire}`);
        }
        session.setTrackedChanges(trackedChangeModes[wire]);
        return { success: true, mode: wire };
      }
      default: throw unknownAction('docxodus_track_changes', action);
    }
  }

  function getContent(args) {
    const format_ = args.format ?? 'markdown';
    if (format_ === 'markdown') return { markdown: session.project().markdown };
    if (format_ === 'version') return { version: session.getVersion() };
    throw new McpToolError(`unsupported content format in the browser endpoint: ${format_}`);
  }

  function mutations(args) {
    const steps = need(args, 'steps');
    if (!Array.isArray(steps) || steps.length === 0) {
      throw new McpToolError('steps must be a non-empty array');
    }
    // `mode` atomic (the catalog's default) vs best_effort. The browser endpoint
    // gets its atomicity from the session's own undo history rather than the
    // server's transaction journal: a failed step rewinds the steps already
    // applied in this batch. `preview` and the transaction journal are the
    // server's and are not reimplemented here.
    const mode = args.mode ?? 'atomic';
    if (mode !== 'atomic' && mode !== 'best_effort') {
      throw new McpToolError(`unsupported batch mode in the browser endpoint: ${mode}`);
    }
    const atomic = mode === 'atomic';
    const applied = [];
    for (const step of steps) {
      const result = dispatch(step.tool, { ...step.args, sessionId });
      applied.push({ tool: step.tool, action: step.args?.action ?? null, result });
      if (atomic && result && result.success === false) {
        for (let i = 0; i < applied.length - 1; i++) session.undo();
        return { success: false, failedStep: applied.length - 1, error: result.error, applied: [] };
      }
    }
    return { success: true, applied };
  }

  return {
    /** Handle one JSON-RPC frame, synchronously, exactly as the stdio server
     *  handles one line. Returns the response frame plus the measured duration. */
    handle(frame) {
      const started = performance.now();
      if (frame?.jsonrpc !== JSONRPC_VERSION) {
        return {
          response: errorFrame(frame?.id, JSON_RPC_ERRORS.invalidRequest, 'invalid JSON-RPC version'),
          ms: performance.now() - started,
        };
      }
      if (frame.method === 'initialize') {
        return {
          response: {
            jsonrpc: JSONRPC_VERSION,
            id: frame.id,
            result: {
              protocolVersion: frame.params?.protocolVersion ?? MCP_PROTOCOL_VERSION,
              capabilities: { tools: {}, resources: {} },
              serverInfo: SERVER_INFO,
            },
          },
          ms: performance.now() - started,
        };
      }
      if (frame.method !== 'tools/call') {
        return {
          response: errorFrame(frame.id, JSON_RPC_ERRORS.methodNotFound,
            `method not found: ${frame.method}`),
          ms: performance.now() - started,
        };
      }

      const tool = frame.params?.name;
      const args = { ...(frame.params?.arguments ?? {}), sessionId };
      let resultJson;
      let isError;
      try {
        resultJson = JSON.stringify(dispatch(tool, args));
        isError = looksLikeFailure(resultJson);
      } catch (err) {
        // Every dispatch failure is a TOOL error, not a protocol error.
        resultJson = JSON.stringify({ error: String(err?.message ?? err) });
        isError = true;
      }
      const ms = performance.now() - started;
      stats.record(tool, ms);
      if (!isError && onMutate) onMutate(tool, args);
      return { response: toolResultFrame(frame.id, resultJson, isError), ms };
    },

    /** Convenience for a scripted caller: mint the request frame, handle it, and
     *  hand back both halves plus the parsed tool result. */
    call(tool, args) {
      const request = requestFrame(nextId++, tool, args);
      const { response, ms } = this.handle(request);
      const text = response.result?.content?.[0]?.text ?? '{}';
      return {
        request,
        response,
        ms,
        isError: Boolean(response.result?.isError ?? response.error),
        result: safeParse(text),
      };
    },

    sessionId,
    stats,
  };
}

function safeParse(text) {
  try { return JSON.parse(text); } catch { return { raw: text }; }
}

/** `docxodus_search` mode `text` is a LITERAL substring search; `DocxSession.grep`
 *  takes a regex. Escaping here keeps the tool's documented semantics. */
export function escapeRegex(literal) {
  return String(literal).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
