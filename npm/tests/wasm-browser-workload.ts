// Browser paths omitted by the original AOT workload (issue #783). Shared by the
// profile recorder and benchmarks/aot-coverage/run.mjs so they execute identical code.
// This function is serialized into page.evaluate: keep it entirely self-contained.
export type BrowserOperation =
  | 'html.bare' | 'html.headers' | 'html.anchors' | 'html.paginated' | 'html.options'
  | 'annotation.create' | 'editor.open' | 'editor.openAsync' | 'editor.openAsync.flow';

export interface BrowserWorkloadInput {
  bytes: number[];
  operation: BrowserOperation;
  iterations: number;
}

export interface BrowserSample {
  wallMs: number;
  engineMs: number;
  calls: Record<string, { count: number; ms: number }>;
  outputLength: number;
  outputHash: string;
  anchors: number;
  /** Rendered page boxes for editor mounts; null for conversions, which paginate client-side later. */
  pages: number | null;
  progress: number;
  firstWindowMs: number | null;
}

export async function runWasmBrowserWorkload(input: BrowserWorkloadInput): Promise<BrowserSample[]> {
  const D = (window as any).Docxodus;
  const bytes = new Uint8Array(input.bytes);
  const op = input.operation;
  const samples: BrowserSample[] = [];
  let calls: BrowserSample['calls'] = {};
  const wrap = (obj: any, prefix: string) => Object.fromEntries(Object.entries(obj).map(([key, value]) => [
    key,
    typeof value !== 'function' ? value : (...args: any[]) => {
      const start = performance.now();
      try { return value.apply(obj, args); }
      finally {
        const call = (calls[prefix + '.' + key] ??= { count: 0, ms: 0 });
        call.count++;
        call.ms += performance.now() - start;
      }
    },
  ]));
  const exports = {
    ...D,
    DocxSessionBridge: wrap(D.DocxSessionBridge, 'session'),
    DocumentConverter: wrap(D.DocumentConverter, 'converter'),
  };

  for (let i = 0; i < input.iterations; i++) {
    const container = document.createElement('div');
    container.style.width = '1024px';
    if (op.startsWith('editor.')) document.body.appendChild(container);
    let editor: any;
    let output = '';
    let progress = 0;
    let firstWindowMs: number | null = null;
    calls = {};
    const start = performance.now();
    try {
      if (op.startsWith('editor.')) {
        const options = {
          paginated: op !== 'editor.openAsync.flow',
          editable: true,
          headerFooter: true, // createRibbonEditor's default; trains the header/footer region path.
          columnWidth: 'section',
          windowSize: 24,
          onProgress: () => {
            progress++;
            firstWindowMs ??= performance.now() - start;
          },
        };
        editor = op === 'editor.open'
          ? D.DocxEditor.open(container, bytes, exports, options)
          : await D.DocxEditor.openAsync(container, bytes, exports, options);
      } else if (op === 'annotation.create') {
        output = exports.DocumentConverter.CreateExternalAnnotationSet(bytes, 'aot-benchmark');
      } else if (op === 'html.bare') {
        output = exports.DocumentConverter.ConvertDocxToHtml(bytes);
      } else {
        output = exports.DocumentConverter.ConvertDocxToHtmlComplete(
          bytes, 'Document', 'docx-', true, '', -1, 'comment-',
          op === 'html.paginated' || op === 'html.options' ? 1 : 0,
          1, 'page-', false, 0, 'annot-',
          op === 'html.options',
          op === 'html.headers' || op === 'html.options',
          false, true, true, false, '',
          op === 'html.anchors' || op === 'html.options',
        );
      }
      // Stop timing before DOM inspection, output hashing, and closing the session.
      const wallMs = performance.now() - start;
      const engineMs = Object.values(calls).reduce((sum, call) => sum + call.ms, 0);
      let root: ParentNode = container;
      if (op.startsWith('editor.')) {
        output = container.innerHTML;
        if (!container.querySelector('[data-anchor]')) throw new Error('mount produced no blocks');
        if (op !== 'editor.open' && progress === 0) throw new Error('openAsync fell back to a synchronous mount');
      } else if (op === 'annotation.create') {
        const parsed = JSON.parse(output);
        if (parsed.DocumentId !== 'aot-benchmark' || typeof parsed.Content !== 'string' || !parsed.Content.length) {
          throw new Error('annotation export failed: ' + output.slice(0, 300));
        }
        // Creation timestamps are the only intentionally variable export fields.
        delete parsed.CreatedAt;
        delete parsed.UpdatedAt;
        output = JSON.stringify(parsed);
      } else {
        if (!output.trimStart().startsWith('<')) throw new Error('conversion failed: ' + output.slice(0, 300));
        root = new DOMParser().parseFromString(output, 'text/html');
      }
      const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(output));
      samples.push({
        wallMs, engineMs, calls: JSON.parse(JSON.stringify(calls)),
        outputLength: output.length,
        outputHash: Array.from(new Uint8Array(hash), b => b.toString(16).padStart(2, '0')).join(''),
        anchors: root.querySelectorAll('[data-anchor]').length,
        pages: op.startsWith('editor.') ? root.querySelectorAll('.page-box').length : null,
        progress, firstWindowMs,
      });
    } finally {
      editor?.close();
      container.remove();
    }
    // Let background compilation and pending browser work proceed between calls.
    await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
  return samples;
}
