import assert from 'node:assert/strict';
import { test } from 'node:test';
import { observeFonts } from '../dist/font-observation.js';

function protocol({ failure, nodes = 1, faces = 1, limit = 10 } = {}) {
  const calls = [];
  const session = {
    async send(method) {
      calls.push(method);
      if (failure?.method === method) throw failure.error;
      if (method === 'Runtime.evaluate') return { result: { objectId: 'nodes' } };
      if (method === 'Runtime.getProperties') return { result: Array.from({ length: nodes }, (_, i) =>
        ({ name: String(i), value: { objectId: `text-${i}` } })) };
      if (method === 'DOM.requestNode') return { nodeId: 1 };
      if (method === 'CSS.getPlatformFontsForNode') return { fonts: Array.from({ length: faces }, () =>
        ({ familyName: 'Actual Face', postScriptName: '', isCustomFont: false, glyphCount: 1 })) };
      if (method === 'Runtime.callFunctionOn') return { result: { value: { text: 'A', requestId: 'font-0001' } } };
      return {};
    },
    async detach() { calls.push('detach'); },
  };
  const context = { newCDPSession: async () => session };
  const report = { fonts: [{ requestId: 'font-0001' }], resources: [], warnings: [], unsupportedContent: [],
    options: { policy: { limits: { renderDiagnostics: limit } } } };
  return { calls, run: () => observeFonts(context, {}, report, 'print') };
}

test('only a missing font-observation protocol is reported as unavailable', async () => {
  const p = protocol({ failure: { method: 'CSS.getPlatformFontsForNode',
    error: new Error("Protocol error (CSS.getPlatformFontsForNode): 'CSS.getPlatformFontsForNode' wasn't found") } });
  assert.deepEqual(await p.run(), { status: 'unavailable', reason: 'protocol_unavailable' });
  assert.deepEqual(p.calls.slice(-2), ['Runtime.releaseObjectGroup', 'detach']);
});

for (const message of ['Target closed', 'Timeout exceeded', 'Operation cancelled']) {
  test(`observation propagates ${message} and releases protocol handles`, async () => {
    const error = new Error(message);
    const p = protocol({ failure: { method: 'CSS.getPlatformFontsForNode', error } });
    await assert.rejects(p.run(), received => received === error);
    assert.deepEqual(p.calls.slice(-2), ['Runtime.releaseObjectGroup', 'detach']);
  });
}

for (const setting of [{ limit: 1 }, { nodes: 3, limit: 3 }, { nodes: 2, faces: 2, limit: 6 }]) {
  test(`exhausted observation budget ${JSON.stringify(setting)} preserves export without partial evidence`, async () => {
    const p = protocol(setting);
    assert.deepEqual(await p.run(), { status: 'unavailable', reason: 'diagnostic_limit' });
    assert.deepEqual(p.calls.slice(-2), ['Runtime.releaseObjectGroup', 'detach']);
  });
}

test('an exactly fitting observation retains complete evidence', async () => {
  const p = protocol({ nodes: 2, limit: 5 });
  const result = await p.run();
  assert.equal(result.status, 'observed');
  assert.equal(result.uses.length, 2);
  assert.deepEqual(p.calls.slice(-2), ['Runtime.releaseObjectGroup', 'detach']);
});
