import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';

const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error('Usage: node summarize.mjs RESULTS_JSON SUMMARY_JSON');
const data = JSON.parse(await readFile(input, 'utf8'));
assert.ok(data.method.warmup < data.method.iterations, 'no warm samples: warmup must be below iterations');
const median = values => {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};
const p95 = values => [...values].sort((a, b) => a - b)[Math.ceil(values.length * 0.95) - 1];
const groups = new Map();
for (const run of data.runs) {
  const key = run.fixture + '|' + run.operation;
  if (!groups.has(key)) groups.set(key, []);
  groups.get(key).push(run);
}
const rows = [];
for (const runs of groups.values()) {
  const row = { fixture: runs[0].fixture, operation: runs[0].operation, arms: {} };
  const hashes = new Set(runs.flatMap(r => r.samples.map(s => s.outputHash)));
  assert.equal(hashes.size, 1, `${row.fixture} ${row.operation}: output changed across calls or profiles`);
  row.outputHash = [...hashes][0];
  row.outputLength = runs[0].samples[0].outputLength;
  row.pages = runs[0].samples[0].pages;
  row.anchors = runs[0].samples[0].anchors;
  for (const arm of data.method.arms) {
    const cases = runs.filter(r => r.arm === arm);
    assert.equal(cases.length, data.method.repetitions, `${arm}: incomplete repetitions`);
    assert.ok(cases.every(r => r.samples.length === data.method.iterations), `${arm}: incomplete samples`);
    const cold = cases.map(r => r.samples[0].wallMs);
    const warm = cases.flatMap(r => r.samples.slice(data.method.warmup));
    const bridgeNames = [...new Set(warm.flatMap(s => Object.keys(s.calls)))];
    row.arms[arm] = {
      firstCallMedianMs: median(cold), firstCallSamplesMs: cold,
      warmSamples: warm.length,
      warmMedianMs: median(warm.map(s => s.wallMs)),
      warmP95Ms: p95(warm.map(s => s.wallMs)),
      warmMedianEngineMs: median(warm.map(s => s.engineMs)),
      warmMedianOtherMs: median(warm.map(s => s.wallMs - s.engineMs)),
      perRepetitionWarmMedianMs: cases.map(r => median(r.samples.slice(data.method.warmup).map(s => s.wallMs))),
      bridgeMedianMs: Object.fromEntries(bridgeNames.map(name => [name, median(warm.map(s => s.calls[name]?.ms ?? 0))])),
    };
  }
  if (row.arms.original && row.arms.expanded) {
    row.warmSpeedup = row.arms.original.warmMedianMs / row.arms.expanded.warmMedianMs;
    row.firstCallSpeedup = row.arms.original.firstCallMedianMs / row.arms.expanded.firstCallMedianMs;
  }
  rows.push(row);
}
const summary = { environment: data.environment, method: data.method, bundles: data.bundles, fixtures: data.fixtures, rows };
await writeFile(output, JSON.stringify(summary, null, 2) + '\n');
for (const row of rows) {
  console.log(`${row.fixture} ${row.operation}: ${Object.entries(row.arms).map(([arm, s]) => `${arm}=${s.warmMedianMs.toFixed(1)} ms`).join(', ')}${row.warmSpeedup ? ` (${row.warmSpeedup.toFixed(2)}x)` : ''}`);
}
console.log('All output hashes match across every measured call and profile.');
