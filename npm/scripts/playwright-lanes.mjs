#!/usr/bin/env node
// `npm test`: run the Playwright suite as its two lanes, one after the other.
//
// The pool lane runs every ordinary spec on several workers at once; the serial lane runs
// the specs that assert a wall-clock, latency or frame-rate bound on one worker, where
// another worker's WASM boot cannot inflate what they measure. playwright.config.ts
// decides which spec belongs to which lane from DOCXODUS_PLAYWRIGHT_LANE; this script
// only sets the lane and forwards every argument (`npm test -- worker.spec.ts`) to both.
// CI runs each lane as its own job with the same variable, so this is also the only
// local command that exercises the same split.
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const LANES = ['pool', 'serial'];

/**
 * The specs that run in the serial lane: each asserts a latency, throughput or frame-rate
 * figure that another worker's WASM boot would inflate. Every other spec is in the pool.
 * playwright.config.ts builds its projects from this list; scripts/playwright-lanes.test.mjs
 * fails when an entry names no spec file or when a spec file ends up in both lanes or neither.
 */
export const serialLaneSpecs = [
  'aot-profile-record.spec.ts', // also writes one profile file at a fixed path
  'demo-arcade-doom.spec.ts', // frames per second
  'editor-latency-bench.spec.ts',
  'editor-move-latency-bench.spec.ts',
  'editor-perf-incremental.spec.ts',
  'incremental-annotations.spec.ts',
  'profiling.spec.ts',
  'wasm-steady-state.spec.ts',
  'worker-prepare.spec.ts',
  'worker-request-deadlines.spec.ts',
  'worker.spec.ts',
];

/**
 * Plan one Playwright invocation per lane. A narrowed run (any argument at all — a spec
 * filter, a project, a grep) usually matches nothing in one of the two lanes, so it
 * tolerates an empty lane; a bare run does not, so a lane that runs nothing fails.
 */
export function planLaneRuns(args, lanes = LANES) {
  const tolerateEmpty = args.length > 0 && !args.includes('--pass-with-no-tests');
  return lanes.map(lane => ({
    lane,
    env: { DOCXODUS_PLAYWRIGHT_LANE: lane },
    args: ['test', ...args, ...(tolerateEmpty ? ['--pass-with-no-tests'] : [])],
  }));
}

function main() {
  const npmDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');
  const cli = createRequire(join(npmDir, 'package.json')).resolve('@playwright/test/cli');
  const failed = [];
  for (const run of planLaneRuns(process.argv.slice(2))) {
    console.log(`\n=== Playwright lane: ${run.lane} ===`);
    const result = spawnSync(process.execPath, [cli, ...run.args], {
      cwd: npmDir, stdio: 'inherit', env: { ...process.env, ...run.env },
    });
    if (result.status !== 0) failed.push(run.lane);
  }
  if (failed.length > 0) {
    console.error(`\nPlaywright lane(s) failed: ${failed.join(', ')}`);
    process.exit(1);
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
