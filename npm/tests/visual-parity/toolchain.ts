import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  accessSync,
  constants,
  lstatSync,
  readFileSync,
  realpathSync,
} from 'node:fs';
import { delimiter } from 'node:path';
import { basename, isAbsolute, resolve } from 'node:path';

const VERSION_TIMEOUT_MS = 15_000;
const VERSION_OUTPUT_BYTES = 1024 * 1024;

/** Resolve one executable once so its invocation, version, and digest cannot observe different PATH entries. */
export function resolveExecutable(command: string, pathValue = process.env.PATH ?? ''): string {
  if (!command || command.includes('\0')) throw new Error('Executable name must be non-empty.');
  const candidates = isAbsolute(command)
    ? [command]
    : pathValue.split(delimiter).map((directory) => resolve(directory || '.', command));
  for (const candidate of candidates) {
    try {
      accessSync(candidate, constants.X_OK);
      const resolved = realpathSync(candidate);
      const metadata = lstatSync(resolved);
      if (metadata.isFile() && !metadata.isSymbolicLink()) return resolved;
    } catch {
      // Keep looking. The final error names the requested command without leaking every PATH entry.
    }
  }
  throw new Error(`Required executable was not found as a regular file: ${command}`);
}

/** A bounded, successful version probe for an already-resolved executable. */
export function commandVersion(executable: string, args: readonly string[]): string {
  if (!isAbsolute(executable)) {
    throw new Error(`Version probes require an absolute executable path: ${executable}`);
  }
  const result = spawnSync(executable, [...args], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: VERSION_TIMEOUT_MS,
    maxBuffer: VERSION_OUTPUT_BYTES,
  });
  if (result.error) throw result.error;
  if (result.signal) throw new Error(`${basename(executable)} version probe ended with ${result.signal}.`);
  if (result.status !== 0) {
    throw new Error(`${basename(executable)} version probe exited with status ${String(result.status)}.`);
  }
  return `${result.stdout ?? ''}\n${result.stderr ?? ''}`.trim();
}

export interface ExecutableEvidence {
  command: string;
  executable: string;
  executableSha256: string;
  version: string;
}

export interface PinnedExecutable {
  path: string;
  evidence: ExecutableEvidence;
}

export function pinExecutable(
  command: string,
  versionArgs: readonly string[],
  pathValue = process.env.PATH ?? '',
): PinnedExecutable {
  const resolvedPath = resolveExecutable(command, pathValue);
  return {
    path: resolvedPath,
    evidence: {
      command,
      executable: basename(resolvedPath),
      executableSha256: createHash('sha256').update(readFileSync(resolvedPath)).digest('hex'),
      version: commandVersion(resolvedPath, versionArgs),
    },
  };
}
