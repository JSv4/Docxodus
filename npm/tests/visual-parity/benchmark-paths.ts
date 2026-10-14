import {
  existsSync,
  lstatSync,
  mkdirSync,
  readdirSync,
  realpathSync,
} from 'node:fs';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';

function isInside(parent: string, candidate: string): boolean {
  const difference = relative(parent, candidate);
  return difference === '' || (!difference.startsWith('..') && !isAbsolute(difference));
}

function nearestExistingAncestor(path: string): string {
  let current = path;
  while (!existsSync(current)) {
    const parent = dirname(current);
    if (parent === current) throw new Error(`No existing ancestor for ${path}`);
    current = parent;
  }
  return current;
}

function assertRegularDirectory(path: string, label: string): void {
  const metadata = lstatSync(path);
  if (!metadata.isDirectory() || metadata.isSymbolicLink()) {
    throw new Error(`${label} must be a non-symlink directory: ${path}`);
  }
}

/** Prepare a canonical artifact directory that cannot resolve into the source repository. */
export function prepareExternalOutputRoot(
  repositoryRoot: string,
  requestedRoot: string,
  retry: number,
  allowedBootstrapArtifacts: ReadonlySet<string>,
): string {
  if (!Number.isSafeInteger(retry) || retry < 0) throw new Error('Retry index must be non-negative.');
  const canonicalRepository = realpathSync(repositoryRoot);
  const configured = resolve(requestedRoot);
  if (isInside(canonicalRepository, configured)) {
    throw new Error(`Generated-PDF parity artifacts must stay outside the repository: ${configured}`);
  }
  const ancestor = realpathSync(nearestExistingAncestor(configured));
  if (isInside(canonicalRepository, ancestor)) {
    throw new Error(`Generated-PDF parity artifact ancestor resolves inside the repository: ${ancestor}`);
  }
  mkdirSync(configured, { recursive: true, mode: 0o700 });
  assertRegularDirectory(configured, 'Configured artifact root');
  const canonicalConfigured = realpathSync(configured);
  if (isInside(canonicalRepository, canonicalConfigured)) {
    throw new Error(`Generated-PDF parity artifact root resolves inside the repository: ${configured}`);
  }

  const output = retry === 0 ? canonicalConfigured : join(canonicalConfigured, `retry-${retry}`);
  mkdirSync(output, { recursive: true, mode: 0o700 });
  assertRegularDirectory(output, 'Attempt artifact root');
  const canonicalOutput = realpathSync(output);
  if (!isInside(canonicalConfigured, canonicalOutput)
    || isInside(canonicalRepository, canonicalOutput)) {
    throw new Error(`Generated-PDF parity retry root escaped its configured root: ${output}`);
  }
  const unexpected: string[] = [];
  for (const name of readdirSync(canonicalOutput)) {
    const path = join(canonicalOutput, name);
    const metadata = lstatSync(path);
    if (!allowedBootstrapArtifacts.has(name)
      || !metadata.isFile() || metadata.isSymbolicLink()) unexpected.push(name);
  }
  if (unexpected.length > 0) {
    throw new Error('Generated-PDF parity output contains stale or unsafe artifacts: '
      + `${unexpected.join(', ')}\nEvidence from a previous run cannot be mixed with this one. `
      + `Remove ${canonicalOutput}, or point DOCXODUS_GENERATED_PDF_PARITY_OUTPUT at a new `
      + 'directory outside the repository.');
  }
  return canonicalOutput;
}

export function assertSafeCaseId(id: string): void {
  if (!/^[a-z0-9](?:[a-z0-9-]{0,62}[a-z0-9])?$/.test(id)) {
    throw new Error(`Generated-PDF parity case has an unsafe artifact identifier: ${id}`);
  }
}

/** Resolve a portable repository-relative path and reject traversal, symlink, and escape cases. */
export function resolveTrackedRegularFile(repositoryRoot: string, repositoryPath: string): string {
  if (!repositoryPath || isAbsolute(repositoryPath) || repositoryPath.includes('\\')
    || repositoryPath.includes('\0')
    || repositoryPath.split('/').some((segment) => segment === '' || segment === '.' || segment === '..')) {
    throw new Error(`Tracked path is not a canonical repository-relative path: ${repositoryPath}`);
  }
  const canonicalRepository = realpathSync(repositoryRoot);
  const lexical = resolve(canonicalRepository, repositoryPath);
  if (!isInside(canonicalRepository, lexical)) {
    throw new Error(`Tracked path escapes the repository: ${repositoryPath}`);
  }
  const metadata = lstatSync(lexical);
  if (!metadata.isFile() || metadata.isSymbolicLink()) {
    throw new Error(`Tracked path must be a regular non-symlink file: ${repositoryPath}`);
  }
  const canonical = realpathSync(lexical);
  if (!isInside(canonicalRepository, canonical)) {
    throw new Error(`Tracked path resolves outside the repository: ${repositoryPath}`);
  }
  return canonical;
}
