import { createHash } from 'node:crypto';
import {
  lstatSync,
  readFileSync,
  readdirSync,
  realpathSync,
} from 'node:fs';
import { extname, join, relative, resolve, sep } from 'node:path';

const MAXIMUM_GRAPH_FILES = 512;
const MAXIMUM_GRAPH_BYTES = 32 * 1024 * 1024;
const MAXIMUM_EVIDENCE_FILE_BYTES = 64 * 1024 * 1024;

export interface FileDigestEvidence {
  bytes: number;
  sha256: string;
}

export interface JavaScriptGraphEvidence extends FileDigestEvidence {
  files: number;
}

export interface GeneratedPdfBuildEvidence {
  schemaVersion: 1;
  exporterJavaScript: JavaScriptGraphEvidence;
  exporterEntry: FileDigestEvidence;
  /** tsc output. Recorded because a stale one means a partial build, but NOT what runs. */
  docxodusRuntimeEntry: FileDigestEvidence;
  /**
   * The esbuild bundle the page actually loads — `assets.ts` serves
   * `./export-browser.bundle.js`, not the tsc output beside it. Fingerprinting only the latter
   * let a rebuild that changed the bundle pass as unchanged, which is the exact partial-build
   * class this evidence exists to catch.
   */
  docxodusBrowserBundle: FileDigestEvidence;
  exportAssetManifest: FileDigestEvidence;
  npmLock: FileDigestEvidence;
  exporterLock: FileDigestEvidence;
}

export function assertBuildOwningLifecycle(active: boolean, lifecycleEvent: string | undefined): void {
  if (active && lifecycleEvent !== 'test:generated-pdf-parity') {
    throw new Error('Run generated-PDF parity through `npm run test:generated-pdf-parity`; '
      + 'direct Playwright invocation does not own the required npm and exporter builds.');
  }
}

function digest(bytes: Uint8Array | string): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function regularFileEvidence(path: string): FileDigestEvidence {
  const metadata = lstatSync(path);
  if (!metadata.isFile() || metadata.isSymbolicLink()) {
    throw new Error(`Build evidence requires a regular non-symlink file: ${path}`);
  }
  if (metadata.size > MAXIMUM_EVIDENCE_FILE_BYTES) {
    throw new Error(`Build-evidence file exceeds ${MAXIMUM_EVIDENCE_FILE_BYTES} bytes: ${path}`);
  }
  const bytes = readFileSync(path);
  return { bytes: bytes.byteLength, sha256: digest(bytes) };
}

/** Digest every emitted JavaScript module, including siblings imported by dist/index.js. */
export function javascriptGraphEvidence(root: string): JavaScriptGraphEvidence {
  const lexicalRoot = resolve(root);
  const rootMetadata = lstatSync(lexicalRoot);
  if (!rootMetadata.isDirectory() || rootMetadata.isSymbolicLink()) {
    throw new Error(`Exporter build graph root must be a non-symlink directory: ${lexicalRoot}`);
  }
  const canonicalRoot = realpathSync(lexicalRoot);
  const files: string[] = [];
  const visit = (directory: string): void => {
    for (const name of readdirSync(directory).sort()) {
      const path = join(directory, name);
      const metadata = lstatSync(path);
      if (metadata.isSymbolicLink()) throw new Error(`Build graph rejects symlinks: ${path}`);
      if (metadata.isDirectory()) {
        visit(path);
      } else if (metadata.isFile() && extname(name) === '.js') {
        files.push(path);
        if (files.length > MAXIMUM_GRAPH_FILES) {
          throw new Error(`Build graph exceeds ${MAXIMUM_GRAPH_FILES} JavaScript files.`);
        }
      } else if (!metadata.isFile()) {
        throw new Error(`Build graph rejects special filesystem entries: ${path}`);
      }
    }
  };
  visit(canonicalRoot);
  if (files.length === 0) throw new Error(`Build graph contains no JavaScript: ${canonicalRoot}`);
  const graph = createHash('sha256');
  let bytes = 0;
  for (const path of files) {
    const content = readFileSync(path);
    bytes += content.byteLength;
    if (bytes > MAXIMUM_GRAPH_BYTES) {
      throw new Error(`Build graph exceeds ${MAXIMUM_GRAPH_BYTES} JavaScript bytes.`);
    }
    const name = relative(canonicalRoot, path).split(sep).join('/');
    graph.update(name).update('\0').update(String(content.byteLength)).update('\0')
      .update(digest(content)).update('\n');
  }
  return { files: files.length, bytes, sha256: graph.digest('hex') };
}

export function captureGeneratedPdfBuildEvidence(repositoryRoot: string): GeneratedPdfBuildEvidence {
  const exporterDist = resolve(repositoryRoot, 'npm-export/dist');
  return {
    schemaVersion: 1,
    exporterJavaScript: javascriptGraphEvidence(exporterDist),
    exporterEntry: regularFileEvidence(join(exporterDist, 'index.js')),
    docxodusRuntimeEntry: regularFileEvidence(resolve(repositoryRoot, 'npm/dist/export-browser.js')),
    docxodusBrowserBundle: regularFileEvidence(
      resolve(repositoryRoot, 'npm/dist/export-browser.bundle.js')),
    exportAssetManifest: regularFileEvidence(resolve(repositoryRoot, 'npm/dist/export-assets.json')),
    npmLock: regularFileEvidence(resolve(repositoryRoot, 'npm/package-lock.json')),
    exporterLock: regularFileEvidence(resolve(repositoryRoot, 'npm-export/package-lock.json')),
  };
}
