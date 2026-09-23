import { randomUUID } from 'node:crypto';
import { mkdir, rename, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { extensionFileName, type RepoExtension } from './extension-repo';

/**
 * Downloads extension APKs into the extensions directory.
 *
 * Only URLs that came from a repository index the user configured are ever
 * fetched; nothing here discovers or suggests sources.
 */

export interface InstallExtensionOptions {
  extensionsDir: string;
  extension: RepoExtension;
  fetchImpl?: typeof fetch;
  /** Guards against a mistyped repo serving something enormous. */
  maxBytes?: number;
  /** Cancels a stalled download; without it a hung repo blocks the install. */
  signal?: AbortSignal;
  /** Applied when no `signal` is given, so a download can never hang forever. */
  timeoutMs?: number;
  /** Injectable filesystem boundary for failure-path tests. */
  fileIo?: ExtensionInstallerFileIo;
}

export interface ExtensionInstallerFileIo {
  mkdir: (dir: string) => Promise<unknown>;
  writeFile: (filePath: string, bytes: Uint8Array) => Promise<void>;
  rename: (from: string, to: string) => Promise<void>;
  removeFile: (filePath: string) => Promise<void>;
}

const DEFAULT_FILE_IO: ExtensionInstallerFileIo = {
  mkdir: (dir) => mkdir(dir, { recursive: true }),
  writeFile: (filePath, bytes) => writeFile(filePath, bytes),
  rename,
  removeFile: (filePath) => rm(filePath, { force: true }),
};

/** APKs are a few MB; anything far past that is not an extension. */
const DEFAULT_MAX_BYTES = 64 * 1024 * 1024;

/** Generous enough for a large APK on a slow link, short of hanging forever. */
const DEFAULT_TIMEOUT_MS = 120_000;

const APK_MAGIC = [0x50, 0x4b, 0x03, 0x04]; // "PK\x03\x04" — APKs are zip archives.

export function looksLikeApk(bytes: Uint8Array): boolean {
  return APK_MAGIC.every((byte, index) => bytes[index] === byte);
}

/**
 * Download one extension into `extensionsDir`, replacing any previous version.
 *
 * The file is named after the package so an update overwrites in place rather
 * than leaving two versions for the bridge to load.
 */
export async function installExtension(options: InstallExtensionOptions): Promise<string> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const fileIo = options.fileIo ?? DEFAULT_FILE_IO;
  const maxBytes = options.maxBytes ?? DEFAULT_MAX_BYTES;
  const signal = options.signal ?? AbortSignal.timeout(options.timeoutMs ?? DEFAULT_TIMEOUT_MS);

  const response = await fetchImpl(options.extension.apkUrl, { signal });
  if (!response.ok) {
    throw new Error(`Downloading ${options.extension.name} failed (${response.status}).`);
  }

  const declared = Number(response.headers.get('content-length') ?? '0');
  if (declared > maxBytes) {
    throw new Error(`${options.extension.name} is larger than the ${maxBytes} byte limit.`);
  }

  const bytes = await readBounded(response, maxBytes, options.extension.name);
  if (!looksLikeApk(bytes)) {
    throw new Error(`${options.extension.name} did not download as an APK.`);
  }

  await fileIo.mkdir(options.extensionsDir);
  const target = resolveTarget(options.extensionsDir, options.extension.pkg);
  const staged = `${target}.${randomUUID()}.tmp`;
  try {
    await fileIo.writeFile(staged, bytes);
    await fileIo.rename(staged, target);
  } finally {
    try {
      await fileIo.removeFile(staged);
    } catch {}
  }
  return target;
}

/**
 * Read the body incrementally and stop the moment the limit is passed.
 *
 * Buffering first and measuring afterwards would let a repo that lies about
 * (or omits) `content-length` push an unbounded amount into memory before the
 * check ever runs.
 */
async function readBounded(
  response: Response,
  maxBytes: number,
  name: string,
): Promise<Uint8Array> {
  const reader = response.body?.getReader();
  if (!reader) {
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.byteLength > maxBytes) {
      throw new Error(`${name} is larger than the ${maxBytes} byte limit.`);
    }
    return bytes;
  }

  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      try {
        await reader.cancel();
      } catch {}
      throw new Error(`${name} is larger than the ${maxBytes} byte limit.`);
    }
    chunks.push(value);
  }

  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

/**
 * Defence in depth against a repository index that smuggles path separators
 * into a package name: the write target must stay inside `extensionsDir`.
 */
function resolveTarget(extensionsDir: string, pkg: string): string {
  const root = path.resolve(extensionsDir);
  const target = path.resolve(root, extensionFileName(pkg));
  if (path.dirname(target) !== root) {
    throw new Error(`Refusing to install ${pkg}: the package name is not a valid file name.`);
  }
  return target;
}

/** Delete an installed extension. Missing files are treated as already gone. */
export async function removeExtension(extensionsDir: string, pkg: string): Promise<void> {
  await rm(resolveTarget(extensionsDir, pkg), { force: true });
}
