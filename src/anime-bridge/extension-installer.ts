import { mkdir, rm, writeFile } from 'node:fs/promises';
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
}

/** APKs are a few MB; anything far past that is not an extension. */
const DEFAULT_MAX_BYTES = 64 * 1024 * 1024;

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
  const maxBytes = options.maxBytes ?? DEFAULT_MAX_BYTES;

  const response = await fetchImpl(options.extension.apkUrl);
  if (!response.ok) {
    throw new Error(`Downloading ${options.extension.name} failed (${response.status}).`);
  }

  const declared = Number(response.headers.get('content-length') ?? '0');
  if (declared > maxBytes) {
    throw new Error(`${options.extension.name} is larger than the ${maxBytes} byte limit.`);
  }

  const bytes = new Uint8Array(await response.arrayBuffer());
  if (bytes.byteLength > maxBytes) {
    throw new Error(`${options.extension.name} is larger than the ${maxBytes} byte limit.`);
  }
  if (!looksLikeApk(bytes)) {
    throw new Error(`${options.extension.name} did not download as an APK.`);
  }

  await mkdir(options.extensionsDir, { recursive: true });
  const target = path.join(options.extensionsDir, extensionFileName(options.extension.pkg));
  await writeFile(target, bytes);
  return target;
}

/** Delete an installed extension. Missing files are treated as already gone. */
export async function removeExtension(extensionsDir: string, pkg: string): Promise<void> {
  await rm(path.join(extensionsDir, extensionFileName(pkg)), { force: true });
}
