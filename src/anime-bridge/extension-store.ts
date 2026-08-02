import { createReadStream } from 'node:fs';
import { readdir, readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { pipeline } from 'node:stream/promises';
import path from 'node:path';
import type { AnimeBridgeClient } from './bridge-client';
import type { BridgeSource } from './bridge-client';
import type { ExtensionLoadFailure, InstalledExtensionView } from '../types/anime-browser';

/**
 * Anime extensions are Aniyomi APKs the user supplies. They are read from a
 * directory rather than fetched from a hardcoded catalogue, so which sources
 * exist is entirely the user's choice.
 */

export interface InstalledExtension {
  /** Absolute path to the .apk. */
  file: string;
  /** File name without extension, used when the bridge reports no name. */
  fallbackName: string;
  /**
   * SHA-256 of the APK. Identifies the build rather than the slot, so the
   * bridge's extension-id cache misses after an in-place upgrade.
   */
  sha256: string;
}

export interface ExtensionSource {
  /** Package-qualified id used by the UI and runtime. */
  id: string;
  /** Raw bridge id, which selects this source inside a factory APK. */
  bridgeId: string;
  name: string;
  lang: string;
  pkg: string;
  file: string;
}

/**
 * Discover every .apk in `directory`. A missing directory yields no extensions.
 *
 * Only a hash is kept, never the bytes: APKs run to several MB each and a
 * base64 copy adds a third on top, so holding the whole set for the lifetime of
 * the Anime Browser would cost far more than re-reading a file on the rare
 * upload. Hashing streams, so peak memory stays flat regardless of APK size.
 */
export async function readInstalledExtensions(directory: string): Promise<InstalledExtension[]> {
  let entries;
  try {
    entries = await readdir(directory, { withFileTypes: true });
  } catch {
    return [];
  }

  const extensions: InstalledExtension[] = [];
  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    if (!entry.isFile() || !entry.name.toLowerCase().endsWith('.apk')) continue;
    const file = path.join(directory, entry.name);
    extensions.push({
      file,
      fallbackName: entry.name.replace(/\.apk$/i, ''),
      sha256: await hashFile(file),
    });
  }
  return extensions;
}

async function hashFile(file: string): Promise<string> {
  const hash = createHash('sha256');
  await pipeline(createReadStream(file), hash);
  return hash.digest('hex');
}

/**
 * Describe what is on disk, for the installed list in the Extensions tab.
 *
 * Built from the directory rather than from a repository catalogue: an APK
 * dropped in by hand, or one whose repository the user has since removed, is
 * still installed and must stay removable.
 */
export function toInstalledExtensionViews(
  extensions: InstalledExtension[],
  sources: ExtensionSource[],
  loadFailures: ExtensionLoadFailure[],
): InstalledExtensionView[] {
  return extensions.map((extension) => {
    const provided = sources.filter((source) => source.file === extension.file);
    const names = [...new Set(provided.map((source) => source.name))];
    return {
      pkg: extension.fallbackName,
      name: names.length > 0 ? names.join(', ') : extension.fallbackName,
      langs: [...new Set(provided.map((source) => source.lang))],
      sourceCount: provided.length,
      error: loadFailures.find((failure) => failure.pkg === extension.fallbackName)?.error ?? null,
    };
  });
}

/**
 * The bridge payload for a specific source inside an extension.
 *
 * The APK is read on demand: after the first upload the bridge answers by
 * extension id, so most calls never touch the file at all.
 */
export function toBridgeSource(extension: InstalledExtension, sourceId?: string): BridgeSource {
  return {
    fingerprint: extension.sha256,
    loadApkBase64: async () => (await readFile(extension.file)).toString('base64'),
    ...(sourceId ? { sourceId } : {}),
  };
}

/**
 * Ask the bridge which sources each extension provides.
 *
 * An extension that fails to load is skipped rather than aborting the scan, so
 * one broken APK cannot hide every working one. Failures are reported through
 * `onError` for surfacing in the UI.
 */
export async function listExtensionSources(
  client: AnimeBridgeClient,
  extensions: InstalledExtension[],
  onError?: (extension: InstalledExtension, error: unknown) => void,
): Promise<ExtensionSource[]> {
  const sources: ExtensionSource[] = [];

  for (const extension of extensions) {
    try {
      const descriptors = await client.listAnimeSources(toBridgeSource(extension));
      for (const descriptor of descriptors) {
        const bridgeId = descriptor.id === undefined ? null : String(descriptor.id);
        if (bridgeId === null || bridgeId.length === 0) continue;
        sources.push({
          id: `${extension.fallbackName}:${bridgeId}`,
          bridgeId,
          name: descriptor.name?.trim() || extension.fallbackName,
          lang: descriptor.lang ?? 'all',
          pkg: extension.fallbackName,
          file: extension.file,
        });
      }
    } catch (error) {
      onError?.(extension, error);
    }
  }

  return sources;
}
