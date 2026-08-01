import { readdir, readFile } from 'node:fs/promises';
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
  apkBase64: string;
}

export interface ExtensionSource {
  /** Stable id: the bridge source id, which selects it inside a factory APK. */
  id: string;
  name: string;
  lang: string;
  pkg: string;
  file: string;
}

/** Read every .apk in `directory`. A missing directory yields no extensions. */
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
    const bytes = await readFile(file);
    extensions.push({
      file,
      fallbackName: entry.name.replace(/\.apk$/i, ''),
      apkBase64: bytes.toString('base64'),
    });
  }
  return extensions;
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

/** The bridge payload for a specific source inside an extension. */
export function toBridgeSource(extension: InstalledExtension, sourceId?: string): BridgeSource {
  return { apkBase64: extension.apkBase64, ...(sourceId ? { sourceId } : {}) };
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
        const id = descriptor.id === undefined ? null : String(descriptor.id);
        if (id === null || id.length === 0) continue;
        sources.push({
          id,
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
