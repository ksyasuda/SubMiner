import electron from 'electron';
import type { Extension, Session } from 'electron';
import { existsSync } from 'node:fs';
import * as path from 'node:path';
import { ensureExtensionCopyAsync } from './yomitan-extension-copy';
import { HACHIDORI_SESSION_PARTITION } from './tokenizer/hachidori-parser-bridge';

export function getHachidoriSession(): Session {
  return electron.session.fromPartition(HACHIDORI_SESSION_PARTITION);
}

export function resolveHachidoriExtensionPath(options: {
  moduleDir: string;
  resourcesPath: string;
  exists?: (candidate: string) => boolean;
}): string {
  const candidates = [
    path.resolve(options.moduleDir, '../../../build/hachidori'),
    path.join(options.resourcesPath, 'hachidori'),
  ];
  const found = candidates.find((candidate) =>
    (options.exists ?? existsSync)(path.join(candidate, 'manifest.json')),
  );
  if (!found) throw new Error('Hachidori is not bundled. Run bun run build:hachidori.');
  return found;
}

/** Separate session keeps settings-only launches from injecting into the other backend's overlay. */
export function createHachidoriExtensionRuntime(userDataPath: string) {
  let extension: Extension | null = null;
  let loading: Promise<Extension> | null = null;
  return {
    getSession: getHachidoriSession,
    ensureLoaded(): Promise<Extension> {
      if (extension) return Promise.resolve(extension);
      if (loading) return loading;
      loading = (async () => {
        const source = resolveHachidoriExtensionPath({
          moduleDir: __dirname,
          resourcesPath: process.resourcesPath,
        });
        const copy = await ensureExtensionCopyAsync(source, userDataPath, {
          extensionName: 'hachidori',
        });
        const session = getHachidoriSession();
        // Electron can reuse an old extension worker after its files change.
        // Drop worker registrations before loading, preserving dictionaries and settings.
        await session.clearStorageData({ storages: ['serviceworkers'] });
        extension = await session.extensions.loadExtension(copy.targetDir, {
          allowFileAccess: true,
        });
        return extension;
      })().finally(() => {
        loading = null;
      });
      return loading;
    },
  };
}
