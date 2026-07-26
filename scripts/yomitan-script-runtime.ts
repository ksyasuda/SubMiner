import fs from 'node:fs';
import path from 'node:path';

import { resolveYomitanExtensionPath as resolveBuiltYomitanExtensionPath } from '../src/core/services/yomitan-extension-paths.js';

// Yomitan bootstrap for CLI scripts that need the app's real tokenizer. Mirrors
// what scripts/get_frequency.ts does inline; new scripts should import this.
export interface YomitanRuntimeState {
  yomitanExt: unknown | null;
  yomitanSession: unknown | null;
  parserWindow: unknown | null;
  parserReadyPromise: Promise<void> | null;
  parserInitPromise: Promise<boolean> | null;
  available: boolean;
  note?: string;
}

export function withTimeout<T>(promise: Promise<T>, timeoutMs: number, label: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`${label} timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    promise
      .then((value) => {
        clearTimeout(timer);
        resolve(value);
      })
      .catch((error) => {
        clearTimeout(timer);
        reject(error);
      });
  });
}

export function destroyParserWindow(window: unknown): void {
  if (!window || typeof window !== 'object') {
    return;
  }
  const candidate = window as { isDestroyed?: () => boolean; destroy?: () => void };
  if (typeof candidate.isDestroyed !== 'function' || typeof candidate.destroy !== 'function') {
    return;
  }
  if (!candidate.isDestroyed()) {
    candidate.destroy();
  }
}

export async function loadElectronModule(): Promise<typeof import('electron') | null> {
  try {
    const electronImport = await import('electron');
    return (electronImport.default ?? electronImport) as typeof import('electron');
  } catch {
    return null;
  }
}

async function createYomitanRuntimeState(
  userDataPath: string,
  extensionPath?: string,
): Promise<YomitanRuntimeState> {
  const state: YomitanRuntimeState = {
    yomitanExt: null,
    yomitanSession: null,
    parserWindow: null,
    parserReadyPromise: null,
    parserInitPromise: null,
    available: false,
  };

  const electronImport = await loadElectronModule();
  if (
    !electronImport ||
    !electronImport.app ||
    typeof electronImport.app.whenReady !== 'function' ||
    !electronImport.session
  ) {
    state.note = electronImport
      ? 'electron runtime not available in this process'
      : 'electron import failed';
    return state;
  }

  try {
    await electronImport.app.whenReady();
    const loadYomitanExtension = (await import('../src/core/services/yomitan-extension-loader.js'))
      .loadYomitanExtension as (options: {
      userDataPath: string;
      extensionPath?: string;
      getYomitanParserWindow: () => unknown;
      setYomitanParserWindow: (window: unknown) => void;
      setYomitanParserReadyPromise: (promise: Promise<void> | null) => void;
      setYomitanParserInitPromise: (promise: Promise<boolean> | null) => void;
      setYomitanExtension: (extension: unknown) => void;
      setYomitanSession: (session: unknown) => void;
    }) => Promise<unknown>;

    const extension = await loadYomitanExtension({
      userDataPath,
      extensionPath,
      getYomitanParserWindow: () => state.parserWindow,
      setYomitanParserWindow: (window) => {
        state.parserWindow = window;
      },
      setYomitanParserReadyPromise: (promise) => {
        state.parserReadyPromise = promise;
      },
      setYomitanParserInitPromise: (promise) => {
        state.parserInitPromise = promise;
      },
      setYomitanExtension: (loaded) => {
        state.yomitanExt = loaded;
      },
      setYomitanSession: (nextSession) => {
        state.yomitanSession = nextSession;
      },
    });

    if (!extension) {
      state.note = 'yomitan extension is not available';
      return state;
    }

    state.yomitanExt = extension;
    state.available = true;
    return state;
  } catch (error) {
    state.note = error instanceof Error ? error.message : 'failed to initialize yomitan extension';
    return state;
  }
}

export async function createYomitanRuntimeStateWithSearch(
  userDataPath: string,
  extensionPath?: string,
): Promise<YomitanRuntimeState> {
  const resolvedExtensionPath = resolveBuiltYomitanExtensionPath({
    explicitPath: extensionPath,
    cwd: process.cwd(),
  });

  if (resolvedExtensionPath) {
    try {
      if (fs.existsSync(path.join(resolvedExtensionPath, 'manifest.json'))) {
        const state = await createYomitanRuntimeState(userDataPath, resolvedExtensionPath);
        if (!state.available && !state.note) {
          state.note = `Failed to load yomitan extension at ${resolvedExtensionPath}`;
        }
        return state;
      }
    } catch {
      // fall through to the unconstrained loader below
    }
  }

  return createYomitanRuntimeState(userDataPath, resolvedExtensionPath ?? undefined);
}
