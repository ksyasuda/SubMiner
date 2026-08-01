import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { AnimeBrowserRuntime } from './anime-browser-runtime';
import type { AnimeBrowserPlayRequest } from '../../types/anime-browser';

export interface AnimeBrowserIpcDeps {
  // Structurally typed so tests can pass a fake without importing Electron.
  ipcMain: {
    handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): unknown;
  };
  runtime: AnimeBrowserRuntime;
}

/**
 * Bridge the renderer to the anime runtime. Page arguments arrive as `unknown`
 * from the renderer, so they are coerced here rather than trusted.
 */
export function registerAnimeBrowserIpcHandlers(deps: AnimeBrowserIpcDeps): void {
  const channels = IPC_CHANNELS.request;
  const { runtime } = deps;
  const handle = (
    channel: string,
    listener: (event: unknown, ...args: unknown[]) => unknown,
  ): void => {
    deps.ipcMain.handle(channel, listener);
  };

  handle(channels.animeBrowserGetSnapshot, () => runtime.getSnapshot());
  handle(channels.animeBrowserEnsureBridge, () => runtime.ensureBridge());
  handle(channels.animeBrowserSelectSource, (_event, sourceId) =>
    runtime.selectSource(String(sourceId)),
  );
  handle(channels.animeBrowserSearch, (_event, query, page) =>
    runtime.search(String(query ?? ''), toPage(page)),
  );
  handle(channels.animeBrowserGetPopular, (_event, page) => runtime.getPopular(toPage(page)));
  handle(channels.animeBrowserGetDetails, (_event, animeUrl, sourceId) =>
    runtime.getDetails(String(animeUrl), toOptionalId(sourceId)),
  );
  handle(channels.animeBrowserGetEpisodes, (_event, animeUrl, sourceId) =>
    runtime.getEpisodes(String(animeUrl), toOptionalId(sourceId)),
  );
  handle(channels.animeBrowserListAvailableExtensions, () => runtime.listAvailableExtensions());
  handle(channels.animeBrowserInstallExtension, (_event, pkg) =>
    runtime.installExtension(String(pkg)),
  );
  handle(channels.animeBrowserRemoveExtension, (_event, pkg) =>
    runtime.removeExtension(String(pkg)),
  );
  handle(channels.animeBrowserRescanExtensions, () => runtime.rescanExtensions());
  handle(channels.animeBrowserAddRepo, (_event, url) => runtime.addRepo(String(url)));
  handle(channels.animeBrowserRemoveRepo, (_event, url) => runtime.removeRepo(String(url)));
  handle(channels.animeBrowserPlayEpisode, (_event, request) =>
    runtime.playEpisode(request as AnimeBrowserPlayRequest),
  );
  handle(channels.animeBrowserGetPreferences, (_event, sourceId) =>
    runtime.getPreferences(String(sourceId)),
  );
  handle(channels.animeBrowserSetPreference, (_event, sourceId, key, value) =>
    runtime.setPreference(String(sourceId), String(key), value as string | string[] | boolean),
  );
}

/**
 * A source id the renderer may omit. Absent means "use the current selection",
 * so an empty value must stay undefined rather than becoming the string "".
 */
function toOptionalId(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

/** Bridge pages are 1-based; anything unusable falls back to the first page. */
function toPage(value: unknown): number {
  const page = Number(value);
  return Number.isFinite(page) && page >= 1 ? Math.floor(page) : 1;
}
