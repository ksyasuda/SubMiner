import type { ChangelogSnapshot } from '../../../types/changelog';
import type { UpdateChannel } from '../../../types/config';
import { createCurlFetch, createGlobalFetch } from '../update/fetch-adapter';
import { fetchLatestStableRelease, type FetchLike } from '../update/release-assets';
import { readBundledChangelog, resolveBundledChangelogPath } from './bundled-changelog';
import { createChangelogSource } from './changelog-source';

export interface ChangelogRuntimeDeps {
  getInstalledVersion: () => string;
  getUpdateChannel: () => UpdateChannel;
  resourcesPath: string;
  appPath: string;
  dirname: string;
  joinPath: (...parts: string[]) => string;
  fileExists: (path: string) => boolean;
  readFile: (path: string) => string;
  logWarn: (message: string) => void;
  /** Injected in tests; production picks curl on POSIX and global fetch on Windows. */
  createFetch?: () => FetchLike;
}

/**
 * curl enforces its own `--max-time`, but the global-fetch transport has no
 * deadline: without this a stalled connection leaves the modal on "Loading
 * changelog..." with no way back except closing it.
 */
export const CHANGELOG_REQUEST_TIMEOUT_MS = 30_000;

export function withRequestTimeout(fetchImpl: FetchLike, timeoutMs: number): FetchLike {
  return (url, init) => {
    if (typeof AbortSignal?.timeout !== 'function') return fetchImpl(url, init);
    return fetchImpl(url, { ...init, signal: AbortSignal.timeout(timeoutMs) });
  };
}

export function createChangelogRuntime(deps: ChangelogRuntimeDeps): {
  getChangelogSnapshot: (options?: { refresh?: boolean }) => Promise<ChangelogSnapshot>;
} {
  // curl matches the updater's transport choice: Electron's global fetch is
  // unreliable for GitHub on some Linux builds.
  const fetchImpl = withRequestTimeout(
    deps.createFetch?.() ??
      (process.platform === 'win32' ? createGlobalFetch() : createCurlFetch()),
    CHANGELOG_REQUEST_TIMEOUT_MS,
  );

  const source = createChangelogSource({
    fetchLatestReleaseTag: async () => {
      const release = await fetchLatestStableRelease({
        fetch: fetchImpl,
        channel: deps.getUpdateChannel(),
      });
      return release?.tag_name ?? null;
    },
    fetchText: async (url) => {
      const response = await fetchImpl(url, {
        headers: { 'User-Agent': 'SubMiner changelog' },
      });
      if (!response.ok) {
        throw new Error(`Changelog request failed with ${response.status}`);
      }
      return await response.text();
    },
    readBundledChangelog: () =>
      readBundledChangelog({
        resolvePath: () =>
          resolveBundledChangelogPath({
            resourcesPath: deps.resourcesPath,
            appPath: deps.appPath,
            dirname: deps.dirname,
            joinPath: deps.joinPath,
            fileExists: deps.fileExists,
          }),
        readFile: deps.readFile,
        logWarn: deps.logWarn,
      }),
    getInstalledVersion: deps.getInstalledVersion,
    now: () => Date.now(),
    logWarn: deps.logWarn,
  });

  return {
    getChangelogSnapshot: (options?: { refresh?: boolean }) => source.getSnapshot(options),
  };
}
