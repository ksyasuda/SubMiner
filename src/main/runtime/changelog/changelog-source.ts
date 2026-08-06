import type { ChangelogSnapshot } from '../../../types/changelog';
import { buildChangelogSnapshot, buildEmptyChangelogSnapshot } from './changelog-snapshot';

const DEFAULT_OWNER = 'ksyasuda';
const DEFAULT_REPO = 'SubMiner';
const DEFAULT_CACHE_TTL_MS = 10 * 60 * 1000;

export interface ChangelogSourceDeps {
  /** Resolves the release the changelog should be read from, or null when unknown. */
  fetchLatestReleaseTag: () => Promise<string | null>;
  fetchText: (url: string) => Promise<string>;
  /** Reads the CHANGELOG.md shipped with the install; null when unavailable. */
  readBundledChangelog: () => string | null;
  getInstalledVersion: () => string;
  now: () => number;
  logWarn: (message: string) => void;
  owner?: string;
  repo?: string;
  cacheTtlMs?: number;
}

export function buildRawChangelogUrl(ref: string, owner: string, repo: string): string {
  return `https://raw.githubusercontent.com/${owner}/${repo}/${encodeURIComponent(ref)}/CHANGELOG.md`;
}

function summarize(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function createChangelogSource(deps: ChangelogSourceDeps): {
  getSnapshot: (options?: { refresh?: boolean }) => Promise<ChangelogSnapshot>;
} {
  const owner = deps.owner ?? DEFAULT_OWNER;
  const repo = deps.repo ?? DEFAULT_REPO;
  const cacheTtlMs = deps.cacheTtlMs ?? DEFAULT_CACHE_TTL_MS;

  let cached: { snapshot: ChangelogSnapshot; fetchedAt: number } | null = null;
  let inFlight: Promise<ChangelogSnapshot> | null = null;

  /** `reason` is the raw failure; each branch phrases it for what it can show. */
  function fallbackToBundled(reason: string): ChangelogSnapshot {
    const bundled = deps.readBundledChangelog();
    if (bundled === null) {
      // Nothing is on screen, so promising a bundled changelog would be a lie.
      return buildEmptyChangelogSnapshot({
        installedVersion: deps.getInstalledVersion(),
        error: `Changelog unavailable: ${reason}`,
      });
    }
    return buildChangelogSnapshot(bundled, {
      installedVersion: deps.getInstalledVersion(),
      source: 'bundled',
      warning: `Showing the bundled changelog: ${reason}`,
    });
  }

  async function loadSnapshot(): Promise<ChangelogSnapshot> {
    let releaseTag: string | null = null;
    try {
      releaseTag = await deps.fetchLatestReleaseTag();
    } catch (error) {
      deps.logWarn(`Changelog release lookup failed: ${summarize(error)}`);
    }

    // Without a release tag the default branch still gives the newest published
    // changelog, so try it before falling back to the bundled copy.
    const ref = releaseTag ?? 'main';
    try {
      const markdown = await deps.fetchText(buildRawChangelogUrl(ref, owner, repo));
      if (markdown.trim().length === 0) {
        throw new Error('Remote changelog was empty.');
      }
      const snapshot = buildChangelogSnapshot(markdown, {
        installedVersion: deps.getInstalledVersion(),
        source: 'remote',
        ...(releaseTag ? { releaseTag } : {}),
      });
      // A 200 that isn't a changelog (a redirect landing page, a renamed repo)
      // parses to nothing; the bundled copy beats showing an empty modal.
      if (snapshot.entries.length === 0) {
        throw new Error('Remote changelog contained no releases.');
      }
      return snapshot;
    } catch (error) {
      const message = summarize(error);
      deps.logWarn(`Changelog download failed (${ref}): ${message}`);
      return fallbackToBundled(message);
    }
  }

  return {
    async getSnapshot(options?: { refresh?: boolean }): Promise<ChangelogSnapshot> {
      const refresh = options?.refresh === true;
      if (!refresh && cached && deps.now() - cached.fetchedAt < cacheTtlMs) {
        return cached.snapshot;
      }
      if (inFlight) return await inFlight;

      inFlight = loadSnapshot()
        .then((snapshot) => {
          // Only a successful remote read is worth caching; a bundled fallback
          // should retry the network on the next open.
          if (snapshot.source === 'remote') {
            cached = { snapshot, fetchedAt: deps.now() };
          } else {
            cached = null;
          }
          return snapshot;
        })
        .finally(() => {
          inFlight = null;
        });

      return await inFlight;
    },
  };
}
