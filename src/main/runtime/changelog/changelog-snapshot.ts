import type { ChangelogSnapshot, ChangelogSourceKind } from '../../../types/changelog';
import { parseChangelog } from '../../../core/utils/changelog-parse';
import { compareSemverLike } from '../update/release-assets';

export function buildChangelogSnapshot(
  markdown: string,
  options: {
    installedVersion: string;
    source: ChangelogSourceKind;
    releaseTag?: string;
    warning?: string;
  },
): ChangelogSnapshot {
  const entries = parseChangelog(markdown);
  const latest = entries.reduce<string | null>(
    (best, entry) =>
      best === null || compareSemverLike(entry.version, best) > 0 ? entry.version : best,
    null,
  );
  const latestEntry = entries.find((entry) => entry.version === latest) ?? entries[0] ?? null;

  return {
    entries,
    installedVersion: options.installedVersion,
    latestVersion: latest,
    expandedGroupKey: latestEntry?.groupKey ?? null,
    source: options.source,
    ...(options.releaseTag ? { releaseTag: options.releaseTag } : {}),
    ...(options.warning ? { warning: options.warning } : {}),
  };
}

export function buildEmptyChangelogSnapshot(options: {
  installedVersion: string;
  error: string;
}): ChangelogSnapshot {
  return {
    entries: [],
    installedVersion: options.installedVersion,
    latestVersion: null,
    expandedGroupKey: null,
    source: 'bundled',
    error: options.error,
  };
}
