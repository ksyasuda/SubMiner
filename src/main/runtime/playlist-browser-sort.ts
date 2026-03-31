import path from 'node:path';

type ParsedEpisodeKey = {
  season: number | null;
  episode: number;
};

type SortToken = string | number;

export type PlaylistBrowserSortedDirectoryItem = {
  path: string;
  basename: string;
  episodeLabel: string | null;
};

const COLLATOR = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: 'base',
});

function parseEpisodeKey(basename: string): ParsedEpisodeKey | null {
  const name = basename.replace(/\.[^.]+$/, '');
  const seasonEpisode = name.match(/(?:^|[^a-z0-9])s(\d{1,2})\s*e(\d{1,3})(?:$|[^a-z0-9])/i);
  if (seasonEpisode) {
    return {
      season: Number(seasonEpisode[1]),
      episode: Number(seasonEpisode[2]),
    };
  }

  const seasonByX = name.match(/(?:^|[^a-z0-9])(\d{1,2})x(\d{1,3})(?:$|[^a-z0-9])/i);
  if (seasonByX) {
    return {
      season: Number(seasonByX[1]),
      episode: Number(seasonByX[2]),
    };
  }

  const namedEpisode = name.match(
    /(?:^|[^a-z0-9])(?:ep|episode|第)\s*(\d{1,3})(?:\s*(?:話|episode|ep))?(?:$|[^a-z0-9])/i,
  );
  if (namedEpisode) {
    return {
      season: null,
      episode: Number(namedEpisode[1]),
    };
  }

  return null;
}

function buildEpisodeLabel(parsed: ParsedEpisodeKey | null): string | null {
  if (!parsed) return null;
  if (parsed.season !== null) {
    return `S${parsed.season}E${parsed.episode}`;
  }
  return `E${parsed.episode}`;
}

function tokenizeNaturalSort(basename: string): SortToken[] {
  return basename
    .toLowerCase()
    .split(/(\d+)/)
    .filter((token) => token.length > 0)
    .map((token) => (/^\d+$/.test(token) ? Number(token) : token));
}

function compareNaturalTokens(left: SortToken[], right: SortToken[]): number {
  const maxLength = Math.max(left.length, right.length);
  for (let index = 0; index < maxLength; index += 1) {
    const a = left[index];
    const b = right[index];
    if (a === undefined) return -1;
    if (b === undefined) return 1;
    if (typeof a === 'number' && typeof b === 'number') {
      if (a !== b) return a - b;
      continue;
    }
    const comparison = COLLATOR.compare(String(a), String(b));
    if (comparison !== 0) return comparison;
  }
  return 0;
}

export function sortPlaylistBrowserDirectoryItems(
  paths: string[],
): PlaylistBrowserSortedDirectoryItem[] {
  return paths
    .map((pathValue) => {
      const basename = path.basename(pathValue);
      const parsed = parseEpisodeKey(basename);
      return {
        path: pathValue,
        basename,
        parsed,
        episodeLabel: buildEpisodeLabel(parsed),
        naturalTokens: tokenizeNaturalSort(basename),
      };
    })
    .sort((left, right) => {
      if (left.parsed && right.parsed) {
        if (
          left.parsed.season !== null &&
          right.parsed.season !== null &&
          left.parsed.season !== right.parsed.season
        ) {
          return left.parsed.season - right.parsed.season;
        }
        if (left.parsed.episode !== right.parsed.episode) {
          return left.parsed.episode - right.parsed.episode;
        }
      } else if (left.parsed && !right.parsed) {
        return -1;
      } else if (!left.parsed && right.parsed) {
        return 1;
      }

      const naturalComparison = compareNaturalTokens(left.naturalTokens, right.naturalTokens);
      if (naturalComparison !== 0) {
        return naturalComparison;
      }
      return COLLATOR.compare(left.basename, right.basename);
    })
    .map(({ path: itemPath, basename, episodeLabel }) => ({
      path: itemPath,
      basename,
      episodeLabel,
    }));
}
