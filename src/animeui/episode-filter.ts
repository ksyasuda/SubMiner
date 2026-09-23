/**
 * Filtering for the episode list.
 *
 * A season's worth of episodes does not fit on screen, and sources hand them
 * over in whatever order they please, so the list needs a way to jump straight
 * to one. The query is deliberately forgiving: a number finds that episode, a
 * range finds a span of them, and anything else is a substring of the name.
 */

export interface FilterableEpisode {
  /** Effective episode number, after the list's fallback numbering. */
  number: number | null;
  name: string;
}

export type EpisodeFilter =
  | { kind: 'number'; value: number; text: string }
  | { kind: 'range'; from: number; to: number }
  | { kind: 'text'; text: string };

const NUMBER = String.raw`\d{1,4}(?:\.\d+)?`;
const RANGE_PATTERN = new RegExp(`^(${NUMBER})\\s*[-–—~]\\s*(${NUMBER})$`);
const NUMBER_PATTERN = new RegExp(`^(${NUMBER})$`);

function normalize(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

/**
 * Read a query into what it asks for, or null when it asks for nothing.
 *
 * A bare number stays a text match as well as a numeric one: sources put the
 * number in the name often enough ("Episode 12 - …") that a source reporting no
 * numbers at all would otherwise filter to nothing.
 */
export function parseEpisodeFilter(query: string): EpisodeFilter | null {
  const text = normalize(query);
  if (!text) return null;

  const range = text.match(RANGE_PATTERN);
  if (range) {
    const first = Number.parseFloat(range[1]!);
    const second = Number.parseFloat(range[2]!);
    // "18-12" is a typo, not an empty range.
    return { kind: 'range', from: Math.min(first, second), to: Math.max(first, second) };
  }

  const number = text.match(NUMBER_PATTERN);
  if (number) {
    return { kind: 'number', value: Number.parseFloat(number[1]!), text: text.toLowerCase() };
  }

  return { kind: 'text', text: text.toLowerCase() };
}

export function matchesEpisodeFilter(episode: FilterableEpisode, filter: EpisodeFilter): boolean {
  const name = episode.name.toLowerCase();
  switch (filter.kind) {
    case 'number':
      return episode.number === filter.value || name.includes(filter.text);
    case 'range':
      return (
        episode.number !== null && episode.number >= filter.from && episode.number <= filter.to
      );
    case 'text':
      return name.includes(filter.text);
  }
}

export function filterEpisodes<T extends FilterableEpisode>(episodes: T[], query: string): T[] {
  const filter = parseEpisodeFilter(query);
  if (!filter) return episodes;
  return episodes.filter((episode) => matchesEpisodeFilter(episode, filter));
}
