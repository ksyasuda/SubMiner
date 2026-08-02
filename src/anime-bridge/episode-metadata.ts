/**
 * Structured metadata for a streamed episode.
 *
 * Extensions hand us two free-form strings — an anime title that usually
 * carries the season ("… Season 3") and an episode label that usually carries
 * the number ("Episode 4: …"). Everything downstream (stats grouping, AniList,
 * the subtitle modals) wants those as separate fields, so they are split once
 * here rather than re-parsed out of the mpv title by each consumer.
 */

/** Where a stream came from, resolved into the fields consumers actually want. */
export interface AnimeStreamMetadata {
  /** The URL handed to mpv. Matches what mpv reports as `path`. */
  mediaPath: string;
  /**
   * Stable identity for this episode. The stream URL carries a per-playback
   * proxy port and token, so it cannot be the key stats stores.
   */
  statsPath: string;
  /** Series name with the season suffix removed. */
  seriesTitle: string;
  seasonNumber: number | null;
  episodeNumber: number | null;
  /** The episode's own name, or null when the label was only a number. */
  episodeTitle: string | null;
  /** Shown by mpv, and the fallback every string parser sees. */
  displayTitle: string;
}

export interface AnimeStreamMetadataInput {
  sourceId: string;
  animeUrl: string;
  animeTitle: string;
  episodeUrl: string;
  episodeName: string;
  /** Extension-reported number; trusted over anything parsed from the label. */
  episodeNumber: number | null;
  /** The URL playback actually uses, after proxy rewriting. */
  mediaPath: string;
}

/**
 * Season suffixes, anchored to the end of the title so a "Season" that is part
 * of the name ("A Season of Snow") cannot be mistaken for one.
 */
const SEASON_SUFFIX_PATTERNS: RegExp[] = [
  /[\s:_-]+season\s*(\d{1,2})\s*$/i,
  /[\s:_-]+(\d{1,2})(?:st|nd|rd|th)\s+season\s*$/i,
  /[\s:_-]+s(\d{1,2})\s*$/i,
  /[\s:_-]*第\s*(\d{1,2})\s*期\s*$/,
  /[\s:_-]+(\d{1,2})\s*期\s*$/,
];

/**
 * Episode labels, most specific first. The trailing group is the episode's own
 * name when the label carries one.
 */
const EPISODE_LABEL_PATTERNS: RegExp[] = [
  /^\s*(?:episodio|épisode|episode|ep|e)\s*[.#]?\s*(\d{1,4}(?:\.\d+)?)\s*(?:[:\-–—.)]+\s*(.*))?$/i,
  /^\s*第\s*(\d{1,4})\s*話\s*(?:[:\-–—]+\s*)?(.*)$/,
  /^\s*(\d{1,4}(?:\.\d+)?)\s*[:\-–—.)]+\s*(.*)$/,
  /^\s*(\d{1,4}(?:\.\d+)?)\s*$/,
];

function collapseWhitespace(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

/**
 * Trims separators a split left dangling on either end. `.` is deliberately not
 * one of them: an episode name often ends in a full stop that belongs to it.
 */
function trimSeparators(value: string): string {
  return collapseWhitespace(value)
    .replace(/^[\s:_\-–—]+/, '')
    .replace(/[\s:_\-–—]+$/, '')
    .trim();
}

function toEpisodeNumber(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) return null;
  return value;
}

/**
 * Split a trailing season marker off an anime title.
 *
 * "Mushoku Tensei: Jobless Reincarnation Season 3" becomes the series plus
 * season 3, which is what both AniList and the stats grouping key want. A title
 * with no marker is returned unchanged with a null season — season 1 is *not*
 * assumed, because "unknown" and "one" behave differently when grouping.
 */
export function splitSeasonFromTitle(animeTitle: string): {
  title: string;
  season: number | null;
} {
  const normalized = collapseWhitespace(animeTitle);
  for (const pattern of SEASON_SUFFIX_PATTERNS) {
    const match = normalized.match(pattern);
    if (!match || match.index === undefined) continue;
    const season = Number.parseInt(match[1]!, 10);
    if (!Number.isInteger(season) || season <= 0) continue;
    const title = trimSeparators(normalized.slice(0, match.index));
    // A title that is *only* a season marker is not a title; keep the original.
    if (!title) continue;
    return { title, season };
  }
  return { title: normalized, season: null };
}

/**
 * Split an episode label into its number and its own name.
 *
 * Sources are inconsistent here: "Episode 4", "4. Title", "第4話 タイトル" and a
 * bare "4" all show up. A label that matches nothing is treated as a pure
 * episode name, which is right for movies and specials.
 */
export function splitEpisodeLabel(episodeName: string): {
  number: number | null;
  title: string | null;
} {
  const normalized = collapseWhitespace(episodeName);
  if (!normalized) return { number: null, title: null };

  for (const pattern of EPISODE_LABEL_PATTERNS) {
    const match = normalized.match(pattern);
    if (!match) continue;
    const parsed = Number.parseFloat(match[1]!);
    if (!Number.isFinite(parsed) || parsed <= 0) continue;
    const title = trimSeparators(match[2] ?? '');
    return { number: parsed, title: title || null };
  }

  return { number: null, title: normalized };
}

function formatEpisodePart(value: number): string {
  return Number.isInteger(value) ? String(value).padStart(2, '0') : String(value);
}

/**
 * The title mpv shows.
 *
 * `SxxEyy` is not just for looks: it is the one form both guessit and
 * SubMiner's own filename parser read reliably, so any consumer that only ever
 * sees the title string still lands on the right series and episode.
 */
export function buildStreamDisplayTitle(
  seriesTitle: string,
  season: number | null,
  episode: number | null,
  episodeTitle: string | null,
): string {
  const parts: string[] = [seriesTitle];
  if (episode !== null) {
    parts.push(
      season !== null
        ? `S${String(season).padStart(2, '0')}E${formatEpisodePart(episode)}`
        : `E${formatEpisodePart(episode)}`,
    );
  } else if (season !== null) {
    parts.push(`S${String(season).padStart(2, '0')}`);
  }

  const head = parts.join(' ');
  return episodeTitle ? `${head} - ${episodeTitle}` : head;
}

/**
 * A per-episode identity that survives across playbacks.
 *
 * The stream URL points at the strip proxy, whose port and token are minted per
 * playback, so keying stats on it makes every rewatch a new video. The source's
 * own episode url is stable, so that is what stats records instead — with the
 * real URL kept as an alias so mpv's path change still finds the row.
 */
export function buildAnimeStreamStatsPath(
  sourceId: string,
  animeUrl: string,
  episodeUrl: string,
): string {
  const source = encodeURIComponent(sourceId || 'unknown');
  const anime = encodeURIComponent(animeUrl || 'unknown');
  const episode = encodeURIComponent(episodeUrl || 'unknown');
  return `animebrowser://${source}/${anime}/${episode}`;
}

export function buildAnimeStreamMetadata(input: AnimeStreamMetadataInput): AnimeStreamMetadata {
  const { title: seriesTitle, season } = splitSeasonFromTitle(input.animeTitle);
  const label = splitEpisodeLabel(input.episodeName);
  const episodeNumber = toEpisodeNumber(input.episodeNumber) ?? label.number;
  const displayTitle = buildStreamDisplayTitle(seriesTitle, season, episodeNumber, label.title);

  return {
    mediaPath: input.mediaPath,
    statsPath: buildAnimeStreamStatsPath(input.sourceId, input.animeUrl, input.episodeUrl),
    seriesTitle,
    seasonNumber: season,
    episodeNumber,
    episodeTitle: label.title,
    // A source that gave us neither a number nor a name leaves the series title
    // alone rather than showing an empty suffix.
    displayTitle: displayTitle || collapseWhitespace(input.animeTitle),
  };
}
