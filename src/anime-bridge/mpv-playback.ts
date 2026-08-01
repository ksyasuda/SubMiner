import { toMpvHeaderFields } from './headers';
import type { ResolvedStream } from './types';

/**
 * Japanese first, always, and for subtitles Japanese *only*: the primary slot
 * belongs to the language being mined, and an English track belongs in the
 * secondary slot, where the `secondarySub` machinery puts it by language tag.
 * For audio, mpv falls back to the first track when nothing matches, so an
 * English-only release still plays.
 */
export const JAPANESE_LANGUAGE_PREFERENCE = 'ja,jpn,jp,japanese';

/**
 * mpv must not scan the filesystem for sidecar subtitles when the "file" is a
 * network stream, and the secondary slot stays empty so the overlay only ever
 * reads one track. Everything else is left to normal track selection, driven
 * by the language preferences above.
 *
 * `alang`/`slang` are set as properties instead of file-local options: their
 * values are comma-separated lists, and a comma inside a `loadfile` option
 * value splits the option list.
 */
const BASE_LOADFILE_OPTIONS = [
  'sub-auto=no',
  'secondary-sid=no',
  'secondary-sub-visibility=no',
  'sub-visibility=yes',
];

/** Matches a language tag or a label such as "Japanese (Sub)" or "[JPN]". */
const JAPANESE_PATTERN = /(^|[^a-z])(ja|jp|jpn|japanese|日本語)([^a-z]|$)/i;
/** Extensions label dub entries in the quality string, e.g. "1080p (Dub)". */
const DUB_PATTERN = /(^|[^a-z])(dub|dubbed|dublado|latino|castellano)([^a-z]|$)/i;
/** The counterpart label for original-audio entries, e.g. "SUB - 1080p". */
const SUBBED_PATTERN = /(^|[^a-z])(sub|subbed|softsub|hardsub|subtitulado|raw)([^a-z]|$)/i;

export type MpvCommand = Array<string | number>;

export interface BuildPlaybackOptions {
  stream: ResolvedStream;
  /** Shown as the mpv window/OSD title. */
  title?: string;
  /** Resume position in seconds. */
  startSeconds?: number;
}

export function isJapaneseTag(value: string): boolean {
  return JAPANESE_PATTERN.test(value);
}

/**
 * Build the mpv `loadfile` option string for a stream.
 *
 * Headers ride as `file-local-options/http-header-fields` so they apply to this
 * file only, and so SubMiner's Anki media path can read them back off mpv when
 * generating card audio and screenshots. Tracks added later with `sub-add` /
 * `audio-add` inherit them too, which is how external tracks on an
 * authenticated host stay reachable.
 */
export function buildLoadfileOptions(options: BuildPlaybackOptions): string {
  const parts = [...BASE_LOADFILE_OPTIONS];

  const headerFields = toMpvHeaderFields(options.stream.headers);
  if (headerFields.length > 0) {
    // Escape the mpv option-list separators so a header never splits the list.
    parts.push(`http-header-fields=${escapeOptionValue(headerFields)}`);
  }

  if (options.startSeconds !== undefined && options.startSeconds > 0) {
    parts.push(`start=${options.startSeconds}`);
  }

  return parts.join(',');
}

/**
 * mpv splits `loadfile` options on commas and `=`-separates keys, so a value
 * containing either must be quoted. Percent-encoding is mpv's own escape for
 * embedded separators in option values.
 *
 * The count is in UTF-8 bytes of the decoded value, not JS string units, so a
 * non-ASCII header value (extensions supply these) would otherwise under-count
 * and mpv would cut the value short.
 */
function escapeOptionValue(value: string): string {
  return `%${Buffer.byteLength(value, 'utf8')}%${value}`;
}

/**
 * Ordered mpv commands that start playback of a resolved stream.
 *
 * The plugin is told subtitles are being managed before the file loads, so the
 * overlay does not flash the source's own tracks during the swap.
 */
export function buildPlaybackCommands(options: BuildPlaybackOptions): MpvCommand[] {
  const commands: MpvCommand[] = [
    ['script-message', 'subminer-managed-subtitles-loading'],
    ['set_property', 'alang', JAPANESE_LANGUAGE_PREFERENCE],
    ['set_property', 'slang', JAPANESE_LANGUAGE_PREFERENCE],
    ['loadfile', options.stream.url, 'replace', -1, buildLoadfileOptions(options)],
  ];

  if (options.title !== undefined && options.title.length > 0) {
    commands.push(['set_property', 'force-media-title', options.title]);
  }

  return commands;
}

/**
 * Commands that attach the extension's external audio and subtitle tracks.
 *
 * These must be sent *after* the file is loading, so they are separate from
 * {@link buildPlaybackCommands}. Every track is added — even the ones we do not
 * select — so they show up in mpv's track menu and can be switched by hand.
 */
export function buildTrackCommands(stream: ResolvedStream): MpvCommand[] {
  return [
    ...buildAddTrackCommands('audio-add', stream.audios, 'Audio'),
    ...buildAddTrackCommands('sub-add', stream.subtitles, 'Subtitle'),
  ];
}

/**
 * Only a Japanese track is ever selected outright — the primary slot is for
 * the mining language. A non-Japanese track is added unselected: for audio,
 * `alang`'s pick off the container stands; for subtitles, the `secondarySub`
 * auto-load matches the track's language tag against the user's configured
 * secondary languages and routes it to `secondary-sid` instead.
 */
function buildAddTrackCommands(
  command: 'audio-add' | 'sub-add',
  tracks: Array<{ url: string; lang: string }>,
  kind: 'Audio' | 'Subtitle',
): MpvCommand[] {
  const unique = dedupeByUrl(tracks);
  const selected = unique.findIndex((track) => isJapaneseTag(track.lang));

  return unique.map((track, index) => [
    command,
    track.url,
    index === selected ? 'select' : 'auto',
    track.lang || `${kind} ${index + 1}`,
    normalizeLangTag(track.lang),
  ]);
}

/** Extension language labels mapped to the tags users put in config. */
const LANG_TAG_BY_LABEL: Record<string, string> = {
  japanese: 'ja',
  日本語: 'ja',
  english: 'en',
  eng: 'en',
  spanish: 'es',
  español: 'es',
  portuguese: 'pt',
  português: 'pt',
  french: 'fr',
  français: 'fr',
  german: 'de',
  deutsch: 'de',
  italian: 'it',
  italiano: 'it',
  indonesian: 'id',
  arabic: 'ar',
  russian: 'ru',
  korean: 'ko',
  chinese: 'zh',
  thai: 'th',
  vietnamese: 'vi',
};

/**
 * mpv's `lang` field is what SubMiner's secondary-subtitle auto-load compares
 * against `secondarySub.secondarySubLanguages`, so a label like "English" must
 * become the tag a user would actually configure. Unknown labels pass through;
 * matching is best-effort, and the raw label stays visible as the track title.
 */
export function normalizeLangTag(lang: string): string {
  const trimmed = lang.trim();
  if (isJapaneseTag(trimmed)) return 'ja';
  const mapped = LANG_TAG_BY_LABEL[trimmed.toLowerCase()];
  if (mapped !== undefined) return mapped;
  if (/^[A-Za-z]{2,3}([-_][A-Za-z0-9]+)?$/.test(trimmed)) {
    return trimmed.split(/[-_]/, 1)[0]?.toLowerCase() ?? trimmed.toLowerCase();
  }
  return trimmed;
}

function dedupeByUrl(
  tracks: Array<{ url: string; lang: string }>,
): Array<{ url: string; lang: string }> {
  const seen = new Set<string>();
  return tracks.filter((track) => {
    if (track.url.length === 0 || seen.has(track.url)) return false;
    seen.add(track.url);
    return true;
  });
}

/**
 * Rank a stream by how likely it is to carry Japanese audio.
 *
 * Sources commonly return the dub and the original as separate entries rather
 * than as two audio tracks of one entry, so the choice of *entry* is the first
 * place a dub can slip in.
 */
function scoreStream(stream: ResolvedStream): number {
  if (stream.audios.some((audio) => isJapaneseTag(audio.lang))) return 2;
  const label = stream.quality;
  if (isJapaneseTag(label) || SUBBED_PATTERN.test(label)) return 1;
  if (DUB_PATTERN.test(label)) return -1;
  return 0;
}

/**
 * Pick the best stream from an extension's video list.
 *
 * Japanese audio outranks the quality hint — a 1080p dub is the wrong file, not
 * a better one. Within the surviving entries the hint decides, and otherwise
 * the extension's own ordering does.
 */
export function selectPreferredStream(
  streams: ResolvedStream[],
  preferredQuality?: string,
): ResolvedStream | null {
  if (streams.length === 0) return null;

  const best = Math.max(...streams.map(scoreStream));
  const candidates = streams.filter((stream) => scoreStream(stream) === best);

  if (preferredQuality !== undefined && preferredQuality.length > 0) {
    const needle = preferredQuality.toLowerCase();
    const match = candidates.find((stream) => stream.quality.toLowerCase().includes(needle));
    if (match) return match;
  }
  return candidates[0] ?? null;
}
