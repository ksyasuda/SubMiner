import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/**
 * Extension subtitle tracks arrive as URLs on the bridge's loopback proxy, and
 * mpv is perfectly happy to stream them. alass is not: it needs a file on disk
 * to use as the timing reference, and the subsync path rejects an external
 * track whose `external-filename` is not an existing file. So every track is
 * downloaded to a temp directory first and mpv is given the local path, the
 * same way the Jellyfin preload caches its delivery URLs.
 *
 * The directory outlives the `sub-add` — alass reads it mid-playback — and is
 * removed when the next episode starts or the runtime shuts down.
 */

/** Extensions mpv and alass both recognise off a filename. */
const KNOWN_SUBTITLE_EXTENSIONS = new Set([
  'srt',
  'ass',
  'ssa',
  'vtt',
  'sub',
  'ttml',
  'smi',
  'sbv',
]);

/** What an unrecognisable track is named; mpv still probes the content. */
const DEFAULT_SUBTITLE_EXTENSION = 'srt';

const DOWNLOAD_TIMEOUT_MS = 15_000;

/** Anything this large is not a subtitle file, and is not worth buffering. */
const MAX_SUBTITLE_BYTES = 32 * 1024 * 1024;

/** How much of the body is decoded to guess the format. */
const SNIFF_BYTES = 1024;

export interface SubtitleTrackRef {
  url: string;
  lang: string;
}

export interface CachedSubtitleTrack extends SubtitleTrackRef {
  /** Where the track came from, kept for logs. */
  sourceUrl: string;
  /** False when the download failed and `url` is still the remote URL. */
  local: boolean;
}

export interface SubtitleCacheResult {
  /** The temp directory to remove later, or null when nothing was cached. */
  dir: string | null;
  tracks: CachedSubtitleTrack[];
}

interface FetchResponseLike {
  ok: boolean;
  status: number;
  arrayBuffer: () => Promise<ArrayBuffer>;
}

export interface SubtitleCacheIo {
  fetch: (
    url: string,
    init: { headers: Record<string, string>; signal: AbortSignal },
  ) => Promise<FetchResponseLike>;
  makeTempDir: (prefix: string) => Promise<string>;
  writeFile: (filePath: string, bytes: Uint8Array) => Promise<void>;
  removeDir: (dir: string) => Promise<void>;
}

export interface CacheSubtitleTracksOptions {
  tracks: SubtitleTrackRef[];
  /** Headers the stream was resolved with; some hosts gate subtitles too. */
  headers?: Record<string, string>;
  io?: SubtitleCacheIo;
  log?: (message: string) => void;
}

export function createSubtitleCacheIo(): SubtitleCacheIo {
  return {
    fetch: (url, init) => fetch(url, init),
    makeTempDir: (prefix) => fs.promises.mkdtemp(prefix),
    writeFile: (filePath, bytes) => fs.promises.writeFile(filePath, bytes),
    removeDir: (dir) => fs.promises.rm(dir, { recursive: true, force: true }),
  };
}

/**
 * Guess a subtitle format from the start of the file.
 *
 * Bridge subtitle URLs are opaque tokens far more often than they are
 * filenames, so the content is the only reliable signal. mpv and alass both
 * pick their parser off the extension, and a `.srt` holding ASS is a parse
 * error rather than a mistimed subtitle.
 */
export function sniffSubtitleExtension(head: string): string | null {
  const text = head.replace(/^\uFEFF/, '').trimStart();
  if (/^\[(script info|v4\+? styles|events)\]/i.test(text)) return 'ass';
  if (/^WEBVTT(\s|$)/.test(text)) return 'vtt';
  if (/^<\?xml/i.test(text) && /<tt[\s>]|ttml/i.test(text)) return 'ttml';
  // Cue-numbered and bare-timestamp SRT; the `.` separator is a common variant.
  if (/^(\d+\s*\r?\n)?\d{1,3}:\d{2}:\d{2}[,.]\d{1,3}\s*-->/.test(text)) return 'srt';
  return null;
}

/** The URL's own extension, when it names a format we know. */
export function subtitleExtensionFromUrl(url: string): string | null {
  const urlPath = (() => {
    try {
      return new URL(url).pathname;
    } catch {
      return url;
    }
  })();
  const extension = path.extname(urlPath).slice(1).toLowerCase();
  return KNOWN_SUBTITLE_EXTENSIONS.has(extension) ? extension : null;
}

/** Content first, then the URL, then a guess mpv can still probe past. */
export function resolveSubtitleExtension(url: string, head: string): string {
  return (
    sniffSubtitleExtension(head) ?? subtitleExtensionFromUrl(url) ?? DEFAULT_SUBTITLE_EXTENSION
  );
}

function dedupeByUrl(tracks: SubtitleTrackRef[]): SubtitleTrackRef[] {
  const seen = new Set<string>();
  return tracks.filter((track) => {
    if (track.url.length === 0 || seen.has(track.url)) return false;
    seen.add(track.url);
    return true;
  });
}

async function downloadTrack(
  io: SubtitleCacheIo,
  dir: string,
  index: number,
  track: SubtitleTrackRef,
  headers: Record<string, string>,
): Promise<string> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), DOWNLOAD_TIMEOUT_MS);
  let bytes: Uint8Array;
  try {
    const response = await io.fetch(track.url, { headers, signal: controller.signal });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    bytes = new Uint8Array(await response.arrayBuffer());
  } finally {
    clearTimeout(timeoutId);
  }

  if (bytes.byteLength === 0) {
    throw new Error('empty response');
  }
  if (bytes.byteLength > MAX_SUBTITLE_BYTES) {
    throw new Error(`response too large (${bytes.byteLength} bytes)`);
  }

  const head = Buffer.from(bytes.subarray(0, SNIFF_BYTES)).toString('utf8');
  const extension = resolveSubtitleExtension(track.url, head);
  const filePath = path.join(dir, `track-${index}.${extension}`);
  // Written byte for byte: re-encoding would corrupt a non-UTF-8 track that
  // mpv's own charset detection would otherwise handle.
  await io.writeFile(filePath, bytes);
  return filePath;
}

/**
 * Download every subtitle track to a fresh temp directory.
 *
 * A track that fails to download keeps its remote URL, so a dead subtitle
 * server costs the alass reference rather than the episode.
 */
export async function cacheSubtitleTracks(
  options: CacheSubtitleTracksOptions,
): Promise<SubtitleCacheResult> {
  const io = options.io ?? createSubtitleCacheIo();
  const tracks = dedupeByUrl(options.tracks);
  if (tracks.length === 0) return { dir: null, tracks: [] };

  const dir = await io.makeTempDir(path.join(os.tmpdir(), 'subminer-anime-subtitles-'));
  const cached = await Promise.all(
    tracks.map(async (track, index): Promise<CachedSubtitleTrack> => {
      try {
        const filePath = await downloadTrack(io, dir, index, track, options.headers ?? {});
        return { url: filePath, lang: track.lang, sourceUrl: track.url, local: true };
      } catch (error) {
        options.log?.(
          `[anime-browser] subtitle download failed (${track.lang || 'unknown'}): ` +
            describeError(error),
        );
        return { url: track.url, lang: track.lang, sourceUrl: track.url, local: false };
      }
    }),
  );

  if (!cached.some((track) => track.local)) {
    await removeSubtitleCache(dir, io);
    return { dir: null, tracks: cached };
  }
  return { dir, tracks: cached };
}

/** Remove a cache directory. Never throws: cleanup is best effort. */
export async function removeSubtitleCache(
  dir: string | null,
  io: SubtitleCacheIo = createSubtitleCacheIo(),
): Promise<void> {
  if (!dir) return;
  try {
    await io.removeDir(dir);
  } catch {}
}

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
