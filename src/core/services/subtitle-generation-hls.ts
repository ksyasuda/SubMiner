import { mkdir, open, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { toRequestHeaders, type ResolvedMpvHttpHeaders } from './mpv-http-headers';

const DOWNLOAD_CONCURRENCY = 4;
const REQUEST_TIMEOUT_MS = 15_000;
const MAX_PLAYLIST_BYTES = 2 * 1024 * 1024;

// Keep complex playlists with FFmpeg, including encryption, byte ranges, variant
// selection, initialization segments and discontinuities that may change tracks.
const SUPPORTED_TAGS = new Set([
  '#EXTM3U',
  '#EXTINF',
  '#EXT-X-VERSION',
  '#EXT-X-TARGETDURATION',
  '#EXT-X-MEDIA-SEQUENCE',
  '#EXT-X-ENDLIST',
  '#EXT-X-PLAYLIST-TYPE',
  '#EXT-X-ALLOW-CACHE',
  '#EXT-X-INDEPENDENT-SEGMENTS',
  '#EXT-X-PROGRAM-DATE-TIME',
]);

function parsePlaylist(text: string, base: string) {
  const lines = text
    .trim()
    .split(/\r?\n/)
    .map((line) => line.trim());
  if (lines[0] !== '#EXTM3U' || !lines.includes('#EXT-X-ENDLIST')) return null;
  const segments: { url: string; filename: string; duration: number }[] = [];
  const localLines: string[] = [];
  let duration: number | undefined;
  for (const line of lines) {
    if (!line) continue;
    if (line.startsWith('#')) {
      if (!SUPPORTED_TAGS.has(line.split(':', 1)[0] ?? '')) return null;
      if (line.startsWith('#EXTINF:')) {
        if (duration !== undefined) return null;
        duration = Number(line.slice(8).split(',', 1)[0]);
        if (!Number.isFinite(duration) || duration <= 0) return null;
      }
      localLines.push(line);
      continue;
    }
    if (duration === undefined) return null;
    let url: URL;
    try {
      url = new URL(line, base);
    } catch {
      return null;
    }
    if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) return null;
    const filename = `segment-${segments.length}.ts`;
    segments.push({ url: url.href, filename, duration });
    localLines.push(filename);
    duration = undefined;
  }
  if (!segments.length || duration !== undefined) return null;
  return { segments, contents: `${localLines.join('\n')}\n` };
}

class UnsupportedSegment extends Error {}

async function downloadSegment(input: {
  url: string;
  destination: string;
  headers: Record<string, string>;
  signal: AbortSignal;
}): Promise<void> {
  const timeout = new AbortController();
  const timer = setTimeout(
    () => timeout.abort(new Error('Episode segment download timed out.')),
    REQUEST_TIMEOUT_MS,
  );
  timer.unref();
  try {
    const response = await fetch(input.url, {
      headers: input.headers,
      signal: AbortSignal.any([input.signal, timeout.signal]),
    });
    if (!response.ok || !response.body) {
      await response.body?.cancel();
      throw new Error(`Episode segment download failed: HTTP ${response.status}`);
    }
    const reader = response.body.getReader();
    try {
      const file = await open(input.destination, 'wx');
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          input.signal.throwIfAborted();
          // Match FFmpeg's inactivity timeout without limiting total transfer time.
          timer.refresh();
          await file.writeFile(value);
        }
      } finally {
        await file.close();
      }
    } finally {
      await reader.cancel().catch(() => undefined);
    }
  } finally {
    clearTimeout(timer);
  }
}

/** Stage plain MPEG-TS VOD segments concurrently without changing their order or timestamps. */
async function stageSubtitleGenerationHls(input: {
  mediaPath: string;
  directory: string;
  httpHeaders: ResolvedMpvHttpHeaders;
  signal?: AbortSignal;
  onProgress?: (percent: number) => void;
  window?: { startSeconds: number; durationSeconds: number };
}): Promise<{ playlistPath: string; seekSeconds: number } | null> {
  input.signal?.throwIfAborted();
  const controller = new AbortController();
  const signal = input.signal
    ? AbortSignal.any([input.signal, controller.signal])
    : controller.signal;
  const headers = toRequestHeaders(input.httpHeaders);
  let playlist: ReturnType<typeof parsePlaylist>;
  try {
    const response = await fetch(input.mediaPath, {
      headers,
      signal: AbortSignal.any([signal, AbortSignal.timeout(REQUEST_TIMEOUT_MS)]),
    });
    if (!response.ok || !response.body) {
      await response.body?.cancel();
      return null;
    }
    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let bytes = 0;
    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        bytes += value.byteLength;
        if (bytes > MAX_PLAYLIST_BYTES) return null;
        chunks.push(value);
      }
    } finally {
      await reader.cancel().catch(() => undefined);
    }
    playlist = parsePlaylist(Buffer.concat(chunks).toString('utf8'), response.url);
  } catch {
    input.signal?.throwIfAborted();
    // FFmpeg supports request contexts and playlist formats that fetch may not.
    return null;
  }
  if (!playlist) return null;
  let seekSeconds = 0;
  if (input.window) {
    const { startSeconds, durationSeconds } = input.window;
    if (
      !Number.isFinite(startSeconds) ||
      startSeconds < 0 ||
      !Number.isFinite(durationSeconds) ||
      durationSeconds <= 0
    )
      throw new Error('Invalid HLS sample window.');
    let position = 0;
    const segments = playlist.segments.filter((segment) => {
      const start = position;
      position += segment.duration;
      if (position <= startSeconds || start >= startSeconds + durationSeconds) return false;
      if (start <= startSeconds) seekSeconds = startSeconds - start;
      return true;
    });
    if (!segments.length) return null;
    playlist = {
      segments,
      contents:
        '#EXTM3U\n#EXT-X-VERSION:3\n#EXT-X-TARGETDURATION:' +
        Math.ceil(Math.max(...segments.map((segment) => segment.duration))) +
        '\n' +
        segments.map((segment) => `#EXTINF:${segment.duration},\n${segment.filename}\n`).join('') +
        '#EXT-X-ENDLIST\n',
    };
  }
  const { segments } = playlist;
  const directory = path.join(input.directory, 'hls');
  await mkdir(directory);
  const totalDuration = segments.reduce((sum, segment) => sum + segment.duration, 0);
  let completedDuration = 0;
  let next = 0;
  input.onProgress?.(0);
  const worker = async () => {
    try {
      while (true) {
        signal.throwIfAborted();
        const segment = segments[next++];
        if (!segment) return;
        await downloadSegment({
          url: segment.url,
          destination: path.join(directory, segment.filename),
          headers,
          signal,
        });
        // URL extensions are often absent or disguised. Only stage MPEG-TS here;
        // packed audio and other formats retain FFmpeg's original demuxing path.
        const file = await open(path.join(directory, segment.filename), 'r');
        try {
          const header = Buffer.alloc(377);
          const { bytesRead } = await file.read(header, 0, header.length, 0);
          if (
            bytesRead < header.length ||
            header[0] !== 0x47 ||
            header[188] !== 0x47 ||
            header[376] !== 0x47
          )
            throw new UnsupportedSegment();
        } finally {
          await file.close();
        }
        signal.throwIfAborted();
        completedDuration += segment.duration;
        input.onProgress?.(Math.floor((completedDuration / totalDuration) * 100));
      }
    } catch (error) {
      controller.abort(error);
      throw error;
    }
  };
  // Drain aborted workers before the caller removes the temporary directory.
  await Promise.allSettled(
    Array.from({ length: Math.min(DOWNLOAD_CONCURRENCY, segments.length) }, worker),
  );
  input.signal?.throwIfAborted();
  if (controller.signal.aborted) {
    if (controller.signal.reason instanceof UnsupportedSegment) {
      await rm(directory, { recursive: true, force: true });
      return null;
    }
    throw controller.signal.reason;
  }
  const localPath = path.join(directory, 'episode.m3u8');
  await writeFile(localPath, playlist.contents, { flag: 'wx' });
  return { playlistPath: localPath, seekSeconds };
}

type HlsDownloadInput = Omit<Parameters<typeof stageSubtitleGenerationHls>[0], 'window'>;

export async function downloadSubtitleGenerationHls(
  input: HlsDownloadInput,
): Promise<string | null> {
  return (await stageSubtitleGenerationHls(input))?.playlistPath ?? null;
}

/** Seek using the VOD segment timeline; some remote HLS seeks exit with empty output. */
export function downloadSubtitleGenerationHlsWindow(
  input: HlsDownloadInput & {
    window: { startSeconds: number; durationSeconds: number };
  },
) {
  return stageSubtitleGenerationHls(input);
}
