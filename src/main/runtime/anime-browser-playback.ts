import type { AnimeBridgeClient, BridgeSource } from '../../anime-bridge/bridge-client';
import { buildAnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import { resolveStream } from '../../anime-bridge/headers';
import { resolveBridgeMediaUrl, routeHlsThroughProxy } from '../../anime-bridge/media-url';
import {
  buildPlaybackCommands,
  buildTrackCommands,
  selectPreferredStream,
} from '../../anime-bridge/mpv-playback';
import { watchPlaybackOutcome } from '../../anime-bridge/playback-outcome';
import { cacheSubtitleTracks, removeSubtitleCache } from '../../anime-bridge/subtitle-cache';
import type { StreamStripProxyHandle } from '../../anime-bridge/stream-strip-proxy';
import type { AnimeBrowserPlayRequest, AnimeBrowserPlayResult } from '../../types/anime-browser';
import type { AnimeBrowserPlaybackDeps } from './anime-browser-runtime-deps';

const TRACK_ATTACH_DELAY_MS = 300;

interface AnimeBrowserPlaybackOptions {
  deps: AnimeBrowserPlaybackDeps;
  bridge: () => Promise<{ client: AnimeBridgeClient; baseUrl: string }>;
  sourceFor: (sourceId: string) => Promise<BridgeSource>;
  stripProxy: () => StreamStripProxyHandle | null;
}

export function createAnimeBrowserPlayback(options: AnimeBrowserPlaybackOptions) {
  const { deps, bridge, sourceFor, stripProxy } = options;
  const wait =
    deps.wait ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  let subtitleCacheDir: string | null = null;
  // Overlapping playEpisode calls share subtitleCacheDir, so each call carries a
  // generation and only writes the shared slot while it is still the newest one.
  let cacheGeneration = 0;

  async function clearSubtitleCache(generation: number): Promise<void> {
    const previousDir = subtitleCacheDir;
    if (generation === cacheGeneration) subtitleCacheDir = null;
    await removeSubtitleCache(previousDir, deps.subtitleCacheIo);
  }

  async function cacheStreamSubtitles(
    stream: {
      headers: Record<string, string>;
      subtitles: Array<{ url: string; lang: string }>;
    },
    generation: number,
  ): Promise<Array<{ url: string; lang: string }>> {
    const cached = await cacheSubtitleTracks({
      tracks: stream.subtitles,
      headers: stream.headers,
      io: deps.subtitleCacheIo,
      log: deps.log,
    });
    if (generation === cacheGeneration) {
      subtitleCacheDir = cached.dir;
    } else {
      await removeSubtitleCache(cached.dir, deps.subtitleCacheIo);
    }

    const localCount = cached.tracks.filter((track) => track.local).length;
    if (cached.tracks.length > 0) {
      deps.log(
        `[anime-browser] cached ${localCount}/${cached.tracks.length} subtitle track(s) to disk` +
          (cached.dir ? ` in ${cached.dir}` : ''),
      );
    }
    return cached.tracks.map((track) => ({ url: track.url, lang: track.lang }));
  }

  async function playEpisode(request: AnimeBrowserPlayRequest): Promise<AnimeBrowserPlayResult> {
    const generation = ++cacheGeneration;
    try {
      const { client, baseUrl } = await bridge();
      const videos = await client.getVideoList(
        await sourceFor(request.sourceId),
        request.episodeUrl,
      );
      const streams = videos
        .map((video) => resolveStream(video))
        .filter((stream): stream is NonNullable<typeof stream> => stream !== null)
        .map((stream) => ({
          ...stream,
          url: resolveBridgeMediaUrl(baseUrl, stream.url),
          audios: stream.audios.map((track) => ({
            ...track,
            url: resolveBridgeMediaUrl(baseUrl, track.url),
          })),
          subtitles: stream.subtitles.map((track) => ({
            ...track,
            url: resolveBridgeMediaUrl(baseUrl, track.url),
          })),
        }));

      const selected = selectPreferredStream(streams, deps.preferredQuality?.());
      if (!selected) {
        return { ok: false, error: 'That source returned no playable video.', quality: null };
      }
      const proxy = stripProxy();
      const stream = proxy
        ? { ...selected, url: routeHlsThroughProxy(selected.url, baseUrl, proxy.origin) }
        : selected;

      if (!(await deps.ensureMpvConnected())) {
        return { ok: false, error: 'mpv is not running and could not be started.', quality: null };
      }

      const watch =
        deps.onPlaybackEndFile && deps.readMpvProperty
          ? watchPlaybackOutcome({
              onEndFile: deps.onPlaybackEndFile,
              readProperty: deps.readMpvProperty,
              wait,
            })
          : null;

      try {
        const metadata = buildAnimeStreamMetadata({
          sourceId: request.sourceId,
          animeUrl: request.animeUrl,
          animeTitle: request.animeTitle,
          episodeUrl: request.episodeUrl,
          episodeName: request.episodeName,
          episodeNumber: request.episodeNumber ?? null,
          mediaPath: stream.url,
        });
        const title = metadata.displayTitle;
        deps.onPlaybackMetadata?.(metadata);
        for (const command of buildPlaybackCommands({ stream, title })) {
          deps.sendMpvCommand(command);
        }
        // The file is already loading; a subtitle cache or track attach failure
        // costs extra tracks, not the episode, so it must not fail playback.
        try {
          await clearSubtitleCache(generation);

          if (stream.audios.length > 0 || stream.subtitles.length > 0) {
            deps.log(
              `[anime-browser] ${stream.audios.length} external audio, ` +
                `${stream.subtitles.length} external subtitle track(s)`,
            );
            const [subtitles] = await Promise.all([
              cacheStreamSubtitles(stream, generation),
              wait(TRACK_ATTACH_DELAY_MS),
            ]);
            for (const command of buildTrackCommands({ ...stream, subtitles })) {
              deps.sendMpvCommand(command);
            }
          }
        } catch (error) {
          deps.log(`[anime-browser] external track setup failed: ${String(error)}`);
        }

        if (watch) {
          const outcome = await watch.wait();
          if (!outcome.ok) {
            deps.log(`[anime-browser] playback failed to start: ${outcome.error}`);
            return { ok: false, error: outcome.error, quality: null };
          }
        }

        deps.showVisibleOverlay?.();
        deps.showMpvOsd?.(title);
        return { ok: true, error: null, quality: stream.quality || null };
      } finally {
        watch?.dispose();
      }
    } catch (error) {
      deps.log(`[anime-browser] playback failed: ${String(error)}`);
      return { ok: false, error: describeError(error), quality: null };
    }
  }

  async function dispose(): Promise<void> {
    const cacheDir = subtitleCacheDir;
    subtitleCacheDir = null;
    await removeSubtitleCache(cacheDir, deps.subtitleCacheIo);
  }

  return { playEpisode, dispose };
}

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
