import type { AnimeBridgeClient, BridgeSource } from '../../anime-bridge/bridge-client';
import { buildAnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import { resolveStream } from '../../anime-bridge/headers';
import { resolveBridgeMediaUrl, routeHlsThroughProxy } from '../../anime-bridge/media-url';
import {
  buildPlaybackCommands,
  buildQueuedPlaybackCommands,
  buildTrackCommands,
  selectPreferredStream,
} from '../../anime-bridge/mpv-playback';
import { watchPlaybackOutcome } from '../../anime-bridge/playback-outcome';
import { cacheSubtitleTracks, removeSubtitleCache } from '../../anime-bridge/subtitle-cache';
import type { StreamStripProxyHandle } from '../../anime-bridge/stream-strip-proxy';
import type { AnimeStreamMetadata } from '../../anime-bridge/episode-metadata';
import type { ResolvedStream } from '../../anime-bridge/types';
import type { AnimeBrowserPlayRequest, AnimeBrowserPlayResult } from '../../types/anime-browser';
import type { AnimeBrowserPlaybackDeps } from './anime-browser-runtime-deps';

const TRACK_ATTACH_DELAY_MS = 300;

interface AnimeBrowserPlaybackOptions {
  deps: AnimeBrowserPlaybackDeps;
  bridge: () => Promise<{ client: AnimeBridgeClient; baseUrl: string }>;
  sourceFor: (sourceId: string) => Promise<BridgeSource>;
  stripProxy: () => StreamStripProxyHandle | null;
}

export interface PreparedAnimeBrowserPlayback {
  request: AnimeBrowserPlayRequest;
  stream: ResolvedStream;
  metadata: AnimeStreamMetadata;
  trackPreparation: Promise<PreparedTrackSetup>;
}

interface PreparedTrackSetup {
  stream: ResolvedStream;
  subtitleCacheDir: string | null;
}

export type PrepareAnimeBrowserPlaybackResult =
  | { ok: true; playback: PreparedAnimeBrowserPlayback; quality: string | null; error: null }
  | { ok: false; playback: null; quality: null; error: string };

export function createAnimeBrowserPlayback(options: AnimeBrowserPlaybackOptions) {
  const { deps, bridge, sourceFor, stripProxy } = options;
  const wait =
    deps.wait ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  let subtitleCacheDir: string | null = null;
  const queuedSubtitleCacheDirs = new Set<string>();
  const queuedTrackPreparations = new Set<Promise<PreparedTrackSetup>>();
  // Overlapping playEpisode calls share mpv and subtitleCacheDir, so each call
  // carries a generation and only acts while it is still the newest one.
  let playbackGeneration = 0;

  async function clearSubtitleCache(generation: number): Promise<void> {
    // A stale call must not delete the directory a newer playback now owns.
    if (generation !== playbackGeneration) return;
    const previousDir = subtitleCacheDir;
    subtitleCacheDir = null;
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
    if (generation === playbackGeneration) {
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

  async function resolveEpisode(request: AnimeBrowserPlayRequest): Promise<{
    stream: ResolvedStream;
    metadata: AnimeStreamMetadata;
  } | null> {
    const { client, baseUrl } = await bridge();
    const videos = await client.getVideoList(await sourceFor(request.sourceId), request.episodeUrl);
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
    if (!selected) return null;
    const proxy = stripProxy();
    const stream = proxy
      ? { ...selected, url: routeHlsThroughProxy(selected.url, baseUrl, proxy.origin) }
      : selected;
    const metadata = buildAnimeStreamMetadata({
      sourceId: request.sourceId,
      animeUrl: request.animeUrl,
      animeTitle: request.animeTitle,
      episodeUrl: request.episodeUrl,
      episodeName: request.episodeName,
      episodeNumber: request.episodeNumber ?? null,
      mediaPath: stream.url,
    });
    return { stream, metadata };
  }

  /** Resolve a queued stream and start preparing its external tracks. */
  async function prepareEpisode(
    request: AnimeBrowserPlayRequest,
  ): Promise<PrepareAnimeBrowserPlaybackResult> {
    try {
      const resolved = await resolveEpisode(request);
      if (!resolved) {
        return {
          ok: false,
          playback: null,
          error: 'That source returned no playable video.',
          quality: null,
        };
      }
      if (!(await deps.ensureMpvConnected())) {
        return {
          ok: false,
          playback: null,
          error: 'mpv is not running and could not be started.',
          quality: null,
        };
      }

      // Do not hold the mpv append behind subtitle downloads. The preparation
      // starts now and activation waits for it only when this item begins.
      const trackPreparation = prepareQueuedTracks(resolved.stream);
      return {
        ok: true,
        playback: {
          request,
          stream: resolved.stream,
          metadata: resolved.metadata,
          trackPreparation,
        },
        error: null,
        quality: resolved.stream.quality || null,
      };
    } catch (error) {
      deps.log(`[anime-browser] queued playback preparation failed: ${String(error)}`);
      return { ok: false, playback: null, error: describeError(error), quality: null };
    }
  }

  function prepareQueuedTracks(stream: ResolvedStream): Promise<PreparedTrackSetup> {
    const preparation = (async (): Promise<PreparedTrackSetup> => {
      try {
        const cached = await cacheSubtitleTracks({
          tracks: stream.subtitles,
          headers: stream.headers,
          io: deps.subtitleCacheIo,
          log: deps.log,
        });
        if (cached.dir) queuedSubtitleCacheDirs.add(cached.dir);
        return {
          stream: {
            ...stream,
            subtitles: cached.tracks.map((track) => ({ url: track.url, lang: track.lang })),
          },
          subtitleCacheDir: cached.dir,
        };
      } catch (error) {
        deps.log(`[anime-browser] queued subtitle preparation failed: ${String(error)}`);
        return { stream, subtitleCacheDir: null };
      }
    })();
    queuedTrackPreparations.add(preparation);
    void preparation.then(() => queuedTrackPreparations.delete(preparation));
    return preparation;
  }

  /** The prepared stream becomes a real mpv playlist entry immediately. */
  function appendEpisode(playback: PreparedAnimeBrowserPlayback): void {
    deps.onPreparedPlaybackMetadata?.(playback.metadata);
    for (const command of buildQueuedPlaybackCommands({
      stream: playback.stream,
      title: playback.metadata.displayTitle,
    })) {
      deps.sendMpvCommand(command);
    }
  }

  /** Attach the prepared external tracks when mpv reaches this playlist item. */
  async function activateEpisode(playback: PreparedAnimeBrowserPlayback): Promise<void> {
    const generation = ++playbackGeneration;
    const preparedTracks = await playback.trackPreparation;
    if (generation !== playbackGeneration) {
      await releasePreparedTracks(preparedTracks);
      return;
    }
    const previousDir = subtitleCacheDir;
    subtitleCacheDir = preparedTracks.subtitleCacheDir;
    if (preparedTracks.subtitleCacheDir) {
      queuedSubtitleCacheDirs.delete(preparedTracks.subtitleCacheDir);
    }
    if (previousDir !== preparedTracks.subtitleCacheDir) {
      await removeSubtitleCache(previousDir, deps.subtitleCacheIo);
    }
    if (generation !== playbackGeneration) return;

    if (preparedTracks.stream.audios.length > 0 || preparedTracks.stream.subtitles.length > 0) {
      await wait(TRACK_ATTACH_DELAY_MS);
      if (generation !== playbackGeneration) return;
      for (const command of buildTrackCommands(preparedTracks.stream)) {
        deps.sendMpvCommand(command);
      }
    }
    deps.showVisibleOverlay?.();
    deps.showMpvOsd?.(playback.metadata.displayTitle);
  }

  async function discardEpisode(playback: PreparedAnimeBrowserPlayback): Promise<void> {
    await releasePreparedTracks(await playback.trackPreparation);
  }

  async function releasePreparedTracks(preparedTracks: PreparedTrackSetup): Promise<void> {
    const dir = preparedTracks.subtitleCacheDir;
    if (!dir || !queuedSubtitleCacheDirs.delete(dir)) return;
    await removeSubtitleCache(dir, deps.subtitleCacheIo);
  }

  async function playEpisode(request: AnimeBrowserPlayRequest): Promise<AnimeBrowserPlayResult> {
    const generation = ++playbackGeneration;
    const isCurrent = (): boolean => generation === playbackGeneration;
    try {
      const resolved = await resolveEpisode(request);
      if (!resolved) {
        return { ok: false, error: 'That source returned no playable video.', quality: null };
      }
      const { stream, metadata } = resolved;

      if (!(await deps.ensureMpvConnected())) {
        return { ok: false, error: 'mpv is not running and could not be started.', quality: null };
      }

      // Resolving the stream can outlast a newer click; from here on every step
      // touches mpv or the overlay, so a superseded call stops instead.
      if (!isCurrent()) return superseded();

      const watch =
        deps.onPlaybackEndFile && deps.readMpvProperty
          ? watchPlaybackOutcome({
              onEndFile: deps.onPlaybackEndFile,
              readProperty: deps.readMpvProperty,
              wait,
            })
          : null;

      try {
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
            if (!isCurrent()) return superseded();
            for (const command of buildTrackCommands({ ...stream, subtitles })) {
              deps.sendMpvCommand(command);
            }
          }
        } catch (error) {
          deps.log(`[anime-browser] external track setup failed: ${String(error)}`);
        }

        if (!isCurrent()) return superseded();

        if (watch) {
          const outcome = await watch.wait();
          // The outcome belongs to whichever file mpv is playing now, so a
          // superseded call must not read it as its own.
          if (!isCurrent()) return superseded();
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
    // Bumping the generation makes any in-flight playEpisode stale, so a cache
    // it is still writing gets removed by that call instead of outliving us.
    playbackGeneration += 1;
    const cacheDir = subtitleCacheDir;
    subtitleCacheDir = null;
    await Promise.all(queuedTrackPreparations);
    const queuedDirs = [...queuedSubtitleCacheDirs];
    queuedSubtitleCacheDirs.clear();
    await Promise.all([
      removeSubtitleCache(cacheDir, deps.subtitleCacheIo),
      ...queuedDirs.map((dir) => removeSubtitleCache(dir, deps.subtitleCacheIo)),
    ]);
  }

  return { playEpisode, prepareEpisode, appendEpisode, activateEpisode, discardEpisode, dispose };
}

function superseded(): AnimeBrowserPlayResult {
  return { ok: false, error: 'A newer episode replaced this playback.', quality: null };
}

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
