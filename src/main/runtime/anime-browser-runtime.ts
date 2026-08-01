import { AnimeBridgeClient } from '../../anime-bridge/bridge-client';
import { resolveStream } from '../../anime-bridge/headers';
import {
  parseAnimeStatus,
  resolveBridgeMediaUrl,
  routeHlsThroughProxy,
} from '../../anime-bridge/media-url';
import {
  watchPlaybackOutcome,
  type PlaybackEndFileEvent,
} from '../../anime-bridge/playback-outcome';
import {
  startStreamStripProxy,
  type StreamStripProxyHandle,
} from '../../anime-bridge/stream-strip-proxy';
import {
  buildPlaybackCommands,
  buildTrackCommands,
  selectPreferredStream,
} from '../../anime-bridge/mpv-playback';
import {
  listExtensionSources,
  readInstalledExtensions,
  toBridgeSource,
  toInstalledExtensionViews,
  type ExtensionSource,
  type InstalledExtension,
} from '../../anime-bridge/extension-store';
import {
  cacheSubtitleTracks,
  removeSubtitleCache,
  type SubtitleCacheIo,
} from '../../anime-bridge/subtitle-cache';
import { interleave, mapSourcesConcurrently } from '../../anime-bridge/multi-source-search';
import { startSidecar, type SidecarHandle } from '../../anime-bridge/sidecar-process';
import {
  fetchRepoCatalogue,
  isValidRepoUrl,
  type RepoExtension,
} from '../../anime-bridge/extension-repo';
import {
  installExtension,
  removeExtension as removeExtensionFile,
} from '../../anime-bridge/extension-installer';
import { PreferenceStore } from '../../anime-bridge/preference-store';
import { applyPreferenceValue, parsePreferences } from '../../anime-bridge/preferences';
import type { SourcePreferenceView } from '../../anime-bridge/preferences';
import type { BundleBinaries } from '../../anime-bridge/sidecar-bundle';
import type { InstallProgress } from './anime-bridge-installer';
import { ALL_SOURCES_ID } from '../../types/anime-browser';
import type {
  AnimeBrowserBridgeState,
  AnimeBrowserDetails,
  AnimeBrowserEntry,
  AnimeBrowserEpisode,
  AnimeBrowserPlayRequest,
  AnimeBrowserPlayResult,
  AnimeBrowserSearchResult,
  AnimeBrowserSearchUpdate,
  AnimeBrowserSnapshot,
  AvailableExtensionsResult,
  ExtensionLoadFailure,
} from '../../types/anime-browser';
import type { BridgeAnimePage } from '../../anime-bridge/types';

export interface AnimeBrowserRuntimeDeps {
  /** Where user-supplied Aniyomi extension APKs live. Read lazily so config edits apply. */
  extensionsDir: () => string;
  /** Configured repository index URLs. Empty unless the user added one. */
  repos: () => string[];
  /** Persists the repository list. Config stays the source of truth. */
  setRepos: (repos: string[]) => void;
  /** JSON file holding each source's saved preference values. */
  preferencesFile: string;
  ensureBinaries: (onProgress: (progress: InstallProgress) => void) => Promise<BundleBinaries>;
  /** Sends mpv an IPC command; same transport the Jellyfin path uses. */
  sendMpvCommand: (command: Array<string | number>) => void;
  /** Brings mpv up if it is not already connected. Resolves false on failure. */
  ensureMpvConnected: () => Promise<boolean>;
  /**
   * Subscribe to mpv end-file events; returns the unsubscribe. Together with
   * readMpvProperty this lets playEpisode confirm playback really started
   * instead of reporting success the moment the commands were written.
   */
  onPlaybackEndFile?: (listener: (event: PlaybackEndFileEvent) => void) => () => void;
  /** One-shot mpv property read; rejects while the property is unavailable. */
  readMpvProperty?: (name: string) => Promise<unknown>;
  showMpvOsd?: (message: string) => void;
  showVisibleOverlay?: () => void;
  /** Lets tests drive the pause between `loadfile` and the track commands. */
  wait?: (ms: number) => Promise<void>;
  /** Overrides the filesystem/network the subtitle cache uses. Tests only. */
  subtitleCacheIo?: SubtitleCacheIo;
  onBridgeState: (state: AnimeBrowserBridgeState) => void;
  /**
   * Streams per-source progress while a search invoke is pending. Optional so
   * a host that has no window to push to can leave it out.
   */
  onSearchUpdate?: (update: AnimeBrowserSearchUpdate) => void;
  preferredQuality?: () => string | undefined;
  log: (message: string) => void;
}

const IDLE_STATE: AnimeBrowserBridgeState = { stage: 'idle', progress: null, message: null };

/** How long to let `loadfile` settle before adding external tracks. */
const TRACK_ATTACH_DELAY_MS = 300;

export function createAnimeBrowserRuntime(deps: AnimeBrowserRuntimeDeps) {
  let bridgeState: AnimeBrowserBridgeState = IDLE_STATE;
  let sidecar: SidecarHandle | null = null;
  let stripProxy: StreamStripProxyHandle | null = null;
  let starting: Promise<SidecarHandle> | null = null;
  let extensions: InstalledExtension[] = [];
  let sources: ExtensionSource[] = [];
  let selectedSourceId: string | null = null;
  let loadFailures: ExtensionLoadFailure[] = [];
  // Monotonic; identifies the newest browse so stale ones stop emitting.
  let searchToken = 0;
  // Temp directory holding the playing episode's downloaded subtitles. Kept
  // until the next episode replaces it, because alass reads it mid-playback.
  let subtitleCacheDir: string | null = null;
  const preferenceStore = new PreferenceStore(deps.preferencesFile);
  const wait =
    deps.wait ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));

  /**
   * Drop the previous episode's downloaded subtitles.
   *
   * Deliberately not done at end-file: alass reads those files for as long as
   * the episode is up, so they only go once another one replaces them.
   */
  async function clearSubtitleCache(): Promise<void> {
    const previousDir = subtitleCacheDir;
    subtitleCacheDir = null;
    await removeSubtitleCache(previousDir, deps.subtitleCacheIo);
  }

  /**
   * Download the stream's subtitle tracks so mpv loads files instead of URLs.
   *
   * alass needs the reference track on disk, and the subsync picker rejects an
   * external track whose `external-filename` is not a real file, so a streamed
   * track cannot be used to retime anything.
   */
  async function cacheStreamSubtitles(stream: {
    headers: Record<string, string>;
    subtitles: Array<{ url: string; lang: string }>;
  }): Promise<Array<{ url: string; lang: string }>> {
    const cached = await cacheSubtitleTracks({
      tracks: stream.subtitles,
      headers: stream.headers,
      io: deps.subtitleCacheIo,
      log: deps.log,
    });
    subtitleCacheDir = cached.dir;

    const localCount = cached.tracks.filter((track) => track.local).length;
    if (cached.tracks.length > 0) {
      deps.log(
        `[anime-browser] cached ${localCount}/${cached.tracks.length} subtitle track(s) to disk` +
          (cached.dir ? ` in ${cached.dir}` : ''),
      );
    }
    return cached.tracks.map((track) => ({ url: track.url, lang: track.lang }));
  }

  function setState(state: AnimeBrowserBridgeState): void {
    bridgeState = state;
    deps.onBridgeState(state);
  }

  async function sourceFor(sourceId: string) {
    const source = sources.find((candidate) => candidate.id === sourceId);
    if (!source) throw new Error('That source is no longer installed. Rescan and try again.');
    const extension = extensions.find((candidate) => candidate.file === source.file);
    if (!extension) throw new Error(`Extension file missing for ${source.name}.`);
    // Saved values ride along on every call; the extension is stateless per request.
    const saved = await preferenceStore.get(source.id);
    return { ...toBridgeSource(extension, source.id), preferences: saved };
  }

  /**
   * The live bridge, starting (or restarting) it first when there is none.
   * A bridge that died out from under the app — killed, crashed, stopped —
   * comes back on the next request instead of failing every call until the
   * app restarts.
   */
  async function bridge(): Promise<{ client: AnimeBridgeClient; baseUrl: string }> {
    if (!sidecar) await ensureBridge();
    if (!sidecar) {
      throw new Error(bridgeState.message ?? 'The anime bridge is not running.');
    }
    return { client: sidecar.client, baseUrl: sidecar.baseUrl };
  }

  async function startBridge(): Promise<SidecarHandle> {
    const binaries = await deps.ensureBinaries((progress) =>
      setState({ stage: progress.stage, progress: progress.progress, message: null }),
    );

    setState({ stage: 'starting', progress: null, message: null });
    const handle = await startSidecar({
      binaries,
      onLog: (line) => deps.log(`[anime-bridge] ${line}`),
    });
    sidecar = handle;

    // A deliberate stop detaches first (dispose nulls `sidecar` before
    // stopping, a restart replaces it), so reaching the body means the bridge
    // died out from under us and the next request should bring it back.
    handle.onExit(({ code, signal }) => {
      if (sidecar !== handle) return;
      sidecar = null;
      starting = null;
      const proxy = stripProxy;
      stripProxy = null;
      void proxy?.close();
      deps.log(
        `[anime-browser] bridge exited unexpectedly (code ${code}, signal ${signal}); ` +
          'it will restart on the next request',
      );
      setState({
        stage: 'idle',
        progress: null,
        message: 'The anime bridge stopped. It restarts on the next search.',
      });
    });

    try {
      const proxy = await startStreamStripProxy({
        upstreamOrigin: () => sidecar?.baseUrl ?? handle.baseUrl,
        log: deps.log,
      });
      // The bridge can die while the proxy is coming up, in which case onExit
      // already ran and cleared `stripProxy`; adopting this one would leak a
      // listening server pointed at a dead upstream.
      if (sidecar === handle) stripProxy = proxy;
      else void proxy.close();
    } catch (error) {
      // Playback still works for undisguised streams; log and carry on.
      deps.log(`[anime-browser] stream proxy failed to start: ${describeError(error)}`);
    }

    await scanExtensions(handle);
    return handle;
  }

  /**
   * Re-read the extensions directory and ask the bridge what each APK provides.
   * Called on start and after any install or removal.
   */
  async function scanExtensions(handle: SidecarHandle): Promise<void> {
    const directory = deps.extensionsDir();
    loadFailures = [];
    extensions = await readInstalledExtensions(directory);
    sources = await listExtensionSources(handle.client, extensions, (extension, error) => {
      const message = describeError(error);
      loadFailures.push({ pkg: extension.fallbackName, error: message });
      deps.log(`[anime-bridge] extension ${extension.fallbackName} failed to load: ${message}`);
    });

    // Keep the current selection if it survived the rescan. "All sources"
    // survives as long as anything is installed.
    const keptAll = selectedSourceId === ALL_SOURCES_ID && sources.length > 0;
    if (!keptAll && !sources.some((source) => source.id === selectedSourceId)) {
      selectedSourceId = sources[0]?.id ?? null;
    }

    setState({
      stage: 'ready',
      progress: null,
      message:
        sources.length === 0
          ? `No anime extensions installed. Add a repository or put .apk files in ${directory}.`
          : null,
    });
  }

  async function ensureBridge(): Promise<AnimeBrowserBridgeState> {
    if (sidecar) return bridgeState;
    // Collapse concurrent callers onto one start; the UI calls this eagerly.
    if (!starting) {
      starting = startBridge().catch((error: unknown) => {
        setState({ stage: 'failed', progress: null, message: describeError(error) });
        starting = null;
        throw error;
      });
    }
    try {
      await starting;
    } catch {
      return bridgeState;
    }
    return bridgeState;
  }

  async function installExtensionFrom(extension: RepoExtension): Promise<void> {
    await installExtension({ extensionsDir: deps.extensionsDir(), extension });
    if (sidecar) await scanExtensions(sidecar);
  }

  function toEntry(
    baseUrl: string,
    anime: { url?: string; title?: string; thumbnail_url?: string },
    source: ExtensionSource,
  ) {
    return {
      url: anime.url ?? '',
      title: anime.title ?? 'Untitled',
      thumbnailUrl: anime.thumbnail_url
        ? resolveBridgeMediaUrl(baseUrl, anime.thumbnail_url)
        : null,
      sourceId: source.id,
      sourceName: source.name,
    } satisfies AnimeBrowserEntry;
  }

  /** The one source a per-anime call must run against. */
  function requireSource(sourceId: string | null): ExtensionSource {
    const source = sources.find((candidate) => candidate.id === sourceId);
    if (!source) {
      throw new Error(
        sourceId === ALL_SOURCES_ID ? 'Pick a single source for this.' : 'Select a source first.',
      );
    }
    return source;
  }

  /**
   * Run a listing call against the selected source, or against every installed
   * source when "all sources" is selected.
   *
   * With one source a failure rejects, as it always has. Across all of them a
   * failure is reported alongside the sources that did answer, so one broken
   * extension cannot blank the grid.
   */
  async function browse(
    page: number,
    fetchPage: (
      source: Awaited<ReturnType<typeof sourceFor>>,
      page: number,
    ) => Promise<BridgeAnimePage>,
  ): Promise<AnimeBrowserSearchResult> {
    const { baseUrl } = await bridge();
    const targets =
      selectedSourceId === ALL_SOURCES_ID ? sources : [requireSource(selectedSourceId)];
    if (targets.length === 0) throw new Error('No sources are installed.');

    // Each source's answer is pushed the moment it lands, so a fast source is
    // on screen while a slow one is still resolving. Guarded by the token: a
    // superseded search stops emitting, and its remaining sources run out
    // quietly.
    const token = ++searchToken;
    const emit = (update: AnimeBrowserSearchUpdate): void => {
      if (token === searchToken) deps.onSearchUpdate?.(update);
    };
    emit({ kind: 'start', token, sourceCount: targets.length });

    const { results, failures } = await mapSourcesConcurrently(targets, async (source) => {
      try {
        const response = await fetchPage(await sourceFor(source.id), page);
        const entries = (response.animes ?? []).map((anime) => toEntry(baseUrl, anime, source));
        emit({ kind: 'result', token, sourceId: source.id, sourceName: source.name, entries });
        return { entries, hasNextPage: response.hasNextPage === true };
      } catch (error) {
        emit({
          kind: 'failure',
          token,
          failure: { sourceId: source.id, sourceName: source.name, error: describeError(error) },
        });
        throw error;
      }
    });

    emit({ kind: 'done', token });

    // A single-source browse has no other source to fall back on, so surface
    // the error the way a direct call would.
    if (targets.length === 1 && failures[0]) throw new Error(failures[0].error);

    return {
      entries: interleave(results.map((result) => result.entries)),
      hasNextPage: results.some((result) => result.hasNextPage),
      failures,
    };
  }

  return {
    getSnapshot(): AnimeBrowserSnapshot {
      return {
        bridge: bridgeState,
        sources: sources.map((source) => ({
          id: source.id,
          name: source.name,
          lang: source.lang,
          pkg: source.pkg,
        })),
        selectedSourceId,
        loadFailures,
        installed: toInstalledExtensionViews(extensions, sources, loadFailures),
        extensionsDir: deps.extensionsDir(),
        repos: deps.repos(),
      };
    },

    ensureBridge,

    /**
     * Extensions available from the configured repositories, annotated with
     * what is installed. Returns nothing when no repository is configured —
     * SubMiner never supplies one.
     */
    async listAvailableExtensions(): Promise<AvailableExtensionsResult> {
      const repos = deps.repos();
      if (repos.length === 0) return { extensions: [], failures: [] };

      const catalogue = await fetchRepoCatalogue(repos);
      const installedPkgs = new Set(extensions.map((extension) => extension.fallbackName));
      return {
        extensions: catalogue.extensions.map((extension) => ({
          pkg: extension.pkg,
          name: extension.name,
          lang: extension.lang,
          version: extension.version,
          nsfw: extension.nsfw,
          repoUrl: extension.repoUrl,
          sourceNames: extension.sourceNames,
          installed: installedPkgs.has(extension.pkg),
        })),
        failures: catalogue.failures,
      };
    },

    /** Download an extension by package name, then rescan. */
    async installExtension(pkg: string): Promise<void> {
      const repos = deps.repos();
      if (repos.length === 0) throw new Error('No extension repository is configured.');

      const catalogue = await fetchRepoCatalogue(repos);
      const match = catalogue.extensions.find((candidate) => candidate.pkg === pkg);
      if (!match) throw new Error(`${pkg} is not offered by any configured repository.`);

      await installExtensionFrom(match);
    },

    /** Remove an installed extension, then rescan. */
    async removeExtension(pkg: string): Promise<void> {
      await removeExtensionFile(deps.extensionsDir(), pkg);
      await preferenceStore.clear(pkg).catch(() => undefined);
      if (sidecar) await scanExtensions(sidecar);
    },

    /**
     * Add a repository index URL. Rejected unless it is an https index URL, so
     * a typo surfaces immediately instead of failing later at fetch time.
     */
    addRepo(url: string): void {
      const trimmed = url.trim();
      if (!isValidRepoUrl(trimmed)) {
        throw new Error('A repository URL must be https and point at a .json index file.');
      }
      const repos = deps.repos();
      if (repos.includes(trimmed)) return;
      deps.setRepos([...repos, trimmed]);
    },

    removeRepo(url: string): void {
      deps.setRepos(deps.repos().filter((candidate) => candidate !== url));
    },

    /** Re-read the extensions directory without restarting the bridge. */
    async rescanExtensions(): Promise<void> {
      if (sidecar) await scanExtensions(sidecar);
    },

    selectSource(sourceId: string): void {
      if (sourceId === ALL_SOURCES_ID && sources.length > 0) {
        selectedSourceId = ALL_SOURCES_ID;
        return;
      }
      if (sources.some((source) => source.id === sourceId)) selectedSourceId = sourceId;
    },

    /**
     * The extension's settings schema, merged with anything saved locally. The
     * extension is the source of truth for structure; saved values only supply
     * what it has no memory of between requests.
     */
    async getPreferences(sourceId: string): Promise<SourcePreferenceView[]> {
      const { client } = await bridge();
      const source = requireSource(sourceId);
      const schema = await client.getSourcePreferences(await sourceFor(source.id));
      return parsePreferences(schema);
    },

    /**
     * Persist one preference and hand the whole array back to the extension so
     * it can react (Jellyfin logs in and populates its library list here).
     */
    async setPreference(
      sourceId: string,
      key: string,
      value: string | string[] | boolean,
    ): Promise<SourcePreferenceView[]> {
      const { client } = await bridge();
      const source = await sourceFor(sourceId);

      // Start from the extension's own schema so saved values never go stale
      // against an updated extension.
      const current =
        source.preferences && source.preferences.length > 0
          ? source.preferences
          : await client.getSourcePreferences(source);

      const updated = applyPreferenceValue(current, key, value);
      await preferenceStore.set(sourceId, updated);

      const refreshed = await client.setSourcePreference({ ...source, preferences: updated }, key);
      if (refreshed.length > 0) await preferenceStore.set(sourceId, refreshed);
      return parsePreferences(refreshed.length > 0 ? refreshed : updated);
    },

    async search(query: string, page = 1): Promise<AnimeBrowserSearchResult> {
      const { client } = await bridge();
      return browse(page, (source, requestedPage) =>
        client.searchAnime(source, query, requestedPage),
      );
    },

    async getPopular(page = 1): Promise<AnimeBrowserSearchResult> {
      const { client } = await bridge();
      return browse(page, (source, requestedPage) => client.getPopularAnime(source, requestedPage));
    },

    async getDetails(animeUrl: string, sourceId?: string): Promise<AnimeBrowserDetails> {
      const { client, baseUrl } = await bridge();
      const source = requireSource(sourceId ?? selectedSourceId);
      const details = await client.getAnimeDetails(await sourceFor(source.id), animeUrl);
      return {
        ...toEntry(baseUrl, { ...details, url: details.url ?? animeUrl }, source),
        description: details.description ?? null,
        author: details.author ?? null,
        genres: details.genres ?? [],
        status: parseAnimeStatus(details.status),
      };
    },

    async getEpisodes(animeUrl: string, sourceId?: string): Promise<AnimeBrowserEpisode[]> {
      const { client } = await bridge();
      const source = requireSource(sourceId ?? selectedSourceId);
      const episodes = await client.getEpisodeList(await sourceFor(source.id), animeUrl);
      return episodes.map((episode) => ({
        url: episode.url ?? '',
        name: episode.name ?? 'Episode',
        number: typeof episode.episode_number === 'number' ? episode.episode_number : null,
        uploadedAt:
          typeof episode.date_upload === 'number' && episode.date_upload > 0
            ? episode.date_upload
            : null,
        scanlator: episode.scanlator ?? null,
      }));
    },

    async playEpisode(request: AnimeBrowserPlayRequest): Promise<AnimeBrowserPlayResult> {
      try {
        const { client, baseUrl } = await bridge();
        const videos = await client.getVideoList(
          await sourceFor(request.sourceId),
          request.episodeUrl,
        );
        const streams = videos
          .map((video) => resolveStream(video))
          .filter((stream): stream is NonNullable<typeof stream> => stream !== null)
          // External tracks come off the same loopback proxy as the video, so
          // they need the same rebase onto the port the bridge really uses.
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
        // HLS goes through the local strip proxy, which undoes fake-image
        // segment disguises mpv cannot decode around.
        const stream = stripProxy
          ? { ...selected, url: routeHlsThroughProxy(selected.url, baseUrl, stripProxy.origin) }
          : selected;

        if (!(await deps.ensureMpvConnected())) {
          return {
            ok: false,
            error: 'mpv is not running and could not be started.',
            quality: null,
          };
        }

        // Subscribed before loadfile so a fast failure cannot slip past it.
        const watch =
          deps.onPlaybackEndFile && deps.readMpvProperty
            ? watchPlaybackOutcome({
                onEndFile: deps.onPlaybackEndFile,
                readProperty: deps.readMpvProperty,
                wait,
              })
            : null;

        try {
          const title = `${request.animeTitle} — ${request.episodeName}`;
          for (const command of buildPlaybackCommands({ stream, title })) {
            deps.sendMpvCommand(command);
          }
          // The old episode's subtitles are dead the moment this one loads,
          // whether or not the new one brings any of its own.
          await clearSubtitleCache();

          if (stream.audios.length > 0 || stream.subtitles.length > 0) {
            deps.log(
              `[anime-browser] ${stream.audios.length} external audio, ` +
                `${stream.subtitles.length} external subtitle track(s)`,
            );
            // mpv attaches added tracks to the file that is loading, so give the
            // loadfile a moment to take effect first. Same pause the Jellyfin
            // subtitle preload uses. The download runs inside that pause.
            const [subtitles] = await Promise.all([
              cacheStreamSubtitles(stream),
              wait(TRACK_ATTACH_DELAY_MS),
            ]);
            for (const command of buildTrackCommands({ ...stream, subtitles })) {
              deps.sendMpvCommand(command);
            }
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
    },

    async dispose(): Promise<void> {
      const handle = sidecar;
      const proxy = stripProxy;
      const cacheDir = subtitleCacheDir;
      sidecar = null;
      stripProxy = null;
      starting = null;
      subtitleCacheDir = null;
      setState(IDLE_STATE);
      await removeSubtitleCache(cacheDir, deps.subtitleCacheIo);
      await proxy?.close();
      await handle?.stop();
    },
  };
}

export type AnimeBrowserRuntime = ReturnType<typeof createAnimeBrowserRuntime>;

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
