import { AnimeBridgeClient } from '../../anime-bridge/bridge-client';
import { parseAnimeStatus, resolveBridgeMediaUrl } from '../../anime-bridge/media-url';
import {
  startStreamStripProxy,
  type StreamStripProxyHandle,
} from '../../anime-bridge/stream-strip-proxy';
import {
  listExtensionSources,
  readInstalledExtensions,
  toBridgeSource,
  toInstalledExtensionViews,
  type ExtensionSource,
  type InstalledExtension,
} from '../../anime-bridge/extension-store';
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
import {
  buildAnimeStreamMetadata,
  buildAnimeStreamStatsPath,
} from '../../anime-bridge/episode-metadata';
import { PreferenceStore } from '../../anime-bridge/preference-store';
import { applyPreferenceValue, parsePreferences } from '../../anime-bridge/preferences';
import type { SourcePreferenceView } from '../../anime-bridge/preferences';
import { ALL_SOURCES_ID } from '../../types/anime-browser';
import type {
  AnimeBrowserBridgeInstall,
  AnimeBrowserBridgeState,
  AnimeBrowserDetails,
  AnimeBrowserEntry,
  AnimeBrowserEpisode,
  AnimeBrowserEpisodeWatchState,
  AnimeBrowserPlayRequest,
  AnimeBrowserPlayResult,
  AnimeBrowserQueueState,
  AnimeBrowserSetWatchedRequest,
  AnimeBrowserWatchStateRequest,
  AnimeBrowserSearchResult,
  AnimeBrowserSearchUpdate,
  AnimeBrowserSnapshot,
  AvailableExtensionsResult,
  ExtensionLoadFailure,
} from '../../types/anime-browser';
import type { BridgeAnimePage, BridgePreference } from '../../anime-bridge/types';
import { findExtensionUpdates, hasExtensionUpdate } from '../../shared/extension-updates';
import { createAnimeBrowserPlayback } from './anime-browser-playback';
import { createAnimeBrowserQueue } from './anime-browser-queue';
import type { AnimeBrowserRuntimeDeps } from './anime-browser-runtime-deps';
export type { AnimeBrowserRuntimeDeps } from './anime-browser-runtime-deps';

const IDLE_STATE: AnimeBrowserBridgeState = {
  stage: 'idle',
  progress: null,
  message: null,
  install: null,
};

/** Stage, progress and message; `install` is carried across every state change. */
type BridgeStateChange = Omit<AnimeBrowserBridgeState, 'install'>;

export function createAnimeBrowserRuntime(deps: AnimeBrowserRuntimeDeps) {
  let bridgeState: AnimeBrowserBridgeState = IDLE_STATE;
  let install: AnimeBrowserBridgeInstall | null = null;
  let sidecar: SidecarHandle | null = null;
  let stripProxy: StreamStripProxyHandle | null = null;
  let starting: Promise<SidecarHandle> | null = null;
  let updating: Promise<AnimeBrowserBridgeState> | null = null;
  let extensionMutationTail = Promise.resolve();
  let extensions: InstalledExtension[] = [];
  let sources: ExtensionSource[] = [];
  let loadFailures: ExtensionLoadFailure[] = [];
  const browserSessions = new Map<
    string,
    { selectedSourceId: string | null; searchToken: number }
  >();
  const preferenceStore = new PreferenceStore(deps.preferencesFile);

  function withExtensionMutation<T>(operation: () => Promise<T>): Promise<T> {
    const result = extensionMutationTail.then(operation);
    extensionMutationTail = result.then(
      () => undefined,
      () => undefined,
    );
    return result;
  }

  function getBrowserSession(sessionId = 'default') {
    const existing = browserSessions.get(sessionId);
    if (existing) return existing;
    const created = { selectedSourceId: sources[0]?.id ?? null, searchToken: 0 };
    browserSessions.set(sessionId, created);
    return created;
  }

  function setState(change: BridgeStateChange): void {
    bridgeState = { ...change, install };
    deps.onBridgeState(bridgeState);
  }

  async function sourceFor(sourceId: string) {
    const source = sources.find((candidate) => candidate.id === sourceId);
    if (!source) throw new Error('That source is no longer installed. Rescan and try again.');
    const extension = extensions.find((candidate) => candidate.file === source.file);
    if (!extension) throw new Error(`Extension file missing for ${source.name}.`);
    // Saved values ride along on every call; the extension is stateless per request.
    const saved = await preferenceStore.get(source.pkg, source.bridgeId);
    return { ...toBridgeSource(extension, source.bridgeId), preferences: saved };
  }

  function overlaySavedPreferences(
    schema: BridgePreference[],
    saved: BridgePreference[],
  ): BridgePreference[] {
    let merged = schema;
    for (const preference of parsePreferences(saved)) {
      merged = applyPreferenceValue(merged, preference.key, preference.value);
    }
    return merged;
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
    const resolved = await deps.ensureBinaries((progress) =>
      setState({ stage: progress.stage, progress: progress.progress, message: null }),
    );
    install = {
      origin: resolved.origin,
      version: resolved.version,
      dir: resolved.dir,
      updateAvailable: resolved.updateAvailable,
    };
    deps.log(
      `[anime-browser] bridge ${resolved.version ?? 'unknown version'} (${resolved.origin}) ` +
        `from ${resolved.dir}`,
    );

    setState({ stage: 'starting', progress: null, message: null });
    const handle = await (deps.startSidecar ?? startSidecar)({
      binaries: resolved,
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
      const proxy = await (deps.startStreamStripProxy ?? startStreamStripProxy)({
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

    await withExtensionMutation(() => scanExtensions(handle));
    void checkForBridgeUpdate(handle);
    return handle;
  }

  /**
   * Ask upstream whether a managed install is behind, after the bridge is up
   * so a slow or failed GitHub call never delays a search. The answer lands
   * in `install.updateAvailable` and is re-broadcast on the current state.
   */
  async function checkForBridgeUpdate(handle: SidecarHandle): Promise<void> {
    if (install === null || install.origin !== 'managed') return;
    try {
      const latest = await deps.checkBridgeUpdate(install);
      // The bridge may have been restarted or updated while we waited.
      if (sidecar !== handle || install === null || latest === install.updateAvailable) return;
      install = { ...install, updateAvailable: latest };
      if (latest !== null) deps.log(`[anime-browser] bridge update available: ${latest}`);
      setState({
        stage: bridgeState.stage,
        progress: bridgeState.progress,
        message: bridgeState.message,
      });
    } catch (error) {
      deps.log(`[anime-browser] bridge update check failed: ${describeError(error)}`);
    }
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

    // Each open browser keeps its own source selection. Reconcile all of them
    // after a rescan without making one renderer change another renderer's UI.
    for (const session of browserSessions.values()) {
      const keptAll = session.selectedSourceId === ALL_SOURCES_ID && sources.length > 0;
      if (!keptAll && !sources.some((source) => source.id === session.selectedSourceId)) {
        session.selectedSourceId = sources[0]?.id ?? null;
      }
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

  /** Stop the bridge and its proxy on purpose, without disturbing playback state. */
  async function stopBridge(): Promise<void> {
    const pendingStart = starting;
    starting = null;
    try {
      await pendingStart;
    } catch {
      // A failed start has no sidecar to stop.
    }
    const handle = sidecar;
    const proxy = stripProxy;
    sidecar = null;
    stripProxy = null;
    await proxy?.close();
    await handle?.stop();
  }

  async function ensureBridge(): Promise<AnimeBrowserBridgeState> {
    // A request that lands while an update is swapping directories would start
    // the old bridge out of a tree that is about to be deleted.
    if (updating) await updating;
    return startIfNeeded();
  }

  async function startIfNeeded(): Promise<AnimeBrowserBridgeState> {
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

  /**
   * Move a managed install to the newest release: download beside it while
   * the old bridge keeps serving, then stop, swap, and restart. A failed
   * download leaves the old bridge running; a failed swap is reported as a
   * failed start, which the next request retries.
   */
  function updateBridge(): Promise<AnimeBrowserBridgeState> {
    if (updating) return updating;
    if (install && install.origin !== 'managed') {
      return Promise.reject(
        new Error(`The bridge in ${install.dir} is managed outside SubMiner; update it there.`),
      );
    }
    updating = (async () => {
      let staged;
      try {
        staged = await deps.stageBridgeUpdate((progress) =>
          setState({ stage: progress.stage, progress: progress.progress, message: null }),
        );
      } catch (error) {
        setState({
          stage: sidecar ? 'ready' : 'failed',
          progress: null,
          message: `Bridge update failed: ${describeError(error)}`,
        });
        return bridgeState;
      }
      await stopBridge();
      try {
        await staged.commit();
      } catch (error) {
        setState({
          stage: 'failed',
          progress: null,
          message: `Bridge update failed: ${describeError(error)}`,
        });
        return bridgeState;
      }
      // Re-resolves the install from disk and re-checks upstream once it is up.
      return startIfNeeded();
    })().finally(() => {
      updating = null;
    });
    return updating;
  }

  async function downloadExtension(extension: RepoExtension): Promise<void> {
    await installExtension({ extensionsDir: deps.extensionsDir(), extension });
  }

  async function installExtensionFrom(extension: RepoExtension): Promise<void> {
    await downloadExtension(extension);
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
    sessionId: string,
    fetchPage: (
      source: Awaited<ReturnType<typeof sourceFor>>,
      page: number,
    ) => Promise<BridgeAnimePage>,
  ): Promise<AnimeBrowserSearchResult> {
    const { baseUrl } = await bridge();
    const session = getBrowserSession(sessionId);
    const targets =
      session.selectedSourceId === ALL_SOURCES_ID
        ? sources
        : [requireSource(session.selectedSourceId)];
    if (targets.length === 0) throw new Error('No sources are installed.');

    // Each source's answer is pushed the moment it lands, so a fast source is
    // on screen while a slow one is still resolving. Guarded by the token: a
    // superseded search stops emitting, and its remaining sources run out
    // quietly.
    const token = ++session.searchToken;
    const emit = (update: AnimeBrowserSearchUpdate): void => {
      if (token === session.searchToken) deps.onSearchUpdate?.(update, sessionId);
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

  /**
   * Which of these episodes have already been watched.
   *
   * The stats database is the only store: playback records every streamed
   * episode under the same derived path, and marks it watched once a session
   * runs far enough. With tracking disabled there is no history to read, so
   * every episode comes back unwatched rather than the call failing.
   *
   * A closure rather than a method, so `setWatched` can reuse it without
   * depending on how the runtime object was called.
   */
  async function getWatchState(
    request: AnimeBrowserWatchStateRequest,
  ): Promise<AnimeBrowserEpisodeWatchState[]> {
    const episodeUrls = request.episodeUrls.filter((url) => url.length > 0);
    if (episodeUrls.length === 0 || !deps.getWatchState) return [];

    const statsPaths = new Map(
      episodeUrls.map((episodeUrl) => [
        episodeUrl,
        buildAnimeStreamStatsPath(request.sourceId, request.animeUrl, episodeUrl),
      ]),
    );

    try {
      const state = await deps.getWatchState([...statsPaths.values()]);
      const watchState: AnimeBrowserEpisodeWatchState[] = [];
      for (const [episodeUrl, statsPath] of statsPaths) {
        const entry = state.get(statsPath);
        if (!entry) continue;
        watchState.push({
          episodeUrl,
          watched: entry.watched,
          lastWatchedMs: entry.lastWatchedMs,
          sessionCount: entry.sessionCount,
        });
      }
      return watchState;
    } catch (error) {
      // Watch marks are decoration on a list that is already usable; a stats
      // read that fails must not take the episode list down with it.
      deps.log(`[anime-browser] watch state lookup failed: ${describeError(error)}`);
      return [];
    }
  }

  const playback = createAnimeBrowserPlayback({
    deps,
    bridge,
    sourceFor,
    stripProxy: () => stripProxy,
  });

  const queue = createAnimeBrowserQueue({
    prepareEpisode: (request) => playback.prepareEpisode(request),
    appendEpisode: (prepared) => playback.appendEpisode(prepared),
    activateEpisode: (prepared) => playback.activateEpisode(prepared),
    discardEpisode: (prepared) => playback.discardEpisode(prepared),
    armNextEpisode: () =>
      deps.sendMpvCommand(['script-message', 'subminer-managed-subtitles-loading']),
    onPlaybackPathChange: deps.onPlaybackPathChange,
    readMpvProperty: deps.readMpvProperty,
    sendMpvCommand: deps.sendMpvCommand,
    onQueueState: deps.onQueueState,
    showMpvOsd: deps.showMpvOsd,
    log: deps.log,
  });

  return {
    getSnapshot(sessionId = 'default'): AnimeBrowserSnapshot {
      const session = getBrowserSession(sessionId);
      return {
        bridge: bridgeState,
        sources: sources.map((source) => ({
          id: source.id,
          name: source.name,
          lang: source.lang,
          pkg: source.pkg,
        })),
        selectedSourceId: session.selectedSourceId,
        loadFailures,
        installed: toInstalledExtensionViews(extensions, sources, loadFailures),
        extensionsDir: deps.extensionsDir(),
        repos: deps.repos(),
      };
    },

    ensureBridge,
    updateBridge,

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
          versionCode: extension.versionCode,
          nsfw: extension.nsfw,
          repoUrl: extension.repoUrl,
          iconUrl: extension.iconUrl,
          sourceNames: extension.sourceNames,
          installed: installedPkgs.has(extension.pkg),
        })),
        failures: catalogue.failures,
      };
    },

    /** Download an extension by package name, then rescan. */
    installExtension(pkg: string): Promise<void> {
      return withExtensionMutation(async () => {
        const repos = deps.repos();
        if (repos.length === 0) throw new Error('No extension repository is configured.');

        const catalogue = await fetchRepoCatalogue(repos);
        const match = catalogue.extensions.find((candidate) => candidate.pkg === pkg);
        if (!match) throw new Error(`${pkg} is not offered by any configured repository.`);

        const current = extensions.find((candidate) => candidate.fallbackName === pkg);
        if (
          current &&
          current.versionCode !== null &&
          !hasExtensionUpdate(current.versionCode, match.versionCode)
        ) {
          return;
        }

        await installExtensionFrom(match);
      });
    },

    /** Download every strictly newer repository build, then rescan once. */
    updateAllExtensions(): Promise<number> {
      return withExtensionMutation(async () => {
        const repos = deps.repos();
        if (repos.length === 0) return 0;

        const catalogue = await fetchRepoCatalogue(repos);
        const installedVersions = extensions.map((extension) => ({
          pkg: extension.fallbackName,
          versionCode: extension.versionCode,
        }));
        const updates = findExtensionUpdates(installedVersions, catalogue.extensions);

        let installedCount = 0;
        try {
          for (const extension of updates) {
            await downloadExtension(extension);
            installedCount += 1;
          }
        } finally {
          if (installedCount > 0 && sidecar) await scanExtensions(sidecar);
        }
        return installedCount;
      });
    },

    /** Remove an installed extension, then rescan. */
    removeExtension(pkg: string): Promise<void> {
      return withExtensionMutation(async () => {
        await removeExtensionFile(deps.extensionsDir(), pkg);
        await preferenceStore.clear(pkg).catch(() => undefined);
        if (sidecar) await scanExtensions(sidecar);
      });
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
    rescanExtensions(): Promise<void> {
      return withExtensionMutation(async () => {
        if (sidecar) await scanExtensions(sidecar);
      });
    },

    selectSource(sourceId: string, sessionId = 'default'): void {
      const session = getBrowserSession(sessionId);
      if (sourceId === ALL_SOURCES_ID && sources.length > 0) {
        session.selectedSourceId = ALL_SOURCES_ID;
        return;
      }
      if (sources.some((source) => source.id === sourceId)) session.selectedSourceId = sourceId;
    },

    /**
     * The extension's settings schema, merged with anything saved locally. The
     * extension is the source of truth for structure; saved values only supply
     * what it has no memory of between requests.
     */
    async getPreferences(sourceId: string): Promise<SourcePreferenceView[]> {
      const { client } = await bridge();
      const source = requireSource(sourceId);
      const bridgeSource = await sourceFor(source.id);
      const schema = await client.getSourcePreferences(bridgeSource);
      return parsePreferences(overlaySavedPreferences(schema, bridgeSource.preferences ?? []));
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
      const extensionSource = requireSource(sourceId);
      const source = await sourceFor(extensionSource.id);

      // Start from the extension's own schema so saved values never go stale
      // against an updated extension.
      const schema = await client.getSourcePreferences(source);
      const current = overlaySavedPreferences(schema, source.preferences ?? []);

      const updated = applyPreferenceValue(current, key, value);
      await preferenceStore.set(extensionSource.pkg, extensionSource.bridgeId, updated);

      const refreshed = await client.setSourcePreference({ ...source, preferences: updated }, key);
      if (refreshed.length > 0) {
        await preferenceStore.set(extensionSource.pkg, extensionSource.bridgeId, refreshed);
      }
      return parsePreferences(refreshed.length > 0 ? refreshed : updated);
    },

    async search(
      query: string,
      page = 1,
      sessionId = 'default',
    ): Promise<AnimeBrowserSearchResult> {
      const { client } = await bridge();
      return browse(page, sessionId, (source, requestedPage) =>
        client.searchAnime(source, query, requestedPage),
      );
    },

    async getPopular(page = 1, sessionId = 'default'): Promise<AnimeBrowserSearchResult> {
      const { client } = await bridge();
      return browse(page, sessionId, (source, requestedPage) =>
        client.getPopularAnime(source, requestedPage),
      );
    },

    async getDetails(
      animeUrl: string,
      sourceId?: string,
      sessionId = 'default',
    ): Promise<AnimeBrowserDetails> {
      const { client, baseUrl } = await bridge();
      const source = requireSource(sourceId ?? getBrowserSession(sessionId).selectedSourceId);
      const details = await client.getAnimeDetails(await sourceFor(source.id), animeUrl);
      return {
        ...toEntry(baseUrl, { ...details, url: details.url ?? animeUrl }, source),
        description: details.description ?? null,
        author: details.author ?? null,
        genres: details.genres ?? [],
        status: parseAnimeStatus(details.status),
      };
    },

    async getEpisodes(
      animeUrl: string,
      sourceId?: string,
      sessionId = 'default',
    ): Promise<AnimeBrowserEpisode[]> {
      const { client } = await bridge();
      const source = requireSource(sourceId ?? getBrowserSession(sessionId).selectedSourceId);
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

    getWatchState,

    /**
     * Set or clear the watch mark on the given episodes, then report the state
     * that write left behind so the browser paints from the store rather than
     * from what it hoped happened.
     */
    async setWatched(
      request: AnimeBrowserSetWatchedRequest,
    ): Promise<AnimeBrowserEpisodeWatchState[]> {
      const episodes = request.episodes.filter((episode) => episode.episodeUrl.length > 0);
      if (episodes.length === 0 || !deps.setWatchState) return [];

      // The same metadata playback records, so an episode marked before it is
      // ever played still lands under the right series, season and episode.
      const marks = episodes.map((episode) => {
        const metadata = buildAnimeStreamMetadata({
          sourceId: request.sourceId,
          animeUrl: request.animeUrl,
          animeTitle: request.animeTitle,
          episodeUrl: episode.episodeUrl,
          episodeName: episode.episodeName,
          episodeNumber: episode.episodeNumber,
          // No stream was resolved: there is no media path to alias.
          mediaPath: '',
        });
        return {
          mediaPath: '',
          statsPath: metadata.statsPath,
          displayTitle: metadata.displayTitle,
          seriesTitle: metadata.seriesTitle,
          seasonNumber: metadata.seasonNumber,
          episodeNumber: metadata.episodeNumber,
        };
      });

      await deps.setWatchState(marks, request.watched);
      return getWatchState({
        sourceId: request.sourceId,
        animeUrl: request.animeUrl,
        episodeUrls: episodes.map((episode) => episode.episodeUrl),
      });
    },

    /**
     * Play now, replacing whatever mpv has. The queue is left standing and
     * re-armed behind this file, so an episode played by hand mid-queue is a
     * detour rather than a reset.
     */
    async playEpisode(request: AnimeBrowserPlayRequest): Promise<AnimeBrowserPlayResult> {
      const result = await playback.playEpisode(request);
      if (result.ok) queue.handlePlaybackStarted();
      return result;
    },

    async queueEpisode(request: AnimeBrowserPlayRequest): Promise<AnimeBrowserQueueState> {
      return await queue.enqueue(request);
    },

    async dequeueEpisode(sourceId: string, episodeUrl: string): Promise<AnimeBrowserQueueState> {
      return await queue.dequeue(sourceId, episodeUrl);
    },

    async clearQueue(): Promise<AnimeBrowserQueueState> {
      return await queue.clear();
    },

    getQueue(): AnimeBrowserQueueState {
      return queue.getState();
    },

    /**
     * Whether mpv has a file open. An mpv that is not running cannot answer,
     * and there is nothing playing in it either, so both read as false.
     */
    async isPlaying(): Promise<boolean> {
      if (!deps.readMpvProperty) return false;
      try {
        return (await deps.readMpvProperty('idle-active')) !== true;
      } catch {
        return false;
      }
    },

    releaseSession(sessionId: string): void {
      const session = browserSessions.get(sessionId);
      if (session) session.searchToken += 1;
      browserSessions.delete(sessionId);
    },

    async dispose(): Promise<void> {
      const stopping = stopBridge();
      setState(IDLE_STATE);
      await queue.dispose();
      await playback.dispose();
      await stopping;
    },
  };
}

export type AnimeBrowserRuntime = ReturnType<typeof createAnimeBrowserRuntime>;

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
