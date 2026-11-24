/** What the episode playlist needs from a prepared Jellyfin item. */
export type JellyfinPlaylistEntry = {
  itemId: string;
  /** Set for episodes; the playlist only queues neighbors within a series. */
  seriesId: string | null;
  /** The URL handed to mpv, which mpv reports back as `path` when it starts the entry. */
  playbackUrl: string;
  /** Per-file mpv options as name/value pairs, applied whenever mpv starts this entry. */
  fileOptions: ReadonlyArray<readonly [name: string, value: string]>;
  /** Shown in mpv's playlist, which otherwise lists the stream URL and its API key. */
  title: string;
};

/**
 * How entries get into mpv's playlist with their title and per-file options.
 * - `loadfile`: options ride on `loadfile`, and the title is set through `playlist/N/title`,
 *   which mpv accepts from 0.42.
 * - `loadlist`: an in-memory m3u carries the title (`#EXTINF`), which older mpv shows too, and
 *   the SubMiner mpv plugin applies the options from its `on_load` hook.
 */
export type MpvEntryLoadMode = 'loadfile' | 'loadlist';

/** Set by the mpv plugin when its `on_load` hook applies `ENTRY_OPTIONS_SCRIPT_MESSAGE` options. */
export const PLUGIN_ENTRY_OPTIONS_USER_DATA = 'user-data/subminer/playlist-entry-options';

/** `<path> <name> <value> ...`: per-file options the mpv plugin applies when `path` loads. */
export const ENTRY_OPTIONS_SCRIPT_MESSAGE = 'subminer-playlist-entry-options';

/** Picks `loadlist` only where it is needed (mpv before 0.42) and the plugin can apply options. */
export function chooseMpvEntryLoadMode(
  mpvVersion: unknown,
  pluginAppliesEntryOptions: boolean,
): MpvEntryLoadMode {
  const match = /(\d+)\.(\d+)/.exec(String(mpvVersion ?? ''));
  if (!match || !pluginAppliesEntryOptions) return 'loadfile';
  const [major, minor] = [Number(match[1]), Number(match[2])];
  return major === 0 && minor < 42 ? 'loadlist' : 'loadfile';
}

// mpv's `%<bytes>%value` quoting takes any value verbatim, commas and colons included.
function formatLoadfileOptions(options: JellyfinPlaylistEntry['fileOptions']): string {
  return options
    .map(([name, value]) =>
      /^[\w.-]*$/.test(value)
        ? `${name}=${value}`
        : `${name}=%${Buffer.byteLength(value, 'utf8')}%${value}`,
    )
    .join(',');
}

type MpvPlaylistPlacement = { kind: 'replace' } | { kind: 'insert-at'; index: number };

/** The mpv commands that load `entry` with its title and options at `placement`. */
export function buildMpvEntryLoadCommands(
  entry: JellyfinPlaylistEntry,
  mode: MpvEntryLoadMode,
  placement: MpvPlaylistPlacement,
): Array<Array<string | number>> {
  if (mode === 'loadlist') {
    const title = entry.title.replace(/[\r\n]+/g, ' ');
    const m3u = `memory://#EXTM3U\n#EXTINF:0,${title}\n${entry.playbackUrl}\n`;
    return [
      [
        'script-message',
        ENTRY_OPTIONS_SCRIPT_MESSAGE,
        entry.playbackUrl,
        ...entry.fileOptions.flat(),
      ],
      placement.kind === 'replace'
        ? ['loadlist', m3u, 'replace']
        : ['loadlist', m3u, 'insert-at', placement.index],
    ];
  }
  const options = formatLoadfileOptions(entry.fileOptions);
  if (placement.kind === 'replace')
    return [['loadfile', entry.playbackUrl, 'replace', -1, options]];
  return [
    ['loadfile', entry.playbackUrl, 'insert-at', placement.index, options],
    // Sent while the entry is still at that index.
    ['set_property', `playlist/${placement.index}/title`, entry.title],
  ];
}

export type JellyfinEpisodePlaylistSettings = {
  enabled: boolean;
  /** Most entries kept in mpv's playlist, the current one included. Raised to at least 3. */
  maxEpisodes: number;
};

/**
 * A window of 3 or more always holds the current episode's immediate neighbors (it only shifts
 * at a series edge, where that neighbor does not exist). Removals then stay at least two steps
 * from the playing entry, so a queue run cannot remove the episode a user just stepped back to.
 */
const MIN_WINDOW_EPISODES = 3;

export type JellyfinQueuedEpisode<Context, Entry> = {
  context: Context;
  entry: Entry;
};

/** Tells the mpv plugin which playlist entries the app starts itself, so it skips its own startup. */
export const MANAGED_PLAYLIST_ENTRIES_SCRIPT_MESSAGE = 'subminer-managed-playlist-entries';

type MpvPlaylistItem = { filename: string; current: boolean };

function parseMpvPlaylist(raw: unknown): MpvPlaylistItem[] {
  if (!Array.isArray(raw)) return [];
  return raw.map((item: unknown) => {
    const record = item && typeof item === 'object' ? (item as Record<string, unknown>) : {};
    return {
      filename: typeof record.filename === 'string' ? record.filename : '',
      current: record.current === true,
    };
  });
}

/**
 * Picks up to `maxEpisodes` ids around `currentId`, centered on it (an extra slot goes forward)
 * and shifted inward at either end of the series so the window stays full.
 */
export function selectEpisodeWindow(
  episodeIds: readonly string[],
  currentId: string,
  maxEpisodes: number,
): { previousIds: string[]; nextIds: string[] } {
  const index = episodeIds.indexOf(currentId);
  const size = Math.min(episodeIds.length, Math.max(1, Math.floor(maxEpisodes)));
  if (index < 0 || size <= 1) return { previousIds: [], nextIds: [] };
  const start = Math.min(Math.max(0, index - Math.floor((size - 1) / 2)), episodeIds.length - size);
  return {
    previousIds: episodeIds.slice(start, index),
    nextIds: episodeIds.slice(index + 1, start + size),
  };
}

/**
 * Keeps the episodes around the playing Jellyfin item in mpv's playlist, so mpv's own playlist
 * navigation (and end-of-file advance) moves through the series. mpv starts those entries by
 * itself; `find` lets the playback handler recognize them from mpv's `path`.
 */
export function createJellyfinEpisodePlaylist<Context, Entry extends JellyfinPlaylistEntry>(deps: {
  getSettings: () => JellyfinEpisodePlaylistSettings;
  listSeriesEpisodeIds: (context: Context, seriesId: string) => Promise<string[]>;
  prepareEntry: (context: Context, itemId: string) => Promise<Entry>;
  readMpvPlaylist: () => Promise<unknown>;
  /** The mode of the explicit playback this playlist follows. */
  getEntryLoadMode: () => MpvEntryLoadMode;
  sendMpvCommand: (command: Array<string | number>) => void;
  logDebug: (message: string, error?: unknown) => void;
}) {
  let queued = new Map<string, JellyfinQueuedEpisode<Context, Entry>>();
  // Episode order per series, fetched once per explicit playback and reused while navigating.
  let seriesEpisodeIds = new Map<string, Promise<string[]>>();
  // Bumped by every reset and queue run, so a slow run never edits a playlist that moved on.
  let generation = 0;

  function getSeriesEpisodeIds(context: Context, seriesId: string): Promise<string[]> {
    let ids = seriesEpisodeIds.get(seriesId);
    if (!ids) {
      ids = deps.listSeriesEpisodeIds(context, seriesId);
      seriesEpisodeIds.set(seriesId, ids);
      // A failed lookup is retried on the next queue run.
      ids.catch(() => seriesEpisodeIds.delete(seriesId));
    }
    return ids;
  }

  async function prepareEntries(context: Context, itemIds: string[]): Promise<Map<string, Entry>> {
    const prepared = new Map<string, Entry>();
    await Promise.all(
      itemIds.map((itemId) =>
        deps.prepareEntry(context, itemId).then(
          (entry) => {
            prepared.set(itemId, entry);
          },
          (error: unknown) => {
            deps.logDebug(`Failed to prepare Jellyfin episode ${itemId} for the playlist`, error);
          },
        ),
      ),
    );
    return prepared;
  }

  async function queueNeighbors(context: Context, current: Entry): Promise<void> {
    const run = ++generation;
    const settings = deps.getSettings();
    let windowIds = [current.itemId];
    if (settings.enabled && current.seriesId) {
      const episodeIds = await getSeriesEpisodeIds(context, current.seriesId);
      const window = selectEpisodeWindow(
        episodeIds,
        current.itemId,
        Math.max(MIN_WINDOW_EPISODES, settings.maxEpisodes),
      );
      windowIds = [...window.previousIds, current.itemId, ...window.nextIds];
    }
    // Entries already queued stay where they are, so only the missing ones are resolved.
    const queuedItemIds = new Set([...queued.values()].map(({ entry }) => entry.itemId));
    const prepared = await prepareEntries(
      context,
      windowIds.filter((itemId) => itemId !== current.itemId && !queuedItemIds.has(itemId)),
    );
    if (run !== generation) return;

    const playlist = parseMpvPlaylist(await deps.readMpvPlaylist());
    if (run !== generation) return;
    const currentIndex = playlist.findIndex((item) => item.current);
    // mpv already left this item, so the run that follows it owns the playlist.
    if (playlist[currentIndex]?.filename !== current.playbackUrl) return;

    // mpv may have moved on since the read, so this only touches entries outside the window
    // and never removes one the user could have just navigated to. `model` mirrors mpv's
    // playlist by item id (null for entries the app did not queue) as the commands edit it.
    const model = playlist.map((item) =>
      item.filename === current.playbackUrl
        ? current.itemId
        : (queued.get(item.filename)?.entry.itemId ?? null),
    );
    for (let index = model.length - 1; index >= 0; index -= 1) {
      const itemId = model[index];
      if (itemId === null || itemId === undefined || windowIds.includes(itemId)) continue;
      deps.sendMpvCommand(['playlist-remove', index]);
      model.splice(index, 1);
    }

    const nextQueued = new Map<string, JellyfinQueuedEpisode<Context, Entry>>();
    for (const queuedEpisode of queued.values()) {
      if (model.includes(queuedEpisode.entry.itemId)) {
        nextQueued.set(queuedEpisode.entry.playbackUrl, queuedEpisode);
      }
    }
    nextQueued.set(current.playbackUrl, { context, entry: current });
    const currentAt = windowIds.indexOf(current.itemId);
    const nextIds = windowIds.slice(currentAt + 1);
    const previousIds = windowIds.slice(0, currentAt).reverse();
    // Every prepared entry is missing from mpv's playlist, so each one gets inserted below.
    for (const itemId of [...nextIds, ...previousIds]) {
      const entry = prepared.get(itemId);
      if (entry) nextQueued.set(entry.playbackUrl, { context, entry });
    }
    queued = nextQueued;
    // Sent before the inserts, so the plugin already knows an entry when mpv starts it.
    deps.sendMpvCommand([
      'script-message',
      MANAGED_PLAYLIST_ENTRIES_SCRIPT_MESSAGE,
      ...queued.keys(),
    ]);

    const mode = deps.getEntryLoadMode();
    const insertEntry = (entry: Entry, index: number) => {
      for (const command of buildMpvEntryLoadCommands(entry, mode, { kind: 'insert-at', index })) {
        deps.sendMpvCommand(command);
      }
      model.splice(index, 0, entry.itemId);
    };
    // A `loadlist` entry already has its title from the m3u.
    if (mode === 'loadfile') {
      deps.sendMpvCommand([
        'set_property',
        `playlist/${model.indexOf(current.itemId)}/title`,
        current.title,
      ]);
    }
    // Walk outward from the current entry, inserting each missing episode beside its
    // neighbor in the window.
    let anchor = model.indexOf(current.itemId);
    for (const itemId of nextIds) {
      const entry = prepared.get(itemId);
      if (model.includes(itemId)) anchor = model.indexOf(itemId);
      else if (entry) insertEntry(entry, ++anchor);
    }
    anchor = model.indexOf(current.itemId);
    for (const itemId of previousIds) {
      const entry = prepared.get(itemId);
      if (model.includes(itemId)) anchor = model.indexOf(itemId);
      else if (entry) insertEntry(entry, anchor);
    }
  }

  return {
    /** Forgets queued entries and cancels queue runs in flight, for a new explicit playback. */
    reset(): void {
      generation += 1;
      queued = new Map();
      seriesEpisodeIds = new Map();
    },
    /** The queued episode whose playback URL mpv reports as `path`, if any. */
    find(path: string): JellyfinQueuedEpisode<Context, Entry> | null {
      return queued.get(path) ?? null;
    },
    /** Fills mpv's playlist with the episodes around `current`, mpv's current entry. */
    queueNeighbors: (context: Context, current: Entry): Promise<void> =>
      queueNeighbors(context, current).catch((error: unknown) => {
        deps.logDebug('Failed to queue adjacent Jellyfin episodes', error);
      }),
  };
}
