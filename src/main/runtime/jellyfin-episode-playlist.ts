/** What the episode playlist needs from a prepared Jellyfin item. */
export type JellyfinPlaylistEntry = {
  itemId: string;
  /** Set for episodes; the playlist only queues neighbors within a series. */
  seriesId: string | null;
  /** The URL handed to mpv, which mpv reports back as `path` when it starts the entry. */
  playbackUrl: string;
  /** Per-file mpv options, applied whenever mpv starts this entry. */
  loadfileOptions: string;
  /** Shown in mpv's playlist, which otherwise lists the stream URL and its API key. */
  title: string;
};

export type JellyfinEpisodePlaylistSettings = {
  enabled: boolean;
  /** Most entries kept in mpv's playlist, the current one included. */
  maxEpisodes: number;
};

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

  async function prepareNeighbors(context: Context, itemIds: string[]): Promise<Entry[]> {
    const slots: Array<Entry | null> = itemIds.map(() => null);
    await Promise.all(
      itemIds.map((itemId, index) =>
        deps.prepareEntry(context, itemId).then(
          (entry) => {
            slots[index] = entry;
          },
          (error: unknown) => {
            deps.logDebug(`Failed to prepare Jellyfin episode ${itemId} for the playlist`, error);
          },
        ),
      ),
    );
    return slots.filter((entry): entry is Entry => entry !== null);
  }

  async function queueNeighbors(context: Context, current: Entry): Promise<void> {
    const run = ++generation;
    const settings = deps.getSettings();
    let previous: Entry[] = [];
    let next: Entry[] = [];
    if (settings.enabled && current.seriesId) {
      const episodeIds = await getSeriesEpisodeIds(context, current.seriesId);
      const window = selectEpisodeWindow(episodeIds, current.itemId, settings.maxEpisodes);
      [previous, next] = await Promise.all([
        prepareNeighbors(context, window.previousIds),
        prepareNeighbors(context, window.nextIds),
      ]);
    }
    if (run !== generation) return;

    const playlist = parseMpvPlaylist(await deps.readMpvPlaylist());
    if (run !== generation) return;
    const currentIndex = playlist.findIndex((item) => item.current);
    // mpv already left this item, so the run that follows it owns the playlist.
    if (playlist[currentIndex]?.filename !== current.playbackUrl) return;

    // Drop the entries queued for the last item, from the end so earlier indices stay valid.
    // Entries still in the window are re-added below with fresh resume positions.
    let currentAt = currentIndex;
    for (let index = playlist.length - 1; index >= 0; index -= 1) {
      if (index === currentIndex || !queued.has(playlist[index]!.filename)) continue;
      deps.sendMpvCommand(['playlist-remove', index]);
      if (index < currentIndex) currentAt -= 1;
    }

    // Each insert is followed by its title, sent while the entry is still at that index.
    const insertEntry = (entry: Entry, index: number) => {
      deps.sendMpvCommand([
        'loadfile',
        entry.playbackUrl,
        'insert-at',
        index,
        entry.loadfileOptions,
      ]);
      deps.sendMpvCommand(['set_property', `playlist/${index}/title`, entry.title]);
      queued.set(entry.playbackUrl, { context, entry });
    };
    queued = new Map([[current.playbackUrl, { context, entry: current }]]);
    deps.sendMpvCommand(['set_property', `playlist/${currentAt}/title`, current.title]);
    next.forEach((entry, offset) => insertEntry(entry, currentAt + 1 + offset));
    // Each insert before the current entry pushes it one slot later.
    previous.forEach((entry, offset) => insertEntry(entry, currentAt + offset));
    deps.sendMpvCommand([
      'script-message',
      MANAGED_PLAYLIST_ENTRIES_SCRIPT_MESSAGE,
      ...queued.keys(),
    ]);
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
