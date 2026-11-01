// Runs AnkiConnect's `sync` (AnkiWeb) after you mine cards. A burst of mining becomes one sync:
// it waits until no card has been mined for `delayMs`, and syncs right away when playback ends.
// Also registers a `sync` action for a manual sync.

const DEFAULT_SETTINGS = {
  delayMs: 30_000,
  syncOnEnd: true,
  timeoutMs: 60_000,
  notify: true,
};

const END_REASONS = new Set(['eof', 'stop', 'quit']);

function activate(ctx) {
  const settings = { ...DEFAULT_SETTINGS, ...ctx.settings };
  // Cards mined since the last sync started.
  let dirty = false;
  let syncing = false;
  let cancelTimer = null;

  function schedule() {
    cancelTimer?.();
    cancelTimer = ctx.after(settings.delayMs, flush);
  }

  async function flush() {
    cancelTimer?.();
    cancelTimer = null;
    // A sync already running picks up new cards when it finishes.
    if (!dirty || syncing) return;
    dirty = false;
    syncing = true;
    try {
      await ctx.anki.invoke('sync', {}, { timeoutMs: settings.timeoutMs });
      ctx.log.info('Synced Anki.');
      if (settings.notify) ctx.notify('Anki synced');
    } catch (error) {
      ctx.log.warn(`Anki sync failed: ${error.message}`);
      ctx.notify(`Anki sync failed: ${error.message}`, { variant: 'error' });
    } finally {
      syncing = false;
      if (dirty) schedule();
    }
  }

  ctx.anki.on('card-mined', () => {
    dirty = true;
    schedule();
  });

  ctx.mpv.on('end-file', ({ reason }) => {
    if (settings.syncOnEnd && END_REASONS.has(reason)) return flush();
  });

  ctx.actions.register('sync', () => {
    dirty = true;
    return flush();
  });
}

module.exports = { activate };
