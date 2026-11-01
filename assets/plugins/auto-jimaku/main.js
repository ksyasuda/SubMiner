// Opens the Jimaku subtitle search when a file starts without a Japanese subtitle track.
// It waits a few seconds after the file loads so external subtitles mpv or SubMiner add count.

const DEFAULT_SETTINGS = {
  // Built-in action to run, e.g. `openTsukihime` or `openSubtitleSelection`.
  action: 'openJimaku',
  // Track `lang` codes that count as Japanese. Track titles with "Japanese" or 日本語 count too.
  languages: ['ja', 'jp', 'jpn'],
  delayMs: 3_000,
  // Also check http(s) streams; Jimaku can only search by name, which streams rarely have.
  streams: false,
};

const JAPANESE_TITLE = /japanese|日本語/i;

function activate(ctx) {
  const settings = { ...DEFAULT_SETTINGS, ...ctx.settings };
  const languages = new Set(settings.languages.map((lang) => lang.toLowerCase()));
  let cancelCheck = null;
  let playbackId = null;

  function hasJapaneseSubtitle(trackList) {
    return trackList.some(
      (track) =>
        track?.type === 'sub' &&
        (languages.has(String(track.lang ?? '').toLowerCase()) ||
          JAPANESE_TITLE.test(String(track.title ?? ''))),
    );
  }

  async function check(media) {
    cancelCheck = null;
    const trackList = await ctx.mpv.getProperty('track-list');
    if (ctx.media.current()?.playbackId !== media.playbackId || !Array.isArray(trackList)) return;
    if (hasJapaneseSubtitle(trackList)) return;
    ctx.log.info(`No Japanese subtitles for ${media.path}; running ${settings.action}.`);
    await ctx.actions.run(settings.action);
  }

  ctx.media.on('change', (media) => {
    if (media?.playbackId === playbackId) return;
    playbackId = media?.playbackId ?? null;
    cancelCheck?.();
    cancelCheck = null;
    if (!media) return;
    if (!settings.streams && /^[a-z]+:\/\//i.test(media.path)) return;
    cancelCheck = ctx.after(settings.delayMs, () => check(media));
  });
}

module.exports = { activate };
