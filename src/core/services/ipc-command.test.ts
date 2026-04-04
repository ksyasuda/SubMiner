import assert from 'node:assert/strict';
import test from 'node:test';
import { handleMpvCommandFromIpc } from './ipc-command';

function createOptions(overrides: Partial<Parameters<typeof handleMpvCommandFromIpc>[1]> = {}) {
  const calls: string[] = [];
  const sentCommands: (string | number)[][] = [];
  const osd: string[] = [];
  const options: Parameters<typeof handleMpvCommandFromIpc>[1] = {
    specialCommands: {
      SUBSYNC_TRIGGER: '__subsync-trigger',
      RUNTIME_OPTIONS_OPEN: '__runtime-options-open',
      JIMAKU_OPEN: '__jimaku-open',
      RUNTIME_OPTION_CYCLE_PREFIX: '__runtime-option-cycle:',
      REPLAY_SUBTITLE: '__replay-subtitle',
      PLAY_NEXT_SUBTITLE: '__play-next-subtitle',
      SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START: '__sub-delay-next-line',
      SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START: '__sub-delay-prev-line',
      YOUTUBE_PICKER_OPEN: '__youtube-picker-open',
      PLAYLIST_BROWSER_OPEN: '__playlist-browser-open',
    },
    triggerSubsyncFromConfig: () => {
      calls.push('subsync');
    },
    openRuntimeOptionsPalette: () => {
      calls.push('runtime-options');
    },
    openJimaku: () => {
      calls.push('jimaku');
    },
    openYoutubeTrackPicker: () => {
      calls.push('youtube-picker');
    },
    openPlaylistBrowser: () => {
      calls.push('playlist-browser');
    },
    runtimeOptionsCycle: () => ({ ok: true }),
    showMpvOsd: (text) => {
      osd.push(text);
    },
    mpvReplaySubtitle: () => {
      calls.push('replay');
    },
    mpvPlayNextSubtitle: () => {
      calls.push('next');
    },
    shiftSubDelayToAdjacentSubtitle: async (direction) => {
      calls.push(`shift:${direction}`);
    },
    mpvSendCommand: (command) => {
      sentCommands.push(command);
    },
    resolveProxyCommandOsd: async () => null,
    isMpvConnected: () => true,
    hasRuntimeOptionsManager: () => true,
    ...overrides,
  };
  return { options, calls, sentCommands, osd };
}

test('handleMpvCommandFromIpc forwards regular mpv commands', () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['cycle', 'pause'], options);
  assert.deepEqual(sentCommands, [['cycle', 'pause']]);
  assert.deepEqual(osd, []);
});

test('handleMpvCommandFromIpc emits osd for subtitle position keybinding proxies', async () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['add', 'sub-pos', 1], options);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(sentCommands, [['add', 'sub-pos', 1]]);
  assert.deepEqual(osd, ['Subtitle position: ${sub-pos}']);
});

test('handleMpvCommandFromIpc emits resolved osd for primary subtitle track keybinding proxies', async () => {
  const { options, sentCommands, osd } = createOptions({
    resolveProxyCommandOsd: async () => 'Subtitle track: Internal #3 - Japanese (active)',
  });
  handleMpvCommandFromIpc(['cycle', 'sid'], options);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(sentCommands, [['cycle', 'sid']]);
  assert.deepEqual(osd, ['Subtitle track: Internal #3 - Japanese (active)']);
});

test('handleMpvCommandFromIpc emits resolved osd for secondary subtitle track keybinding proxies', async () => {
  const { options, sentCommands, osd } = createOptions({
    resolveProxyCommandOsd: async () =>
      'Secondary subtitle track: External #8 - English Commentary',
  });
  handleMpvCommandFromIpc(['set_property', 'secondary-sid', 'auto'], options);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(sentCommands, [['set_property', 'secondary-sid', 'auto']]);
  assert.deepEqual(osd, ['Secondary subtitle track: External #8 - English Commentary']);
});

test('handleMpvCommandFromIpc emits osd for subtitle delay keybinding proxies', async () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['add', 'sub-delay', 0.1], options);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(sentCommands, [['add', 'sub-delay', 0.1]]);
  assert.deepEqual(osd, ['Subtitle delay: ${sub-delay}']);
});

test('handleMpvCommandFromIpc dispatches special subtitle-delay shift command', () => {
  const { options, calls, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['__sub-delay-next-line'], options);
  assert.deepEqual(calls, ['shift:next']);
  assert.deepEqual(sentCommands, []);
  assert.deepEqual(osd, []);
});

test('handleMpvCommandFromIpc dispatches special youtube picker open command', () => {
  const { options, calls, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['__youtube-picker-open'], options);
  assert.deepEqual(calls, ['youtube-picker']);
  assert.deepEqual(sentCommands, []);
  assert.deepEqual(osd, []);
});

test('handleMpvCommandFromIpc dispatches special jimaku open command', () => {
  const { options, calls, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['__jimaku-open'], options);
  assert.deepEqual(calls, ['jimaku']);
  assert.deepEqual(sentCommands, []);
  assert.deepEqual(osd, []);
});

test('handleMpvCommandFromIpc dispatches special playlist browser open command', async () => {
  const { options, calls, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['__playlist-browser-open'], options);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(calls, ['playlist-browser']);
  assert.deepEqual(sentCommands, []);
  assert.deepEqual(osd, []);
});

test('handleMpvCommandFromIpc surfaces playlist browser open rejections via mpv osd', async () => {
  const { options, osd } = createOptions({
    openPlaylistBrowser: async () => {
      throw new Error('overlay failed');
    },
  });

  handleMpvCommandFromIpc(['__playlist-browser-open'], options);
  await new Promise((resolve) => setImmediate(resolve));

  assert.deepEqual(osd, ['Playlist browser failed: overlay failed']);
});

test('handleMpvCommandFromIpc does not forward commands while disconnected', () => {
  const { options, sentCommands, osd } = createOptions({
    isMpvConnected: () => false,
  });
  handleMpvCommandFromIpc(['add', 'sub-pos', 1], options);
  assert.deepEqual(sentCommands, []);
  assert.deepEqual(osd, []);
});
