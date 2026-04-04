import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildMpvCommandFromIpcRuntimeMainDepsHandler } from './ipc-mpv-command-main-deps';

test('ipc mpv command main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildMpvCommandFromIpcRuntimeMainDepsHandler({
    triggerSubsyncFromConfig: () => calls.push('subsync'),
    openRuntimeOptionsPalette: () => calls.push('palette'),
    openJimaku: () => calls.push('jimaku'),
    openYoutubeTrackPicker: () => {
      calls.push('youtube-picker');
    },
    openPlaylistBrowser: () => {
      calls.push('playlist-browser');
    },
    cycleRuntimeOption: () => ({ ok: false as const, error: 'x' }),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    replayCurrentSubtitle: () => calls.push('replay'),
    playNextSubtitle: () => calls.push('next'),
    shiftSubDelayToAdjacentSubtitle: async (direction) => {
      calls.push(`shift:${direction}`);
    },
    sendMpvCommand: (command) => calls.push(`cmd:${command.join(':')}`),
    getMpvClient: () => ({ connected: true, requestProperty: async () => null }),
    isMpvConnected: () => true,
    hasRuntimeOptionsManager: () => false,
  })();

  deps.triggerSubsyncFromConfig();
  deps.openRuntimeOptionsPalette();
  deps.openJimaku();
  void deps.openYoutubeTrackPicker();
  void deps.openPlaylistBrowser();
  assert.deepEqual(deps.cycleRuntimeOption('anki.nPlusOneMatchMode', 1), { ok: false, error: 'x' });
  deps.showMpvOsd('hello');
  deps.replayCurrentSubtitle();
  deps.playNextSubtitle();
  void deps.shiftSubDelayToAdjacentSubtitle('next');
  deps.sendMpvCommand(['show-text', 'ok']);
  assert.equal(typeof deps.getMpvClient()?.requestProperty, 'function');
  assert.equal(deps.isMpvConnected(), true);
  assert.equal(deps.hasRuntimeOptionsManager(), false);
  assert.deepEqual(calls, [
    'subsync',
    'palette',
    'jimaku',
    'youtube-picker',
    'playlist-browser',
    'osd:hello',
    'replay',
    'next',
    'shift:next',
    'cmd:show-text:ok',
  ]);
});
