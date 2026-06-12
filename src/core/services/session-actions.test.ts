import assert from 'node:assert/strict';
import test from 'node:test';

import { dispatchSessionAction, type SessionActionExecutorDeps } from './session-actions';

function createDeps(overrides: Partial<SessionActionExecutorDeps> = {}) {
  const calls: string[] = [];
  const deps: SessionActionExecutorDeps = {
    toggleStatsOverlay: () => calls.push('stats'),
    toggleVisibleOverlay: () => calls.push('visible'),
    copyCurrentSubtitle: () => calls.push('copy'),
    copySubtitleCount: (count) => calls.push(`copy:${count}`),
    updateLastCardFromClipboard: async () => {
      calls.push('update');
    },
    triggerFieldGrouping: async () => {
      calls.push('field-grouping');
    },
    triggerSubsyncFromConfig: async () => {
      calls.push('subsync');
    },
    mineSentenceCard: async () => {
      calls.push('mine');
    },
    mineSentenceCount: (count) => calls.push(`mine:${count}`),
    toggleSecondarySub: () => calls.push('secondary'),
    toggleSubtitleSidebar: () => calls.push('sidebar'),
    toggleNotificationHistory: () => calls.push('notification-history'),
    markLastCardAsAudioCard: async () => {
      calls.push('audio');
    },
    markActiveVideoWatched: async () => {
      calls.push('mark-watched');
      return true;
    },
    openRuntimeOptionsPalette: () => calls.push('runtime-options'),
    openSessionHelp: () => calls.push('session-help'),
    openCharacterDictionaryManager: () => calls.push('character-dictionary-manager'),
    openControllerSelect: () => calls.push('controller-select'),
    openControllerDebug: () => calls.push('controller-debug'),
    openJimaku: () => calls.push('jimaku'),
    openYoutubeTrackPicker: () => {
      calls.push('youtube');
    },
    openPlaylistBrowser: () => {
      calls.push('playlist');
    },
    replayCurrentSubtitle: () => calls.push('replay'),
    playNextSubtitle: () => calls.push('play-next'),
    cycleRuntimeOption: () => ({ ok: true }),
    playNextPlaylistItem: () => calls.push('playlist-next'),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    ...overrides,
  };
  return { calls, deps };
}

test('dispatchSessionAction marks watched and advances playlist after success', async () => {
  const { calls, deps } = createDeps();

  await dispatchSessionAction({ actionId: 'markWatched' }, deps);

  assert.deepEqual(calls, ['mark-watched', 'osd:Marked as watched', 'playlist-next']);
});

test('dispatchSessionAction does not advance playlist when mark watched no-ops', async () => {
  const { calls, deps } = createDeps({
    markActiveVideoWatched: async () => {
      calls.push('mark-watched');
      return false;
    },
  });

  await dispatchSessionAction({ actionId: 'markWatched' }, deps);

  assert.deepEqual(calls, ['mark-watched']);
});

test('dispatchSessionAction opens the character dictionary manager', async () => {
  const { calls, deps } = createDeps();

  await dispatchSessionAction({ actionId: 'openCharacterDictionaryManager' }, deps);

  assert.deepEqual(calls, ['character-dictionary-manager']);
});
