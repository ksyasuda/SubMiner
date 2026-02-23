import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildOverlayShortcutsRuntimeMainDepsHandler } from './overlay-shortcuts-runtime-main-deps';

test('overlay shortcuts runtime main deps builder maps lifecycle and action callbacks', async () => {
  const calls: string[] = [];
  let shortcutsRegistered = false;
  const deps = createBuildOverlayShortcutsRuntimeMainDepsHandler({
    getConfiguredShortcuts: () => ({ copySubtitle: 's' } as never),
    getShortcutsRegistered: () => shortcutsRegistered,
    setShortcutsRegistered: (registered) => {
      shortcutsRegistered = registered;
      calls.push(`registered:${registered}`);
    },
    isOverlayRuntimeInitialized: () => true,
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    openRuntimeOptionsPalette: () => calls.push('runtime-options'),
    openJimaku: () => calls.push('jimaku'),
    markAudioCard: async () => {
      calls.push('mark-audio');
    },
    copySubtitleMultiple: (timeoutMs) => calls.push(`copy-multi:${timeoutMs}`),
    copySubtitle: () => calls.push('copy'),
    toggleSecondarySubMode: () => calls.push('toggle-sub'),
    updateLastCardFromClipboard: async () => {
      calls.push('update-last-card');
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
    mineSentenceMultiple: (timeoutMs) => calls.push(`mine-multi:${timeoutMs}`),
    cancelPendingMultiCopy: () => calls.push('cancel-copy'),
    cancelPendingMineSentenceMultiple: () => calls.push('cancel-mine'),
  })();

  assert.equal(deps.isOverlayRuntimeInitialized(), true);
  assert.equal(deps.getShortcutsRegistered(), false);
  deps.setShortcutsRegistered(true);
  assert.equal(shortcutsRegistered, true);
  deps.showMpvOsd('x');
  deps.openRuntimeOptionsPalette();
  deps.openJimaku();
  await deps.markAudioCard();
  deps.copySubtitleMultiple(5000);
  deps.copySubtitle();
  deps.toggleSecondarySubMode();
  await deps.updateLastCardFromClipboard();
  await deps.triggerFieldGrouping();
  await deps.triggerSubsyncFromConfig();
  await deps.mineSentenceCard();
  deps.mineSentenceMultiple(3000);
  deps.cancelPendingMultiCopy();
  deps.cancelPendingMineSentenceMultiple();
  assert.deepEqual(calls, [
    'registered:true',
    'osd:x',
    'runtime-options',
    'jimaku',
    'mark-audio',
    'copy-multi:5000',
    'copy',
    'toggle-sub',
    'update-last-card',
    'field-grouping',
    'subsync',
    'mine',
    'mine-multi:3000',
    'cancel-copy',
    'cancel-mine',
  ]);
});
