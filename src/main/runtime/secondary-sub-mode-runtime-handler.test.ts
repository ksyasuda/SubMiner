import assert from 'node:assert/strict';
import test from 'node:test';
import { createCycleSecondarySubModeRuntimeHandler } from './secondary-sub-mode-runtime-handler';

test('secondary sub mode runtime handler composes deps for runtime call', () => {
  const calls: string[] = [];
  const handleCycleSecondarySubMode = createCycleSecondarySubModeRuntimeHandler({
    cycleSecondarySubModeMainDeps: {
      getSecondarySubMode: () => 'hidden',
      setSecondarySubMode: () => calls.push('set-mode'),
      getLastSecondarySubToggleAtMs: () => 10,
      setLastSecondarySubToggleAtMs: () => calls.push('set-ts'),
      broadcastToOverlayWindows: (channel, mode) => calls.push(`broadcast:${channel}:${mode}`),
      showMpvOsd: (text) => calls.push(`osd:${text}`),
    },
    cycleSecondarySubMode: (deps) => {
      deps.setSecondarySubMode('romaji' as never);
      deps.broadcastSecondarySubMode('romaji' as never);
      deps.showMpvOsd('romaji');
    },
  });

  handleCycleSecondarySubMode();
  assert.deepEqual(calls, ['set-mode', 'broadcast:secondary-subtitle:mode:romaji', 'osd:romaji']);
});
