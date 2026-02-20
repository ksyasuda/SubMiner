import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildCycleSecondarySubModeMainDepsHandler } from './secondary-sub-mode-main-deps';
import type { SecondarySubMode } from '../../types';

test('cycle secondary sub mode main deps builder maps state and broadcasts with channel', () => {
  const calls: string[] = [];
  let mode: SecondarySubMode = 'hover';
  let lastToggleAt = 100;
  const deps = createBuildCycleSecondarySubModeMainDepsHandler({
    getSecondarySubMode: () => mode,
    setSecondarySubMode: (nextMode) => {
      mode = nextMode;
      calls.push(`set-mode:${nextMode}`);
    },
    getLastSecondarySubToggleAtMs: () => lastToggleAt,
    setLastSecondarySubToggleAtMs: (timestampMs) => {
      lastToggleAt = timestampMs;
      calls.push(`set-ts:${timestampMs}`);
    },
    broadcastToOverlayWindows: (channel, nextMode) => calls.push(`broadcast:${channel}:${nextMode}`),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
  })();

  assert.equal(deps.getSecondarySubMode(), 'hover');
  assert.equal(deps.getLastSecondarySubToggleAtMs(), 100);
  deps.setSecondarySubMode('visible');
  deps.setLastSecondarySubToggleAtMs(200);
  deps.broadcastSecondarySubMode('visible');
  deps.showMpvOsd('Secondary subtitle: visible');
  assert.equal(mode, 'visible');
  assert.equal(lastToggleAt, 200);
  assert.deepEqual(calls, [
    'set-mode:visible',
    'set-ts:200',
    'broadcast:secondary-subtitle:mode:visible',
    'osd:Secondary subtitle: visible',
  ]);
});
