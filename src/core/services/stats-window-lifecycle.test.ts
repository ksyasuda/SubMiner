import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createStatsWindowLayerSuspensionState,
  isStatsWindowLayerSuspended,
  resetStatsWindowLayerSuspension,
  restoreStatsWindowLayer,
  suspendStatsWindowLayer,
} from './stats-window-layer';

test('stats window layer suspension reset clears missed native dialog closes', () => {
  const state = createStatsWindowLayerSuspensionState();

  assert.equal(suspendStatsWindowLayer(state), true);
  assert.equal(suspendStatsWindowLayer(state), false);
  assert.equal(isStatsWindowLayerSuspended(state), true);

  resetStatsWindowLayerSuspension(state);

  assert.equal(isStatsWindowLayerSuspended(state), false);
  assert.equal(restoreStatsWindowLayer(state), false);
  assert.equal(suspendStatsWindowLayer(state), true);
});
