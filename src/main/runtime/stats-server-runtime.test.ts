import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isSelfOwnedBackgroundStatsDaemonState,
  shouldClearAppStateStatsServerOnStop,
} from './stats-server-runtime';

test('detects self-owned background stats daemon state', () => {
  assert.equal(
    isSelfOwnedBackgroundStatsDaemonState({ pid: process.pid, port: 6969, startedAtMs: 1 }),
    true,
  );
});

test('stats server app-state reference should be cleared after private server stop', () => {
  assert.equal(shouldClearAppStateStatsServerOnStop({ hadStatsServer: true }), true);
});
