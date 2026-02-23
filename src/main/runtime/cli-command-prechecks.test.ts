import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleTexthookerOnlyModeTransitionHandler } from './cli-command-prechecks';

test('texthooker precheck no-ops when mode is disabled', () => {
  let warmups = 0;
  const handlePrecheck = createHandleTexthookerOnlyModeTransitionHandler({
    isTexthookerOnlyMode: () => false,
    setTexthookerOnlyMode: () => {},
    commandNeedsOverlayRuntime: () => true,
    startBackgroundWarmups: () => {
      warmups += 1;
    },
    logInfo: () => {},
  });

  handlePrecheck({ start: true, texthooker: false } as never);
  assert.equal(warmups, 0);
});

test('texthooker precheck disables mode and warms up on start command', () => {
  let mode = true;
  let warmups = 0;
  let logs = 0;
  const handlePrecheck = createHandleTexthookerOnlyModeTransitionHandler({
    isTexthookerOnlyMode: () => mode,
    setTexthookerOnlyMode: (enabled) => {
      mode = enabled;
    },
    commandNeedsOverlayRuntime: () => false,
    startBackgroundWarmups: () => {
      warmups += 1;
    },
    logInfo: () => {
      logs += 1;
    },
  });

  handlePrecheck({ start: true, texthooker: false } as never);
  assert.equal(mode, false);
  assert.equal(warmups, 1);
  assert.equal(logs, 1);
});

test('texthooker precheck no-ops for texthooker command', () => {
  let mode = true;
  const handlePrecheck = createHandleTexthookerOnlyModeTransitionHandler({
    isTexthookerOnlyMode: () => mode,
    setTexthookerOnlyMode: (enabled) => {
      mode = enabled;
    },
    commandNeedsOverlayRuntime: () => true,
    startBackgroundWarmups: () => {},
    logInfo: () => {},
  });

  handlePrecheck({ start: true, texthooker: true } as never);
  assert.equal(mode, true);
});
