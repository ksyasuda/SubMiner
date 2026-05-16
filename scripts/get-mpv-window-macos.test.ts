import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const source = readFileSync('scripts/get-mpv-window-macos.swift', 'utf8');

test('minimized Accessibility windows are validated by PID and socket before reporting minimized', () => {
  const minimizedAssignmentIndex = source.indexOf('foundMinimizedTargetWindow = true');
  assert.notEqual(minimizedAssignmentIndex, -1);

  const loopStartIndex = source.lastIndexOf('for window in windows', minimizedAssignmentIndex);
  assert.notEqual(loopStartIndex, -1);

  const pidExtractionIndex = source.indexOf(
    'AXUIElementGetPid(window, &windowPid)',
    loopStartIndex,
  );
  const appPidMatchIndex = source.indexOf('windowPid != app.processIdentifier', loopStartIndex);
  const socketCheckIndex = source.indexOf('if !windowHasTargetSocket(windowPid)', loopStartIndex);

  assert.ok(
    pidExtractionIndex > loopStartIndex && pidExtractionIndex < minimizedAssignmentIndex,
    'window PID must be extracted before accepting a minimized window',
  );
  assert.ok(
    appPidMatchIndex > pidExtractionIndex && appPidMatchIndex < minimizedAssignmentIndex,
    'window PID must match the owning app before accepting a minimized window',
  );
  assert.ok(
    socketCheckIndex > appPidMatchIndex && socketCheckIndex < minimizedAssignmentIndex,
    'target socket must be validated before accepting a minimized window',
  );
});
