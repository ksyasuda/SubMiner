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

test('focused mpv window follows the frontmost mpv app signal', () => {
  const focusHelperIndex = source.indexOf('private func isFocusedMpvWindow');
  assert.notEqual(focusHelperIndex, -1);

  const nextFunctionIndex = source.indexOf('\nprivate func ', focusHelperIndex + 1);
  const focusHelperBody = source.slice(focusHelperIndex, nextFunctionIndex);

  assert.ok(
    focusHelperBody.includes('frontmost.pid == ownerPid'),
    'matching frontmost PID should mark the mpv window focused',
  );
  assert.ok(
    focusHelperBody.includes('frontmost.isMpv && windowHasTargetSocket(ownerPid)'),
    'frontmost mpv app should mark the target mpv window focused even when PIDs differ',
  );
  assert.ok(
    source.includes('focused: isFocusedMpvWindow(ownerPid: windowPid, frontmost: frontmost)'),
    'Accessibility path should use the shared focused mpv helper',
  );
  assert.ok(
    source.includes('focused: isFocusedMpvWindow(ownerPid: ownerPid, frontmost: frontmost)'),
    'CoreGraphics path should use the shared focused mpv helper',
  );
});

test('frontmost mpv app emits active state when geometry lookup misses', () => {
  assert.ok(
    /case\s+\.active:/.test(source),
    'helper should expose an active state without window geometry',
  );
  assert.ok(
    source.includes('if windowHasTargetSocket(frontmost.pid)'),
    'active state should still accept a matching target socket when available',
  );
  assert.ok(
    source.includes('return targetMpvSocketPath != nil'),
    'active state should preserve frontmost mpv even if command-line socket detection fails',
  );
  assert.ok(
    source.includes('return .active'),
    'lookup should preserve active mpv state after geometry lookup misses',
  );
  assert.ok(source.includes('print("active")'), 'active state should be printed for the tracker');
});

test('frontmost non-mpv app emits inactive state when geometry lookup misses', () => {
  assert.ok(
    /case\s+\.inactive:/.test(source),
    'helper should expose an inactive state without window geometry',
  );
  assert.ok(
    source.includes('if frontmost != nil'),
    'helper should distinguish a known non-mpv frontmost app from an unknown miss',
  );
  assert.ok(source.includes('return .inactive'), 'known non-mpv focus should return inactive');
  assert.ok(
    source.includes('print("inactive")'),
    'inactive state should be printed for the tracker',
  );
});
