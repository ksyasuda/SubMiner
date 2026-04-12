import test from 'node:test';
import assert from 'node:assert/strict';
import { findWindowsMpvTargetWindowHandle } from './windows-helper';
import type { MpvPollResult } from './win32';

test('findWindowsMpvTargetWindowHandle prefers the focused mpv window', () => {
  const result: MpvPollResult = {
    matches: [
      {
        hwnd: 111,
        bounds: { x: 0, y: 0, width: 1280, height: 720 },
        area: 1280 * 720,
        isForeground: false,
      },
      {
        hwnd: 222,
        bounds: { x: 10, y: 10, width: 800, height: 600 },
        area: 800 * 600,
        isForeground: true,
      },
    ],
    focusState: true,
    windowState: 'visible',
  };

  assert.equal(findWindowsMpvTargetWindowHandle(result), 222);
});

test('findWindowsMpvTargetWindowHandle falls back to the largest visible mpv window', () => {
  const result: MpvPollResult = {
    matches: [
      {
        hwnd: 111,
        bounds: { x: 0, y: 0, width: 640, height: 360 },
        area: 640 * 360,
        isForeground: false,
      },
      {
        hwnd: 222,
        bounds: { x: 10, y: 10, width: 1920, height: 1080 },
        area: 1920 * 1080,
        isForeground: false,
      },
    ],
    focusState: false,
    windowState: 'visible',
  };

  assert.equal(findWindowsMpvTargetWindowHandle(result), 222);
});

test('findWindowsMpvTargetWindowHandle returns null when no mpv windows are visible', () => {
  const result: MpvPollResult = {
    matches: [],
    focusState: false,
    windowState: 'not-found',
  };

  assert.equal(findWindowsMpvTargetWindowHandle(result), null);
});
