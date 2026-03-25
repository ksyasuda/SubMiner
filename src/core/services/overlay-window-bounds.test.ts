import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeOverlayWindowBoundsForPlatform } from './overlay-window-bounds';

test('normalizeOverlayWindowBoundsForPlatform returns original geometry outside Windows', () => {
  const geometry = { x: 150, y: 90, width: 1200, height: 675 };
  assert.deepEqual(normalizeOverlayWindowBoundsForPlatform(geometry, 'linux', null), geometry);
});

test('normalizeOverlayWindowBoundsForPlatform returns original geometry on Windows when screen is unavailable', () => {
  const geometry = { x: 150, y: 90, width: 1200, height: 675 };
  assert.deepEqual(normalizeOverlayWindowBoundsForPlatform(geometry, 'win32', null), geometry);
});

test('normalizeOverlayWindowBoundsForPlatform converts Windows physical pixels to DIP', () => {
  assert.deepEqual(
    normalizeOverlayWindowBoundsForPlatform(
      {
        x: 150,
        y: 75,
        width: 1920,
        height: 1080,
      },
      'win32',
      {
        screenToDipRect: (_window, rect) => ({
          x: Math.round(rect.x / 1.5),
          y: Math.round(rect.y / 1.5),
          width: Math.round(rect.width / 1.5),
          height: Math.round(rect.height / 1.5),
        }),
      },
    ),
    {
      x: 100,
      y: 50,
      width: 1280,
      height: 720,
    },
  );
});
