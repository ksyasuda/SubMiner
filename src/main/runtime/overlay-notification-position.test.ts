import assert from 'node:assert/strict';
import test from 'node:test';

import { withConfiguredOverlayNotificationPosition } from './overlay-notification-position';

test('overlay notification payloads inherit configured overlay position', () => {
  assert.deepEqual(
    withConfiguredOverlayNotificationPosition(
      { title: 'SubMiner', body: 'Ready' },
      { notifications: { overlayPosition: 'top' } },
    ),
    { title: 'SubMiner', body: 'Ready', position: 'top' },
  );
});

test('overlay notification payload position can override configured position', () => {
  assert.deepEqual(
    withConfiguredOverlayNotificationPosition(
      { title: 'SubMiner', body: 'Ready', position: 'top-left' },
      { notifications: { overlayPosition: 'top-right' } },
    ),
    { title: 'SubMiner', body: 'Ready', position: 'top-left' },
  );
});
