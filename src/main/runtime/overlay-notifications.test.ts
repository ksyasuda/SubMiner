import assert from 'node:assert/strict';
import test from 'node:test';

import type { OverlayNotificationPayload } from '../../types';
import {
  createConfiguredNotificationHandler,
  createShowOverlayNotificationHandler,
  createShowLogicalOsdHandler,
} from './overlay-notifications.js';

test('logical OSD prefers overlay notifications when available', () => {
  const calls: string[] = [];
  const showLogicalOsd = createShowLogicalOsdHandler({
    showOverlayNotification: (payload) => {
      calls.push(`overlay:${payload.kind}:${payload.message}`);
      return true;
    },
    showMpvOsd: (message) => {
      calls.push(`osd:${message}`);
    },
  });

  const result = showLogicalOsd({
    kind: 'success',
    message: 'Subtitle annotations loaded',
  });

  assert.equal(result, 'overlay');
  assert.deepEqual(calls, ['overlay:success:Subtitle annotations loaded']);
});

test('logical OSD normalizes spinner-frame loading messages for overlay notifications', () => {
  const calls: string[] = [];
  const showLogicalOsd = createShowLogicalOsdHandler({
    showOverlayNotification: (payload) => {
      calls.push(`overlay:${payload.kind}:${payload.message}`);
      return true;
    },
    showMpvOsd: (message) => {
      calls.push(`osd:${message}`);
    },
  });

  showLogicalOsd('Loading subtitle annotations |');

  assert.deepEqual(calls, ['overlay:loading:Loading subtitle annotations']);
});

test('logical OSD falls back to mpv OSD when overlay notifications are unavailable', () => {
  const calls: string[] = [];
  const showLogicalOsd = createShowLogicalOsdHandler({
    showOverlayNotification: () => false,
    showMpvOsd: (message) => {
      calls.push(`osd:${message}`);
    },
  });

  const result = showLogicalOsd({
    kind: 'loading',
    message: 'Loading subtitle annotations',
  });

  assert.equal(result, 'osd');
  assert.deepEqual(calls, ['osd:Loading subtitle annotations']);
});

test('overlay notifications send to the visible overlay when it is enabled and visible', () => {
  const calls: string[] = [];
  const showOverlayNotification = createShowOverlayNotificationHandler({
    isOverlayRuntimeInitialized: () => true,
    getVisibleOverlayVisible: () => true,
    getMainWindow: () =>
      ({
        isDestroyed: () => false,
        isVisible: () => true,
        webContents: {
          isLoading: () => false,
          getURL: () => 'file:///overlay.html',
          once: () => undefined,
          send: (channel: string, payload: OverlayNotificationPayload) => {
            calls.push(`${channel}:${payload.kind}:${payload.message}`);
          },
        },
      }) as never,
    notificationChannel: 'overlay:notification',
  });

  const shown = showOverlayNotification({
    kind: 'success',
    message: 'Overlay ready',
  });

  assert.equal(shown, true);
  assert.deepEqual(calls, ['overlay:notification:success:Overlay ready']);
});

test('overlay notifications return false when the visible overlay is hidden', () => {
  const showOverlayNotification = createShowOverlayNotificationHandler({
    isOverlayRuntimeInitialized: () => true,
    getVisibleOverlayVisible: () => false,
    getMainWindow: () => null,
    notificationChannel: 'overlay:notification',
  });

  assert.equal(
    showOverlayNotification({
      kind: 'info',
      message: 'Hidden overlay fallback',
    }),
    false,
  );
});

test('overlay notifications queue until renderer load finishes when window is visible but still loading', () => {
  const calls: string[] = [];
  let didFinishLoad: (() => void) | null = null;
  const showOverlayNotification = createShowOverlayNotificationHandler({
    isOverlayRuntimeInitialized: () => true,
    getVisibleOverlayVisible: () => true,
    getMainWindow: () =>
      ({
        isDestroyed: () => false,
        isVisible: () => true,
        webContents: {
          isLoading: () => true,
          getURL: () => 'about:blank',
          once: (_event: string, listener: () => void) => {
            didFinishLoad = listener;
          },
          send: (channel: string, payload: OverlayNotificationPayload) => {
            calls.push(`${channel}:${payload.kind}:${payload.message}`);
          },
        },
      }) as never,
    notificationChannel: 'overlay:notification',
  });

  const shown = showOverlayNotification({
    kind: 'loading',
    message: 'Loading subtitle annotations',
  });

  assert.equal(shown, true);
  assert.deepEqual(calls, []);

  if (didFinishLoad === null) {
    throw new Error('expected did-finish-load listener');
  }
  const runDidFinishLoad: () => void = didFinishLoad;
  runDidFinishLoad();

  assert.deepEqual(calls, ['overlay:notification:loading:Loading subtitle annotations']);
});

test('configured notifications treat osd as the overlay-or-fallback channel', () => {
  const calls: string[] = [];
  const showConfiguredNotification = createConfiguredNotificationHandler({
    getNotificationType: () => 'osd',
    showLogicalOsd: (payload) => {
      calls.push(`logical:${payload.kind}:${payload.message}`);
      return 'overlay';
    },
    showDesktopNotification: () => {
      calls.push('desktop');
    },
  });

  showConfiguredNotification('SubMiner', {
    kind: 'info',
    message: 'Config reload failed',
  });

  assert.deepEqual(calls, ['logical:info:Config reload failed']);
});

test('configured notifications send both logical OSD and desktop notifications for both', () => {
  const calls: string[] = [];
  const showConfiguredNotification = createConfiguredNotificationHandler({
    getNotificationType: () => 'both',
    showLogicalOsd: (payload) => {
      calls.push(`logical:${payload.kind}:${payload.message}`);
      return 'overlay';
    },
    showDesktopNotification: (title, options) => {
      calls.push(`desktop:${title}:${options.body}`);
    },
  });

  showConfiguredNotification('SubMiner', {
    kind: 'warning',
    message: 'Restart required',
  });

  assert.deepEqual(calls, [
    'logical:warning:Restart required',
    'desktop:SubMiner:Restart required',
  ]);
});

test('configured notifications suppress all channels for none', () => {
  const calls: string[] = [];
  const showConfiguredNotification = createConfiguredNotificationHandler({
    getNotificationType: () => 'none',
    showLogicalOsd: () => {
      calls.push('logical');
      return 'overlay';
    },
    showDesktopNotification: () => {
      calls.push('desktop');
    },
  });

  showConfiguredNotification('SubMiner', {
    kind: 'error',
    message: 'Overlay failed to start',
  });

  assert.deepEqual(calls, []);
});

test('configured notifications default missing types to logical OSD only', () => {
  const calls: string[] = [];
  const showConfiguredNotification = createConfiguredNotificationHandler({
    getNotificationType: () => undefined,
    showLogicalOsd: (payload) => {
      calls.push(`logical:${payload.kind}:${payload.message}`);
      return 'overlay';
    },
    showDesktopNotification: () => {
      calls.push('desktop');
    },
  });

  const payload: OverlayNotificationPayload = {
    kind: 'success',
    message: 'Card updated',
  };
  showConfiguredNotification('SubMiner', payload);

  assert.deepEqual(calls, ['logical:success:Card updated']);
});
