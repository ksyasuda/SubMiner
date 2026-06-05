import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getPlaybackFeedbackNotificationOptions,
  notifyConfiguredStatus,
} from './configured-status-notification';

test('notifyConfiguredStatus routes both to overlay and system without osd', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Subsync: choose engine and source', {
    getNotificationType: () => 'both',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) =>
      calls.push(
        `overlay:${payload.id ?? ''}:${payload.title}:${payload.body}:${payload.variant}:${payload.persistent ? 'pin' : 'auto'}`,
      ),
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, [
    'overlay::SubMiner:Subsync: choose engine and source:info:auto',
    'desktop:SubMiner:Subsync: choose engine and source',
  ]);
});

test('notifyConfiguredStatus routes pre-overlay status to osd only', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay loading...', {
    getNotificationType: () => 'both',
    isOverlayReady: () => false,
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) =>
      calls.push(`overlay:${payload.id ?? ''}:${payload.body ?? ''}`),
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['osd:Overlay loading...']);
});

test('notifyConfiguredStatus keeps osd-system on legacy surfaces', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay loading...', {
    getNotificationType: () => 'osd-system',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['osd:Overlay loading...', 'desktop:SubMiner:Overlay loading...']);
});

test('notifyConfiguredStatus can suppress desktop delivery for progress ticks', () => {
  const calls: string[] = [];

  notifyConfiguredStatus(
    'Subsync: syncing |',
    {
      getNotificationType: () => 'both',
      showOsd: (message) => {
        calls.push(`osd:${message}`);
      },
      showOverlayNotification: (payload) =>
        calls.push(
          `overlay:${payload.id ?? ''}:${payload.title}:${payload.body}:${payload.variant}:${payload.persistent ? 'pin' : 'auto'}`,
        ),
      showDesktopNotification: (title, options) =>
        calls.push(`desktop:${title}:${options.body ?? ''}`),
    },
    {
      id: 'subsync-status',
      title: 'Subsync',
      variant: 'progress',
      persistent: true,
      desktop: false,
    },
  );

  assert.deepEqual(calls, ['overlay:subsync-status:Subsync:Subsync: syncing |:progress:pin']);
});

test('notifyConfiguredStatus routes feedback through overlay without desktop delivery', () => {
  const calls: string[] = [];

  notifyConfiguredStatus(
    'Primary subtitle: hover',
    {
      getNotificationType: () => 'both',
      showOsd: (message) => {
        calls.push(`osd:${message}`);
      },
      showOverlayNotification: (payload) =>
        calls.push(`overlay:${payload.title}:${payload.body ?? ''}`),
      showDesktopNotification: (title, options) =>
        calls.push(`desktop:${title}:${options.body ?? ''}`),
    },
    { delivery: 'feedback' },
  );

  assert.deepEqual(calls, ['overlay:SubMiner:Primary subtitle: hover']);
});

test('notifyConfiguredStatus routes osd-system feedback through osd only', () => {
  const calls: string[] = [];

  notifyConfiguredStatus(
    'Secondary subtitle: visible',
    {
      getNotificationType: () => 'osd-system',
      showOsd: (message) => {
        calls.push(`osd:${message}`);
      },
      showDesktopNotification: (title, options) =>
        calls.push(`desktop:${title}:${options.body ?? ''}`),
    },
    { delivery: 'feedback' },
  );

  assert.deepEqual(calls, ['osd:Secondary subtitle: visible']);
});

test('notifyConfiguredStatus suppresses system-only feedback', () => {
  const calls: string[] = [];

  notifyConfiguredStatus(
    'Primary subtitle: visible',
    {
      getNotificationType: () => 'system',
      showOsd: (message) => {
        calls.push(`osd:${message}`);
      },
      showDesktopNotification: (title, options) =>
        calls.push(`desktop:${title}:${options.body ?? ''}`),
    },
    { delivery: 'feedback' },
  );

  assert.deepEqual(calls, []);
});

test('playback feedback options reuse subtitle mode notification ids', () => {
  assert.deepEqual(getPlaybackFeedbackNotificationOptions('Primary subtitle: hover'), {
    id: 'primary-subtitle-mode-feedback',
  });
  assert.deepEqual(getPlaybackFeedbackNotificationOptions('Secondary subtitle: hidden'), {
    id: 'secondary-subtitle-mode-feedback',
  });
  assert.deepEqual(getPlaybackFeedbackNotificationOptions('Secondary subtitle track: English'), {});
});

test('notifyConfiguredStatus falls back to desktop if overlay is unavailable', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay unavailable.', {
    getNotificationType: () => 'overlay',
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['desktop:SubMiner:Overlay unavailable.']);
});
