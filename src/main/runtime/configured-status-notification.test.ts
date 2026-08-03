import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getPlaybackFeedbackNotificationOptions,
  getSubsyncStatusNotificationOptions,
  getYoutubeFlowStatusNotificationOptions,
  notifyConfiguredStatus,
} from './configured-status-notification';
import { createOverlayNotificationDelivery } from './overlay-notification-delivery';

test('notifyConfiguredStatus routes both to overlay and system without osd', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Subsync: choose engine and subtitles', {
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
    'overlay::SubMiner:Subsync: choose engine and subtitles:info:auto',
    'desktop:SubMiner:Subsync: choose engine and subtitles',
  ]);
});

test('notifyConfiguredStatus queues overlay for pre-overlay both status and preserves desktop', () => {
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

  assert.deepEqual(calls, ['overlay::Overlay loading...', 'desktop:SubMiner:Overlay loading...']);
});

test('notifyConfiguredStatus queues overlay for pre-overlay overlay-only status', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay loading...', {
    getNotificationType: () => 'overlay',
    isOverlayReady: () => false,
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) =>
      calls.push(`overlay:${payload.id ?? ''}:${payload.body ?? ''}`),
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['overlay::Overlay loading...']);
});

test('notifyConfiguredStatus routes pre-overlay system status to desktop only', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay loading...', {
    getNotificationType: () => 'system',
    isOverlayReady: () => false,
    showOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload) =>
      calls.push(`overlay:${payload.id ?? ''}:${payload.body ?? ''}`),
    showDesktopNotification: (title, options) =>
      calls.push(`desktop:${title}:${options.body ?? ''}`),
  });

  assert.deepEqual(calls, ['desktop:SubMiner:Overlay loading...']);
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

test('notifyConfiguredStatus queues osd status when mpv osd is unavailable', () => {
  const calls: string[] = [];

  notifyConfiguredStatus(
    'YouTube media cache is downloading.',
    {
      getNotificationType: () => 'osd',
      showOsd: (message) => {
        calls.push(`osd:${message}`);
        return false;
      },
      queueOsd: (message, options) => {
        calls.push(`queue:${options.id ?? ''}:${message}`);
      },
      showDesktopNotification: (title, options) =>
        calls.push(`desktop:${title}:${options.body ?? ''}`),
    },
    {
      id: 'youtube-media-cache-status',
      title: 'YouTube media cache',
      variant: 'progress',
      persistent: true,
    },
  );

  assert.deepEqual(calls, [
    'osd:YouTube media cache is downloading.',
    'queue:youtube-media-cache-status:YouTube media cache is downloading.',
  ]);
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

test('subsync progress keeps the osd spinner frame but strips it from the overlay card', () => {
  const calls: string[] = [];

  for (const frame of ['|', '/', '-', '\\']) {
    const message = `Subsync: syncing ${frame}`;
    notifyConfiguredStatus(
      message,
      {
        getNotificationType: () => 'both',
        showOsd: (osdMessage) => {
          calls.push(`osd:${osdMessage}`);
        },
        showOverlayNotification: (payload) =>
          calls.push(
            `overlay:${payload.body}:${payload.variant}:${payload.persistent ? 'pin' : 'auto'}`,
          ),
        showDesktopNotification: (title, options) =>
          calls.push(`desktop:${title}:${options.body ?? ''}`),
      },
      getSubsyncStatusNotificationOptions(message),
    );
  }

  assert.deepEqual(calls, [
    'overlay:Subsync: syncing:progress:pin',
    'overlay:Subsync: syncing:progress:pin',
    'overlay:Subsync: syncing:progress:pin',
    'overlay:Subsync: syncing:progress:pin',
  ]);

  calls.length = 0;
  notifyConfiguredStatus(
    'Subsync: syncing /',
    {
      getNotificationType: () => 'osd',
      showOsd: (osdMessage) => {
        calls.push(`osd:${osdMessage}`);
      },
      showOverlayNotification: (payload) => calls.push(`overlay:${payload.body}`),
      showDesktopNotification: (title, options) =>
        calls.push(`desktop:${title}:${options.body ?? ''}`),
    },
    getSubsyncStatusNotificationOptions('Subsync: syncing /'),
  );

  assert.deepEqual(calls, ['osd:Subsync: syncing /']);
});

test('subsync result notifications keep their message intact', () => {
  assert.equal(
    getSubsyncStatusNotificationOptions('Subtitle synchronized with ffsubsync').overlayBody,
    'Subtitle synchronized with ffsubsync',
  );
  const failure = getSubsyncStatusNotificationOptions('ffsubsync synchronization failed: boom');
  assert.equal(failure.variant, 'error');
  assert.equal(failure.overlayBody, 'ffsubsync synchronization failed: boom');
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

test('youtube flow status options route picker opening as one-shot configured status', () => {
  assert.deepEqual(getYoutubeFlowStatusNotificationOptions('Opening YouTube subtitle picker...'), {
    id: 'youtube-subtitles-status',
    title: 'YouTube subtitles',
    variant: 'info',
    persistent: false,
    desktop: true,
  });
});

test('youtube flow status options route loaded messages as transient success', () => {
  assert.deepEqual(getYoutubeFlowStatusNotificationOptions('Subtitles loaded.'), {
    id: 'youtube-subtitles-status',
    title: 'YouTube subtitles',
    variant: 'success',
    persistent: false,
    desktop: true,
  });
  assert.deepEqual(
    getYoutubeFlowStatusNotificationOptions('Primary and secondary subtitles loaded.'),
    {
      id: 'youtube-subtitles-status',
      title: 'YouTube subtitles',
      variant: 'success',
      persistent: false,
      desktop: true,
    },
  );
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

test('overlay notification delivery queues until an overlay window is ready', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) => sent.push(`${payload.id ?? ''}:${'body' in payload ? payload.body : ''}`),
  });

  delivery.send({ id: 'startup-tokenization', title: 'Subtitle tokenization', body: 'Loading' });
  delivery.send({ id: 'character-dictionary-auto-sync', title: 'Dictionary', body: 'Building' });

  assert.equal(delivery.getQueuedCount(), 2);
  assert.deepEqual(sent, []);

  ready = true;
  delivery.flush();

  assert.equal(delivery.getQueuedCount(), 0);
  assert.deepEqual(sent, [
    'startup-tokenization:Loading',
    'character-dictionary-auto-sync:Building',
  ]);
});

test('overlay notification delivery upserts queued progress by notification id', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) => sent.push(`${payload.id ?? ''}:${'body' in payload ? payload.body : ''}`),
  });

  delivery.send({ id: 'startup-subtitle-annotations', title: 'Subtitle annotations', body: '|' });
  delivery.send({ id: 'startup-subtitle-annotations', title: 'Subtitle annotations', body: '/' });
  delivery.send({ id: 'startup-tokenization', title: 'Subtitle tokenization', body: 'Ready' });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, ['startup-subtitle-annotations:/', 'startup-tokenization:Ready']);
});

test('overlay notification delivery preserves queued events with distinct history ids', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push(
        `${payload.id ?? ''}:${'historyId' in payload ? payload.historyId : ''}:${'body' in payload ? payload.body : ''}`,
      ),
  });

  delivery.send({
    id: 'character-dictionary-auto-sync',
    historyId: 'character-dictionary-auto-sync-checking',
    title: 'Character dictionary',
    body: 'Checking character dictionary...',
    persistent: true,
  });
  delivery.send({
    id: 'character-dictionary-auto-sync',
    historyId: 'character-dictionary-auto-sync-building',
    title: 'Character dictionary',
    body: 'Building character dictionary...',
    persistent: true,
  });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, [
    'character-dictionary-auto-sync:character-dictionary-auto-sync-checking:Checking character dictionary...',
    'character-dictionary-auto-sync:character-dictionary-auto-sync-building:Building character dictionary...',
  ]);
});

test('overlay notification delivery preserves queued startup progress before terminal update', () => {
  const sent: string[] = [];
  const scheduled: Array<() => void> = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push(
        `${payload.id ?? ''}:${'body' in payload ? payload.body : ''}:${'persistent' in payload && payload.persistent ? 'pin' : 'auto'}`,
      ),
    scheduleFlushRetry: (callback) => {
      scheduled.push(callback);
    },
  });

  delivery.send({
    id: 'startup-tokenization',
    title: 'Subtitle tokenization',
    body: 'Loading subtitle tokenization...',
    variant: 'progress',
    persistent: true,
  });
  delivery.send({
    id: 'startup-tokenization',
    title: 'Subtitle tokenization',
    body: 'Subtitle tokenization ready',
    variant: 'success',
    persistent: false,
  });

  ready = true;
  delivery.flush();
  scheduled.shift()?.();

  assert.deepEqual(sent, [
    'startup-tokenization:Loading subtitle tokenization...:pin',
    'startup-tokenization:Subtitle tokenization ready:auto',
  ]);
});

test('overlay notification delivery defers terminal update after first queued progress paint', () => {
  const sent: string[] = [];
  const scheduled: Array<() => void> = [];
  const delays: number[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push(
        `${payload.id ?? ''}:${'body' in payload ? payload.body : ''}:${'persistent' in payload && payload.persistent ? 'pin' : 'auto'}`,
      ),
    scheduleFlushRetry: (callback, delayMs) => {
      scheduled.push(callback);
      delays.push(delayMs);
    },
    terminalUpdateDelayMs: 750,
  });

  delivery.send({
    id: 'startup-subtitle-annotations',
    title: 'Subtitle annotations',
    body: 'Loading subtitle annotations |',
    variant: 'progress',
    persistent: true,
  });
  delivery.send({
    id: 'startup-subtitle-annotations',
    title: 'Subtitle annotations',
    body: 'Subtitle annotations loaded',
    variant: 'success',
    persistent: false,
  });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, ['startup-subtitle-annotations:Loading subtitle annotations |:pin']);
  assert.equal(delivery.getQueuedCount(), 1);
  assert.deepEqual(delays, [750]);

  scheduled.shift()?.();

  assert.equal(delivery.getQueuedCount(), 0);
  assert.deepEqual(sent, [
    'startup-subtitle-annotations:Loading subtitle annotations |:pin',
    'startup-subtitle-annotations:Subtitle annotations loaded:auto',
  ]);
});

test('overlay notification delivery retries flush when lifecycle fires before window readiness settles', () => {
  const sent: string[] = [];
  const scheduled: Array<() => void> = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) => sent.push(`${payload.id ?? ''}:${'body' in payload ? payload.body : ''}`),
    scheduleFlushRetry: (callback) => {
      scheduled.push(callback);
    },
  });

  delivery.send({ id: 'startup-tokenization', title: 'Subtitle tokenization', body: 'Loading' });
  delivery.flush();

  assert.equal(delivery.getQueuedCount(), 1);
  assert.equal(scheduled.length, 1);
  assert.deepEqual(sent, []);

  ready = true;
  scheduled.shift()?.();

  assert.equal(delivery.getQueuedCount(), 0);
  assert.deepEqual(sent, ['startup-tokenization:Loading']);
});

test('overlay notification delivery drops queued notification when dismissed before flush', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push('dismiss' in payload ? `dismiss:${payload.id}` : `show:${payload.id ?? ''}`),
  });

  delivery.send({ id: 'overlay-loading-status', title: 'SubMiner', body: 'Overlay loading' });
  delivery.send({ id: 'overlay-loading-status', dismiss: true });

  ready = true;
  delivery.flush();

  assert.deepEqual(sent, []);
});

test('overlay notification delivery removes queued notification when dismissed at readiness', () => {
  const sent: string[] = [];
  let ready = false;
  const delivery = createOverlayNotificationDelivery({
    hasReadyOverlayWindow: () => ready,
    send: (payload) =>
      sent.push('dismiss' in payload ? `dismiss:${payload.id}` : `show:${payload.id ?? ''}`),
  });

  delivery.send({ id: 'overlay-loading-status', title: 'SubMiner', body: 'Overlay loading' });

  ready = true;
  delivery.send({ id: 'overlay-loading-status', dismiss: true });
  delivery.flush();

  assert.deepEqual(sent, ['dismiss:overlay-loading-status']);
});
