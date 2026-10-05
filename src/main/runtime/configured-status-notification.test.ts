import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getPlaybackFeedbackNotificationOptions,
  getSubsyncStatusNotificationOptions,
  getYoutubeFlowStatusNotificationOptions,
  notifyConfiguredStatus,
} from './configured-status-notification';

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

test('notifyConfiguredStatus routes overlay-only status to the overlay', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay loading...', {
    getNotificationType: () => 'overlay',
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

test('notifyConfiguredStatus routes system status to desktop only', () => {
  const calls: string[] = [];

  notifyConfiguredStatus('Overlay loading...', {
    getNotificationType: () => 'system',
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
