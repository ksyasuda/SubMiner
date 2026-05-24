import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createYoutubePrimarySubtitleNotificationRuntime,
  type YoutubePrimarySubtitleNotificationTimer,
} from './youtube-primary-subtitle-notification';

function createTimerHarness() {
  let nextId = 1;
  const timers = new Map<number, () => void | Promise<void>>();
  return {
    schedule: (fn: () => void | Promise<void>): YoutubePrimarySubtitleNotificationTimer => {
      const id = nextId++;
      timers.set(id, fn);
      return { id };
    },
    clear: (timer: YoutubePrimarySubtitleNotificationTimer | null) => {
      if (!timer) {
        return;
      }
      if (typeof timer === 'object' && 'id' in timer) {
        timers.delete(timer.id);
      }
    },
    runAll: () => {
      const pending = [...timers.values()];
      timers.clear();
      for (const fn of pending) {
        void fn();
      }
    },
    runAllAsync: async () => {
      const pending = [...timers.values()];
      timers.clear();
      for (const fn of pending) {
        await fn();
      }
    },
    size: () => timers.size,
  };
}

test('notifier reports missing preferred primary subtitle once for youtube media', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja', 'jpn'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleSubtitleTrackChange(null);
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 2, lang: 'en', title: 'English', external: true },
  ]);

  assert.equal(timers.size(), 1);
  timers.runAll();
  timers.runAll();

  assert.deepEqual(notifications, [
    'Primary subtitle failed to download or load. Try again from the subtitle modal.',
  ]);
});

test('notifier suppresses failure when preferred primary subtitle is selected', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja', 'jpn'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 5, lang: 'ja', title: 'Japanese', external: true },
  ]);
  runtime.handleSubtitleTrackChange(5);
  timers.runAll();

  assert.deepEqual(notifications, []);
});

test('notifier suppresses failure when selected track is marked active before sid arrives', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja', 'jpn'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleSubtitleTrackChange(null);
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 5, lang: 'ja', title: 'Japanese', external: false, selected: true },
  ]);
  timers.runAll();

  assert.deepEqual(notifications, []);
});

test('notifier suppresses failure when any external subtitle track is selected', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja', 'jpn'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 5, lang: '', title: 'auto-ja-orig.ja-orig.vtt', external: true },
  ]);
  runtime.handleSubtitleTrackChange(5);
  timers.runAll();

  assert.deepEqual(notifications, []);
});

test('notifier resets when media changes away from youtube', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleMediaPathChange('/tmp/video.mkv');
  timers.runAll();

  assert.deepEqual(notifications, []);
});

test('notifier ignores empty and null media paths and waits for track list before reporting', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange(null);
  runtime.handleMediaPathChange('');
  assert.equal(timers.size(), 0);

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleSubtitleTrackChange(7);
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 7, lang: 'ja', title: 'Japanese', external: true },
  ]);
  timers.runAll();
  assert.deepEqual(notifications, []);
});

test('notifier suppresses timer while app-owned youtube flow is still settling', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.setAppOwnedFlowInFlight(true);
  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');

  assert.equal(timers.size(), 0);

  runtime.setAppOwnedFlowInFlight(false);
  assert.equal(timers.size(), 1);

  timers.runAll();
  assert.deepEqual(notifications, [
    'Primary subtitle failed to download or load. Try again from the subtitle modal.',
  ]);
});

test('notifier suppresses stale delayed failure after primary subtitle load is confirmed', () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=abc');
  runtime.handleSubtitleTrackChange(null);
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 2, lang: 'en', title: 'English', external: true },
  ]);
  runtime.markCurrentMediaPrimarySubtitleLoaded();

  assert.equal(timers.size(), 0);
  timers.runAll();
  assert.deepEqual(notifications, []);
});

test('notifier suppresses delayed failure when live mpv state has downloaded primary selected', async () => {
  const notifications: string[] = [];
  const timers = createTimerHarness();
  let liveStateReads = 0;
  const runtime = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => ['ja'],
    notifyFailure: (message) => {
      notifications.push(message);
    },
    schedule: (fn) => timers.schedule(fn),
    clearSchedule: (timer) => timers.clear(timer),
    getCurrentSubtitleState: async () => {
      liveStateReads += 1;
      return {
        sid: 22,
        trackList: [
          {
            type: 'sub',
            id: 1,
            lang: 'en',
            title: 'English',
            external: true,
            selected: true,
            'main-selection': 1,
          },
          {
            type: 'sub',
            id: 22,
            lang: 'ja',
            title: 'manual-ja.ja.srt',
            external: true,
            selected: true,
            'main-selection': 0,
            'external-filename': '/tmp/subminer-youtube-subtitles-aahLWu/manual-ja.ja.srt',
          },
        ],
      };
    },
  });

  runtime.handleMediaPathChange('https://www.youtube.com/watch?v=uO2jfacqjYQ');
  runtime.handleSubtitleTrackChange(null);
  runtime.handleSubtitleTrackListChange([
    { type: 'sub', id: 1, lang: 'en', title: 'English', external: true, selected: false },
  ]);

  assert.equal(timers.size(), 1);
  await timers.runAllAsync();

  assert.equal(liveStateReads, 1);
  assert.deepEqual(notifications, []);
});
