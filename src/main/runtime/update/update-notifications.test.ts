import test from 'node:test';
import assert from 'node:assert/strict';
import { notifyUpdateAvailable } from './update-notifications';
import type { OverlayNotificationPayload } from '../../../types/notification';

test('notifyUpdateAvailable routes notification surfaces from config', async () => {
  const calls: string[] = [];
  const deps = {
    showSystemNotification: (title: string, body: string) => {
      calls.push(`system:${title}:${body}`);
    },
    showOsdNotification: async (message: string) => {
      calls.push(`osd:${message}`);
    },
    showOverlayNotification: (payload: OverlayNotificationPayload) => {
      calls.push(`overlay:${payload.title}:${payload.body ?? ''}`);
    },
    log: (message: string) => {
      calls.push(`log:${message}`);
    },
  };

  await notifyUpdateAvailable({ notificationType: 'overlay', version: '0.15.0' }, deps);
  await notifyUpdateAvailable({ notificationType: 'system', version: '0.15.0' }, deps);
  await notifyUpdateAvailable({ notificationType: 'both', version: '0.15.0' }, deps);
  await notifyUpdateAvailable({ notificationType: 'osd-system', version: '0.15.0' }, deps);
  await notifyUpdateAvailable({ notificationType: 'none', version: '0.15.0' }, deps);

  assert.deepEqual(calls, [
    'overlay:SubMiner update available:SubMiner v0.15.0 is available',
    'system:SubMiner update available:SubMiner v0.15.0 is available',
    'overlay:SubMiner update available:SubMiner v0.15.0 is available',
    'system:SubMiner update available:SubMiner v0.15.0 is available',
    'osd:SubMiner v0.15.0 is available',
    'system:SubMiner update available:SubMiner v0.15.0 is available',
  ]);
});

test('notifyUpdateAvailable adds install and changelog actions to overlay update notifications', async () => {
  const payloads: OverlayNotificationPayload[] = [];

  await notifyUpdateAvailable(
    { notificationType: 'overlay', version: '0.15.0' },
    {
      showSystemNotification: () => {},
      showOsdNotification: async () => {},
      showOverlayNotification: (nextPayload) => {
        payloads.push(nextPayload);
      },
      log: () => {},
    },
  );

  const payload = payloads[0];
  assert.ok(payload);
  assert.deepEqual(payload.actions, [
    { id: 'install-update', label: 'Update' },
    { id: 'view-changelog', label: "What's New", keepOpen: true },
  ]);
  assert.equal(payload.id, 'subminer-update-available');
  assert.equal(payload.persistent, true);
});

test('notifyUpdateAvailable logs osd fallback when overlay notification fails', async () => {
  const calls: string[] = [];

  await notifyUpdateAvailable(
    { notificationType: 'osd', version: '0.15.0' },
    {
      showSystemNotification: () => {
        calls.push('system');
      },
      showOsdNotification: async () => {
        throw new Error('mpv disconnected');
      },
      showOverlayNotification: () => {
        calls.push('overlay');
      },
      log: (message) => {
        calls.push(message);
      },
    },
  );

  assert.deepEqual(calls, ['Update OSD notification failed: mpv disconnected']);
});

test('notifyUpdateAvailable logs non-error osd failures with thrown value', async () => {
  const calls: string[] = [];

  await notifyUpdateAvailable(
    { notificationType: 'osd', version: '0.15.0' },
    {
      showSystemNotification: () => {
        calls.push('system');
      },
      showOsdNotification: async () => {
        throw 'mpv disconnected';
      },
      showOverlayNotification: () => {
        calls.push('overlay');
      },
      log: (message) => {
        calls.push(message);
      },
    },
  );

  assert.deepEqual(calls, ['Update OSD notification failed: mpv disconnected']);
});
