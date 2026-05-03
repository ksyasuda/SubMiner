import assert from 'node:assert/strict';
import test from 'node:test';
import {
  clearJellyfinAuthSessionAndRefreshTray,
  isJellyfinConfiguredForTray,
  toggleJellyfinDiscoveryFromTray,
} from './jellyfin-tray-discovery';

test('detects Jellyfin tray configuration only when auth session fields exist', () => {
  assert.equal(
    isJellyfinConfiguredForTray({
      getResolvedJellyfinConfig: () => ({
        enabled: true,
        serverUrl: 'http://server:8096',
        accessToken: 'token',
        userId: 'user',
      }),
    }),
    true,
  );

  assert.equal(
    isJellyfinConfiguredForTray({
      getResolvedJellyfinConfig: () => ({
        enabled: true,
        serverUrl: 'http://server:8096',
        accessToken: 'token',
        userId: '',
      }),
    }),
    false,
  );
});

test('clears stored auth, stops active discovery, and refreshes tray', () => {
  const calls: string[] = [];

  clearJellyfinAuthSessionAndRefreshTray({
    clearStoredSession: () => calls.push('clear'),
    getRemoteSession: () => ({ advertiseNow: async () => true }),
    stopRemoteSession: () => calls.push('stop'),
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message) => calls.push(`error:${message}`),
    },
  });

  assert.deepEqual(calls, ['clear', 'stop', 'refresh']);
});

test('clear auth still refreshes tray when clear or stop throws', () => {
  const calls: string[] = [];

  clearJellyfinAuthSessionAndRefreshTray({
    clearStoredSession: () => {
      throw new Error('clear failed');
    },
    getRemoteSession: () => ({ advertiseNow: async () => true }),
    stopRemoteSession: () => {
      throw new Error('stop failed');
    },
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message) => calls.push(`error:${message}`),
    },
  });

  assert.deepEqual(calls, [
    'error:Failed to clear Jellyfin auth session.',
    'error:Failed to stop Jellyfin discovery while clearing auth session.',
    'refresh',
  ]);
});

test('starts explicit discovery and advertises cast target from tray', async () => {
  const calls: string[] = [];
  let session: { advertiseNow: () => Promise<boolean> } | null = null;

  await toggleJellyfinDiscoveryFromTray({
    getRemoteSession: () => session,
    stopRemoteSession: () => calls.push('stop'),
    startRemoteSession: async (options) => {
      assert.deepEqual(options, { explicit: true });
      calls.push('start');
      session = {
        advertiseNow: async () => {
          calls.push('advertise');
          return true;
        },
      };
    },
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message) => calls.push(`error:${message}`),
    },
    showMpvOsd: (message) => calls.push(`osd:${message}`),
  });

  assert.deepEqual(calls, [
    'start',
    'advertise',
    'info:Jellyfin discovery started; cast target is visible in server sessions.',
    'osd:Jellyfin discovery started',
    'refresh',
  ]);
});

test('starts explicit discovery and reports pending visibility from tray', async () => {
  const calls: string[] = [];
  let session: { advertiseNow: () => Promise<boolean> } | null = null;

  await toggleJellyfinDiscoveryFromTray({
    getRemoteSession: () => session,
    stopRemoteSession: () => calls.push('stop'),
    startRemoteSession: async (options) => {
      assert.deepEqual(options, { explicit: true });
      calls.push('start');
      session = {
        advertiseNow: async () => {
          calls.push('advertise');
          return false;
        },
      };
    },
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message) => calls.push(`error:${message}`),
    },
    showMpvOsd: (message) => calls.push(`osd:${message}`),
  });

  assert.deepEqual(calls, [
    'start',
    'advertise',
    'warn:Jellyfin discovery started, but cast target is not visible yet.',
    'osd:Jellyfin discovery started; waiting for visibility',
    'refresh',
  ]);
});

test('stops active discovery from tray', async () => {
  const calls: string[] = [];

  await toggleJellyfinDiscoveryFromTray({
    getRemoteSession: () => ({ advertiseNow: async () => true }),
    stopRemoteSession: () => calls.push('stop'),
    startRemoteSession: async () => {
      calls.push('start');
    },
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message) => calls.push(`error:${message}`),
    },
    showMpvOsd: (message) => calls.push(`osd:${message}`),
  });

  assert.deepEqual(calls, [
    'stop',
    'info:Jellyfin discovery stopped.',
    'osd:Jellyfin discovery stopped',
    'refresh',
  ]);
});

test('warns and refreshes tray when explicit discovery cannot create a session', async () => {
  const calls: string[] = [];

  await toggleJellyfinDiscoveryFromTray({
    getRemoteSession: () => null,
    stopRemoteSession: () => calls.push('stop'),
    startRemoteSession: async () => {
      calls.push('start');
    },
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message) => calls.push(`error:${message}`),
    },
    showMpvOsd: (message) => calls.push(`osd:${message}`),
  });

  assert.deepEqual(calls, [
    'start',
    'warn:Jellyfin discovery could not start. Configure Jellyfin first.',
    'osd:Jellyfin discovery unavailable',
    'refresh',
  ]);
});

test('reports discovery toggle failures and still refreshes tray', async () => {
  const calls: string[] = [];
  const error = new Error('boom');

  await toggleJellyfinDiscoveryFromTray({
    getRemoteSession: () => null,
    stopRemoteSession: () => calls.push('stop'),
    startRemoteSession: async () => {
      throw error;
    },
    refreshTrayMenu: () => calls.push('refresh'),
    logger: {
      info: (message) => calls.push(`info:${message}`),
      warn: (message) => calls.push(`warn:${message}`),
      error: (message, actualError) => {
        calls.push(`error:${message}`);
        assert.equal(actualError, error);
      },
    },
    showMpvOsd: (message) => calls.push(`osd:${message}`),
  });

  assert.deepEqual(calls, [
    'error:Failed to toggle Jellyfin discovery.',
    'osd:Jellyfin discovery failed',
    'refresh',
  ]);
});
