import assert from 'node:assert/strict';
import test from 'node:test';

import {
  buildDiscordPresenceActivity,
  createDiscordPresenceService,
  type DiscordActivityPayload,
  type DiscordPresenceSnapshot,
} from './discord-presence';

const baseConfig = {
  enabled: true,
  clientId: '1234',
  detailsTemplate: 'Watching {title}',
  stateTemplate: '{status}',
  largeImageKey: 'subminer-logo',
  largeImageText: 'SubMiner',
  smallImageKey: 'study',
  smallImageText: 'Sentence Mining',
  buttonLabel: 'GitHub',
  buttonUrl: 'https://github.com/sudacode/SubMiner',
  updateIntervalMs: 10_000,
  debounceMs: 200,
} as const;

const baseSnapshot: DiscordPresenceSnapshot = {
  mediaTitle: 'Sousou no Frieren E01',
  mediaPath: '/media/Frieren/E01.mkv',
  subtitleText: '旅立ち',
  paused: false,
  connected: true,
  sessionStartedAtMs: 1_700_000_000_000,
};

test('buildDiscordPresenceActivity maps polished payload fields', () => {
  const payload = buildDiscordPresenceActivity(baseConfig, baseSnapshot);
  assert.equal(payload.details, 'Watching Sousou no Frieren E01');
  assert.equal(payload.state, 'Watching');
  assert.equal(payload.largeImageKey, 'subminer-logo');
  assert.equal(payload.smallImageKey, 'study');
  assert.equal(payload.buttons?.[0]?.label, 'GitHub');
  assert.equal(payload.startTimestamp, 1_700_000_000);
});

test('buildDiscordPresenceActivity falls back to idle when disconnected', () => {
  const payload = buildDiscordPresenceActivity(baseConfig, {
    ...baseSnapshot,
    connected: false,
    mediaPath: null,
  });
  assert.equal(payload.state, 'Idle');
});

test('service deduplicates identical activity updates and throttles interval', async () => {
  const sent: DiscordActivityPayload[] = [];
  const timers = new Map<number, () => void>();
  let timerId = 0;
  let nowMs = 100_000;

  const service = createDiscordPresenceService({
    config: baseConfig,
    createClient: () => ({
      login: async () => {},
      setActivity: async (activity) => {
        sent.push(activity);
      },
      clearActivity: async () => {},
      destroy: () => {},
    }),
    now: () => nowMs,
    setTimeoutFn: (callback) => {
      const id = ++timerId;
      timers.set(id, callback);
      return id as unknown as ReturnType<typeof setTimeout>;
    },
    clearTimeoutFn: (id) => {
      timers.delete(id as unknown as number);
    },
  });

  await service.start();
  service.publish(baseSnapshot);
  timers.get(1)?.();
  await Promise.resolve();
  assert.equal(sent.length, 1);

  service.publish(baseSnapshot);
  timers.get(2)?.();
  await Promise.resolve();
  assert.equal(sent.length, 1);

  nowMs += 10_001;
  service.publish({ ...baseSnapshot, paused: true });
  timers.get(3)?.();
  await Promise.resolve();
  assert.equal(sent.length, 2);
  assert.equal(sent[1]?.state, 'Paused');
});

test('service handles login failure and stop without throwing', async () => {
  let destroyed = false;
  const service = createDiscordPresenceService({
    config: baseConfig,
    createClient: () => ({
      login: async () => {
        throw new Error('discord not running');
      },
      setActivity: async () => {},
      clearActivity: async () => {},
      destroy: () => {
        destroyed = true;
      },
    }),
  });

  await assert.doesNotReject(async () => service.start());
  await assert.doesNotReject(async () => service.stop());
  assert.equal(destroyed, false);
});
