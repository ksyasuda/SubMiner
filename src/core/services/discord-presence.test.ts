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
  presenceStyle: 'default' as const,
  updateIntervalMs: 10_000,
  debounceMs: 200,
} as const;

const BASE_SESSION_STARTED_AT_MS = 1_700_000 * 1_000_000;

const baseSnapshot: DiscordPresenceSnapshot = {
  mediaTitle: 'Sousou no Frieren E01',
  mediaPath: '/media/Frieren/E01.mkv',
  subtitleText: '旅立ち',
  currentTimeSec: 95,
  mediaDurationSec: 1450,
  paused: false,
  connected: true,
  sessionStartedAtMs: BASE_SESSION_STARTED_AT_MS,
};

test('buildDiscordPresenceActivity maps polished payload fields (default style)', () => {
  const payload = buildDiscordPresenceActivity(baseConfig, baseSnapshot);
  assert.equal(payload.details, 'Sousou no Frieren E01');
  assert.equal(payload.state, 'Playing 01:35 / 24:10');
  assert.equal(payload.largeImageKey, 'subminer-logo');
  assert.equal(payload.smallImageKey, 'study');
  assert.equal(payload.smallImageText, '日本語学習中');
  assert.equal(payload.buttons, undefined);
  assert.equal(payload.startTimestamp, Math.floor(BASE_SESSION_STARTED_AT_MS / 1000));
});

test('buildDiscordPresenceActivity falls back to idle with default style', () => {
  const payload = buildDiscordPresenceActivity(baseConfig, {
    ...baseSnapshot,
    connected: false,
    mediaPath: null,
  });
  assert.equal(payload.state, 'Idle');
  assert.equal(payload.details, 'Sentence Mining');
});

test('buildDiscordPresenceActivity uses meme style fallback', () => {
  const memeConfig = { ...baseConfig, presenceStyle: 'meme' as const };
  const payload = buildDiscordPresenceActivity(memeConfig, {
    ...baseSnapshot,
    connected: false,
    mediaPath: null,
  });
  assert.equal(payload.details, 'Mining and crafting (Anki cards)');
  assert.equal(payload.smallImageText, 'Sentence Mining');
});

test('buildDiscordPresenceActivity uses japanese style', () => {
  const jpConfig = { ...baseConfig, presenceStyle: 'japanese' as const };
  const payload = buildDiscordPresenceActivity(jpConfig, {
    ...baseSnapshot,
    connected: false,
    mediaPath: null,
  });
  assert.equal(payload.details, '文の採掘中');
  assert.equal(payload.smallImageText, 'イマージョン学習');
});

test('buildDiscordPresenceActivity uses minimal style', () => {
  const minConfig = { ...baseConfig, presenceStyle: 'minimal' as const };
  const payload = buildDiscordPresenceActivity(minConfig, {
    ...baseSnapshot,
    connected: false,
    mediaPath: null,
  });
  assert.equal(payload.details, 'SubMiner');
  assert.equal(payload.smallImageKey, undefined);
  assert.equal(payload.smallImageText, undefined);
});

test('buildDiscordPresenceActivity shows media title regardless of style', () => {
  for (const presenceStyle of ['default', 'meme', 'japanese', 'minimal'] as const) {
    const payload = buildDiscordPresenceActivity({ ...baseConfig, presenceStyle }, baseSnapshot);
    assert.equal(payload.details, 'Sousou no Frieren E01');
    assert.equal(payload.state, 'Playing 01:35 / 24:10');
  }
});

test('service deduplicates identical updates and sends changed timeline', async () => {
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
  service.publish({ ...baseSnapshot, paused: true, currentTimeSec: 100 });
  timers.get(3)?.();
  await Promise.resolve();
  assert.equal(sent.length, 2);
  assert.equal(sent[1]?.state, 'Paused 01:40 / 24:10');
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
