import assert from 'node:assert/strict';
import test from 'node:test';
import { createPlayJellyfinItemInMpvHandler } from './jellyfin-playback-launch';

const baseSession = {
  serverUrl: 'http://localhost:8096',
  accessToken: 'token',
  userId: 'uid',
  username: 'alice',
};

const baseClientInfo = {
  clientName: 'SubMiner',
  clientVersion: '1.0.0',
  deviceId: 'did',
};

test('playback handler throws when mpv is not connected', async () => {
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => false,
    getMpvClient: () => null,
    resolvePlaybackPlan: async () => {
      throw new Error('unreachable');
    },
    applyJellyfinMpvDefaults: () => {},
    sendMpvCommand: () => {},
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
  });

  await assert.rejects(
    () =>
      handler({
        session: baseSession,
        clientInfo: baseClientInfo,
        jellyfinConfig: {},
        itemId: 'item-1',
      }),
    /MPV not connected and auto-launch failed/,
  );
});

test('playback handler drives mpv commands and playback state', async () => {
  const commands: Array<Array<string | number>> = [];
  const scheduled: Array<{ delay: number; callback: () => void }> = [];
  const calls: string[] = [];
  const activeStates: Array<Record<string, unknown>> = [];
  const reportPayloads: Array<Record<string, unknown>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 1',
      startTimeTicks: 12_000_000,
      audioStreamIndex: 1,
      subtitleStreamIndex: 2,
    }),
    applyJellyfinMpvDefaults: () => calls.push('defaults'),
    sendMpvCommand: (command) => commands.push(command),
    armQuitOnDisconnect: () => calls.push('arm'),
    schedule: (callback, delayMs) => {
      scheduled.push({ delay: delayMs, callback });
    },
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => calls.push('preload'),
    setActivePlayback: (state) => activeStates.push(state as Record<string, unknown>),
    setLastProgressAtMs: (value) => calls.push(`progress:${value}`),
    reportPlaying: (payload) => reportPayloads.push(payload as Record<string, unknown>),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-1',
  });

  assert.deepEqual(commands.slice(0, 5), [
    ['set_property', 'sub-auto', 'no'],
    ['loadfile', 'https://stream.example/video.m3u8', 'replace'],
    ['set_property', 'force-media-title', '[Jellyfin/direct] Episode 1'],
    ['set_property', 'sid', 'no'],
    ['seek', 1.2, 'absolute+exact'],
  ]);
  assert.equal(scheduled.length, 1);
  assert.equal(scheduled[0]?.delay, 500);
  scheduled[0]?.callback();
  assert.deepEqual(commands[commands.length - 1], ['set_property', 'sid', 'no']);

  assert.ok(calls.includes('defaults'));
  assert.ok(calls.includes('arm'));
  assert.ok(calls.includes('preload'));
  assert.ok(calls.includes('progress:0'));
  assert.ok(calls.includes('osd:Jellyfin direct: Episode 1'));

  assert.equal(activeStates.length, 1);
  assert.equal(activeStates[0]?.playMethod, 'DirectPlay');
  assert.equal(reportPayloads.length, 1);
  assert.equal(reportPayloads[0]?.eventName, 'start');
});
