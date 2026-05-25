import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildPlayJellyfinItemInMpvMainDepsHandler } from './jellyfin-playback-launch-main-deps';

test('play jellyfin item in mpv main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildPlayJellyfinItemInMpvMainDepsHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'u',
      mode: 'direct',
      title: 't',
      itemTitle: 't',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    applyJellyfinMpvDefaults: () => calls.push('defaults'),
    showVisibleOverlay: () => calls.push('visible-overlay'),
    sendMpvCommand: (command) => calls.push(`cmd:${command[0]}`),
    armQuitOnDisconnect: () => calls.push('arm'),
    schedule: (_callback, delayMs) => calls.push(`schedule:${delayMs}`),
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {
      calls.push('preload');
    },
    setActivePlayback: () => calls.push('active'),
    setLastProgressAtMs: () => calls.push('progress'),
    reportPlaying: () => calls.push('report'),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
  })();

  assert.equal(await deps.ensureMpvConnectedForPlayback(), true);
  assert.equal(typeof deps.getMpvClient(), 'object');
  assert.deepEqual(
    await deps.resolvePlaybackPlan({
      session: {
        serverUrl: 'http://localhost:8096',
        accessToken: 'token',
        userId: 'uid',
        username: 'alice',
      },
      clientInfo: {
        clientName: 'SubMiner',
        clientVersion: '1.0.0',
        deviceId: 'did',
      },
      jellyfinConfig: {},
      itemId: 'i',
    }),
    {
      url: 'u',
      mode: 'direct',
      title: 't',
      itemTitle: 't',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    },
  );
  deps.applyJellyfinMpvDefaults({ connected: true, send: () => {} });
  deps.showVisibleOverlay();
  deps.sendMpvCommand(['show-text', 'x']);
  deps.armQuitOnDisconnect();
  deps.schedule(() => {}, 500);
  assert.equal(deps.convertTicksToSeconds(20_000_000), 2);
  deps.preloadExternalSubtitles({
    session: {
      serverUrl: 'http://localhost:8096',
      accessToken: 'token',
      userId: 'uid',
      username: 'alice',
    },
    clientInfo: {
      clientName: 'SubMiner',
      clientVersion: '1.0.0',
      deviceId: 'did',
    },
    itemId: 'i',
  });
  deps.setActivePlayback({ itemId: 'i', mediaSourceId: undefined, playMethod: 'DirectPlay' });
  deps.setLastProgressAtMs(0);
  deps.reportPlaying({
    itemId: 'i',
    mediaSourceId: undefined,
    playMethod: 'DirectPlay',
    eventName: 'start',
  });
  deps.showMpvOsd('ok');

  assert.deepEqual(calls, [
    'defaults',
    'visible-overlay',
    'cmd:show-text',
    'arm',
    'schedule:500',
    'preload',
    'active',
    'progress',
    'report',
    'osd:ok',
  ]);
});
