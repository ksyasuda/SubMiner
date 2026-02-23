import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleJellyfinPlayCommand } from './jellyfin-cli-play';

const baseSession = {
  serverUrl: 'http://localhost',
  accessToken: 'token',
  userId: 'user-id',
  username: 'user',
};

const baseClientInfo = {
  clientName: 'SubMiner',
  clientVersion: '1.0.0',
  deviceId: 'device-id',
};

const baseConfig = {
  defaultLibraryId: '',
};

test('play handler no-ops when play flag is disabled', async () => {
  let called = false;
  const handlePlay = createHandleJellyfinPlayCommand({
    playJellyfinItemInMpv: async () => {
      called = true;
    },
    logWarn: () => {},
  });

  const handled = await handlePlay({
    args: {
      jellyfinPlay: false,
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: baseConfig,
  });

  assert.equal(handled, false);
  assert.equal(called, false);
});

test('play handler warns when item id is missing', async () => {
  const warnings: string[] = [];
  const handlePlay = createHandleJellyfinPlayCommand({
    playJellyfinItemInMpv: async () => {
      throw new Error('should not play');
    },
    logWarn: (message) => warnings.push(message),
  });

  const handled = await handlePlay({
    args: {
      jellyfinPlay: true,
      jellyfinItemId: '',
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: baseConfig,
  });

  assert.equal(handled, true);
  assert.deepEqual(warnings, ['Ignoring --jellyfin-play without --jellyfin-item-id.']);
});

test('play handler runs playback with stream overrides', async () => {
  let called = false;
  const received: {
    itemId: string;
    audioStreamIndex?: number;
    subtitleStreamIndex?: number;
    setQuitOnDisconnectArm?: boolean;
  } = {
    itemId: '',
  };
  const handlePlay = createHandleJellyfinPlayCommand({
    playJellyfinItemInMpv: async (params) => {
      called = true;
      received.itemId = params.itemId;
      received.audioStreamIndex = params.audioStreamIndex;
      received.subtitleStreamIndex = params.subtitleStreamIndex;
      received.setQuitOnDisconnectArm = params.setQuitOnDisconnectArm;
    },
    logWarn: () => {},
  });

  const handled = await handlePlay({
    args: {
      jellyfinPlay: true,
      jellyfinItemId: 'item-1',
      jellyfinAudioStreamIndex: 2,
      jellyfinSubtitleStreamIndex: 3,
    } as never,
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: baseConfig,
  });

  assert.equal(handled, true);
  assert.equal(called, true);
  assert.equal(received.itemId, 'item-1');
  assert.equal(received.audioStreamIndex, 2);
  assert.equal(received.subtitleStreamIndex, 3);
  assert.equal(received.setQuitOnDisconnectArm, true);
});
