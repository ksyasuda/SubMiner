import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLauncherYoutubeSubgenConfig } from './config/youtube-subgen-config.js';
import { parseLauncherJellyfinConfig } from './config/jellyfin-config.js';
import { readExternalYomitanProfilePath } from './config.js';
import {
  getPluginConfigCandidates,
  parsePluginRuntimeConfigContent,
} from './config/plugin-runtime-config.js';
import { getDefaultSocketPath } from './types.js';

test('parseLauncherYoutubeSubgenConfig keeps only valid typed values', () => {
  const parsed = parseLauncherYoutubeSubgenConfig({
    ai: {
      enabled: true,
      apiKey: 'shared-key',
      baseUrl: 'https://openrouter.ai/api',
      model: 'openrouter/shared-model',
      systemPrompt: 'Legacy shared prompt.',
      requestTimeoutMs: 12000,
    },
    youtubeSubgen: {
      whisperBin: '/usr/bin/whisper',
      whisperModel: '/models/base.bin',
      whisperVadModel: '/models/vad.bin',
      whisperThreads: 6.8,
      fixWithAi: true,
      ai: {
        model: 'openrouter/subgen-model',
        systemPrompt: 'Fix subtitles only.',
      },
    },
    youtube: {
      primarySubLanguages: ['ja', 42, 'en'],
    },
    secondarySub: {
      secondarySubLanguages: ['eng', true, 'deu'],
    },
    jimaku: {
      apiKey: 'abc',
      apiKeyCommand: 'pass show key',
      apiBaseUrl: 'https://jimaku.cc',
      languagePreference: 'ja',
      maxEntryResults: 8.7,
    },
  });

  assert.equal('mode' in parsed, false);
  assert.deepEqual(parsed.primarySubLanguages, ['ja', 'en']);
  assert.deepEqual(parsed.secondarySubLanguages, ['eng', 'deu']);
  assert.equal(parsed.whisperVadModel, '/models/vad.bin');
  assert.equal(parsed.whisperThreads, 6);
  assert.equal(parsed.fixWithAi, true);
  assert.equal(parsed.ai?.enabled, true);
  assert.equal(parsed.ai?.apiKey, 'shared-key');
  assert.equal(parsed.ai?.model, 'openrouter/subgen-model');
  assert.equal(parsed.ai?.systemPrompt, 'Fix subtitles only.');
  assert.equal(parsed.ai?.requestTimeoutMs, 12000);
  assert.equal(parsed.jimakuLanguagePreference, 'ja');
  assert.equal(parsed.jimakuMaxEntryResults, 8);
});

test('parseLauncherJellyfinConfig omits legacy token and user id fields', () => {
  const parsed = parseLauncherJellyfinConfig({
    jellyfin: {
      enabled: true,
      serverUrl: 'https://jf.example',
      username: 'alice',
      accessToken: 'legacy-token',
      userId: 'legacy-user',
      pullPictures: true,
    },
  });

  assert.equal(parsed.enabled, true);
  assert.equal(parsed.serverUrl, 'https://jf.example');
  assert.equal(parsed.username, 'alice');
  assert.equal(parsed.pullPictures, true);
  assert.equal('accessToken' in parsed, false);
  assert.equal('userId' in parsed, false);
});

test('parsePluginRuntimeConfigContent reads socket path and startup gate options', () => {
  const parsed = parsePluginRuntimeConfigContent(`
# comment
socket_path = /tmp/custom.sock # trailing comment
auto_start = yes
auto_start_visible_overlay = true
auto_start_pause_until_ready = 1
`);
  assert.equal(parsed.socketPath, '/tmp/custom.sock');
  assert.equal(parsed.autoStart, true);
  assert.equal(parsed.autoStartVisibleOverlay, true);
  assert.equal(parsed.autoStartPauseUntilReady, true);
});

test('parsePluginRuntimeConfigContent falls back to disabled startup gate options', () => {
  const parsed = parsePluginRuntimeConfigContent(`
auto_start = maybe
auto_start_visible_overlay = no
auto_start_pause_until_ready = off
`);
  assert.equal(parsed.autoStart, false);
  assert.equal(parsed.autoStartVisibleOverlay, false);
  assert.equal(parsed.autoStartPauseUntilReady, false);
});

test('getPluginConfigCandidates resolves Windows mpv script-opts path', () => {
  assert.deepEqual(
    getPluginConfigCandidates({
      platform: 'win32',
      homeDir: 'C:\\Users\\tester',
      appDataDir: 'C:\\Users\\tester\\AppData\\Roaming',
    }),
    ['C:\\Users\\tester\\AppData\\Roaming\\mpv\\script-opts\\subminer.conf'],
  );
});

test('getDefaultSocketPath returns Windows named pipe default', () => {
  assert.equal(getDefaultSocketPath('win32'), '\\\\.\\pipe\\subminer-socket');
});

test('readExternalYomitanProfilePath detects configured external profile paths', () => {
  assert.equal(
    readExternalYomitanProfilePath({
      yomitan: {
        externalProfilePath: '  ~/.config/gsm_overlay  ',
      },
    }),
    '~/.config/gsm_overlay',
  );
  assert.equal(
    readExternalYomitanProfilePath({
      yomitan: {
        externalProfilePath: '   ',
      },
    }),
    null,
  );
  assert.equal(
    readExternalYomitanProfilePath({
      yomitan: null,
    }),
    null,
  );
  assert.equal(
    readExternalYomitanProfilePath({
      yomitan: {
        externalProfilePath: 123,
      },
    } as never),
    null,
  );
});
