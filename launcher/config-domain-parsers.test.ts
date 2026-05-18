import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLauncherYoutubeSubgenConfig } from './config/youtube-subgen-config.js';
import { parseLauncherJellyfinConfig } from './config/jellyfin-config.js';
import { parseLauncherMpvConfig } from './config/mpv-config.js';
import { readExternalYomitanProfilePath } from './config.js';
import {
  buildPluginRuntimeScriptOptParts,
  parsePluginRuntimeConfigFromMainConfig,
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

test('parseLauncherMpvConfig reads launch mode preference', () => {
  const parsed = parseLauncherMpvConfig({
    mpv: {
      launchMode: ' maximized ',
      executablePath: 'ignored-here',
      socketPath: '/tmp/custom.sock',
      backend: 'x11',
      autoStartSubMiner: false,
      pauseUntilOverlayReady: false,
      subminerBinaryPath: '/opt/SubMiner/SubMiner.AppImage',
      aniskipEnabled: false,
      aniskipButtonKey: 'F8',
    },
  });

  assert.equal(parsed.launchMode, 'maximized');
  assert.equal(parsed.socketPath, '/tmp/custom.sock');
  assert.equal(parsed.backend, 'x11');
  assert.equal(parsed.autoStartSubMiner, false);
  assert.equal(parsed.pauseUntilOverlayReady, false);
  assert.equal(parsed.subminerBinaryPath, '/opt/SubMiner/SubMiner.AppImage');
  assert.equal(parsed.aniskipEnabled, false);
  assert.equal(parsed.aniskipButtonKey, 'F8');
});

test('parseLauncherMpvConfig ignores invalid launch mode values', () => {
  const parsed = parseLauncherMpvConfig({
    mpv: {
      launchMode: 'wide',
    },
  });

  assert.equal(parsed.launchMode, undefined);
});

test('parsePluginRuntimeConfigFromMainConfig maps config.jsonc values over plugin defaults', () => {
  const parsed = parsePluginRuntimeConfigFromMainConfig({
    auto_start_overlay: false,
    texthooker: {
      launchAtStartup: false,
    },
    mpv: {
      socketPath: '/tmp/config.sock',
      backend: 'sway',
      autoStartSubMiner: true,
      pauseUntilOverlayReady: true,
      subminerBinaryPath: '/opt/SubMiner/SubMiner.AppImage',
      aniskipEnabled: false,
      aniskipButtonKey: 'F8',
    },
  });

  assert.equal(parsed.socketPath, '/tmp/config.sock');
  assert.equal(parsed.backend, 'sway');
  assert.equal(parsed.autoStart, true);
  assert.equal(parsed.autoStartVisibleOverlay, false);
  assert.equal(parsed.autoStartPauseUntilReady, true);
  assert.equal(parsed.binaryPath, '/opt/SubMiner/SubMiner.AppImage');
  assert.equal(parsed.texthookerEnabled, false);
  assert.equal(parsed.aniskipEnabled, false);
  assert.equal(parsed.aniskipButtonKey, 'F8');
});

test('parsePluginRuntimeConfigFromMainConfig defaults to background-only managed startup', () => {
  const parsed = parsePluginRuntimeConfigFromMainConfig(null);

  assert.equal(parsed.autoStart, true);
  assert.equal(parsed.autoStartVisibleOverlay, false);
  assert.equal(parsed.autoStartPauseUntilReady, true);
  assert.equal(parsed.texthookerEnabled, false);
  assert.equal(parsed.aniskipEnabled, true);
  assert.equal(parsed.aniskipButtonKey, 'TAB');
});

test('buildPluginRuntimeScriptOptParts emits config values that override plugin defaults', () => {
  assert.deepEqual(
    buildPluginRuntimeScriptOptParts(
      {
        socketPath: '/tmp/config.sock',
        binaryPath: '/opt/SubMiner/SubMiner.AppImage',
        backend: 'x11',
        autoStart: true,
        autoStartVisibleOverlay: false,
        autoStartPauseUntilReady: true,
        texthookerEnabled: false,
        aniskipEnabled: false,
        aniskipButtonKey: 'F8',
      },
      '/fallback/SubMiner.AppImage',
    ),
    [
      'subminer-binary_path=/opt/SubMiner/SubMiner.AppImage',
      'subminer-socket_path=/tmp/config.sock',
      'subminer-backend=x11',
      'subminer-auto_start=yes',
      'subminer-auto_start_visible_overlay=no',
      'subminer-auto_start_pause_until_ready=yes',
      'subminer-texthooker_enabled=no',
      'subminer-aniskip_enabled=no',
      'subminer-aniskip_button_key=F8',
    ],
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
