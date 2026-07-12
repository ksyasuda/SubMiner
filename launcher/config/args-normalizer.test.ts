import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { withProcessExitIntercept } from '../test-support/exit-intercept.js';
import {
  applyInvocationsToArgs,
  applyRootOptionsToArgs,
  createDefaultArgs,
} from './args-normalizer.js';

function withTempDir<T>(fn: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-launcher-args-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('createDefaultArgs normalizes configured language codes and env thread override', () => {
  const originalThreads = process.env.SUBMINER_WHISPER_THREADS;
  process.env.SUBMINER_WHISPER_THREADS = '7';

  try {
    const parsed = createDefaultArgs({
      primarySubLanguages: [' JA ', 'jpn', 'ja'],
      secondarySubLanguages: ['en', 'ENG', ''],
      whisperThreads: 2,
    });

    assert.deepEqual(parsed.youtubePrimarySubLangs, ['ja', 'jpn']);
    assert.deepEqual(parsed.youtubeSecondarySubLangs, ['en', 'eng']);
    assert.deepEqual(parsed.youtubeAudioLangs, ['ja', 'jpn', 'en', 'eng']);
    assert.equal(parsed.whisperThreads, 7);
    assert.equal(parsed.youtubeWhisperSourceLanguage, 'ja');
    assert.equal(parsed.profile, '');
  } finally {
    if (originalThreads === undefined) {
      delete process.env.SUBMINER_WHISPER_THREADS;
    } else {
      process.env.SUBMINER_WHISPER_THREADS = originalThreads;
    }
  }
});

test('createDefaultArgs seeds mpv profile from launcher config', () => {
  const parsed = createDefaultArgs({}, { profile: 'anime' });

  assert.equal(parsed.profile, 'anime');
});

test('createDefaultArgs seeds log level from launcher logging config', () => {
  const parsed = createDefaultArgs({}, {}, { level: 'debug', rotation: 14 });

  assert.equal(parsed.logLevel, 'debug');
  assert.equal(parsed.logRotation, 14);
});

test('applyRootOptionsToArgs appends CLI mpv profile to configured profile', () => {
  const parsed = createDefaultArgs({}, { profile: 'anime' });

  applyRootOptionsToArgs(parsed, { profile: 'hdr' }, undefined);

  assert.equal(parsed.profile, 'anime,hdr');
});

test('applyRootOptionsToArgs maps file, directory, and url targets', () => {
  withTempDir((dir) => {
    const filePath = path.join(dir, 'movie.mkv');
    const folderPath = path.join(dir, 'anime');
    fs.writeFileSync(filePath, 'x');
    fs.mkdirSync(folderPath);

    const fileParsed = createDefaultArgs({});
    applyRootOptionsToArgs(fileParsed, {}, filePath);
    assert.equal(fileParsed.targetKind, 'file');
    assert.equal(fileParsed.target, filePath);

    const dirParsed = createDefaultArgs({});
    applyRootOptionsToArgs(dirParsed, {}, folderPath);
    assert.equal(dirParsed.directory, folderPath);
    assert.equal(dirParsed.target, '');
    assert.equal(dirParsed.targetKind, '');

    const urlParsed = createDefaultArgs({});
    applyRootOptionsToArgs(urlParsed, {}, 'https://example.test/video');
    assert.equal(urlParsed.targetKind, 'url');
    assert.equal(urlParsed.target, 'https://example.test/video');
  });
});

test('applyRootOptionsToArgs rejects unsupported targets', () => {
  const parsed = createDefaultArgs({});

  const error = withProcessExitIntercept(() => {
    applyRootOptionsToArgs(parsed, {}, '/definitely/missing/subminer-target');
  });

  assert.equal(error.code, 1);
  assert.match(error.message, /exit:1/);
  assert.match(error.stderr, /Not a file, directory, or supported URL/);
});

test('applyInvocationsToArgs maps config and jellyfin invocation state', () => {
  const parsed = createDefaultArgs({});

  applyInvocationsToArgs(parsed, {
    jellyfinInvocation: {
      action: 'play',
      play: true,
      server: 'https://jf.example',
      username: 'alice',
      password: 'secret',
      logLevel: 'debug',
    },
    configInvocation: {
      action: 'show',
      logLevel: 'warn',
    },
    settingsInvocation: null,
    mpvInvocation: null,
    appInvocation: null,
    dictionaryTriggered: false,
    dictionaryTarget: null,
    dictionaryLogLevel: null,
    dictionaryCandidates: false,
    dictionarySelect: false,
    dictionaryAnilistId: null,
    statsTriggered: false,
    statsBackground: false,
    statsStop: false,
    statsCleanup: false,
    statsCleanupVocab: false,
    statsCleanupLifetime: false,
    statsLogLevel: null,
    syncTriggered: false,
    syncHost: null,
    syncSnapshotPath: null,
    syncMergePath: null,
    syncDirection: 'both',
    syncRemoteCmd: null,
    syncDbPath: null,
    syncForce: false,
    syncJson: false,
    syncCheck: false,
    syncMakeTemp: false,
    syncRemoveTempPath: '',
    syncLogLevel: null,
    syncUiTriggered: false,
    syncUiLogLevel: null,
    doctorTriggered: false,
    doctorLogLevel: null,
    doctorRefreshKnownWords: false,
    logsTriggered: false,
    logsExport: false,
    texthookerTriggered: false,
    texthookerLogLevel: null,
    texthookerOpenBrowser: false,
  });

  assert.equal(parsed.jellyfin, false);
  assert.equal(parsed.jellyfinPlay, true);
  assert.equal(parsed.jellyfinDiscovery, false);
  assert.equal(parsed.jellyfinLogin, false);
  assert.equal(parsed.jellyfinLogout, false);
  assert.equal(parsed.jellyfinServer, 'https://jf.example');
  assert.equal(parsed.jellyfinUsername, 'alice');
  assert.equal(parsed.jellyfinPassword, 'secret');
  assert.equal(parsed.configShow, true);
  assert.equal(parsed.logLevel, 'warn');
});

test('applyInvocationsToArgs maps settings invocation to settings window', () => {
  const parsed = createDefaultArgs({});

  applyInvocationsToArgs(parsed, {
    jellyfinInvocation: null,
    configInvocation: null,
    settingsInvocation: {
      logLevel: undefined,
    },
    mpvInvocation: null,
    appInvocation: null,
    dictionaryTriggered: false,
    dictionaryTarget: null,
    dictionaryLogLevel: null,
    dictionaryCandidates: false,
    dictionarySelect: false,
    dictionaryAnilistId: null,
    statsTriggered: false,
    statsBackground: false,
    statsStop: false,
    statsCleanup: false,
    statsCleanupVocab: false,
    statsCleanupLifetime: false,
    statsLogLevel: null,
    syncTriggered: false,
    syncHost: null,
    syncSnapshotPath: null,
    syncMergePath: null,
    syncDirection: 'both',
    syncRemoteCmd: null,
    syncDbPath: null,
    syncForce: false,
    syncJson: false,
    syncCheck: false,
    syncMakeTemp: false,
    syncRemoveTempPath: '',
    syncLogLevel: null,
    syncUiTriggered: false,
    syncUiLogLevel: null,
    doctorTriggered: false,
    doctorLogLevel: null,
    doctorRefreshKnownWords: false,
    logsTriggered: false,
    logsExport: false,
    texthookerTriggered: false,
    texthookerLogLevel: null,
    texthookerOpenBrowser: false,
  });

  assert.equal(parsed.settings, true);
  assert.equal(parsed.configPath, false);
});

test('applyInvocationsToArgs fails when config invocation has no action', () => {
  const parsed = createDefaultArgs({});

  const error = withProcessExitIntercept(() => {
    applyInvocationsToArgs(parsed, {
      jellyfinInvocation: null,
      configInvocation: {
        action: undefined,
      },
      settingsInvocation: null,
      mpvInvocation: null,
      appInvocation: null,
      dictionaryTriggered: false,
      dictionaryTarget: null,
      dictionaryLogLevel: null,
      dictionaryCandidates: false,
      dictionarySelect: false,
      dictionaryAnilistId: null,
      statsTriggered: false,
      statsBackground: false,
      statsStop: false,
      statsCleanup: false,
      statsCleanupVocab: false,
      statsCleanupLifetime: false,
      statsLogLevel: null,
      syncTriggered: false,
      syncHost: null,
      syncSnapshotPath: null,
      syncMergePath: null,
      syncDirection: 'both',
      syncRemoteCmd: null,
      syncDbPath: null,
      syncForce: false,
      syncJson: false,
      syncCheck: false,
      syncMakeTemp: false,
      syncRemoveTempPath: '',
      syncLogLevel: null,
      syncUiTriggered: false,
      syncUiLogLevel: null,
      doctorTriggered: false,
      doctorLogLevel: null,
      doctorRefreshKnownWords: false,
      logsTriggered: false,
      logsExport: false,
      texthookerTriggered: false,
      texthookerLogLevel: null,
      texthookerOpenBrowser: false,
    });
  });

  assert.equal(error.code, 1);
  assert.match(error.stderr, /Unknown config action: \(none\)/);
});

test('applyInvocationsToArgs maps texthooker browser-open request', () => {
  const parsed = createDefaultArgs({});

  applyInvocationsToArgs(parsed, {
    jellyfinInvocation: null,
    configInvocation: null,
    settingsInvocation: null,
    mpvInvocation: null,
    appInvocation: null,
    dictionaryTriggered: false,
    dictionaryTarget: null,
    dictionaryLogLevel: null,
    dictionaryCandidates: false,
    dictionarySelect: false,
    dictionaryAnilistId: null,
    statsTriggered: false,
    statsBackground: false,
    statsStop: false,
    statsCleanup: false,
    statsCleanupVocab: false,
    statsCleanupLifetime: false,
    statsLogLevel: null,
    syncTriggered: false,
    syncHost: null,
    syncSnapshotPath: null,
    syncMergePath: null,
    syncDirection: 'both',
    syncRemoteCmd: null,
    syncDbPath: null,
    syncForce: false,
    syncJson: false,
    syncCheck: false,
    syncMakeTemp: false,
    syncRemoveTempPath: '',
    syncLogLevel: null,
    syncUiTriggered: false,
    syncUiLogLevel: null,
    doctorTriggered: false,
    doctorLogLevel: null,
    doctorRefreshKnownWords: false,
    logsTriggered: false,
    logsExport: false,
    texthookerTriggered: true,
    texthookerLogLevel: null,
    texthookerOpenBrowser: true,
  });

  assert.equal(parsed.texthookerOnly, true);
  assert.equal(parsed.texthookerOpenBrowser, true);
});
