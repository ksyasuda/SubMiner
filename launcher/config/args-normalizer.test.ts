import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  applyInvocationsToArgs,
  applyRootOptionsToArgs,
  createDefaultArgs,
} from './args-normalizer.js';

class ExitSignal extends Error {
  code: number;

  constructor(code: number) {
    super(`exit:${code}`);
    this.code = code;
  }
}

function withProcessExitIntercept(callback: () => void): ExitSignal {
  const originalExit = process.exit;
  try {
    process.exit = ((code?: number) => {
      throw new ExitSignal(code ?? 0);
    }) as typeof process.exit;
    callback();
  } catch (error) {
    if (error instanceof ExitSignal) {
      return error;
    }
    throw error;
  } finally {
    process.exit = originalExit;
  }

  throw new Error('expected process.exit');
}

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
    doctorTriggered: false,
    doctorLogLevel: null,
    doctorRefreshKnownWords: false,
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

test('applyInvocationsToArgs maps bare config invocation to settings window', () => {
  const parsed = createDefaultArgs({});

  applyInvocationsToArgs(parsed, {
    jellyfinInvocation: null,
    configInvocation: {
      action: undefined,
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
    doctorTriggered: false,
    doctorLogLevel: null,
    doctorRefreshKnownWords: false,
    texthookerTriggered: false,
    texthookerLogLevel: null,
    texthookerOpenBrowser: false,
  });

  assert.equal(parsed.configSettings, true);
  assert.equal(parsed.configPath, false);
});

test('applyInvocationsToArgs maps texthooker browser-open request', () => {
  const parsed = createDefaultArgs({});

  applyInvocationsToArgs(parsed, {
    jellyfinInvocation: null,
    configInvocation: null,
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
    doctorTriggered: false,
    doctorLogLevel: null,
    doctorRefreshKnownWords: false,
    texthookerTriggered: true,
    texthookerLogLevel: null,
    texthookerOpenBrowser: true,
  });

  assert.equal(parsed.texthookerOnly, true);
  assert.equal(parsed.texthookerOpenBrowser, true);
});
