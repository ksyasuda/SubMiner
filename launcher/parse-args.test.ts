import test from 'node:test';
import assert from 'node:assert/strict';
import { parseArgs } from './config';

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

  throw new Error('expected parseArgs to exit');
}

test('parseArgs captures passthrough args for app subcommand', () => {
  const parsed = parseArgs(['app', '--anilist', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.appPassthrough, true);
  assert.deepEqual(parsed.appArgs, ['--anilist', '--log-level', 'debug']);
});

test('parseArgs supports bin alias for app subcommand', () => {
  const parsed = parseArgs(['bin', '--anilist-status'], 'subminer', {});

  assert.equal(parsed.appPassthrough, true);
  assert.deepEqual(parsed.appArgs, ['--anilist-status']);
});

test('parseArgs keeps all args after app verbatim', () => {
  const parsed = parseArgs(['app', '--start', '--anilist-setup', '-h'], 'subminer', {});

  assert.equal(parsed.appPassthrough, true);
  assert.deepEqual(parsed.appArgs, ['--start', '--anilist-setup', '-h']);
});

test('parseArgs captures mpv args string', () => {
  const parsed = parseArgs(['--args', '--pause=yes --title="movie night"'], 'subminer', {});

  assert.equal(parsed.mpvArgs, '--pause=yes --title="movie night"');
});

test('parseArgs maps jellyfin play action and log-level override', () => {
  const parsed = parseArgs(['jellyfin', 'play', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.jellyfinPlay, true);
  assert.equal(parsed.logLevel, 'debug');
});

test('parseArgs forwards jellyfin password-store option', () => {
  const parsed = parseArgs(['jf', 'setup', '--password-store', 'gnome-libsecret'], 'subminer', {});

  assert.equal(parsed.jellyfin, true);
  assert.equal(parsed.passwordStore, 'gnome-libsecret');
});

test('parseArgs maps config show action', () => {
  const parsed = parseArgs(['config', 'show'], 'subminer', {});

  assert.equal(parsed.configShow, true);
  assert.equal(parsed.configPath, false);
});

test('parseArgs maps mpv idle action', () => {
  const parsed = parseArgs(['mpv', 'idle'], 'subminer', {});

  assert.equal(parsed.mpvIdle, true);
  assert.equal(parsed.mpvStatus, false);
});

test('parseArgs captures youtube mode forwarding', () => {
  const parsed = parseArgs(['youtube', 'https://example.com', '--mode', 'generate'], 'subminer', {});

  assert.equal(parsed.target, 'https://example.com');
  assert.equal(parsed.youtubeMode, 'generate');
});

test('parseArgs maps dictionary command and log-level override', () => {
  const parsed = parseArgs(['dictionary', '.', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.dictionary, true);
  assert.equal(parsed.dictionaryTarget, process.cwd());
  assert.equal(parsed.logLevel, 'debug');
});

test('parseArgs maps stats command and log-level override', () => {
  const parsed = parseArgs(['stats', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.stats, true);
  assert.equal(parsed.logLevel, 'debug');
});

test('parseArgs maps stats background flag', () => {
  const parsed = parseArgs(['stats', '-b'], 'subminer', {}) as ReturnType<typeof parseArgs> & {
    statsBackground?: boolean;
    statsStop?: boolean;
  };

  assert.equal(parsed.stats, true);
  assert.equal(parsed.statsBackground, true);
  assert.equal(parsed.statsStop, false);
});

test('parseArgs maps stats stop flag', () => {
  const parsed = parseArgs(['stats', '-s'], 'subminer', {}) as ReturnType<typeof parseArgs> & {
    statsBackground?: boolean;
    statsStop?: boolean;
  };

  assert.equal(parsed.stats, true);
  assert.equal(parsed.statsStop, true);
  assert.equal(parsed.statsBackground, false);
});

test('parseArgs maps stats cleanup to vocab mode by default', () => {
  const parsed = parseArgs(['stats', 'cleanup'], 'subminer', {});

  assert.equal(parsed.stats, true);
  assert.equal(parsed.statsCleanup, true);
  assert.equal(parsed.statsCleanupVocab, true);
});

test('parseArgs maps explicit stats cleanup vocab flag', () => {
  const parsed = parseArgs(['stats', 'cleanup', '-v'], 'subminer', {});

  assert.equal(parsed.stats, true);
  assert.equal(parsed.statsCleanup, true);
  assert.equal(parsed.statsCleanupVocab, true);
});

test('parseArgs maps lifetime stats cleanup flag', () => {
  const parsed = parseArgs(['stats', 'cleanup', '--lifetime'], 'subminer', {});

  assert.equal(parsed.stats, true);
  assert.equal(parsed.statsCleanup, true);
  assert.equal(parsed.statsCleanupVocab, false);
  assert.equal(parsed.statsCleanupLifetime, true);
});

test('parseArgs rejects cleanup-only stats flags without cleanup action', () => {
  const error = withProcessExitIntercept(() => {
    parseArgs(['stats', '--vocab'], 'subminer', {});
  });

  assert.equal(error.code, 1);
  assert.match(error.message, /exit:1/);
});

test('parseArgs maps stats rebuild action to cleanup lifetime mode', () => {
  const parsed = parseArgs(['stats', 'rebuild'], 'subminer', {});

  assert.equal(parsed.stats, true);
  assert.equal(parsed.statsCleanup, true);
  assert.equal(parsed.statsCleanupVocab, false);
  assert.equal(parsed.statsCleanupLifetime, true);
});

test('parseArgs maps doctor refresh-known-words flag', () => {
  const parsed = parseArgs(['doctor', '--refresh-known-words'], 'subminer', {});

  assert.equal(parsed.doctor, true);
  assert.equal(parsed.doctorRefreshKnownWords, true);
});
