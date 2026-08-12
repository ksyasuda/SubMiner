import test from 'node:test';
import assert from 'node:assert/strict';
import { parseArgs } from './config';
import { withProcessExitIntercept } from './test-support/exit-intercept.js';

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

test('parseArgs appends CLI mpv profile to configured mpv profile', () => {
  const parsed = parseArgs(['--profile', 'hdr'], 'subminer', {}, { profile: 'anime' });

  assert.equal(parsed.profile, 'anime,hdr');
});

test('parseArgs maps root settings window option', () => {
  const parsed = parseArgs(['--settings'], 'subminer', {});

  assert.equal(parsed.settings, true);
});

test('parseArgs maps root watch history flags', () => {
  const shortParsed = parseArgs(['-H'], 'subminer', {});
  const longParsed = parseArgs(['--history'], 'subminer', {});
  const rofiParsed = parseArgs(['-R', '-H'], 'subminer', {});
  const defaultParsed = parseArgs([], 'subminer', {});

  assert.equal(shortParsed.history, true);
  assert.equal(longParsed.history, true);
  assert.equal(rofiParsed.history, true);
  assert.equal(rofiParsed.useRofi, true);
  assert.equal(defaultParsed.history, false);
});

test('parseArgs maps root update flags without conflicting with jellyfin username', () => {
  const shortParsed = parseArgs(['-u'], 'subminer', {});
  const longParsed = parseArgs(['--update'], 'subminer', {});
  const jellyfinParsed = parseArgs(['jellyfin', 'setup', '-u', 'kyle'], 'subminer', {});

  assert.equal(shortParsed.update, true);
  assert.equal(longParsed.update, true);
  assert.equal(jellyfinParsed.update, false);
  assert.equal(jellyfinParsed.jellyfin, true);
  assert.equal(jellyfinParsed.jellyfinUsername, 'kyle');
});

test('parseArgs maps root version flags without conflicting with stats vocab flag', () => {
  const shortParsed = parseArgs(['-v'], 'subminer', {});
  const longParsed = parseArgs(['--version'], 'subminer', {});
  const statsParsed = parseArgs(['stats', 'cleanup', '-v'], 'subminer', {});

  assert.equal(shortParsed.version, true);
  assert.equal(longParsed.version, true);
  assert.equal(statsParsed.version, false);
  assert.equal(statsParsed.statsCleanupVocab, true);
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

test('parseArgs maps settings command to settings window', () => {
  const parsed = parseArgs(['settings'], 'subminer', {});

  assert.equal(parsed.settings, true);
  assert.equal(parsed.configPath, false);
  assert.equal(parsed.configShow, false);
});

test('parseArgs maps config path action to config path output', () => {
  const parsed = parseArgs(['config', 'path'], 'subminer', {});

  assert.equal(parsed.configPath, true);
  assert.equal(parsed.settings, false);
});

test('parseArgs rejects removed config open and launch actions', () => {
  const openExit = withProcessExitIntercept(() => {
    parseArgs(['config', 'open'], 'subminer', {});
  });
  const exit = withProcessExitIntercept(() => {
    parseArgs(['config', 'launch'], 'subminer', {});
  });

  assert.equal(openExit.code, 1);
  assert.match(openExit.stderr, /Unknown config action: open/);
  assert.equal(exit.code, 1);
  assert.match(exit.stderr, /Unknown config action: launch/);
});

test('parseArgs requires an explicit action for the config subcommand', () => {
  const exit = withProcessExitIntercept(() => {
    parseArgs(['config'], 'subminer', {});
  });

  assert.equal(exit.code, 1);
  assert.match(exit.stderr, /Unknown config action: \(none\)/);
});

test('parseArgs maps mpv idle action', () => {
  const parsed = parseArgs(['mpv', 'idle'], 'subminer', {});

  assert.equal(parsed.mpvIdle, true);
  assert.equal(parsed.mpvStatus, false);
});

test('parseArgs applies configured mpv launch mode default', () => {
  const parsed = parseArgs([], 'subminer', {}, { launchMode: 'maximized' });

  assert.equal(parsed.launchMode, 'maximized');
});

test('parseArgs maps dictionary command and log-level override', () => {
  const parsed = parseArgs(['dictionary', '.', '--log-level', 'debug'], 'subminer', {});

  assert.equal(parsed.dictionary, true);
  assert.equal(parsed.dictionaryTarget, process.cwd());
  assert.equal(parsed.logLevel, 'debug');
});

test('parseArgs maps dictionary candidate lookup and manual selection', () => {
  const candidateParsed = parseArgs(['dictionary', '--candidates', '.'], 'subminer', {});
  assert.equal(candidateParsed.dictionaryCandidates, true);
  assert.equal(candidateParsed.dictionaryTarget, process.cwd());

  const selectParsed = parseArgs(['dictionary', '--select', '21355', '.'], 'subminer', {});
  assert.equal(selectParsed.dictionarySelect, true);
  assert.equal(selectParsed.dictionaryAnilistId, 21355);
  assert.equal(selectParsed.dictionaryTarget, process.cwd());
});

test('parseArgs rejects conflicting dictionary candidate and selection modes', () => {
  const exit = withProcessExitIntercept(() => {
    parseArgs(['dictionary', '--candidates', '--select', '21355', '.'], 'subminer', {});
  });

  assert.equal(exit.code, 1);
  assert.match(exit.stderr, /Dictionary --candidates and --select cannot be combined/);
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

test('parseArgs maps duplicate-line stats cleanup flags', () => {
  const parsed = parseArgs(
    ['stats', 'cleanup', '--duplicate-lines', '--dry-run', '--lookback-days', '30'],
    'subminer',
    {},
  );

  assert.equal(parsed.statsCleanup, true);
  assert.equal(parsed.statsCleanupVocab, false);
  assert.equal(parsed.statsCleanupDuplicateLines, true);
  assert.equal(parsed.statsCleanupDryRun, true);
  assert.equal(parsed.statsCleanupLookbackDays, 30);

  const fractional = parseArgs(
    ['stats', 'cleanup', '--duplicate-lines', '--lookback-days', '1.5'],
    'subminer',
    {},
  );
  assert.equal(fractional.statsCleanupLookbackDays, 1);
});

test('parseArgs rejects duplicate-line flags without the duplicate-lines mode', () => {
  const error = withProcessExitIntercept(() => {
    parseArgs(['stats', 'cleanup', '--dry-run'], 'subminer', {});
  });

  assert.equal(error.code, 1);
  assert.match(error.stderr, /--dry-run and --lookback-days require --duplicate-lines/);
});

test('parseArgs rejects combining lifetime and duplicate-line cleanup modes', () => {
  const error = withProcessExitIntercept(() => {
    parseArgs(['stats', 'cleanup', '--lifetime', '--duplicate-lines'], 'subminer', {});
  });

  assert.equal(error.code, 1);
  assert.match(error.stderr, /Stats cleanup runs one mode at a time/);
});

test('parseArgs rejects unusable lookback windows', () => {
  for (const value of ['0', '0.5', '-5', 'soon']) {
    const error = withProcessExitIntercept(() => {
      parseArgs(
        ['stats', 'cleanup', '--duplicate-lines', '--lookback-days', value],
        'subminer',
        {},
      );
    });

    assert.equal(error.code, 1);
    assert.match(error.stderr, /--lookback-days must be at least one day/);
  }
});

test('parseArgs rejects cleanup-only stats flags without cleanup action', () => {
  const error = withProcessExitIntercept(() => {
    parseArgs(['stats', '--vocab'], 'subminer', {});
  });

  assert.equal(error.code, 1);
  assert.match(error.message, /exit:1/);
  assert.match(
    error.stderr,
    /Stats --vocab, --lifetime and --duplicate-lines flags require the cleanup action/,
  );
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

test('parseArgs maps logs export flag', () => {
  const parsed = parseArgs(['logs', '-e'], 'subminer', {});

  assert.equal(parsed.logsExport, true);
});

test('parseArgs requires an explicit logs action', () => {
  const exit = withProcessExitIntercept(() => {
    parseArgs(['logs'], 'subminer', {});
  });

  assert.equal(exit.code, 1);
  assert.match(exit.stderr, /Logs command requires -e or --export/);
});
