import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  assertYomitanDictionaryMutationSafe,
  observeYomitanDictionaryCount,
} from './yomitan-dictionary-integrity';

function withTempDir(run: (directory: string) => void): void {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-dictionary-integrity-'));
  try {
    run(directory);
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
}

test('dictionary integrity observation establishes and updates a non-empty baseline', () => {
  withTempDir((userDataPath) => {
    assert.deepEqual(observeYomitanDictionaryCount(userDataPath, 5), {
      safe: true,
      previousCount: null,
    });
    assert.deepEqual(observeYomitanDictionaryCount(userDataPath, 3), {
      safe: true,
      previousCount: 5,
    });
    const statePath = path.join(userDataPath, 'yomitan-dictionary-integrity.json');
    const unchangedTimestamp = new Date('2000-01-01T00:00:00.000Z');
    fs.utimesSync(statePath, unchangedTimestamp, unchangedTimestamp);
    assert.deepEqual(observeYomitanDictionaryCount(userDataPath, 3), {
      safe: true,
      previousCount: 3,
    });
    assert.equal(fs.statSync(statePath).mtimeMs, unchangedTimestamp.getTime());
    assert.deepEqual(JSON.parse(fs.readFileSync(statePath, 'utf8')), { lastKnownNonEmptyCount: 3 });
    assert.deepEqual(fs.readdirSync(userDataPath), ['yomitan-dictionary-integrity.json']);
  });
});

test('dictionary integrity permits an empty profile before dictionaries are installed', () => {
  withTempDir((userDataPath) => {
    assert.deepEqual(observeYomitanDictionaryCount(userDataPath, 0), {
      safe: true,
      previousCount: null,
    });
  });
});

test('dictionary integrity rejects invalid counts without creating a safety record', () => {
  withTempDir((userDataPath) => {
    for (const invalidCount of [Number.NaN, Number.POSITIVE_INFINITY, -1]) {
      assert.deepEqual(observeYomitanDictionaryCount(userDataPath, invalidCount), {
        safe: false,
        previousCount: null,
        message: 'SubMiner could not verify Yomitan dictionary storage: invalid dictionary count.',
      });
    }

    assert.equal(
      fs.existsSync(path.join(userDataPath, 'yomitan-dictionary-integrity.json')),
      false,
    );
  });
});

test('dictionary integrity migrates the previous count from first-run setup state', () => {
  withTempDir((userDataPath) => {
    fs.writeFileSync(
      path.join(userDataPath, 'setup-state.json'),
      JSON.stringify({
        version: 4,
        status: 'completed',
        completedAt: '2026-08-18T00:00:00.000Z',
        completionSource: 'user',
        yomitanSetupMode: 'internal',
        lastSeenYomitanDictionaryCount: 4,
        pluginInstallStatus: 'installed',
        pluginInstallPathSummary: null,
        windowsMpvShortcutPreferences: {
          startMenuEnabled: true,
          desktopEnabled: false,
        },
        windowsMpvShortcutLastStatus: 'installed',
        bunInstallStatus: 'installed',
        launcherInstallStatus: 'installed',
        launcherInstallPath: '/home/tester/.local/bin/subminer',
      }),
      'utf8',
    );

    assert.throws(
      () => assertYomitanDictionaryMutationSafe(userDataPath, 0),
      /reported zero dictionaries after previously reporting 4/,
    );
  });
});

test('dictionary integrity blocks automatic mutation after a non-empty profile becomes empty', () => {
  withTempDir((userDataPath) => {
    observeYomitanDictionaryCount(userDataPath, 6);

    assert.throws(
      () => assertYomitanDictionaryMutationSafe(userDataPath, 0),
      /reported zero dictionaries after previously reporting 6/,
    );
    assert.deepEqual(observeYomitanDictionaryCount(userDataPath, 0), {
      safe: false,
      previousCount: 6,
      message:
        'Yomitan reported zero dictionaries after previously reporting 6. SubMiner blocked automatic dictionary changes because Chromium storage may have been reset. Close SubMiner and restore or inspect the profile before changing dictionaries.',
    });
  });
});

test('dictionary integrity fails closed when its state is malformed', () => {
  withTempDir((userDataPath) => {
    const statePath = path.join(userDataPath, 'yomitan-dictionary-integrity.json');
    fs.writeFileSync(statePath, '{}', 'utf8');

    assert.throws(
      () => assertYomitanDictionaryMutationSafe(userDataPath, 2),
      (error: unknown) => {
        assert.ok(error instanceof Error);
        assert.match(error.message, /could not verify Yomitan dictionary storage/);
        assert.equal(error.message.includes(statePath), true);
        return true;
      },
    );
  });
});
