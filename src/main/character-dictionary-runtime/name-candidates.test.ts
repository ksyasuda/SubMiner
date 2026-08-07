import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { CHARACTER_DICTIONARY_FORMAT_VERSION } from './constants';
import { createCharacterNameCandidateLookup } from './name-candidates';

function writeSnapshot(outputDir: string, mediaId: number, entries: Array<[string, string]>): void {
  const snapshotsDir = path.join(outputDir, 'snapshots');
  fs.mkdirSync(snapshotsDir, { recursive: true });
  fs.writeFileSync(
    path.join(snapshotsDir, `anilist-${mediaId}.json`),
    JSON.stringify({
      formatVersion: CHARACTER_DICTIONARY_FORMAT_VERSION,
      mediaId,
      mediaTitle: `title-${mediaId}`,
      entryCount: entries.length,
      updatedAt: 1,
      termEntries: entries.map(([term, reading]) => [
        term,
        reading,
        'name main',
        '',
        100,
        [],
        0,
        '',
      ]),
      images: [],
    }),
  );
}

function withTempDir<T>(run: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-name-candidates-'));
  try {
    return run(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('collects terms and readings for the current media', () => {
  withTempDir((dir) => {
    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['湊', 'みなと'],
    ]);
    writeSnapshot(dir, 2, [['カズマ', 'かずま']]);

    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
    });
    const candidates = lookup.get();

    assert.ok(candidates);
    assert.deepEqual([...candidates.forms].sort(), ['みなと', 'ミナト', '湊'].sort());
    // Deduplicated: both entries share the みなと reading.
    assert.equal(candidates.forms.length, 3);
  });
});

test('returns null without a media scope so the scanner stays exhaustive', () => {
  withTempDir((dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);

    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => null,
    });

    assert.equal(lookup.get(), null);
  });
});

test('returns null for a media with no cached snapshot', () => {
  withTempDir((dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);

    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 999,
    });

    assert.equal(lookup.get(), null);
  });
});

test('key changes when the snapshot content changes', () => {
  withTempDir((dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);
    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
    });
    const first = lookup.get();

    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['アクア', 'あくあ'],
    ]);
    lookup.invalidate();
    const second = lookup.get();

    assert.ok(first && second);
    assert.notEqual(first.key, second.key);
    assert.equal(second.forms.length, 4);
  });
});

// The lookup runs once per subtitle line, so it must not stat the snapshot
// directory every call. Asserted behaviorally: an unannounced on-disk change is
// invisible until the recheck interval elapses, which can only be true if the
// filesystem is not consulted per lookup.
test('does not re-read the snapshot directory on every lookup', () => {
  withTempDir((dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);
    let nowMs = 1_000_000;
    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
      now: () => nowMs,
    });

    assert.equal(lookup.get()?.forms.length, 2);

    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['アクア', 'あくあ'],
    ]);

    nowMs += 1000;
    assert.equal(lookup.get()?.forms.length, 2, 'expected the cached list within the interval');

    nowMs += 10_000;
    assert.equal(lookup.get()?.forms.length, 4, 'expected a refresh past the interval');
  });
});

test('invalidate picks up a snapshot change immediately', () => {
  withTempDir((dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);
    let nowMs = 1_000_000;
    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
      now: () => nowMs,
    });

    assert.equal(lookup.get()?.forms.length, 2);

    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['アクア', 'あくあ'],
    ]);
    nowMs += 1;
    lookup.invalidate();

    assert.equal(lookup.get()?.forms.length, 4);
  });
});
