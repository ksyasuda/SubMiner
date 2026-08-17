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

async function withTempDir<T>(run: (dir: string) => Promise<T> | T): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-name-candidates-'));
  try {
    return await run(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

// The snapshot index rebuilds in the background while lookups serve stale data, so tests poll the
// probe until the refresh they triggered has landed.
async function waitForRefresh<T>(probe: () => T | null | undefined): Promise<T> {
  const deadline = Date.now() + 5000;
  for (;;) {
    const value = probe();
    if (value !== null && value !== undefined) {
      return value;
    }
    if (Date.now() > deadline) {
      throw new Error('timed out waiting for background snapshot refresh');
    }
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

test('collects terms and readings for the current media', async () => {
  await withTempDir(async (dir) => {
    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['湊', 'みなと'],
    ]);
    writeSnapshot(dir, 2, [['カズマ', 'かずま']]);

    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
    });
    const candidates = await waitForRefresh(() => lookup.get());

    assert.deepEqual([...candidates.forms].sort(), ['みなと', 'ミナト', '湊'].sort());
    // Deduplicated: both entries share the みなと reading.
    assert.equal(candidates.forms.length, 3);
  });
});

test('returns null without a media scope so the scanner stays exhaustive', async () => {
  await withTempDir(async (dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);

    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => null,
    });

    // The explicitly-scoped probe proves the index has loaded before the unscoped case is judged.
    await waitForRefresh(() => lookup.get(1));
    assert.equal(lookup.get(), null);
  });
});

test('returns null for a media with no cached snapshot', async () => {
  await withTempDir(async (dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);

    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 999,
    });

    await waitForRefresh(() => lookup.get(1));
    assert.equal(lookup.get(), null);
  });
});

test('key changes when the snapshot content changes', async () => {
  await withTempDir(async (dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);
    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
    });
    const first = await waitForRefresh(() => lookup.get());

    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['アクア', 'あくあ'],
    ]);
    lookup.invalidate();
    const second = await waitForRefresh(() => {
      const candidates = lookup.get();
      return candidates && candidates.forms.length === 4 ? candidates : null;
    });

    assert.notEqual(first.key, second.key);
  });
});

// The lookup runs once per subtitle line, so it must not stat the snapshot
// directory every call. Asserted behaviorally: an unannounced on-disk change is
// invisible until the recheck interval elapses, which can only be true if the
// filesystem is not consulted per lookup.
test('does not re-read the snapshot directory on every lookup', async () => {
  await withTempDir(async (dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);
    let nowMs = 1_000_000;
    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
      now: () => nowMs,
    });

    await waitForRefresh(() => lookup.get());
    assert.equal(lookup.get()?.forms.length, 2);

    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['アクア', 'あくあ'],
    ]);

    nowMs += 1000;
    assert.equal(lookup.get()?.forms.length, 2, 'expected the cached list within the interval');

    nowMs += 10_000;
    const refreshed = await waitForRefresh(() => {
      const candidates = lookup.get();
      return candidates && candidates.forms.length === 4 ? candidates : null;
    });
    assert.equal(refreshed.forms.length, 4, 'expected a refresh past the interval');
  });
});

test('invalidate picks up a snapshot change on the next refresh', async () => {
  await withTempDir(async (dir) => {
    writeSnapshot(dir, 1, [['ミナト', 'みなと']]);
    let nowMs = 1_000_000;
    const lookup = createCharacterNameCandidateLookup({
      outputDir: dir,
      getCurrentMediaId: () => 1,
      now: () => nowMs,
    });

    await waitForRefresh(() => lookup.get());
    assert.equal(lookup.get()?.forms.length, 2);

    writeSnapshot(dir, 1, [
      ['ミナト', 'みなと'],
      ['アクア', 'あくあ'],
    ]);
    nowMs += 1;
    lookup.invalidate();

    const refreshed = await waitForRefresh(() => {
      const candidates = lookup.get();
      return candidates && candidates.forms.length === 4 ? candidates : null;
    });
    assert.equal(refreshed.forms.length, 4);
  });
});
