import test from 'node:test';
import assert from 'node:assert/strict';
import { applySearchUpdate, idleSearchProgress, summarizeProgress } from './search-progress';
import type { AnimeBrowserEntry, SourceSearchFailure } from '../types/anime-browser';

const entry = (title: string): AnimeBrowserEntry => ({
  url: `/a/${title}`,
  title,
  thumbnailUrl: null,
  sourceId: 's1',
  sourceName: 'Source One',
});

const failure: SourceSearchFailure = {
  sourceId: 's2',
  sourceName: 'Source Two',
  error: 'login required',
};

test('a search accumulates results and failures update by update', () => {
  let progress = idleSearchProgress();

  let applied = applySearchUpdate(progress, { kind: 'start', token: 1, sourceCount: 3 });
  assert.ok(applied);
  assert.equal(applied.started, true);
  progress = applied.progress;

  applied = applySearchUpdate(progress, {
    kind: 'result',
    token: 1,
    sourceId: 's1',
    sourceName: 'Source One',
    entries: [entry('A'), entry('B')],
  });
  assert.ok(applied);
  assert.deepEqual(
    applied.entries.map((item) => item.title),
    ['A', 'B'],
  );
  progress = applied.progress;

  applied = applySearchUpdate(progress, { kind: 'failure', token: 1, failure });
  assert.ok(applied);
  progress = applied.progress;

  assert.equal(progress.entryCount, 2);
  assert.equal(progress.sourcesDone, 2);
  assert.deepEqual(progress.failures, [failure]);
  assert.equal(progress.done, false);

  applied = applySearchUpdate(progress, { kind: 'done', token: 1 });
  assert.ok(applied);
  assert.equal(applied.progress.done, true);
});

test('updates from a superseded search are dropped entirely', () => {
  let progress = idleSearchProgress();
  progress = applySearchUpdate(progress, { kind: 'start', token: 1, sourceCount: 2 })!.progress;
  progress = applySearchUpdate(progress, { kind: 'start', token: 2, sourceCount: 2 })!.progress;

  // The first search's straggler results and completion must not touch token 2.
  assert.equal(
    applySearchUpdate(progress, {
      kind: 'result',
      token: 1,
      sourceId: 's1',
      sourceName: 'Source One',
      entries: [entry('stale')],
    }),
    null,
  );
  assert.equal(applySearchUpdate(progress, { kind: 'done', token: 1 }), null);
  assert.equal(progress.entryCount, 0);
});

test('an older start cannot reset a newer search', () => {
  let progress = idleSearchProgress();
  progress = applySearchUpdate(progress, { kind: 'start', token: 5, sourceCount: 1 })!.progress;

  assert.equal(applySearchUpdate(progress, { kind: 'start', token: 4, sourceCount: 9 }), null);
  assert.equal(applySearchUpdate(progress, { kind: 'start', token: 5, sourceCount: 9 }), null);
  assert.equal(progress.sourceCount, 1);
});

test('summarizeProgress counts sources and names failures', () => {
  let progress = idleSearchProgress();
  progress = applySearchUpdate(progress, { kind: 'start', token: 1, sourceCount: 5 })!.progress;
  progress = applySearchUpdate(progress, {
    kind: 'result',
    token: 1,
    sourceId: 's1',
    sourceName: 'Source One',
    entries: [entry('A')],
  })!.progress;

  assert.equal(summarizeProgress(progress), 'Searching… 1/5 sources · 1 result');

  progress = applySearchUpdate(progress, { kind: 'failure', token: 1, failure })!.progress;
  assert.equal(
    summarizeProgress(progress),
    'Searching… 2/5 sources · 1 result · unavailable: Source Two',
  );
});
