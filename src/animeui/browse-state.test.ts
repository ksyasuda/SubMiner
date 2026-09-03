import test from 'node:test';
import assert from 'node:assert/strict';
import {
  beginBrowse,
  beginNextPage,
  capture,
  createBrowseState,
  failBrowse,
  finishBrowse,
  LatestRequest,
  safeUploadDate,
  soleBrowseRequest,
  takeUnseenEntries,
} from './browse-state';
import type { AnimeBrowserEntry } from '../types/anime-browser';

test('a new browse replaces results and starts from page one', () => {
  const started = beginBrowse(createBrowseState(), 'frieren');

  assert.deepEqual(started.request, {
    id: 1,
    query: 'frieren',
    page: 1,
    append: false,
  });
  assert.equal(started.state.loading, true);
});

test('the next page appends only when the current result has another page', () => {
  const first = beginBrowse(createBrowseState(), '');
  const ready = finishBrowse(first.state, first.request.id, true);
  const next = beginNextPage(ready);

  assert.ok(next);
  assert.deepEqual(next.request, {
    id: 2,
    query: '',
    page: 2,
    append: true,
  });
  assert.equal(beginNextPage(next.state), null, 'cannot overlap page requests');
  assert.equal(beginNextPage(finishBrowse(next.state, next.request.id, false)), null);
});

test('a stale page completion cannot change the active browse state', () => {
  const first = beginBrowse(createBrowseState(), 'old');
  const current = beginBrowse(first.state, 'new');

  assert.equal(finishBrowse(current.state, first.request.id, true), current.state);
});

test('stream updates are correlated only when one browse request is in flight', () => {
  const first = beginBrowse(createBrowseState(), 'old');
  const second = beginBrowse(first.state, 'new');
  const inFlight = new Map([
    [second.request.id, second.request],
    [first.request.id, first.request],
  ]);

  assert.equal(soleBrowseRequest(inFlight), null, 'start order is ambiguous while calls overlap');
  inFlight.delete(first.request.id);
  assert.equal(soleBrowseRequest(inFlight), second.request);
});

test('a failed next page remains retryable', () => {
  const first = beginBrowse(createBrowseState(), '');
  const ready = finishBrowse(first.state, first.request.id, true);
  const next = beginNextPage(ready)!;
  const failed = failBrowse(next.state, next.request);
  const retry = beginNextPage(failed);

  assert.ok(retry);
  assert.equal(retry.request.page, 2);
});

test('latest request tokens invalidate closed and superseded details', () => {
  const requests = new LatestRequest();
  const first = requests.begin();
  const second = requests.begin();

  assert.equal(requests.isCurrent(first), false);
  assert.equal(requests.isCurrent(second), true);
  requests.cancel();
  assert.equal(requests.isCurrent(second), false);
});

test('safeUploadDate ignores malformed timestamps', () => {
  assert.equal(safeUploadDate(Date.UTC(2025, 3, 2)), '2025-04-02');
  assert.equal(safeUploadDate(Number.NaN), null);
  assert.equal(safeUploadDate(Number.POSITIVE_INFINITY), null);
});

test('capture turns a rejected IPC operation into a displayable failure', async () => {
  const failure = new Error('bridge disconnected');
  const result = await capture(async () => Promise.reject(failure));

  assert.deepEqual(result, { ok: false, error: failure });
});

test('takeUnseenEntries deduplicates streamed and final-page entries', () => {
  const entry = (sourceId: string, url: string): AnimeBrowserEntry => ({
    sourceId,
    sourceName: sourceId,
    url,
    title: url,
    thumbnailUrl: null,
  });
  const seen = new Set<string>();

  assert.deepEqual(takeUnseenEntries([entry('one', '/a')], seen), [entry('one', '/a')]);
  assert.deepEqual(takeUnseenEntries([entry('one', '/a'), entry('two', '/a')], seen), [
    entry('two', '/a'),
  ]);
});
