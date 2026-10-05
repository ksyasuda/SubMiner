import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeleteProgressToast } from '../components/common/DeleteProgressToast';
import { apiClient } from './api-client';
import {
  beginDeleteTask,
  getDeleteProgressSnapshot,
  resetDeleteProgress,
  subscribeDeleteProgress,
  trackDelete,
} from './delete-progress';

test('delete progress store reports the oldest label and notifies subscribers', () => {
  resetDeleteProgress();
  let notifications = 0;
  const unsubscribe = subscribeDeleteProgress(() => {
    notifications += 1;
  });

  try {
    assert.deepEqual(getDeleteProgressSnapshot(), { count: 0, label: null });

    const endFirst = beginDeleteTask('Deleting session');
    const endSecond = beginDeleteTask('Deleting episode');
    assert.deepEqual(getDeleteProgressSnapshot(), { count: 2, label: 'Deleting session' });
    assert.equal(notifications, 2);

    endFirst();
    assert.deepEqual(getDeleteProgressSnapshot(), { count: 1, label: 'Deleting episode' });

    endFirst();
    assert.deepEqual(
      getDeleteProgressSnapshot(),
      { count: 1, label: 'Deleting episode' },
      'ending the same task twice must not drop another task',
    );

    endSecond();
    assert.deepEqual(getDeleteProgressSnapshot(), { count: 0, label: null });
  } finally {
    unsubscribe();
    resetDeleteProgress();
  }
});

test('trackDelete clears the indicator even when the request rejects', async () => {
  resetDeleteProgress();
  await assert.rejects(
    trackDelete('Deleting library entry', async () => {
      assert.equal(getDeleteProgressSnapshot().count, 1);
      throw new Error('boom');
    }),
    /boom/,
  );
  assert.deepEqual(getDeleteProgressSnapshot(), { count: 0, label: null });
});

test('DeleteProgressToast stays hidden while idle and reports active deletes', () => {
  resetDeleteProgress();
  assert.equal(renderToStaticMarkup(<DeleteProgressToast />), '');

  const end = beginDeleteTask('Deleting library entry');
  try {
    const markup = renderToStaticMarkup(<DeleteProgressToast />);
    assert.match(markup, /role="status"/);
    assert.match(markup, /Deleting library entry/);
    assert.match(markup, /animate-indeterminate/);
  } finally {
    end();
  }

  const endFirst = beginDeleteTask('Deleting session');
  const endSecond = beginDeleteTask('Deleting episode');
  try {
    assert.match(renderToStaticMarkup(<DeleteProgressToast />), /Deleting 2 items/);
  } finally {
    endFirst();
    endSecond();
    resetDeleteProgress();
  }
});

test('every api client delete registers with the global progress indicator', async () => {
  const originalFetch = globalThis.fetch;
  const seenCounts: number[] = [];
  globalThis.fetch = (async () => {
    seenCounts.push(getDeleteProgressSnapshot().count);
    return new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } });
  }) as typeof globalThis.fetch;

  try {
    resetDeleteProgress();
    await apiClient.deleteSession(1);
    await apiClient.deleteSessions([1, 2]);
    await apiClient.deleteVideo(3);
    await apiClient.deleteAnime(4);

    assert.deepEqual(seenCounts, [1, 1, 1, 1]);
    assert.deepEqual(getDeleteProgressSnapshot(), { count: 0, label: null });
  } finally {
    globalThis.fetch = originalFetch;
    resetDeleteProgress();
  }
});
