import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
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

test('the delete indicator is mounted once at the app root, not inside tab panels', () => {
  const srcDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const read = (relativePath: string): string =>
    fs.readFileSync(path.join(srcDir, relativePath), 'utf8');

  const app = read('App.tsx');
  assert.match(app, /<DeleteProgressToast \/>/);
  // Sibling of the confirm dialog: outside every `hidden` tab panel, so the
  // indicator survives tab switches and detail-view navigation.
  assert.match(app, /<DeleteConfirmDialog \/>\s*<DeleteProgressToast \/>/);

  for (const tab of [
    'components/overview/OverviewTab.tsx',
    'components/sessions/SessionsTab.tsx',
  ]) {
    assert.doesNotMatch(read(tab), /DeleteProgressToast/, `${tab} must not mount its own toast`);
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
