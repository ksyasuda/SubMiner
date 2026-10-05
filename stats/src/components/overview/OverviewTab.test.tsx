import assert from 'node:assert/strict';
import test from 'node:test';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { apiClient } from '../../lib/api-client';
import { setDeleteConfirmPresenter } from '../../lib/delete-confirm';
import { localDayFromMs } from '../../lib/formatters';
import type { OverviewData, SessionSummary } from '../../types/stats';
import { OverviewTab } from './OverviewTab';
import { installDom } from '../../test-utils/dom';

function session(sessionId: number): SessionSummary {
  return {
    sessionId,
    canonicalTitle: `Test episode ${sessionId}`,
    videoId: null,
    animeId: null,
    animeTitle: 'Test anime',
    startedAtMs: Date.now(),
    endedAtMs: Date.now(),
    totalWatchedMs: 60_000,
    activeWatchedMs: 60_000,
    linesSeen: 1,
    tokensSeen: 10,
    cardsMined: 2,
    lookupCount: 0,
    lookupHits: 0,
    yomitanLookupCount: 0,
    knownWordsSeen: 1,
    knownWordRate: 0.1,
  };
}

function overview(sessions: SessionSummary[]): OverviewData {
  const count = sessions.length;
  return {
    sessions,
    rollups: count
      ? [
          {
            rollupDayOrMonth: localDayFromMs(Date.now()),
            videoId: null,
            totalSessions: count,
            totalActiveMin: count,
            totalLinesSeen: count,
            totalTokensSeen: count * 10,
            totalCards: count * 2,
            cardsPerHour: null,
            tokensPerMin: null,
            lookupHitRate: null,
          },
        ]
      : [],
    hints: {
      totalSessions: count,
      activeSessions: 0,
      episodesToday: count,
      activeAnimeCount: count ? 1 : 0,
      totalEpisodesWatched: count,
      totalAnimeCompleted: 0,
      totalActiveMin: count,
      activeDays: count ? 1 : 0,
      totalCards: count * 2,
      totalTokensSeen: count * 10,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
      newWordsToday: count,
      newWordsThisWeek: count,
    },
  };
}

function metric(container: HTMLElement, label: string): string {
  const element = [...container.querySelectorAll('div')].find(
    (element) => element.textContent === label,
  );
  assert.ok(element?.parentElement, `expected metric ${label}`);
  return (element.parentElement.textContent ?? '').replace(label, '').trim();
}

for (const mode of ['session', 'day', 'anime', 'failed'] as const) {
  test(`Overview refreshes all dependent data after ${mode} deletion`, async () => {
    const uninstallDom = installDom();
    const original = { ...apiClient };
    const restoreConfirm = setDeleteConfirmPresenter(() => true);
    let sessions =
      mode === 'session' || mode === 'failed'
        ? [session(1)]
        : [
            { ...session(1), animeId: 7 },
            { ...session(2), animeId: 7 },
          ];
    const reads = { overview: 0, sessions: 0, calendar: 0, words: 0 };
    const deletedIds: number[] = [];
    apiClient.getOverview = async () => {
      reads.overview++;
      return overview(sessions);
    };
    apiClient.getSessions = async () => {
      reads.sessions++;
      return sessions;
    };
    apiClient.getStreakCalendar = async () => {
      reads.calendar++;
      return [{ epochDay: localDayFromMs(Date.now()), totalActiveMin: sessions.length }];
    };
    apiClient.getKnownWordsSummary = async () => {
      reads.words++;
      return { totalUniqueWords: 10 + sessions.length, knownWordCount: 5 };
    };
    apiClient.getCoverImages = async () => ({ anime: {}, media: {} });
    async function deleteSessions(ids: number[]) {
      if (mode === 'failed') throw new Error('Deletion failed');
      deletedIds.push(...ids);
      sessions = sessions.filter((session) => !ids.includes(session.sessionId));
    }
    apiClient.deleteSession = async (id) => deleteSessions([id]);
    apiClient.deleteSessions = deleteSessions;
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    try {
      await act(async () =>
        root.render(
          <OverviewTab onNavigateToMediaDetail={() => {}} onNavigateToSession={() => {}} />,
        ),
      );
      const count = sessions.length;
      assert.match(metric(container, 'Cards Mined Today'), new RegExp(`^${count * 2}`));
      assert.match(metric(container, 'Known Words'), new RegExp(`/ ${10 + count}`));
      assert.equal(container.querySelectorAll('.bg-ctp-green\\/30').length, 1);
      const prefix =
        mode === 'day'
          ? 'Delete all sessions from '
          : mode === 'anime'
            ? 'Delete all sessions for '
            : 'Delete session ';
      const button = [...container.querySelectorAll('button')].find((button) =>
        button.getAttribute('aria-label')?.startsWith(prefix),
      );
      assert.ok(button, `expected ${mode} delete button`);
      await act(async () => button.click());
      if (mode === 'failed') {
        assert.match(container.textContent ?? '', /Deletion failed/);
        assert.match(metric(container, 'Cards Mined Today'), /^2/);
        assert.deepEqual(reads, { overview: 1, sessions: 1, calendar: 1, words: 1 });
        assert.deepEqual(deletedIds, []);
      } else {
        assert.deepEqual(deletedIds, count === 1 ? [1] : [1, 2]);
        for (const label of [
          'Cards Mined Today',
          'Episodes Today',
          'Sessions',
          'Cards Mined',
          'Words Today',
          'New Words Today',
        ]) {
          assert.equal(metric(container, label), '0', `${label} should reflect the deletion`);
        }
        assert.match(metric(container, 'Known Words'), /5\s*\/ 10/);
        assert.equal(container.querySelectorAll('.bg-ctp-green\\/30').length, 0);
        assert.match(container.textContent ?? '', /No sessions yet/);
        assert.deepEqual(reads, { overview: 2, sessions: 2, calendar: 2, words: 2 });
      }
    } finally {
      await act(async () => root.unmount());
      restoreConfirm();
      Object.assign(apiClient, original);
      uninstallDom();
    }
  });
}

for (const failedRead of ['overview', 'sessions', 'calendar'] as const) {
  test(`Overview stays available during a failed ${failedRead} refresh and can retry`, async () => {
    const uninstallDom = installDom();
    const original = { ...apiClient };
    const restoreConfirm = setDeleteConfirmPresenter(() => true);
    let sessions = [session(1)];
    let refreshing = false;
    let failRefresh = true;
    let rejectRefresh = (_error: Error) => {};
    const pendingRefresh = new Promise<never>((_resolve, reject) => {
      rejectRefresh = reject;
    });
    apiClient.getOverview = async () => {
      if (refreshing && failRefresh && failedRead === 'overview') return pendingRefresh;
      return overview(sessions);
    };
    apiClient.getSessions = async () => {
      if (refreshing && failRefresh && failedRead === 'sessions') return pendingRefresh;
      return sessions;
    };
    apiClient.getStreakCalendar = async () => {
      if (refreshing && failRefresh && failedRead === 'calendar') return pendingRefresh;
      return [];
    };
    apiClient.getKnownWordsSummary = async () => ({
      totalUniqueWords: 0,
      knownWordCount: 0,
    });
    apiClient.deleteSession = async () => {
      sessions = [];
      refreshing = true;
    };
    // A calendar-only failure still lets the overview totals refresh.
    const staleCards = failedRead === 'calendar' ? '0' : '2';
    const container = document.createElement('div');
    document.body.append(container);
    const root = createRoot(container);
    try {
      await act(async () =>
        root.render(
          <OverviewTab onNavigateToMediaDetail={() => {}} onNavigateToSession={() => {}} />,
        ),
      );
      const deleteButton = [...container.querySelectorAll('button')].find((button) =>
        button.getAttribute('aria-label')?.startsWith('Delete session '),
      );
      assert.ok(deleteButton);
      await act(async () => deleteButton.click());
      assert.match(container.textContent ?? '', /Tracking Snapshot/);
      assert.equal(metric(container, 'Cards Mined Today'), staleCards);
      assert.match(container.textContent ?? '', /No sessions yet/);

      await act(async () => rejectRefresh(new Error(`${failedRead} unavailable`)));
      assert.match(container.textContent ?? '', /Tracking Snapshot/);
      assert.equal(metric(container, 'Cards Mined Today'), staleCards);
      assert.match(container.querySelector('[role="alert"]')?.textContent ?? '', /unavailable/);

      const retryButton = [...container.querySelectorAll('button')].find(
        (button) => button.textContent === 'Retry',
      );
      assert.ok(retryButton, 'failed refresh should offer a retry');
      failRefresh = false;
      await act(async () => retryButton.click());
      assert.equal(metric(container, 'Cards Mined Today'), '0');
      assert.equal(container.querySelector('[role="alert"]'), null);
    } finally {
      await act(async () => root.unmount());
      restoreConfirm();
      Object.assign(apiClient, original);
      uninstallDom();
    }
  });
}
