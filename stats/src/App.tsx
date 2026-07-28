import { Suspense, lazy, useCallback, useState } from 'react';
import { DeleteConfirmDialog } from './components/layout/DeleteConfirmDialog';
import { DeleteProgressToast } from './components/common/DeleteProgressToast';
import { TabBar } from './components/layout/TabBar';
import { OverviewTab } from './components/overview/OverviewTab';
import { useExcludedWords } from './hooks/useExcludedWords';
import type { TabId } from './components/layout/TabBar';
import {
  closeMediaDetail,
  createInitialStatsView,
  navigateToAnime as navigateToAnimeState,
  navigateToSession as navigateToSessionState,
  openAnimeEpisodeDetail,
  openOverviewMediaDetail,
  openSessionsMediaDetail,
  switchTab,
} from './lib/stats-navigation';

const AnimeTab = lazy(() =>
  import('./components/anime/AnimeTab').then((module) => ({
    default: module.AnimeTab,
  })),
);
const TrendsTab = lazy(() =>
  import('./components/trends/TrendsTab').then((module) => ({
    default: module.TrendsTab,
  })),
);
const VocabularyTab = lazy(() =>
  import('./components/vocabulary/VocabularyTab').then((module) => ({
    default: module.VocabularyTab,
  })),
);
const SearchTab = lazy(() =>
  import('./components/search/SearchTab').then((module) => ({
    default: module.SearchTab,
  })),
);
const SessionsTab = lazy(() =>
  import('./components/sessions/SessionsTab').then((module) => ({
    default: module.SessionsTab,
  })),
);
const MediaDetailView = lazy(() =>
  import('./components/library/MediaDetailView').then((module) => ({
    default: module.MediaDetailView,
  })),
);
const WordDetailPanel = lazy(() =>
  import('./components/vocabulary/WordDetailPanel').then((module) => ({
    default: module.WordDetailPanel,
  })),
);

function LoadingSurface({ label, overlay = false }: { label: string; overlay?: boolean }) {
  return (
    <div
      aria-busy="true"
      className={
        overlay
          ? 'fixed inset-x-4 bottom-4 z-50 rounded-xl border border-ctp-surface1 bg-ctp-mantle/95 p-4 text-sm text-ctp-overlay2 shadow-2xl backdrop-blur'
          : 'rounded-xl border border-ctp-surface1 bg-ctp-surface0/70 p-4 text-sm text-ctp-overlay2'
      }
    >
      {label}
    </div>
  );
}

export function App() {
  const [viewState, setViewState] = useState(createInitialStatsView);
  const [mountedTabs, setMountedTabs] = useState<Set<TabId>>(() => new Set(['overview']));
  const [globalWordId, setGlobalWordId] = useState<number | null>(null);
  const { excluded, isExcluded, toggleExclusion, removeExclusion, clearAll } = useExcludedWords();
  const { activeTab, selectedAnimeId, focusedSessionId, mediaDetail } = viewState;

  const activateTab = useCallback((tabId: TabId) => {
    setViewState((prev) => switchTab(prev, tabId));
    setMountedTabs((prev) => {
      if (prev.has(tabId)) return prev;
      const next = new Set(prev);
      next.add(tabId);
      return next;
    });
  }, []);

  const navigateToAnime = useCallback((animeId: number) => {
    setViewState((prev) => navigateToAnimeState(prev, animeId));
    setMountedTabs((prev) => {
      if (prev.has('anime')) return prev;
      const next = new Set(prev);
      next.add('anime');
      return next;
    });
  }, []);

  const navigateToSession = useCallback((sessionId: number) => {
    setViewState((prev) => navigateToSessionState(prev, sessionId));
    setMountedTabs((prev) => {
      if (prev.has('sessions')) return prev;
      const next = new Set(prev);
      next.add('sessions');
      return next;
    });
  }, []);

  const navigateToEpisodeDetail = useCallback(
    (animeId: number, videoId: number, sessionId: number | null = null) => {
      setViewState((prev) => openAnimeEpisodeDetail(prev, animeId, videoId, sessionId));
    },
    [],
  );

  const navigateToOverviewMediaDetail = useCallback(
    (videoId: number, sessionId: number | null = null) => {
      setViewState((prev) => openOverviewMediaDetail(prev, videoId, sessionId));
    },
    [],
  );

  const navigateToSessionsMediaDetail = useCallback((videoId: number) => {
    setViewState((prev) => openSessionsMediaDetail(prev, videoId));
  }, []);

  const openWordDetail = useCallback((wordId: number) => {
    setGlobalWordId(wordId);
  }, []);

  const handleTabChange = useCallback(
    (tabId: TabId) => {
      activateTab(tabId);
    },
    [activateTab],
  );

  return (
    <div className="h-screen min-h-0 overflow-hidden flex flex-col bg-ctp-base">
      <header className="px-4 pt-3 pb-0">
        <button
          type="button"
          onClick={() => handleTabChange('overview')}
          className="flex items-center gap-2 mb-2 hover:opacity-80 transition-opacity"
        >
          <img src="/favicon.png" alt="" className="h-6 object-contain" />
          <h1 className="text-lg font-semibold text-ctp-text">SubMiner Stats</h1>
        </button>
        <TabBar activeTab={activeTab} onTabChange={handleTabChange} />
      </header>
      <main className="flex-1 min-h-0 overflow-y-auto p-4">
        {mediaDetail ? (
          <Suspense fallback={<LoadingSurface label="Loading media detail..." />}>
            <MediaDetailView
              videoId={mediaDetail.videoId}
              initialExpandedSessionId={mediaDetail.initialSessionId}
              onConsumeInitialExpandedSession={() =>
                setViewState((prev) =>
                  prev.mediaDetail
                    ? {
                        ...prev,
                        mediaDetail: {
                          ...prev.mediaDetail,
                          initialSessionId: null,
                        },
                      }
                    : prev,
                )
              }
              onBack={() => setViewState((prev) => closeMediaDetail(prev))}
              backLabel={
                mediaDetail.origin.type === 'overview'
                  ? 'Back to Overview'
                  : mediaDetail.origin.type === 'sessions'
                    ? 'Back to Sessions'
                    : 'Back to Library'
              }
              onNavigateToAnime={navigateToAnime}
            />
          </Suspense>
        ) : (
          <>
            {mountedTabs.has('overview') ? (
              <section
                id="panel-overview"
                role="tabpanel"
                aria-labelledby="tab-overview"
                hidden={activeTab !== 'overview'}
                className="animate-fade-in"
              >
                <OverviewTab
                  onNavigateToMediaDetail={navigateToOverviewMediaDetail}
                  onNavigateToSession={navigateToSession}
                  isActive={activeTab === 'overview'}
                />
              </section>
            ) : null}
            {mountedTabs.has('anime') ? (
              <section
                id="panel-anime"
                role="tabpanel"
                aria-labelledby="tab-anime"
                hidden={activeTab !== 'anime'}
                className="animate-fade-in"
              >
                <Suspense fallback={<LoadingSurface label="Loading library..." />}>
                  <AnimeTab
                    initialAnimeId={selectedAnimeId}
                    onClearInitialAnime={() =>
                      setViewState((prev) => ({ ...prev, selectedAnimeId: null }))
                    }
                    onNavigateToWord={openWordDetail}
                    onOpenEpisodeDetail={navigateToEpisodeDetail}
                  />
                </Suspense>
              </section>
            ) : null}
            {mountedTabs.has('trends') ? (
              <section
                id="panel-trends"
                role="tabpanel"
                aria-labelledby="tab-trends"
                hidden={activeTab !== 'trends'}
                className="animate-fade-in"
              >
                <Suspense fallback={<LoadingSurface label="Loading trends..." />}>
                  <TrendsTab />
                </Suspense>
              </section>
            ) : null}
            {mountedTabs.has('vocabulary') ? (
              <section
                id="panel-vocabulary"
                role="tabpanel"
                aria-labelledby="tab-vocabulary"
                hidden={activeTab !== 'vocabulary'}
                className="animate-fade-in"
              >
                <Suspense fallback={<LoadingSurface label="Loading vocabulary..." />}>
                  <VocabularyTab
                    onNavigateToAnime={navigateToAnime}
                    onOpenWordDetail={openWordDetail}
                    excluded={excluded}
                    isExcluded={isExcluded}
                    onRemoveExclusion={removeExclusion}
                    onClearExclusions={clearAll}
                  />
                </Suspense>
              </section>
            ) : null}
            {mountedTabs.has('search') ? (
              <section
                id="panel-search"
                role="tabpanel"
                aria-labelledby="tab-search"
                hidden={activeTab !== 'search'}
                className="animate-fade-in"
              >
                <Suspense fallback={<LoadingSurface label="Loading search..." />}>
                  <SearchTab />
                </Suspense>
              </section>
            ) : null}
            {mountedTabs.has('sessions') ? (
              <section
                id="panel-sessions"
                role="tabpanel"
                aria-labelledby="tab-sessions"
                hidden={activeTab !== 'sessions'}
                className="animate-fade-in"
              >
                <Suspense fallback={<LoadingSurface label="Loading sessions..." />}>
                  <SessionsTab
                    initialSessionId={focusedSessionId}
                    onClearInitialSession={() =>
                      setViewState((prev) => ({ ...prev, focusedSessionId: null }))
                    }
                    onNavigateToMediaDetail={navigateToSessionsMediaDetail}
                  />
                </Suspense>
              </section>
            ) : null}
          </>
        )}
      </main>
      {globalWordId !== null ? (
        <Suspense fallback={<LoadingSurface label="Loading word detail..." overlay />}>
          <WordDetailPanel
            wordId={globalWordId}
            onClose={() => setGlobalWordId(null)}
            onSelectWord={openWordDetail}
            onNavigateToAnime={navigateToAnime}
            isExcluded={isExcluded}
            onToggleExclusion={toggleExclusion}
          />
        </Suspense>
      ) : null}
      <DeleteConfirmDialog />
      <DeleteProgressToast />
    </div>
  );
}
