import { useState, useCallback } from 'react';
import { TabBar } from './components/layout/TabBar';
import { OverviewTab } from './components/overview/OverviewTab';
import { TrendsTab } from './components/trends/TrendsTab';
import { AnimeTab } from './components/anime/AnimeTab';
import { MediaDetailView } from './components/library/MediaDetailView';
import { VocabularyTab } from './components/vocabulary/VocabularyTab';
import { SessionsTab } from './components/sessions/SessionsTab';
import { WordDetailPanel } from './components/vocabulary/WordDetailPanel';
import { useExcludedWords } from './hooks/useExcludedWords';
import type { TabId } from './components/layout/TabBar';
import {
  closeMediaDetail,
  createInitialStatsView,
  navigateToAnime as navigateToAnimeState,
  navigateToSession as navigateToSessionState,
  openAnimeEpisodeDetail,
  openOverviewMediaDetail,
  switchTab,
} from './lib/stats-navigation';

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
    <div className="min-h-screen flex flex-col bg-ctp-base">
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
      <main className="flex-1 overflow-y-auto p-4">
        {mediaDetail ? (
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
              mediaDetail.origin.type === 'overview' ? 'Back to Overview' : 'Back to Library'
            }
          />
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
                <AnimeTab
                  initialAnimeId={selectedAnimeId}
                  onClearInitialAnime={() =>
                    setViewState((prev) => ({ ...prev, selectedAnimeId: null }))
                  }
                  onNavigateToWord={openWordDetail}
                  onOpenEpisodeDetail={navigateToEpisodeDetail}
                />
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
                <TrendsTab />
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
                <VocabularyTab
                  onNavigateToAnime={navigateToAnime}
                  onOpenWordDetail={openWordDetail}
                  excluded={excluded}
                  isExcluded={isExcluded}
                  onRemoveExclusion={removeExclusion}
                  onClearExclusions={clearAll}
                />
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
                <SessionsTab
                  initialSessionId={focusedSessionId}
                  onClearInitialSession={() =>
                    setViewState((prev) => ({ ...prev, focusedSessionId: null }))
                  }
                />
              </section>
            ) : null}
          </>
        )}
      </main>
      <WordDetailPanel
        wordId={globalWordId}
        onClose={() => setGlobalWordId(null)}
        onSelectWord={openWordDetail}
        onNavigateToAnime={navigateToAnime}
        isExcluded={isExcluded}
        onToggleExclusion={toggleExclusion}
      />
    </div>
  );
}
