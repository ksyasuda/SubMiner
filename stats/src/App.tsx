import { useState, useCallback } from 'react';
import { TabBar } from './components/layout/TabBar';
import { OverviewTab } from './components/overview/OverviewTab';
import { TrendsTab } from './components/trends/TrendsTab';
import { AnimeTab } from './components/anime/AnimeTab';
import { LibraryTab } from './components/library/LibraryTab';
import { VocabularyTab } from './components/vocabulary/VocabularyTab';
import { SessionsTab } from './components/sessions/SessionsTab';
import { WordDetailPanel } from './components/vocabulary/WordDetailPanel';
import { useExcludedWords } from './hooks/useExcludedWords';
import type { TabId } from './components/layout/TabBar';

export function App() {
  const [activeTab, setActiveTab] = useState<TabId>('overview');
  const [mountedTabs, setMountedTabs] = useState<Set<TabId>>(() => new Set(['overview']));
  const [selectedAnimeId, setSelectedAnimeId] = useState<number | null>(null);
  const [focusedSessionId, setFocusedSessionId] = useState<number | null>(null);
  const [globalWordId, setGlobalWordId] = useState<number | null>(null);
  const { excluded, isExcluded, toggleExclusion, removeExclusion, clearAll } = useExcludedWords();

  const activateTab = useCallback((tabId: TabId) => {
    setActiveTab(tabId);
    setMountedTabs((prev) => {
      if (prev.has(tabId)) return prev;
      const next = new Set(prev);
      next.add(tabId);
      return next;
    });
  }, []);

  const navigateToAnime = useCallback((animeId: number) => {
    activateTab('anime');
    setSelectedAnimeId(animeId);
  }, [activateTab]);

  const navigateToSession = useCallback((sessionId: number) => {
    activateTab('sessions');
    setFocusedSessionId(sessionId);
  }, [activateTab]);

  const openWordDetail = useCallback((wordId: number) => {
    setGlobalWordId(wordId);
  }, []);

  const handleTabChange = useCallback((tabId: TabId) => {
    activateTab(tabId);
    setSelectedAnimeId(null);
    if (tabId !== 'sessions') {
      setFocusedSessionId(null);
    }
  }, [activateTab]);

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
        {mountedTabs.has('overview') ? (
          <section
            id="panel-overview"
            role="tabpanel"
            aria-labelledby="tab-overview"
            hidden={activeTab !== 'overview'}
            className="animate-fade-in"
          >
            <OverviewTab onNavigateToSession={navigateToSession} />
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
              onClearInitialAnime={() => setSelectedAnimeId(null)}
              onNavigateToWord={openWordDetail}
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
        {mountedTabs.has('library') ? (
          <section
            id="panel-library"
            role="tabpanel"
            aria-labelledby="tab-library"
            hidden={activeTab !== 'library'}
            className="animate-fade-in"
          >
            <LibraryTab />
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
              onClearInitialSession={() => setFocusedSessionId(null)}
            />
          </section>
        ) : null}
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
