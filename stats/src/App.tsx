import { useState, useCallback } from 'react';
import { TabBar } from './components/layout/TabBar';
import { OverviewTab } from './components/overview/OverviewTab';
import { TrendsTab } from './components/trends/TrendsTab';
import { AnimeTab } from './components/anime/AnimeTab';
import { VocabularyTab } from './components/vocabulary/VocabularyTab';
import { SessionsTab } from './components/sessions/SessionsTab';
import { WordDetailPanel } from './components/vocabulary/WordDetailPanel';
import { useExcludedWords } from './hooks/useExcludedWords';
import type { TabId } from './components/layout/TabBar';

export function App() {
  const [activeTab, setActiveTab] = useState<TabId>('overview');
  const [selectedAnimeId, setSelectedAnimeId] = useState<number | null>(null);
  const [globalWordId, setGlobalWordId] = useState<number | null>(null);
  const { excluded, isExcluded, toggleExclusion, removeExclusion, clearAll } = useExcludedWords();

  const navigateToAnime = useCallback((animeId: number) => {
    setActiveTab('anime');
    setSelectedAnimeId(animeId);
  }, []);

  const openWordDetail = useCallback((wordId: number) => {
    setGlobalWordId(wordId);
  }, []);

  const handleTabChange = useCallback((tabId: TabId) => {
    setActiveTab(tabId);
    setSelectedAnimeId(null);
  }, []);

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
        {activeTab === 'overview' ? (
          <section id="panel-overview" role="tabpanel" aria-labelledby="tab-overview" key="overview" className="animate-fade-in">
            <OverviewTab />
          </section>
        ) : null}
        {activeTab === 'anime' ? (
          <section id="panel-anime" role="tabpanel" aria-labelledby="tab-anime" key="anime" className="animate-fade-in">
            <AnimeTab
              initialAnimeId={selectedAnimeId}
              onClearInitialAnime={() => setSelectedAnimeId(null)}
              onNavigateToWord={openWordDetail}
            />
          </section>
        ) : null}
        {activeTab === 'trends' ? (
          <section id="panel-trends" role="tabpanel" aria-labelledby="tab-trends" key="trends" className="animate-fade-in">
            <TrendsTab />
          </section>
        ) : null}
        {activeTab === 'vocabulary' ? (
          <section id="panel-vocabulary" role="tabpanel" aria-labelledby="tab-vocabulary" key="vocabulary" className="animate-fade-in">
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
        {activeTab === 'sessions' ? (
          <section id="panel-sessions" role="tabpanel" aria-labelledby="tab-sessions" key="sessions" className="animate-fade-in">
            <SessionsTab />
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
