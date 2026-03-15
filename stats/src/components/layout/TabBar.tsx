export type TabId = 'overview' | 'anime' | 'trends' | 'vocabulary' | 'sessions';

interface Tab {
  id: TabId;
  label: string;
}

const TABS: Tab[] = [
  { id: 'overview', label: 'Overview' },
  { id: 'anime', label: 'Anime' },
  { id: 'trends', label: 'Trends' },
  { id: 'vocabulary', label: 'Vocabulary' },
  { id: 'sessions', label: 'Sessions' },
];

interface TabBarProps {
  activeTab: TabId;
  onTabChange: (tabId: TabId) => void;
}

export function TabBar({ activeTab, onTabChange }: TabBarProps) {
  return (
    <nav className="flex border-b border-ctp-surface1" role="tablist" aria-label="Stats tabs">
      {TABS.map((tab) => (
        <button
          key={tab.id}
          id={`tab-${tab.id}`}
          type="button"
          role="tab"
          aria-controls={`panel-${tab.id}`}
          aria-selected={activeTab === tab.id}
          tabIndex={activeTab === tab.id ? 0 : -1}
          onClick={() => onTabChange(tab.id)}
          className={`px-4 py-2.5 text-sm font-medium transition-colors
            ${
              activeTab === tab.id
                ? 'text-ctp-text border-b-2 border-ctp-lavender'
                : 'text-ctp-subtext0 hover:text-ctp-subtext1'
            }`}
        >
          {tab.label}
        </button>
      ))}
    </nav>
  );
}
