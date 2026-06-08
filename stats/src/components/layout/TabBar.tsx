import { useRef, type KeyboardEvent } from 'react';

export type TabId = 'overview' | 'anime' | 'trends' | 'vocabulary' | 'search' | 'sessions';

interface Tab {
  id: TabId;
  label: string;
}

const TABS: Tab[] = [
  { id: 'overview', label: 'Overview' },
  { id: 'anime', label: 'Library' },
  { id: 'trends', label: 'Trends' },
  { id: 'vocabulary', label: 'Vocabulary' },
  { id: 'search', label: 'Search' },
  { id: 'sessions', label: 'Sessions' },
];

interface TabBarProps {
  activeTab: TabId;
  onTabChange: (tabId: TabId) => void;
}

export function TabBar({ activeTab, onTabChange }: TabBarProps) {
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([]);

  const activateAtIndex = (index: number) => {
    const tab = TABS[index];
    if (!tab) return;
    tabRefs.current[index]?.focus();
    onTabChange(tab.id);
  };

  const onTabKeyDown = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      event.preventDefault();
      activateAtIndex((index + 1) % TABS.length);
      return;
    }
    if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      event.preventDefault();
      activateAtIndex((index - 1 + TABS.length) % TABS.length);
      return;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      activateAtIndex(0);
      return;
    }
    if (event.key === 'End') {
      event.preventDefault();
      activateAtIndex(TABS.length - 1);
    }
  };

  return (
    <nav
      className="flex border-b border-ctp-surface1"
      role="tablist"
      aria-label="Stats tabs"
      aria-orientation="horizontal"
    >
      {TABS.map((tab, index) => (
        <button
          key={tab.id}
          id={`tab-${tab.id}`}
          ref={(element) => {
            tabRefs.current[index] = element;
          }}
          type="button"
          role="tab"
          aria-controls={`panel-${tab.id}`}
          aria-selected={activeTab === tab.id}
          tabIndex={activeTab === tab.id ? 0 : -1}
          onClick={() => onTabChange(tab.id)}
          onKeyDown={(event) => onTabKeyDown(event, index)}
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
