import { useRef, useMemo, type KeyboardEvent } from 'react';
import { useTranslation } from '../../i18n';

export type TabId = 'overview' | 'anime' | 'trends' | 'vocabulary' | 'search' | 'sessions';

interface Tab {
  id: TabId;
  i18nKey: string;
}

const TAB_KEYS: Record<TabId, string> = {
  overview: 'stats.tab.overview',
  anime: 'stats.tab.library',
  trends: 'stats.tab.trends',
  vocabulary: 'stats.tab.vocabulary',
  search: 'stats.tab.search',
  sessions: 'stats.tab.sessions',
};

const TAB_IDS = Object.keys(TAB_KEYS) as TabId[];

interface TabBarProps {
  activeTab: TabId;
  onTabChange: (tabId: TabId) => void;
}

export function TabBar({ activeTab, onTabChange }: TabBarProps) {
  const { t } = useTranslation();
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([]);

  const tabs = useMemo<Tab[]>(() => TAB_IDS.map((id) => ({ id, i18nKey: TAB_KEYS[id] })), []);

  const activateAtIndex = (index: number) => {
    const tab = tabs[index];
    if (!tab) return;
    tabRefs.current[index]?.focus();
    onTabChange(tab.id);
  };

  const onTabKeyDown = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
      event.preventDefault();
      activateAtIndex((index + 1) % tabs.length);
      return;
    }
    if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      event.preventDefault();
      activateAtIndex((index - 1 + tabs.length) % tabs.length);
      return;
    }
    if (event.key === 'Home') {
      event.preventDefault();
      activateAtIndex(0);
      return;
    }
    if (event.key === 'End') {
      event.preventDefault();
      activateAtIndex(tabs.length - 1);
    }
  };

  return (
    <nav
      className="flex border-b border-ctp-surface1"
      role="tablist"
      aria-label={t('stats.tabsAria')}
      aria-orientation="horizontal"
    >
      {tabs.map((tab, index) => (
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
          {t(tab.i18nKey)}
        </button>
      ))}
    </nav>
  );
}
