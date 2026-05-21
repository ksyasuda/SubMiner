import {
  filterSessionHelpSections,
  getSessionHelpSectionTabId,
  SESSION_HELP_TABS,
  type SessionHelpSection,
  type SessionHelpTabId,
} from './session-help-sections';

function countRows(sections: SessionHelpSection[]): number {
  return sections.reduce((count, section) => count + section.rows.length, 0);
}

function sectionMatchesTab(section: SessionHelpSection, tabId: SessionHelpTabId): boolean {
  return getSessionHelpSectionTabId(section) === tabId;
}

export function buildVisibleSessionHelpSections(
  sections: SessionHelpSection[],
  tabId: SessionHelpTabId,
  query: string,
): SessionHelpSection[] {
  if (query.trim()) return filterSessionHelpSections(sections, query);
  return sections.filter((section) => sectionMatchesTab(section, tabId));
}

export function createSessionHelpTabBar(
  sections: SessionHelpSection[],
  activeTabId: SessionHelpTabId,
  onSelect: (tabId: SessionHelpTabId) => void,
): HTMLElement {
  const tabBar = document.createElement('div');
  tabBar.className = 'session-help-tabs';

  for (const tab of SESSION_HELP_TABS) {
    const tabSections = sections.filter((section) => sectionMatchesTab(section, tab.id));
    if (tabSections.length === 0) continue;

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'session-help-tab';
    button.dataset.sessionHelpTab = tab.id;
    button.setAttribute('aria-pressed', String(tab.id === activeTabId));
    if (tab.id === activeTabId) button.classList.add('active');
    button.textContent = `${tab.label} ${countRows(tabSections)}`;
    button.addEventListener('click', () => onSelect(tab.id));
    tabBar.appendChild(button);
  }

  return tabBar;
}
