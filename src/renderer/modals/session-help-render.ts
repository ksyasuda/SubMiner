import type { SessionHelpItem, SessionHelpSection } from './session-help-sections';
import { i18n } from '../../i18n/index.js';

function createShortcutRow(row: SessionHelpItem, globalIndex: number): HTMLButtonElement {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'session-help-item';
  button.tabIndex = -1;
  button.dataset.sessionHelpIndex = String(globalIndex);

  const left = document.createElement('div');
  left.className = 'session-help-item-left';
  const shortcut = document.createElement('span');
  shortcut.className = 'session-help-key';
  shortcut.textContent = row.shortcut;
  left.appendChild(shortcut);

  const right = document.createElement('div');
  right.className = 'session-help-item-right';
  const action = document.createElement('span');
  action.className = 'session-help-action';
  action.textContent = row.action;
  right.appendChild(action);

  if (row.color) {
    const dot = document.createElement('span');
    dot.className = 'session-help-color-dot';
    dot.style.backgroundColor = row.color;
    right.insertBefore(dot, action);
  }

  button.appendChild(left);
  button.appendChild(right);
  return button;
}

const SECTION_ICON: Record<string, string> = {
  [i18n.t('sessionHelp.sections.playback')]: '▶',
  [i18n.t('sessionHelp.sections.visual')]: '◉',
  [i18n.t('sessionHelp.sections.sync')]: '⟲',
  [i18n.t('sessionHelp.section.mining')]: '✦',
  [i18n.t('sessionHelp.section.stats')]: '◉',
  [i18n.t('sessionHelp.section.overlayControls')]: '◈',
  [i18n.t('sessionHelp.section.modals')]: '▣',
  [i18n.t('sessionHelp.sections.runtime')]: '⚙',
  [i18n.t('sessionHelp.sections.system')]: '◆',
  [i18n.t('sessionHelp.sections.other')]: '…',
  [i18n.t('sessionHelp.section.fixedOverlay')]: '◇',
  [i18n.t('sessionHelp.section.yChords')]: '⌘',
  [i18n.t('sessionHelp.section.globalShortcuts')]: '◆',
  [i18n.t('sessionHelp.colorLegend')]: '◈',
};

export function createSessionHelpSectionNode(
  section: SessionHelpSection,
  sectionIndex: number,
  globalIndexMap: number[],
): HTMLElement {
  const sectionNode = document.createElement('section');
  sectionNode.className = 'session-help-section';

  const title = document.createElement('h3');
  title.className = 'session-help-section-title';
  const icon = SECTION_ICON[section.title] ?? '•';
  title.textContent = `${icon} ${section.title}`;
  sectionNode.appendChild(title);

  const list = document.createElement('div');
  list.className = 'session-help-item-list';

  section.rows.forEach((row, rowIndex) => {
    const globalIndex = (globalIndexMap[sectionIndex] ?? 0) + rowIndex;
    const button = createShortcutRow(row, globalIndex);
    list.appendChild(button);
  });

  sectionNode.appendChild(list);
  return sectionNode;
}
