import type { SessionHelpItem, SessionHelpSection } from './session-help-sections';

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
  'Playback and navigation': '▶',
  'Visual feedback': '◉',
  'Subtitle sync': '⟲',
  'Mining and capture': '✦',
  'Stats and progress': '◉',
  'Overlay controls': '◈',
  'Modals and tools': '▣',
  'Runtime settings': '⚙',
  'System actions': '◆',
  'Other shortcuts': '…',
  'Fixed overlay controls': '◇',
  'Y chords': '⌘',
  'Global shortcuts': '◆',
  'Color legend': '◈',
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
