import type {
  ChangelogEntry,
  ChangelogItem,
  ChangelogSection,
  ChangelogSnapshot,
} from '../../types/changelog';
import { compareSemverLike } from '../../core/utils/semver-compare';

const SECTION_ICON: Record<string, string> = {
  Added: '✦',
  Changed: '⟲',
  Fixed: '✔',
  Docs: '▤',
  Internal: '⚙',
  'Breaking Changes': '⚠',
  Changes: '•',
};

export type ChangelogEntryBadge = 'installed' | 'newer' | null;

export function resolveEntryBadge(
  entryVersion: string,
  installedVersion: string,
): ChangelogEntryBadge {
  const comparison = compareSemverLike(entryVersion, installedVersion);
  if (comparison === 0) return 'installed';
  if (comparison > 0) return 'newer';
  return null;
}

/**
 * Version entries in the newest major.minor line start expanded, mirroring the
 * docs-site changelog where older lines sit behind "Previous Versions".
 */
export function shouldEntryStartExpanded(
  entry: ChangelogEntry,
  snapshot: Pick<ChangelogSnapshot, 'expandedGroupKey'>,
): boolean {
  if (!snapshot.expandedGroupKey) return false;
  return entry.groupKey === snapshot.expandedGroupKey;
}

type InlineToken =
  | { kind: 'text'; value: string }
  | { kind: 'code'; value: string }
  | { kind: 'strong'; value: string }
  | { kind: 'link'; value: string; href: string };

/**
 * Minimal inline-markdown tokenizer for changelog bullets: backtick code,
 * bold, and links. Anything else stays literal text.
 */
export function tokenizeInlineMarkdown(text: string): InlineToken[] {
  const tokens: InlineToken[] = [];
  const pattern = /`([^`]+)`|\*\*([^*]+)\*\*|\[([^\]]+)\]\(([^)\s]+)\)/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      tokens.push({ kind: 'text', value: text.slice(lastIndex, match.index) });
    }
    if (match[1] !== undefined) {
      tokens.push({ kind: 'code', value: match[1] });
    } else if (match[2] !== undefined) {
      tokens.push({ kind: 'strong', value: match[2] });
    } else if (match[3] !== undefined && match[4] !== undefined) {
      tokens.push({ kind: 'link', value: match[3], href: match[4] });
    }
    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    tokens.push({ kind: 'text', value: text.slice(lastIndex) });
  }
  return tokens;
}

function appendInlineMarkdown(target: HTMLElement, text: string): void {
  for (const token of tokenizeInlineMarkdown(text)) {
    if (token.kind === 'text') {
      target.appendChild(document.createTextNode(token.value));
      continue;
    }
    if (token.kind === 'code') {
      const code = document.createElement('code');
      code.className = 'changelog-code';
      code.textContent = token.value;
      target.appendChild(code);
      continue;
    }
    if (token.kind === 'strong') {
      const strong = document.createElement('strong');
      strong.textContent = token.value;
      target.appendChild(strong);
      continue;
    }
    // Links stay inert: the overlay has nowhere to navigate to.
    const link = document.createElement('span');
    link.className = 'changelog-link';
    link.textContent = token.value;
    link.title = token.href;
    target.appendChild(link);
  }
}

function createItemList(items: ChangelogItem[], depth: number): HTMLUListElement {
  const list = document.createElement('ul');
  list.className = depth === 0 ? 'changelog-items' : 'changelog-items changelog-items-nested';

  for (const item of items) {
    const listItem = document.createElement('li');
    listItem.className = 'changelog-item';

    const text = document.createElement('span');
    text.className = 'changelog-item-text';
    appendInlineMarkdown(text, item.text);
    listItem.appendChild(text);

    if (item.children.length > 0) {
      listItem.appendChild(createItemList(item.children, depth + 1));
    }
    list.appendChild(listItem);
  }
  return list;
}

function createSectionNode(section: ChangelogSection): HTMLElement {
  const node = document.createElement('section');
  node.className = 'changelog-section';

  const title = document.createElement('h4');
  title.className = 'changelog-section-title';
  title.textContent = `${SECTION_ICON[section.heading] ?? '•'} ${section.heading}`;
  node.appendChild(title);

  node.appendChild(createItemList(section.items, 0));
  return node;
}

function createInternalNode(sections: ChangelogSection[]): HTMLElement {
  const details = document.createElement('details');
  details.className = 'changelog-internal';

  const summary = document.createElement('summary');
  summary.className = 'changelog-internal-summary';
  summary.textContent = 'Internal changes';
  details.appendChild(summary);

  for (const section of sections) {
    details.appendChild(createSectionNode(section));
  }
  return details;
}

export function createChangelogEntryNode(
  entry: ChangelogEntry,
  options: { expanded: boolean; badge: ChangelogEntryBadge; index: number },
): HTMLDetailsElement {
  const details = document.createElement('details');
  details.className = 'changelog-entry';
  details.open = options.expanded;
  details.dataset.changelogVersion = entry.version;

  const summary = document.createElement('summary');
  summary.className = 'changelog-entry-summary';
  summary.dataset.changelogIndex = String(options.index);
  summary.tabIndex = -1;

  const version = document.createElement('span');
  version.className = 'changelog-entry-version';
  version.textContent = `v${entry.version}`;
  summary.appendChild(version);

  if (entry.date) {
    const date = document.createElement('span');
    date.className = 'changelog-entry-date';
    date.textContent = entry.date;
    summary.appendChild(date);
  }

  if (options.badge) {
    const badge = document.createElement('span');
    badge.className = `changelog-entry-badge changelog-entry-badge-${options.badge}`;
    badge.textContent = options.badge === 'installed' ? 'Installed' : 'New';
    summary.appendChild(badge);
  }

  details.appendChild(summary);

  const body = document.createElement('div');
  body.className = 'changelog-entry-body';
  const publicSections = entry.sections.filter((section) => !section.internal);
  const internalSections = entry.sections.filter((section) => section.internal);

  for (const section of publicSections) {
    body.appendChild(createSectionNode(section));
  }
  if (internalSections.length > 0) {
    body.appendChild(createInternalNode(internalSections));
  }
  if (publicSections.length === 0 && internalSections.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'changelog-empty-entry';
    empty.textContent = 'No release notes recorded for this version.';
    body.appendChild(empty);
  }

  details.appendChild(body);
  return details;
}

export function describeChangelogSource(snapshot: ChangelogSnapshot): string {
  if (snapshot.source === 'remote') {
    return snapshot.releaseTag
      ? `Latest release ${snapshot.releaseTag}`
      : 'Latest published changelog';
  }
  return 'Bundled changelog';
}
