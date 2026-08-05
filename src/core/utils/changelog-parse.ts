import type { ChangelogEntry, ChangelogItem, ChangelogSection } from '../../types/changelog';

const VERSION_HEADING = /^##\s+v(\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?)\s*(?:\(([^)]*)\))?\s*$/;
const SECTION_HEADING = /^###\s+(.+?)\s*$/;
const BULLET = /^(\s*)[-*]\s+(.*)$/;

/**
 * Entries are grouped by `major.minor` so the whole current minor line renders
 * expanded, matching how docs-site/changelog.md splits current vs previous.
 */
export function resolveChangelogGroupKey(version: string): string {
  const match = version.match(/^(\d+)\.(\d+)/);
  if (!match) return version;
  return `${match[1]}.${match[2]}`;
}

/**
 * Parses the repo CHANGELOG.md into version entries. Bullets keep their inline
 * markdown and their nesting: older entries group related notes under a bold
 * lead bullet with indented children, and flattening them loses that structure.
 */
export function parseChangelog(markdown: string): ChangelogEntry[] {
  const entries: ChangelogEntry[] = [];
  let entry: ChangelogEntry | null = null;
  let section: ChangelogSection | null = null;
  let internal = false;
  // Open bullets from outermost to innermost, used to place the next bullet.
  let openItems: Array<{ indent: number; item: ChangelogItem }> = [];

  function startSection(heading: string): void {
    section = { heading, items: [], internal };
    openItems = [];
    entry?.sections.push(section);
  }

  function addBullet(indent: number, text: string): void {
    if (!section) {
      // Bullets before any "###" heading (older entries) land in a generic group.
      startSection('Changes');
    }
    const item: ChangelogItem = { text, children: [] };

    while (openItems.length > 0 && (openItems[openItems.length - 1]?.indent ?? 0) >= indent) {
      openItems.pop();
    }
    const parent = openItems[openItems.length - 1];
    if (parent) {
      parent.item.children.push(item);
    } else {
      section?.items.push(item);
    }
    openItems.push({ indent, item });
  }

  function appendContinuation(text: string): void {
    const current = openItems[openItems.length - 1];
    if (!current) return;
    current.item.text = `${current.item.text} ${text}`;
  }

  for (const rawLine of markdown.split(/\r?\n/)) {
    const line = rawLine.trimEnd();
    const trimmed = line.trim();

    const versionMatch = trimmed.match(VERSION_HEADING);
    if (versionMatch) {
      const version = versionMatch[1] ?? '';
      entry = {
        version,
        date: versionMatch[2]?.trim() ?? '',
        groupKey: resolveChangelogGroupKey(version),
        sections: [],
      };
      entries.push(entry);
      section = null;
      internal = false;
      openItems = [];
      continue;
    }

    if (!entry) continue;

    if (trimmed.startsWith('<details')) {
      internal = true;
      section = null;
      openItems = [];
      continue;
    }
    if (trimmed.startsWith('</details')) {
      internal = false;
      section = null;
      openItems = [];
      continue;
    }
    if (trimmed.startsWith('<summary')) continue;

    const sectionMatch = trimmed.match(SECTION_HEADING);
    if (sectionMatch) {
      startSection(sectionMatch[1] ?? '');
      continue;
    }

    const bulletMatch = line.match(BULLET);
    if (bulletMatch) {
      addBullet((bulletMatch[1] ?? '').length, bulletMatch[2] ?? '');
      continue;
    }

    // A blank line ends the current bullet run; an indented non-bullet line is a
    // wrapped continuation of the bullet above it.
    if (!trimmed) {
      continue;
    }
    if (/^\s/.test(line)) {
      appendContinuation(trimmed);
    }
  }

  return entries.map((item) => ({
    ...item,
    sections: item.sections.filter((entrySection) => entrySection.items.length > 0),
  }));
}
